"""Importação, revisão e exportação de YAML exclusivamente em memória.

Não há diretórios temporários, extração de ZIP nem escrita de metadados em disco.
Os documentos round-trip ficam vivos apenas no objeto Batch da sessão.
"""
from __future__ import annotations

import base64
import binascii
import copy
import hashlib
import io
import json
import math
import os
import stat
import time
import zipfile
import zlib
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import PurePosixPath
from typing import Any

from ruamel.yaml import YAML
from ruamel.yaml.comments import CommentedMap
from ruamel.yaml.scalarstring import DoubleQuotedScalarString, LiteralScalarString
from ruamel.yaml.tokens import (AliasToken, TagToken, BlockMappingStartToken,
                               BlockSequenceStartToken, BlockEndToken,
                               FlowMappingStartToken, FlowMappingEndToken,
                               FlowSequenceStartToken, FlowSequenceEndToken)

from .rules import get_catalog, validate_tag

def _setting(name, default, minimum, maximum):
    try:
        value = int(os.environ.get("METADATA_" + name, default))
        if not minimum <= value <= maximum:
            raise ValueError
        return value
    except (ValueError, TypeError):
        raise ValueError(f"Configuração METADATA_{name} inválida; use inteiro entre {minimum} e {maximum}.") from None


MAX_FILE_BYTES = _setting("MAX_FILE_BYTES", 10 * 1024 * 1024, 1024, 100 * 1024 * 1024)
MAX_BATCH_BYTES = _setting("MAX_BATCH_BYTES", 50 * 1024 * 1024, 1024, 250 * 1024 * 1024)
MAX_FILES = _setting("MAX_FILES", 500, 1, 5000)
MAX_DEPTH = _setting("MAX_DEPTH", 64, 4, 128)
MAX_NODES = _setting("MAX_NODES", 100_000, 100, 1_000_000)
MAX_ZIP_ENTRIES = _setting("MAX_ZIP_ENTRIES", 1000, 1, 10_000)
MAX_ZIP_RATIO = _setting("MAX_ZIP_RATIO", 200, 2, 1000)
MAX_PROCESS_SECONDS = _setting("MAX_PROCESS_SECONDS", 30, 1, 120)
STANDARD_TAGS = {"str", "int", "float", "bool", "null", "timestamp", "seq", "map"}


def _check_deadline(deadline):
    if deadline is not None and time.monotonic() > deadline:
        raise ValueError("Tempo máximo de processamento atingido; divida o lote em arquivos menores.")


class FrozenDict(dict):
    """Snapshot imutável compatível com o serializador JSON padrão."""
    def _immutable(self, *args, **kwargs):
        raise TypeError("Snapshot original é somente leitura.")
    __setitem__ = __delitem__ = clear = pop = popitem = setdefault = update = __ior__ = _immutable
    def __deepcopy__(self, memo):
        return self


def _plain(value):
    if isinstance(value, dict):
        return {str(k): _plain(v) for k, v in value.items()}
    if isinstance(value, (list, tuple)):
        return [_plain(v) for v in value]
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if isinstance(value, str):
        return str(value)
    if isinstance(value, float) and not math.isfinite(value):
        return str(value)
    if value is None or isinstance(value, (int, float, bool)):
        return value
    return str(value)


def _freeze(value):
    if isinstance(value, dict):
        return FrozenDict({k: _freeze(v) for k, v in value.items()})
    if isinstance(value, list):
        return tuple(_freeze(v) for v in value)
    return value


@dataclass
class Document:
    filename: str
    roots: list[Any]
    source_sha256: str
    explicit_start: bool = False


@dataclass
class Batch:
    documents: list[Document] = field(default_factory=list)
    nodes: list[dict] = field(default_factory=list)
    errors: list[dict] = field(default_factory=list)
    duplicates: list[dict] = field(default_factory=list)
    file_count: int = 0
    processed_count: int = 0
    graph_count: int = 0
    _bindings: dict = field(default_factory=dict, repr=False)
    _blocked_ids: set = field(default_factory=set, repr=False)


def _parser():
    yaml = YAML(typ="rt")
    yaml.preserve_quotes = True
    yaml.allow_duplicate_keys = False
    yaml.width = 4096
    return yaml


def _safe_filename(name: str) -> str:
    if not isinstance(name, str) or not name or len(name) > 240 or any(ord(ch) < 32 for ch in name):
        raise ValueError("Nome de arquivo inválido.")
    # Windows separators/drive letters and ZIP's POSIX paths share the same check.
    if any(ch in name for ch in '\\:<>"|?*') or name.startswith("/"):
        raise ValueError("Caminho absoluto ou separador de arquivo não permitido.")
    parts = name.split("/")
    if any(p in ("", ".", "..") for p in parts):
        raise ValueError("Caminho de arquivo inseguro.")
    reserved = {"CON", "PRN", "AUX", "NUL"} | {f"{prefix}{n}" for prefix in ("COM", "LPT") for n in range(1, 10)}
    if any(p.endswith((".", " ")) or p.split(".", 1)[0].upper() in reserved for p in parts):
        raise ValueError("Nome de arquivo incompatível com caminhos seguros no Windows.")
    return str(PurePosixPath(name))


def _check_graph(value, depth=0, active=None, budget=None, deadline=None):
    active = active if active is not None else set()
    budget = budget if budget is not None else [MAX_NODES]
    budget[0] -= 1
    if budget[0] % 256 == 0:
        _check_deadline(deadline)
    if budget[0] < 0:
        raise ValueError("YAML excede o limite de elementos.")
    if depth > MAX_DEPTH:
        raise ValueError("YAML excede o limite de profundidade.")
    if isinstance(value, (dict, list)):
        if id(value) in active:
            raise ValueError("Referência cíclica não permitida.")
        active.add(id(value))
        entries = list(value.items()) if isinstance(value, dict) else list(enumerate(value))
        for key, child in entries:
            if isinstance(value, dict) and not isinstance(key, str):
                raise ValueError("Chaves YAML devem ser textos simples.")
            _check_graph(child, depth + 1, active, budget, deadline)
        active.remove(id(value))


def _load(raw: bytes, deadline=None, graph_budget=None):
    _check_deadline(deadline)
    if len(raw) > MAX_FILE_BYTES:
        raise ValueError("Arquivo excede o limite de tamanho configurado.")
    try:
        source = raw.decode("utf-8-sig")
    except UnicodeDecodeError as exc:
        raise ValueError("Arquivo deve usar codificação UTF-8.") from exc
    parser = _parser()
    token_count = 0
    depth = 0
    # Inspeção sintática ANTES de construir objetos: aliases e tags arbitrárias
    # são rejeitados sem permitir expansão, ciclos ou construtores externos.
    for token in parser.scan(source):
        token_count += 1
        if token_count % 256 == 0:
            _check_deadline(deadline)
        if token_count > MAX_NODES * 8:
            raise ValueError("YAML excede o limite de elementos.")
        if isinstance(token, (BlockMappingStartToken, BlockSequenceStartToken, FlowMappingStartToken, FlowSequenceStartToken)):
            depth += 1
            if depth > MAX_DEPTH:
                raise ValueError("YAML excede o limite de profundidade.")
        elif isinstance(token, (BlockEndToken, FlowMappingEndToken, FlowSequenceEndToken)):
            depth -= 1
        if isinstance(token, AliasToken):
            raise ValueError("Aliases YAML não são permitidos; envie valores explícitos.")
        if isinstance(token, TagToken):
            handle, suffix = token.value
            standard = (handle == "!!" and suffix in STANDARD_TAGS) or (handle is None and suffix.startswith("tag:yaml.org,2002:") and suffix.rsplit(":", 1)[-1] in STANDARD_TAGS)
            if not standard:
                raise ValueError("Tags YAML customizadas ou executáveis não são permitidas.")
    try:
        roots = list(_parser().load_all(source))
    except (ValueError, OverflowError):
        # Explicit standard tags such as !!int can raise built-in exceptions
        # that quote their scalar value. Never expose that source content.
        raise ValueError("YAML inválido; confira os tipos declarados e a sintaxe.") from None
    _check_deadline(deadline)
    if not roots or all(root is None for root in roots):
        raise ValueError("Arquivo YAML vazio.")
    _check_graph(roots, budget=graph_budget, deadline=deadline)
    return roots, source.lstrip().startswith("---")


def _sid(level, path):
    return hashlib.sha256(f"{level}\0{path}".encode("utf-8")).hexdigest()[:24]


def _extract(batch: Batch, document: Document, deadline=None):
    result: list[tuple[dict, dict]] = []
    def fail(path, message):
        batch.errors.append({"filename": document.filename, "field": path, "message": message})
    def collection(value, level, context, fieldpath):
        if isinstance(value, list):
            for index, item in enumerate(value):
                visit(item, level, context, f"{fieldpath}[{index}]")
        elif isinstance(value, dict):
            # A singular declaration can occur in a plural collection.
            if any(k in value for k in ("id", "name")):
                visit(value, level, context, fieldpath)
            else:
                for key, item in value.items():
                    visit(item, level, context, f"{fieldpath}.{key}", key)
        else:
            fail(fieldpath, "Esperada lista ou mapa de objetos.")
    def visit(data, level, context, fieldpath, key_hint=None):
        _check_deadline(deadline)
        if not isinstance(data, dict):
            fail(fieldpath, "Declaração do objeto deve ser um mapa YAML.")
            return
        expected = {"Catalog": 1, "Schema": 2, "Table": 3, "View": 3, "Column": 4}[level]
        identity = data.get("id")
        name = data.get("name", key_hint)
        if identity is not None and not isinstance(identity, str):
            fail(f"{fieldpath}.id", "Identificador deve ser texto.")
            return
        if name is not None and not isinstance(name, str):
            fail(f"{fieldpath}.name", "Nome deve ser texto.")
            return
        if identity:
            parts = identity.split(".")
            if len(parts) == 1 and context:
                parts = [*context, identity]
        elif name:
            parts = name.split(".") if "." in name else [*context, name]
        else:
            fail(fieldpath, "Objeto sem nome ou identificador inequívoco.")
            return
        if len(parts) != expected or any(not p.strip() or p != p.strip() for p in parts):
            fail(f"{fieldpath}.id", f"Caminho inequívoco de {expected} nível(is) esperado para {level}.")
            return
        if context and parts[:-1] != context:
            fail(f"{fieldpath}.id", "Identificador diverge da hierarquia que contém o objeto.")
            return
        if name and "." not in name and name != parts[-1]:
            fail(f"{fieldpath}.name", "Nome técnico diverge do identificador do objeto.")
            return
        tags = data.get("tags", {})
        if not isinstance(tags, dict):
            fail(f"{fieldpath}.tags", "Tags devem ser um mapa de nomes e valores.")
            return
        for editable in ("description", "business_name"):
            if data.get(editable) is not None and not isinstance(data[editable], str):
                fail(f"{fieldpath}.{editable}", "Campo de negócio deve ser texto; o valor original foi preservado.")
        path = ".".join(parts)
        obj_id = _sid(level, path)
        if level == "Column":
            obj_id = _sid(context_level[0], ".".join(parts[:-1]))
        plain_tags = _plain(tags)
        node = {"id": _sid(level, path), "level": level, "path": path,
                "catalog": parts[0], "schema": parts[1] if len(parts) > 1 else "",
                "name": parts[-1], "object_id": obj_id, "filename": document.filename,
                "source_sha256": document.source_sha256,
                "description": _plain(data.get("description")),
                "business_name": _plain(data.get("business_name")), "tags": plain_tags,
                "_native_tag_types": {str(k): type(v).__name__ for k, v in tags.items() if not isinstance(v, str)},
                "original_description": _freeze(_plain(data.get("description"))),
                "original_business_name": _freeze(_plain(data.get("business_name"))),
                "original_tags": _freeze(plain_tags)}
        for key in ("type", "data_type", "last_modified", "updated_at", "nullable"):
            if key in data:
                node[key] = _plain(data[key])
        result.append((node, data))
        child_context = parts
        for key, child_level in (("schemas", "Schema"), ("schema", "Schema"), ("tables", "Table"), ("table", "Table"), ("views", "View"), ("view", "View"), ("columns", "Column")):
            if key not in data:
                continue
            allowed = (level == "Catalog" and child_level == "Schema") or (level == "Schema" and child_level in ("Table", "View")) or (level in ("Table", "View") and child_level == "Column")
            if not allowed:
                fail(f"{fieldpath}.{key}", "Nível de objeto incompatível com a hierarquia.")
                continue
            old = context_level[0]
            if child_level == "Column":
                context_level[0] = level
            if key in ("schema", "table", "view"):
                visit(data[key], child_level, child_context, f"{fieldpath}.{key}")
            else:
                collection(data[key], child_level, child_context, f"{fieldpath}.{key}")
            context_level[0] = old
    context_level = ["Table"]
    for index, root in enumerate(document.roots):
        if not isinstance(root, dict):
            fail(f"documento[{index}]", "Raiz YAML deve ser um mapa com catalog/schema/table/view ou coleções correspondentes.")
            continue
        found = False
        for key, level in (("catalog", "Catalog"), ("catalogs", "Catalog"), ("schema", "Schema"), ("schemas", "Schema"), ("table", "Table"), ("tables", "Table"), ("view", "View"), ("views", "View")):
            if key in root:
                found = True
                if key.endswith("s"):
                    collection(root[key], level, [], key)
                else:
                    visit(root[key], level, [], key)
        if not found:
            explicit = root.get("kind", root.get("object_type"))
            if isinstance(explicit, str) and explicit.capitalize() in ("Catalog", "Schema", "Table", "View"):
                visit(root, explicit.capitalize(), [], "objeto")
            else:
                fail(f"documento[{index}]", "Estrutura não reconhecida; nenhum objeto foi descartado ou convertido automaticamente.")
    return result


def parse_batch(files: list[dict], references: dict | None = None) -> Batch:
    get_catalog(references)
    batch = Batch()
    deadline = time.monotonic() + MAX_PROCESS_SECONDS
    if not isinstance(files, list) or len(files) > MAX_FILES:
        raise ValueError(f"Envie uma lista com até {MAX_FILES} arquivos.")
    expanded: list[tuple[str, bytes]] = []
    total = 0
    names = set()
    def error(filename, message, field="arquivo"):
        batch.errors.append({"filename": str(filename), "field": field, "message": message})
    def add(name, raw):
        nonlocal total
        _check_deadline(deadline)
        name = _safe_filename(name)
        if len(raw) > MAX_FILE_BYTES:
            raise ValueError("Arquivo YAML excede o limite de tamanho configurado.")
        total += len(raw)
        if total > MAX_BATCH_BYTES:
            raise ValueError("Lote descompactado excede o limite de tamanho configurado.")
        if len(expanded) >= MAX_FILES:
            raise ValueError(f"Lote contém mais de {MAX_FILES} arquivos YAML.")
        if name in names:
            error(name, "Nome de arquivo duplicado no lote; renomeie ou substitua o arquivo explicitamente.")
            return
        names.add(name)
        expanded.append((name, raw))
    for item in files:
        name = item.get("name", "arquivo") if isinstance(item, dict) else "arquivo"
        try:
            _check_deadline(deadline)
            name = _safe_filename(name)
            encoded = item.get("content")
            if not isinstance(encoded, str) or len(encoded) > (MAX_BATCH_BYTES * 4 // 3 + 8):
                raise ValueError("Conteúdo base64 ausente ou maior que o limite do lote.")
            raw = base64.b64decode(encoded, validate=True)
            suffix = PurePosixPath(name).suffix.lower()
            if suffix in (".yaml", ".yml"):
                add(name, raw)
            elif suffix == ".zip":
                if len(raw) > MAX_BATCH_BYTES:
                    raise ValueError("ZIP excede o limite de tamanho configurado.")
                with zipfile.ZipFile(io.BytesIO(raw)) as archive:
                    infos = archive.infolist()
                    if len(infos) > MAX_ZIP_ENTRIES:
                        raise ValueError("ZIP contém entradas em excesso.")
                    # Preflight all entries before decompressing any one of them.
                    planned = []
                    uncompressed = 0
                    for info in infos:
                        _check_deadline(deadline)
                        original_entry = info.orig_filename
                        if info.is_dir():
                            _safe_filename(original_entry.rstrip("/"))
                            continue
                        entry = _safe_filename(original_entry)
                        if stat.S_ISLNK(info.external_attr >> 16):
                            raise ValueError("Links simbólicos não são permitidos no ZIP.")
                        if info.flag_bits & 1:
                            raise ValueError("ZIP protegido por senha não é suportado.")
                        if PurePosixPath(entry).suffix.lower() not in (".yaml", ".yml"):
                            raise ValueError("ZIP deve conter somente YAMLs; arquivos aninhados ou outros formatos não são aceitos.")
                        if info.file_size > MAX_FILE_BYTES or (info.file_size > 1024 * 1024 and info.file_size / max(info.compress_size, 1) > MAX_ZIP_RATIO):
                            raise ValueError("Entrada ZIP excede o tamanho ou a taxa de compactação permitidos.")
                        uncompressed += info.file_size
                        if total + uncompressed > MAX_BATCH_BYTES or len(expanded) + len(planned) >= MAX_FILES:
                            raise ValueError("ZIP excede os limites configurados do lote.")
                        planned.append((info, entry))
                    if not planned:
                        raise ValueError("ZIP não contém arquivos YAML.")
                    for info, entry in planned:
                        add(f"{PurePosixPath(name).stem}/{entry}", archive.read(info))
            else:
                raise ValueError("Formato não suportado; envie .yaml, .yml ou .zip.")
        except zlib.error:
            error(name, "Conteúdo ZIP corrompido; envie um novo arquivo compactado.")
        except (ValueError, binascii.Error, zipfile.BadZipFile, RuntimeError, OSError) as exc:
            error(name, str(exc))
    batch.file_count = len(expanded)
    seen = {}
    graph_remaining = MAX_NODES
    for name, raw in expanded:
        try:
            candidate_budget = [graph_remaining]
            roots, explicit_start = _load(raw, deadline, candidate_budget)
            document = Document(name, roots, hashlib.sha256(raw).hexdigest(), explicit_start)
            extracted = _extract(batch, document, deadline)
            batch.documents.append(document)
            graph_remaining = candidate_budget[0]
            batch.graph_count = MAX_NODES - graph_remaining
            batch.processed_count += 1
            for node, mapping in extracted:
                if node["id"] in seen:
                    previous = seen[node["id"]]
                    same = _plain(mapping) == _plain(batch._bindings[node["id"]])
                    duplicate = {"id": node["id"], "path": node["path"], "level": node["level"], "filenames": [previous["filename"], node["filename"]], "conflict": not same,
                                 "message": "Objeto duplicado com conteúdo divergente." if not same else "Objeto repetido com conteúdo idêntico."}
                    batch.duplicates.append(duplicate)
                    batch._blocked_ids.add(node["id"])
                    previous["conflict"] = True
                    error(name, duplicate["message"] + " Resolva a duplicidade antes de gerar YAMLs.", node["path"])
                    continue
                seen[node["id"]] = node
                batch.nodes.append(node)
                batch._bindings[node["id"]] = mapping
        except Exception as exc:
            # Parser diagnostics may quote sensitive source snippets. Expose only
            # a concise location and safe reason, never a stack trace/source line.
            if isinstance(exc, ValueError):
                message = str(exc)
            else:
                mark = getattr(exc, "problem_mark", None)
                where = f" na linha {mark.line + 1}, coluna {mark.column + 1}" if mark else ""
                message = f"YAML inválido{where}; verifique sintaxe, chaves repetidas e estrutura."
            error(name, message)
    return batch


def _styled(value, old=None):
    if isinstance(value, str):
        if value in ("true", "false"):
            return DoubleQuotedScalarString(value)
        if isinstance(old, str) and type(old) is not str:
            try:
                return type(old)(value)
            except TypeError:
                pass
        if "\n" in value:
            return LiteralScalarString(value)
    return value


def apply_edits(batch: Batch, edits: list[dict], references: dict | None = None) -> list[dict]:
    catalog = get_catalog(references)
    if not isinstance(edits, list) or len(edits) > MAX_NODES:
        raise ValueError("Lista de alterações inválida ou grande demais.")
    by_id = {node["id"]: node for node in batch.nodes}
    reports = []
    values_by_key = {}
    for edit in edits:
        if not isinstance(edit, dict):
            continue
        key = (edit.get("id"), edit.get("field"), edit.get("tag"))
        if any(v is not None and not isinstance(v, str) for v in key):
            continue
        values_by_key.setdefault(key, set()).add(json.dumps(edit.get("value"), ensure_ascii=False, sort_keys=True, default=str))
    for edit in edits:
        if not isinstance(edit, dict):
            reports.append({"id": None, "field": None, "before": None, "after": None, "status": "rejected", "message": "Alteração deve ser um objeto."})
            continue
        nid, fieldname, tag, value = edit.get("id"), edit.get("field"), edit.get("tag"), edit.get("value")
        if any(v is not None and not isinstance(v, str) for v in (nid, fieldname, tag)):
            reports.append({"id": None, "field": None, "before": None, "after": None, "status": "rejected", "message": "Identificador, campo e nome da tag devem ser textos."})
            continue
        node = by_id.get(nid)
        before = node.get("tags", {}).get(tag) if node and fieldname == "tags" else node.get(fieldname) if node else None
        report = {"id": nid, "field": fieldname, "before": copy.deepcopy(before), "after": _plain(value), "status": "rejected"}
        if tag is not None:
            report["tag"] = tag
        reason = None
        if node is None:
            reason = "Objeto desconhecido na sessão atual."
        elif nid in batch._blocked_ids:
            reason = "Objeto duplicado; resolva os arquivos conflitantes antes de editar."
        elif len(values_by_key[(nid, fieldname, tag)]) > 1:
            reason = "Alterações conflitantes para o mesmo campo."
        elif fieldname not in ("description", "business_name", "tags"):
            reason = "Campo técnico ou não editável."
        elif fieldname == "tags" and (not isinstance(tag, str) or not tag):
            reason = "Nome da tag ausente."
        elif fieldname == "tags" and tag not in node["tags"] and tag not in catalog[node["level"]]:
            reason = "Não é permitido criar tag que não consta no catálogo deste nível."
        elif value is not None and not isinstance(value, str):
            reason = 'Alterações devem ser textos; use "true" e "false" para tags booleanas.'
        elif isinstance(value, str) and len(value) > (10_000 if fieldname == "tags" else 100_000):
            reason = "Texto excede o limite permitido."
        elif fieldname == "tags" and value is not None:
            invalid = [i for i in validate_tag(node["level"], tag, value, references, catalog=catalog) if i["severity"] == "error"]
            if invalid:
                reason = invalid[0]["message"]
        if reason:
            report["message"] = reason
        elif type(before) is type(value) and before == value and not (fieldname == "tags" and tag in node.get("_native_tag_types", {})) and not (value is None and ((fieldname == "tags" and tag in node["tags"]) or (fieldname != "tags" and fieldname in batch._bindings[nid]))):
            report["status"] = "unchanged"
        else:
            mapping = batch._bindings[nid]
            if fieldname == "tags":
                node.get("_native_tag_types", {}).pop(tag, None)
                if "tags" not in mapping:
                    mapping["tags"] = CommentedMap()
                if value is None:
                    mapping["tags"].pop(tag, None)
                    node["tags"].pop(tag, None)
                else:
                    mapping["tags"][tag] = _styled(value, mapping["tags"].get(tag))
                    node["tags"][tag] = value
            else:
                if value is None:
                    mapping.pop(fieldname, None)
                else:
                    mapping[fieldname] = _styled(value, mapping.get(fieldname))
                node[fieldname] = value
            report["status"] = "applied"
        reports.append(report)
    return reports


def render_yaml_zip(batch: Batch, report: list[dict]) -> bytes:
    if batch.duplicates:
        raise ValueError("Resolva os objetos duplicados antes de gerar YAMLs atualizados.")
    output = io.BytesIO()
    deadline = time.monotonic() + MAX_PROCESS_SECONDS
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for document in batch.documents:
            _check_deadline(deadline)
            buffer = io.StringIO()
            yaml = _parser()
            yaml.explicit_start = document.explicit_start or len(document.roots) > 1
            yaml.dump_all(document.roots, buffer)
            encoded = buffer.getvalue().encode("utf-8")
            _load(encoded, deadline)  # Verificação sintática antes de liberar o download.
            archive.writestr(_safe_filename(document.filename), encoded)
        archive.writestr("report.json", json.dumps(report, ensure_ascii=False, indent=2, default=str).encode("utf-8"))
    return output.getvalue()

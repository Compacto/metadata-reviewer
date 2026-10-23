"""Conversão YAML ↔ Excel. Sem sessões de aprovação nem persistência de arquivos."""
from __future__ import annotations

from io import BytesIO
from zipfile import ZipFile, ZIP_DEFLATED

from . import engine
from .engine import parse_batch, apply_edits, render_yaml_zip
from .owner_workbook import build_owner_workbook, read_owner_workbook
from .rules import validate_nodes

MAX_WORKBOOK_BYTES = 15 * 1024 * 1024
MAX_WORKBOOKS = 50
MAX_WORKBOOK_BATCH_BYTES = 50 * 1024 * 1024


def _batch(files):
    if not isinstance(files, list) or not files:
        raise ValueError("Selecione pelo menos um YAML ou ZIP.")
    batch = parse_batch(files)
    if batch.errors:
        first = batch.errors[0]
        raise ValueError(
            f"O lote não foi convertido: {first.get('filename', 'arquivo')} · "
            f"{first.get('field', 'arquivo')} · {first.get('message', 'Arquivo inválido')}"
            + (f" Há mais {len(batch.errors) - 1} ocorrência(s)." if len(batch.errors) > 1 else "")
        )
    if not batch.nodes or not batch.documents:
        raise ValueError("Os arquivos não contêm objetos de metadados reconhecidos.")
    return batch


def _warnings(nodes, issues, workbook_name=None):
    index = {n["id"]: n for n in nodes}
    result = []
    for issue in issues:
        node = index.get(issue.get("id"), {})
        result.append({
            "filename": node.get("filename", ""),
            "path": node.get("path", "Referência de governança"),
            "field": issue.get("field", ""),
            "message": issue.get("message", "Verificar com Governança."),
            "severity": issue.get("severity", "pending"),
            **({"workbook": workbook_name} if workbook_name else {}),
        })
    return result


def create_owner_excel(files):
    """Build one portable workbook from individual YAMLs and/or ZIP archives."""
    batch = _batch(files)
    issues = validate_nodes(batch.nodes)
    content = build_owner_workbook(files, batch.nodes, issues)
    if len(content) > MAX_WORKBOOK_BYTES:
        raise ValueError("O Excel excede 15 MiB. Divida os YAMLs em lotes menores.")
    return {
        "content": content,
        "filename": "revisao_metadados_owner.xlsx",
        "summary": {
            "files": len(batch.documents),
            "objects": sum(n["level"] in ("Table", "View") for n in batch.nodes),
            "columns": sum(n["level"] == "Column" for n in batch.nodes),
        },
        "warnings": _warnings(batch.nodes, issues),
    }


def _updated_yaml_files(batch):
    # Reuse the audited round-trip serializer and syntax recheck. The deployment
    # ZIP contains YAMLs only; internal diagnostic records are kept out of it.
    output = render_yaml_zip(batch, [])
    with ZipFile(BytesIO(output)) as archive:
        return [
            {"name": item.filename, "content": archive.read(item)}
            for item in archive.infolist()
            if item.filename.lower().endswith((".yaml", ".yml"))
        ]


def convert_owner_excels(workbooks):
    """Recover original metadata from returned workbooks and apply owner edits.

    Existing source inconsistencies are advisory and never suppress an object.
    Invalid proposed field formats or new consistency errors fail the conversion
    as a whole, so the analyst never receives a silently partial deployment ZIP.
    """
    if not isinstance(workbooks, list) or not 1 <= len(workbooks) <= MAX_WORKBOOKS:
        raise ValueError(f"Selecione entre 1 e {MAX_WORKBOOKS} planilhas Excel.")
    total = 0
    for item in workbooks:
        if (not isinstance(item, dict) or not isinstance(item.get("name"), str)
                or not isinstance(item.get("content"), bytes)):
            raise ValueError("Planilha sem nome ou conteúdo válido.")
        if not 1 <= len(item["name"]) <= 240 or any(ord(c) < 32 for c in item["name"]):
            raise ValueError("Nome de planilha inválido.")
        size = len(item["content"])
        if not size or size > MAX_WORKBOOK_BYTES:
            raise ValueError("Cada Excel deve ter até 15 MiB.")
        total += size
    if total > MAX_WORKBOOK_BATCH_BYTES:
        raise ValueError("O lote de planilhas excede 50 MiB. Divida a conversão em lotes menores.")

    yamls, changes, warnings, seen_ids = [], [], [], set()
    object_count = 0
    source_bytes = output_bytes = source_files = graph_count = 0
    for position, item in enumerate(workbooks, 1):
        try:
            decoded = read_owner_workbook(item["content"])
        except ValueError as exc:
            raise ValueError(f'{item["name"]}: {exc}') from None
        source_bytes += sum(len(f["content"]) * 3 // 4 - (len(f["content"]) - len(f["content"].rstrip("=")))
                            for f in decoded["files"])
        if source_bytes > engine.MAX_BATCH_BYTES:
            raise ValueError("As origens recuperadas dos Excels excedem o limite total do lote. Converta menos planilhas por vez.")
        batch = _batch(decoded["files"])
        source_files += len(batch.documents)
        graph_count += batch.graph_count
        if source_files > engine.MAX_FILES or graph_count > engine.MAX_NODES:
            raise ValueError("As origens recuperadas dos Excels excedem o limite de arquivos ou elementos do lote.")
        ids = {n["id"] for n in batch.nodes}
        if seen_ids & ids:
            raise ValueError("O mesmo objeto aparece em mais de um Excel. Envie apenas uma versão de cada objeto.")
        seen_ids.update(ids)
        before = validate_nodes(batch.nodes)
        report = apply_edits(batch, decoded.get("edits", []))
        if any(r["status"] == "rejected" for r in report):
            reason = next(r.get("message", "Alteração inválida.") for r in report if r["status"] == "rejected")
            raise ValueError("Uma alteração do Excel não pode ser aplicada: " + reason)
        after = validate_nodes(batch.nodes)
        old_errors = {(i.get("id"), i.get("field"), i.get("code"))
                      for i in before if i.get("severity") == "error"}
        introduced = [i for i in after if i.get("severity") == "error"
                      and (i.get("id"), i.get("field"), i.get("code")) not in old_errors]
        if introduced:
            raise ValueError("As alterações introduzem uma inconsistência: " + introduced[0]["message"]
                             + " Corrija as propostas correspondentes no Excel e reenvie.")
        for record in _updated_yaml_files(batch):
            output_bytes += len(record["content"])
            if output_bytes > engine.MAX_BATCH_BYTES:
                raise ValueError("Os YAMLs gerados excedem o limite total do lote. Converta menos planilhas por vez.")
            name = record["name"] if len(workbooks) == 1 else f"lote_{position:02d}/" + record["name"]
            yamls.append({"name": name, "content": record["content"]})
        object_count += sum(n["level"] in ("Table", "View") for n in batch.nodes)
        changes.extend({"workbook": item["name"], **row} for row in report if row["status"] == "applied")
        warnings.extend(_warnings(batch.nodes, after, item["name"]))
        warnings.extend({"workbook": item["name"], **row} for row in decoded.get("warnings", [])
                        if row.get("source") not in {"original", "proposal"})
        warnings.extend({"workbook": item["name"], "message": "Comentário do Owner: " + str(note.get("comment", note.get("message", ""))),
                         "path": note.get("path", ""), "field": "Comentários do Owner"}
                        for note in decoded.get("notes", []))

    names = [y["name"] for y in yamls]
    if len(names) != len(set(names)):
        raise ValueError("Há nomes de arquivos repetidos no pacote final. Separe as versões antes de converter.")
    buffer = BytesIO()
    with ZipFile(buffer, "w", compression=ZIP_DEFLATED) as archive:
        for record in yamls:
            archive.writestr(record["name"], record["content"])
    return {
        "content": buffer.getvalue(), "filename": "yamls_revisados.zip",
        "yamls": yamls, "changes": changes, "warnings": warnings,
        "summary": {"files": len(yamls), "objects": object_count, "changes": len(changes)},
    }

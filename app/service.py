"""Application operations with no filesystem persistence or global user state."""
from __future__ import annotations

import copy
from datetime import datetime, timezone
from io import BytesIO
import json
from zipfile import ZipFile, ZIP_DEFLATED

from .engine import apply_edits, parse_batch, render_yaml_zip
from .excel import export_workbook, preview_workbook
from .rules import get_catalog, validate_nodes

STATUSES = ["Pending", "In review", "Awaiting clarification", "Validated", "Validated with remarks", "Rejected"]
REFERENCE_KEYS = {"nm_dom", "nm_subdom", "nm_sist_orig", "ct_area_neg", "dominios", "subdominios", "sistemas_origem", "areas_negocio"}


def validate_input(files, edits, reviews, refs):
    if not isinstance(files, list) or any(not isinstance(f, dict) for f in files):
        raise ValueError("Lote de arquivos inválido.")
    if not isinstance(edits, list) or len(edits) > 100_000:
        raise ValueError("Lista de propostas inválida ou excessiva.")
    for e in edits:
        if (not isinstance(e, dict) or not isinstance(e.get("id"), str)
                or not isinstance(e.get("field"), str)
                or (e.get("tag") is not None and not isinstance(e["tag"], str))
                or (e.get("value") is not None and not isinstance(e["value"], str))):
            raise ValueError("Propostas devem usar texto, inclusive true e false nas tags.")
        if len(e.get("value") or "") > 32_000:
            raise ValueError("Um campo proposto excede 32 mil caracteres.")
    if not isinstance(reviews, dict):
        raise ValueError("As revisões devem ser identificadas por objeto.")
    if not isinstance(refs, dict) or set(refs) - REFERENCE_KEYS:
        raise ValueError("Referências aceitas: nm_dom, nm_subdom, nm_sist_orig e ct_area_neg.")
    for values in refs.values():
        if (not isinstance(values, list) or len(values) > 10_000
                or any(not isinstance(v, str) or not v.strip() or len(v) > 250 for v in values)):
            raise ValueError("Listas oficiais devem conter textos não vazios de até 250 caracteres.")


def normalize_reviews(nodes, supplied, issues):
    ids = {n["id"] for n in nodes}
    if set(supplied) - ids:
        raise ValueError("Há revisões de objetos que não pertencem ao lote atual.")
    result = {}
    affected = {i.get("id") for i in issues if i.get("id")}
    now = datetime.now(timezone.utc).isoformat()
    children = {}
    for n in nodes:
        if n["level"] == "Column":
            children.setdefault(n.get("object_id"), set()).add(n["id"])
    for n in nodes:
        raw = supplied.get(n["id"], {})
        if not isinstance(raw, dict):
            raise ValueError("Revisão inválida.")
        status = raw.get("status", "Pending")
        if not isinstance(status, str) or status not in STATUSES or type(raw.get("completed", False)) is not bool:
            raise ValueError("Status ou campo de conclusão inválido.")
        completed = raw.get("completed", False)
        has_issues = bool(({n["id"]} | children.get(n["id"], set())) & affected)
        if status in {"Rejected", "Awaiting clarification"}:
            pass  # A completed review may explicitly reject the object.
        elif completed:
            status = "Validated with remarks" if has_issues or status == "Validated with remarks" or bool(raw.get("pending")) else "Validated"
        elif status in {"Validated", "Validated with remarks"}:
            status = "In review"
        timestamp = raw.get("validated_at") or now
        if not isinstance(timestamp, str):
            raise ValueError("Data da revisão deve ser um texto ISO com data e hora.")
        try:
            datetime.fromisoformat(timestamp.replace("Z", "+00:00"))
        except ValueError:
            raise ValueError("Data da revisão deve ser um texto ISO com data e hora.") from None
        clean = {"status": status, "completed": completed, "authenticated": False,
                 "validated_at": timestamp if completed else ""}
        for key in ("owner", "comments", "pending", "pending_owner", "final_comments"):
            value = raw.get(key, "")
            if not isinstance(value, str) or len(value) > 32_000:
                raise ValueError("Comentário ou responsável inválido.")
            clean[key] = value
        for key in ("tag_comments", "column_comments"):
            value = raw.get(key, {})
            if (not isinstance(value, dict) or len(value) > 10_000 or any(
                    not isinstance(k, str) or not isinstance(v, str) or len(v) > 32_000 for k, v in value.items())):
                raise ValueError("Comentários de revisão inválidos.")
            clean[key] = copy.deepcopy(value)
        result[n["id"]] = clean
    return result


def prepare(files, edits=None, reviews=None, refs=None):
    """Rebuild from originals and explicit edits. Return only in-memory objects."""
    edits = [] if edits is None else edits
    reviews = {} if reviews is None else reviews
    refs = {} if refs is None else refs
    validate_input(files, edits, reviews, refs)
    batch = parse_batch(files, refs)
    report = apply_edits(batch, edits, refs)
    issues = validate_nodes(batch.nodes, refs)
    normalized = normalize_reviews(batch.nodes, reviews, issues)
    objects = [n for n in batch.nodes if n["level"] in {"Table", "View"}]
    progress = sum(normalized[n["id"]]["completed"] for n in objects)
    return {"batch": batch, "nodes": batch.nodes, "report": report, "issues": issues,
            "reviews": normalized, "errors": batch.errors, "duplicates": batch.duplicates,
            "catalog": get_catalog(refs), "summary": {
                "objects": len(objects), "tables": sum(n["level"] == "Table" for n in objects),
                "views": sum(n["level"] == "View" for n in objects),
                "validated": sum(normalized[n["id"]]["status"] == "Validated" for n in objects),
                "with_reservations": sum(normalized[n["id"]]["status"] == "Validated with remarks" for n in objects),
                "rejected": sum(normalized[n["id"]]["status"] == "Rejected" for n in objects),
                "pending": sum(normalized[n["id"]]["status"] in {"Pending", "In review", "Awaiting clarification"} for n in objects),
                "issues": len(issues), "technical_issues": sum(bool(i.get("id")) for i in issues),
                "progress": round(100 * progress / len(objects)) if objects else 0,
                "files_received": len(files), "files_processed": len(batch.documents),
                "files_with_errors": len({e.get("filename") for e in batch.errors}),
            }}


def ensure_exportable(state):
    if not state["nodes"]:
        raise ValueError("Importe ao menos um YAML com objetos reconhecidos.")
    if state["errors"]:
        raise ValueError("Resolva ou remova os arquivos com erro antes de exportar o lote inteiro.")
    if any(r.get("status") == "rejected" for r in state["report"]):
        raise ValueError("Existem propostas rejeitadas. Corrija-as antes de exportar.")


def export_xlsx(files, edits=None, reviews=None, refs=None):
    state = prepare(files, edits, reviews, refs)
    ensure_exportable(state)
    return export_workbook(state["nodes"], state["reviews"], state["issues"], refs or {})


def merge_edits(existing, incoming):
    """Latest explicitly confirmed proposal replaces an earlier session proposal."""
    result = {(e.get("id"), e.get("field"), e.get("tag")): e for e in existing}
    result.update({(e.get("id"), e.get("field"), e.get("tag")): e for e in incoming})
    return list(result.values())


def preview_xlsx(content, files, edits=None, reviews=None, refs=None):
    if not isinstance(content, bytes) or len(content) > 15 * 1024 * 1024:
        raise ValueError("Selecione uma planilha .xlsx de até 15 MiB.")
    state = prepare(files, edits, reviews, refs)
    ensure_exportable(state)
    result = preview_workbook(content, state["nodes"], state["reviews"], refs or {})
    proposed_reviews = {**state["reviews"], **result.get("reviews", {})}
    trial = prepare(files, merge_edits(edits or [], result.get("edits", [])), proposed_reviews, refs)
    result["reviews"] = trial["reviews"]
    result["issues"] = trial["issues"]
    return result


def final_zip(files, edits=None, reviews=None, refs=None, confirmed=False):
    """Only emit complete source files whose objects are all approved and clean.

    A mixed file cannot be split safely without changing the original contract.
    Such files are skipped atomically and explained in the ignored-object report.
    """
    if confirmed is not True:
        raise ValueError("Confirme explicitamente a geração dos YAMLs revisados.")
    state = prepare(files, edits, reviews, refs)
    ensure_exportable(state)
    nodes = state["nodes"]
    objects = [n for n in nodes if n["level"] in {"Table", "View"}]
    by_file = {}
    for n in objects:
        by_file.setdefault(n["filename"], []).append(n)
    affected = {i.get("id") for i in state["issues"] if i.get("id")}
    affected_files = {n["filename"] for n in nodes if n["id"] in affected}
    children_by_object = {}
    for n in nodes:
        if n["level"] == "Column":
            children_by_object.setdefault(n.get("object_id"), set()).add(n["id"])
    eligible, ignored, approved_ids = set(), [], set()
    for filename, group in by_file.items():
        problems = []
        for n in group:
            owned_ids = {n["id"]} | children_by_object.get(n["id"], set())
            reason = []
            review = state["reviews"][n["id"]]
            if review["status"] != "Validated" or not review["completed"]:
                reason.append("Somente objetos com status Validated e revisão concluída entram no pacote final.")
            if owned_ids & affected:
                reason.append("Há inconsistências técnicas ou informações necessárias pendentes no objeto ou nas colunas.")
            if filename in affected_files and not (owned_ids & affected):
                reason.append("Há inconsistências técnicas ou pendências em outro nível deste documento; o arquivo não está tecnicamente conforme.")
            if reason:
                problems.append((n, " ".join(reason)))
        if problems:
            for n in group:
                reason = next((why for obj, why in problems if obj["id"] == n["id"]),
                              "Arquivo contém outro objeto ainda não aprovado; o documento foi preservado integralmente e não foi incluído.")
                ignored.append({"id": n["id"], "object": n["path"], "filename": filename,
                                "status": state["reviews"][n["id"]]["status"], "reason": reason})
        else:
            eligible.add(filename)
            approved_ids.update(n["id"] for n in group)
    for doc in state["batch"].documents:
        if doc.filename not in by_file:
            ignored.append({"filename": doc.filename, "reason": "Documento de referência sem Table/View aprovado para o pacote final."})
    export_batch = copy.copy(state["batch"])
    export_batch.documents = [d for d in state["batch"].documents if d.filename in eligible]
    export_batch.nodes = [n for n in nodes if n["filename"] in eligible]
    exported_ids = {n["id"] for n in export_batch.nodes}
    applied = [r for r in state["report"] if r.get("id") in exported_ids and r.get("status") == "applied"]
    deferred = [r for r in state["report"] if r.get("id") not in exported_ids and r.get("status") == "applied"]
    report = applied + [{"type": "ignored_object", **r} for r in ignored]
    report += [{"type": "pending", **i} for i in state["issues"]]
    report += [{"type": "functional_review", "id": key, **value}
               for key, value in state["reviews"].items() if key in exported_ids]
    content = render_yaml_zip(export_batch, report)
    buffer = BytesIO(content)
    with ZipFile(buffer, "a", compression=ZIP_DEFLATED) as archive:
        archive.writestr("relatorios/alteracoes_aplicadas.json", json.dumps(applied, ensure_ascii=False, indent=2))
        archive.writestr("relatorios/objetos_ignorados_e_pendencias.json", json.dumps({
            "objects": ignored, "pending": state["issues"], "deferred_changes": deferred,
            "files_included": len(eligible), "objects_included": len(approved_ids),
            "note": "Campos não revisados e estruturas originais foram preservados. Categorias não identificadas exigem revisão da Governança."
        }, ensure_ascii=False, indent=2))
    return buffer.getvalue()

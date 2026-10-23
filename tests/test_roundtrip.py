"""Full service workflow using only synthetic in-memory YAML and XLSX bytes."""
import base64
from io import BytesIO
import json
from unittest.mock import patch
from zipfile import ZipFile

import pytest

from app import service
from app.excel import DELETE_TOKEN
from test_excel import edit_field, mutate, row_number


def upload(name, *, suffix="", extra=""):
    text = f'''# comentário sintético preservado
table:
  id: catalogo_teste.silver.{name}
  name: {name}
  description: "Descrição original{suffix}"
  business_name: "Nome original"
  format: delta  # detalhe técnico preservado
  tags:
    fl_mask: "false"
    fl_apt_cons: "false"
  columns:
    - name: coluna_teste
      type: STRING
      description: "Coluna original"
      nullable: false
      tags:
        fl_atr_mask: "false"
        fl_atr_pii: "false"
{extra}'''
    return {"name": name + ".yaml", "content": base64.b64encode(text.encode()).decode()}


def table_id(state, name):
    return next(n["id"] for n in state["nodes"] if n["level"] == "Table" and n["name"] == name)


def test_excel_service_roundtrip_exports_only_validated_files_and_preserves_unknowns():
    files = [upload("aprovada", extra="ingest:\n  marker: preservar\n"), upload("rejeitada"), upload("pendente")]
    state = service.prepare(files)
    approved = table_id(state, "aprovada")
    rejected = table_id(state, "rejeitada")
    content = service.export_xlsx(files, reviews=state["reviews"])
    content = edit_field(content, "Objects", approved, "Descrição proposta", "Descrição aprovada pelo Owner")
    content = edit_field(content, "Approval", approved, "Status final", "Validated")
    content = edit_field(content, "Approval", approved, "Data Owner", "Responsável sintético")
    content = edit_field(content, "Approval", rejected, "Status final", "Rejected")
    content = edit_field(content, "Approval", rejected, "Validation completed", "Sim")
    with patch("tempfile.mkstemp", side_effect=AssertionError("No temporary files")), patch("tempfile.NamedTemporaryFile", side_effect=AssertionError("No temporary files")):
        preview = service.preview_xlsx(content, files, reviews=state["reviews"])
        assert not preview["rejected"] and not preview["missing"]
        result = service.final_zip(files, preview["edits"], preview["reviews"], confirmed=True)
    assert preview["reviews"][approved]["status"] == "Validated"
    assert preview["reviews"][rejected]["status"] == "Rejected"
    with ZipFile(BytesIO(result)) as archive:
        assert {name for name in archive.namelist() if name.endswith(".yaml")} == {"aprovada.yaml"}
        output = archive.read("aprovada.yaml").decode()
        assert "Descrição aprovada pelo Owner" in output
        assert "# comentário sintético preservado" in output
        assert "# detalhe técnico preservado" in output
        assert "marker: preservar" in output
        report = json.loads(archive.read("relatorios/objetos_ignorados_e_pendencias.json"))
    assert report["files_included"] == 1 and report["objects_included"] == 1
    assert {entry["status"] for entry in report["objects"]} == {"Rejected", "Pending"}


def test_service_blank_keeps_existing_and_explicit_delete_changes_yaml():
    files = [upload("campos")]
    state = service.prepare(files)
    obj_id = table_id(state, "campos")
    column = next(n for n in state["nodes"] if n["level"] == "Column")
    content = service.export_xlsx(files, reviews=state["reviews"])
    content = edit_field(content, "Objects", obj_id, "Descrição proposta", None)
    content = edit_field(content, "Tags", column["id"], "Valor proposto", None, "fl_atr_pii")
    content = edit_field(content, "Objects", obj_id, "Nome de negócio proposto", DELETE_TOKEN)
    content = edit_field(content, "Approval", obj_id, "Status final", "Validated")
    preview = service.preview_xlsx(content, files, reviews=state["reviews"])
    assert not preview["rejected"]
    assert preview["edits"] == [{"id": obj_id, "field": "business_name", "value": None}]
    output = service.final_zip(files, preview["edits"], preview["reviews"], confirmed=True)
    with ZipFile(BytesIO(output)) as archive:
        updated = archive.read("campos.yaml").decode()
    assert "business_name:" not in updated
    assert "Descrição original" in updated and 'fl_atr_pii: "false"' in updated


def test_service_preview_rejects_old_fingerprint_even_with_same_object_id():
    files = [upload("versionada")]
    state = service.prepare(files)
    obj_id = table_id(state, "versionada")
    content = edit_field(service.export_xlsx(files), "Objects", obj_id, "Descrição proposta", "Alteração antiga")
    updated_files = [upload("versionada", suffix=" versão 2")]
    preview = service.preview_xlsx(content, updated_files)
    assert not preview["edits"] and any(item["code"] == "immutable" for item in preview["rejected"])


@pytest.mark.parametrize("mode", ["duplicate", "missing"])
def test_service_invalid_duplicate_or_missing_record_never_applies_related_edit(mode):
    files = [upload("duplicada")]
    state = service.prepare(files)
    column = next(n for n in state["nodes"] if n["level"] == "Column")
    content = service.export_xlsx(files)
    content = edit_field(content, "Columns", column["id"], "fl_atr_pii proposto", "true")
    row = row_number(content, "Tags", column["id"], "fl_atr_pii")
    content = mutate(content, "Tags", **{"duplicate_row" if mode == "duplicate" else "delete_row": row})
    preview = service.preview_xlsx(content, files)
    assert not preview["edits"]
    assert preview["rejected"] or preview["missing"]


def test_service_final_package_requires_explicit_confirmation():
    files = [upload("confirmacao")]
    with pytest.raises(ValueError, match="explicitamente"):
        service.final_zip(files)


def test_service_file_with_mixed_approval_is_skipped_atomically():
    source = '''catalog:
  name: catalogo_teste
  schemas:
    - name: silver
      tables:
        - name: aprovada
          description: objeto aprovado
        - name: pendente
          description: objeto pendente
'''
    files = [{"name": "multiplos.yaml", "content": base64.b64encode(source.encode()).decode()}]
    state = service.prepare(files)
    approved = table_id(state, "aprovada")
    reviews = state["reviews"]
    reviews[approved].update({"completed": True, "status": "Validated"})
    output = service.final_zip(files, reviews=reviews, confirmed=True)
    with ZipFile(BytesIO(output)) as archive:
        assert not any(name.endswith(".yaml") for name in archive.namelist())
        report = json.loads(archive.read("relatorios/objetos_ignorados_e_pendencias.json"))
    assert len(report["objects"]) == 2 and report["files_included"] == 0

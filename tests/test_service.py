import base64
from io import BytesIO
import json
from zipfile import ZipFile

import pytest
from openpyxl import load_workbook

from app.service import export_xlsx, final_zip, merge_edits, prepare, preview_xlsx


SOURCE = '''# Contrato sintético
catalogs:
  dev_demo:
    tags: {nm_class_inf: restrita}
    schemas:
      gold:
        tags: {nm_class_inf: restrita}
        tables:
          pedidos:
            description: Descrição inicial
            business_name: Pedidos
            format: iceberg
            custom: preservar
            tags:
              nm_environment: dev
              nm_class_inf: interna
              fl_mask: "false"
              fl_apt_cons: "false"
            columns:
              - name: codigo
                type: STRING
                description: Código do pedido
                tags: {nm_atr_class_inf: interna, fl_atr_mask: "false"}
'''


def upload(text=SOURCE, name="contrato.yaml"):
    return {"name": name, "content": base64.b64encode(text.encode()).decode()}


def test_streamlit_service_complete_roundtrip_and_scope():
    files = [upload()]
    initial = prepare(files)
    obj = next(n for n in initial["nodes"] if n["level"] == "Table")
    edits = [{"id": obj["id"], "field": "description", "value": "Descrição revisada"}]
    reviews = {obj["id"]: {"completed": True, "status": "Validated", "owner": "Pessoa de teste"}}
    reviewed = prepare(files, edits, reviews)
    assert reviewed["reviews"][obj["id"]]["status"] == "Validated"
    assert reviewed["summary"]["validated"] == 1
    # Completing a table never completes its columns or upper levels.
    assert all(not r["completed"] for key, r in reviewed["reviews"].items() if key != obj["id"])
    assert all(not i.get("id") for i in reviewed["issues"])
    book = export_xlsx(files, edits, reviews)
    wb = load_workbook(BytesIO(book))
    assert [s.title for s in wb if s.sheet_state == "visible"] == ["Read Me", "Objects", "Tags", "Columns", "Approval"]
    preview = preview_xlsx(book, files, edits, reviews)
    assert not preview["rejected"]
    assert preview["reviews"][obj["id"]]["status"] == "Validated"
    with pytest.raises(ValueError, match="explicitamente"):
        final_zip(files, edits, reviews)
    result = final_zip(files, edits, reviews, confirmed=True)
    with ZipFile(BytesIO(result)) as z:
        text = z.read("contrato.yaml").decode()
        assert "Descrição revisada" in text
        assert "# Contrato sintético" in text
        assert "custom: preservar" in text
        assert "format: iceberg" in text
        report = json.loads(z.read("relatorios/alteracoes_aplicadas.json"))
        assert report[0]["after"] == "Descrição revisada"


def test_pending_and_rejected_never_become_approved_in_final_package():
    files = [upload()]
    obj = next(n for n in prepare(files)["nodes"] if n["level"] == "Table")
    for status, completed in [("Pending", False), ("Awaiting clarification", False), ("Rejected", True), ("Validated with remarks", True)]:
        reviews = {obj["id"]: {"status": status, "completed": completed}}
        state = prepare(files, reviews=reviews)
        if status == "Rejected":
            assert state["reviews"][obj["id"]]["status"] == "Rejected"
        with ZipFile(BytesIO(final_zip(files, reviews=reviews, confirmed=True))) as z:
            assert not any(name.endswith((".yaml", ".yml")) for name in z.namelist())
            assert json.loads(z.read("relatorios/objetos_ignorados_e_pendencias.json"))["objects"]


def test_technical_errors_are_never_hidden_by_functional_completion():
    files = [upload(SOURCE.replace('fl_mask: "false"', 'fl_mask: "false"').replace('fl_atr_mask: "false"', 'fl_atr_mask: "true"'))]
    obj = next(n for n in prepare(files)["nodes"] if n["level"] == "Table")
    state = prepare(files, reviews={obj["id"]: {"completed": True, "status": "Validated"}})
    assert state["reviews"][obj["id"]]["status"] == "Validated with remarks"
    assert any(i["code"] == "mask_inconsistent" for i in state["issues"])


def test_partial_import_cannot_silently_export_subset():
    files = [upload(), upload("table: [bad", "invalid.yml")]
    state = prepare(files)
    assert state["errors"] and state["nodes"]
    with pytest.raises(ValueError, match="lote inteiro"):
        export_xlsx(files)
    with pytest.raises(ValueError):
        final_zip(files, confirmed=True)


def test_keep_catalog_notes_and_validate_reference_inputs():
    files = [upload()]
    catalog = next(n for n in prepare(files)["nodes"] if n["level"] == "Catalog")
    state = prepare(files, reviews={catalog["id"]: {"tag_comments": {"nm_class_inf": "Conferir limite"}}})
    assert state["reviews"][catalog["id"]]["tag_comments"]["nm_class_inf"] == "Conferir limite"
    with pytest.raises(ValueError):
        prepare(files, refs={"nm_dom": "invalid"})
    with pytest.raises(ValueError):
        prepare(files, reviews={"unknown-id": {}})
    with pytest.raises(ValueError):
        prepare(files, edits=[{"id": "x", "field": "tags", "value": True}])


def test_latest_confirmed_edit_replaces_earlier_proposal():
    old = {"id": "object", "field": "description", "value": "before"}
    new = {**old, "value": "after"}
    assert merge_edits([old], [new]) == [new]

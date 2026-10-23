"""Native Streamlit flows with synthetic uploads and downloads only in RAM."""
import base64
from io import BytesIO
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

from streamlit.testing.v1 import AppTest

from app.service import export_xlsx, prepare, preview_xlsx


APP = Path(__file__).resolve().parents[1] / "streamlit_app.py"
SYNTHETIC = b'''table:
  id: dev_demo.gold.asset
  description: Initial synthetic description
  tags:
    nm_environment: "dev"
  columns:
    - name: code
      type: STRING
      description: Synthetic identifier
      tags:
        fl_atr_pii: "false"
'''


def app():
    return AppTest.from_file(str(APP), default_timeout=30).run()


def button(at, label):
    return next(w for w in at.button if w.label == label)


def checkbox(at, label):
    return next(w for w in at.checkbox if w.label == label)


def seeded():
    at = app()
    files = [{"name": "synthetic.yaml", "content": base64.b64encode(SYNTHETIC).decode("ascii")}]
    prepared = prepare(files)
    at.session_state["mr"].update(files=files, prepared=prepared, reviews=prepared["reviews"])
    at.run()
    assert not at.exception
    return at


def navigate(at, page):
    next(w for w in at.radio if w.label == "Navegação").set_value(page).run()
    assert not at.exception


def revise(at):
    button(at, "Abrir revisão do objeto").click().run()
    assert not at.exception


def workbook_description(content):
    """Edit a generated XLSX without openpyxl.save or temporary files."""
    with ZipFile(BytesIO(content)) as archive:
        entries = {name: archive.read(name) for name in archive.namelist()}
    ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    root = ET.fromstring(entries["xl/worksheets/sheet2.xml"])
    cell = root.find(f".//{ns}c[@r='H5']")
    assert cell is not None
    for child in list(cell):
        cell.remove(child)
    cell.set("t", "inlineStr")
    ET.SubElement(ET.SubElement(cell, ns + "is"), ns + "t").text = "Revised through synthetic Excel"
    entries["xl/worksheets/sheet2.xml"] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    output = BytesIO()
    with ZipFile(output, "w", ZIP_DEFLATED) as archive:
        for name, value in entries.items():
            archive.writestr(name, value)
    return output.getvalue()


def test_empty_app_and_reference_navigation():
    at = app()
    assert not at.exception
    assert any(w.value == "Governança de dados" for w in at.title)
    assert not at.session_state["mr"]["files"]
    navigate(at, "Regras e referências")
    assert any("Catálogo conferido com o PDF" in w.value for w in at.info)
    assert any("Pendência global da Governança" in w.value for w in at.warning)


def test_individual_review_save_and_export_in_memory():
    at = seeded()
    assert any("Pendência global da Governança" in w.value for w in at.warning)
    revise(at)
    next(w for w in at.text_area if w.label == "Descrição proposta").set_value("Synthetic reviewed description")
    checkbox(at, "Validation completed").check()
    button(at, "Salvar revisão em memória").click().run()
    assert not at.exception
    state = at.session_state["mr"]
    table = next(n for n in state["prepared"]["nodes"] if n["level"] == "Table")
    column = next(n for n in state["prepared"]["nodes"] if n["level"] == "Column")
    assert table["description"] == "Synthetic reviewed description"
    assert state["reviews"][table["id"]]["status"] == "Validated"
    assert state["reviews"][column["id"]]["completed"] is False
    navigate(at, "Arquivos e downloads")
    button(at, "Export Excel for Data Owner").click().run()
    assert not at.exception
    assert at.session_state["mr"]["xlsx"].startswith(b"PK")


def test_preview_explicit_apply_zip_confirmation_and_clear():
    at = seeded()
    state = at.session_state["mr"]
    original = export_xlsx(state["files"], state["edits"], state["reviews"], state["refs"])
    revised = workbook_description(original)
    state["preview"] = preview_xlsx(revised, state["files"], state["edits"], state["reviews"], state["refs"])
    navigate(at, "Arquivos e downloads")
    assert button(at, "Aplicar alterações aceitas").disabled
    assert not state["edits"]
    checkbox(at, "Revisei a comparação e autorizo aplicar as alterações aceitas à sessão").check().run()
    button(at, "Aplicar alterações aceitas").click().run()
    assert not at.exception
    state = at.session_state["mr"]
    table = next(n for n in state["prepared"]["nodes"] if n["level"] == "Table")
    assert table["description"] == "Revised through synthetic Excel"
    navigate(at, "Visão geral")
    revise(at)
    checkbox(at, "Validation completed").check()
    button(at, "Salvar revisão em memória").click().run()
    navigate(at, "Arquivos e downloads")
    assert button(at, "Gerar YAMLs atualizados (ZIP)").disabled
    checkbox(at, "Confirmo gerar novas versões apenas dos arquivos elegíveis; os originais não serão sobrescritos").check().run()
    button(at, "Gerar YAMLs atualizados (ZIP)").click().run()
    assert not at.exception
    with ZipFile(BytesIO(at.session_state["mr"]["zip"])) as archive:
        assert "synthetic.yaml" in archive.namelist()
        assert b"Revised through synthetic Excel" in archive.read("synthetic.yaml")
    checkbox(at, "Descartar arquivos e alterações desta sessão").check().run()
    button(at, "Limpar agora").click().run()
    assert not at.exception
    assert at.session_state["mr_epoch"] == 1
    assert not at.session_state["mr"]["files"]
    assert not at.session_state["mr"]["reviews"]
    assert not at.session_state["mr"]["prepared"]


def test_completed_rejection_is_not_promoted_to_validation():
    at = seeded()
    revise(at)
    next(w for w in at.selectbox if w.label == "Status da revisão").set_value("Rejected")
    checkbox(at, "Validation completed").check()
    button(at, "Salvar revisão em memória").click().run()
    assert not at.exception
    state = at.session_state["mr"]
    table = next(n for n in state["prepared"]["nodes"] if n["level"] == "Table")
    assert state["reviews"][table["id"]]["status"] == "Rejected"

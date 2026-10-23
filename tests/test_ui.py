"""Exercise the two native Streamlit conversions using only synthetic bytes."""
import base64
from io import BytesIO
from pathlib import Path
from unittest.mock import Mock, patch
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

import pytest
from streamlit.testing.v1 import AppTest

from app import ui


APP = Path(__file__).resolve().parents[1] / "streamlit_app.py"
SYNTHETIC = b"table:\n  id: demo.gold.asset\n  description: Synthetic\n"


class Upload(BytesIO):
    def __init__(self, name, content):
        super().__init__(content)
        self.name = name
        self.size = len(content)


@pytest.fixture
def session():
    """AppTest does not implement file_uploader; inject framework upload objects."""
    uploads = {"yaml_upload": [], "excel_upload": []}
    excel_result = {"content": b"PK synthetic workbook", "filename": "revisao_owner.xlsx",
                    "summary": {"files": 1, "objects": 1, "columns": 0}, "warnings": []}
    yaml_result = {"content": b"PK synthetic zip", "filename": "yamls_revisados.zip",
                   "yamls": [{"name": "synthetic.yaml", "content": SYNTHETIC}],
                   "summary": {"files": 1, "objects": 1, "changes": 1}, "changes": [], "warnings": []}

    def upload(*args, **kwargs):
        return uploads["yaml_upload" if kwargs["key"].endswith("_yaml_upload") else "excel_upload"]

    excel = Mock(return_value=excel_result)
    yaml = Mock(return_value=yaml_result)
    with patch.object(ui.st, "file_uploader", side_effect=upload), \
            patch.object(ui.service, "create_owner_excel", excel, create=True), \
            patch.object(ui.service, "convert_owner_excels", yaml, create=True):
        at = AppTest.from_file(str(APP), default_timeout=30).run()
        assert not at.exception
        yield at, uploads, excel, yaml


def button(at, label):
    return next(widget for widget in at.button if widget.label == label)


def downloads(at):
    return {widget.label for widget in at.get("download_button")}


def test_only_two_tabs_for_offline_owner_review(session):
    at, _, _, _ = session
    assert [tab.label for tab in at.tabs] == ["1 · YAML para Excel", "2 · Excel para YAML"]
    assert at.title[0].value == "Metadata Review"
    assert not at.radio
    assert not at.checkbox
    assert not at.text_input
    assert not at.text_area
    assert not downloads(at)
    assert at.session_state["mr"] == {"excel": None, "yaml": None}
    assert any("Não é necessário anexar os YAMLs originais" in element.value for element in at.caption)


def test_yaml_and_zip_batch_to_excel_with_explicit_action(session):
    at, uploads, excel, yaml = session
    uploads["yaml_upload"] = [Upload("synthetic.yaml", SYNTHETIC), Upload("batch.zip", b"PK synthetic")]
    at.run()
    excel.assert_not_called()
    button(at, "Gerar Excel para o Owner").click().run()
    assert not at.exception
    assert excel.call_args.args[0] == [
        {"name": "synthetic.yaml", "content": base64.b64encode(SYNTHETIC).decode("ascii")},
        {"name": "batch.zip", "content": base64.b64encode(b"PK synthetic").decode("ascii")},
    ]
    yaml.assert_not_called()
    assert downloads(at) == {"Baixar Excel para o Owner"}
    assert at.session_state["mr"]["excel"]["content"] == b"PK synthetic workbook"
    assert any("Excel pronto" in widget.value for widget in at.success)


def test_governance_converts_excel_without_any_yaml_upload(session):
    at, uploads, excel, yaml = session
    uploads["excel_upload"] = [Upload("owner.xlsx", b"PK synthetic returned workbook")]
    button(at, "Gerar YAMLs").click().run()
    assert not at.exception
    yaml.assert_called_once_with([{"name": "owner.xlsx", "content": b"PK synthetic returned workbook"}])
    excel.assert_not_called()
    assert downloads(at) == {"Baixar YAMLs em ZIP", "Baixar YAML individual"}
    assert at.session_state["mr"]["excel"] is None


def test_multiple_returned_excels_need_no_approval_gate(session):
    at, uploads, _, yaml = session
    uploads["excel_upload"] = [Upload("a.xlsx", b"a"), Upload("b.xlsx", b"b")]
    yaml.return_value = {**yaml.return_value, "yamls": [
        {"name": "a.yaml", "content": b"a"}, {"name": "b.yaml", "content": b"b"}]}
    assert not button(at, "Gerar YAMLs").disabled
    button(at, "Gerar YAMLs").click().run()
    assert not at.exception
    assert len(yaml.call_args.args[0]) == 2
    assert downloads(at) == {"Baixar YAMLs em ZIP"}


def test_new_source_or_failed_conversion_removes_stale_download(session):
    at, uploads, excel, _ = session
    uploads["yaml_upload"] = [Upload("a.yaml", SYNTHETIC)]
    button(at, "Gerar Excel para o Owner").click().run()
    assert downloads(at) == {"Baixar Excel para o Owner"}
    uploads["yaml_upload"] = [Upload("a.yaml", SYNTHETIC + b"# changed\n")]
    at.run()
    assert not downloads(at)
    assert at.session_state["mr"]["excel"] is None
    excel.side_effect = ValueError("Formato de YAML inválido.")
    button(at, "Gerar Excel para o Owner").click().run()
    assert not at.exception
    assert not downloads(at)
    assert any("Formato de YAML" in error.value for error in at.error)


def test_changed_workbook_removes_old_yamls_and_preserves_other_tab(session):
    at, uploads, _, _ = session
    uploads["yaml_upload"] = [Upload("synthetic.yaml", SYNTHETIC)]
    button(at, "Gerar Excel para o Owner").click().run()
    uploads["excel_upload"] = [Upload("owner.xlsx", b"before")]
    button(at, "Gerar YAMLs").click().run()
    assert len(downloads(at)) == 3
    uploads["excel_upload"] = [Upload("owner.xlsx", b"after")]
    at.run()
    assert not at.exception
    assert downloads(at) == {"Baixar Excel para o Owner"}
    assert at.session_state["mr"]["yaml"] is None


def test_excessive_batch_is_refused_before_conversion(session):
    at, uploads, excel, _ = session
    oversized = Upload("large.zip", b"synthetic")
    oversized.size = ui.MAX_FILE_BYTES + 1
    uploads["yaml_upload"] = [oversized]
    at.run()
    assert not at.exception
    assert button(at, "Gerar Excel para o Owner").disabled
    assert at.error
    excel.assert_not_called()


def test_errors_do_not_render_imported_markdown_or_unexpected_exception(session):
    at, uploads, excel, _ = session
    uploads["yaml_upload"] = [Upload("synthetic.yaml", SYNTHETIC)]
    excel.side_effect = ValueError("![metadata](https://example.invalid/pixel)")
    button(at, "Gerar Excel para o Owner").click().run()
    assert not at.exception
    assert "\\!\\[metadata\\]\\(https://example\\.invalid/pixel\\)" in at.error[0].value
    excel.side_effect = RuntimeError("sensitive internal diagnostic")
    button(at, "Gerar Excel para o Owner").click().run()
    assert not at.exception
    assert "sensitive" not in at.error[0].value
    assert not downloads(at)


def test_optional_warnings_remain_collapsed_and_do_not_block_download(session):
    at, uploads, excel, _ = session
    uploads["yaml_upload"] = [Upload("synthetic.yaml", SYNTHETIC)]
    excel.return_value = {**excel.return_value, "warnings": [
        {"message": "Confira a tag ![x](https://example.invalid/a)", "filename": "source.yaml"}]}
    button(at, "Gerar Excel para o Owner").click().run()
    assert not at.exception
    assert "Baixar Excel para o Owner" in downloads(at)
    assert any(widget.label == "Pontos para Governança · 1" for widget in at.expander)
    assert "\\!\\[x\\]" in at.warning[0].value


def test_missing_uploads_are_actionable_and_session_reset_releases_bytes(session):
    at, uploads, excel, yaml = session
    button(at, "Gerar Excel para o Owner").click().run()
    button(at, "Gerar YAMLs").click().run()
    excel.assert_not_called()
    yaml.assert_not_called()
    assert at.warning
    uploads["yaml_upload"] = [Upload("synthetic.yaml", SYNTHETIC)]
    button(at, "Gerar Excel para o Owner").click().run()
    with patch.object(ui, "release_session_files") as release:
        button(at, "Limpar sessão").click().run()
        release.assert_called_once()
    assert not at.exception
    assert at.session_state["mr_epoch"] == 1
    assert at.session_state["mr"] == {"excel": None, "yaml": None}
    assert not downloads(at)


def test_legacy_in_app_review_session_is_discarded(session):
    at, _, _, _ = session
    at.session_state["mr"] = {"files": [{"name": "legacy.yaml", "content": "sensitive"}], "reviews": {}}
    with patch.object(ui, "release_session_files") as release:
        at.run()
        release.assert_called_once()
    assert not at.exception
    assert at.session_state["mr"] == {"excel": None, "yaml": None}


def test_real_ui_roundtrip_after_clearing_originals():
    uploads = {"yaml_upload": [Upload("synthetic.yaml", SYNTHETIC)], "excel_upload": []}

    def upload(*args, **kwargs):
        return uploads["yaml_upload" if kwargs["key"].endswith("_yaml_upload") else "excel_upload"]

    with patch.object(ui.st, "file_uploader", side_effect=upload):
        at = AppTest.from_file(str(APP), default_timeout=30).run()
        button(at, "Gerar Excel para o Owner").click().run()
        assert not at.exception
        content = at.session_state["mr"]["excel"]["content"]
        # Simulate an Owner's permitted description edit without saving a file.
        with ZipFile(BytesIO(content)) as workbook:
            entries = {name: workbook.read(name) for name in workbook.namelist()}
        namespace = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
        book = ET.fromstring(entries["xl/workbook.xml"])
        index = next(i + 1 for i, item in enumerate(book.find(namespace + "sheets")) if item.get("name") == "Objetos")
        sheet_path = f"xl/worksheets/sheet{index}.xml"
        sheet = ET.fromstring(entries[sheet_path])
        cell = sheet.find(f".//{namespace}c[@r='H5']")
        assert cell is not None
        for child in list(cell):
            cell.remove(child)
        cell.set("t", "inlineStr")
        ET.SubElement(ET.SubElement(cell, namespace + "is"), namespace + "t").text = "Owner proposed description"
        entries[sheet_path] = ET.tostring(sheet, encoding="utf-8", xml_declaration=True)
        edited = BytesIO()
        with ZipFile(edited, "w", ZIP_DEFLATED) as workbook:
            for name, value in entries.items():
                workbook.writestr(name, value)
        uploads["yaml_upload"] = []
        button(at, "Limpar sessão").click().run()
        assert at.session_state["mr"] == {"excel": None, "yaml": None}
        uploads["excel_upload"] = [Upload("owner_revised.xlsx", edited.getvalue())]
        button(at, "Gerar YAMLs").click().run()
        assert not at.exception
        assert not at.error
        result = at.session_state["mr"]["yaml"]
        assert result["summary"]["changes"] == 1
        assert b"Owner proposed description" in result["yamls"][0]["content"]
        with ZipFile(BytesIO(result["content"])) as archive:
            assert archive.namelist() == ["synthetic.yaml"]

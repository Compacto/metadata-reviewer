"""Adversarial and memory-only checks, using synthetic metadata exclusively."""
import base64
import io
import json
import zipfile
from contextlib import ExitStack
from pathlib import Path
from unittest.mock import patch

import pytest

from app import engine, rules
from app.excel import export_workbook, preview_workbook


SAFE = "table: {id: 'catalogo_teste.gold.objeto_teste', description: 'Somente teste', tags: {fl_mask: 'false'}}\n"


def upload(raw, name="sintetico.yaml"):
    if isinstance(raw, str):
        raw = raw.encode("utf-8")
    return {"name": name, "content": base64.b64encode(raw).decode("ascii")}


def zipped(entries):
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, value in entries:
            archive.writestr(name, value)
    return upload(buffer.getvalue(), "sintetico.zip")


def test_complete_processing_never_opens_disk_files_or_temporary_storage():
    packed = zipped([("contratos/sintetico.yaml", SAFE)])
    # Imports happen above; all operations that can create a file are denied
    # throughout parsing, Excel writing/reading and updated YAML ZIP generation.
    with ExitStack() as stack:
        for target in ("builtins.open", "tempfile.NamedTemporaryFile", "tempfile.TemporaryFile", "tempfile.mkstemp"):
            stack.enter_context(patch(target, side_effect=AssertionError("Persistent I/O attempted")))
        stack.enter_context(patch.object(Path, "write_bytes", side_effect=AssertionError("Persistent write attempted")))
        stack.enter_context(patch.object(Path, "write_text", side_effect=AssertionError("Persistent write attempted")))
        batch = engine.parse_batch([packed])
        assert not batch.errors
        issues = rules.validate_nodes(batch.nodes)
        workbook = export_workbook(batch.nodes, {}, issues)
        preview = preview_workbook(workbook, batch.nodes, {})
        assert not preview["rejected"]
        report = engine.apply_edits(batch, [{"id": batch.nodes[0]["id"], "field": "description", "value": "Texto revisado de teste"}])
        result = engine.render_yaml_zip(batch, report)
        assert isinstance(workbook, bytes) and isinstance(result, bytes)
        with zipfile.ZipFile(io.BytesIO(result)) as archive:
            assert "Texto revisado de teste" in archive.read("sintetico/contratos/sintetico.yaml").decode()


@pytest.mark.parametrize("source", [
    "table:\n  id: 'catalogo_teste.gold.objeto_teste'\n  tags:\n    CONTEUDO_PRIVADO_SENTINELA: [",
    "table: {id: 'catalogo_teste.gold.objeto_teste', segredo: CONTEUDO_PRIVADO_SENTINELA, segredo: outro}",
    "table: !!python/object/apply:os.system ['CONTEUDO_PRIVADO_SENTINELA']",
    "table: {id: 'catalogo_teste.gold.objeto_teste', unknown: !!int CONTEUDO_PRIVADO_SENTINELA}",
    "table: {id: 'catalogo_teste.gold.objeto_teste', unknown: !!float CONTEUDO_PRIVADO_SENTINELA}",
    b"table: \xffCONTEUDO_PRIVADO_SENTINELA",
])
def test_parser_diagnostics_never_echo_input_content(source):
    batch = engine.parse_batch([upload(source)])
    assert batch.errors
    assert "CONTEUDO" not in json.dumps(batch.errors)
    assert not batch.documents


@pytest.mark.parametrize("name", ["../objeto.yaml", "/objeto.yaml", "C:/objeto.yaml", "pasta/../objeto.yaml", "NUL.yaml", "pasta./objeto.yaml", "objeto\n.yaml"])
def test_unsafe_direct_paths_are_rejected_without_any_extraction(name):
    batch = engine.parse_batch([upload(SAFE, name)])
    assert batch.errors
    assert not batch.nodes


def test_oversized_file_does_not_corrupt_an_accepted_file(monkeypatch):
    monkeypatch.setattr(engine, "MAX_FILE_BYTES", len(SAFE.encode()))
    batch = engine.parse_batch([upload(SAFE), upload(SAFE + "# excesso\n", "grande.yaml")])
    assert len(batch.documents) == 1
    assert len(batch.nodes) == 1
    assert any(error["filename"] == "grande.yaml" for error in batch.errors)


def test_batch_graph_budget_applies_across_files_without_discarding_accepted_objects(monkeypatch):
    monkeypatch.setattr(engine, "MAX_NODES", 20)
    many_values = "table: {id: 'catalogo_teste.gold.grande', extra: [" + ",".join("x" for _ in range(30)) + "]}"
    batch = engine.parse_batch([upload(many_values, "grande.yaml"), upload(SAFE)])
    assert batch.errors
    assert len(batch.documents) == 1
    assert batch.nodes[0]["name"] == "objeto_teste"


def test_zip_entries_include_directories_in_the_configurable_limit(monkeypatch):
    packed = zipped([("um/", ""), ("dois/", ""), ("sintetico.yaml", SAFE)])
    monkeypatch.setattr(engine, "MAX_ZIP_ENTRIES", 2)
    batch = engine.parse_batch([packed])
    assert batch.errors
    assert not batch.documents


def test_high_expansion_zip_is_rejected_before_reading_payload():
    packed = zipped([("explosao.yaml", "a" * (1024 * 1024 + 1))])
    with patch.object(zipfile.ZipFile, "read", side_effect=AssertionError("Payload decompressed before preflight")):
        batch = engine.parse_batch([packed])
    assert batch.errors
    assert "compactação" in batch.errors[0]["message"]


def test_cooperative_deadline_stops_work_with_safe_diagnostic():
    with patch.object(engine.time, "monotonic", side_effect=[0.0] + [1000.0] * 20):
        batch = engine.parse_batch([upload(SAFE)])
    assert not batch.documents
    assert "Tempo máximo" in batch.errors[0]["message"]


def test_configuration_accepts_numeric_env_only_and_rejects_unbounded_values(monkeypatch):
    monkeypatch.setenv("METADATA_MAX_FILES", "12")
    assert engine._setting("MAX_FILES", 500, 1, 5000) == 12
    for value in ("0", "99999999999999999999999", "c:/nao_ler_arquivo"):
        monkeypatch.setenv("METADATA_MAX_FILES", value)
        with pytest.raises(ValueError, match="Configuração METADATA_MAX_FILES inválida"):
            engine._setting("MAX_FILES", 500, 1, 5000)


def test_unknown_native_boolean_is_a_technical_error_and_explicit_null_deletion_works():
    batch = engine.parse_batch([upload("table: {id: 'catalogo_teste.gold.objeto_teste', tags: {tag_extra: true, nm_dom: null}}")])
    assert any(issue["code"] == "native_boolean" for issue in rules.validate_nodes(batch.nodes))
    nid = batch.nodes[0]["id"]
    report = engine.apply_edits(batch, [{"id": nid, "field": "tags", "tag": "nm_dom", "value": None}])
    assert report[0]["status"] == "applied"
    assert "nm_dom" not in batch.nodes[0]["tags"]
    assert "nm_dom" in batch.nodes[0]["original_tags"]


def test_unquoted_date_tag_requires_an_explicit_string_confirmation():
    batch = engine.parse_batch([upload("table: {id: 'catalogo_teste.gold.objeto_teste', tags: {dt_aprv_gov: 2026-10-23}}")])
    node = batch.nodes[0]
    assert node["tags"]["dt_aprv_gov"] == "2026-10-23"
    assert any(issue["code"] == "native_tag_type" for issue in rules.validate_nodes(batch.nodes))
    report = engine.apply_edits(batch, [{"id": node["id"], "field": "tags", "tag": "dt_aprv_gov", "value": "2026-10-23"}])
    assert report[0]["status"] == "applied"
    assert not any(issue["code"] == "native_tag_type" for issue in rules.validate_nodes(batch.nodes))
    with zipfile.ZipFile(io.BytesIO(engine.render_yaml_zip(batch, report))) as archive:
        reparsed = engine.parse_batch([upload(archive.read("sintetico.yaml"))])
    assert not reparsed.nodes[0]["_native_tag_types"]


def test_governance_category_gap_is_visible_separately_from_technical_approval():
    issues = rules.validate_nodes(engine.parse_batch([upload(SAFE)]).nodes)
    category = [issue for issue in issues if issue["code"] == "category_not_identified"]
    assert len(category) == 1
    assert category[0]["id"] == "" and category[0]["scope"] == "governance"
    assert not [issue for issue in issues if issue["id"]]


def test_malformed_edit_identifiers_are_rejected_without_uncaught_errors():
    batch = engine.parse_batch([upload(SAFE)])
    report = engine.apply_edits(batch, [{"id": [], "field": "description", "value": "x"}, {"id": batch.nodes[0]["id"], "field": {}, "value": "x"}])
    assert all(row["status"] == "rejected" for row in report)

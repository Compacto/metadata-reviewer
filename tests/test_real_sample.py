"""Opt-in real contract validation. No corporate fixture or result is persisted.

Set METADATA_SAMPLE_YAML to the external sample supplied by the user. Only its
bytes are read; all workbook manipulation and exported contracts stay in RAM.
"""
from __future__ import annotations

from copy import deepcopy
import os
from pathlib import Path

import pytest

from app import service
from app.engine import parse_batch
from test_conversion import (
    assert_same_documents,
    convert,
    edit_owner_field,
    output_files,
    upload,
    make_zip,
)


@pytest.fixture
def real_sample():
    configured = os.environ.get("METADATA_SAMPLE_YAML")
    if not configured:
        pytest.skip("Set METADATA_SAMPLE_YAML to exercise the external user sample.")
    path = Path(configured)
    assert path.is_file(), "The externally configured sample is unavailable."
    return {"name": path.name, "raw": path.read_bytes()}


def test_real_sample_owner_hand_off_preserves_whole_contract_and_comments(real_sample):
    files = [upload(real_sample["raw"], real_sample["name"])]
    excel = service.create_owner_excel(files)
    result = convert(excel["content"])
    updated = output_files(result)
    assert set(updated) == {real_sample["name"]}
    assert_same_documents(files, updated)
    source = real_sample["raw"].decode("utf-8-sig")
    exported = updated[real_sample["name"]].decode("utf-8")
    comments = [line.strip() for line in source.splitlines() if line.lstrip().startswith("#")]
    assert all(comment in exported for comment in comments)
    assert result["changes"] == []


def test_real_sample_owner_edits_only_requested_description_column_and_tag(real_sample):
    files = [upload(real_sample["raw"], real_sample["name"])]
    original = parse_batch(files)
    assert not original.errors
    table = next(node for node in original.nodes if node["level"] in {"Table", "View"})
    column = next(node for node in original.nodes if node["level"] == "Column")
    excel = service.create_owner_excel(files)["content"]
    excel = edit_owner_field(excel, sheet="Objetos", name=table["name"], value="Descrição revisada pelo Owner no teste local")
    excel = edit_owner_field(excel, sheet="Colunas", name=column["name"], value="Descrição da coluna revisada pelo Owner no teste local")
    replacement = "false" if table.get("tags", {}).get("fl_metadata_ai_generate") == "true" else "true"
    excel = edit_owner_field(excel, sheet="Tags", name=table["path"], field="tags", tag="fl_metadata_ai_generate", value=replacement)
    result = convert(excel)
    updated = output_files(result)
    after = parse_batch([upload(raw, name) for name, raw in updated.items()])
    assert not after.errors
    before_by_id = {node["id"]: deepcopy(node) for node in original.nodes}
    after_by_id = {node["id"]: node for node in after.nodes}
    assert before_by_id.keys() == after_by_id.keys()
    assert after_by_id[table["id"]]["description"] == "Descrição revisada pelo Owner no teste local"
    assert after_by_id[column["id"]]["description"] == "Descrição da coluna revisada pelo Owner no teste local"
    assert after_by_id[table["id"]]["tags"]["fl_metadata_ai_generate"] == replacement
    for node_id, before in before_by_id.items():
        after_node = after_by_id[node_id]
        assert before["path"] == after_node["path"] and before["level"] == after_node["level"]
        expected_tags = dict(before["tags"])
        if node_id == table["id"]:
            expected_tags["fl_metadata_ai_generate"] = replacement
        assert after_node["tags"] == expected_tags
        assert after_node["business_name"] == before["business_name"]
        if node_id not in {table["id"], column["id"]}:
            assert after_node["description"] == before["description"]


def test_real_sample_inside_zip_roundtrips_without_original_upload(real_sample):
    packed = make_zip([(f"contratos/{real_sample['name']}", real_sample["raw"])])
    files = [upload(packed, "lote_exemplo.zip")]
    excel = service.create_owner_excel(files)
    result = convert(excel["content"])
    updated = output_files(result)
    assert set(updated) == {f"lote_exemplo/contratos/{real_sample['name']}"}
    assert_same_documents(files, updated)
    assert result["changes"] == []

"""Column governance tags are reviewable in both friendly workbook views."""
from io import BytesIO
from unittest.mock import patch
from zipfile import ZipFile

import openpyxl
import pytest

from app import engine, owner_workbook as ow, rules
from test_owner_workbook import SOURCE, edit, find_row, header_col, upload, workbook


DOCUMENTED_COLUMN_TAGS = {
    "fl_reg_apl", "fl_atr_pii", "nm_atr_pii", "fl_atr_sens",
    "fl_atr_mask", "nm_atr_class_inf", "fl_metadata_ai_generate",
}
ALL_COLUMN_TAGS_SOURCE = SOURCE.replace(
    'tags: {fl_atr_pii: "false", nm_atr_class_inf: interna, personalizada: original}',
    'tags: {fl_reg_apl: "false", fl_atr_pii: "false", nm_atr_pii: nao_aplicavel, '
    'fl_atr_sens: "false", fl_atr_mask: "false", nm_atr_class_inf: interna, '
    'fl_metadata_ai_generate: "false", personalizada: original}',
)
COLUMN_PATH = "exemplo.silver.vendas.codigo"


def tags_rows(raw, path=None):
    book = openpyxl.load_workbook(BytesIO(raw), data_only=False)
    try:
        sheet = book["Tags"]
        headers = [cell.value for cell in sheet[ow.HEADER_ROW + 1]]
        rows = [dict(zip(headers, [cell.value for cell in row])) | {"row": row[0].row}
                for row in sheet.iter_rows(min_row=ow.HEADER_ROW + 2)
                if row[0].value == "Coluna"]
        return [row for row in rows if path is None or row["Objeto"] == path]
    finally:
        book.close()


def tag_row(raw, tag, path=COLUMN_PATH):
    return next(row["row"] for row in tags_rows(raw, path) if row["Tag técnica"] == tag)


def proposed_tag(raw, tag, value):
    column = header_col(raw, "Tags", "Valor proposto")
    row = tag_row(raw, tag)
    return edit(raw, "Tags", {f"{column}{row}": value})


def resulting_nodes(result):
    batch = engine.parse_batch(result["files"])
    report = engine.apply_edits(batch, result["edits"])
    assert all(item["status"] == "applied" for item in report)
    with ZipFile(BytesIO(engine.render_yaml_zip(batch, report))) as archive:
        files = [{"name": name, "content": upload(archive.read(name).decode())["content"]}
                 for name in archive.namelist() if name.endswith((".yaml", ".yml"))]
    regenerated = engine.parse_batch(files)
    assert not regenerated.errors
    return regenerated.nodes


def test_tags_sheet_lists_all_seven_documented_column_tags_with_column_guidance():
    raw = workbook(ALL_COLUMN_TAGS_SOURCE)
    rows = tags_rows(raw, COLUMN_PATH)
    by_tag = {row["Tag técnica"]: row for row in rows}
    assert set(by_tag) == DOCUMENTED_COLUMN_TAGS | {"personalizada"}
    assert len(rows) == len(by_tag)
    catalog = rules.get_catalog()["Column"]
    for tag in DOCUMENTED_COLUMN_TAGS:
        row = by_tag[tag]
        assert row["Tipo de objeto"] == "Coluna"
        assert row["Informação de governança"] == ow.TAG_LABELS[tag]
        assert row["O que significa"] == catalog[tag]["description"]
        assert row["Valores aceitos / orientação"] == ow._allowed(catalog[tag])
        assert row["Valor proposto"] in (None, "")
    assert by_tag["fl_atr_pii"]["Valor atual"] == "Não"
    assert by_tag["nm_atr_pii"]["Valor atual"] == "nao_aplicavel"
    assert by_tag["personalizada"]["Valor atual"] == "original"


def test_only_existing_column_tags_become_tags_sheet_rows():
    raw = workbook()
    actual = {(row["Objeto"], row["Tag técnica"]) for row in tags_rows(raw)}
    assert actual == {
        (COLUMN_PATH, "fl_atr_pii"),
        (COLUMN_PATH, "nm_atr_class_inf"),
        (COLUMN_PATH, "personalizada"),
        ("exemplo.silver.vendas.observacao", "fl_atr_mask"),
    }
    result = ow.read_owner_workbook(raw)
    assert result["files"] == [upload()]
    assert result["edits"] == []


@pytest.mark.parametrize("proposal,expected", [
    ("Sim", "true"), ("true", "true"), ("Não", "false"), ("false", "false"),
])
def test_boolean_edit_from_tags_roundtrips_as_column_string_without_changing_table(proposal, expected):
    source = ALL_COLUMN_TAGS_SOURCE
    if expected == "false":
        source = source.replace('fl_atr_pii: "false"', 'fl_atr_pii: "true"')
    raw = proposed_tag(workbook(source), "fl_atr_pii", proposal)
    result = ow.read_owner_workbook(raw)
    original = engine.parse_batch([upload(source)]).nodes
    column = next(node for node in original if node["path"] == COLUMN_PATH)
    assert result["edits"] == [{"id": column["id"], "field": "tags",
                                "tag": "fl_atr_pii", "value": expected}]
    nodes = resulting_nodes(result)
    revised = next(node for node in nodes if node["path"] == COLUMN_PATH)
    assert revised["tags"]["fl_atr_pii"] == expected
    assert type(revised["tags"]["fl_atr_pii"]) is str
    table = next(node for node in nodes if node["level"] == "Table")
    assert table["tags"] == next(node["tags"] for node in original if node["level"] == "Table")
    assert "fl_atr_pii" not in table["tags"]


def test_unknown_existing_column_tag_can_be_reviewed_in_tags_sheet():
    result = ow.read_owner_workbook(proposed_tag(workbook(), "personalizada", "revisada"))
    assert len(result["edits"]) == 1
    assert result["edits"][0]["tag"] == "personalizada"
    nodes = resulting_nodes(result)
    assert next(node for node in nodes if node["path"] == COLUMN_PATH)["tags"]["personalizada"] == "revisada"
    assert any(warning["code"] == "unknown_tag" and warning["source"] == "proposal"
               for warning in result["warnings"])


@pytest.mark.parametrize("columns_proposal,tags_proposal,expected", [
    ("Sim", "Sim", "true"),
    ("Sim", "true", "true"),
    (ow.DELETE_TOKEN, ow.DELETE_TOKEN, None),
])
def test_same_column_proposal_in_both_views_is_applied_once(columns_proposal, tags_proposal, expected):
    raw = workbook()
    row = find_row(raw, "Colunas", "Coluna", "codigo")
    column = header_col(raw, "Colunas", "Dado pessoal — proposto")
    raw = edit(raw, "Colunas", {f"{column}{row}": columns_proposal})
    raw = proposed_tag(raw, "fl_atr_pii", tags_proposal)
    result = ow.read_owner_workbook(raw)
    assert len(result["edits"]) == 1
    assert result["edits"][0]["tag"] == "fl_atr_pii"
    assert result["edits"][0]["value"] == expected
    node = next(node for node in resulting_nodes(result) if node["path"] == COLUMN_PATH)
    if expected is None:
        assert "fl_atr_pii" not in node["tags"]
    else:
        assert node["tags"]["fl_atr_pii"] == expected


@pytest.mark.parametrize("columns_proposal,tags_proposal", [
    ("Sim", "Não"),  # An explicit proposal retaining the old value still conflicts.
    ("Não", "Sim"),
    (ow.DELETE_TOKEN, "Sim"),
])
def test_conflicting_column_proposals_reject_import_and_identify_both_cells(columns_proposal, tags_proposal):
    raw = workbook()
    row = find_row(raw, "Colunas", "Coluna", "codigo")
    column = header_col(raw, "Colunas", "Dado pessoal — proposto")
    column_cell = f"{column}{row}"
    tags_cell = f'{header_col(raw, "Tags", "Valor proposto")}{tag_row(raw, "fl_atr_pii")}'
    raw = edit(raw, "Colunas", {column_cell: columns_proposal})
    raw = edit(raw, "Tags", {tags_cell: tags_proposal})
    with pytest.raises(ValueError) as error:
        ow.read_owner_workbook(raw)
    message = str(error.value)
    assert "Colunas" in message and column_cell in message
    assert "Tags" in message and tags_cell in message


@pytest.mark.parametrize("original_value", [
    "true", "17", "null", "[um, dois]", "{um: dois}", '"valor_fora_da_lista"',
])
def test_blank_column_tag_proposal_preserves_malformed_original_values(original_value):
    source = SOURCE.replace('fl_atr_pii: "false"', f"fl_atr_pii: {original_value}")
    raw = workbook(source)
    assert tag_row(raw, "fl_atr_pii")
    result = ow.read_owner_workbook(raw)
    assert result["edits"] == []
    assert result["files"] == [upload(source)]


@pytest.mark.parametrize("version", ["metadata-reviewer-owner-v1", "metadata-reviewer-owner-v2"])
def test_pre_column_tags_workbook_versions_still_accept_existing_column_edits(version):
    with patch.object(ow, "VERSION", version):
        raw = workbook()
    assert tags_rows(raw) == []
    row = find_row(raw, "Colunas", "Coluna", "codigo")
    column = header_col(raw, "Colunas", "Dado pessoal — proposto")
    result = ow.read_owner_workbook(edit(raw, "Colunas", {f"{column}{row}": "Sim"}))
    assert len(result["edits"]) == 1
    assert result["edits"][0]["tag"] == "fl_atr_pii"
    assert result["edits"][0]["value"] == "true"
    nodes = resulting_nodes(result)
    assert next(node for node in nodes if node["path"] == COLUMN_PATH)["tags"]["fl_atr_pii"] == "true"

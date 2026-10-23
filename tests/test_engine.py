import base64
import io
import json
import stat
import unittest
import zipfile

from app.engine import apply_edits, parse_batch, render_yaml_zip
from app.rules import get_catalog, validate_nodes


def upload(text, name="contrato.yaml"):
    return {"name": name, "content": base64.b64encode(text.encode("utf-8")).decode("ascii")}


def zip_upload(entries, name="lote.zip"):
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for filename, content in entries:
            archive.writestr(filename, content)
    return {"name": name, "content": base64.b64encode(buffer.getvalue()).decode("ascii")}


CONTRACT = '''# comentário inicial
table:
  id: "dev_exemplo.gold.contrato_sintetico"
  name: "contrato_sintetico"
  description: |
    Descrição original.
  business_name: "Contrato"
  format: iceberg  # informação técnica preservada
  tags:
    nm_environment: "dev"
    nm_dom: "dominio_sintetico"
    nm_class_inf: "restrita"
    fl_mask: "true"
    ct_qld_status: "em_avaliacao"
    fl_apt_cons: "false"
    tag_nao_documentada: "preservar"
  columns:
    - name: "id_sintetico"
      type: "STRING"
      nullable: false
      description: "Identificador de teste."
      tags:
        fl_atr_mask: "true"
        nm_atr_class_inf: "restrita"
ingest:
  job:
    entry_point: nao_executar
'''


class EngineTests(unittest.TestCase):
    def test_roundtrip_preserves_unknown_fields_comments_and_originals(self):
        batch = parse_batch([upload(CONTRACT)])
        self.assertEqual([], batch.errors)
        table, column = batch.nodes
        self.assertEqual("Table", table["level"])
        self.assertEqual(table["id"], column["object_id"])
        report = apply_edits(batch, [
            {"id": table["id"], "field": "description", "value": "Descrição revisada.\n"},
            {"id": column["id"], "field": "tags", "tag": "fl_atr_mask", "value": "false"},
        ])
        self.assertTrue(all(r["status"] == "applied" for r in report))
        self.assertEqual("true", column["original_tags"]["fl_atr_mask"])
        with self.assertRaises(TypeError):
            column["original_tags"]["fl_atr_mask"] = "false"
        output = render_yaml_zip(batch, report)
        with zipfile.ZipFile(io.BytesIO(output)) as archive:
            result = archive.read("contrato.yaml").decode("utf-8")
            self.assertIn("# comentário inicial", result)
            self.assertIn("# informação técnica preservada", result)
            self.assertIn("entry_point: nao_executar", result)
            self.assertIn('tag_nao_documentada: "preservar"', result)
            self.assertIn('fl_atr_mask: "false"', result)
            self.assertEqual(report, json.loads(archive.read("report.json")))
        self.assertEqual("true", parse_batch([upload(CONTRACT)]).nodes[1]["tags"]["fl_atr_mask"])

    def test_stable_ids_and_source_fingerprint_independent_of_filename(self):
        a = parse_batch([upload(CONTRACT, "a.yaml")]).nodes
        b = parse_batch([upload(CONTRACT, "b.yaml")]).nodes
        self.assertEqual([n["id"] for n in a], [n["id"] for n in b])
        self.assertEqual(a[0]["source_sha256"], b[0]["source_sha256"])
        changed = parse_batch([upload(CONTRACT.replace("Descrição original.", "Outro texto."))]).nodes
        self.assertNotEqual(a[0]["source_sha256"], changed[0]["source_sha256"])

    def test_hierarchy_and_standalone_levels(self):
        text = '''catalogs:
  dev_teste:
    tags: {nm_environment: dev}
    schemas:
      gold:
        tables:
          tabela:
            columns:
              - name: coluna
        views:
          visao:
            columns:
              - name: coluna_view
'''
        batch = parse_batch([upload(text), upload("schema: {id: 'outro.gold', description: 'Camada'}", "schema.yml")])
        self.assertEqual([], batch.errors)
        levels = [n["level"] for n in batch.nodes]
        self.assertEqual(["Catalog", "Schema", "Table", "Column", "View", "Column", "Schema"], levels)
        self.assertEqual(batch.nodes[4]["id"], batch.nodes[5]["object_id"])
        self.assertEqual("dev_teste.gold.visao.coluna_view", batch.nodes[5]["path"])

    def test_unrecognized_and_ambiguous_structure_is_reported(self):
        for text in ("qualquer: {coisa: valor}", "schema: {name: gold}", "table: {id: 'c.s.t', name: outra}"):
            batch = parse_batch([upload(text)])
            self.assertTrue(batch.errors)
            self.assertFalse(batch.nodes)

    def test_native_booleans_invalid_and_string_edit_fixes(self):
        batch = parse_batch([upload(CONTRACT.replace('fl_mask: "true"', "fl_mask: true"))])
        self.assertTrue(any(i["code"] == "native_boolean" for i in validate_nodes(batch.nodes)))
        node = batch.nodes[0]
        report = apply_edits(batch, [{"id": node["id"], "field": "tags", "tag": "fl_mask", "value": "true"}])
        self.assertEqual("applied", report[0]["status"])
        self.assertFalse(any(i["code"] == "native_boolean" for i in validate_nodes(batch.nodes)))
        self.assertIs(True, node["original_tags"]["fl_mask"])

    def test_external_lists_unknown_tags_and_level_differences(self):
        batch = parse_batch([upload(CONTRACT)])
        issues = validate_nodes(batch.nodes)
        self.assertTrue(any(i["message"] == "Lista oficial não carregada" for i in issues))
        self.assertTrue(any(i["code"] == "unknown_tag" and i["severity"] == "pending" for i in issues))
        issues = validate_nodes(batch.nodes, {"nm_dom": ["dominio_sintetico"]})
        self.assertFalse(any(i["code"] == "external_list_missing" for i in issues))
        self.assertNotIn("nm_sist_orig", get_catalog()["View"])
        self.assertEqual("Não identificada", get_catalog()["Table"]["nm_dom"]["category"])

    def test_consistency_mask_consumption_and_classification(self):
        text = CONTRACT.replace('fl_mask: "true"', 'fl_mask: "false"').replace('fl_apt_cons: "false"', 'fl_apt_cons: "true"').replace('nm_class_inf: "restrita"', 'nm_class_inf: "publica"')
        batch = parse_batch([upload(text)])
        codes = {i["code"] for i in validate_nodes(batch.nodes)}
        self.assertTrue({"mask_inconsistent", "consumption_quality", "consumption_approval", "classification_too_low", "classification_review"}.issubset(codes))
        self.assertNotIn("fl_pii", batch.nodes[0]["tags"])

    def test_invalid_edits_rejected_deletion_explicit_and_original_preserved(self):
        batch = parse_batch([upload(CONTRACT)])
        node = batch.nodes[0]
        edits = [
            {"id": node["id"], "field": "id", "value": "outro"},
            {"id": node["id"], "field": "tags", "tag": "fl_legado", "value": True},
            {"id": node["id"], "field": "tags", "tag": "nm_environment", "value": "lab"},
            {"id": node["id"], "field": "tags", "tag": "inventada", "value": "valor"},
            {"id": node["id"], "field": "tags", "tag": "nm_retencao", "value": "2_anos"},
            {"id": node["id"], "field": "tags", "tag": "dt_aprv_gov", "value": "2026-02-30"},
            {"id": node["id"], "field": "tags", "tag": "nm_dom", "value": None},
        ]
        report = apply_edits(batch, edits)
        self.assertEqual(["rejected"] * 6 + ["applied"], [r["status"] for r in report])
        self.assertNotIn("nm_dom", node["tags"])
        self.assertEqual("dominio_sintetico", node["original_tags"]["nm_dom"])

    def test_conflicting_edits_are_not_last_write_wins(self):
        batch = parse_batch([upload(CONTRACT)])
        nid = batch.nodes[0]["id"]
        report = apply_edits(batch, [{"id": nid, "field": "description", "value": "A"}, {"id": nid, "field": "description", "value": "B"}])
        self.assertEqual(["rejected", "rejected"], [r["status"] for r in report])
        self.assertEqual("Descrição original.\n", batch.nodes[0]["description"])

    def test_duplicate_object_versions_are_visible_and_export_blocked(self):
        batch = parse_batch([upload(CONTRACT, "a.yaml"), upload(CONTRACT.replace("Descrição original.", "Alteração."), "b.yaml")])
        self.assertEqual(2, len(batch.documents))
        self.assertTrue(batch.duplicates)
        self.assertTrue(batch.duplicates[0]["conflict"])
        self.assertTrue(batch.errors)
        with self.assertRaises(ValueError):
            render_yaml_zip(batch, [])

    def test_zip_import_no_disk_extraction_and_dangerous_entries_rejected(self):
        batch = parse_batch([zip_upload([("pasta/arquivo.yaml", CONTRACT)])])
        self.assertFalse(batch.errors)
        self.assertEqual("lote/pasta/arquivo.yaml", batch.documents[0].filename)
        for entries in ([('../escape.yaml', CONTRACT)], [('C:/absoluto.yaml', CONTRACT)], [('aninhado.zip', b'PK')], [('texto.txt', 'invalido')]):
            batch = parse_batch([zip_upload(entries)])
            self.assertTrue(batch.errors)
            self.assertFalse(batch.documents)
        bad_separator = zip_upload([("pasta/escape.yaml", CONTRACT)])
        bad_separator["content"] = base64.b64encode(base64.b64decode(bad_separator["content"]).replace(b"pasta/escape.yaml", b"pasta\\escape.yaml")).decode()
        self.assertTrue(parse_batch([bad_separator]).errors)
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w") as archive:
            link = zipfile.ZipInfo("link.yaml")
            link.create_system = 3
            link.external_attr = (stat.S_IFLNK | 0o777) << 16
            archive.writestr(link, "destino")
        batch = parse_batch([{"name": "links.zip", "content": base64.b64encode(buffer.getvalue()).decode()}])
        self.assertTrue(batch.errors)

    def test_unsafe_yaml_rejected_before_construction(self):
        adversarial = [
            "table: !!python/object/apply:os.system ['comando_nao_executado']",
            "table: !custom {id: 'c.s.t'}",
            "table: &a {id: 'c.s.t', unknown: *a}",
            "table: {id: 'c.s.t', id: 'c.s.outro'}",
            "a: " + "[" * 80 + "0" + "]" * 80,
        ]
        for text in adversarial:
            batch = parse_batch([upload(text)])
            self.assertTrue(batch.errors, text)
            self.assertFalse(batch.documents)
            self.assertNotIn("comando_nao_executado", batch.errors[0]["message"])

    def test_multiple_yaml_documents_preserved(self):
        batch = parse_batch([upload("---\ncatalog: {name: primeiro}\n---\ncatalog: {name: segundo}\n")])
        self.assertFalse(batch.errors)
        with zipfile.ZipFile(io.BytesIO(render_yaml_zip(batch, []))) as archive:
            reparsed = parse_batch([upload(archive.read("contrato.yaml").decode())])
        self.assertEqual(2, len(reparsed.nodes))


if __name__ == "__main__":
    unittest.main()

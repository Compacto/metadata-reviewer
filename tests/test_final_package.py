"""Final packages contain only complete, explicitly approved source documents."""
import base64
import io
import json
import zipfile

import pytest

from app import service


def upload(text, name):
    return {"name": name, "content": base64.b64encode(text.encode()).decode()}


def contract(name, extra=""):
    return f'''# contrato sintético
table:
  id: "catalogo_teste.gold.{name}"
  description: "Descrição sintética"
  tags: {{fl_mask: "false"}}
ingest:
  extra_preservado: "valor_sintetico"
{extra}'''


def package(content):
    with zipfile.ZipFile(io.BytesIO(content)) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def completed(nodes, **review):
    return {n["id"]: {"status": "Validated", "completed": True, **review}
            for n in nodes if n["level"] in ("Table", "View")}


def test_only_approved_source_files_are_emitted_and_applied_report_excludes_deferred_edits():
    files = [upload(contract("aprovada"), "aceito.yaml"), upload(contract("pendente"), "pendente.yaml")]
    initial = service.prepare(files)
    approved, pending = initial["nodes"]
    reviews = completed([approved])
    edits = [{"id": approved["id"], "field": "description", "value": "Revisão aprovada sintética"},
             {"id": pending["id"], "field": "description", "value": "Revisão ainda pendente sintética"}]
    result = package(service.final_zip(files, edits, reviews, confirmed=True))
    assert {name for name in result if name.endswith((".yaml", ".yml"))} == {"aceito.yaml"}
    assert b"# contrato sint" in result["aceito.yaml"]
    assert "Revisão aprovada sintética" in result["aceito.yaml"].decode()
    assert 'extra_preservado: "valor_sintetico"' in result["aceito.yaml"].decode()
    applied = json.loads(result["relatorios/alteracoes_aplicadas.json"])
    assert [row["id"] for row in applied] == [approved["id"]]
    ignored = json.loads(result["relatorios/objetos_ignorados_e_pendencias.json"])
    assert ignored["files_included"] == 1 and ignored["objects_included"] == 1
    assert ignored["objects"][0]["id"] == pending["id"]
    assert ignored["deferred_changes"][0]["id"] == pending["id"]


def test_a_mixed_document_is_skipped_atomically_without_unvalidated_originals():
    files = [upload('''tables:
  - id: "catalogo_teste.gold.aprovada"
    tags: {fl_mask: "false"}
  - id: "catalogo_teste.gold.pendente"
    tags: {fl_mask: "false"}
''', "misto.yaml")]
    initial = service.prepare(files)
    result = package(service.final_zip(files, reviews=completed([initial["nodes"][0]]), confirmed=True))
    assert not any(name.endswith((".yaml", ".yml")) for name in result)
    ignored = json.loads(result["relatorios/objetos_ignorados_e_pendencias.json"])
    assert len(ignored["objects"]) == 2
    assert ignored["files_included"] == 0
    assert any("outro objeto" in row["reason"] for row in ignored["objects"])


@pytest.mark.parametrize("status", ["Pending", "In review", "Awaiting clarification", "Validated with remarks", "Rejected"])
def test_unapproved_decisions_are_kept_out_of_final_yaml(status):
    files = [upload(contract("decisao"), "decisao.yaml")]
    initial = service.prepare(files)
    review = {initial["nodes"][0]["id"]: {"status": status, "completed": status in ("Validated with remarks", "Rejected", "Awaiting clarification")}}
    state = service.prepare(files, reviews=review)
    if status in ("Rejected", "Awaiting clarification"):
        assert state["reviews"][initial["nodes"][0]["id"]]["status"] == status
    result = package(service.final_zip(files, reviews=review, confirmed=True))
    assert "decisao.yaml" not in result


def test_business_pending_text_prevents_accidental_validated_export():
    files = [upload(contract("pendencia"), "pendencia.yaml")]
    initial = service.prepare(files)
    reviews = completed(initial["nodes"], pending="Aguardar decisão de negócio sintética")
    state = service.prepare(files, reviews=reviews)
    assert state["reviews"][initial["nodes"][0]["id"]]["status"] == "Validated with remarks"
    assert "pendencia.yaml" not in package(service.final_zip(files, reviews=reviews, confirmed=True))


def test_parent_technical_error_blocks_whole_hierarchical_source_document():
    files = [upload('''catalog:
  name: catalogo_teste
  tags: {nm_environment: "invalido"}
  schemas:
    gold:
      tables:
        tabela_teste:
          tags: {fl_mask: "false"}
''', "hierarquia.yaml")]
    initial = service.prepare(files)
    reviews = completed(initial["nodes"])
    assert any(issue.get("id") == initial["nodes"][0]["id"] for issue in initial["issues"])
    result = package(service.final_zip(files, reviews=reviews, confirmed=True))
    assert "hierarquia.yaml" not in result
    ignored = json.loads(result["relatorios/objetos_ignorados_e_pendencias.json"])
    assert ignored["files_included"] == 0 and ignored["objects"]


def test_technical_errors_and_external_list_pendencies_override_requested_validated_status():
    text = contract("erro").replace('fl_mask: "false"', 'fl_mask: "invalido", nm_dom: "dominio_sem_lista"')
    files = [upload(text, "tecnico.yaml")]
    initial = service.prepare(files)
    reviews = completed(initial["nodes"])
    state = service.prepare(files, reviews=reviews)
    assert state["reviews"][initial["nodes"][0]["id"]]["status"] == "Validated with remarks"
    assert "tecnico.yaml" not in package(service.final_zip(files, reviews=reviews, confirmed=True))


def test_category_reference_warning_remains_visible_without_blocking_clean_approval():
    files = [upload(contract("categoria"), "categoria.yaml")]
    initial = service.prepare(files)
    assert any(issue.get("scope") == "governance" for issue in initial["issues"])
    result = package(service.final_zip(files, reviews=completed(initial["nodes"]), confirmed=True))
    assert "categoria.yaml" in result
    pending = json.loads(result["relatorios/objetos_ignorados_e_pendencias.json"])["pending"]
    assert any(issue.get("scope") == "governance" for issue in pending)


@pytest.mark.parametrize("confirmed", [False, None, 1, "true"])
def test_generation_requires_explicit_boolean_confirmation(confirmed):
    files = [upload(contract("confirmacao"), "confirmacao.yaml")]
    with pytest.raises(ValueError, match="explicitamente"):
        service.final_zip(files, confirmed=confirmed)


@pytest.mark.parametrize("kwargs", [{"edits": {}}, {"reviews": []}, {"refs": []}])
def test_falsey_malformed_parameters_are_not_silently_defaulted(kwargs):
    with pytest.raises(ValueError):
        service.prepare([upload(contract("parametros"), "parametros.yaml")], **kwargs)


@pytest.mark.parametrize("timestamp", ["data_invalida", {"conteudo": "nao_e_data"}, 12345])
def test_review_timestamp_is_validated_as_iso_text(timestamp):
    files = [upload(contract("data"), "data.yaml")]
    initial = service.prepare(files)
    with pytest.raises(ValueError):
        service.prepare(files, reviews=completed(initial["nodes"], validated_at=timestamp))

"""Interface Streamlit. Conteúdo importado permanece somente na RAM da sessão.

Não usa cache, diretórios temporários, serviços externos ou persistência.
"""
from __future__ import annotations

import base64
import copy
import html
import json
import re
from datetime import datetime, timezone

import streamlit as st

from app.rules import get_catalog
from app.privacy import release_session_files
from app.service import export_xlsx, final_zip, prepare, preview_xlsx


STATUSES = ["Pending", "In review", "Awaiting clarification", "Validated", "Validated with remarks", "Rejected"]
LAYERS = {
    "bronze": "Dados recebidos da origem, preservando suas características iniciais.",
    "silver": "Dados tratados, padronizados e preparados para uso.",
    "gold": "Dados consolidados e preparados para análises e consumo.",
    "data_product": "Produtos de dados destinados a necessidades específicas de negócio.",
    "information_schema": "Metadados técnicos sobre os objetos do catálogo.",
}
STYLE = """<style>
:root { --navy:#142b40; --teal:#008a78; }
.stApp { background:#f5f7fa; }
[data-testid="stHeader"] {background:#f5f7fae8;}
[data-testid="stSidebar"] {background:#142b40;color:#ccdae5;border-right:1px solid #244357;}
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"], [data-testid="stSidebar"] label,
[data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 {color:#c8dbe5;}
[data-testid="stSidebar"] .stCaption {color:#90abbc;}
[data-testid="stSidebar"] [data-testid="stExpander"] {border-color:#345165;}
[data-testid="stSidebar"] button {background:#203d52;color:#d6e6ee;border-color:#36566b;}
.block-container {padding-top:4rem;max-width:1500px;padding-bottom:2rem;}
h1 {color:#17354d;font-weight:700!important;letter-spacing:-1.2px;font-size:2.1rem!important;}
h2,h3 {color:#23465d;letter-spacing:-.4px;}
[data-testid="stMetric"] {background:white;border:1px solid #e0e8ef;border-radius:10px;padding:19px 21px;}
[data-testid="stMetricLabel"] p {font-size:.77rem;color:#748b9b;}
[data-testid="stMetricValue"] {font-size:2rem;color:#24485f;}
[data-testid="stVerticalBlockBorderWrapper"]>div {border-color:#dfe8ee!important;border-radius:10px;}
.stButton>button,.stDownloadButton>button {border-radius:6px;font-size:.82rem;}
[data-testid="stFileUploader"] {border-radius:8px;}
[data-testid="stFileUploaderDropzone"] {background:#f0f8f5;border:1px dashed #a7d2c5;}
[data-baseweb="tab-list"] {gap:22px;}
.eyebrow {color:#7e96a6;font-size:.67rem;letter-spacing:1.8px;font-weight:650;margin-bottom:6px;}
.brand {font-size:1.63rem;font-weight:700;color:#edf7fa;letter-spacing:-1px;line-height:1.2;}
.brand em {color:#38c6ad;font-style:normal;}
.brand small {font-size:.59rem;color:#88aabb;display:block;letter-spacing:1.9px;margin-top:8px;margin-bottom:25px;}
.hero {background:white;border:1px solid #e0e9ee;border-radius:10px;padding:26px 29px;margin:20px 0 24px;}
.hero .kicker {color:#179382;font-size:.65rem;letter-spacing:1.4px;font-weight:650;}
.hero h2 {font-size:1.45rem;max-width:650px;margin:9px 0 5px;line-height:1.4;}
.hero p {color:#8195a4;font-size:.83rem;margin:6px 0 0;line-height:1.75;}
.tech-grid {display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:19px;padding:21px;background:#edf2f6;border-radius:7px;margin:12px 0 20px;}
.tech-label {font-size:.67rem;color:#8499a8;margin-bottom:6px;}
.tech-value {font-size:.79rem;color:#365970;overflow-wrap:anywhere;}
.status-pill {display:inline-block;padding:4px 10px;background:#e8f3ef;color:#238573;border-radius:5px;font-size:.73rem;}
.memory-note {font-size:.7rem;color:#7e98a9;line-height:1.8;border-top:1px solid #355165;padding-top:15px;margin-top:22px;}
.page-foot {font-size:.68rem;color:#95a8b5;border-top:1px solid #dfe7ed;margin-top:30px;padding-top:16px;}
.stTabs [data-baseweb="tab-panel"] {padding-top:20px;}
@media(max-width:700px) {.block-container{padding-left:1rem;padding-right:1rem;}.tech-grid{grid-template-columns:repeat(2,minmax(0,1fr));}.hero{padding:20px;}h1{font-size:1.8rem!important;}}
</style>"""


def _fresh() -> dict:
    return {"files": [], "edits": [], "reviews": {}, "refs": {}, "prepared": None,
            "preview": None, "xlsx": None, "zip": None, "selected": None,
            "revision": 0, "notice": "", "page": "Visão geral"}


def _state() -> dict:
    if "mr" not in st.session_state:
        st.session_state.mr = _fresh()
    if "mr_epoch" not in st.session_state:
        st.session_state.mr_epoch = 0
    return st.session_state.mr


def _key(name: str) -> str:
    return f"mr_{st.session_state.mr_epoch}_{name}"


def _safe(value) -> str:
    return html.escape("" if value is None else str(value), quote=True)


def _text(value) -> str:
    return "" if value is None else str(value)


def _md(value) -> str:
    """Render imported identifiers/errors as text, never Markdown images/links."""
    return re.sub(r"([\\`*_{}\[\]()#+.!>|~-])", r"\\\1", _text(value))


def _nodes(s: dict) -> list[dict]:
    return (s.get("prepared") or {}).get("nodes", [])


def _issues(s: dict, node: dict | None = None) -> list[dict]:
    issues = (s.get("prepared") or {}).get("issues", [])
    if node is None:
        return issues
    ids = {node["id"]}
    if node["level"] in ("Table", "View"):
        ids |= {n["id"] for n in _nodes(s) if n["level"] == "Column" and n["object_id"] == node["id"]}
    return [i for i in issues if i.get("id", i.get("node_id")) in ids]


def _review(s: dict, nid: str) -> dict:
    return s["reviews"].get(nid, {"status": "Pending", "completed": False})


def _status(s: dict, node: dict) -> str:
    status = _review(s, node["id"]).get("status", "Pending")
    if status == "Validated" and _issues(s, node):
        return "Validated with remarks"
    return status


def _merge_edits(existing: list, incoming: list) -> list:
    result = {(e.get("id"), e.get("field"), e.get("tag")): e for e in existing}
    for edit in incoming:
        result[(edit.get("id"), edit.get("field"), edit.get("tag"))] = edit
    return list(result.values())


def _refresh(s: dict, *, edits=None, reviews=None, files=None, refs=None) -> bool:
    candidate = {"files": s["files"] if files is None else files,
                 "edits": s["edits"] if edits is None else edits,
                 "reviews": s["reviews"] if reviews is None else reviews,
                 "refs": s["refs"] if refs is None else refs}
    try:
        with st.spinner("Processando e validando em memória…"):
            result = prepare(candidate["files"], candidate["edits"], candidate["reviews"], candidate["refs"])
    except ValueError as exc:
        st.error(_md(exc))
        return False
    except Exception:
        st.error("Não foi possível concluir esta operação. Confira os arquivos e tente novamente.")
        return False
    s.update(candidate)
    s["prepared"] = result
    s["reviews"] = result.get("reviews", candidate["reviews"])
    s["revision"] += 1
    s["preview"] = s["xlsx"] = s["zip"] = None
    return True


def _show_records(records, *, limit=250) -> None:
    if not records:
        st.caption("Nenhum registro.")
        return
    if isinstance(records, dict):
        records = [{"campo": k, "informação": _text(v)} for k, v in records.items()]
    rows = []
    for record in list(records)[:limit]:
        if isinstance(record, dict):
            rows.append({str(k): _text(v) if isinstance(v, (dict, list, tuple)) else v for k, v in record.items()})
        else:
            rows.append({"informação": str(record)})
    st.dataframe(rows, hide_index=True, use_container_width=True)
    if len(records) > limit:
        st.caption(f"Exibindo os primeiros {limit} de {len(records)} registros; o relatório completo acompanha o download.")


def _show_issues(issues: list, *, key: str) -> None:
    if not issues:
        st.success("Nenhuma inconsistência identificada neste recorte.")
        return
    for index, issue in enumerate(issues[:60]):
        message = f"{issue.get('field', 'Validação')} · {issue.get('message', 'Revisão necessária')}"
        if issue.get("severity") == "error":
            st.error(_md(message), icon="🚩")
        else:
            st.warning(_md(message), icon="⚠️")
        if issue.get("path") or issue.get("filename"):
            st.caption(_md(f"{issue.get('path', '')} · {issue.get('filename', '')}"))
    if len(issues) > 60:
        st.caption(f"Mostrando 60 de {len(issues)} pendências.")
        with st.expander("Consultar relatório completo"):
            _show_records(issues, limit=1000)


def _go_review(s: dict, nid: str) -> None:
    s["selected"] = nid
    s["page"] = "Revisar metadados"
    s["pending_navigation"] = "Revisar metadados"


def _sidebar(s: dict) -> str:
    with st.sidebar:
        st.markdown('<div class="brand">metadata<em> review</em><small>DATA GOVERNANCE WORKSPACE</small></div>', unsafe_allow_html=True)
        page = st.radio("Navegação", ["Visão geral", "Revisar metadados", "Arquivos e downloads", "Regras e referências"], key=_key("navigation"), label_visibility="collapsed")
        s["page"] = page
        st.divider()
        st.caption("EXPLORAR CATÁLOGO")
        nodes = _nodes(s)
        if not nodes:
            st.caption("Sua hierarquia de dados aparecerá aqui após a importação.")
        for cat in sorted({n["catalog"] for n in nodes}):
            objects = [n for n in nodes if n["catalog"] == cat and n["level"] in ("Table", "View")]
            done = sum(_review(s, n["id"]).get("completed", False) for n in objects)
            with st.expander(_md(f"▦ {cat} · {done}/{len(objects)}"), expanded=True):
                cat_nodes = [n for n in nodes if n["level"] == "Catalog" and n["catalog"] == cat]
                for n in cat_nodes:
                    st.button("Revisar catálogo", key=_key(f"cat_{n['id']}"), use_container_width=True, on_click=_go_review, args=(s, n["id"]))
                for schema in sorted({n["schema"] for n in nodes if n["catalog"] == cat and n.get("schema")}):
                    schema_objects = [n for n in objects if n["schema"] == schema]
                    complete = sum(_review(s, n["id"]).get("completed", False) for n in schema_objects)
                    with st.expander(_md(f"▤ {schema} · {complete}/{len(schema_objects)}"), expanded=True):
                        for n in nodes:
                            if n["level"] == "Schema" and n["catalog"] == cat and n["schema"] == schema:
                                st.button("Revisar schema", key=_key(f"schema_{n['id']}"), use_container_width=True, on_click=_go_review, args=(s, n["id"]))
                        for n in schema_objects:
                            marker = "✓" if _review(s, n["id"]).get("completed") else "○"
                            st.button(_md(f"{marker} {n['name']}"), key=_key(f"obj_{n['id']}"), help=f"{n['level']} · {_status(s, n)}", use_container_width=True, on_click=_go_review, args=(s, n["id"]))
                            if s["selected"] == n["id"] or any(c["id"] == s["selected"] and c.get("object_id") == n["id"] for c in nodes):
                                with st.expander("Colunas", expanded=False):
                                    for col in [c for c in nodes if c["level"] == "Column" and c["object_id"] == n["id"]]:
                                        st.button(_md(f"└ {col['name']}"), key=_key(f"col_{col['id']}"), use_container_width=True, on_click=_go_review, args=(s, col["id"]))
        st.markdown('<div class="memory-note">◈ &nbsp; Sessão privada em memória<br>Uploads e alterações não são gravados em disco. Para retomar, reimporte os originais e o Excel.</div>', unsafe_allow_html=True)
        with st.expander("Limpar sessão"):
            confirm = st.checkbox("Descartar arquivos e alterações desta sessão", key=_key("clear_confirm"))
            if st.button("Limpar agora", disabled=not confirm, key=_key("clear"), use_container_width=True):
                release_session_files()
                epoch = st.session_state.mr_epoch + 1
                for k in list(st.session_state):
                    del st.session_state[k]
                st.session_state.mr_epoch = epoch
                st.session_state.mr = _fresh()
                st.rerun()
    return page


def _upload(s: dict) -> None:
    with st.container(border=True):
        st.subheader("Importar YAMLs ou lotes ZIP")
        st.caption("Arraste os arquivos para a área abaixo. O processamento ocorre em memória no computador que executa o servidor.")
        with st.form(_key("upload_form"), clear_on_submit=False):
            files = st.file_uploader("Arquivos de metadados", type=["yaml", "yml", "zip"], accept_multiple_files=True, key=_key("yaml_upload"))
            mode = st.radio("Como importar", ["Adicionar à sessão", "Substituir a sessão"], horizontal=True)
            confirmed = st.checkbox("Confirmo descartar arquivos e alterações atuais ao substituir", key=_key("replace_confirm")) if s["files"] else True
            submitted = st.form_submit_button("Importar e validar", type="primary", use_container_width=True)
        if submitted:
            if not files:
                st.warning("Selecione pelo menos um arquivo YAML, YML ou ZIP.")
                return
            if mode == "Substituir a sessão" and not confirmed:
                st.warning("Confirme a substituição para descartar a sessão atual.")
                return
            if sum(f.size for f in files) > 50 * 1024 * 1024:
                st.error("O lote excede o limite de 50 MiB. Divida os arquivos em lotes menores.")
                return
            if any(f.size > 10 * 1024 * 1024 for f in files):
                st.error("Cada arquivo deve ter até 10 MiB.")
                return
            incoming = [{"name": f.name, "content": base64.b64encode(f.getvalue()).decode("ascii")} for f in files]
            replace = mode == "Substituir a sessão"
            current = [] if replace else s["files"]
            for entry in incoming:
                other = next((f for f in current if f["name"] == entry["name"]), None)
                if other and other["content"] != entry["content"]:
                    st.error(_md(f"Há uma versão diferente de {entry['name']} na sessão. Remova a versão anterior em Arquivos e downloads ou substitua a sessão explicitamente."))
                    return
            dedup = {(f["name"], f["content"]): f for f in [*current, *incoming]}
            if _refresh(s, files=list(dedup.values()), edits=[] if replace else s["edits"], reviews={} if replace else s["reviews"]):
                if replace:
                    s["selected"] = None
                s["notice"] = "Importação processada. Confira o relatório e as pendências antes de exportar."
                st.rerun()
        st.caption("Limites: 10 MiB por arquivo e 50 MiB por lote. ZIPs são inspecionados sem extração para o disco.")


def _import_report(s: dict) -> None:
    p = s.get("prepared") or {}
    batch = p.get("batch")
    errors = p.get("errors", getattr(batch, "errors", []))
    duplicates = p.get("duplicates", getattr(batch, "duplicates", []))
    count = getattr(batch, "file_count", len(s["files"]))
    accepted = getattr(batch, "processed_count", len({n["filename"] for n in _nodes(s)}))
    st.caption(f"Lote: {count} arquivo(s) identificado(s) · {accepted} processado(s) · {len(errors)} erro(s) · {len(duplicates)} duplicidade(s).")
    if errors or duplicates:
        st.error("O lote contém arquivos ou objetos com problemas. Consulte o relatório; a exportação final é controlada para evitar descarte silencioso.")
        with st.expander("Relatório de importação", expanded=True):
            if errors:
                st.write("**Arquivos com erro**")
                _show_records(errors)
            if duplicates:
                st.write("**Duplicidades e conflitos entre versões**")
                _show_records(duplicates)
            st.caption("Remova a origem conflitante em Arquivos e downloads e importe a versão desejada. A aplicação não escolhe uma versão automaticamente.")


def _dashboard(s: dict) -> None:
    st.markdown('<div class="eyebrow">WORKSPACE DE VALIDAÇÃO</div>', unsafe_allow_html=True)
    st.title("Governança de dados")
    st.caption("Uma visão clara dos metadados. Uma revisão de cada vez.")
    if not s["files"]:
        st.markdown('<div class="hero"><div class="kicker">DO YAML À VALIDAÇÃO DE NEGÓCIO</div><h2>Prepare os metadados para revisão.</h2><p>Importe seus arquivos, revise os objetos e gere a planilha para o Data Owner.</p></div>', unsafe_allow_html=True)
        _upload(s)
        cols = st.columns(3)
        for col, title, body in zip(cols, ["01 · Importe", "02 · Revise", "03 · Compartilhe"], ["YAMLs individuais ou lotes compactados, organizados pela hierarquia real.", "Descrições, tags e colunas com orientações e pendências visíveis.", "Excel para validação e novos YAMLs após aprovação explícita."]):
            with col:
                st.markdown(f"**{title}**")
                st.caption(body)
        return
    objects = [n for n in _nodes(s) if n["level"] in ("Table", "View")]
    governance = [i for i in _issues(s) if i.get("scope") == "governance"]
    if governance:
        st.warning("Pendência global da Governança: as categorias [G], [T] e [REDUNDANTE] não são atribuídas por tag no PDF. Elas permanecem como Não identificada e não foram inferidas. Essa lacuna é separada da conformidade técnica dos objetos.")
    completed = sum(bool(_review(s, n["id"]).get("completed")) for n in objects)
    clean = sum(_status(s, n) == "Validated" for n in objects)
    remarks = sum(_status(s, n) == "Validated with remarks" for n in objects)
    cols = st.columns(4)
    cols[0].metric("Objetos importados", len(objects), help="Somente tabelas e views; colunas são atributos dos objetos.")
    cols[0].caption(f"{sum(n['level']=='Table' for n in objects)} tabelas · {sum(n['level']=='View' for n in objects)} views")
    cols[1].metric("Aguardando conclusão", len(objects) - completed)
    cols[1].caption("Revisão funcional ainda não concluída")
    cols[2].metric("Objetos validados", clean)
    cols[2].caption(f"{remarks} com ressalvas")
    cols[3].metric("Inconsistências técnicas", sum(i.get("severity") == "error" for i in _issues(s)))
    cols[3].caption(f"{sum(i.get('severity') != 'error' for i in _issues(s))} pendências de revisão")
    st.write("")
    with st.container(border=True):
        left, right = st.columns([5, 1])
        left.markdown("**Progresso da validação funcional**")
        right.markdown(f"**{round(100 * completed / len(objects)) if objects else 0}%**")
        st.progress(completed / len(objects) if objects else 0)
        st.caption(f"{completed} de {len(objects)} objetos com revisão concluída. Conclusão funcional não significa conformidade técnica.")
    _import_report(s)
    st.subheader("Seus objetos de dados")
    st.caption("Encontre um ativo para revisar seus metadados e suas pendências.")
    search = st.text_input("Pesquisar objetos", placeholder="Nome técnico, nome de negócio, domínio ou tag", key=_key("search"))
    f1, f2, f3, f4 = st.columns([1, 1, 1.5, 1])
    layer = f1.selectbox("Camada", ["Todas", *sorted({n.get('schema', '') for n in objects})], key=_key("layer"))
    kind = f2.selectbox("Tipo", ["Todos", "Table", "View"], key=_key("kind"))
    status = f3.selectbox("Status", ["Todos", *STATUSES], key=_key("status"))
    pending = f4.checkbox("Com pendências", key=_key("pending_filter"))
    found = []
    for n in objects:
        haystack = " ".join([n["path"], _text(n.get("business_name")), json.dumps(n.get("tags", {}), ensure_ascii=False)]).casefold()
        if search.casefold() not in haystack or (layer != "Todas" and n["schema"] != layer) or (kind != "Todos" and n["level"] != kind) or (status != "Todos" and _status(s, n) != status) or (pending and not _issues(s, n)):
            continue
        found.append(n)
    rows = [{"Objeto": n["name"], "Catálogo": n["catalog"], "Camada": n["schema"], "Tipo": n["level"], "Status": _status(s, n), "Pendências": len(_issues(s, n))} for n in found]
    if rows:
        st.dataframe(rows, hide_index=True, use_container_width=True)
        selected = st.selectbox("Objeto para revisar", [n["id"] for n in found], format_func=lambda nid: next(n["path"] for n in found if n["id"] == nid), key=_key("dashboard_selection"))
        st.button("Abrir revisão do objeto", type="primary", key=_key("open_review"), on_click=_go_review, args=(s, selected))
    else:
        st.info("Nenhum objeto corresponde aos filtros selecionados.")
    st.caption(f"{len(found)} de {len(objects)} objetos.")


def _tag_editor(s: dict, node: dict) -> None:
    catalog = (s.get("prepared") or {}).get("catalog", get_catalog(s["refs"]))
    rules = catalog.get(node["level"], {})
    tags = node.get("tags", {})
    original = node.get("original_tags", {})
    st.caption("Somente as tags presentes neste objeto são editadas aqui. Valores ausentes não recebem padrões automáticos.")
    if not tags and not original:
        st.info("Este objeto não contém tags. Nenhuma tag foi criada automaticamente.")
        return
    rows = []
    for tag in dict.fromkeys([*original, *tags]):
        rule = rules.get(tag, {})
        matching = [i for i in _issues(s, node) if i.get("id") == node["id"] and tag in _text(i.get("field"))]
        allowed = rule.get("values")
        allowed_text = ", ".join(map(str, allowed)) if allowed is not None else "Lista oficial não carregada" if rule.get("external") else rule.get("format", "Texto livre; revisão necessária" if not rule else "Texto livre")
        situation = "Inválida" if any(i.get("severity") == "error" for i in matching) else "Requer revisão" if matching else "Alterada" if tags.get(tag) != original.get(tag) else "Válida"
        rows.append({"Tag técnica": tag, "Significado": rule.get("description", "Tag não identificada no padrão fornecido."), "Valor atual": _text(original.get(tag)), "Valor proposto": "[exclusão explícita]" if tag not in tags else _text(tags[tag]), "Valores permitidos": allowed_text, "Categoria": rule.get("category", "Não identificada"), "Situação": situation})
    st.dataframe(rows, use_container_width=True, hide_index=True)
    tag = st.selectbox("Tag para revisar", list(dict.fromkeys([*original, *tags])), key=_key(f"tagselect_{node['id']}"))
    rule = rules.get(tag, {})
    current = _text(tags.get(tag))
    removed = tag not in tags
    st.info(rule.get("description", "Tag não identificada no padrão. O valor será preservado e a validação permanecerá pendente até revisão da Governança."), icon="ℹ")
    if rule.get("external") and not rule.get("values"):
        st.warning("Lista oficial não carregada. Carregue a referência em Regras e referências.")
    form_key = _key(f"tag_form_{node['id']}_{tag}_{s['revision']}")
    with st.form(form_key):
        options = rule.get("values")
        if options:
            keep = "Manter exclusão atual" if removed else "Manter valor atual"
            choices = [keep, *options]
            proposed = st.selectbox("Valor proposto", choices, index=0, help="Os valores permitidos seguem o catálogo. Manter valor atual preserva inclusive um valor que ainda precise de correção.")
            value = (None if removed else current) if proposed == keep else proposed
        else:
            proposed = st.text_input("Valor proposto", value=current, help=rule.get("format", "Vazio preserva o valor atual."), max_chars=10000)
            value = proposed if proposed != "" else None if removed else current
        delete = st.checkbox("Excluir explicitamente esta tag", help="A exclusão remove apenas esta tag e pode gerar uma pendência de informação ausente.")
        confirm_delete = st.checkbox("Confirmo a exclusão selecionada")
        comments = st.text_area("Comentário ou justificativa da tag", value=_review(s, node["id"]).get("tag_comments", {}).get(tag, ""), max_chars=10000)
        apply = st.form_submit_button("Aplicar revisão da tag", type="primary")
    if apply:
        if delete and not confirm_delete:
            st.warning("Confirme explicitamente a exclusão da tag.")
            return
        new_edits = _merge_edits(s["edits"], [{"id": node["id"], "field": "tags", "tag": tag, "value": None if delete else value}])
        reviews = copy.deepcopy(s["reviews"])
        reviews.setdefault(node["id"], {}).setdefault("tag_comments", {})[tag] = comments
        if _refresh(s, edits=new_edits, reviews=reviews):
            s["notice"] = "Revisão da tag processada. Valores rejeitados permanecem no relatório de alterações."
            st.rerun()


def _business_review(s: dict, node: dict) -> None:
    review = _review(s, node["id"])
    st.markdown("#### Informações de negócio")
    st.caption("A identificação técnica permanece somente leitura. Campos vazios preservam o valor atual.")
    left, right = st.columns(2)
    left.text_area("Descrição original", value=_text(node.get("original_description")), disabled=True, key=_key(f"original_desc_{node['id']}"))
    with st.form(_key(f"business_form_{node['id']}_{s['revision']}")):
        description = st.text_area("Descrição proposta", value=_text(node.get("description")), height=120, max_chars=32000)
        business = None
        if node.get("original_business_name") is not None or node.get("business_name") is not None:
            business = st.text_input("Nome de negócio proposto", value=_text(node.get("business_name")), max_chars=32000)
        else:
            st.caption("Nome de negócio não informado no YAML; nenhum campo foi criado automaticamente.")
        with st.expander("Exclusão explícita de campos"):
            remove_desc = st.checkbox("Excluir descrição de negócio")
            remove_business = st.checkbox("Excluir nome de negócio", disabled=business is None)
            confirm_delete = st.checkbox("Confirmo remover os campos selecionados")
        st.markdown("#### Revisão funcional individual")
        status = st.selectbox("Status da revisão", STATUSES, index=STATUSES.index(review.get("status", "Pending")) if review.get("status", "Pending") in STATUSES else 0)
        completed = st.checkbox("Validation completed", value=bool(review.get("completed")))
        owner = st.text_input("Data Owner (opcional)", value=review.get("owner", ""), help="Identificação informada pelo usuário. Esta aplicação local não autentica a identidade do responsável.", max_chars=500)
        comments = st.text_area("Comentários e esclarecimentos", value=review.get("comments", ""), max_chars=10000)
        pending_text = st.text_area("Pendências do Data Owner", value=review.get("pending", ""), max_chars=10000, help="Pendências de negócio também mantêm a conclusão como Validated with remarks.")
        final_comments = st.text_area("Comentários finais da decisão", value=review.get("final_comments", ""), max_chars=10000)
        pending_owner = st.text_input("Responsável por tratar as pendências (opcional)", value=review.get("pending_owner", ""), max_chars=500)
        submit = st.form_submit_button("Salvar revisão em memória", type="primary", use_container_width=True)
    if _issues(s, node):
        st.warning("Há pendências técnicas neste objeto ou em suas colunas. Ao concluir como validado, o status será Validated with remarks; as pendências continuarão visíveis.")
    if review.get("validated_at"):
        st.caption(f"Última conclusão informada: {review['validated_at']}. Responsável autodeclarado, sem autenticação.")
    if submit:
        if (remove_desc or remove_business) and not confirm_delete:
            st.warning("Confirme explicitamente a exclusão dos campos selecionados.")
            return
        edits = []
        if remove_desc:
            edits.append({"id": node["id"], "field": "description", "value": None})
        elif description != "":
            edits.append({"id": node["id"], "field": "description", "value": description})
        if remove_business:
            edits.append({"id": node["id"], "field": "business_name", "value": None})
        elif business:
            edits.append({"id": node["id"], "field": "business_name", "value": business})
        if completed and status not in ("Rejected", "Awaiting clarification"):
            status = "Validated with remarks" if _issues(s, node) or pending_text.strip() or status == "Validated with remarks" else "Validated"
        elif status in ("Validated", "Validated with remarks"):
            st.warning("Marque Validation completed para registrar a conclusão ou escolha um status em andamento.")
            return
        reviews = copy.deepcopy(s["reviews"])
        reviews[node["id"]] = {**review, "status": status, "completed": completed, "owner": owner, "comments": comments, "pending": pending_text, "final_comments": final_comments, "pending_owner": pending_owner,
                                "validated_at": datetime.now(timezone.utc).isoformat(timespec="seconds") if completed else ""}
        if _refresh(s, edits=_merge_edits(s["edits"], edits), reviews=reviews):
            s["notice"] = "Revisão salva somente na memória desta sessão. Exporte o Excel antes de encerrar."
            st.rerun()


def _columns(s: dict, node: dict) -> None:
    columns = [n for n in _nodes(s) if n["level"] == "Column" and n["object_id"] == node["id"]]
    if not columns:
        st.info("Nenhuma coluna declarada neste objeto.")
        return
    left, right = st.columns([3, 1.5])
    search = left.text_input("Pesquisar colunas", key=_key(f"col_search_{node['id']}"))
    status = right.selectbox("Situação da coluna", ["Todas", "Com pendências", *STATUSES], key=_key(f"col_status_{node['id']}"))
    filtered = [c for c in columns if search.casefold() in (c["name"] + " " + _text(c.get("description"))).casefold() and (status == "Todas" or status == "Com pendências" and _issues(s, c) or _status(s, c) == status)]
    privacy = ["fl_reg_apl", "fl_atr_pii", "nm_atr_pii", "fl_atr_sens", "fl_atr_mask", "nm_atr_class_inf", "fl_metadata_ai_generate"]
    present = [tag for tag in privacy if any(tag in c.get("tags", {}) for c in filtered)]
    rows = [{"Coluna": c["name"], "Tipo": c.get("data_type", c.get("type", "")), "Descrição atual": c.get("original_description"), "Descrição proposta": c.get("description"), **{tag: c.get("tags", {}).get(tag, "Não informado") for tag in present}, "Status": _status(s, c), "Pendências": len(_issues(s, c))} for c in filtered]
    if not rows:
        st.info("Nenhuma coluna corresponde aos filtros.")
        return
    st.dataframe(rows, hide_index=True, use_container_width=True)
    col = st.selectbox("Coluna para revisar", [c["id"] for c in filtered], format_func=lambda nid: next(c["name"] for c in filtered if c["id"] == nid), key=_key(f"column_select_{node['id']}"))
    st.button("Abrir revisão da coluna", key=_key(f"open_column_{node['id']}"), on_click=_go_review, args=(s, col))
    st.caption("A revisão da tabela ou view não conclui automaticamente suas colunas. Tags ausentes não recebem valores fictícios.")


def _review_page(s: dict) -> None:
    st.markdown('<div class="eyebrow">REVISÃO DE METADADOS</div>', unsafe_allow_html=True)
    nodes = _nodes(s)
    if not nodes:
        st.title("Revisar metadados")
        st.info("Importe seus arquivos na Visão geral para iniciar a revisão.")
        return
    ids = [n["id"] for n in nodes]
    if s["selected"] not in ids:
        s["selected"] = next((n["id"] for n in nodes if n["level"] in ("Table", "View")), ids[0])
    selected = st.selectbox("Objeto selecionado", ids, index=ids.index(s["selected"]), format_func=lambda nid: next(f"{n['level']} · {n['path']}" for n in nodes if n["id"] == nid), key=_key(f"review_selection_{s['selected']}"))
    if selected != s["selected"]:
        s["selected"] = selected
    node = next(n for n in nodes if n["id"] == s["selected"])
    st.title(_md(node["name"]))
    st.caption(_md(node["path"]))
    st.markdown(f'<span class="status-pill">{_safe(_status(s, node))}</span>', unsafe_allow_html=True)
    cells = [("Catálogo", node["catalog"]), ("Schema / camada", node["schema"] or "—"), ("Nível técnico", node["level"]), ("Arquivo de origem", node["filename"]), ("Domínio", node.get("tags", {}).get("nm_dom", "Não informado")), ("Subdomínio", node.get("tags", {}).get("nm_subdom", "Não informado"))]
    if node.get("tags", {}).get("nm_purpose"):
        cells.append(("Finalidade", node["tags"]["nm_purpose"]))
    if node.get("last_modified") or node.get("updated_at"):
        cells.append(("Última alteração informada", node.get("last_modified", node.get("updated_at"))))
    st.markdown('<div class="tech-grid">' + "".join(f'<div><div class="tech-label">{_safe(label)}</div><div class="tech-value">{_safe(value)}</div></div>' for label, value in cells) + "</div>", unsafe_allow_html=True)
    if node.get("schema") in LAYERS:
        st.caption(f"{node['schema']}: {LAYERS[node['schema']]} O significado operacional depende dos padrões internos da empresa.")
    tabs = st.tabs(["Visão geral", "Tags de governança", "Colunas", f"Pendências ({len(_issues(s, node))})"])
    with tabs[0]:
        _business_review(s, node)
    with tabs[1]:
        _tag_editor(s, node)
    with tabs[2]:
        if node["level"] in ("Table", "View"):
            _columns(s, node)
        elif node["level"] == "Column":
            st.caption(f"Tipo técnico: {node.get('data_type', node.get('type', 'Não informado'))}")
            st.info("Edite a descrição e a revisão individual em Visão geral e as tags desta coluna em Tags de governança.")
            parent = next((n for n in nodes if n["id"] == node.get("object_id") and n["level"] in ("Table", "View")), None)
            if parent:
                st.button("Voltar ao objeto de dados", key=_key("parent_object"), on_click=_go_review, args=(s, parent["id"]))
        else:
            st.info("Selecione uma tabela ou view para revisar suas colunas. A aprovação deste nível é independente.")
    with tabs[3]:
        _show_issues(_issues(s, node), key=node["id"])


def _downloads(s: dict) -> None:
    st.markdown('<div class="eyebrow">IMPORTAÇÃO E ENTREGA</div>', unsafe_allow_html=True)
    st.title("Arquivos e downloads")
    st.caption("Prepare a validação em Excel e converta somente as revisões aprovadas em novos YAMLs.")
    _upload(s)
    if not s["files"]:
        return
    _import_report(s)
    with st.expander("Arquivos nesta sessão e controle de versões"):
        st.dataframe([{"Arquivo": f["name"], "Tamanho aproximado (KiB)": round(len(f["content"]) * 3 / 4 / 1024, 1)} for f in s["files"]], hide_index=True, use_container_width=True)
        selected = st.selectbox("Origem para remover", range(len(s["files"])), format_func=lambda i: s["files"][i]["name"], key=_key("remove_source"))
        confirmed = st.checkbox("Confirmo remover esta origem e revalidar o lote", key=_key("remove_confirm"))
        if st.button("Remover origem selecionada", disabled=not confirmed, key=_key("remove_file")):
            files = [f for i, f in enumerate(s["files"]) if i != selected]
            try:
                retained_ids = {n["id"] for n in prepare(files, [], {}, s["refs"])["nodes"]}
                edits = [e for e in s["edits"] if e["id"] in retained_ids]
                reviews = {nid: r for nid, r in s["reviews"].items() if nid in retained_ids}
                if _refresh(s, files=files, edits=edits, reviews=reviews):
                    st.rerun()
            except ValueError as exc:
                st.error(_md(exc))
    st.subheader("1 · Planilha para o Data Owner")
    st.caption("Cinco abas: Read Me, Objects, Tags, Columns e Approval. Células editáveis, listas controladas e identificação estável dos objetos.")
    if st.button("Export Excel for Data Owner", type="primary", key=_key("export_excel")):
        try:
            with st.spinner("Preparando o Excel em memória…"):
                s["xlsx"] = export_xlsx(s["files"], s["edits"], s["reviews"], s["refs"])
        except ValueError as exc:
            st.error(_md(exc))
        except Exception:
            st.error("Não foi possível gerar a planilha. Confira o relatório de importação.")
    if s.get("xlsx"):
        st.download_button("Baixar planilha de validação", data=s["xlsx"], file_name="metadata_review.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=_key("download_excel"), use_container_width=True)
    st.divider()
    st.subheader("2 · Importar a revisão preenchida")
    st.caption("Reimporte primeiro os mesmos YAMLs de origem. A comparação usa identificadores e o conteúdo original, sem depender da ordem das linhas.")
    with st.form(_key("reimport_form")):
        workbook = st.file_uploader("Excel preenchido pelo Data Owner", type=["xlsx"], key=_key("xlsx_upload"))
        submitted = st.form_submit_button("Import Validated Excel", use_container_width=True)
    if submitted:
        if workbook is None:
            st.warning("Selecione a planilha XLSX validada.")
        elif workbook.size > 15 * 1024 * 1024:
            st.error("A planilha excede o limite de 15 MiB.")
        else:
            try:
                with st.spinner("Comparando alterações, conflitos e regras…"):
                    s["preview"] = preview_xlsx(workbook.getvalue(), s["files"], s["edits"], s["reviews"], s["refs"])
            except ValueError as exc:
                st.error(_md(exc))
            except Exception:
                st.error("A planilha não pôde ser lida. Use uma planilha gerada por esta aplicação.")
    preview = s.get("preview")
    if preview:
        changes = preview.get("changes", preview.get("edits", []))
        rejected = preview.get("rejected", [])
        missing = preview.get("missing", [])
        st.markdown("#### Comparação antes de aplicar")
        a, b, c = st.columns(3)
        a.metric("Alterações propostas", len(changes))
        b.metric("Registros rejeitados", len(rejected))
        c.metric("Registros ausentes", len(missing))
        with st.expander("Alterações aceitas para revisão", expanded=True):
            _show_records(changes)
        if rejected:
            with st.expander("Alterações rejeitadas", expanded=True):
                _show_records(rejected)
        if missing:
            with st.expander("Registros ausentes ou não encontrados", expanded=True):
                _show_records(missing)
        if preview.get("issues"):
            with st.expander("Pendências após as alterações propostas"):
                _show_issues(preview["issues"], key="preview")
        confirm = st.checkbox("Revisei a comparação e autorizo aplicar as alterações aceitas à sessão", key=_key(f"apply_preview_confirm_{s['revision']}"))
        if st.button("Aplicar alterações aceitas", disabled=not confirm, key=_key("apply_preview"), type="primary"):
            reviews = copy.deepcopy(s["reviews"])
            for nid, review in preview.get("reviews", {}).items():
                reviews[nid] = {**reviews.get(nid, {}), **review}
            if _refresh(s, edits=_merge_edits(s["edits"], preview.get("edits", [])), reviews=reviews):
                s["notice"] = "Alterações aceitas aplicadas e revalidadas. Confira as pendências antes de gerar os YAMLs."
                st.rerun()
    st.divider()
    st.subheader("3 · Gerar pacote de YAMLs aprovados")
    st.warning("Somente objetos com status Validated e sem pendências técnicas são elegíveis. Fontes que misturam objetos aprovados e não aprovados são excluídas por inteiro para preservar a estrutura original. O pacote contém o relatório de alterações e exclusões.")
    excluded = [n for n in _nodes(s) if n["level"] in ("Table", "View") and _status(s, n) != "Validated"]
    if excluded:
        with st.expander(f"Objetos ainda não elegíveis ({len(excluded)})"):
            _show_records([{"Objeto": n["path"], "Status": _status(s, n), "Pendências": len(_issues(s, n)), "Arquivo": n["filename"]} for n in excluded])
    confirmed = st.checkbox("Confirmo gerar novas versões apenas dos arquivos elegíveis; os originais não serão sobrescritos", key=_key(f"zip_confirm_{s['revision']}"))
    if st.button("Gerar YAMLs atualizados (ZIP)", disabled=not confirmed, key=_key("generate_zip"), type="primary"):
        try:
            with st.spinner("Revalidando e gerando o ZIP em memória…"):
                s["zip"] = final_zip(s["files"], s["edits"], s["reviews"], s["refs"], confirmed=True)
        except ValueError as exc:
            st.error(_md(exc))
        except Exception:
            st.error("Não foi possível gerar o pacote. Confira a revisão e os relatórios.")
    if s.get("zip"):
        st.download_button("Baixar YAMLs aprovados e relatórios", data=s["zip"], file_name="metadata_review_approved.zip", mime="application/zip", key=_key("download_zip"), use_container_width=True)
    with st.expander("Relatório de alterações da sessão"):
        _show_records((s.get("prepared") or {}).get("report", []), limit=1000)


def _references(s: dict) -> None:
    st.markdown('<div class="eyebrow">PADRÃO DE GOVERNANÇA</div>', unsafe_allow_html=True)
    st.title("Regras e referências")
    st.caption("Valores controlados, orientações de preenchimento e listas oficiais externas.")
    st.info("Catálogo conferido com o PDF DAI-Tags - Databricks Unity Catalog-101026-010433.pdf, que não é distribuído com a aplicação. A imagem de referência não foi disponibilizada; a hierarquia é derivada dos arquivos importados.")
    st.caption("Divergência documentada: a seção View do PDF registra ct_risco*. A aplicação utiliza ct_risco, conforme a transcrição funcional fornecida; a Governança deve confirmar essa grafia.")
    st.warning("Pendência global da Governança: o PDF define [G], [T] e [REDUNDANTE], mas não atribui a categoria de cada tag. Essas categorias permanecem como Não identificada e não são inferidas; a lacuna é distinta das inconsistências técnicas dos objetos.")
    left, right = st.columns([1.2, 1])
    with left:
        with st.container(border=True):
            st.subheader("Listas oficiais externas")
            st.caption("Um arquivo JSON com listas de textos. Não use valores inventados; listas vazias continuam pendentes.")
            for key, label in [("nm_dom", "Domínios"), ("nm_subdom", "Subdomínios"), ("nm_sist_orig", "Sistemas de origem"), ("ct_area_neg", "Áreas de negócio")]:
                values = s["refs"].get(key, [])
                st.write(f"**{label}** · {len(values)} valores" if values else f"**{label}** · Lista oficial não carregada")
            st.code('{\n  "nm_dom": [],\n  "nm_subdom": [],\n  "nm_sist_orig": [],\n  "ct_area_neg": []\n}', language="json")
            with st.form(_key("refs_form")):
                reference = st.file_uploader("Arquivo de listas oficiais (.json)", type=["json"], key=_key("refs_upload"))
                apply = st.form_submit_button("Carregar e revalidar listas oficiais", type="primary", use_container_width=True)
            if apply:
                if reference is None:
                    st.warning("Selecione um arquivo JSON de referência.")
                elif reference.size > 2 * 1024 * 1024:
                    st.error("O arquivo de referência excede 2 MiB.")
                else:
                    try:
                        refs = json.loads(reference.getvalue().decode("utf-8-sig"))
                        keys = {"nm_dom", "nm_subdom", "nm_sist_orig", "ct_area_neg"}
                        if not isinstance(refs, dict) or set(refs) - keys:
                            raise ValueError("Use somente as chaves nm_dom, nm_subdom, nm_sist_orig e ct_area_neg.")
                        if any(not isinstance(v, list) or len(v) > 10000 or any(not isinstance(x, str) or not x.strip() or len(x) > 250 for x in v) for v in refs.values()):
                            raise ValueError("Cada chave deve conter uma lista de textos não vazios, com até 10.000 valores.")
                        refs = {k: list(dict.fromkeys(v)) for k, v in refs.items()}
                        get_catalog(refs)
                    except (ValueError, UnicodeDecodeError):
                        st.error("Referência inválida. Use o formato JSON mostrado acima, com listas de textos oficiais.")
                    else:
                        if s["files"]:
                            ok = _refresh(s, refs=refs)
                        else:
                            s["refs"] = refs
                            ok = True
                        if ok:
                            s["notice"] = "Listas oficiais carregadas somente nesta sessão e validações atualizadas."
                            st.rerun()
    with right:
        with st.container(border=True):
            st.subheader("Entenda a hierarquia")
            st.caption("Catalog → Schema → Table/View → Column")
            st.write("**Catalog** organiza schemas. **Schema** agrupa ativos. **Table/View** são ativos de dados. **Column** representa seus atributos.")
            for layer, meaning in LAYERS.items():
                st.markdown(f"**{layer}**")
                st.caption(meaning)
            st.caption("O significado operacional de cada camada depende dos padrões internos da empresa. Camada de dados e nível de aplicação da tag são conceitos distintos.")
    st.subheader("Catálogo de tags por nível")
    st.caption("[G] Governed Tag: opções oficiais. [T] Tag Livre: formato definido. [REDUNDANTE]: repetição documentada em outro nível. Uma categoria desconhecida não é inferida.")
    for level, tags in get_catalog(s["refs"]).items():
        with st.expander(f"{level} · {len(tags)} tags"):
            rows = []
            for tag, rule in tags.items():
                allowed = rule.get("values")
                rows.append({"Tag": tag, "Significado": rule["description"], "Categoria": rule["category"], "Valores / formato": ", ".join(allowed) if allowed is not None else "Lista oficial não carregada" if rule.get("external") else rule.get("format", "Texto livre")})
            st.dataframe(rows, hide_index=True, use_container_width=True)
    with st.expander("Validação funcional e técnica"):
        st.write("Validation completed registra a conclusão individual do trabalho de revisão. A conformidade técnica depende dos valores permitidos, formatos e regras de consistência. Pendências continuam visíveis após a conclusão; nesses casos, o status é Validated with remarks.")
        st.write("Pending: revisão pendente. In review: revisão em andamento. Awaiting clarification: aguardando esclarecimento. Validated: concluído sem pendências identificadas. Validated with remarks: concluído com ressalvas. Rejected: rejeitado pelo responsável.")
        st.caption("O responsável é informado manualmente e não é uma identidade autenticada. A aplicação registra a data da ação em UTC.")


def run() -> None:
    st.set_page_config(page_title="Metadata Review · Governança de dados", page_icon="◈", layout="wide", initial_sidebar_state="expanded")
    st.markdown(STYLE, unsafe_allow_html=True)
    s = _state()
    if s.get("pending_navigation"):
        st.session_state[_key("navigation")] = s.pop("pending_navigation")
    page = _sidebar(s)
    if s.get("notice"):
        st.success(s.pop("notice"))
    rejected = [r for r in (s.get("prepared") or {}).get("report", []) if r.get("status") == "rejected"]
    if rejected:
        st.warning(f"{len(rejected)} proposta(s) não foram aplicadas. Corrija o valor na revisão ou descarte as propostas rejeitadas antes de exportar.")
        with st.expander("Propostas rejeitadas e motivos", expanded=True):
            _show_records(rejected)
            if st.button("Descartar propostas rejeitadas", key=_key("discard_rejected")):
                keys = {(r.get("id"), r.get("field"), r.get("tag")) for r in rejected}
                edits = [e for e in s["edits"] if (e.get("id"), e.get("field"), e.get("tag")) not in keys]
                if _refresh(s, edits=edits):
                    st.rerun()
    if page == "Visão geral":
        _dashboard(s)
    elif page == "Revisar metadados":
        _review_page(s)
    elif page == "Arquivos e downloads":
        _downloads(s)
    else:
        _references(s)
    st.markdown('<div class="page-foot">Metadata Review · Processamento em memória no servidor local · Nenhuma alteração no Databricks · Exporte o Excel antes de encerrar a sessão.</div>', unsafe_allow_html=True)

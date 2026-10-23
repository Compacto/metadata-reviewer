"""Two conversions for an offline owner review; uploaded bytes stay in memory."""
from __future__ import annotations

import base64
import hashlib
import re

import streamlit as st

from app import service
from app.engine import MAX_BATCH_BYTES, MAX_FILE_BYTES, MAX_FILES
from app.privacy import release_session_files


STYLE = """<style>
:root { --navy:#142b40; --teal:#087f82; }
.stApp {background:#f5f7fa;}
[data-testid="stHeader"] {background:#f5f7fae8;}
.block-container {max-width:1150px;padding-top:3.5rem;padding-bottom:2rem;}
h1 {color:#142b40;letter-spacing:-1px;font-size:2.3rem!important;}
h2,h3 {color:#23465d;letter-spacing:-.4px;}
[data-testid="stFileUploaderDropzone"] {background:#f0f8f6;border:1px dashed #a7d2c5;}
[data-testid="stForm"] {border:0;padding:0;}
[data-testid="stVerticalBlockBorderWrapper"]>div {border-color:#dfe8ee!important;border-radius:10px;}
.stButton>button,.stDownloadButton>button {border-radius:7px;}
[data-baseweb="tab-list"] {gap:24px;border-bottom:1px solid #dfe7ed;}
.stTabs [data-baseweb="tab-panel"] {padding-top:18px;}
.journey {height:100%;background:white;border:1px solid #dfe8ee;border-radius:10px;padding:18px 20px;margin:12px 0 24px;}
.journey strong {display:block;font-size:.91rem;color:#24485f;margin-bottom:7px;}
.journey p {font-size:.8rem;color:#667f90;line-height:1.6;margin:0;}
.step {color:#087f82;font-weight:700;margin-right:7px;}
.page-foot {font-size:.75rem;color:#718a9b;border-top:1px solid #dfe7ed;margin-top:28px;padding-top:16px;}
@media(max-width:700px) {.block-container{padding-left:1rem;padding-right:1rem;}.journey{margin:3px 0 8px;}h1{font-size:1.8rem!important;}}
</style>"""


def _state() -> dict:
    if "mr" not in st.session_state:
        st.session_state.mr = {"excel": None, "yaml": None}
    elif set(st.session_state.mr) != {"excel", "yaml"}:
        _reset()
    if "mr_epoch" not in st.session_state:
        st.session_state.mr_epoch = 0
    return st.session_state.mr


def _key(name: str) -> str:
    return f"mr_{st.session_state.mr_epoch}_{name}"


def _md(value) -> str:
    """Imported names and diagnostics must never become Markdown images/links."""
    return re.sub(r"([\\`*_{}\[\]()#+.!>|~-])", r"\\\1", str(value))


def _fingerprint(uploads) -> str:
    digest = hashlib.sha256()
    for upload in uploads or []:
        name = upload.name.encode("utf-8")
        content = upload.getvalue()
        digest.update(len(name).to_bytes(8, "big"))
        digest.update(name)
        digest.update(len(content).to_bytes(8, "big"))
        digest.update(content)
    return digest.hexdigest()


def _upload_problem(uploads) -> str | None:
    if len(uploads or []) > MAX_FILES:
        return f"Envie até {MAX_FILES} arquivos por conversão."
    if any(upload.size > MAX_FILE_BYTES for upload in uploads or []):
        return f"Cada arquivo deve ter até {MAX_FILE_BYTES // (1024 * 1024)} MB."
    if sum(upload.size for upload in uploads or []) > MAX_BATCH_BYTES:
        return f"O lote deve ter até {MAX_BATCH_BYTES // (1024 * 1024)} MB no total."
    return None


def _forget(slot: str) -> None:
    _state()[slot] = None


def _reset() -> None:
    release_session_files()
    epoch = st.session_state.get("mr_epoch", 0) + 1
    for key in list(st.session_state):
        if key == "mr" or key == "mr_epoch" or key.startswith("mr_"):
            del st.session_state[key]
    st.session_state.mr_epoch = epoch
    st.session_state.mr = {"excel": None, "yaml": None}


def _count(summary: dict, field: str) -> int:
    value = summary.get(field, 0)
    return value if isinstance(value, int) and value >= 0 else 0


def _warnings(warnings) -> None:
    if not warnings:
        return
    with st.expander(f"Pontos para Governança · {len(warnings)}", expanded=False):
        st.caption("Avisos de metadados para conferência. Eles não substituem a revisão do Owner.")
        for warning in warnings[:60]:
            if isinstance(warning, dict):
                message = warning.get("message", "Confira este metadado com a Governança.")
                context = warning.get("path") or warning.get("filename")
                st.warning(_md(message))
                if context:
                    st.caption(_md(context))
            else:
                st.warning(_md(warning))
        if len(warnings) > 60:
            st.caption(f"Exibindo 60 de {len(warnings)} avisos.")


def _current_result(state: dict, slot: str, fingerprint: str):
    result = state.get(slot)
    if result and result.get("fingerprint") != fingerprint:
        state[slot] = None
        return None
    return result


def _to_excel(state: dict) -> None:
    st.subheader("Prepare o Excel para o Owner")
    st.write("Envie um YAML, vários YAMLs ou um lote ZIP. O Excel organiza as informações e indica os campos que o Owner pode alterar.")
    uploads = st.file_uploader(
        "Selecione YAMLs ou lotes ZIP",
        type=["yaml", "yml", "zip"],
        accept_multiple_files=True,
        key=_key("yaml_upload"),
        on_change=_forget,
        args=("excel",),
        help="Arquivos processados apenas na memória da sessão. Cada upload pode ter até 10 MB.",
    )
    problem = _upload_problem(uploads)
    fingerprint = "invalid" if problem else _fingerprint(uploads)
    _current_result(state, "excel", fingerprint)
    if problem:
        st.error(problem)
    with st.form(_key("excel_form")):
        submitted = st.form_submit_button("Gerar Excel para o Owner", type="primary", use_container_width=True, disabled=bool(problem))
    if submitted:
        state["excel"] = None
        if not uploads:
            st.warning("Selecione ao menos um arquivo YAML ou ZIP para gerar o Excel.")
        else:
            try:
                with st.spinner("Organizando os metadados no Excel…"):
                    files = [{"name": f.name, "content": base64.b64encode(f.getvalue()).decode("ascii")} for f in uploads]
                    result = service.create_owner_excel(files)
                state["excel"] = {"fingerprint": fingerprint, **result}
            except ValueError as exc:
                st.error(_md(exc))
            except Exception:
                st.error("Não foi possível gerar o Excel. Confira os arquivos e tente novamente.")
    result = _current_result(state, "excel", fingerprint)
    if result:
        summary = result.get("summary", {})
        st.success(f"Excel pronto: {_count(summary, 'files')} YAML(s), {_count(summary, 'objects')} objeto(s) e {_count(summary, 'columns')} coluna(s).")
        st.download_button(
            "Baixar Excel para o Owner",
            data=result["content"],
            file_name=result["filename"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=_key("excel_download"),
            type="primary",
            use_container_width=True,
        )
        st.caption("Envie este Excel ao Owner. Após a revisão, ele deve devolver a planilha à Governança.")
        _warnings(result.get("warnings", []))


def _to_yaml(state: dict) -> None:
    st.caption("Use texto ou valores nos campos preenchidos. Se preparar respostas com fórmulas, use Colar especial > Valores no Excel antes de enviar.")
    st.subheader("Gere os YAMLs revisados")
    st.write("Governança: anexe o Excel devolvido pelo Owner para gerar os YAMLs com as alterações da planilha.")
    st.caption("Use um Excel gerado por este app. Não é necessário anexar os YAMLs originais.")
    uploads = st.file_uploader(
        "Selecione os Excels revisados",
        type=["xlsx"],
        accept_multiple_files=True,
        key=_key("excel_upload"),
        on_change=_forget,
        args=("yaml",),
        help="Você pode converter uma planilha ou várias planilhas devolvidas pelos Owners.",
    )
    problem = _upload_problem(uploads)
    fingerprint = "invalid" if problem else _fingerprint(uploads)
    _current_result(state, "yaml", fingerprint)
    if problem:
        st.error(problem)
    with st.form(_key("yaml_form")):
        submitted = st.form_submit_button("Gerar YAMLs", type="primary", use_container_width=True, disabled=bool(problem))
    if submitted:
        state["yaml"] = None
        if not uploads:
            st.warning("Selecione ao menos um Excel revisado para gerar os YAMLs.")
        else:
            try:
                with st.spinner("Aplicando as alterações do Excel aos YAMLs…"):
                    workbooks = [{"name": f.name, "content": f.getvalue()} for f in uploads]
                    result = service.convert_owner_excels(workbooks)
                state["yaml"] = {"fingerprint": fingerprint, **result}
            except ValueError as exc:
                st.error(_md(exc))
            except Exception:
                st.error("Não foi possível gerar os YAMLs. Confira se os Excels foram gerados por este app e tente novamente.")
    result = _current_result(state, "yaml", fingerprint)
    if result:
        summary = result.get("summary", {})
        st.success(f"YAMLs prontos: {_count(summary, 'files')} arquivo(s), {_count(summary, 'objects')} objeto(s) e {_count(summary, 'changes')} alteração(ões).")
        st.download_button(
            "Baixar YAMLs em ZIP",
            data=result["content"],
            file_name=result["filename"],
            mime="application/zip",
            key=_key("zip_download"),
            type="primary",
            use_container_width=True,
        )
        yamls = result.get("yamls", [])
        if len(yamls) == 1:
            st.download_button(
                "Baixar YAML individual",
                data=yamls[0]["content"],
                file_name=yamls[0]["name"].rsplit("/", 1)[-1],
                mime="application/yaml",
                key=_key("yaml_download"),
                use_container_width=True,
            )
        st.caption("O analista de Governança pode baixar os arquivos e realizar a ingestão no ambiente.")
        _warnings(result.get("warnings", []))


def run() -> None:
    st.set_page_config(page_title="Metadata Review · YAML e Excel", page_icon="◈", layout="wide", initial_sidebar_state="collapsed")
    st.markdown(STYLE, unsafe_allow_html=True)
    state = _state()
    title, reset = st.columns([4, 1])
    with title:
        st.title("Metadata Review")
        st.caption("Metadados claros para o Owner. YAMLs prontos para a Governança.")
    with reset:
        st.button("Limpar sessão", key=_key("clear"), on_click=_reset, use_container_width=True, help="Descarta os uploads e downloads desta sessão.")
    steps = [
        ("1", "Gere o Excel", "Anexe YAMLs ou um ZIP e baixe a planilha de revisão."),
        ("2", "Owner revisa", "O Owner altera os campos indicados no Excel e devolve à Governança."),
        ("3", "Gere os YAMLs", "A Governança importa o Excel revisado e baixa os YAMLs para ingestão."),
    ]
    for column, (number, label, description) in zip(st.columns(3), steps):
        with column:
            st.markdown(f'<div class="journey"><strong><span class="step">{number}</span>{label}</strong><p>{description}</p></div>', unsafe_allow_html=True)
    excel_tab, yaml_tab = st.tabs(["1 · YAML para Excel", "2 · Excel para YAML"])
    with excel_tab:
        _to_excel(state)
    with yaml_tab:
        _to_yaml(state)
    with st.expander("Privacidade dos arquivos", expanded=False):
        st.write("Uploads e conversões são processados na memória da sessão. Os arquivos não são gravados no servidor nem enviados a serviços externos. Use Limpar sessão para descartar os dados carregados e os downloads gerados.")
        st.write("O Excel contém os metadados necessários para gerar os YAMLs depois; trate-o como documento corporativo. Guarde e compartilhe os downloads conforme as regras da sua organização.")
    st.markdown('<div class="page-foot">Processamento em memória · Revisão pelo Owner no Excel · Ingestão pelo analista de Governança</div>', unsafe_allow_html=True)

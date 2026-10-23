"""Release this tab's framework-owned byte buffers on explicit session reset.

These APIs are covered against the pinned Streamlit release. No user content is
written, logged, or retained by this helper; other active sessions keep their
own uploads and download references.
"""
from streamlit import runtime
from streamlit.runtime.scriptrunner import get_script_run_ctx


def release_session_files() -> bool:
    context = get_script_run_ctx(suppress_warning=True)
    if context is None or not runtime.exists():
        return False
    instance = runtime.get_instance()
    instance.uploaded_file_mgr.remove_session_files(context.session_id)
    instance.media_file_mgr.clear_session_refs(context.session_id)
    # Streamlit gives downloadable files a mark-then-delete grace pass. An
    # explicit reset revokes those downloads immediately, so finish both passes.
    instance.media_file_mgr.remove_orphaned_files()
    instance.media_file_mgr.remove_orphaned_files()
    return True

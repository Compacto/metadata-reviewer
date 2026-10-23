"""Exercise the installed upload/download transport with disk writes denied."""
import builtins
from contextlib import ExitStack, contextmanager
from types import SimpleNamespace
from unittest.mock import patch

import pytest
import streamlit
from streamlit import config
from streamlit.runtime.media_file_manager import MediaFileManager
from streamlit.runtime.media_file_storage import MediaFileKind, MediaFileStorageError
from streamlit.runtime.memory_media_file_storage import MemoryMediaFileStorage
from streamlit.runtime.memory_uploaded_file_manager import MemoryUploadedFileManager
from streamlit.runtime.uploaded_file_manager import UploadedFileRec
from streamlit.web.server.media_file_handler import MediaFileHandler
from streamlit.web.server.upload_file_request_handler import UploadFileRequestHandler
from tornado.testing import AsyncHTTPTestCase
from tornado.web import Application

from app import privacy


@contextmanager
def deny_disk_writes():
    original_open = builtins.open
    def guarded_open(file, mode="r", *args, **kwargs):
        if any(flag in mode for flag in "wax+"):
            raise AssertionError("Transport attempted to write uploaded content to disk")
        return original_open(file, mode, *args, **kwargs)
    with ExitStack() as stack:
        for target in ("tempfile.TemporaryFile", "tempfile.NamedTemporaryFile", "tempfile.mkstemp", "tempfile.SpooledTemporaryFile"):
            stack.enter_context(patch(target, side_effect=AssertionError("Transport attempted temporary-file storage")))
        stack.enter_context(patch("builtins.open", side_effect=guarded_open))
        yield


class UploadTransportTests(AsyncHTTPTestCase):
    def get_app(self):
        # Resolve framework configuration before the I/O-denial assertion scope.
        config.get_option("server.maxUploadSize")
        self.manager = MemoryUploadedFileManager("/upload")
        self.media = MemoryMediaFileStorage("/media")
        MediaFileHandler.initialize_storage(self.media)
        return Application([
            (r"/upload/(?P<session_id>[^/]+)/(?P<file_id>[^/]+)", UploadFileRequestHandler,
             {"file_mgr": self.manager, "is_active_session": lambda session: session == "synthetic-session"}),
            (r"/media/(.*)", MediaFileHandler, {"path": ""}),
        ], xsrf_cookies=False)

    def test_more_than_one_megabyte_upload_stays_in_memory_in_actual_http_handler(self):
        assert streamlit.__version__ == "1.50.0", "Re-audit transport before changing the Streamlit pin."
        content = b"synthetic metadata transport bytes; " * 70_000
        assert len(content) > 1024 * 1024
        boundary = "SYNTHETICBOUNDARY"
        body = (f"--{boundary}\r\nContent-Disposition: form-data; name=\"file\"; filename=\"synthetic.yaml\"\r\n"
                "Content-Type: application/yaml\r\n\r\n").encode() + content + f"\r\n--{boundary}--\r\n".encode()
        with deny_disk_writes():
            response = self.fetch("/upload/synthetic-session/file-one", method="PUT", body=body,
                                  headers={"Content-Type": f"multipart/form-data; boundary={boundary}"})
            assert response.code == 204
            records = self.manager.get_files("synthetic-session", ["file-one"])
            assert len(records) == 1 and records[0].data == content
            self.manager.remove_session_files("synthetic-session")
            assert not self.manager.get_files("synthetic-session", ["file-one"])

    def test_download_is_served_from_bytes_without_reading_or_writing_a_file(self):
        content = b"synthetic xlsx transport bytes; " * 70_000
        with deny_disk_writes():
            identifier = self.media.load_and_get_id(content, "application/octet-stream", MediaFileKind.DOWNLOADABLE, "synthetic.xlsx")
            response = self.fetch(self.media.get_url(identifier))
            assert response.code == 200 and response.body == content
            assert "attachment" in response.headers["Content-Disposition"]


def test_reset_releases_current_session_framework_bytes_and_preserves_another_active_session():
    uploads = MemoryUploadedFileManager("/upload")
    uploads.add_file("current", UploadedFileRec("file-current", "synthetic.yaml", "application/yaml", b"current synthetic"))
    uploads.add_file("other", UploadedFileRec("file-other", "synthetic.yaml", "application/yaml", b"other synthetic"))
    storage = MemoryMediaFileStorage("/media")
    media = MediaFileManager(storage)
    with patch("streamlit.runtime.media_file_manager._get_session_id", side_effect=["current", "other"]):
        current_url = media.add(b"current synthetic", "application/octet-stream", "coordinate", "current.zip", is_for_static_download=True)
        other_url = media.add(b"other synthetic", "application/octet-stream", "coordinate", "other.zip", is_for_static_download=True)
    instance = SimpleNamespace(uploaded_file_mgr=uploads, media_file_mgr=media)
    with patch.object(privacy, "get_script_run_ctx", return_value=SimpleNamespace(session_id="current")), \
            patch.object(privacy.runtime, "exists", return_value=True), \
            patch.object(privacy.runtime, "get_instance", return_value=instance):
        assert privacy.release_session_files()
    assert not uploads.get_files("current", ["file-current"])
    assert uploads.get_files("other", ["file-other"])[0].data == b"other synthetic"
    with pytest.raises(MediaFileStorageError):
        storage.get_file(current_url.rsplit("/", 1)[-1])
    assert storage.get_file(other_url.rsplit("/", 1)[-1]).content == b"other synthetic"


@pytest.mark.parametrize("context,exists", [(None, False), (None, True), (SimpleNamespace(session_id="synthetic"), False)])
def test_reset_without_active_runtime_has_no_side_effects(context, exists):
    with patch.object(privacy, "get_script_run_ctx", return_value=context), \
            patch.object(privacy.runtime, "exists", return_value=exists), \
            patch.object(privacy.runtime, "get_instance") as get_instance:
        assert privacy.release_session_files() is False
        get_instance.assert_not_called()

"""Shared test fixtures."""

from unittest.mock import MagicMock, patch

import pytest

# The patch.dict(sys.modules) below drops every module a test imported.
# FastMCP 4 caches the prefab classes, so a prefab_ui imported again in a
# later test no longer matches the cache and preview cards serialize as plain
# dumps. Loaded once here, it stays. Not installed without the apps extra.
try:
    import prefab_ui.app  # noqa: F401
    import prefab_ui.components.base  # noqa: F401
except ImportError:
    pass


@pytest.fixture(autouse=True)
def mock_windows_com_modules():
    """Mock Windows-only COM modules so unit tests can run on non-Windows CI."""
    win32com = MagicMock()
    pythoncom = MagicMock()
    with patch.dict(
        "sys.modules",
        {
            "win32com": win32com,
            "win32com.client": win32com.client,
            "pythoncom": pythoncom,
        },
    ):
        yield


@pytest.fixture(autouse=True)
def answers_count_as_human(monkeypatch):
    """Fake dialogs answer at once. Count every answer as a person's, so the
    no-dialog note (confirmation.NO_DIALOG_MS) appears only in the tests that
    set the threshold themselves (tests/test_no_dialog.py)."""
    from mcp_sap_gui import confirmation
    monkeypatch.setattr(confirmation, "NO_DIALOG_MS", 0.0)

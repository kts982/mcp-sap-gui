"""Shared test fixtures."""

from unittest.mock import MagicMock, patch

import pytest


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

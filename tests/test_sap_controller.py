"""Tests for SAP GUI Controller."""

import pytest
from unittest.mock import MagicMock, patch


class TestSAPGUIController:
    """Tests for SAPGUIController class."""

    def test_init_without_pywin32(self):
        """Test that initialization fails gracefully without pywin32."""
        with patch.dict('sys.modules', {'win32com': None, 'win32com.client': None}):
            # Force reimport
            import importlib
            from mcp_sap_gui import sap_controller
            importlib.reload(sap_controller)

            with pytest.raises(sap_controller.SAPGUINotAvailableError):
                sap_controller.SAPGUIController()

    def test_is_connected_false_by_default(self):
        """Test that is_connected returns False initially."""
        with patch('win32com.client'):
            from mcp_sap_gui.sap_controller import SAPGUIController
            controller = SAPGUIController()
            assert controller.is_connected is False

    def test_require_session_raises_when_not_connected(self):
        """Test that operations fail when not connected."""
        with patch('win32com.client'):
            from mcp_sap_gui.sap_controller import (
                SAPGUIController,
                SAPGUINotConnectedError
            )
            controller = SAPGUIController()

            with pytest.raises(SAPGUINotConnectedError):
                controller.get_session_info()


class TestVKey:
    """Tests for VKey enum."""

    def test_vkey_values(self):
        """Test that VKey has expected values."""
        from mcp_sap_gui.sap_controller import VKey

        assert VKey.ENTER == 0
        assert VKey.F1 == 1
        assert VKey.F3 == 3
        assert VKey.F8 == 8
        assert VKey.F11 == 11
        assert VKey.F12 == 12


class TestSessionInfo:
    """Tests for SessionInfo dataclass."""

    def test_session_info_creation(self):
        """Test SessionInfo can be created."""
        from mcp_sap_gui.sap_controller import SessionInfo

        info = SessionInfo(
            system_name="DEV",
            system_number="00",
            client="100",
            user="TESTUSER",
            language="EN",
            transaction="MM03",
            program="SAPLMGMM",
            screen_number=100,
            session_number=0,
        )

        assert info.system_name == "DEV"
        assert info.client == "100"
        assert info.transaction == "MM03"

"""
SAP GUI Controller - COM automation wrapper for SAP GUI Scripting API.

This module provides a Python interface to SAP GUI for Windows via COM automation.
It wraps the SAP GUI Scripting API to enable programmatic interaction with SAP.

Requirements:
- Windows OS
- SAP GUI for Windows installed
- SAP GUI Scripting enabled (transaction RZ11, parameter sapgui/user_scripting)
- pywin32 package

Reference:
- SAP GUI Scripting API: https://help.sap.com/docs/sap_gui_for_windows
- SAP Note 480149 - SAP GUI Scripting Security
"""

# Re-export all public symbols so existing imports continue to work:
#   from mcp_sap_gui.sap_controller import SAPGUIController, VKey, ...
from .controller import SAPGUIControllerBase
from .discovery import DiscoveryMixin
from .fields import FieldsMixin
from .models import (  # noqa: F401
    _TOOLBAR_BUTTON_TYPES,
    SAPGUIError,
    SAPGUINotAvailableError,
    SAPGUINotConnectedError,
    ScreenElement,
    SessionInfo,
    VKey,
    _strip_tcode_prefix,
)
from .tables import TablesMixin
from .trees import TreesMixin


class SAPGUIController(
    FieldsMixin,
    TablesMixin,
    TreesMixin,
    DiscoveryMixin,
    SAPGUIControllerBase,
):
    """
    Controller for SAP GUI Scripting API via COM automation.

    This class provides methods to interact with SAP GUI programmatically,
    including connecting to systems, navigating transactions, reading/writing
    fields, and extracting data from tables.

    Composed from mixins:
    - SAPGUIControllerBase: connection, navigation, screen info
    - FieldsMixin: field read/write, buttons, checkboxes, combos, textedit
    - TablesMixin: ALV grid + TableControl reading, selection, cell ops
    - TreesMixin: tree reading, expand/collapse, node interaction
    - DiscoveryMixin: popups, toolbars, shell, screen elements, screenshots

    Example usage:
        controller = SAPGUIController()
        controller.connect("DEV - Development System")
        controller.execute_transaction("MM03")
        controller.set_field("wnd[0]/usr/ctxtRMMG1-MATNR", "MAT-001")
        controller.send_vkey(VKey.ENTER)
        description = controller.read_field("wnd[0]/usr/txtMAKT-MAKTX")
    """
    pass

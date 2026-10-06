"""
Data models, enums, and exceptions for SAP GUI Controller.

This module contains all shared types used across the controller modules.
"""

from dataclasses import dataclass
from enum import IntEnum


# SAP GUI Virtual Keys
class VKey(IntEnum):
    """SAP GUI virtual key codes."""
    ENTER = 0
    F1 = 1   # Help
    F2 = 2
    F3 = 3   # Back
    F4 = 4   # Dropdown/Search help
    F5 = 5   # Refresh
    F6 = 6
    F7 = 7
    F8 = 8   # Execute
    F9 = 9
    F10 = 10
    F11 = 11  # Save
    F12 = 12  # Cancel
    SHIFT_F1 = 13
    SHIFT_F2 = 14
    SHIFT_F3 = 15  # Exit
    SHIFT_F4 = 16
    SHIFT_F5 = 17
    SHIFT_F6 = 18
    SHIFT_F7 = 19
    SHIFT_F8 = 20
    SHIFT_F9 = 21
    CTRL_S = 11     # Save (same as F11)
    # Ctrl+letter codes per the VKey table of the scripting API guide
    # (7.60 and 8.10): 32-34 are Ctrl+F8..F10, not Ctrl+F/G/P.
    CTRL_F = 71     # Find
    CTRL_G = 84     # Continue search
    CTRL_P = 86     # Print
    ESC = 12        # Cancel (same as F12)


@dataclass
class SessionInfo:
    """Information about the current SAP session."""
    system_name: str
    system_number: str
    client: str
    user: str
    language: str
    transaction: str
    program: str
    screen_number: int
    session_number: int
    # SAP GUI for Windows release of the bound engine, e.g. "8.10 PL0".
    sap_gui_version: str = ""
    # The server allows read-only scripting (sapgui/user_scripting_set_readonly,
    # or sapgui/nwbc_scripting in SAP Business Client): nothing can be set.
    scripting_read_only: bool = False


@dataclass
class ScreenElement:
    """Information about a screen element.

    No ``visible`` flag: SAP GUI Scripting has no Visible property (the 8.10
    guide says so), so it read as True for every element.
    """
    id: str
    type: str
    name: str
    text: str
    changeable: bool


class SAPGUIError(Exception):
    """Exception raised for SAP GUI errors."""
    pass


class SAPGUINotAvailableError(SAPGUIError):
    """Exception raised when SAP GUI is not available."""
    pass


class SAPGUINotConnectedError(SAPGUIError):
    """Exception raised when not connected to SAP."""
    pass


# GetToolbarButtonType() returns strings per SAP GUI Scripting API v8.00,
# but some SAP GUI versions may return numeric values. Handle both.
_TOOLBAR_BUTTON_TYPES = {
    0: "Button", 1: "ButtonAndMenu", 2: "Menu",
    3: "Separator", 4: "CheckBox", 5: "Group",
    "Button": "Button", "ButtonAndMenu": "ButtonAndMenu",
    "Menu": "Menu", "Separator": "Separator",
    "CheckBox": "CheckBox", "Group": "Group",
}


def _strip_tcode_prefix(tcode: str) -> str:
    """Strip SAP command prefixes (/n, /N, /o, /O, /*) from a transaction code."""
    for prefix in ("/n", "/N", "/o", "/O", "/*"):
        tcode = tcode.removeprefix(prefix)
    return tcode

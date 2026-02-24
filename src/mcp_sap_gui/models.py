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
    SHIFT_F3 = 15  # Back (same as F3)
    SHIFT_F4 = 16
    SHIFT_F5 = 17
    SHIFT_F6 = 18
    SHIFT_F7 = 19
    SHIFT_F8 = 20
    SHIFT_F9 = 21
    CTRL_S = 11     # Save (same as F11)
    CTRL_F = 32     # Find
    CTRL_G = 33     # Continue search
    CTRL_P = 34     # Print
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


@dataclass
class ScreenElement:
    """Information about a screen element."""
    id: str
    type: str
    name: str
    text: str
    changeable: bool
    visible: bool


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

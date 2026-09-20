"""MCP Server for SAP GUI Scripting interaction."""

__version__ = "0.4.0"

from .sap_controller import (  # noqa: F401
    SAPGUIController,
    SAPGUIError,
    SAPGUINotAvailableError,
    SAPGUINotConnectedError,
    ScreenElement,
    SessionInfo,
    VKey,
)

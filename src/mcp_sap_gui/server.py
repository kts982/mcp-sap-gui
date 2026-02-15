"""
MCP Server for SAP GUI Scripting.

This module implements a Model Context Protocol (MCP) server that exposes
SAP GUI automation capabilities to AI assistants like Claude.

Security Note:
    This server provides powerful automation capabilities. Production deployments
    should implement:
    - Transaction whitelisting
    - Read-only modes
    - Audit logging
    - Rate limiting
    - User confirmation for write operations
"""

import asyncio
import json
import logging
from concurrent.futures import ThreadPoolExecutor
from typing import Optional, Dict, Any, List
from dataclasses import dataclass, field

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import (
    Tool,
    TextContent,
    ImageContent,
)

from .sap_controller import (
    SAPGUIController,
    SAPGUIError,
    SAPGUINotAvailableError,
    SAPGUINotConnectedError,
    VKey,
    SessionInfo,
)

logger = logging.getLogger(__name__)


@dataclass
class ServerConfig:
    """Configuration for the MCP SAP GUI server."""

    # Security settings
    read_only: bool = False
    allowed_transactions: Optional[List[str]] = None  # None = all allowed
    blocked_transactions: List[str] = field(default_factory=lambda: [
        "SU01", "SU10", "SU01D",  # User administration
        "PFCG", "SU53",           # Role administration
        "SM21", "ST22",           # System logs
        "SE16N",                  # Table maintenance
    ])
    require_confirmation_for_writes: bool = True

    # Behavior settings
    auto_connect_existing: bool = True
    default_language: str = "EN"
    max_table_rows: int = 500


class MCPSAPGUIServer:
    """
    MCP Server implementation for SAP GUI automation.

    This server exposes SAP GUI Scripting capabilities through the MCP protocol,
    allowing AI assistants to interact with SAP systems.
    """

    def __init__(self, config: Optional[ServerConfig] = None):
        """Initialize the MCP server."""
        self.config = config or ServerConfig()
        self.controller = SAPGUIController()
        self.server = Server("mcp-sap-gui")
        # Single-thread executor so all COM calls stay on the same thread
        self._com_executor = ThreadPoolExecutor(max_workers=1)
        self._setup_handlers()

    def _setup_handlers(self):
        """Set up MCP request handlers."""

        @self.server.list_tools()
        async def list_tools() -> List[Tool]:
            """Return list of available tools."""
            tools = [
                # Connection tools
                Tool(
                    name="sap_connect",
                    description="Connect to an SAP system by its name in SAP Logon Pad. "
                                "Optionally provide credentials for automatic login.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "system_description": {
                                "type": "string",
                                "description": "Exact system name as shown in SAP Logon "
                                               "(e.g., 'DEV - Development System')"
                            },
                            "client": {
                                "type": "string",
                                "description": "SAP client number (optional)"
                            },
                            "user": {
                                "type": "string",
                                "description": "SAP username (optional)"
                            },
                            "password": {
                                "type": "string",
                                "description": "SAP password (optional)"
                            },
                            "language": {
                                "type": "string",
                                "description": "Login language, e.g., 'EN' (optional)"
                            }
                        },
                        "required": ["system_description"]
                    }
                ),
                Tool(
                    name="sap_connect_existing",
                    description="Connect to an already open SAP session. "
                                "Use this when SAP is already logged in.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "connection_index": {
                                "type": "integer",
                                "description": "Index of the connection (0 = first)",
                                "default": 0
                            },
                            "session_index": {
                                "type": "integer",
                                "description": "Index of the session (0 = first)",
                                "default": 0
                            }
                        }
                    }
                ),
                Tool(
                    name="sap_list_connections",
                    description="List all open SAP connections and sessions",
                    inputSchema={"type": "object", "properties": {}}
                ),
                Tool(
                    name="sap_get_session_info",
                    description="Get information about the current SAP session "
                                "(system, client, user, transaction, screen)",
                    inputSchema={"type": "object", "properties": {}}
                ),

                # Navigation tools
                Tool(
                    name="sap_execute_transaction",
                    description="Execute an SAP transaction code (e.g., MM03, VA01, SE80)",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "tcode": {
                                "type": "string",
                                "description": "Transaction code to execute"
                            }
                        },
                        "required": ["tcode"]
                    }
                ),
                Tool(
                    name="sap_send_key",
                    description="Send a keyboard key. Common keys: Enter, F1 (Help), "
                                "F3 (Back), F4 (Search help), F5 (Refresh), F8 (Execute), "
                                "F11 (Save), F12 (Cancel)",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "key": {
                                "type": "string",
                                "description": "Key to send: 'Enter', 'F1'-'F12', "
                                               "'Back', 'Save', 'Cancel', 'Execute'",
                                "enum": ["Enter", "F1", "F2", "F3", "F4", "F5", "F6",
                                         "F7", "F8", "F9", "F10", "F11", "F12",
                                         "Back", "Save", "Cancel", "Execute", "Refresh"]
                            }
                        },
                        "required": ["key"]
                    }
                ),
                Tool(
                    name="sap_get_screen_info",
                    description="Get information about the current SAP screen "
                                "(transaction, program, screen number, title, status message)",
                    inputSchema={"type": "object", "properties": {}}
                ),

                # Field tools
                Tool(
                    name="sap_read_field",
                    description="Read the value of a field on the current SAP screen",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "field_id": {
                                "type": "string",
                                "description": "SAP GUI field ID (e.g., 'wnd[0]/usr/txtMATNR')"
                            }
                        },
                        "required": ["field_id"]
                    }
                ),
                Tool(
                    name="sap_set_field",
                    description="Set a value in a field on the current SAP screen",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "field_id": {
                                "type": "string",
                                "description": "SAP GUI field ID"
                            },
                            "value": {
                                "type": "string",
                                "description": "Value to set"
                            }
                        },
                        "required": ["field_id", "value"]
                    }
                ),
                Tool(
                    name="sap_press_button",
                    description="Press a button on the current SAP screen",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "button_id": {
                                "type": "string",
                                "description": "SAP GUI button ID "
                                               "(e.g., 'wnd[0]/tbar[1]/btn[8]' for Execute)"
                            }
                        },
                        "required": ["button_id"]
                    }
                ),

                Tool(
                    name="sap_select_checkbox",
                    description="Select or deselect a checkbox on the current SAP screen",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "checkbox_id": {
                                "type": "string",
                                "description": "SAP GUI checkbox ID (e.g., 'wnd[0]/usr/chkFLAG')"
                            },
                            "selected": {
                                "type": "boolean",
                                "description": "True to check, False to uncheck",
                                "default": True
                            }
                        },
                        "required": ["checkbox_id"]
                    }
                ),

                # Table tools
                Tool(
                    name="sap_read_table",
                    description="Read data from an ALV grid or table on the current screen",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "table_id": {
                                "type": "string",
                                "description": "SAP GUI table/grid ID"
                            },
                            "max_rows": {
                                "type": "integer",
                                "description": "Maximum rows to read (default 100)",
                                "default": 100
                            }
                        },
                        "required": ["table_id"]
                    }
                ),
                Tool(
                    name="sap_get_alv_toolbar",
                    description="Get all toolbar buttons from an ALV grid. Returns button IDs, "
                                "texts, and types. Use this to discover available actions.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "grid_id": {
                                "type": "string",
                                "description": "SAP GUI grid ID"
                            }
                        },
                        "required": ["grid_id"]
                    }
                ),
                Tool(
                    name="sap_press_alv_toolbar_button",
                    description="Press a toolbar button on an ALV grid (e.g., sort, filter, "
                                "export, custom actions). Use sap_get_alv_toolbar to find button IDs.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "grid_id": {
                                "type": "string",
                                "description": "SAP GUI grid ID"
                            },
                            "button_id": {
                                "type": "string",
                                "description": "Toolbar button ID (from sap_get_alv_toolbar)"
                            }
                        },
                        "required": ["grid_id", "button_id"]
                    }
                ),
                Tool(
                    name="sap_select_alv_context_menu_item",
                    description="Select an item from an opened ALV context menu. "
                                "First use sap_press_alv_toolbar_button on a Menu button "
                                "to open the menu, then use this to select an item. "
                                "Alternatively, pass toolbar_button_id to open the menu "
                                "and select the item in one atomic call (recommended).",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "grid_id": {
                                "type": "string",
                                "description": "SAP GUI grid ID"
                            },
                            "menu_item_id": {
                                "type": "string",
                                "description": "Function code / ID of the menu item"
                            },
                            "toolbar_button_id": {
                                "type": "string",
                                "description": "Optional: toolbar button ID to open the menu first "
                                               "(e.g., 'METHODS'). Combines open+select in one call."
                            }
                        },
                        "required": ["grid_id", "menu_item_id"]
                    }
                ),
                Tool(
                    name="sap_select_table_row",
                    description="Select a row in a table/grid",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "table_id": {"type": "string"},
                            "row": {"type": "integer", "description": "Row index (0-based)"}
                        },
                        "required": ["table_id", "row"]
                    }
                ),
                Tool(
                    name="sap_double_click_cell",
                    description="Double-click a cell in a table/grid (often opens details)",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "table_id": {"type": "string"},
                            "row": {"type": "integer"},
                            "column": {"type": "string", "description": "Column name"}
                        },
                        "required": ["table_id", "row", "column"]
                    }
                ),

                # Tree tools
                Tool(
                    name="sap_read_tree",
                    description="Read data from a tree control (SAP.TableTreeControl, "
                                "SAP.ColumnTreeControl, etc.). Returns node hierarchy "
                                "with texts and column values.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "tree_id": {
                                "type": "string",
                                "description": "SAP GUI tree ID "
                                               "(e.g., 'wnd[0]/usr/shell/shellcont[0]/shell')"
                            },
                            "max_nodes": {
                                "type": "integer",
                                "description": "Maximum nodes to read (default 200)",
                                "default": 200
                            }
                        },
                        "required": ["tree_id"]
                    }
                ),
                Tool(
                    name="sap_expand_tree_node",
                    description="Expand a folder node in a tree control to reveal its children",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "tree_id": {"type": "string", "description": "SAP GUI tree ID"},
                            "node_key": {"type": "string", "description": "Key of the node to expand"}
                        },
                        "required": ["tree_id", "node_key"]
                    }
                ),
                Tool(
                    name="sap_collapse_tree_node",
                    description="Collapse a folder node in a tree control",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "tree_id": {"type": "string", "description": "SAP GUI tree ID"},
                            "node_key": {"type": "string", "description": "Key of the node to collapse"}
                        },
                        "required": ["tree_id", "node_key"]
                    }
                ),
                Tool(
                    name="sap_select_tree_node",
                    description="Select a node in a tree control",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "tree_id": {"type": "string", "description": "SAP GUI tree ID"},
                            "node_key": {"type": "string", "description": "Key of the node to select"}
                        },
                        "required": ["tree_id", "node_key"]
                    }
                ),
                Tool(
                    name="sap_double_click_tree_node",
                    description="Double-click a node in a tree control (often opens details or drills down)",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "tree_id": {"type": "string", "description": "SAP GUI tree ID"},
                            "node_key": {"type": "string", "description": "Key of the node to double-click"}
                        },
                        "required": ["tree_id", "node_key"]
                    }
                ),

                # Discovery tools
                Tool(
                    name="sap_get_screen_elements",
                    description="Discover all elements on the current SAP screen. "
                                "Useful for finding field IDs when working with a new screen.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "container_id": {
                                "type": "string",
                                "description": "Container to enumerate (default: main user area)",
                                "default": "wnd[0]/usr"
                            }
                        }
                    }
                ),
                Tool(
                    name="sap_screenshot",
                    description="Take a screenshot of the current SAP window",
                    inputSchema={"type": "object", "properties": {}}
                ),
            ]

            # Filter out write/mutating tools if read-only mode
            if self.config.read_only:
                write_tools = {
                    "sap_set_field", "sap_press_button", "sap_select_checkbox",
                    "sap_select_table_row", "sap_double_click_cell",
                    "sap_execute_transaction", "sap_send_key",
                    "sap_press_alv_toolbar_button", "sap_select_alv_context_menu_item",
                    "sap_expand_tree_node", "sap_collapse_tree_node",
                    "sap_select_tree_node", "sap_double_click_tree_node",
                }
                tools = [t for t in tools if t.name not in write_tools]

            return tools

        @self.server.call_tool()
        async def call_tool(name: str, arguments: Dict[str, Any]) -> List[TextContent | ImageContent]:
            """Handle tool calls."""
            try:
                result = await self._handle_tool(name, arguments)

                # Handle screenshot specially (return image)
                if name == "sap_screenshot" and "data" in result:
                    return [
                        ImageContent(
                            type="image",
                            data=result["data"],
                            mimeType="image/png"
                        )
                    ]

                # Return JSON for all other tools
                return [TextContent(type="text", text=json.dumps(result, indent=2))]

            except Exception as e:
                logger.error("Tool %s failed: %s", name, e)
                return [TextContent(type="text", text=json.dumps({"error": str(e)}))]

    async def _handle_tool(self, name: str, arguments: Dict[str, Any]) -> Dict[str, Any]:
        """Route tool calls to appropriate handlers."""

        # Run synchronous SAP operations in thread pool
        loop = asyncio.get_running_loop()

        # Connection tools
        if name == "sap_connect":
            # Extract args explicitly to avoid passing unexpected keys
            # and to prevent password from leaking into logs/tracebacks
            connect_args = {
                "system_description": arguments["system_description"],
            }
            for key in ("client", "user", "password", "language"):
                if key in arguments:
                    connect_args[key] = arguments[key]
            info = await loop.run_in_executor(
                self._com_executor, lambda: self.controller.connect(**connect_args)
            )
            return info.__dict__ if hasattr(info, '__dict__') else info

        elif name == "sap_connect_existing":
            info = await loop.run_in_executor(
                self._com_executor,
                lambda: self.controller.connect_to_existing_session(
                    arguments.get("connection_index", 0),
                    arguments.get("session_index", 0)
                )
            )
            return info.__dict__ if hasattr(info, '__dict__') else info

        elif name == "sap_list_connections":
            return await loop.run_in_executor(
                self._com_executor, self.controller.list_connections
            )

        elif name == "sap_get_session_info":
            info = await loop.run_in_executor(
                self._com_executor, self.controller.get_session_info
            )
            return info.__dict__ if hasattr(info, '__dict__') else info

        # Navigation tools
        elif name == "sap_execute_transaction":
            if self.config.read_only:
                return {"error": "Write operations disabled in read-only mode"}
            tcode = arguments["tcode"]
            if self._is_transaction_blocked(tcode):
                return {"error": f"Transaction {tcode} is blocked by security policy"}
            return await loop.run_in_executor(
                self._com_executor, lambda: self.controller.execute_transaction(tcode)
            )

        elif name == "sap_send_key":
            if self.config.read_only:
                return {"error": "Write operations disabled in read-only mode"}
            vkey = self._parse_key(arguments["key"])
            return await loop.run_in_executor(
                self._com_executor, lambda: self.controller.send_vkey(vkey)
            )

        elif name == "sap_get_screen_info":
            return await loop.run_in_executor(
                self._com_executor, self.controller.get_screen_info
            )

        # Field tools
        elif name == "sap_read_field":
            return await loop.run_in_executor(
                self._com_executor, lambda: self.controller.read_field(arguments["field_id"])
            )

        elif name == "sap_set_field":
            if self.config.read_only:
                return {"error": "Write operations disabled in read-only mode"}
            return await loop.run_in_executor(
                self._com_executor, lambda: self.controller.set_field(
                    arguments["field_id"], arguments["value"]
                )
            )

        elif name == "sap_press_button":
            if self.config.read_only:
                return {"error": "Write operations disabled in read-only mode"}
            return await loop.run_in_executor(
                self._com_executor, lambda: self.controller.press_button(arguments["button_id"])
            )

        elif name == "sap_select_checkbox":
            if self.config.read_only:
                return {"error": "Write operations disabled in read-only mode"}
            return await loop.run_in_executor(
                self._com_executor, lambda: self.controller.select_checkbox(
                    arguments["checkbox_id"], arguments.get("selected", True)
                )
            )

        # Table tools
        elif name == "sap_read_table":
            max_rows = min(
                arguments.get("max_rows", 100),
                self.config.max_table_rows
            )
            return await loop.run_in_executor(
                self._com_executor, lambda: self.controller.read_table(
                    arguments["table_id"], max_rows
                )
            )

        elif name == "sap_get_alv_toolbar":
            return await loop.run_in_executor(
                self._com_executor, lambda: self.controller.get_alv_toolbar(
                    arguments["grid_id"]
                )
            )

        elif name == "sap_press_alv_toolbar_button":
            if self.config.read_only:
                return {"error": "Write operations disabled in read-only mode"}
            return await loop.run_in_executor(
                self._com_executor, lambda: self.controller.press_alv_toolbar_button(
                    arguments["grid_id"], arguments["button_id"]
                )
            )

        elif name == "sap_select_alv_context_menu_item":
            if self.config.read_only:
                return {"error": "Write operations disabled in read-only mode"}
            return await loop.run_in_executor(
                self._com_executor, lambda: self.controller.select_alv_context_menu_item(
                    arguments["grid_id"], arguments["menu_item_id"],
                    arguments.get("toolbar_button_id")
                )
            )

        elif name == "sap_select_table_row":
            if self.config.read_only:
                return {"error": "Write operations disabled in read-only mode"}
            return await loop.run_in_executor(
                self._com_executor, lambda: self.controller.select_table_row(
                    arguments["table_id"], arguments["row"]
                )
            )

        elif name == "sap_double_click_cell":
            if self.config.read_only:
                return {"error": "Write operations disabled in read-only mode"}
            return await loop.run_in_executor(
                self._com_executor, lambda: self.controller.double_click_table_cell(
                    arguments["table_id"], arguments["row"], arguments["column"]
                )
            )

        # Tree tools
        elif name == "sap_read_tree":
            max_nodes = min(
                arguments.get("max_nodes", 200),
                self.config.max_table_rows
            )
            return await loop.run_in_executor(
                self._com_executor, lambda: self.controller.read_tree(
                    arguments["tree_id"], max_nodes
                )
            )

        elif name == "sap_expand_tree_node":
            if self.config.read_only:
                return {"error": "Write operations disabled in read-only mode"}
            return await loop.run_in_executor(
                self._com_executor, lambda: self.controller.expand_tree_node(
                    arguments["tree_id"], arguments["node_key"]
                )
            )

        elif name == "sap_collapse_tree_node":
            if self.config.read_only:
                return {"error": "Write operations disabled in read-only mode"}
            return await loop.run_in_executor(
                self._com_executor, lambda: self.controller.collapse_tree_node(
                    arguments["tree_id"], arguments["node_key"]
                )
            )

        elif name == "sap_select_tree_node":
            if self.config.read_only:
                return {"error": "Write operations disabled in read-only mode"}
            return await loop.run_in_executor(
                self._com_executor, lambda: self.controller.select_tree_node(
                    arguments["tree_id"], arguments["node_key"]
                )
            )

        elif name == "sap_double_click_tree_node":
            if self.config.read_only:
                return {"error": "Write operations disabled in read-only mode"}
            return await loop.run_in_executor(
                self._com_executor, lambda: self.controller.double_click_tree_node(
                    arguments["tree_id"], arguments["node_key"]
                )
            )

        # Discovery tools
        elif name == "sap_get_screen_elements":
            elements = await loop.run_in_executor(
                self._com_executor, lambda: self.controller.get_screen_elements(
                    arguments.get("container_id", "wnd[0]/usr")
                )
            )
            return {
                "element_count": len(elements),
                "elements": [e.__dict__ for e in elements]
            }

        elif name == "sap_screenshot":
            return await loop.run_in_executor(
                self._com_executor, self.controller.take_screenshot
            )

        else:
            return {"error": f"Unknown tool: {name}"}

    def _is_transaction_blocked(self, tcode: str) -> bool:
        """Check if a transaction is blocked."""
        tcode_upper = tcode.upper().removeprefix("/N").removeprefix("/O")

        # Check blocklist
        if tcode_upper in self.config.blocked_transactions:
            return True

        # Check allowlist if configured
        if self.config.allowed_transactions is not None:
            return tcode_upper not in self.config.allowed_transactions

        return False

    def _parse_key(self, key: str) -> int:
        """Parse key name to VKey code."""
        key_map = {
            "Enter": VKey.ENTER,
            "F1": VKey.F1,
            "F2": VKey.F2,
            "F3": VKey.F3,
            "Back": VKey.F3,
            "F4": VKey.F4,
            "F5": VKey.F5,
            "Refresh": VKey.F5,
            "F6": VKey.F6,
            "F7": VKey.F7,
            "F8": VKey.F8,
            "Execute": VKey.F8,
            "F9": VKey.F9,
            "F10": VKey.F10,
            "F11": VKey.F11,
            "Save": VKey.F11,
            "F12": VKey.F12,
            "Cancel": VKey.F12,
        }
        if key not in key_map:
            raise ValueError(f"Unknown key: '{key}'. Valid keys: {', '.join(key_map.keys())}")
        return key_map[key]

    async def run(self):
        """Run the MCP server."""
        try:
            async with stdio_server() as (read_stream, write_stream):
                await self.server.run(
                    read_stream, write_stream, self.server.create_initialization_options()
                )
        finally:
            self._com_executor.shutdown(wait=False)


def main():
    """Main entry point."""
    import argparse

    parser = argparse.ArgumentParser(description="MCP Server for SAP GUI")
    parser.add_argument("--read-only", action="store_true",
                        help="Run in read-only mode (no write operations)")
    parser.add_argument("--allowed-transactions", nargs="*",
                        help="Whitelist of allowed transaction codes")
    parser.add_argument("--debug", action="store_true",
                        help="Enable debug logging")
    args = parser.parse_args()

    if args.debug:
        logging.basicConfig(level=logging.DEBUG)
    else:
        logging.basicConfig(level=logging.INFO)

    config = ServerConfig(
        read_only=args.read_only,
        allowed_transactions=args.allowed_transactions,
    )

    server = MCPSAPGUIServer(config)
    asyncio.run(server.run())


if __name__ == "__main__":
    main()

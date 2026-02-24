#!/usr/bin/env python3
"""
Basic usage example for SAP GUI Controller.

This example shows how to use the SAP GUI Controller directly
(without the MCP server) for simple automation tasks.

Requirements:
- SAP Logon Pad must be running
- SAP GUI Scripting must be enabled
"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from mcp_sap_gui.sap_controller import SAPGUIController, VKey


def example_display_material():
    """Example: Display material master data."""

    controller = SAPGUIController()

    try:
        # Connect to an existing session (SAP must be logged in)
        print("Connecting to existing SAP session...")
        session_info = controller.connect_to_existing_session()
        print(f"Connected to {session_info.system_name} as {session_info.user}")

        # Execute MM03 (Display Material)
        print("\nExecuting transaction MM03...")
        controller.execute_transaction("MM03")

        # Enter material number
        material = "TEST-001"
        print(f"Entering material: {material}")
        controller.set_field("wnd[0]/usr/ctxtRMMG1-MATNR", material)

        # Press Enter
        controller.send_vkey(VKey.ENTER)

        # Select all views and continue
        controller.send_vkey(VKey.ENTER)

        # Read some fields
        screen = controller.get_screen_info()
        print(f"\nCurrent screen: {screen}")

        # Try to read material description
        desc = controller.read_field("wnd[0]/usr/txtMAKT-MAKTX")
        print(f"Material description: {desc.get('value', 'N/A')}")

        # Go back
        controller.press_back()

    except Exception as e:
        print(f"Error: {e}")


def example_discover_screen():
    """Example: Discover elements on current screen."""

    controller = SAPGUIController()

    try:
        # Connect to existing session
        session_info = controller.connect_to_existing_session()
        print(f"Connected to {session_info.system_name}")
        print(f"Current transaction: {session_info.transaction}")

        # Get all screen elements
        print("\nDiscovering screen elements...")
        elements = controller.get_screen_elements()

        print(f"\nFound {len(elements)} elements:")
        for elem in elements:
            if elem.changeable:  # Only show input fields
                print(f"  {elem.id}")
                print(f"    Type: {elem.type}")
                print(f"    Text: {elem.text[:50] if elem.text else ''}")
                print()

    except Exception as e:
        print(f"Error: {e}")


def example_list_connections():
    """Example: List all open SAP connections."""

    controller = SAPGUIController()

    try:
        connections = controller.list_connections()

        if not connections:
            print("No SAP connections found. Is SAP Logon Pad running?")
            return

        print(f"Found {len(connections)} connection(s):\n")
        for conn in connections:
            print(f"Connection {conn['index']}: {conn['description']}")
            for sess in conn['sessions']:
                print(f"  Session {sess['index']}: {sess['user']} - {sess['transaction']}")

    except Exception as e:
        print(f"Error: {e}")


if __name__ == "__main__":
    print("SAP GUI Controller Examples")
    print("=" * 50)

    print("\n1. Listing connections...")
    example_list_connections()

    # Uncomment to run other examples:
    # print("\n2. Discovering screen elements...")
    # example_discover_screen()

    # print("\n3. Display material...")
    # example_display_material()

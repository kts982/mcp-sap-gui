#!/usr/bin/env python3
"""Simple CLI to test SAP GUI Controller directly."""

import sys
sys.path.insert(0, 'src')

from mcp_sap_gui.sap_controller import SAPGUIController, VKey


def main():
    controller = SAPGUIController()

    while True:
        print("\n=== SAP GUI Controller Test CLI ===")
        print("1. List connections")
        print("2. Connect to existing session")
        print("3. Get session info")
        print("4. Execute transaction")
        print("5. Get screen elements")
        print("6. Read field")
        print("7. Set field")
        print("8. Send key (Enter/F3/F8)")
        print("9. Get screen info")
        print("0. Exit")

        choice = input("\nChoice: ").strip()

        try:
            if choice == "1":
                conns = controller.list_connections()
                if not conns:
                    print("No connections found. Is SAP Logon running?")
                else:
                    for c in conns:
                        print(f"Connection {c['index']}: {c['description']}")
                        for s in c['sessions']:
                            print(f"  Session {s['index']}: {s['user']} @ {s['transaction']}")

            elif choice == "2":
                conn_idx = int(input("Connection index [0]: ") or "0")
                sess_idx = int(input("Session index [0]: ") or "0")
                info = controller.connect_to_existing_session(conn_idx, sess_idx)
                print(f"Connected to {info.system_name} as {info.user}")

            elif choice == "3":
                info = controller.get_session_info()
                print(f"System: {info.system_name}")
                print(f"Client: {info.client}")
                print(f"User: {info.user}")
                print(f"Transaction: {info.transaction}")
                print(f"Screen: {info.screen_number}")

            elif choice == "4":
                tcode = input("Transaction code: ").strip()
                result = controller.execute_transaction(tcode)
                print(f"Result: {result}")

            elif choice == "5":
                container = input("Container ID [wnd[0]/usr]: ").strip() or "wnd[0]/usr"
                elements = controller.get_screen_elements(container)
                print(f"\nFound {len(elements)} elements:")
                for e in elements[:20]:  # Limit output
                    if e.changeable or e.text:
                        print(f"  {e.id}")
                        print(f"    Type: {e.type}, Text: {e.text[:40] if e.text else ''}")
                if len(elements) > 20:
                    print(f"  ... and {len(elements) - 20} more")

            elif choice == "6":
                field_id = input("Field ID: ").strip()
                result = controller.read_field(field_id)
                print(f"Value: {result.get('value', 'N/A')}")
                print(f"Type: {result.get('type', 'N/A')}")

            elif choice == "7":
                field_id = input("Field ID: ").strip()
                value = input("Value: ").strip()
                result = controller.set_field(field_id, value)
                print(f"Result: {result}")

            elif choice == "8":
                key = input("Key (enter/f3/f8): ").strip().lower()
                key_map = {"enter": VKey.ENTER, "f3": VKey.F3, "f8": VKey.F8}
                vkey = key_map.get(key, VKey.ENTER)
                result = controller.send_vkey(vkey)
                print(f"Result: {result}")

            elif choice == "9":
                info = controller.get_screen_info()
                print(f"Transaction: {info.get('transaction')}")
                print(f"Program: {info.get('program')}")
                print(f"Screen: {info.get('screen_number')}")
                print(f"Title: {info.get('title')}")
                print(f"Status: {info.get('status_message')}")

            elif choice == "0":
                print("Bye!")
                break

            else:
                print("Invalid choice")

        except Exception as e:
            print(f"Error: {e}")


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""Diagnose SAP GUI COM object structure."""

import win32com.client

def main():
    print("=== SAP GUI COM Diagnostic ===\n")

    # Step 1: Get SAPGUI object
    print("Step 1: Getting SAPGUI object...")
    try:
        sap_gui = win32com.client.GetObject("SAPGUI")
        print(f"  OK - Type: {type(sap_gui)}")
    except Exception as e:
        print(f"  FAILED: {e}")
        print("\n  Make sure SAP Logon Pad is running!")
        return

    # Step 2: Get Scripting Engine
    print("\nStep 2: Getting ScriptingEngine...")
    try:
        app = sap_gui.GetScriptingEngine
        print(f"  OK - Type: {type(app)}")
        print(f"  Children.Count: {app.Children.Count}")
    except Exception as e:
        print(f"  FAILED: {e}")
        return

    if app.Children.Count == 0:
        print("\n  No connections found. Please open/login to an SAP system first.")
        return

    # Step 3: Examine first connection
    print(f"\nStep 3: Examining {app.Children.Count} connection(s)...")
    for i in range(app.Children.Count):
        print(f"\n  Connection {i}:")
        conn = app.Children(i)

        # Try to list available attributes
        print(f"    Type: {type(conn)}")

        # Try common properties
        for prop in ['Id', 'Description', 'ConnectionString', 'Name', 'Type', 'Children']:
            try:
                val = getattr(conn, prop)
                if prop == 'Children':
                    print(f"    {prop}: <collection with {val.Count} items>")
                else:
                    print(f"    {prop}: {val}")
            except Exception as e:
                print(f"    {prop}: ERROR - {e}")

        # Step 4: Examine sessions in this connection
        print(f"\n    Sessions ({conn.Children.Count}):")
        for j in range(conn.Children.Count):
            print(f"\n      Session {j}:")
            try:
                sess = conn.Children(j)
                print(f"        Type: {type(sess)}")

                # Try session properties
                for prop in ['Id', 'Name', 'Type', 'Info']:
                    try:
                        val = getattr(sess, prop)
                        print(f"        {prop}: {val}")
                    except Exception as e:
                        print(f"        {prop}: ERROR - {e}")

                # Try Info properties
                print("\n        Info properties:")
                try:
                    info = sess.Info
                    for iprop in ['SystemName', 'Client', 'User', 'Transaction',
                                  'Program', 'ScreenNumber', 'Language', 'SystemNumber']:
                        try:
                            val = getattr(info, iprop)
                            print(f"          {iprop}: {val}")
                        except Exception as e:
                            print(f"          {iprop}: ERROR - {e}")
                except Exception as e:
                    print(f"        Info: ERROR - {e}")

            except Exception as e:
                print(f"        ERROR accessing session: {e}")

    print("\n=== Diagnostic Complete ===")


if __name__ == "__main__":
    main()

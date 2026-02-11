#!/usr/bin/env python3
"""Diagnose SAP GUI COM - matching zsapconnect approach exactly."""

import win32com.client

def main():
    print("=== SAP GUI COM Diagnostic v2 ===\n")

    # Method 1: Exactly like zsapconnect's find_sap_gui()
    print("Method 1: GetObject('SAPGUI').GetScriptingEngine")
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        print(f"  Application: {application}")
        print(f"  Type: {application.Type if hasattr(application, 'Type') else 'N/A'}")
        print(f"  Children.Count: {application.Children.Count}")

        if application.Children.Count > 0:
            conn = application.Children(0)
            print(f"\n  Connection 0:")
            print(f"    Type: {conn.Type}")
            print(f"    Children.Count: {conn.Children.Count}")

            if conn.Children.Count > 0:
                sess = conn.Children(0)
                print(f"\n    Session 0:")
                print(f"      Type: {sess.Type}")
                print(f"      Id: {sess.Id}")
                info = sess.Info
                print(f"      User: {info.User}")
                print(f"      Transaction: {info.Transaction}")
            else:
                print("    No sessions in connection!")
        else:
            print("  No connections!")
    except Exception as e:
        print(f"  Error: {e}")

    # Method 2: Try Dispatch instead of GetObject
    print("\n" + "="*50)
    print("Method 2: Using Dispatch to create new instance")
    try:
        sap = win32com.client.Dispatch("SapROTWr.SapROTWrapper")
        rot = sap.QueryInterface(win32com.client.pythoncom.IID_IRunningObjectTable)
        print(f"  ROT Wrapper created")
        # This is a different approach...
    except Exception as e:
        print(f"  Error (expected if not available): {e}")

    # Method 3: Try getting active sessions via different ROT entry
    print("\n" + "="*50)
    print("Method 3: Checking SAP GUI process")
    try:
        import subprocess
        result = subprocess.run(['tasklist', '/FI', 'IMAGENAME eq saplogon.exe'],
                              capture_output=True, text=True)
        if 'saplogon.exe' in result.stdout:
            print("  saplogon.exe is running")
        else:
            print("  saplogon.exe NOT found")

        result = subprocess.run(['tasklist', '/FI', 'IMAGENAME eq sapgui.exe'],
                              capture_output=True, text=True)
        if 'sapgui.exe' in result.stdout:
            print("  sapgui.exe is running (active session)")
        else:
            print("  sapgui.exe NOT found")
    except Exception as e:
        print(f"  Error: {e}")

    # Method 4: Check scripting settings hint
    print("\n" + "="*50)
    print("SAP GUI Scripting Checklist:")
    print("  1. SAP GUI Options > Accessibility & Scripting > Scripting")
    print("     - 'Enable scripting' must be checked")
    print("     - 'Notify when script attaches' can cause popups")
    print("  2. Server-side: Transaction RZ11")
    print("     - Parameter: sapgui/user_scripting = TRUE")
    print("  3. User authorization: S_SCR object")

    print("\n=== End Diagnostic ===")


if __name__ == "__main__":
    main()

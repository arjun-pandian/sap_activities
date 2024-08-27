import sys
import os
import pyautogui

current_directory = os.getcwd()
functions_directory = os.path.abspath(os.path.join(current_directory, 'Files', 'Common'))
sys.path.append(functions_directory)

from sap_functions import sap_login, create_session

def check_login(system, client, skip_nex=False):
    try:
        sap_login(system, client)
        session = create_session()
        pyautogui.hotkey('win', 'up')
        if not skip_nex:
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nex"
            session.findById("wnd[0]").sendVKey(0)
        return True
    except Exception as e:
        print(f"Unable to login to {system} ({client}): {e}")
        return False

def open_systems(specific_system=None):
    systems = [
        ("CP5", "700"), ("PK9", "215"), ("PKS", "100"),
        ("PKX", "100"), ("PL5", "012"), ("PM6", "600"),
        ("PMP", "810"), ("PR3", "900"), ("PR5", "700"),
        ("VA5", "700"), ("DE5", "200"), ("CV5", "700"),
        ("CD5", "700")
    ]

    failed_logins = []

    if specific_system:
        for system, client in systems:
            if system == specific_system:
                if check_login(system, client, skip_nex=True):
                    print(f"Successfully logged into {system} ({client})")
                else:
                    failed_logins.append((system, client))
                break
        else:
            print(f"System {specific_system} not found in the list.")
    else:
        for system, client in systems:
            if check_login(system, client):
                print(f"Successfully logged into {system} ({client})")
            else:
                failed_logins.append((system, client))

    if failed_logins:
        print("\nSummary of failed logins:")
        for system, client in failed_logins:
            print(f"Failed to login to {system} ({client})")
    else:
        print("\nAll systems logged in successfully.")

if __name__ == "__main__":
    user_input = input("Enter a system code to open (or press Enter to open all systems): ").strip()
    
    if user_input:
        open_systems(specific_system=user_input)
    else:
        open_systems()

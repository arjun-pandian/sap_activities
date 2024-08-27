import sys
import os
current_directory = os.getcwd()
functions_directory = os.path.abspath(os.path.join(current_directory, 'Files', 'Common'))
sys.path.append(functions_directory)

import pyautogui
from sap_functions import sap_login, create_session

def check_login(system, client):
    try:
        sap_login(system, client)
        session = create_session()
        pyautogui.hotkey('win', 'up')
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nex"
        session.findById("wnd[0]").sendVKey(0)
        return True
    except Exception as e:
        print(f"Unable to login to {system} ({client}): {e}")
        return False

def open_systems():
    systems = [
        ("CP5", "700"), ("PK9", "215"), ("PKS", "100"),
        ("PKX", "100"), ("PL5", "012"), ("PM6", "600"),
        ("PMP", "810"), ("PR3", "900"), ("PR5", "700")
    ]

    for system, client in systems:
        if check_login(system, client):
            print(f"Successfully logged into {system} ({client})")
        else:
            print(f"Failed to login to {system} ({client})")

if __name__ == "__main__":
    open_systems()

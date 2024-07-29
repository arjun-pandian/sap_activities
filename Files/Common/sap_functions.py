from pywinauto import application, timings
import pyautogui
import time
from docx import Document
import subprocess
import pandas as pd
import win32com.client
from datetime import datetime, date, timedelta
import re
import json
import pygetwindow as gw
import os

current_directory = os.getcwd()

today = None
yesterday = None

def update_dates():
    global today, yesterday
    today = (date.today()).strftime("%d.%m.%Y")
    yesterday = (date.today() - timedelta(days=1)).strftime("%d.%m.%Y")

update_dates()

def read_credentials(file_path):
    try:
        with open(file_path, 'r') as f:
            credentials = json.load(f)
        return credentials.get('systems', {})
    except Exception as e:
        print(f"Error reading credentials file: {e}")
        return {}

def sap_login(system_name, client):
    try:
        file_path = os.path.abspath(os.path.join(current_directory, 'Files','Common','system_details.json'))
        systems_info = read_credentials(file_path)

        system_info = systems_info.get(system_name)
        if not system_info:
            raise ValueError(f"System {system_name} not found in credentials file")

        username = system_info.get('username')
        password = system_info.get('password')
        notes = system_info.get('notes')

        if not username or not password:
            raise ValueError(f"Missing credentials for system {system_name}")

        sap_gui_path = "C:/Program Files (x86)/SAP/FrontEnd/SapGui/sapshcut.exe"
        command = f'"{sap_gui_path}" -system="{system_name}" -guiparm="{notes}" -client={client} -user={username} -pw={password} -language=en'
        subprocess.Popen(command)
        time.sleep(10)  

        sap_window = None

        for _ in range(10):
            windows = gw.getAllTitles()
            for title in windows:
                if system_name in title:
                    sap_window = gw.getWindowsWithTitle(title)[0]
                    break
            if sap_window:
                break
            time.sleep(1)

        if sap_window:
            sap_window.activate()
            print(f"Logged in to {system_name}")
        else:
            print(f"Failed to login to {system_name}")

    except Exception as e:
        print(f"Error during SAP login: {e}")


def create_session():
    try:
        SapGuiAuto = win32com.client.GetObject('SAPGUI')
        application = SapGuiAuto.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)
        return session
    except Exception as e:
        print("Error creating SAP session:", e)
        return None

def take_and_save_screenshot(system_name, image_name):
    time.sleep(1)
    screenshot = pyautogui.screenshot()

    folder_path = os.path.abspath(os.path.join(current_directory, 'Scripts_run','Manual_Monitoring','Screenshots'))
    system_folder_path = os.path.join(folder_path, system_name)

    if not os.path.exists(system_folder_path):
        os.makedirs(system_folder_path)

    screenshot_file = f"{system_folder_path}/{image_name}.png"
    screenshot.save(screenshot_file)
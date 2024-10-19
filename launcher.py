import os
import sys
import requests
import subprocess
import logging
import winshell
from win32com.client import Dispatch

# Determine if we're running as a script or frozen executable
if getattr(sys, 'frozen', False):
    # we are running in a bundle
    bundle_dir = sys._MEIPASS
else:
    # we are running in a normal Python environment
    bundle_dir = os.path.dirname(os.path.abspath(__file__))

# Set up logging in the ASX folder on the main drive
ASX_FILES_DIR = r"C:\ASX"
if not os.path.exists(ASX_FILES_DIR):
    os.makedirs(ASX_FILES_DIR)
logging.basicConfig(filename=os.path.join(ASX_FILES_DIR, 'asx_launcher.log'), level=logging.DEBUG)

REQUIRED_FILES = {
    'main.py': 'https://github.com/pristinefr/files/releases/download/gamerfunk/main.py',
    'index.html': 'https://github.com/pristinefr/files/releases/download/gamerfunk/index.html',
    'script.js': 'https://github.com/pristinefr/files/releases/download/gamerfunk/script.js',
    'styles.css': 'https://github.com/pristinefr/files/releases/download/gamerfunk/styles.css'
}

def download_file(url, filename):
    try:
        response = requests.get(url)
        response.raise_for_status()
        with open(filename, 'wb') as f:
            f.write(response.content)
        logging.info(f"Downloaded: {filename}")
    except requests.RequestException as e:
        logging.error(f"Failed to download {filename}: {str(e)}")
        raise

def check_and_download_files():
    for file, url in REQUIRED_FILES.items():
        file_path = os.path.join(ASX_FILES_DIR, file)
        if not os.path.exists(file_path):
            logging.info(f"{file} not found. Downloading...")
            download_file(url, file_path)
        else:
            logging.info(f"{file} already exists.")

def run_main_app():
    main_script = os.path.join(ASX_FILES_DIR, 'main.py')
    try:
        python_exe = sys.executable
        subprocess.Popen([python_exe, main_script], creationflags=subprocess.CREATE_NO_WINDOW)
    except subprocess.CalledProcessError as e:
        logging.error(f"Failed to run main application: {str(e)}")
        raise

def create_desktop_shortcut():
    desktop = winshell.desktop()
    path = os.path.join(desktop, "ASX Loader.lnk")
    target = os.path.join(ASX_FILES_DIR, "main.py")
    
    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(path)
    shortcut.Targetpath = sys.executable
    shortcut.Arguments = f'"{target}"'
    shortcut.WorkingDirectory = ASX_FILES_DIR
    shortcut.IconLocation = sys.executable
    shortcut.save()
    
    logging.info(f"Created desktop shortcut: {path}")

if __name__ == '__main__':
    try:
        check_and_download_files()
        run_main_app()
        create_desktop_shortcut()
    except Exception as e:
        logging.error(f"An error occurred: {str(e)}")
        print(f"An error occurred. Please check the log file in {ASX_FILES_DIR} for details.")
    sys.exit()




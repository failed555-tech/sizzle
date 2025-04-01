import os
import tkinter as tk
from tkinter import messagebox, simpledialog
import shutil
import sys
import subprocess
import time
import urllib.request
import winreg

VERSION = "1.0.0"  # Current version of the application


def install_dependencies():
    try:
        import winshell
        from win32com.client import Dispatch
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pywin32", "winshell"])
        os.execl(sys.executable, sys.executable, *sys.argv)  # Restart script after installing


def download_file(url, file_path):
    try:
        urllib.request.urlretrieve(url, file_path)
    except Exception as e:
        messagebox.showwarning("Warning", f"Failed to download {os.path.basename(file_path)}: {e}")


def get_latest_version():
    version_url = "https://github.com/failed555-tech/sizzle/blob/119770de5620cda44d06703c537e761884433d22/version.txt"  # Update with actual URL
    try:
        with urllib.request.urlopen(version_url) as response:
            return response.read().decode("utf-8").strip()
    except Exception as e:
        messagebox.showwarning("Warning", f"Failed to check latest version: {e}")
        return VERSION  # Assume current version if check fails


def check_for_updates():
    latest_version = get_latest_version()
    if latest_version != VERSION:
        messagebox.showinfo("Update Available", f"A new version ({latest_version}) is available.")


def install_to_documents():
    documents_path = os.path.join(os.environ['USERPROFILE'], 'Documents', 'SizzlesUtilities')
    script_path = os.path.abspath(sys.argv[0])
    new_script_path = os.path.join(documents_path, os.path.basename(script_path))
    
    if not os.path.exists(documents_path):
        os.makedirs(documents_path)
    
    files = {
        "Sizzles Utilities Icon": "https://raw.githubusercontent.com/failed555-tech/sizzle/57a85d72cd7eac096a1e89abc17835d089223a3b/sizzlesutilitiesOFFICIAL%20icon.ico",
        "Text-Shortcut": "https://github.com/failed555-tech/sizzle/blob/537fbeec323a5174e8366b08fef9af64490eeb5b/Text-Shortcut.exe",  
        "Quick Window Resize": "https://github.com/failed555-tech/sizzle/blob/4a05bfebea8028edfb97a1724537b7b8b243dadd/quickwindowresize.exe",    
        "Fast Emoji Input": "https://github.com/failed555-tech/sizzle/blob/2d53df9ba60df5b4f6bdb1af9ba1c18540cea5cc/fastemojiinput.exe",  
        "How To Use These Utilities": "https://github.com/failed555-tech/sizzle/blob/7096d20df9d0d9877e3a2eed5dd78ad4a2f0ec9d/how%20to%20use%20Sizzles%20Utilities.pdf"        
    }
    
    downloaded_files = {}
    for display_name, url in files.items():
        file_name = url.split("/")[-1]
        file_path = os.path.join(documents_path, file_name)
        if not os.path.exists(file_path):
            download_file(url, file_path)
        downloaded_files[display_name] = file_path
    
    if script_path != new_script_path:
        shutil.copy(script_path, new_script_path)
        time.sleep(1)
        subprocess.Popen([sys.executable, new_script_path], creationflags=subprocess.DETACHED_PROCESS)
        sys.exit()
    
    return downloaded_files


def create_shortcut():
    desktop = os.path.join(os.environ['USERPROFILE'], 'Desktop')
    shortcut_path = os.path.join(desktop, "Sizzles Utilities.lnk")
    documents_path = os.path.join(os.environ['USERPROFILE'], 'Documents', 'SizzlesUtilities')
    script_path = os.path.join(documents_path, os.path.basename(sys.argv[0]))
    icon_path = os.path.join(documents_path, "sizzlesutilitiesOFFICIAL%20icon.ico")  
    
    try:
        import winshell
        from win32com.client import Dispatch
        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortcut(shortcut_path)
        shortcut.TargetPath = sys.executable  
        shortcut.Arguments = f'"{script_path}"'
        shortcut.WorkingDirectory = documents_path
        shortcut.Description = "Sizzles Utilities Shortcut"
        if os.path.exists(icon_path):
            shortcut.IconLocation = icon_path
        shortcut.Save()
    except Exception as e:
        messagebox.showwarning("Warning", f"Failed to create shortcut: {e}")


def create_ui(downloaded_files):
    root = tk.Tk()
    root.title(f"Sizzles Utilities - v{VERSION}")
    root.geometry("300x300")
    
    version_label = tk.Label(root, text=f"Current Version: {VERSION}")
    version_label.pack(pady=5)
    
    def on_button_click(file_path):
        launch_file(file_path)
    
    for name, path in downloaded_files.items():
        btn = tk.Button(root, text=name, command=lambda p=path: on_button_click(p))
        btn.pack(pady=5)
    
    exit_btn = tk.Button(root, text="Exit", command=root.quit)
    exit_btn.pack(pady=5)
    
    root.mainloop()


def launch_file(file_path):
    try:
        if os.path.exists(file_path):
            os.startfile(file_path)  
        else:
            messagebox.showerror("Error", f"File not found: {file_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Error launching file: {e}")


if __name__ == "__main__":
    install_dependencies()
    check_for_updates()
    downloaded_files = install_to_documents()
    create_shortcut()
    create_ui(downloaded_files)

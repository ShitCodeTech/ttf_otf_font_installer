import os
import shutil
import winreg
import ctypes
import tkinter as tk
from tkinter import filedialog

def is_admin():
    """
    Checks if the script is running with administrator privileges.
    """
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def install_font(font_path):
    """
    Installs a TTF or OTF font on Windows system-wide.

    :param font_path: Path to the font file (.ttf or .otf)
    """
    if not os.path.exists(font_path):
        print(f"Font file {font_path} does not exist.")
        return

    font_name = os.path.basename(font_path)
    font_extension = os.path.splitext(font_name)[1].lower()

    if font_extension not in ['.ttf', '.otf']:
        print(f"Unsupported font format: {font_extension}. Only TTF and OTF are supported.")
        return

    font_dest_dir = os.path.join(os.environ['WINDIR'], 'Fonts')
    font_dest_path = os.path.join(font_dest_dir, font_name)

    try:
        shutil.copy(font_path, font_dest_path)
        print(f"Copied {font_name} to {font_dest_path}")

        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts", 0, winreg.KEY_SET_VALUE) as reg_key:
            winreg.SetValueEx(reg_key, font_name, 0, winreg.REG_SZ, font_name)
        print(f"Added {font_name} to the Windows registry")

        os.system('RUNDLL32.EXE USER32.DLL,UpdatePerUserSystemParameters')
        print(f"Font {font_name} installed successfully and Windows notified.")

    except PermissionError:
        print(f"Permission denied: Unable to copy {font_name}. Please run the script as an administrator.")
    except Exception as e:
        print(f"Failed to install {font_name}: {e}")

def install_fonts_from_directory(directory):
    """
    Installs all TTF and OTF fonts found in the given directory.

    :param directory: Path to the directory containing font files.
    """
    for root, _, files in os.walk(directory):
        for file in files:
            if file.lower().endswith(('.ttf', '.otf')):
                font_path = os.path.join(root, file)
                install_font(font_path)

def select_directory():
    """
    Opens a dialog to select a directory and installs all fonts in it.
    """
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    directory = filedialog.askdirectory(title="Select the directory containing fonts")
    
    if directory:
        install_fonts_from_directory(directory)
    else:
        print("No directory selected.")

if __name__ == "__main__":
    if is_admin():
        select_directory()
    else:
        print("This script requires administrator privileges. Please restart the script as an administrator.")

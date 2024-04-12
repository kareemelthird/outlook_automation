import os
import sys
import winshell

def find_pythonw():
    # Check a common path for Python installation; adjust as necessary for typical installs
    common_paths = [
        r"C:\Python39\pythonw.exe",  # Adjust the version number as necessary
        r"C:\Python310\pythonw.exe",
        "C:\Python310\pythonw.exe",
        "C:\Python311\pythonw.exe",
        "C:\Python312\pythonw.exe",
        "C:\Python313\pythonw.exe",
        r"C:\Program Files\Python39\pythonw.exe",
        r"C:\Program Files\Python310\pythonw.exe"
    ]
    for path in common_paths:
        if os.path.exists(path):
            return path
    # Fallback to the current executable if none of the common paths fit
    return sys.executable.replace("python.exe", "pythonw.exe")

def create_shortcut_to_startup():
    # Use the helper function to find pythonw.exe
    pythonw_path = find_pythonw()
    
    # Determine the absolute path to the script you want to run at startup
    script_directory = os.path.dirname(os.path.abspath(__file__))
    script_path = os.path.join(script_directory, "background.py")
    
    # Path to the Startup folder for the current user
    startup_folder = winshell.startup()
    
    # Full path for the new shortcut in the Startup folder
    shortcut_path = os.path.join(startup_folder, "BackgroundApp.lnk")
    
    # Create a shortcut to run the script using pythonw.exe
    with winshell.shortcut(shortcut_path) as shortcut:
        shortcut.path = pythonw_path
        shortcut.arguments = f'"{script_path}"'
        shortcut.description = "Run my Python script silently on startup"
        shortcut.working_directory = script_directory
        shortcut.icon_location = (script_path, 0)

    print(f"Shortcut created: {shortcut_path}")

if __name__ == "__main__":
    create_shortcut_to_startup()

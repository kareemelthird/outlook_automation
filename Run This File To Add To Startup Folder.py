import os
import winshell
import sys

def create_shortcut_to_startup():
    # Find the pythonw.exe path in the same directory as the current Python interpreter
    pythonw_path = sys.executable.replace("python.exe", "pythonw.exe")
    
    # Absolute path to the script you want to run at startup
    script_path = r"D:\py\pr\Outook_Automation_Tool\background.pyw"
    
    # Path to the Startup folder
    startup_folder = winshell.startup()
    
    # Full path for the new shortcut
    shortcut_path = os.path.join(startup_folder, "BackgroundApp.lnk")
    
    # Create a shortcut to run the script using pythonw.exe
    with winshell.shortcut(shortcut_path) as shortcut:
        shortcut.path = pythonw_path
        shortcut.arguments = script_path
        shortcut.description = "Run my Python script silently on startup"
        shortcut.working_directory = os.path.dirname(script_path)
        shortcut.icon_location = (script_path, 0)

    print(f"Shortcut created: {shortcut_path}")

if __name__ == "__main__":
    create_shortcut_to_startup()

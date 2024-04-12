import os
import sys
import winshell

def create_shortcut_to_startup():
    # Automatically find the pythonw.exe path in the same directory as the current Python interpreter
    pythonw_path = sys.executable.replace("python.exe", "pythonw.exe") if "python.exe" in sys.executable else sys.executable
    
    # Determine the absolute path to the script you want to run at startup
    script_directory = os.path.dirname(os.path.abspath(__file__))  # Directory of this script
    script_path = os.path.join(script_directory, "background.pyw")  # Adjust filename if necessary
    
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
        shortcut.icon_location = (script_path, 0)  # Optionally, specify an icon

    print(f"Shortcut created: {shortcut_path}")

if __name__ == "__main__":
    create_shortcut_to_startup()

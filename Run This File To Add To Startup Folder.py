import os
import sys
import winshell

def create_shortcut_to_startup():
    # Determine the path to pythonw.exe based on the Python executable
    python_executable = sys.executable.replace("python.exe", "pythonw.exe")
    
    # Get the path to the script
    script_path = os.path.abspath("background.pyw")
    
    # Determine the Startup folder for the current user
    startup_folder = winshell.startup()
    
    # Define the path where the shortcut will be created
    shortcut_path = os.path.join(startup_folder, "BackgroundApp.lnk")
    
    # Create a shortcut in the startup folder
    with winshell.shortcut(shortcut_path) as shortcut:
        shortcut.path = python_executable
        shortcut.arguments = f'"{script_path}"'
        shortcut.description = "Run Background Automation Tool"
        shortcut.working_directory = os.path.dirname(script_path)
        shortcut.icon_location = (script_path, 0)  # Optionally specify an icon

    print(f"Shortcut created in Startup folder: {shortcut_path}")

if __name__ == "__main__":
    create_shortcut_to_startup()

import os
import sys
import winshell

def find_pythonw():
    # Check a common path for Python installation; adjust as necessary for typical installs
    common_paths = [
        r"C:\Python39\pythonw.exe",
        r"C:\Python310\pythonw.exe",
        r"C:\Python311\pythonw.exe",
        r"C:\Program Files\Python39\pythonw.exe",
        r"C:\Program Files\Python310\pythonw.exe",
        r"C:\Program Files\Python311\pythonw.exe",
    ]
    for path in common_paths:
        if os.path.exists(path):
            return path
    # Fallback to the executable replacing 'python.exe' with 'pythonw.exe'
    return sys.executable.replace("python.exe", "pythonw.exe")

def create_shortcut_and_run():
    pythonw_path = find_pythonw()
    
    script_directory = os.path.dirname(os.path.abspath(__file__))
    script_path = os.path.join(script_directory, "background.py")  # Ensure you have a .pyw version of your script
    
    startup_folder = winshell.startup()
    shortcut_path = os.path.join(startup_folder, "K3 Outlook Automation.lnk")
    
    with winshell.shortcut(shortcut_path) as shortcut:
        shortcut.path = pythonw_path
        shortcut.arguments = f'"{script_path}"'
        shortcut.description = "Run my Python script silently on startup"
        shortcut.working_directory = script_directory
        shortcut.icon_location = (script_path, 0)

    print(f"Shortcut created: {shortcut_path}")

    # Run the shortcut
    os.startfile(shortcut_path)
    print("Shortcut executed.")

if __name__ == "__main__":
    create_shortcut_and_run()

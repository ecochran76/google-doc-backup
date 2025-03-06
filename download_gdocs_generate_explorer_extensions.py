import os
import sys
import shutil
import subprocess

# Constants
SCRIPT_NAME = "download_google_doc.pyw"
REG_FILE_NAME = "add_context_menu.reg"
REMOVE_REG_FILE_NAME = "remove_context_menu.reg"

# Registry formatting constants
ESCAPED_QUOTE = '\\"'  # Represents \" in the .reg file

# Helper Functions
def escape_registry_path(path):
    """Escape a path for use in a Windows registry .reg file, excluding outer quotes."""
    return path.replace('"', ESCAPED_QUOTE).replace("\\", "\\\\")

def build_command_string(executable_path, script_path):
    """Build a properly quoted and escaped command string for the registry."""
    escaped_exec = escape_registry_path(executable_path)
    escaped_script = escape_registry_path(script_path)
    # Build the command with outer quotes and escaped inner quotes
    return f'"{ESCAPED_QUOTE}{escaped_exec}{ESCAPED_QUOTE} {ESCAPED_QUOTE}{escaped_script}{ESCAPED_QUOTE} {ESCAPED_QUOTE}%1{ESCAPED_QUOTE}"'

def create_registry_entry(file_extension, menu_text, command_string):
    """Create a registry entry for a given file extension."""
    return f"""[HKEY_CLASSES_ROOT\\SystemFileAssociations\\.{file_extension}\\shell\\DownloadAsOffice]
@="{menu_text}"

[HKEY_CLASSES_ROOT\\SystemFileAssociations\\.{file_extension}\\shell\\DownloadAsOffice\\command]
@={command_string}
"""

def create_removal_entry(file_extension):
    """Create a registry removal entry for a given file extension."""
    return f"""[-HKEY_CLASSES_ROOT\\SystemFileAssociations\\.{file_extension}\\shell\\DownloadAsOffice]
"""

def write_reg_file(filename, content):
    """Write content to a .reg file."""
    with open(filename, "w", encoding="utf-8") as reg_file:
        reg_file.write(content.strip() + "\n")

# Main Logic
def main():
    # Get paths
    script_dir = os.path.dirname(os.path.abspath(__file__))
    script_path = os.path.join(script_dir, SCRIPT_NAME)

    # Find pythonw.exe
    pythonw_path = shutil.which("pythonw")
    if not pythonw_path:
        print("⚠️  Could not find pythonw.exe. Make sure Python is installed and added to PATH.")
        sys.exit(1)

    # Verify script exists
    if not os.path.exists(script_path):
        print(f"⚠️  Script not found at: {script_path}")
        sys.exit(1)

    # Build the command string
    command_string = build_command_string(pythonw_path, script_path)

    # File extensions to configure
    file_extensions = ["gdoc", "gsheet", "gslides"]
    menu_text = "Download as Office File"

    # Generate .reg content
    reg_content = "Windows Registry Editor Version 5.00\n\n"
    for ext in file_extensions:
        reg_content += create_registry_entry(ext, menu_text, command_string)

    # Generate removal .reg content
    remove_reg_content = "Windows Registry Editor Version 5.00\n\n"
    for ext in file_extensions:
        remove_reg_content += create_removal_entry(ext)

    # Write the files
    write_reg_file(REG_FILE_NAME, reg_content)
    write_reg_file(REMOVE_REG_FILE_NAME, remove_reg_content)

    # Output results
    print(f"✅ Successfully generated {os.path.abspath(REG_FILE_NAME)}")
    print(f"✅ Successfully generated {os.path.abspath(REMOVE_REG_FILE_NAME)}")

    # Ask to apply changes
    apply_now = input("Would you like to apply the registry changes now? (y/n): ").strip().lower()
    if apply_now == "y":
        try:
            subprocess.run(["regedit.exe", "/s", REG_FILE_NAME], check=True)
            print("✅ Context menu added successfully! Right-click any .gdoc, .gsheet, or .gslides file to use it.")
        except Exception as e:
            print(f"⚠️  Failed to apply registry changes: {e}")
    else:
        print("ℹ️  You can manually apply the context menu by double-clicking the generated .reg file.")

    # Provide removal instructions
    print("\nTo remove the context menu later, run:")
    print(f"  regedit.exe /s \"{os.path.abspath(REMOVE_REG_FILE_NAME)}\"")

if __name__ == "__main__":
    main()
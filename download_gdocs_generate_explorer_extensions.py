import os
import sys
import shutil
import subprocess

# Constants
SCRIPT_NAME = "download_google_doc.pyw"
REG_FILE_NAME = "add_context_menu.reg"
REMOVE_REG_FILE_NAME = "remove_context_menu.reg"

# Get paths
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SCRIPT_PATH = os.path.join(SCRIPT_DIR, SCRIPT_NAME)

# Find the correct `pythonw.exe` location
PYTHONW_PATH = shutil.which("pythonw")
if not PYTHONW_PATH:
    print("⚠️  Could not find pythonw.exe. Make sure Python is installed and added to PATH.")
    sys.exit(1)

# Generate the .reg file content
REG_CONTENT = f"""Windows Registry Editor Version 5.00

[HKEY_CLASSES_ROOT\\SystemFileAssociations\\.gdoc\\shell\\DownloadAsOffice]
@="Download as Office File"

[HKEY_CLASSES_ROOT\\SystemFileAssociations\\.gdoc\\shell\\DownloadAsOffice\\command]
@="\\"{PYTHONW_PATH}\\" \\"{SCRIPT_PATH}\\" \\"%1\\""

[HKEY_CLASSES_ROOT\\SystemFileAssociations\\.gsheet\\shell\\DownloadAsOffice]
@="Download as Office File"

[HKEY_CLASSES_ROOT\\SystemFileAssociations\\.gsheet\\shell\\DownloadAsOffice\\command]
@="\\"{PYTHONW_PATH}\\" \\"{SCRIPT_PATH}\\" \\"%1\\""

[HKEY_CLASSES_ROOT\\SystemFileAssociations\\.gslides\\shell\\DownloadAsOffice]
@="Download as Office File"

[HKEY_CLASSES_ROOT\\SystemFileAssociations\\.gslides\\shell\\DownloadAsOffice\\command]
@="\\"{PYTHONW_PATH}\\" \\"{SCRIPT_PATH}\\" \\"%1\\""
"""

# Generate the removal .reg file
REMOVE_REG_CONTENT = """Windows Registry Editor Version 5.00

[-HKEY_CLASSES_ROOT\\SystemFileAssociations\\.gdoc\\shell\\DownloadAsOffice]

[-HKEY_CLASSES_ROOT\\SystemFileAssociations\\.gsheet\\shell\\DownloadAsOffice]

[-HKEY_CLASSES_ROOT\\SystemFileAssociations\\.gslides\\shell\\DownloadAsOffice]
"""

# Write the `.reg` files
def write_reg_file(filename, content):
    """Writes a .reg file with the given content."""
    with open(filename, "w", encoding="utf-8") as reg_file:
        reg_file.write(content)

write_reg_file(REG_FILE_NAME, REG_CONTENT)
write_reg_file(REMOVE_REG_FILE_NAME, REMOVE_REG_CONTENT)

print(f"✅ Successfully generated {REG_FILE_NAME} and {REMOVE_REG_FILE_NAME}")

# Ask to apply the registry changes
apply_now = input("Would you like to apply the registry changes now? (y/n): ").strip().lower()

if apply_now == "y":
    try:
        subprocess.run(["regedit.exe", "/s", REG_FILE_NAME], check=True)
        print("✅ Context menu added successfully! Right-click any .gdoc, .gsheet, or .gslides file to use it.")
    except Exception as e:
        print(f"⚠️  Failed to apply registry changes: {e}")
else:
    print("ℹ️  You can manually apply the context menu by double-clicking the generated .reg file.")

# Ask if user wants to enable removal option
print("\nTo remove the context menu later, run:")
print(f"  regedit.exe /s {REMOVE_REG_FILE_NAME}")

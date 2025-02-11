import os
import re
import sys
import logging
from datetime import datetime
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive

# Get the absolute path of the script directory
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# Set up logging (optional, useful for debugging)
logging.basicConfig(filename=os.path.join(SCRIPT_DIR, "download_google_doc.log"), level=logging.INFO)

# Define the path to `client_secrets.json`
CLIENT_SECRETS_PATH = os.path.join(SCRIPT_DIR, "client_secrets.json")

# Authenticate and create the PyDrive client
gauth = GoogleAuth()

try:
    gauth.LoadClientConfigFile(CLIENT_SECRETS_PATH)  # Explicitly load client_secrets.json
except InvalidConfigError:
    logging.error(f"Missing or invalid client_secrets.json file at {CLIENT_SECRETS_PATH}")
    print(f"⚠️  Error: Could not find {CLIENT_SECRETS_PATH}. Please check its location.")
    sys.exit(1)

# Load credentials from the same directory as the script
CREDENTIALS_PATH = os.path.join(SCRIPT_DIR, "credentials.json")

try:
    gauth.LoadCredentialsFile(CREDENTIALS_PATH)

    if gauth.credentials is None:
        gauth.LocalWebserverAuth()
        gauth.SaveCredentialsFile(CREDENTIALS_PATH)
    elif gauth.access_token_expired:
        gauth.Refresh()
        gauth.SaveCredentialsFile(CREDENTIALS_PATH)
    else:
        gauth.Authorize()

except Exception as e:
    logging.error(f"Google Drive authentication failed: {str(e)}")
    print(f"⚠️  Google Drive authentication failed: {str(e)}. Please re-authenticate.")

    # Delete credentials file and retry authentication
    if os.path.exists(CREDENTIALS_PATH):
        os.remove(CREDENTIALS_PATH)
    
    gauth.LocalWebserverAuth()
    gauth.SaveCredentialsFile(CREDENTIALS_PATH)

drive = GoogleDrive(gauth)

def sanitize_filename(filename):
    # Replace any characters that are not allowed in filenames
    return re.sub(r'[\\/*?:"<>|]', "", filename)

def get_export_mimetype(file_type):
    # Map Google Workspace file types to their corresponding MS Office MIME types
    mimetype_map = {
        'application/vnd.google-apps.document': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'application/vnd.google-apps.spreadsheet': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.google-apps.presentation': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    }
    return mimetype_map.get(file_type, None)

def get_extension_from_mimetype(mimetype):
    extension_map = {
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document': 'docx',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': 'xlsx',
        'application/vnd.openxmlformats-officedocument.presentationml.presentation': 'pptx',
    }
    return extension_map.get(mimetype, '')

def find_file_in_drive(path, filename, expected_mime_type):
    """Find the file in Google Drive using its folder structure and return the file ID."""
    folder_id = 'root'

    print(f"🔍 Searching for folder path: {path}")

    for folder_name in path.split(os.path.sep):
        print(f"📂 Looking for folder: {folder_name}")

        # Check if it's a shortcut
        if folder_name.startswith('.shortcut-targets-by-id'):
            target_id = folder_name.split('-')[-1]
            print(f"🔗 Using shortcut target ID: {target_id}")
            folder_id = target_id
            continue

        # Search for the folder
        query = f"'{folder_id}' in parents and trashed=false and mimeType='application/vnd.google-apps.folder' and title='{folder_name}'"
        file_list = drive.ListFile({'q': query}).GetList()
        
        if not file_list:
            print(f"⚠️  Folder '{folder_name}' not found in Google Drive.")
            raise ValueError(f"Folder '{folder_name}' not found in Google Drive.")
        
        folder_id = file_list[0]['id']
        print(f"✅ Found folder '{folder_name}', ID: {folder_id}")

    # Search for the file in the final folder
    print(f"🔍 Searching for file: {filename} with MIME type {expected_mime_type}")
    query = f"'{folder_id}' in parents and trashed=false and title='{filename}' and mimeType='{expected_mime_type}'"
    file_list = drive.ListFile({'q': query}).GetList()

    if not file_list:
        print(f"⚠️  File '{filename}' with expected MIME type '{expected_mime_type}' not found in Google Drive.")
        raise ValueError(f"File '{filename}' with expected MIME type '{expected_mime_type}' not found in Google Drive.")

    file_id = file_list[0]['id']
    print(f"✅ Found file '{filename}', ID: {file_id}")
    
    return file_id

def download_google_file_as_ms_office(file_id, original_extension, add_timestamp, output_directory):
    """Download the Google file as an MS Office document."""
    file = drive.CreateFile({'id': file_id})
    file.FetchMetadata()

    # Debugging: Print retrieved file information
    print(f"📄 Found File: {file['title']} (ID: {file_id})")
    print(f"🔍 MIME Type: {file['mimeType']}")

    # Get export format
    export_mimetype = get_export_mimetype(file['mimeType'])
    if export_mimetype is None:
        print("⚠️  Unsupported Google Workspace file type.")
        return

    # Determine the file extension (.docx, .xlsx, .pptx)
    file_extension = get_extension_from_mimetype(export_mimetype)

    # Generate timestamp if required
    timestamp = ""
    if add_timestamp:
        timestamp = "_" + datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

    # Ensure output directory exists
    os.makedirs(output_directory, exist_ok=True)

    # Construct the output filename
    output_filename = os.path.join(output_directory, f"{sanitize_filename(file['title'])}{original_extension}{timestamp}.{file_extension}")

    # Debugging: Print output filename
    print(f"⬇️  Saving as: {output_filename}")

    # Download file
    file.GetContentFile(output_filename, mimetype=export_mimetype)
    print(f"✅ File downloaded successfully: {output_filename}")

def find_deepest_match(file_path, backup_path):
    """
    Find the deepest matching folder component between the file path and backup path.
    Returns the new path where the file should be saved.
    """
    file_parts = os.path.normpath(file_path).split(os.sep)
    backup_parts = os.path.normpath(backup_path).split(os.sep)

    # Find the deepest match
    match_index = -1
    for i, folder in enumerate(file_parts):
        if folder in backup_parts:
            match_index = i

    if match_index == -1:
        print(f"⚠️  No matching component found between '{file_path}' and backup '{backup_path}'.")
        return None  # No match found

    # Construct the new backup path
    matched_part = os.path.join(*file_parts[:match_index + 1])
    remaining_part = os.path.join(*file_parts[match_index + 1:])

    # Replace the matched portion with the backup path
    new_path = os.path.join(backup_path, remaining_part)
    
    return new_path

def main(input_path, add_timestamp, backup_path=None):
    """Process the Google Doc/Sheet/Slides file from Windows Explorer."""
    input_path = os.path.abspath(input_path)

    if not os.path.exists(input_path):
        print(f"⚠️  Error: File does not exist: {input_path}")
        return

    print(f"Processing: {input_path}")

    # Extract file extension
    folder_path, full_filename = os.path.split(input_path)
    filename, file_extension = os.path.splitext(full_filename)

    # Supported Google file types
    mime_type_map = {
        '.gdoc': 'application/vnd.google-apps.document',
        '.gsheet': 'application/vnd.google-apps.spreadsheet',
        '.gslides': 'application/vnd.google-apps.presentation',
    }

    if file_extension not in mime_type_map:
        print(f"⚠️  Unsupported file type: {file_extension}")
        return

    expected_mime_type = mime_type_map[file_extension]

    # Convert the Windows-style path to a Drive-compatible path
    drive_path = os.path.relpath(folder_path, "H:\\My Drive")

    # Find the file ID in Google Drive
    file_id = find_file_in_drive(drive_path, filename, expected_mime_type)

    # Determine output directory
    if backup_path:
        backup_path = os.path.abspath(backup_path)
        new_output_path = find_deepest_match(folder_path, backup_path)
        if new_output_path:
            os.makedirs(new_output_path, exist_ok=True)  # Ensure backup directory exists
            print(f"📁 Redirecting download to: {new_output_path}")
        else:
            print("⚠️  Backup path does not match any part of the file's directory. Using default location.")
            new_output_path = folder_path
    else:
        new_output_path = folder_path

    # Download the corresponding MS Office file with optional timestamp
    download_google_file_as_ms_office(file_id, file_extension, add_timestamp, new_output_path)

if __name__ == "__main__":
    print("🚀 Script started!")  # Confirms script execution

    if len(sys.argv) < 2:
        print("Usage: python script.py <path_to_gdoc_gsheet_gslides_file> [--timestamp] [--backup <backup_path>]")
        sys.exit(1)

    # Parse arguments
    args = sys.argv[1:]
    input_path = args[0].strip('"')  # Remove extra quotes from Windows paths
    add_timestamp = "--timestamp" in args  # Check if timestamp flag is present

    # Extract backup path if --backup is provided
    backup_path = None
    if "--backup" in args:
        backup_index = args.index("--backup")
        if backup_index + 1 < len(args):  # Ensure an argument is provided
            backup_path = args[backup_index + 1].strip('"')
            print(f"🗂 Backup enabled: {backup_path}")
        else:
            print("⚠️  Error: --backup requires a path argument.")
            sys.exit(1)

    # Call main function with parsed arguments
    main(input_path, add_timestamp, backup_path)

    print("✅ Script completed!")  # Confirms script ran successfully

import os
import re
import sys
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive

# Authenticate and create the PyDrive client
gauth = GoogleAuth()

# Try to load saved client credentials
gauth.LoadCredentialsFile("credentials.json")
if gauth.credentials is None:
    # Authenticate if they're not there
    gauth.LocalWebserverAuth()
    # Save the current credentials to a file
    gauth.SaveCredentialsFile("credentials.json")
elif gauth.access_token_expired:
    # Refresh them if expired
    gauth.Refresh()
    # Save the updated credentials to a file
    gauth.SaveCredentialsFile("credentials.json")
else:
    # Initialize the saved creds
    gauth.Authorize()

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
    folder_id = 'root'
    for folder_name in path.split(os.path.sep):
        # Check if it's a shortcut
        if folder_name.startswith('.shortcut-targets-by-id'):
            # Extract the target ID from the folder name
            target_id = folder_name.split('-')[-1]
            folder_id = target_id
            continue
        
        # Search for the folder
        file_list = drive.ListFile({'q': f"'{folder_id}' in parents and trashed=false and mimeType='application/vnd.google-apps.folder' and title='{folder_name}'"}).GetList()
        if not file_list:
            raise ValueError(f"Folder '{folder_name}' not found in Google Drive.")
        folder_id = file_list[0]['id']
    
    # Search for the file in the final folder
    file_list = drive.ListFile({'q': f"'{folder_id}' in parents and trashed=false and title='{filename}' and mimeType='{expected_mime_type}'"}).GetList()
    if not file_list:
        raise ValueError(f"File '{filename}' with expected MIME type '{expected_mime_type}' not found in Google Drive.")
    
    return file_list[0]['id']

def download_google_file_as_ms_office(file_id):
    # Get the file from Google Drive
    file = drive.CreateFile({'id': file_id})
    file.FetchMetadata()
    
    # Get the title of the document and sanitize it for use as a filename
    doc_title = sanitize_filename(file['title'])
    
    # Determine the export MIME type based on the Google Workspace file type
    export_mimetype = get_export_mimetype(file['mimeType'])
    if export_mimetype is None:
        raise ValueError("Unsupported Google Workspace file type.")
    
    # Determine the appropriate file extension
    file_extension = get_extension_from_mimetype(export_mimetype)
    output_filename = f"{doc_title}.{file_extension}"
    
    # Download the file in the corresponding MS Office format
    file.GetContentFile(output_filename, mimetype=export_mimetype)
    print(f"File downloaded successfully and saved as {output_filename}")

# Main function to process the input file
def main(input_path):
    if not os.path.exists(input_path):
        print(f"Input path {input_path} does not exist.")
        return

    # Split the input path into the folder path and the filename
    folder_path, full_filename = os.path.split(input_path)
    filename, file_extension = os.path.splitext(full_filename)
    
    # Determine the expected MIME type based on the file extension
    mime_type_map = {
        '.gdoc': 'application/vnd.google-apps.document',
        '.gsheet': 'application/vnd.google-apps.spreadsheet',
        '.gslides': 'application/vnd.google-apps.presentation',
    }
    
    if file_extension not in mime_type_map:
        raise ValueError(f"Unsupported file extension '{file_extension}'. Only .gdoc, .gsheet, and .gslides are supported.")
    
    expected_mime_type = mime_type_map[file_extension]
    
    # Convert the Windows-style path to a Drive-compatible path
    drive_path = os.path.relpath(folder_path, "H:\\My Drive")
    
    # Find the file ID in Google Drive
    file_id = find_file_in_drive(drive_path, filename, expected_mime_type)
    
    # Download the corresponding MS Office file
    download_google_file_as_ms_office(file_id)

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python script.py <path_to_gdoc_gsheet_gslides_file>")
        sys.exit(1)

    input_path = sys.argv[1]
    main(input_path)

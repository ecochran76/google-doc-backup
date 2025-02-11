import os
import re
import sys
import logging
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive

# Set up logging (useful for debugging)
logging.basicConfig(filename=os.path.expanduser("~/download_google_doc.log"), level=logging.INFO)

# Authenticate and create the PyDrive client
gauth = GoogleAuth()
gauth.LoadCredentialsFile("credentials.json")

if gauth.credentials is None:
    gauth.LocalWebserverAuth()
    gauth.SaveCredentialsFile("credentials.json")
elif gauth.access_token_expired:
    gauth.Refresh()
    gauth.SaveCredentialsFile("credentials.json")
else:
    gauth.Authorize()

drive = GoogleDrive(gauth)

def sanitize_filename(filename):
    """Remove invalid filename characters."""
    return re.sub(r'[\\/*?:"<>|]', "", filename)

def get_export_mimetype(file_type):
    """Map Google Workspace file types to their corresponding MS Office MIME types."""
    mimetype_map = {
        'application/vnd.google-apps.document': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'application/vnd.google-apps.spreadsheet': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.google-apps.presentation': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    }
    return mimetype_map.get(file_type, None)

def get_extension_from_mimetype(mimetype):
    """Return the correct file extension based on MIME type."""
    extension_map = {
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document': 'docx',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': 'xlsx',
        'application/vnd.openxmlformats-officedocument.presentationml.presentation': 'pptx',
    }
    return extension_map.get(mimetype, '')

def find_file_in_drive(path, filename, expected_mime_type):
    """Find the file in Google Drive using its folder structure and return the file ID."""
    folder_id = 'root'
    for folder_name in path.split(os.path.sep):
        if folder_name.startswith('.shortcut-targets-by-id'):
            target_id = folder_name.split('-')[-1]
            folder_id = target_id
            continue
        
        file_list = drive.ListFile({'q': f"'{folder_id}' in parents and trashed=false and mimeType='application/vnd.google-apps.folder' and title='{folder_name}'"}).GetList()
        if not file_list:
            raise ValueError(f"Folder '{folder_name}' not found in Google Drive.")
        folder_id = file_list[0]['id']
    
    file_list = drive.ListFile({'q': f"'{folder_id}' in parents and trashed=false and title='{filename}' and mimeType='{expected_mime_type}'"}).GetList()
    if not file_list:
        raise ValueError(f"File '{filename}' with expected MIME type '{expected_mime_type}' not found in Google Drive.")
    
    return file_list[0]['id']

def download_google_file_as_ms_office(file_id):
    """Download the Google file as an MS Office document."""
    file = drive.CreateFile({'id': file_id})
    file.FetchMetadata()
    
    doc_title = sanitize_filename(file['title'])
    export_mimetype = get_export_mimetype(file['mimeType'])
    
    if export_mimetype is None:
        raise ValueError("Unsupported Google Workspace file type.")
    
    file_extension = get_extension_from_mimetype(export_mimetype)
    output_filename = f"{doc_title}.{file_extension}"
    
    file.GetContentFile(output_filename, mimetype=export_mimetype)
    logging.info(f"File downloaded successfully as {output_filename}")

def main(input_path):
    """Process the Google Doc/Sheet/Slides file from Windows Explorer."""
    if not os.path.exists(input_path):
        logging.error(f"Input path does not exist: {input_path}")
        return

    folder_path, full_filename = os.path.split(input_path)
    filename, file_extension = os.path.splitext(full_filename)
    
    mime_type_map = {
        '.gdoc': 'application/vnd.google-apps.document',
        '.gsheet': 'application/vnd.google-apps.spreadsheet',
        '.gslides': 'application/vnd.google-apps.presentation',
    }
    
    if file_extension not in mime_type_map:
        logging.error(f"Unsupported file extension: {file_extension}")
        return
    
    expected_mime_type = mime_type_map[file_extension]
    drive_path = os.path.relpath(folder_path, "H:\\My Drive")
    
    file_id = find_file_in_drive(drive_path, filename, expected_mime_type)
    download_google_file_as_ms_office(file_id)

if __name__ == "__main__":
    if len(sys.argv) < 2:
        logging.error("No file specified.")
        sys.exit(1)

    input_path = sys.argv[1].strip('"')  # Handle paths passed with quotes
    main(input_path)

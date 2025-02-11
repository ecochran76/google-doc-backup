import os
import json
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

def extract_file_id(file_path):
    try:
        with open(file_path, 'r') as file:
            content = file.read()
            file_data = json.loads(content)
            return file_data['doc_id']
    except Exception as e:
        raise ValueError(f"Failed to extract file ID from {file_path}: {e}")

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
    file_extension = export_mimetype.split('.')[-1]
    output_filename = f"{doc_title}.{file_extension}"
    
    # Download the file in the corresponding MS Office format
    file.GetContentFile(output_filename, mimetype=export_mimetype)
    print(f"File downloaded successfully and saved as {output_filename}")

# Main function to process the input file
def main(input_file):
    if not os.path.exists(input_file):
        print(f"Input file {input_file} does not exist.")
        return

    # Extract the Drive ID from the input file
    file_id = extract_file_id(input_file)
    
    # Download the corresponding MS Office file
    download_google_file_as_ms_office(file_id)

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python script.py <path_to_gdoc_gsheet_gslides_file>")
        sys.exit(1)

    input_file = sys.argv[1]
    main(input_file)

from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import re
import os

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

def extract_file_id(file_url):
    # Use regex to extract the file ID from the URL
    match = re.search(r'/d/([a-zA-Z0-9-_]+)', file_url)
    if match:
        return match.group(1)
    else:
        raise ValueError("Could not extract file ID from the URL.")

def sanitize_filename(filename):
    # Replace any characters that are not allowed in filenames
    return re.sub(r'[\\/*?:"<>|]', "", filename)

def download_google_doc_as_docx(file_url):
    # Extract file ID from the Google Doc URL
    file_id = extract_file_id(file_url)
    
    # Get the file from Google Drive
    file = drive.CreateFile({'id': file_id})
    file.FetchMetadata()
    
    # Get the title of the document and sanitize it for use as a filename
    doc_title = sanitize_filename(file['title'])
    
    # Set the output filename with .docx extension
    output_filename = f"{doc_title}.docx"
    
    # Download the file as DOCX
    file.GetContentFile(output_filename, mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    print(f"File downloaded successfully and saved as {output_filename}")

# Example usage
file_url = 'https://docs.google.com/document/d/1X7fBINZl-5rHuZ5-I_YclaJdzocW5CU1q1_M4p2q_js?usp=drive_fs'
download_google_doc_as_docx(file_url)

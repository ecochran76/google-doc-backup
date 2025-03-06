import os
import re
import sys
import glob
import logging
import argparse
from datetime import datetime, timedelta
try:
    from dateutil import parser as date_parser
except ImportError:
    print("The 'python-dateutil' package is required. Please install it using pip install python-dateutil")
    sys.exit(1)
from pydrive.auth import GoogleAuth, InvalidConfigError
from pydrive.drive import GoogleDrive

# Get the absolute path of the script directory
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# Set up logging
log_file = os.path.join(SCRIPT_DIR, "download_google_doc.log")
logging.basicConfig(filename=log_file, level=logging.INFO, 
                    format='%(asctime)s %(levelname)s: %(message)s')

# Global cache to store folder IDs (to reduce redundant API calls)
folder_cache = {}

# Define the path to client_secrets.json
CLIENT_SECRETS_PATH = os.path.join(SCRIPT_DIR, "client_secrets.json")

# Authenticate and create the PyDrive client
gauth = GoogleAuth()
try:
    gauth.LoadClientConfigFile(CLIENT_SECRETS_PATH)
except InvalidConfigError:
    logging.error(f"Missing or invalid client_secrets.json file at {CLIENT_SECRETS_PATH}")
    print(f"⚠️  Error: Could not find {CLIENT_SECRETS_PATH}. Please check its location.")
    sys.exit(1)

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
    logging.exception("Google Drive authentication failed")
    print(f"⚠️  Google Drive authentication failed: {e}. Please re-authenticate.")
    if os.path.exists(CREDENTIALS_PATH):
        os.remove(CREDENTIALS_PATH)
    gauth.LocalWebserverAuth()
    gauth.SaveCredentialsFile(CREDENTIALS_PATH)

drive = GoogleDrive(gauth)

def sanitize_filename(filename):
    """Replace invalid filename characters with underscores."""
    return re.sub(r'[\\/*?:"<>|]', "_", filename)

def get_export_mimetype(file_type):
    """Map Google Workspace file types to their corresponding MS Office MIME types."""
    mapping = {
        'application/vnd.google-apps.document': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'application/vnd.google-apps.spreadsheet': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.google-apps.presentation': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    }
    return mapping.get(file_type)

def get_original_extension(mime_type):
    """Map Google Workspace MIME types to their original file extensions."""
    mapping = {
        'application/vnd.google-apps.document': 'gdoc',
        'application/vnd.google-apps.spreadsheet': 'gsheet',
        'application/vnd.google-apps.presentation': 'gslides',
    }
    return mapping.get(mime_type, '')

def get_export_extension(export_mime_type):
    """Map MS Office MIME types to their file extensions."""
    mapping = {
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document': 'docx',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': 'xlsx',
        'application/vnd.openxmlformats-officedocument.presentationml.presentation': 'pptx',
    }
    return mapping.get(export_mime_type, '')

def parse_date_input(date_str):
    """
    Parse a date string into RFC 3339 UTC format.
    
    Accepts absolute dates in many common formats as well as negative values.
    A negative value (e.g. "-3600", "-1h", "-1 day") is treated as a relative time offset from now.
    """
    if not date_str.startswith('-'):
        try:
            dt = date_parser.parse(date_str)
            # If dt is naive, assume local time then convert to UTC
            if dt.tzinfo is None:
                dt = dt.astimezone().astimezone(tz=timedelta(0))
            else:
                dt = dt.astimezone(tz=timedelta(0))
            return dt.strftime("%Y-%m-%dT%H:%M:%SZ")
        except Exception as e:
            print(f"Error parsing date '{date_str}': {e}")
            sys.exit(1)
    else:
        # Interpret as a relative time offset from now.
        rel = date_str[1:].strip()
        m = re.match(r"(\d+)\s*([smhdw]?)", rel)
        if m:
            value = int(m.group(1))
            unit = m.group(2)
            if unit in ['s', '']:
                delta = timedelta(seconds=value)
            elif unit == 'm':
                delta = timedelta(minutes=value)
            elif unit == 'h':
                delta = timedelta(hours=value)
            elif unit == 'd':
                delta = timedelta(days=value)
            elif unit == 'w':
                delta = timedelta(weeks=value)
            else:
                print(f"Unrecognized time unit in relative date: {date_str}")
                sys.exit(1)
            dt = datetime.utcnow() - delta
            return dt.strftime("%Y-%m-%dT%H:%M:%SZ")
        else:
            print(f"Error parsing relative time: {date_str}")
            sys.exit(1)

def find_folder_id(drive_path):
    """Find the folder ID in Google Drive based on the relative path from root.
       Uses a global cache to avoid repeated API calls."""
    global folder_cache
    if not drive_path or drive_path == ".":
        return 'root'
    if drive_path in folder_cache:
        return folder_cache[drive_path]

    folder_id = 'root'
    for folder_name in drive_path.split(os.path.sep):
        if not folder_name:
            continue
        logging.info(f"Looking for folder: {folder_name}")
        print(f"📂 Looking for folder: {folder_name}")

        if folder_name.startswith('.shortcut-targets-by-id'):
            target_id = folder_name.split('-')[-1]
            logging.info(f"Using shortcut target ID: {target_id}")
            print(f"🔗 Using shortcut target ID: {target_id}")
            folder_id = target_id
            continue

        query = f"'{folder_id}' in parents and trashed=false and mimeType='application/vnd.google-apps.folder' and title='{folder_name}'"
        file_list = drive.ListFile({'q': query}).GetList()
        if not file_list:
            logging.error(f"Folder '{folder_name}' not found in Google Drive.")
            print(f"⚠️  Folder '{folder_name}' not found in Google Drive.")
            raise ValueError(f"Folder '{folder_name}' not found in Google Drive.")
        folder_id = file_list[0]['id']
        logging.info(f"Found folder '{folder_name}', ID: {folder_id}")
        print(f"✅ Found folder '{folder_name}', ID: {folder_id}")

    folder_cache[drive_path] = folder_id
    return folder_id

def strip_duplicate_suffix(filename):
    """Remove the '(number)' suffix added by Windows for duplicates, preserving trailing spaces."""
    match = re.match(r"^(.*?)\s*\(\d+\)$", filename)
    if match:
        return match.group(1)
    return filename

def find_files_in_drive(drive, folder_id, base_path="", filename=None, depth=0, max_depth=float('inf'),
                        newer_than=None, older_than=None):
    """Find Google Docs/Sheets/Slides files recursively. Groups files by title and recurses into subfolders."""
    mime_types = {
        'application/vnd.google-apps.document',
        'application/vnd.google-apps.spreadsheet',
        'application/vnd.google-apps.presentation',
    }
    
    query = f"'{folder_id}' in parents and trashed=false"
    if filename:
        base_filename = strip_duplicate_suffix(filename)
        query += f" and title='{base_filename}'"
    if newer_than:
        query += f" and modifiedDate > '{newer_than}'"
    if older_than:
        query += f" and modifiedDate < '{older_than}'"

    file_list = drive.ListFile({'q': query}).GetList()
    matching_files = [f for f in file_list if f['mimeType'] in mime_types]
    folders = [f for f in file_list if f['mimeType'] == 'application/vnd.google-apps.folder']

    # Group matching files by title
    files_by_title = {}
    for file in matching_files:
        title = file['title']
        files_by_title.setdefault(title, []).append(file)

    # Sort each group by createdDate and add a suffix for duplicates
    for title, files in files_by_title.items():
        files.sort(key=lambda x: x.get('createdDate', ''))
        for i, f in enumerate(files):
            suffix = "" if i == 0 else f" ({i})"
            files_by_title[title][i] = (f['id'], suffix, f['mimeType'], base_path, f)
            logging.info(f"File '{title}{suffix}', ID: {f['id']}, MIME: {f['mimeType']}, Path: {base_path}, Created: {f.get('createdDate', 'N/A')}")
            print(f"✅ File '{title}{suffix}', ID: {f['id']}, MIME: {f['mimeType']}, Path: {base_path}, Created: {f.get('createdDate', 'N/A')}")

    # Recurse into subfolders if depth is less than max_depth
    if depth < max_depth:
        for folder in folders:
            subfolder_id = folder['id']
            subfolder_title = folder['title']
            subfolder_path = os.path.join(base_path, subfolder_title) if base_path else subfolder_title
            logging.info(f"Recursing into subfolder: {subfolder_title} (ID: {subfolder_id}) at depth {depth + 1}")
            print(f"📁 Recursing into subfolder: {subfolder_title} (ID: {subfolder_id})")
            subfolder_files = find_files_in_drive(drive, subfolder_id, subfolder_path, filename,
                                                  depth + 1, max_depth, newer_than, older_than)
            for title, file_entries in subfolder_files.items():
                files_by_title.setdefault(title, []).extend(file_entries)

    total_files = sum(len(files) for files in files_by_title.values())
    logging.info(f"Found {total_files} Google Docs/Sheets/Slides files in folder ID '{folder_id}' (depth {depth}).")
    print(f"✅ Found {total_files} Google Docs/Sheets/Slides files at depth {depth}.")
    return files_by_title

def download_google_file_as_ms_office(file_id, suffix, mime_type, subfolder_path, add_timestamp, 
                                      output_directory, dry_run=False, file_meta=None):
    """Download a Google file as an MS Office document, preserving the subfolder structure.
       Uses provided metadata if available to avoid extra API calls."""
    meta = file_meta if file_meta else drive.CreateFile({'id': file_id}).FetchMetadata()
    logging.info(f"Retrieved file: {meta['title']} (ID: {file_id}, MIME: {meta['mimeType']})")
    print(f"📄 Found File: {meta['title']} (ID: {file_id})")
    print(f"🔍 MIME Type: {meta['mimeType']}")
    export_mimetype = get_export_mimetype(meta['mimeType'])
    if export_mimetype is None:
        logging.error(f"Unsupported MIME type: {meta['mimeType']}")
        print("⚠️  Unsupported Google Workspace file type.")
        return
    original_extension = f".{get_original_extension(mime_type)}"
    file_extension = get_export_extension(export_mimetype)
    timestamp = "_" + datetime.now().strftime("%Y-%m-%d_%H-%M-%S") if add_timestamp else ""
    full_output_directory = os.path.join(output_directory, subfolder_path)
    if not dry_run:
        os.makedirs(full_output_directory, exist_ok=True)
    output_filename = os.path.join(full_output_directory,
                                   f"{sanitize_filename(meta['title'])}{suffix}{original_extension}{timestamp}.{file_extension}")
    logging.info(f"{'[Dry Run] Would download' if dry_run else 'Downloading'} to: {output_filename}")
    print(f"{'📜 [Dry Run] Would save' if dry_run else '⬇️  Saving'} as: {output_filename}")
    if not dry_run:
        file_to_download = drive.CreateFile({'id': file_id})
        file_to_download.GetContentFile(output_filename, mimetype=export_mimetype)
        logging.info(f"Download successful: {output_filename}")
        print(f"✅ File downloaded successfully: {output_filename}")

def find_deepest_match(file_path, backup_path):
    """Find the deepest matching folder component between the file path and backup path, starting from the root."""
    file_parts = os.path.normpath(file_path).split(os.sep)
    backup_parts = os.path.normpath(backup_path).split(os.sep)
    match_index = -1
    for i in range(min(len(file_parts), len(backup_parts))):
        if file_parts[i] != backup_parts[i]:
            break
        match_index = i
    if match_index < 0:
        logging.warning(f"No matching component found between '{file_path}' and '{backup_path}'.")
        print(f"⚠️  No matching component found between '{file_path}' and backup '{backup_path}'.")
        return backup_path
    remaining_part = os.path.join(*file_parts[match_index + 1:]) if (match_index + 1) < len(file_parts) else ""
    new_path = os.path.join(backup_path, remaining_part)
    return new_path

def process_path(input_path, add_timestamp, backup_path, dry_run, max_depth, newer_than, older_than):
    """Process a single input path (file or directory)."""
    input_path = os.path.abspath(input_path)
    mime_type_map = {
        '.gdoc': 'application/vnd.google-apps.document',
        '.gsheet': 'application/vnd.google-apps.spreadsheet',
        '.gslides': 'application/vnd.google-apps.presentation',
    }
    try:
        if os.path.isdir(input_path):
            logging.info(f"Input is a directory: {input_path}")
            print(f"📁 Processing directory: {input_path}")
            drive_path = os.path.relpath(input_path, "H:\\My Drive")
            folder_id = find_folder_id(drive_path)
            files_by_title = find_files_in_drive(drive, folder_id, max_depth=max_depth,
                                                 newer_than=newer_than, older_than=older_than)
            output_directory = input_path if not backup_path else find_deepest_match(input_path, os.path.abspath(backup_path))
            if output_directory:
                if not dry_run:
                    os.makedirs(output_directory, exist_ok=True)
                if backup_path:
                    logging.info(f"Redirecting download to: {output_directory}")
                    print(f"📁 Redirecting download to: {output_directory}")
            else:
                logging.warning("Backup path does not match; using default location.")
                print("⚠️  Backup path does not match any part of the directory. Using default location.")
                output_directory = input_path
            for title, file_entries in files_by_title.items():
                for file_id, suffix, mime_type, subfolder_path, file_meta in file_entries:
                    download_google_file_as_ms_office(file_id, suffix, mime_type, subfolder_path, add_timestamp,
                                                      output_directory, dry_run, file_meta)
        else:
            if not os.path.exists(input_path):
                logging.error(f"File does not exist: {input_path}")
                print(f"⚠️  Error: File does not exist: {input_path}")
                return
            logging.info(f"Processing file: {input_path}")
            print(f"Processing: {input_path}")
            folder_path, full_filename = os.path.split(input_path)
            filename, file_extension = os.path.splitext(full_filename)
            if file_extension not in mime_type_map:
                logging.error(f"Unsupported file type: {file_extension}")
                print(f"⚠️  Unsupported file type: {file_extension}")
                return
            drive_path = os.path.relpath(folder_path, "H:\\My Drive")
            folder_id = find_folder_id(drive_path)
            files_by_title = find_files_in_drive(drive, folder_id, filename=filename,
                                                 max_depth=max_depth, newer_than=newer_than, older_than=older_than)
            output_directory = folder_path if not backup_path else find_deepest_match(folder_path, os.path.abspath(backup_path))
            if output_directory:
                if not dry_run:
                    os.makedirs(output_directory, exist_ok=True)
                if backup_path:
                    logging.info(f"Redirecting download to: {output_directory}")
                    print(f"📁 Redirecting download to: {output_directory}")
            else:
                logging.warning("Backup path does not match; using default location.")
                print("⚠️  Backup path does not match any part of the file's directory. Using default location.")
                output_directory = folder_path
            base_filename = strip_duplicate_suffix(filename)
            if base_filename in files_by_title:
                for file_id, suffix, mime_type, subfolder_path, file_meta in files_by_title[base_filename]:
                    download_google_file_as_ms_office(file_id, suffix, mime_type, subfolder_path, add_timestamp,
                                                      output_directory, dry_run, file_meta)
            else:
                logging.error(f"File '{base_filename}' not found in Google Drive.")
                print(f"⚠️  File '{base_filename}' not found in Google Drive.")
    except Exception as e:
        logging.exception(f"Processing failed for {input_path}")
        print(f"⚠️  Processing failed for {input_path}: {e}")

def get_file_drive_path(file_meta):
    """
    Given a file's metadata (which should include a 'parents' field),
    compute a relative folder path representing its location in Drive.
    If the top-level folder is "My Drive", it is removed; otherwise, 
    "Shared With Me" is prepended.
    """
    parents = file_meta.get("parents", [])
    if not parents:
        return ""
    # Use the first parent ID.
    parent_id = parents[0]["id"] if isinstance(parents[0], dict) else parents[0]
    path_parts = []
    # Walk up the folder tree until reaching the root.
    while parent_id and parent_id != "root":
        if parent_id in folder_cache:
            folder_info = folder_cache[parent_id]
        else:
            folder = drive.CreateFile({'id': parent_id})
            folder.FetchMetadata(fields="id,title,parents")
            folder_info = {
                "title": folder["title"],
                "parents": folder.get("parents", [])
            }
            folder_cache[parent_id] = folder_info
        path_parts.append(folder_info["title"])
        if folder_info["parents"]:
            parent_id = folder_info["parents"][0]["id"] if isinstance(folder_info["parents"][0], dict) else folder_info["parents"][0]
        else:
            break
    # Reverse the list so the path starts from the top.
    reversed_path = list(reversed(path_parts))
    if reversed_path:
        if reversed_path[0] == "My Drive":
            # Remove "My Drive" from the path.
            reversed_path = reversed_path[1:]
        else:
            # For files not in "My Drive", assume they are in a shared folder.
            reversed_path.insert(0, "Shared With Me")
    if not reversed_path:
        return ""
    return os.path.join(*reversed_path)


def process_global_search(add_timestamp, backup_path, dry_run, max_depth, newer_than, older_than, title_filter):
    """
    Perform a global search over the entire Drive using date filters and an optional title filter.
    The query is built without a folder constraint, so files matching the criteria from anywhere in Drive are returned.
    For each file, the folder hierarchy is computed from its parents so that it is downloaded into the appropriate
    subfolder under the backup path.
    """
    # Build a query that omits any "in parents" clause.
    mime_types = (
        "mimeType = 'application/vnd.google-apps.document' or "
        "mimeType = 'application/vnd.google-apps.spreadsheet' or "
        "mimeType = 'application/vnd.google-apps.presentation'"
    )
    query = f"trashed=false and ({mime_types})"
    if title_filter:
        query += f" and title='{title_filter}'"
    if newer_than:
        query += f" and modifiedDate > '{newer_than}'"
    if older_than:
        query += f" and modifiedDate < '{older_than}'"
    
    file_list = drive.ListFile({'q': query}).GetList()
    
    output_directory = os.path.abspath(backup_path) if backup_path else os.getcwd()
    if not dry_run:
        os.makedirs(output_directory, exist_ok=True)
    if backup_path:
        logging.info(f"Redirecting download to: {output_directory}")
        print(f"📁 Redirecting download to: {output_directory}")
    
    for file in file_list:
        # Compute the relative folder path from the file's parent metadata.
        subfolder_path = get_file_drive_path(file)
        # For simplicity, we assign an empty suffix (you could add duplicate handling if needed)
        suffix = ""
        download_google_file_as_ms_office(
            file_id=file['id'],
            suffix=suffix,
            mime_type=file['mimeType'],
            subfolder_path=subfolder_path,
            add_timestamp=add_timestamp,
            output_directory=output_directory,
            dry_run=dry_run,
            file_meta=file
        )



def parse_arguments():
    """Parse command line arguments using argparse for better robustness."""
    parser = argparse.ArgumentParser(description="Download Google Docs/Sheets/Slides as Office files.")
    # Make paths optional (nargs="*")
    parser.add_argument("paths", nargs="*", help="Local input file or directory paths (wildcards supported). Leave empty for a global Drive search.")
    parser.add_argument("--timestamp", action="store_true", help="Append timestamp to output filename.")
    parser.add_argument("--backup", type=str, help="Redirect downloads to backup path.")
    parser.add_argument("--dry-run", action="store_true", help="Simulate download without writing files.")
    parser.add_argument("--max-depth", type=int, default=float('inf'), help="Maximum recursion depth for folder search.")
    parser.add_argument("--newer-than", type=str, help="Filter files newer than given date. Accepts absolute dates (e.g., '2023-01-01T00:00:00') or relative times (e.g., '-1d', '-3600').")
    parser.add_argument("--older-than", type=str, help="Filter files older than given date. Accepts absolute dates or relative times.")
    parser.add_argument("--title", type=str, help="Search for files with a specific title.")
    return parser.parse_args()

def main():
    print("🚀 Script started!")
    args = parse_arguments()
    # Process date arguments into RFC 3339 format
    if args.newer_than:
        args.newer_than = parse_date_input(args.newer_than)
    if args.older_than:
        args.older_than = parse_date_input(args.older_than)
    # If no input paths are provided, perform a global search.
    if not args.paths:
        print("No input paths provided. Performing a global search on your Drive.")
        process_global_search(args.timestamp, args.backup, args.dry_run, args.max_depth,
                              args.newer_than, args.older_than, args.title)
    else:
        # Expand wildcards for input paths
        input_paths = []
        for path in args.paths:
            expanded = glob.glob(path.strip('"'))
            if expanded:
                input_paths.extend(expanded)
            else:
                input_paths.append(path.strip('"'))
        for input_path in input_paths:
            process_path(input_path, args.timestamp, args.backup, args.dry_run, args.max_depth,
                         args.newer_than, args.older_than)
    print("✅ Script completed!")

if __name__ == "__main__":
    main()

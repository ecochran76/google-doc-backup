#!/usr/bin/env python
# download_google_doc.pyw

import os
import re
import sys
import glob
import json
import math
import logging
import argparse
import subprocess
import shutil
import io
from datetime import datetime, timedelta

try:
    from dateutil import parser as date_parser
except ImportError:
    print("The 'python-dateutil' package is required. Please install it using pip install python-dateutil")
    sys.exit(1)

from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
from googleapiclient.discovery import build  # FIX: Import for v3 service to fetch shared drive names
from googleapiclient.http import MediaIoBaseDownload
from googleapiclient.errors import HttpError

# Get the absolute path of the script directory
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# Set up logging
log_file = os.path.join(SCRIPT_DIR, "download_google_doc.log")
logging.basicConfig(filename=log_file, level=logging.INFO,
                    format='%(asctime)s %(levelname)s: %(message)s')

# Global cache to store folder IDs (and folder metadata for shared files)
folder_cache = {}

# Define the path to client_secrets.json and credentials.json
CLIENT_SECRETS_PATH = os.path.join(SCRIPT_DIR, "client_secrets.json")
CREDENTIALS_PATH = os.path.join(SCRIPT_DIR, "credentials.json")

def load_and_authenticate():
    """
    Load credentials from file and ensure they are valid.
    If credentials are missing, expired, or invalid, remove stale credentials
    and reauthenticate using LocalWebserverAuth.
    Returns a GoogleAuth instance with valid credentials.
    """
    gauth = GoogleAuth()
    gauth.settings['client_config_file'] = CLIENT_SECRETS_PATH  # Specify the client secrets file path
    # Attempt to load stored credentials.
    if os.path.exists(CREDENTIALS_PATH):
        try:
            gauth.LoadCredentialsFile(CREDENTIALS_PATH)
        except Exception as e:
            logging.error("Error loading credentials: %s", e, exc_info=True)
            os.remove(CREDENTIALS_PATH)
            gauth.credentials = None

    # If credentials exist but the token is expired, remove them.
    if gauth.credentials and gauth.access_token_expired:
        logging.info("Credentials expired. Removing stale credentials.")
        os.remove(CREDENTIALS_PATH)
        gauth.credentials = None

    # If no valid credentials are found, perform authentication.
    if not gauth.credentials:
        print("No valid credentials found. Authenticating with Google...")
        gauth.LocalWebserverAuth()
        gauth.SaveCredentialsFile(CREDENTIALS_PATH)
    else:
        try:
            gauth.Authorize()
        except Exception as e:
            logging.warning("Authorization error: %s", e, exc_info=True)
            if os.path.exists(CREDENTIALS_PATH):
                os.remove(CREDENTIALS_PATH)
            print("Authorization failed. Reauthenticating with Google...")
            gauth.LocalWebserverAuth()
            gauth.SaveCredentialsFile(CREDENTIALS_PATH)
    return gauth

# Authenticate and create the drive instance.
gauth = load_and_authenticate()
drive = GoogleDrive(gauth)
drive_v3 = build('drive', 'v3', credentials=gauth.credentials)  # FIX: Create v3 service for shared drive names

# ListFile params to search across My Drive, shared-with-me, and shared drives
ALL_DRIVES_PARAMS = {
    # Newer flags
    'corpora': 'allDrives',
    'includeItemsFromAllDrives': True,
    'supportsAllDrives': True,
    # Legacy flags (keep for older PyDrive)
    'supportsTeamDrives': True,
    'includeTeamDriveItems': True,
}

SHARED_DRIVE_MARKERS = {
    'shared drives',
    'shared drive',
    'team drives',
    'team drive',
}

shared_drive_id_cache = {}
shared_drive_name_cache = {}


def normalize_local_drive_path(local_path):
    """Derive a Drive-relative path (e.g. 'Shared drives/Team') from a local Drive for desktop path."""
    if not local_path:
        return ''
    normalized = os.path.normpath(local_path)
    parts = [segment for segment in normalized.split(os.path.sep) if segment]
    lowered = [segment.strip().lower() for segment in parts]

    def rebuild(start_index, prefix=None):
        remainder = parts[start_index + 1:]
        if prefix is None:
            return os.path.join(*remainder) if remainder else ''
        return os.path.join(prefix, *remainder) if remainder else prefix

    if '.shortcut-targets-by-id' in lowered:
        idx = lowered.index('.shortcut-targets-by-id')
        return os.path.join(*parts[idx:])

    for marker in SHARED_DRIVE_MARKERS:
        marker_lower = marker.lower()
        if marker_lower in lowered:
            idx = lowered.index(marker_lower)
            return rebuild(idx, 'Shared drives')

    for candidate in ('my drive', 'drive'):
        if candidate in lowered:
            idx = lowered.index(candidate)
            return rebuild(idx, None)

    return os.path.relpath(normalized, os.path.dirname(normalized))


def list_drive_files(params):
    merged = dict(ALL_DRIVES_PARAMS)
    if params:
        merged.update(params)
    return drive.ListFile(merged).GetList()


def ensure_shared_drive_cache():
    global shared_drive_id_cache, shared_drive_name_cache
    if shared_drive_name_cache:
        return
    if drive_v3 is None:
        return
    page_token = None
    while True:
        try:
            response = drive_v3.drives().list(pageSize=100, pageToken=page_token).execute()
        except HttpError as e:
            logging.warning('Failed to list shared drives: %s', e)
            break
        for drive_meta in response.get('drives', []):
            drive_id = drive_meta.get('id')
            if not drive_id:
                continue
            shared_drive_id_cache[drive_id] = drive_meta
            name_key = drive_meta.get('name', '').strip().lower()
            if name_key:
                shared_drive_name_cache[name_key] = drive_meta
        page_token = response.get('nextPageToken')
        if not page_token:
            break


def get_shared_drive_metadata_by_name(name):
    if not name:
        return None
    key = name.strip().lower()
    if key in shared_drive_name_cache:
        return shared_drive_name_cache[key]
    ensure_shared_drive_cache()
    return shared_drive_name_cache.get(key)


def get_shared_drive_name(drive_id):
    if not drive_id:
        return None
    if drive_id in shared_drive_id_cache:
        return shared_drive_id_cache[drive_id].get('name')
    if drive_v3 is None:
        return None
    try:
        response = drive_v3.drives().get(driveId=drive_id).execute()
    except HttpError as e:
        logging.warning('Failed to fetch shared drive name for %s: %s', drive_id, e)
        return None
    if response:
        shared_drive_id_cache[drive_id] = response
        name_key = response.get('name', '').strip().lower()
        if name_key:
            shared_drive_name_cache[name_key] = response
        return response.get('name')
    return None


def fetch_folder_metadata(folder_id):
    global folder_cache
    cached = folder_cache.get(folder_id)
    if isinstance(cached, dict) and 'title' in cached and 'parents' in cached and 'driveId' in cached:
        return cached
    metadata = None
    if drive_v3 is not None:
        try:
            response = drive_v3.files().get(
                fileId=folder_id,
                fields='id,name,parents,driveId',
                supportsAllDrives=True
            ).execute()
        except HttpError as e:
            logging.error('Failed to fetch metadata via Drive v3 for folder ID %s: %s', folder_id, e)
        else:
            metadata = {
                'title': response.get('name', ''),
                'parents': response.get('parents', []),
                'driveId': response.get('driveId'),
            }
    if metadata is None:
        file_obj = drive.CreateFile({'id': folder_id})
        file_obj['supportsAllDrives'] = True
        file_obj['supportsTeamDrives'] = True
        try:
            file_obj.FetchMetadata(fields='id,title,parents,driveId')
        except Exception as e:
            logging.error('Failed to fetch metadata for folder ID %s: %s', folder_id, e)
            return None
        metadata = {
            'title': file_obj.get('title', ''),
            'parents': file_obj.get('parents', []),
            'driveId': file_obj.get('driveId'),
        }
    folder_cache[folder_id] = metadata
    return metadata

def add_long_path_prefix(path):
    """
    On Windows, prepend the extended-length path prefix if not already present.
    This allows you to use paths longer than 260 characters.
    """
    if os.name == "nt":
        abs_path = os.path.abspath(path)
        if abs_path.startswith("\\\\?\\"):
            return abs_path
        if abs_path.startswith("\\\\"):
            # UNC paths
            return "\\\\?\\UNC\\" + abs_path.lstrip("\\")
        else:
            return "\\\\?\\" + abs_path
    return path

def sanitize_filename(filename, max_length=100):
    """
    Replace invalid characters with underscores, strip whitespace,
    and optionally truncate the filename (excluding extension) to max_length.
    """
    sanitized = re.sub(r'[\\/*?:"<>|]', "_", filename).strip()
    name, ext = os.path.splitext(sanitized)
    if len(name) > max_length:
        name = name[:max_length]
    return name + ext

def get_export_mimetype(file_type):
    """
    Map Google Workspace file types to their corresponding MS Office MIME types.
    """
    mapping = {
        'application/vnd.google-apps.document': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'application/vnd.google-apps.spreadsheet': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.google-apps.presentation': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    }
    return mapping.get(file_type)

def get_original_extension(mime_type):
    """
    Map Google Workspace MIME types to their original file extensions.
    """
    mapping = {
        'application/vnd.google-apps.document': 'gdoc',
        'application/vnd.google-apps.spreadsheet': 'gsheet',
        'application/vnd.google-apps.presentation': 'gslides',
    }
    return mapping.get(mime_type, '')

def get_export_extension(export_mime_type):
    """
    Map MS Office MIME types to their file extensions.
    """
    mapping = {
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document': 'docx',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': 'xlsx',
        'application/vnd.openxmlformats-officedocument.presentationml.presentation': 'pptx',
    }
    return mapping.get(export_mime_type, '')

def parse_date_input(date_str):
    """
    Parse a date string into RFC 3339 UTC format.
    Accepts absolute dates or relative times (e.g., "-1d").
    """
    if not date_str.startswith('-'):
        try:
            dt = date_parser.parse(date_str)
            if dt.tzinfo is None:
                dt = dt.astimezone().astimezone(tz=timedelta(0))
            else:
                dt = dt.astimezone(tz=timedelta(0))
            return dt.strftime("%Y-%m-%dT%H:%M:%SZ")
        except Exception as e:
            print(f"Error parsing date '{date_str}': {e}")
            sys.exit(1)
    else:
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

def find_folder_id(drive_path, starting_folder_id='root'):
    """
    Find the folder ID in Google Drive based on a relative path.

    Supports Drive for Desktop shortcut mounts and explicitly prefixed shared
    drives (e.g., 'Shared drives/Team/Folder'). The function first tries an
    exact match, then falls back to word-based search. If a segment cannot be
    resolved, it is skipped after logging the available options.
    """
    global folder_cache
    if not drive_path:
        return starting_folder_id

    cached_value = folder_cache.get(drive_path)
    if isinstance(cached_value, str):
        return cached_value

    cache_path = drive_path
    segments = [segment for segment in drive_path.split(os.path.sep) if segment]
    if not segments:
        folder_cache[cache_path] = starting_folder_id
        return starting_folder_id

    folder_id = starting_folder_id
    segments_to_process = list(segments)

    if segments[0] == '.shortcut-targets-by-id' and len(segments) >= 3:
        drive_id = segments[1]
        metadata = fetch_folder_metadata(drive_id)
        if metadata:
            online_title = metadata.get('title', segments[2])
        else:
            logging.warning('Failed to fetch online title for drive ID %s', drive_id)
            online_title = segments[2]
        segments_to_process = segments[2:]
        if segments_to_process:
            segments_to_process[0] = online_title
        cache_path = os.path.sep.join(segments_to_process)
        folder_id = drive_id
    elif segments[0].strip().lower() in SHARED_DRIVE_MARKERS:
        if len(segments) < 2:
            logging.error('Shared drive name missing in path: %s', drive_path)
            print(f"?? Shared drive name missing in path: {drive_path}")
            folder_cache[cache_path] = starting_folder_id
            return starting_folder_id
        shared_drive_name = segments[1].strip()
        shared_drive = get_shared_drive_metadata_by_name(shared_drive_name)
        if not shared_drive:
            logging.error("Shared drive '%s' not found.", shared_drive_name)
            print(f"?? Shared drive '{shared_drive_name}' not found.")
            folder_cache[cache_path] = starting_folder_id
            return starting_folder_id
        folder_id = shared_drive.get('id', starting_folder_id)
        segments_to_process = segments[2:]
    else:
        segments_to_process = segments

    def normalize(title):
        return re.sub(r"\s+", " ", title.replace(":", " ")).strip().lower()

    for folder_name in segments_to_process:
        if not folder_name:
            continue

        metadata = None
        current_title = ''
        if folder_id != 'root':
            metadata = fetch_folder_metadata(folder_id)
            if metadata:
                current_title = metadata.get('title', '')
        if metadata and normalize(current_title) == normalize(folder_name):
            logging.info("Current folder '%s' already matches '%s'; skipping search", current_title, folder_name)
            print(f"? Current folder '{current_title}' already matches '{folder_name}'; skipping search")
            continue

        logging.info('Looking for folder: %s under %s', folder_name, folder_id)
        print(f"?? Looking for folder: {folder_name} under {folder_id}")

        query = (f"'{folder_id}' in parents and trashed=false and "
                 f"mimeType='application/vnd.google-apps.folder' and title='{folder_name}'")
        try:
            file_list = list_drive_files({'q': query})
        except Exception as e:
            logging.error('Exact match query failed: %s', e)
            print(f"?? Exact match query failed: {e}")
            file_list = []

        if not file_list and (':' not in folder_name) and re.search(r"\s{2,}", folder_name):
            alt_name = re.sub(r"\s{2,}", ': ', folder_name)
            logging.info('Exact match retry with colon: %s', alt_name)
            print(f"?? Exact match retry with colon: {alt_name}")
            query = (f"'{folder_id}' in parents and trashed=false and "
                     f"mimeType='application/vnd.google-apps.folder' and title='{alt_name}'")
            try:
                file_list = list_drive_files({'q': query})
            except Exception as e:
                logging.error('Exact match alternative query failed: %s', e)
                print(f"?? Exact match alternative query failed: {e}")
                file_list = []

        if not file_list and any(char in folder_name for char in " :;,_-"):
            logging.info("Exact match failed for '%s', attempting word-based search", folder_name)
            print(f"?? Exact match failed for '{folder_name}', attempting word-based search")
            words = [word.strip() for word in re.split(r"[\s:;,_-]+", folder_name) if word.strip()]
            if not words:
                logging.error('No valid words in folder name: %s', folder_name)
                print(f"?? No valid words in folder name: {folder_name}")
                continue
            search_terms = ' '.join(words)
            query = (f"'{folder_id}' in parents and trashed=false and "
                     f"mimeType='application/vnd.google-apps.folder' and title contains '{search_terms}'")
            logging.info('Word-based query: %s', query)
            print(f"?? Word-based query: {query}")
            try:
                file_list = list_drive_files({'q': query})
                target_length = sum(len(word) for word in words)
                matches = []
                pattern = '.*'.join(re.escape(word) for word in words)
                for folder in file_list:
                    folder_title = folder['title']
                    folder_words = [w for w in re.split(r"[\s:;,_-]+", folder_title) if w]
                    folder_length = sum(len(w) for w in folder_words)
                    if folder_length == target_length and re.search(pattern, folder_title, re.IGNORECASE):
                        matches.append(folder)
                if matches:
                    if len(matches) > 1:
                        logging.warning("Multiple matches for '%s' under %s: %s. Using first match.",
                                        folder_name, folder_id, [f['title'] for f in matches])
                        print(f"?? Multiple matches for '{folder_name}' under {folder_id}: {[f['title'] for f in matches]}. Using first match.")
                    file_list = [matches[0]]
                else:
                    file_list = []
            except Exception as e:
                logging.error('Word-based search failed: %s', e)
                print(f"?? Word-based search failed: {e}")
                file_list = []

        if not file_list:
            all_folders_query = (f"'{folder_id}' in parents and trashed=false and "
                                 f"mimeType='application/vnd.google-apps.folder'")
            try:
                all_folders = list_drive_files({'q': all_folders_query})
                available = [f['title'] for f in all_folders]
                logging.error("No matching folder found for '%s' under %s. Available folders: %s",
                              folder_name, folder_id, available)
                print(f"?? No matching folder found for '{folder_name}' under {folder_id}. Available folders: {available}")
            except Exception as e:
                logging.error('Failed to list subfolders: %s', e)
                print(f"?? Failed to list subfolders: {e}")
            logging.warning("Skipping folder segment '%s'", folder_name)
            print(f"?? Skipping folder segment '{folder_name}'")
            continue

        folder_id = file_list[0]['id']
        logging.info("Found folder '%s', ID: %s", file_list[0]['title'], folder_id)
        print(f"? Found folder '{file_list[0]['title']}', ID: {folder_id}")

    folder_cache[cache_path] = folder_id
    return folder_id

def strip_duplicate_suffix(filename):
    """Remove the '(number)' suffix added by Windows for duplicates."""
    match = re.match(r"^(.*?)\s*\(\d+\)$", filename)
    if match:
        return match.group(1)
    return filename

def find_files_in_drive(drive, folder_id, base_path="", filename=None, depth=0, max_depth=float('inf'),
                        newer_than=None, older_than=None):
    """
    Recursively find Google Docs/Sheets/Slides files.
    Groups files by title and recurses into subfolders.
    """
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
    
    file_list = list_drive_files({'q': query})
    matching_files = [f for f in file_list if f['mimeType'] in mime_types]
    folders = [f for f in file_list if f['mimeType'] == 'application/vnd.google-apps.folder']
    files_by_title = {}
    for file in matching_files:
        title = file['title']
        files_by_title.setdefault(title, []).append(file)
    for title, files in files_by_title.items():
        files.sort(key=lambda x: x.get('createdDate', ''))
        for i, f in enumerate(files):
            suffix = "" if i == 0 else f" ({i})"
            files_by_title[title][i] = (f['id'], suffix, f['mimeType'], base_path, f)
            logging.info("File '%s%s', ID: %s, MIME: %s, Path: %s", title, suffix, f['id'], f['mimeType'], base_path)
            print(f"✅ File '{title}{suffix}', ID: {f['id']}, MIME: {f['mimeType']}, Path: {base_path}")
    if depth < max_depth:
        for folder in folders:
            subfolder_id = folder['id']
            subfolder_title = folder['title']
            subfolder_path = os.path.join(base_path, subfolder_title) if base_path else subfolder_title
            logging.info("Recursing into subfolder: %s (ID: %s) at depth %d", subfolder_title, subfolder_id, depth + 1)
            print(f"📁 Recursing into subfolder: {subfolder_title} (ID: {subfolder_id})")
            subfolder_files = find_files_in_drive(drive, subfolder_id, subfolder_path, filename,
                                                  depth + 1, max_depth, newer_than, older_than)
            for title, file_entries in subfolder_files.items():
                files_by_title.setdefault(title, []).extend(file_entries)
    total_files = sum(len(files) for files in files_by_title.values())
    logging.info("Found %d files in folder ID '%s' (depth %d).", total_files, folder_id, depth)
    print(f"✅ Found {total_files} files at depth {depth}.")
    return files_by_title

def get_file_drive_path(file_meta):
    """
    Given a file's metadata, compute a relative folder path representing its location in Drive.
    If the top-level folder is 'My Drive', it is removed; otherwise, the path is
    prefixed with either 'Shared drives' (when applicable) or 'Shared With Me'.
    Folder names are sanitized.
    """
    parents = file_meta.get('parents', [])
    if not parents:
        return ''
    parent_id = parents[0]['id'] if isinstance(parents[0], dict) else parents[0]
    path_parts = []
    root_drive_id = None

    while parent_id and parent_id != 'root':
        folder_info = folder_cache.get(parent_id)
        if not isinstance(folder_info, dict) or 'driveId' not in folder_info:
            folder_info = fetch_folder_metadata(parent_id)
            if not folder_info:
                break
        sanitized_title = sanitize_filename(folder_info['title'])
        path_parts.append((sanitized_title, folder_info.get('driveId')))
        parent_refs = folder_info.get('parents', [])
        if not parent_refs:
            root_drive_id = folder_info.get('driveId')
            break
        parent_ref = parent_refs[0]
        if isinstance(parent_ref, dict):
            if parent_ref.get('isRoot'):
                root_drive_id = folder_info.get('driveId', root_drive_id)
            parent_id = parent_ref.get('id')
        else:
            parent_id = parent_ref

    reversed_parts = list(reversed(path_parts))
    titles_only = [title for title, _ in reversed_parts]

    if titles_only:
        if titles_only[0] == 'My Drive':
            titles_only = titles_only[1:]
        else:
            drive_id = root_drive_id or (reversed_parts[0][1] if reversed_parts else None)
            if drive_id:
                shared_drive_name = get_shared_drive_name(drive_id) or titles_only[0]
                if shared_drive_name:
                    titles_only[0] = sanitize_filename(shared_drive_name)
                titles_only.insert(0, 'Shared drives')
            else:
                titles_only.insert(0, 'Shared With Me')

    if not titles_only:
        return ''
    return os.path.join(*titles_only)

def export_google_file_via_v3(file_id, export_mimetype, destination_path):
    """Export a Google-native file using the Drive v3 API with shared-drive support."""
    if drive_v3 is None:
        raise RuntimeError('Drive v3 service is not initialized')
    request = drive_v3.files().export_media(
        fileId=file_id,
        mimeType=export_mimetype
    )
    if 'supportsAllDrives' not in request.uri:
        separator = '&' if '?' in request.uri else '?'
        request.uri = f"{request.uri}{separator}supportsAllDrives=true&supportsTeamDrives=true"
    with io.FileIO(add_long_path_prefix(destination_path), 'wb') as fh:
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()
            # status can be None for small files; no progress logging needed.
    return destination_path

def download_google_file_as_ms_office(file_id, suffix, mime_type, subfolder_path, add_timestamp,
                                      output_directory, dry_run=False, file_meta=None,
                                      prune_newest=None, prune_staggered=None, no_clobber=False):
    """
    Download a Google file as an MS Office document, preserving the subfolder structure.

    Behavior:
      - If --no-clobber is set and the plain target exists, skip downloading.
      - Otherwise, if --timestamp is not specified and the plain target exists, rename it into a backup.
      - Then download the new file. If --timestamp is specified, its name includes the Google file's modified timestamp.
      - If --staggered or --newest are specified, existing backups are pruned.
      - Finally, update the downloaded file's modification time to match that on Google.
      - Also, skip downloading if a file already exists that shares the same modified time.
    """
    meta = file_meta if file_meta else drive.CreateFile({'id': file_id})
    meta['supportsAllDrives'] = True
    meta['supportsTeamDrives'] = True
    if file_meta is None:
        try:
            meta.FetchMetadata(fields='id,title,mimeType,modifiedDate,parents,driveId')
        except Exception as e:
            logging.error('Failed to fetch metadata for %s: %s', file_id, e)
            print(f"?? Failed to fetch metadata for {file_id}: {e}")
            return
    logging.info("Retrieved file: %s (ID: %s, MIME: %s)", meta['title'], file_id, meta['mimeType'])
    print(f"📄 Found File: {meta['title']} (ID: {file_id})")
    print(f"🔍 MIME Type: {meta['mimeType']}")

    export_mimetype = get_export_mimetype(meta['mimeType'])
    if export_mimetype is None:
        logging.error("Unsupported MIME type: %s", meta['mimeType'])
        print("⚠️  Unsupported file type.")
        return

    orig_ext = get_original_extension(mime_type)  # e.g. "gsheet"
    file_extension = get_export_extension(export_mimetype)  # e.g. "xlsx"
    full_output_directory = os.path.join(output_directory, subfolder_path)
    try:
        if not dry_run:
            os.makedirs(add_long_path_prefix(full_output_directory), exist_ok=True)
    except Exception as e:
        logging.error("Failed to create directory %s: %s", full_output_directory, e)
        print(f"⚠️  Failed to create directory {full_output_directory}: {e}")
        return
    base_name = f"{sanitize_filename(meta['title'])}{suffix}{'.' + orig_ext}"
    plain_target = os.path.join(full_output_directory, f"{base_name}.{file_extension}")

    # Determine timestamp from the Google file's modifiedDate.
    google_mod_str = meta.get("modifiedDate")
    if google_mod_str:
        try:
            mod_dt = date_parser.parse(google_mod_str)
            timestamp_str = mod_dt.strftime("%Y-%m-%d_%H-%M-%S")
        except Exception as e:
            logging.warning("Failed to parse modifiedDate: %s", e)
            timestamp_str = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    else:
        timestamp_str = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

    if add_timestamp:
        new_file_name = f"{base_name}_{timestamp_str}.{file_extension}"
    else:
        new_file_name = f"{base_name}.{file_extension}"
    new_target = os.path.join(full_output_directory, new_file_name)

    # If not using timestamp, check if the existing file has the same modified time as on Google.
    if not add_timestamp:
        target_path = add_long_path_prefix(plain_target)
        if os.path.exists(target_path):
            try:
                local_mod_ts = os.path.getmtime(target_path)
                if google_mod_str:
                    google_mod_ts = date_parser.parse(google_mod_str).timestamp()
                    # Allow for minor differences (within 1 second)
                    if abs(local_mod_ts - google_mod_ts) < 1.0:
                        logging.info("Skipping download; %s exists and is up-to-date.", plain_target)
                        print(f"⏩ Skipping download: {plain_target} already exists with matching modification time.")
                        return
            except Exception as e:
                logging.warning("Unable to compare modification times for %s: %s", plain_target, e)

    if no_clobber and os.path.exists(add_long_path_prefix(plain_target)):
        logging.info("Skipping download; %s exists (--no-clobber enabled)", plain_target)
        print(f"⏩ Skipping download: {plain_target} exists")
        return

    # Backup existing file (using Google modified time for the backup name) if not using timestamp.
    if not add_timestamp and os.path.exists(add_long_path_prefix(plain_target)):
        try:
            backup_timestamp = timestamp_str  # Use the Google modified timestamp
            backup_name = f"{base_name}_{backup_timestamp}.{file_extension}"
            backup_path = os.path.join(full_output_directory, backup_name)
            if not dry_run:
                os.rename(add_long_path_prefix(plain_target), add_long_path_prefix(backup_path))
            logging.info("Renamed %s to backup %s", plain_target, backup_path)
            print(f"🗑 Renamed existing file to backup: {backup_path}")
        except Exception as e:
            logging.error("Error renaming file %s to backup: %s", plain_target, e)
            print(f"⚠️ Error renaming file {plain_target} to backup: {e}")

    logging.info("%s to: %s", "[Dry Run] Would download" if dry_run else "Downloading", new_target)
    print(f"{'[Dry Run] Would save' if dry_run else '⬇️  Saving'} as: {new_target}")
    try:
        if not dry_run:
            try:
                export_google_file_via_v3(file_id, export_mimetype, new_target)
            except RuntimeError:
                file_to_download = drive.CreateFile({'id': file_id})
                file_to_download['supportsAllDrives'] = True
                file_to_download['supportsTeamDrives'] = True
                file_to_download.GetContentFile(add_long_path_prefix(new_target), mimetype=export_mimetype)
            logging.info("Download successful: %s", new_target)
            print(f"✅ File downloaded successfully: {new_target}")
            # Update local file modification time to match Google's modifiedDate.
            if google_mod_str:
                try:
                    mod_dt = date_parser.parse(google_mod_str)
                    mod_ts = mod_dt.timestamp()
                    os.utime(add_long_path_prefix(new_target), (mod_ts, mod_ts))
                    logging.info("Updated modification time for %s to %s", new_target, mod_dt.isoformat())
                    print(f"⏰ Updated file modification time for {new_target}")
                except Exception as e:
                    logging.warning("Failed to update modification time for %s: %s", new_target, e)
    except Exception as e:
        logging.error("Error downloading file to %s: %s", new_target, e)
        print(f"⚠️ Error downloading file to {new_target}: {e}")
        return

    if prune_newest is not None or prune_staggered is not None:
        pattern = os.path.join(full_output_directory, f"{base_name}_*.{file_extension}")
        existing_backups = glob.glob(pattern)
        ts_regex = re.compile(rf"^{re.escape(base_name)}_(\d{{4}}-\d{{2}}-\d{{2}}_\d{{2}}-\d{{2}}-\d{{2}})\.{re.escape(file_extension)}$")
        backups_with_ts = []
        for f in existing_backups:
            m = ts_regex.match(os.path.basename(f))
            if m:
                ts_str = m.group(1)
                try:
                    ts = datetime.strptime(ts_str, "%Y-%m-%d_%H-%M-%S")
                    backups_with_ts.append((f, ts))
                except Exception as e:
                    logging.warning("Failed to parse timestamp from %s: %s", f, e)
        if prune_newest is not None:
            backups_with_ts.sort(key=lambda x: x[1], reverse=True)
            if len(backups_with_ts) > prune_newest:
                for f, ts in backups_with_ts[prune_newest:]:
                    try:
                        if not dry_run:
                            os.remove(add_long_path_prefix(f))
                        logging.info("Pruned backup (newest mode): %s", f)
                        print(f"🗑 Pruned backup (newest): {f}")
                    except Exception as e:
                        logging.error("Error pruning backup %s: %s", f, e)
                        print(f"⚠️ Error pruning backup {f}: {e}")
        elif prune_staggered is not None:
            backups_with_ts.sort(key=lambda x: x[1])
            m_val = len(backups_with_ts)
            n_val = prune_staggered
            if m_val > n_val:
                if n_val == 1:
                    keep_indices = [m_val - 1]
                else:
                    keep_indices = []
                    for i in range(n_val):
                        if i == 0:
                            keep_indices.append(0)
                        elif i == n_val - 1:
                            keep_indices.append(m_val - 1)
                        else:
                            idx = int(round(math.exp((math.log(m_val - 1)) * i / (n_val - 1))))
                            idx = min(max(idx, 0), m_val - 1)
                            keep_indices.append(idx)
                    keep_indices = sorted(set(keep_indices))
                files_to_keep = {backups_with_ts[i][0] for i in keep_indices}
                for f, ts in backups_with_ts:
                    if f not in files_to_keep:
                        try:
                            if not dry_run:
                                os.remove(add_long_path_prefix(f))
                            logging.info("Pruned backup (staggered mode): %s", f)
                            print(f"🗑 Pruned backup (staggered): {f}")
                        except Exception as e:
                            logging.error("Error pruning backup %s: %s", f, e)
                            print(f"⚠️ Error pruning backup {f}: {e}")


def process_global_search(add_timestamp, backup_path, dry_run, max_depth, newer_than, older_than, title_filter, prune_newest, prune_staggered, no_clobber):
    """
    Perform a global search on Drive using the given filters.
    Each file is downloaded into a subfolder (computed from its Drive path) under the backup directory.
    """
    mime_types = (
        "mimeType = 'application/vnd.google-apps.document' or "
        "mimeType = 'application/vnd.google-apps.spreadsheet' or "
        "mimeType = 'application/vnd.google-apps.presentation'"
    )
    query = f"trashed=false and ({mime_types})"
    if title_filter:
        query += f" and title contains '{title_filter}'"
    if newer_than:
        query += f" and modifiedDate > '{newer_than}'"
    if older_than:
        query += f" and modifiedDate < '{older_than}'"

    file_list = list_drive_files({'q': query})
    output_directory = os.path.abspath(backup_path) if backup_path else os.getcwd()
    try:
        if not dry_run:
            os.makedirs(add_long_path_prefix(output_directory), exist_ok=True)
    except Exception as e:
        logging.error("Failed to create directory %s: %s", output_directory, e)
        print(f"⚠️ Failed to create directory {output_directory}: {e}")
        return
    if backup_path:
        logging.info("Redirecting download to: %s", output_directory)
        print(f"📁 Redirecting download to: {output_directory}")
    for file in file_list:
        subfolder_path = get_file_drive_path(file)
        suffix = ""
        download_google_file_as_ms_office(
            file_id=file['id'],
            suffix=suffix,
            mime_type=file['mimeType'],
            subfolder_path=subfolder_path,
            add_timestamp=add_timestamp,
            output_directory=output_directory,
            dry_run=dry_run,
            file_meta=file,
            prune_newest=prune_newest,
            prune_staggered=prune_staggered,
            no_clobber=no_clobber
        )

def process_path(input_path, add_timestamp, backup_path, dry_run, max_depth, newer_than, older_than, prune_newest, prune_staggered, no_clobber):
    """Process a local input path (file or directory)."""
    input_path = os.path.abspath(input_path)
    mime_type_map = {
        '.gdoc': 'application/vnd.google-apps.document',
        '.gsheet': 'application/vnd.google-apps.spreadsheet',
        '.gslides': 'application/vnd.google-apps.presentation',
    }
    try:
        if os.path.isdir(input_path):
            logging.info("Input is a directory: %s", input_path)
            print(f"📁 Processing directory: {input_path}")
            drive_path = normalize_local_drive_path(input_path)
            folder_id = find_folder_id(drive_path)
            files_by_title = find_files_in_drive(drive, folder_id, max_depth=max_depth,
                                                 newer_than=newer_than, older_than=older_than)
            output_directory = os.path.abspath(backup_path) if backup_path else input_path
            try:
                if not dry_run:
                    os.makedirs(add_long_path_prefix(output_directory), exist_ok=True)
            except Exception as e:
                logging.error("Failed to create directory %s: %s", output_directory, e)
                print(f"⚠️ Failed to create directory {output_directory}: {e}")
                return
            if backup_path:
                logging.info("Redirecting download to: %s", output_directory)
                print(f"📁 Redirecting download to: {output_directory}")
            for title, file_entries in files_by_title.items():
                for file_id, suffix, mime_type, subfolder_path, file_meta in file_entries:
                    download_google_file_as_ms_office(file_id, suffix, mime_type, subfolder_path, add_timestamp,
                                                      output_directory, dry_run, file_meta,
                                                      prune_newest, prune_staggered, no_clobber)
        else:
            if not os.path.exists(input_path):
                logging.error("File does not exist: %s", input_path)
                print(f"⚠️ Error: {input_path} does not exist")
                return
            logging.info("Processing file: %s", input_path)
            print(f"Processing: {input_path}")
            folder_path, full_filename = os.path.split(input_path)
            filename, file_extension = os.path.splitext(full_filename)
            if file_extension not in mime_type_map:
                logging.error("Unsupported file type: %s", file_extension)
                print(f"⚠️ Unsupported file type: {file_extension}")
                return
            drive_path = normalize_local_drive_path(folder_path)
            folder_id = find_folder_id(drive_path)
            files_by_title = find_files_in_drive(drive, folder_id, filename=filename,
                                                 max_depth=max_depth, newer_than=newer_than, older_than=older_than)
            output_directory = os.path.abspath(backup_path) if backup_path else folder_path
            try:
                if not dry_run:
                    os.makedirs(add_long_path_prefix(output_directory), exist_ok=True)
            except Exception as e:
                logging.error("Failed to create directory %s: %s", output_directory, e)
                print(f"⚠️ Failed to create directory {output_directory}: {e}")
                return
            if backup_path:
                logging.info("Redirecting download to: %s", output_directory)
                print(f"📁 Redirecting download to: {output_directory}")
            else:
                logging.warning("Backup path not matched; using default location.")
                print("⚠️ Backup path not matched; using default location.")
            base_filename = strip_duplicate_suffix(filename)
            if base_filename in files_by_title:
                for file_id, suffix, mime_type, subfolder_path, file_meta in files_by_title[base_filename]:
                    download_google_file_as_ms_office(file_id, suffix, mime_type, subfolder_path, add_timestamp,
                                                      output_directory, dry_run, file_meta,
                                                      prune_newest, prune_staggered, no_clobber)
            else:
                logging.error("File '%s' not found in Drive.", base_filename)
                print(f"⚠️ File '{base_filename}' not found in Drive.")
    except Exception as e:
        logging.exception("Processing failed for %s", input_path)
        print(f"⚠️ Processing failed for {input_path}: {e}")

def get_clasp_command():
    if os.name == 'nt':
        # Try to locate clasp.cmd first.
        cmd_path = shutil.which("clasp.cmd")
        if cmd_path:
            return ["clasp.cmd"]
        # If not found, try clasp.ps1.
        ps1_path = shutil.which("clasp.ps1")
        if ps1_path:
            # Use PowerShell to execute the script.
            return ["powershell", "-ExecutionPolicy", "Bypass", "-File", ps1_path]
        return None
    else:
        cmd_path = shutil.which("clasp")
        return [cmd_path] if cmd_path else None

def backup_standalone_scripts(backup_directory, dry_run):
    clasp_cmd_list = get_clasp_command()
    if clasp_cmd_list is None:
        print("Error: CLASP command not found.")
        if os.name == "nt":
            print("On Windows, please install Node.js and then run: npm install -g @google/clasp")
        else:
            print("On Unix-like systems, please install Node.js (which includes npm) and then run: npm install -g @google/clasp")
        sys.exit(1)

    # Create a common backup directory for all standalone script projects.
    script_backup_dir = os.path.join(backup_directory, "AppScript")
    try:
        if not dry_run:
            os.makedirs(script_backup_dir, exist_ok=True)
    except Exception as e:
        print(f"⚠️ Failed to create script backup directory {script_backup_dir}: {e}")
        return

    query = "mimeType='application/vnd.google-apps.script' and trashed=false"
    try:
        script_files = list_drive_files({'q': query})
    except Exception as e:
        print(f"⚠️ Error querying standalone script projects: {e}")
        return

    if not script_files:
        print("No standalone script projects found.")
        return

    print(f"📄 Found {len(script_files)} standalone script project(s).")
    for script_file in script_files:
        project_id = script_file['id']
        project_title = sanitize_filename(script_file['title'])
        # Create a folder for each project.
        project_backup_dir = os.path.join(script_backup_dir, project_title)
        print(f"📁 Backing up standalone script project: {project_title} (ID: {project_id})")

        if dry_run:
            print(f"⏱ [Dry Run] Would run: { ' '.join(clasp_cmd_list + ['clone', project_id]) } or pull if folder exists")
            continue

        # Create the subfolder if it doesn't exist.
        os.makedirs(project_backup_dir, exist_ok=True)
        clasp_config_path = os.path.join(project_backup_dir, ".clasp.json")
        if os.path.exists(clasp_config_path):
            # Folder exists and contains a CLASP configuration file → perform pull.
            print(f"🔄 Folder exists. Running clasp pull in {project_backup_dir}")
            try:
                subprocess.run(clasp_cmd_list + ["pull"], cwd=project_backup_dir, check=True)
                print(f"✅ Successfully updated {project_title} in {project_backup_dir}")
            except subprocess.CalledProcessError as e:
                print(f"⚠️ CLASP pull failed for project {project_title}: {e}")
        else:
            # Folder does not have a CLASP configuration file → perform clone.
            print(f"⬇️  Folder missing configuration. Running clasp clone in {project_backup_dir}")
            try:
                subprocess.run(clasp_cmd_list + ["clone", project_id], cwd=project_backup_dir, check=True)
                print(f"✅ Successfully cloned {project_title} into {project_backup_dir}")
            except subprocess.CalledProcessError as e:
                print(f"⚠️ CLASP clone failed for project {project_title}: {e}")

def parse_arguments():
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(description="Download Google Docs/Sheets/Slides as Office files.")
    parser.add_argument("paths", nargs="*", help="Local file/directory paths (wildcards supported). Leave empty for global Drive search.")
    parser.add_argument("--timestamp", action="store_true", help="Append timestamp (from Google modified date) to new file name.")
    parser.add_argument("--backup", type=str, help="Redirect downloads to backup path.")
    parser.add_argument("--dry-run", action="store_true", help="Simulate download without writing files.")
    parser.add_argument("--max-depth", type=int, default=float('inf'), help="Max recursion depth for folder search.")
    parser.add_argument("--newer-than", type=str, help="Filter files newer than a date (absolute or relative, e.g. '-1d').")
    parser.add_argument("--older-than", type=str, help="Filter files older than a date (absolute or relative).")
    parser.add_argument("--title", type=str, help="Search for files with a specific title (partial match).")
    parser.add_argument("--staggered", type=int, help="Retain up to <n> timestamped backups in staggered mode.")
    parser.add_argument("--newest", type=int, help="Retain up to <n> most recent timestamped backups.")
    parser.add_argument("--no-clobber", action="store_true", help="Do not re-download if the target exists.")
    parser.add_argument("--no-scripts", action="store_true", help="Do not back up standalone Apps Script projects.")
    return parser.parse_args()


def main():
    print("Script started!")
    args = parse_arguments()
    # Process the --newer-than argument:
    if args.newer_than:
        if args.newer_than.lower() == "last run":
            if os.path.exists(log_file):
                mod_time = os.path.getmtime(log_file)
                dt = datetime.fromtimestamp(mod_time)
                args.newer_than = dt.strftime("%Y-%m-%dT%H:%M:%SZ")
                print(f"⏱ Using log file modified time as --newer-than: {args.newer_than}")
            else:
                print("⚠️ Log file not found; ignoring 'last run'")
                args.newer_than = None
        else:
            args.newer_than = parse_date_input(args.newer_than)
    if args.older_than:
        args.older_than = parse_date_input(args.older_than)
    prune_newest = args.newest if args.newest is not None else None
    prune_staggered = None if prune_newest is not None else args.staggered

    if not args.paths:
        print("No input paths provided. Performing a global search on your Drive.")
        process_global_search(args.timestamp, args.backup, args.dry_run, args.max_depth,
                              args.newer_than, args.older_than, args.title, prune_newest, prune_staggered, args.no_clobber)
    else:
        input_paths = []
        for path in args.paths:
            expanded = glob.glob(path.strip('"'))
            if expanded:
                input_paths.extend(expanded)
            else:
                input_paths.append(path.strip('"'))
        for input_path in input_paths:
            process_path(input_path, args.timestamp, args.backup, args.dry_run, args.max_depth,
                         args.newer_than, args.older_than, prune_newest, prune_staggered, args.no_clobber)
    # Now, if not suppressed, back up standalone Apps Script projects.
    if not args.no_scripts:
        backup_dir = os.path.abspath(args.backup) if args.backup else os.getcwd()
        print("🔄 Attempting to back up standalone Apps Script projects...")
        backup_standalone_scripts(backup_dir, args.dry_run)
    else:
        print("ℹ️ Standalone script backup suppressed (--no-scripts provided).")
    print("✅ Script completed!")


if __name__ == "__main__":
    main()
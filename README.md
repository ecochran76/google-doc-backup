Below is an updated README.md that includes a section with instructions on how to obtain your `client_secrets.json` file:

---

```markdown
# google_doc_backup

**google_doc_backup** is a command-line tool that backs up your Google Docs, Sheets, and Slides to Microsoft Office formats (DOCX, XLSX, and PPTX). It uses PyDrive to authenticate with Google Drive, supports global searches based on date and title filters, and organizes files in your backup directory by their Google Drive folder hierarchy. It also includes optional version control features such as timestamping, backup pruning, and a no‑clobber mode to prevent re-downloading of existing files.

## Features

- **File Conversion:**  
  Download Google Docs as DOCX, Google Sheets as XLSX, and Google Slides as PPTX.

- **Backup Organization:**  
  Preserve your Google Drive folder structure when saving files locally.

- **Version Control Options:**  
  - Append a timestamp to downloaded files (`--timestamp`).
  - Automatically rename existing files to backups using their OS modification time.
  - Retain a specified number of backups using either a staggered or newest-pruning strategy (`--staggered` or `--newest`).
  - Optionally prevent re-downloading of files if they already exist (`--no-clobber`).

- **Global Search:**  
  Perform searches across your entire Drive based on modification dates (`--newer-than` / `--older-than`) and title (`--title`).

- **Long Path Support (Windows):**  
  Handles file paths longer than 260 characters by automatically prepending the Windows extended-length prefix.

## Installation

1. **Clone the Repository:**

   ```bash
   git clone https://github.com/yourusername/google_doc_backup.git
   cd google_doc_backup
   ```

2. **Install in Editable Mode:**

   ```bash
   pip install -e .
   ```

   Alternatively, if published to PyPI:

   ```bash
   pip install google_doc_backup
   ```

3. **Install Dependencies:**

   Ensure that the required packages are installed (listed in `requirements.txt`):

   ```bash
   pip install pydrive python-dateutil
   ```

## Obtaining client_secrets.json

This tool uses the Google Drive API for accessing your files. To authenticate, you need a `client_secrets.json` file:

1. Visit the [Google Developers Console](https://console.developers.google.com/).
2. Create a new project (or select an existing one).
3. Enable the **Google Drive API** for your project.
4. Go to **Credentials** and click on **Create Credentials** > **OAuth client ID**.
5. Configure the consent screen if prompted.
6. Choose **Desktop app** as the application type.
7. Download the generated `client_secrets.json` file.
8. Place the `client_secrets.json` file in the same directory as the script (or in the directory specified by your package).

## Usage

After installation, you can use the package via the command-line interface:

```bash
google-doc-backup [options] [paths...]
```

### Command-Line Options

- `paths`:  
  Local file or directory paths (wildcards supported). If omitted, a global Drive search is performed.

- `--timestamp`:  
  Append the current timestamp to the new file name.

- `--backup <backup_path>`:  
  Specify a destination directory for backups. Files will be organized into subfolders matching their Drive folder hierarchy.

- `--no-clobber`:  
  Do not re-download a file if it already exists.

- `--max-depth <n>`:  
  Maximum recursion depth for folder search (default: unlimited).

- `--newer-than <date>`:  
  Filter files newer than the specified date. Supports absolute dates (e.g., `2023-01-01T00:00:00`) or relative times (e.g., `-1d`).

- `--older-than <date>`:  
  Filter files older than the specified date.

- `--title <title>`:  
  Search for files with a specific title (partial match).

- `--staggered <n>`:  
  Retain up to _n_ timestamped backups in staggered mode.

- `--newest <n>`:  
  Retain up to _n_ most recent timestamped backups.

### Examples

- **Global Search with Versioning:**  
  Back up files with titles containing "Workouts" modified within the last 30 days. Save backups to `E:\SyncThing\Cloud\Google` using staggered backup retention (up to 5 backups):

  ```bash
  google-doc-backup --newer-than="-30d" --backup "E:\SyncThing\Cloud\Google" --staggered=5 --title "Workouts"
  ```

- **Prevent Overwrites:**  
  Skip downloading if the target file already exists:

  ```bash
  google-doc-backup --no-clobber --backup "E:\SyncThing\Cloud\Google" --title "Report"
  ```

- **Local Directory Processing:**  
  Back up all files in a local folder (which corresponds to a folder in "My Drive"):

  ```bash
  google-doc-backup "H:\My Drive\Documents\Reports"
  ```

## Authentication

The tool uses PyDrive for authentication. Place your `client_secrets.json` file in the same directory as the script, and the tool will handle credential management automatically. For first-time use, a browser window will open to authorize access to your Google Drive.

## Contributing

Contributions are welcome! Please open issues or submit pull requests on the [GitHub repository](https://github.com/yourusername/google_doc_backup).

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Disclaimer

This tool relies on the Google Drive API, which is subject to change. Use at your own risk and ensure you comply with Google’s API usage policies.

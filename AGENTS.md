# Repository Guidelines

## Project Structure & Module Organization
The root directory holds the operational scripts. `download_google_doc.pyw` is the main entry point; it handles authentication, Drive traversal, exports, and Apps Script syncing. Supplementary utilities (`extract_google_file_id.py`, `gdoc_download_url.py`, `download_gdocs_generate_explorer_extensions.py`) cover ID parsing, URL generation, and Windows Explorer integrations. Generated OAuth artifacts (`client_secrets.json`, `credentials.json`) must stay local; never push real credentials. Logs land in `download_google_doc.log`. No dedicated package directory exists; if you add one, update `setup.py` accordingly.

## Build, Test, and Development Commands
Create a virtual environment and install runtime dependencies with `python -m venv .venv`, `.venv\Scripts\activate`, and `pip install -r requirements.txt`. For editable work, run `pip install -e .` so the CLIs resolve correctly. Use `python download_google_doc.pyw --help` to review options and `python download_google_doc.pyw --dry-run --backup "D:\Backups\Drive"` for safe validation. The Apps Script backup path depends on `clasp`; verify setup inside a project folder with `npx clasp status`.

## Coding Style & Naming Conventions
Follow PEP 8: four-space indentation, `snake_case` for functions and variables, and uppercase constants for paths (e.g., `CLIENT_SECRETS_PATH`). Keep functions focused; prefer extracting helpers instead of extending the main script. Favor structured `logging` calls over prints, and document any Windows-only behavior in inline comments.

## Testing Guidelines
There is no automated suite yet. When contributing logic, add focused `pytest` modules under `tests/` and stub Drive requests so tests run offline. Always exercise relevant command combinations with `--dry-run` before running destructive syncs, and note the commands in your PR. For Apps Script flows, confirm `--no-scripts` and full backups behave as expected.

## Commit & Pull Request Guidelines
Recent commits mix imperative subjects and Conventional Commit prefixes such as `feat:`. Keep subjects under 72 characters, present tense, and add a scope when it clarifies impact (`feat(clasp): tighten status checks`). Reference manual test logs or screenshots for Explorer menu changes, and link issues when available. PRs should describe risk areas, mention credential handling, and call out any follow-up work.

## Security & Configuration Tips
Store OAuth secrets outside version control and sanitize logs before sharing. Regenerate `credentials.json` after permission updates and prune stale tokens if authentication loops appear. Avoid embedding personal paths in committed scripts; prefer config files or example placeholders.

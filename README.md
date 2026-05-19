# google-doc-backup

`google-doc-backup` exports Google Docs, Sheets, and Slides into local Office-format backups. The current service path is tenant-scoped, `uv`-managed, and prefers the local `gws` Google Workspace CLI with the legacy direct Google API path kept as a fallback.

The default local backup target is:

```text
/mnt/e/SyncThing/Cloud/Google-Docs
```

## What It Backs Up

- My Drive, under `My Drive/` when `my_drive_root_mode = "scoped"`.
- Files shared with the authenticated user, under `Shared With Me/<owner>/`.
- Shared Drives visible to the authenticated user, under `Shared drives/<drive name>/`.
- Standalone Apps Script projects under `AppScript/` when `clasp` is installed and scripts are enabled.

Google-native files are exported as:

- Docs: `.docx`
- Sheets: `.xlsx`
- Slides: `.pptx`

Each run writes a JSON manifest under the tenant state directory so service runs can be audited without reading the backup tree directly.

## Install

Use `uv` for the Python environment:

```bash
uv sync
uv run python download_google_doc.pyw --help
```

The repo also exposes the historical script entry point directly. Prefer `uv run python download_google_doc.pyw ...` until the CLI is split into importable modules.

## Tenant Configuration

Tenant config lives outside the repo:

```text
~/.config/google-doc-backup/tenants/<tenant>.toml
```

Default state lives outside the repo:

```text
~/.local/state/google-doc-backup/tenants/<tenant>/
```

Example default tenant config:

```toml
tenant = "default"
backend = "auto"
backup_root = "/mnt/e/SyncThing/Cloud/Google-Docs"
include_my_drive = true
include_shared_with_me = true
include_shared_drives = true
include_apps_script = true
my_drive_root_mode = "scoped"
staggered = 5
dry_run = false
```

Backend modes:

- `auto`: use `gws` first, then fall back to the direct API path when the `gws` readiness probe fails.
- `gws`: require `gws`.
- `direct-api`: use the legacy PyDrive/direct Google API path.

Inspect resolved config without authenticating:

```bash
uv run python download_google_doc.pyw --tenant default --show-config
```

## Operations

Run readiness checks:

```bash
uv run python download_google_doc.pyw --tenant default --doctor
```

Run a no-write backup plan:

```bash
uv run python download_google_doc.pyw --tenant default --dry-run --show-config
```

Run a backup manually:

```bash
uv run python download_google_doc.pyw --tenant default
```

Disable Apps Script backup when `clasp` is unavailable:

```bash
uv run python download_google_doc.pyw --tenant default --no-scripts
```

Recover local `gws` auth with:

```bash
gws auth sync-gog
```

If the shell wrapper cannot find `gog`, set `GWS_GOG_BIN` to the real binary path and retry.

## User Service

Install or update the user-scoped systemd timer:

```bash
uv run python download_google_doc.pyw --tenant default --install-user-service
```

Preview generated service files without writing them:

```bash
uv run python download_google_doc.pyw --tenant default --install-user-service --service-dry-run
```

The default timer runs daily at `03:30` local time. On this workstation the installed service currently includes `--no-scripts` because `clasp` is optional and was not available during setup.

Inspect the timer:

```bash
systemctl --user list-timers 'google-doc-backup@default.timer' --no-pager
```

## Retention

Timestamped backup retention can be configured per tenant or overridden per run:

```toml
staggered = 5
```

`staggered = 5` keeps five staggered historical exports per file using the existing retention algorithm. `newest = N` keeps the newest `N` timestamped backups instead; if both are configured, `newest` wins.

CLI overrides:

```bash
uv run python download_google_doc.pyw --tenant default --staggered 5
uv run python download_google_doc.pyw --tenant default --newest 5
```

## Legacy Root Migration

Older backups may have My Drive folders directly at the backup root. Plan a migration into the scoped `My Drive/` root without moving files:

```bash
uv run python download_google_doc.pyw --tenant default --plan-migrate-my-drive-root
```

Review JSON output:

```bash
uv run python download_google_doc.pyw --tenant default --plan-migrate-my-drive-root --migration-plan-format json
```

Apply only after reviewing a zero-collision plan:

```bash
uv run python download_google_doc.pyw --tenant default --apply-migrate-my-drive-root
```

The apply command aborts before moving anything if the backup root is missing or any target path already exists.

## Direct API Fallback

The direct API path still uses local OAuth artifacts:

- `client_secrets.json`
- `credentials.json`

Keep real credentials out of version control. Regenerate `credentials.json` after permission changes, and prune stale tokens if authentication loops appear.

## Development

Run the narrow validation checks:

```bash
uv lock --check
uv run python -m py_compile download_google_doc.pyw
uv run python download_google_doc.pyw --tenant default --doctor
uv run python download_google_doc.pyw --tenant default --dry-run --show-config
```

When adding logic, prefer focused offline `pytest` coverage under `tests/` with fake Drive or backend responses.

## Windows Console Encoding

On Windows consoles that default to cp1252, set `PYTHONIOENCODING=utf-8` before launching the CLI so Unicode log messages do not trigger `UnicodeEncodeError`:

```cmd
set PYTHONIOENCODING=utf-8 && python download_google_doc.pyw --backup "E:\Backups\Drive"
```

# Google Docs Backup Configuration

Tenant-specific runtime configuration lives outside the repo.

Default tenant config path:

```text
~/.config/google-doc-backup/tenants/default.toml
```

Default tenant state path:

```text
~/.local/state/google-doc-backup/tenants/default/
```

Example tenant config:

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

For older backup trees with root-level My Drive files, use the migration planner before switching `my_drive_root_mode` from `legacy` to `scoped`.

Backend values:

- `auto`: prefer `gws`; fall back to the direct API path when the `gws` readiness probe fails.
- `gws`: require the Google Workspace CLI backend.
- `direct-api`: use the legacy PyDrive/direct Drive API path.

My Drive root modes:

- `legacy`: preserve the current backup layout by writing My Drive folders directly under `backup_root`.
- `scoped`: write My Drive content under `backup_root/My Drive/`.

Shared surfaces always use explicit roots:

- `Shared With Me/<owner>/...`
- `Shared drives/<drive name>/...`

Backup retention:

- `staggered = 5` keeps up to five timestamped backups per exported file using the existing staggered pruning algorithm.
- `newest = N` can be used instead to keep the newest `N` timestamped backups. If both are set, `newest` wins.
- CLI flags `--staggered` and `--newest` override tenant config for that run.

For local `gws` auth recovery, prefer:

```bash
gws auth sync-gog
```

If `gog` is not resolved correctly by the shell wrapper, set `GWS_GOG_BIN` to the real binary path before retrying.

Inspect the resolved config without authenticating:

```bash
uv run python download_google_doc.pyw --tenant default --show-config
```

Check service readiness:

```bash
uv run python download_google_doc.pyw --tenant default --doctor
```

Each backup or dry-run backup writes a JSON manifest under the tenant state directory:

```text
~/.local/state/google-doc-backup/tenants/default/runs/
~/.local/state/google-doc-backup/tenants/default/latest-run.json
```

Install and enable the default user systemd timer:

```bash
uv run python download_google_doc.pyw --tenant default --install-user-service
```

The default timer runs daily at `03:30` local time and currently invokes the backup with `--no-scripts` because `clasp` is optional and may not be installed.

The installed default service command is expected to look like:

```bash
uv run python /home/ecochran76/workspace.local/google-doc-backup/download_google_doc.pyw --tenant default --backend auto --no-scripts
```

Preview the service and timer files without writing them:

```bash
uv run python download_google_doc.pyw --tenant default --install-user-service --service-dry-run
```

Plan a legacy root migration without moving files:

```bash
uv run python download_google_doc.pyw --tenant default --plan-migrate-my-drive-root
```

The planner excludes already-scoped roots such as `My Drive`, `Shared With Me`, `Shared drives`, and `AppScript`, then reports top-level legacy entries that would move under `My Drive/`. Use JSON output for review or automation:

```bash
uv run python download_google_doc.pyw --tenant default --plan-migrate-my-drive-root --migration-plan-format json
```

Apply the migration only after reviewing a zero-collision plan:

```bash
uv run python download_google_doc.pyw --tenant default --apply-migrate-my-drive-root
```

The apply command aborts before moving anything if the backup root is missing or any target path already exists.

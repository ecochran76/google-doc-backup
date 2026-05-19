# Google Docs Backup Service Roadmap

## Current State

This repo is currently a single-script Google Docs backup utility that is being converted into a tenant-scoped backup service. `download_google_doc.pyw` now owns authentication and backend construction inside `main()`, exports Google Docs, Sheets, and Slides to Office formats, and can preserve Drive hierarchy under tenant-configured backup roots. It also has CLASP-based standalone Apps Script backup support.

The next target is a user-scoped, tenant-aware backup service that uses `gws` as the primary Google Workspace integration and keeps direct API/PyDrive support as a compatibility fallback.

## Target Operating Model

- Python environment and execution are managed with `uv`.
- Repo code remains generic; tenant-specific config and runtime state live outside the repo.
- Tenant config lives under a user config directory, for example `~/.config/google-doc-backup/tenants/<tenant>.toml`.
- Tenant runtime state lives under a user state directory, for example `~/.local/state/google-doc-backup/tenants/<tenant>/`.
- The default target tenant config should point backup output at `/mnt/e/SyncThing/Cloud/Google-Docs`.
- The backup service should cover:
  - My Drive
  - files and folders shared with the authenticated user
  - Shared Drives visible to the authenticated user
  - standalone Apps Script projects
- `gws` is the preferred backend for Drive listing, export, metadata reads, and shared-drive enumeration.
- Direct Drive API/PyDrive remains available as an explicit fallback backend until `gws` coverage is complete.

## P0 | Safety And Project Foundation

Status: COMPLETE FOR CURRENT SERVICE SLICE

### Goals

- Move the repo to a reliable `uv` workflow.
- Stop importing/authenticating Google services as a side effect of importing the script.
- Establish tenant-scoped config and state boundaries before expanding sync behavior.
- Preserve existing backup output and avoid destructive writes while the service model is introduced.

### Work

- Done: add `pyproject.toml` with runtime dependencies currently expressed in `requirements.txt`.
- Done: generate `uv.lock` after dependencies are declared.
- Done: add a `uv run` validation path for:
  - CLI help
  - dry-run backup planning
  - focused unit tests once tests exist
- Done: refactor the script so `main()` owns authentication and backend construction.
- Done: add config loading with explicit precedence:
  - CLI flags
  - tenant config file
  - documented defaults
- Done: add a tenant config schema that can express:
  - tenant name
  - backend preference: `gws`, `direct-api`, or `auto`
  - backup root
  - include/exclude scopes for My Drive, Shared With Me, Shared Drives, and Apps Script
  - dry-run defaults
  - retention policy
- Done: add a non-secret example config under docs, not a real tenant config.
- Deferred: split the legacy script into importable modules once the first service slice has settled.

### Acceptance

- `uv run python download_google_doc.pyw --help` works without system-package assumptions.
- A dry run can resolve the configured backup root from user config.
- Repo-local credentials are no longer required for the `gws` path.
- Direct API credentials remain local-only and are not moved into the repo.

## P1 | Backend Abstraction And GWS Primary Path

Status: IN PROGRESS

### Goals

- Introduce a small backend interface for Drive operations.
- Implement a `gws` backend first, then adapt the existing direct API code behind the same interface.
- Make fallback behavior explicit and observable.

### Work

- Done: define initial backend operations:
  - list Google-native files
  - list folders
  - list shared drives
  - get file metadata
  - export Google-native files
- Done: implement `GwsBackend` using commands such as:
  - `gws drive files list --params ...`
  - `gws drive files export --params ... --output ...`
  - `gws drive drives list --params ...`
- Done: add backend selection:
  - `--backend gws`
  - `--backend direct-api`
  - `--backend auto`
- Done: for `auto`, try `gws` first and fall back to direct API when the `gws` readiness probe fails.
- Done: classify `gws` failure kinds so `auto` does not fall back on adapter validation or JSON decode failures.
- Done: record backend used per run in the run manifest.
- Done: add guidance for `gws auth sync-gog` as the preferred local auth recovery path when `gws` credentials are missing or stale.

### Acceptance

- Done: a dry run can enumerate at least one Drive scope through `gws`.
- Done: global dry-run planning uses `gws` when available and computes a shared-drive output path.
- Done: export uses `gws` when available; live smoke exported a non-empty DOCX to a temporary directory.
- Direct API fallback is available but not silently preferred.
- Done: backend errors are classified enough to distinguish auth/API/discovery/internal fallback from validation/decode bugs.

## P2 | Drive Surface Coverage

Status: IN PROGRESS

### Goals

- Make backup coverage intentional across My Drive, Shared With Me, and Shared Drives.
- Avoid duplicate or ambiguous output paths when the same title appears in multiple folders or scopes.
- Preserve existing My Drive backups under `/mnt/e/SyncThing/Cloud/Google-Docs` while adding new roots for shared surfaces.

### Work

- Done: define canonical output roots:
  - `My Drive/<path>` for personal Drive items when `my_drive_root_mode = "scoped"`
  - legacy root-level My Drive paths when `my_drive_root_mode = "legacy"`
  - `Shared With Me/<owner-or-folder-context>/<path>`
  - `Shared drives/<drive name>/<path>`
  - `AppScript/<project name>`
- Done: add a migration compatibility mode for the existing backup tree that currently appears to contain My Drive folders directly at the root.
- Done: add a no-write legacy root migration planner for reviewing top-level moves into `My Drive/`.
- Done: add a guarded apply command that aborts on missing roots or target collisions.
- Remaining: add duplicate handling based on file ID and Drive path, not title alone.
- Done: add a local run manifest mapping:
  - file ID
  - source scope
  - source path
  - exported path
  - mime type
  - modified time
  - backend
  - last status
- Remaining: handle shortcuts deliberately:
  - record shortcut source
  - resolve target when allowed
  - avoid loops
- Done: shared-drive pagination exists through the `gws` backend; dry-run and export paths can render Shared Drive targets.
- Done: add include/exclude filters for My Drive, Shared With Me, and Shared Drives.

### Acceptance

- Partial: dry-run output reports discovered and included counts for My Drive, Shared With Me, and Shared Drives. Apps Script, skipped, unchanged, and would-download counts remain for the sync-engine slice.
- Done: existing root-level My Drive backup content is not reorganized unless `my_drive_root_mode = "scoped"` is set.
- Done: legacy root migration is reviewable through a no-write planner before any future apply command exists.
- Done: legacy root migration can be applied explicitly after a zero-collision review.
- Done: shared-drive exports land under a distinct `Shared drives/` root.
- Done: shared-with-me exports land under a distinct `Shared With Me/` root.

## P3 | Robust Sync Engine

Status: IN PROGRESS

### Goals

- Convert one-shot download behavior into a resumable, auditable sync.
- Make repeated runs cheap and predictable.
- Separate scan, plan, apply, and report phases.

### Work

- Remaining: split commands into explicit subcommands:
  - `scan`
  - `plan`
  - `sync`
  - `status`
  - `doctor`
- Done: add run manifests under user state, not the repo.
- Done: add atomic download behavior:
  - write to temporary file
  - verify non-empty output
  - set modification time
  - rename into place
- Remaining: add retry/backoff for transient Google or `gws` failures.
- Remaining: add stale/local-only detection without deleting by default.
- Done: add configurable retention for timestamped historical exports.
- Partial: add structured JSON output for automation; migration planning already supports JSON output.
- Remaining: add clear exit codes for success, partial success, auth failure, validation/config failure, and unexpected failure.

### Acceptance

- Partial: `uv run google-doc-backup plan --tenant <tenant>` is not implemented, but existing dry-run commands produce no-write plans.
- `uv run google-doc-backup sync --tenant <tenant>` can resume after interruption.
- Done: every backup run leaves a state manifest and concise human-readable summary.
- No deletion occurs without an explicit reviewed flag.

## P4 | Service Operation

Status: IN PROGRESS

### Goals

- Run backups on a predictable local schedule.
- Keep logs, status, and recovery instructions operator-friendly.
- Preserve manual dry-run and one-shot operation.

### Work

- Done: add user-systemd unit and timer templates.
- Done: add install/update commands for user-scoped service files.
- Done: add `doctor` checks for:
  - `uv`
  - `gws`
  - `gws` auth
  - backup root writability
  - SyncThing path availability
  - optional direct API fallback credentials
  - optional CLASP availability
- Partial: per-run manifests exist under user state; log rotation remains.
- Remaining: add a status command that reports last successful run, last partial run, and next scheduled run.

### Acceptance

- Done: service install is dry-run previewable before writing systemd files.
- Partial: timer is installed and inspectable through systemd; repo-local `status` command remains.
- Missing `/mnt/e/SyncThing/Cloud/Google-Docs` is reported as an operator problem, not treated as an instruction to create an unexpected substitute path.

## P5 | Tests And Release Readiness

Status: PLANNED

### Goals

- Make backup planning testable without live Google calls.
- Keep direct API fallback from regressing while `gws` becomes primary.
- Prepare the repo for versioned releases.

### Work

- Add pytest tests around:
  - config resolution
  - path mapping
  - manifest writes
  - backend fallback decisions
  - duplicate title handling
  - shared-drive path rendering
- Add command-level tests with fake backends.
- Add fixture examples for My Drive, Shared With Me, Shared Drives, shortcuts, and duplicate titles.
- Update README after the first working `uv` and tenant-config slice lands.
- Define a semantic versioning policy for the CLI/service.

### Acceptance

- `uv run pytest` passes offline.
- `uv run python download_google_doc.pyw --help` or the replacement CLI help works.
- README documents `uv`, tenant config, `gws` auth, dry-run planning, and service install flow.

## Immediate Next Slice

Continue hardening the installed service:

1. Add a repo-local status command that summarizes latest manifest status and next systemd timer run.
2. Add offline pytest coverage for config resolution, scoped path mapping, retention selection, and manifest writes.
3. Add retry/backoff around transient `gws` export/list failures.
4. Decide whether to install `clasp` and remove `--no-scripts` from the user service, or keep Apps Script backup as a manual path.

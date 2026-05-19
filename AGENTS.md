# Repository Guidelines

## Project Structure & Module Organization
The root directory holds the operational scripts. `download_google_doc.pyw` is the main entry point; it handles authentication, Drive traversal, exports, and Apps Script syncing. Supplementary utilities (`extract_google_file_id.py`, `gdoc_download_url.py`, `download_gdocs_generate_explorer_extensions.py`) cover ID parsing, URL generation, and Windows Explorer integrations. Generated OAuth artifacts (`client_secrets.json`, `credentials.json`) must stay local; never push real credentials. Logs land in `download_google_doc.log`. No dedicated package directory exists; if you add one, update `setup.py` accordingly.

## Build, Test, and Development Commands
Use `uv` for Python environment management. Run `uv sync` to install locked dependencies, `uv run python download_google_doc.pyw --help` to review options, and `uv run python download_google_doc.pyw --tenant default --show-config` to inspect tenant config without authenticating. Use `uv run python download_google_doc.pyw --dry-run --backup "D:\Backups\Drive"` for safe validation. The Apps Script backup path depends on `clasp`; verify setup inside a project folder with `npx clasp status`.
On Windows consoles that default to cp1252, set `PYTHONIOENCODING=utf-8` before running CLI commands so emoji log messages don't trigger UnicodeEncodeError (e.g. `set PYTHONIOENCODING=utf-8 && python download_google_doc.pyw ...`).

## Coding Style & Naming Conventions
Follow PEP 8: four-space indentation, `snake_case` for functions and variables, and uppercase constants for paths (e.g., `CLIENT_SECRETS_PATH`). Keep functions focused; prefer extracting helpers instead of extending the main script. Favor structured `logging` calls over prints, and document any Windows-only behavior in inline comments.

## Testing Guidelines
There is no automated suite yet. When contributing logic, add focused `pytest` modules under `tests/` and stub Drive requests so tests run offline. Always exercise relevant command combinations with `--dry-run` before running destructive syncs, and note the commands in your PR. For Apps Script flows, confirm `--no-scripts` and full backups behave as expected.

## Commit & Pull Request Guidelines
Recent commits mix imperative subjects and Conventional Commit prefixes such as `feat:`. Keep subjects under 72 characters, present tense, and add a scope when it clarifies impact (`feat(clasp): tighten status checks`). Reference manual test logs or screenshots for Explorer menu changes, and link issues when available. PRs should describe risk areas, mention credential handling, and call out any follow-up work.

## Security & Configuration Tips
Store OAuth secrets outside version control and sanitize logs before sharing. Regenerate `credentials.json` after permission updates and prune stale tokens if authentication loops appear. Avoid embedding personal paths in committed scripts; prefer config files or example placeholders.

## Policy Loading Contract

- `AGENTS.md` is a routing surface, not a one-time pointer.
- Re-read the relevant policy files under `docs/dev/policies/` at the start of any non-trivial turn.
- Re-read the relevant policy files when task scope changes mid-session.
- When behavior is ambiguous, prefer re-reading policy over improvising from stale assumptions.

## Policy Re-read Triggers

- re-read planning-related policy before opening, revising, or closing a substantive plan
- re-read documentation-related policy before changing docs, contracts, or canonical authorities
- re-read validation and closeout policy before claiming work complete

## Policy Entry

This repo keeps its durable repo-local policy under `docs/dev/policies/`.

Read and follow:
- `docs/dev/policies/0001-policy-management.md`
- `docs/dev/policies/0002-policy-upgrade-management.md`
- `docs/dev/policies/0003-policy-adoption-feedback-loop.md`
- `docs/dev/policies/0004-git-worktree-hygiene.md`
- `docs/dev/policies/0005-commit-history-discipline.md`
- `docs/dev/policies/0006-branch-and-integration-strategy.md`
- `docs/dev/policies/0007-commit-and-push-cadence.md`
- `docs/dev/policies/0008-versioning-and-release.md`
- `docs/dev/policies/0009-turn-closeout.md`
- `docs/dev/policies/0010-validation-and-handoff.md`

## Scope

- `AGENTS.md` includes repo-local guidance plus the policy entry section.
- The durable policy body lives under `docs/dev/policies/`.
- Keep repo-specific commands, environment details, and operational caveats in this file or adjacent local docs.

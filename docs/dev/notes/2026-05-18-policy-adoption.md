# 2026-05-18 Policy Adoption

## Decision

Adopted the shared `standalone-library` profile from the repo-policy-selector bundle.

## Source

- Policy selector bundle: `repo-policy-selector` v0.1.13
- Release ref: `v0.1.13`
- Source commit: `b38c90694e15562819d28a4338c5d148dc5171fd`
- Selected repo purpose: `library-cli`
- Execution bias: `max-token-efficiency`

## Adopted Modules

- `policy-management`
- `policy-upgrade-management`
- `policy-adoption-feedback-loop`
- `git-worktree-hygiene`
- `commit-history-discipline`
- `branch-and-integration-strategy`
- `commit-and-push-cadence`
- `versioning-and-release`
- `turn-closeout`
- `validation-and-handoff`

## Local Fit

The repo is a lightweight Python CLI and operational utility for Google Drive document backup. It does not currently use roadmap/runbook planning, multi-lane product planning, tenant runtime state, or website-maintenance workflows, so the heavier product and operations profiles were not adopted.

Keep repo-specific guidance in `AGENTS.md`, especially credential handling, Windows console encoding, PyDrive/CLASP validation, dry-run expectations, and the current root-script layout.

## Follow-up

- Re-check policy fit when the repo gains automated tests, release automation, or a package publishing workflow.
- If roadmap or runbook discipline becomes useful later, review `repo-product-engineering` modules instead of adding ad hoc planning rules.

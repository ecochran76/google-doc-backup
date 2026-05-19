# Policy | Validation And Handoff

## Policy

- Run the relevant validation for the touched surface before commit, handoff, or merge preparation.
- Prefer targeted verification that matches the changed area, and widen to broader suites when the impact is user-visible or cross-cutting.
- Include concrete pass/fail evidence in the handoff or closeout note.
- Keep handoff notes concise, explicit about remaining risk, and clear about the next recommended action.
- When live or manual smoke matters for the changed surface, record whether it was run and what it proved.
- Distinguish validation run by the primary agent from validation reported by a subagent or delegated worker.
- If validation was delegated, record whether the primary agent independently verified the result or accepted the delegated evidence as-is.
- For failed, timed-out, incomplete, or unknown subagent statuses, state what was trusted, what was ignored, and what remains unverified.
## Adoption Notes

Use this module when the repo:
- has multiple test or smoke surfaces with different scopes
- expects evidence-backed closeout notes
- needs clear verification and residual-risk communication before review or release

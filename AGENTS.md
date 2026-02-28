# Repository Collaboration Rules

## Multi-Template Change Confirmation

When a request could apply to both dashboard templates (`Index.html` and `IndexMECOE.html`), ask this before implementing:

`Should I update both templates or only one?`

If the user explicitly names one file/template, update only that one.
If the user says "both", implement equivalent changes in both where applicable.

## Execution Tracking

For any multi-step request, maintain an explicit execution checklist in your working notes and keep statuses current (`pending`, `in_progress`, `completed`).

When reporting progress to the user:
- State what step is currently in progress.
- State what was completed since the last update.
- State the next step.

## Pre-Ship Validation Gate

Before recommending "ship" or "done", run a validation pass and report results:
- Scope check: confirm only intended files/features were changed.
- Functional check: confirm key flows still work for the target template(s).
- Regression check: confirm no unintended cross-template impact.
- Docs check: confirm user-facing docs match the implemented behavior.

If any validation step is incomplete or failing, do not mark the work ready to ship. Report blockers and required fixes first.

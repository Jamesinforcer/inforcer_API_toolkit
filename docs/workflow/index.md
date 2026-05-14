# SOP Generator — Two-Phase Workflow

The SOP generator uses a selection workflow so **you** decide which tenant policies end up in the SOP, instead of relying on auto-matching alone.

## The three phases

| Phase | What you do | Output |
|-------|-------------|--------|
| **[Phase 1 — Export](phase-1-export.md)** | Run the tool, pick tenants, get an Excel workbook listing every policy | `output\<tenant>_Policy_Selection.xlsx` |
| **[Phase 2 — Review](phase-2-review.md)** | Open the workbook in Excel, flip `Include` to `Yes` / `No` | Reviewed workbook (saved & closed) |
| **[Phase 3 — Build](phase-3-build.md)** | Run the tool again, point it at the reviewed workbook | `output\<tenant>_SOP.docx` |

## Why three phases?

The selection step gives you visibility and control:

- You see **every** tenant policy before it lands in the SOP.
- Policies that look like they already match a canonical SOP control are pre-ticked `Yes`, so the common case is fast.
- Anything you un-tick is still recorded in **Appendix A** of the final SOP, so nothing is silently dropped.
- Re-running Phase 3 always re-fetches live policies from Inforcer, so the final document reflects current tenant state.

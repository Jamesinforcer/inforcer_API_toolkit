# Phase 3 — Build the SOP

## Steps

1. Run `run.bat` again.
2. Choose option **[2] Build an SOP from an Excel selection file**.
3. Pick the workbook (the tool lists every `*_Policy_Selection.xlsx` in `output\`).

The tool re-fetches the tenant's policies live from Inforcer (so any recent changes are picked up), then builds `output\<tenant>_SOP.docx`.

## What's in the generated SOP

### Part 2 — the 7 product sections

Uses a hybrid layout:

- Each canonical control row from the template is matched against the **ticked** policies only. Rows with at least one ticked match get their "Policy Display Name" column filled; rows with zero ticked matches are removed.
- Ticked policies that don't line up with any canonical row are appended to their product section as new rows. The Control column uses a layperson label (e.g. `E8-IBP - DfB - MacOS - Antivirus` becomes `MacOS Antivirus policy`) and the Description column is a one-sentence plain-English summary of what the policy configures.

### Appendix A — Tenant Policies Not Included in SOP

Lists every un-ticked policy grouped by product, so nothing is silently dropped.

### Appendix B — Tenant Metadata

Records tenant ID, counts, and the report generation time.

## After the build

The `output\` folder opens automatically when the run completes — your finished `<tenant>_SOP.docx` will be inside.

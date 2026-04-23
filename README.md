# Inforcer SOP Toolkit

Windows tool for Inforcer-managed tenants:

**Baseline SOP Generator** (`inforcer_sop_generator.py`) — builds tenant-specific
SOP Word documents from the `MSP-M365-Baseline-SOP.docx` template, driven by
a per-tenant Excel selection workbook.

Paired with `setup.bat` (one-time install) and `run.bat` (launcher).

---

## One-time setup

1. Put all files (`inforcer_sop_generator.py`, `run.bat`, `setup.bat`,
   `requirements.txt`, and `MSP-M365-Baseline-SOP.docx`) into a single folder.
2. Double-click **`setup.bat`**. It will:
   - create a virtual environment (`.venv`)
   - install the required packages (`requests`, `python-docx`, `openpyxl`)
   - prompt for your Inforcer API key and save it permanently as
     the Windows user env var `INFORCER_API_KEY`
3. Close any open `run.bat` windows and re-open so the key is picked up.

---

## Running the tool

Double-click **`run.bat`**. The SOP generator starts and asks whether you want
to export a selection workbook (Phase 1) or build an SOP from a reviewed one
(Phase 2). Output files land in the `output\` folder, which opens automatically
when the run completes.

---

## SOP Generator — two-phase workflow

The SOP generator uses a selection workflow so you decide which tenant
policies end up in the SOP instead of relying on auto-matching alone.

### Phase 1 — Export a selection workbook

1. Run the tool (double-click `run.bat`).
2. Choose option **[1] Export policies to an Excel selection file**.
3. Pick one or more tenants.
4. For each tenant, the tool writes
   `output\<tenant>_Policy_Selection.xlsx`.

The workbook has one row per tenant policy, with these columns:

| Column | Meaning |
| --- | --- |
| Include | `Yes` / `No` dropdown. Drives what ends up in the SOP. Policies that look like they already match a canonical SOP control are pre-ticked `Yes`. |
| Policy Name | Friendly name as reported by Inforcer. |
| Product | Product area (e.g. Entra, Intune). |
| Primary / Secondary Group | Inforcer category metadata. |
| Suggested Section | The Part 2 section the policy will land in (2.1 – 2.7). |
| Matches Template Control | Canonical controls this policy would populate. |
| Settings Summary | Compact JSON-ish view of the policy's current settings — so you can eyeball what would be inserted into the SOP. |
| Policy ID | Internal ID used to re-fetch the policy in Phase 2. Leave it alone. |

Rows 1–3 carry the tenant metadata (name, Inforcer ID, region). Don't edit
those — they're how Phase 2 knows which tenant this workbook belongs to.

### Phase 2 — Review in Excel

Open the workbook and flip `Include` to `Yes` / `No` as needed. Save and close
it before moving to Phase 3 (the build step needs to open the file and
Excel's lock will block it).

### Phase 3 — Build the SOP

1. Run `run.bat` again.
2. Choose option **[2] Build an SOP from an Excel selection file**.
3. Pick the workbook (the tool lists every
   `*_Policy_Selection.xlsx` in `output\`).

The tool re-fetches the tenant's policies live from Inforcer (so any recent
changes are picked up), then builds `output\<tenant>_SOP.docx`:

- **Part 2 (the 7 product sections)** uses a hybrid layout:
  - Each canonical control row from the template is matched against the
    **ticked** policies only. Rows with at least one ticked match get their
    "Policy Display Name" column filled; rows with zero ticked matches are
    removed.
  - Ticked policies that don't line up with any canonical row are appended
    to their product section as new rows. The Control column uses a
    layperson label (e.g. `E8-IBP - DfB - MacOS - Antivirus` becomes
    `MacOS Antivirus policy`) and the Description column is a one-sentence
    plain-English summary of what the policy configures.
- **Appendix A — Tenant Policies Not Included in SOP** lists every un-ticked
  policy grouped by product, so nothing is silently dropped.
- **Appendix B — Tenant Metadata** records tenant ID, counts, and the
  report generation time.

### Command-line flags (optional)

The interactive menu covers the common cases. For scripted / repeat runs:

```
# Phase 1 — export selection workbooks
python inforcer_sop_generator.py --export-selection
python inforcer_sop_generator.py --export-selection --tenant "Contoso"
python inforcer_sop_generator.py --export-selection --all

# Phase 2 — build an SOP from a ticked workbook
python inforcer_sop_generator.py --from-selection "output/Contoso_Policy_Selection.xlsx"

# Legacy one-shot auto-match (skips the selection step — every policy is
# considered for Part 2)
python inforcer_sop_generator.py --auto-match --tenant "Contoso"
```

Global flags: `--region {anz,eu,uk,us}` (default `anz`), `--out DIR`
(default `./output`), `--template PATH` (default `MSP-M365-Baseline-SOP.docx`).

---

## Troubleshooting

- **`INFORCER_API_KEY is not set`** — re-run `setup.bat` to save the key,
  then close and re-open `run.bat`.
- **`openpyxl package is required`** — re-run `setup.bat` to install it.
- **`template file not found`** — make sure `MSP-M365-Baseline-SOP.docx`
  is in the same folder as the Python script.
- **Workbook region is unknown** — the `Region` cell at row 3 of the
  workbook's Policies sheet must be one of `anz`, `eu`, `uk`, `us`.
- **File locked** — close the file in Word/Excel and re-run; the tool
  retries a few times on `PermissionError`.

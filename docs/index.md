# Inforcer SOP Toolkit

Windows tool for Inforcer-managed tenants.

## What it does

**Baseline SOP Generator** (`inforcer_sop_generator.py`) builds tenant-specific SOP Word documents from the `MSP-M365-Baseline-SOP.docx` template, driven by a per-tenant Excel selection workbook.

Paired with `setup.bat` (one-time install) and `run.bat` (launcher).

## Quick links

- [Getting Started](getting-started.md) — one-time setup and running the tool
- [Workflow](workflow/index.md) — the three-phase selection workflow
- [Command-Line Flags](cli.md) — scripted / repeat runs
- [Troubleshooting](troubleshooting.md) — common errors and fixes

## How it works at a glance

1. **Export** a per-tenant Excel selection workbook listing every policy.
2. **Review** in Excel — tick `Yes` / `No` against each policy.
3. **Build** the SOP — the tool re-fetches policies live from Inforcer and produces a Word document.

The selection step puts you in control of which tenant policies end up in the SOP, instead of relying on auto-matching alone.

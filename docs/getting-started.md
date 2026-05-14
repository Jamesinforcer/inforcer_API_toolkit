# Getting Started

## One-time setup

1. Put all files (`inforcer_sop_generator.py`, `run.bat`, `setup.bat`, `requirements.txt`, and `MSP-M365-Baseline-SOP.docx`) into a single folder.

2. Double-click **`setup.bat`**. It will:
    - create a virtual environment (`.venv`)
    - install the required packages (`requests`, `python-docx`, `openpyxl`)
    - prompt for your Inforcer API key and save it permanently as the Windows user environment variable `INFORCER_API_KEY`

3. Close any open `run.bat` windows and re-open them so the key is picked up.

!!! tip "Where is the API key stored?"
    `setup.bat` saves your key as a persistent Windows user environment variable. You only need to run setup once per machine — the key survives reboots and new terminal windows.

## Running the tool

Double-click **`run.bat`**.

The SOP generator starts and asks whether you want to:

- **Export** a selection workbook (Phase 1), or
- **Build** an SOP from a reviewed one (Phase 3).

Output files land in the `output\` folder, which opens automatically when the run completes.

## Next step

Continue to the [Workflow overview](workflow/index.md) to see how the three-phase process fits together.

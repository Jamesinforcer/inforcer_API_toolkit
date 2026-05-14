# Troubleshooting

## `INFORCER_API_KEY is not set`

Re-run `setup.bat` to save the key, then close and re-open `run.bat`.

The key is stored as a Windows user environment variable, so already-open terminal windows won't see it until they're restarted.

## `openpyxl package is required`

Re-run `setup.bat` to install it. This pulls in all required packages (`requests`, `python-docx`, `openpyxl`).

## `template file not found`

Make sure `MSP-M365-Baseline-SOP.docx` is in the same folder as the Python script.

## Workbook region is unknown

The `Region` cell at row 3 of the workbook's Policies sheet must be one of `anz`, `eu`, `uk`, `us`.

If you've manually edited the workbook, check that row 3 still contains a valid region code.

## File locked

Close the file in Word/Excel and re-run. The tool retries a few times on `PermissionError`, but a persistent lock will eventually fail the build.

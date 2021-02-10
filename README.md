# Spreadsheet Commander

A basic spreadsheet transfer program.

Failsafe: If the program acts weird while controlling the keyboard,
smash the mouse cursor into the corner of the screen to abort.

## Usage:
```
python3 spr_cmdr.py [OPTIONS]

Options:
    --csv CSV_FILE                      Parse as a CSV (Comma Separated Values) file.
    -e, -x, --xl, --excel EXCEL_FILE    Parse as an Excel-compatible file (see --excel-sheet-*).
    --excel-sheet-num SHEET_NUM         Choose sheet number (zero-indexed) to parse in Excel-compatible files. Defaults to the first sheet.
    --excel-sheet-name SHEET_NAME       Choose sheet name to parse in Excel-compatible files. Has priority over --excel-sheet-num.
    -c, --col-head-idx INDEX            Choose which column is the header column (which will be ignored). Zero-indexed, default None.
    -r, --row-head-idx INDEX            Choose which row is the header row (which will be ignored). Zero-indexed, default None.
    -t, --timer SECONDS                 Determines how many seconds the program waits to start transfer. Default 3.
    -n, --nan                           Pass this option to enable NaN filtering.
```

## Special Thanks
Thanks to my teacher (@BlueGadgetEngineer) for giving me the concept for this program as a challenge.

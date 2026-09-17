# nuh-helper

Helper library for enabling data studies: utilities for study enablement such as date shifting, profiling, and related workflows.

## Notebook Installation

```bash
!pip install git+https://github.com/Health-Informatics-UoN/nuh-helper.git
```

## Modules

- **`nuh_helper.date_shift`** — Date shifting for patient data in Excel/DataFrames (consistent shifts per patient ID, reproducible via linking tables).
- **`nuh_helper.profile`** - Profile a dataset into a Scan Report

## Usage; Date Shifting

### Date shifting (basic example)

```python
from nuh_helper import shift_excel_dates_inplace


    source_file = Path(__file__).parent / (
        "data/structural.with-blank.xlsx" if good else "data/structural.bad-blank.xlsx"
    )
    output_path = tmp_path / "target.xlsx"
    linking_table_old = tmp_path / "linking_table_old.csv"
    linking_table_out = tmp_path / "linking_table_out.csv"

# Configure which sheets and columns to shift
sheet_configs = {

    # shift the dates in the "paige" sheet
    "paige": {
        # patient IDs are in the "ptid" column
        "patient_id_col": "ptid",
        "date_columns": [
            # shift the dates in the DOB column
            "dob",
        ],
        "text_columns": [
            # the food column has no dates
            "food",
        ],
        "header_row": 0, # Optional: zero-based row index for column names. defaults to 0
        "skip_rows_after_header": [],
    },

    # ignore the "stuff" page
    "stuff": "skip",
}


# Shift dates in the Excel file
shift_excel_dates_inplace(
    # input / output files
    input_file="input.xlsx",
    output_file="output.xlsx",

    # columns used to compute a list of patients
    patient_sheet="paige",
    patient_id_col="ptid",


    sheet_configs=sheet_configs,
    min_shift_days=-15,  # Lower range
    max_shift_days=15,   # Upper range
    seed=42,             # For reproducibility

    # optional - csv files saving how many days each patient was shifted
    # ... do not distribute these
    # ... and probably best to keep them somewhere
    linking_table_path="linking_table_old.csv",
    linking_table_output="linking_table_new.csv",
)
```

### Reproducible Shifts with Linking Table

To use the same shifts across multiple runs, save and reuse a linking table:

```python
# First run: generate and save shifts
shift_excel_dates_inplace(
    input_file="input.xlsx",
    output_file="output.xlsx",
    patient_sheet="paige",
    patient_id_col="ptid",
    sheet_configs=sheet_configs,
    min_shift_days=-15,  # Lower range
    max_shift_days=15,   # Upper range
    seed=42,             # For reproducibility

    # optional - csv files saving how many days each patient was shifted
    # ... do not distribute these
    # ... and probably best to keep them somewhere specific
    linking_table_output="linking_table_first.csv",
)

# Second run: reuse the same shifts
shift_excel_dates_inplace(
    input_file="input_final.xlsx",
    output_file="output_second.xlsx",
    patient_sheet="paige",
    patient_id_col="ptid",
    sheet_configs=sheet_configs,
    min_shift_days=-15,  # Lower range
    max_shift_days=15,   # Upper range
    seed=42,             # For reproducibility

    # write to a new file this time
    linking_table_path="linking_table_first.csv",
    linking_table_output="linking_table_second.csv",
)
```

### Passing Non Dates (in `date_columns`)

Studies frequently include data that's not parsable as a date in the date columns.
Rarely is this a typo, it can be text like `Record missing` or `2007-09-UN` to signify that information is only partially available.
This has previously been one or two dozen entries across hundreds of cells that need to be copied to the final output unchanged.
To accommodate this, each column in each sheet can have a fixed set of strings that are passed through as-is

```python
sheet_configs = {
    "page-data": {
        "patient_id_col": "pid",
        "date_columns": ["dob"],
        "text_columns": [
            "food",
        ],
        "header_row": 1,
        "skip_rows_after_header": [],
        "pass_as_is": {  # the parameter is here
            "dob": [  # any column can have "as is" values added
                "missing",  # the values are each listed here
            ]
        },
    }
}

```

The intended workflow is ...

1. create sheet configurations
2. execute the date shifting and note non-date values detected
3. gradually build up the list of approved values per column

While this does require repeated manual intervention ...

- It's faster than searching and restoring the fields manually
- It doesn't sacrifice any control or the ability to inspect/detect problematic values

### Key parameters (date shifting)

- `input_file`: Path to input Excel file
- `output_file`: Path to output Excel file with shifted dates
- `patient_sheet`: Name of the sheet containing patient IDs
- `patient_id_col`: Name of the column containing patient IDs
- `sheet_configs`: Dictionary mapping sheet names to configuration dicts, or, the string 'skip' if that sheet should be skipped but is a valid part of the CDM.

- `patient_header_row`: (Optional) Zero-based header row for the patient sheet (default: 0). If the patient sheet is in `sheet_configs`, that sheet’s `header_row` is used instead. (... or is it?)
- `patient_skip_rows`: (Optional) Zero-based row indices to exclude from patient data (e.g. a data-type row). If the patient sheet is in `sheet_configs`, that sheet’s `skip_rows_after_header` is used instead.

- `min_shift_days` / `max_shift_days`: Range of days to shift (default: -15 to 15)
- `linking_table_path`: (Optional) Path to existing linking table CSV for reproducibility
- `linking_table_output`: (Optional) Path to save the linking table CSV
- `seed`: (Optional) Random seed for generating shifts

#### Sheet Configs Per-Page Options

- `patient_id_col`: Patient ID column name in that sheet
- `date_columns`: List of date column names to shift
- `text_columns`:
  - a per-sheet list of columns that have text only and shouldn't have dates
  - these need to be explicitly specified to prevent extra columns in the CDM
- `header_row`: (Optional) Zero-based row index for the row that contains column names
- `skip_rows_after_header`: (Optional) List of zero-based row indices to exclude from data (e.g. a data-type row immediately below the header)
- `pass_as_is`: (Optional) Dict mapping column names to lists of "non dates" that are passed through without being changed

### Excel layout (header row and merged cells)

Sheets can have a non-standard layout: e.g. a merged title row, then a description row, then the actual column names, then a data-type row. Configure as follows:

- Set `header_row` to the **zero-based index of the row that contains the column names** (the row you use for config: `patient_id_col`, `date_columns`).
- Set `skip_rows_after_header` to the indices of any rows **below the header** that should not be treated as data (e.g. a data-type row).
- **Merged cells**: The library reads the header row via openpyxl and resolves merged cells (value taken from the top-left of each merge), so column names are correct even when the sheet has merged cells. Merged ranges in the description area (rows above the header) are preserved when writing the output.

### Date shifting features

- Shifts dates consistently across multiple Excel sheets
- `shift_excel_dates_inplace`: full formatting preservation (cell styles, column widths, conditional formatting, etc.)
- Preserves Excel structure (description rows and merged cells in that area)
- Correct header detection with merged cells (openpyxl-based resolution)
- Optional skip of rows after the header (e.g. data-type row) via `skip_rows_after_header`
- Supports flexible date parsing (handles various formats and placeholders like "Unknown")
- Reproducible shifts via linking tables

## Usage; Dataset Profile

Profile a dataset and generate a Scan Report.

```python
from nuh_helper import generate_scan_report


csv_files = [
    "patients.csv",
]

generate_scan_report(csv_files, min_cell_count=5)
```

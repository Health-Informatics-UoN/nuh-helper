# nuh-helper

Helper library for enabling data studies: utilities for study enablement such as date shifting, profiling, and related workflows.

## Notebook Installation

```bash
!pip install git+https://github.com/Health-Informatics-UoN/nuh-helper.git
```

## Modules

- **`nuh_helper.date_shift`** — Date shifting for patient data in Excel/DataFrames (consistent shifts per patient ID, reproducible via linking tables).
- **`nuh_helper.profile`** - Profile a dataset into a Scan Report

## Usage

### Date shifting (basic example)

```python
from nuh_helper import shift_excel_dates
# or: from nuh_helper.date_shift import shift_excel_dates

# Configure which sheets and columns to shift
sheet_configs = {
    "patients": {
        "patient_id_col": "patient_id",
        "date_columns": ["dob", "date_of_diagnosis"],
        "header_row": 1,  # Optional: zero-based row index for column names
    },
    "labs": {
        "patient_id_col": "patient_id",
        "date_columns": ["test_date"],
        "header_row": 1,
    },
}

# Shift dates in the Excel file
shift_excel_dates(
    input_file="input.xlsx",
    output_file="output.xlsx",
    patient_sheet="patients",
    patient_id_col="patient_id",
    sheet_configs=sheet_configs,
    min_shift_days=-15,  # Lower range
    max_shift_days=15,   # Upper range
    seed=42,             # For reproducibility
    date_format="YYYY-MM-DD",
)
```

### Excluding fixed study dates with `shift_exceptions`

Some columns contain a mix of patient-specific dates (which should be shifted) and fixed study-wide dates (e.g. an end-of-study date) that must remain unchanged. Use `shift_exceptions` in `sheet_configs` to list any date values that should never be shifted:

```python
sheet_configs = {
    "patients": {
        "patient_id_col": "patient_id",
        "date_columns": ["last_alive"],
        "shift_exceptions": {
            "last_alive": ["2024-12-31"],  # end-of-study date — never shift
        },
    },
}

shift_excel_dates(
    input_file="input.xlsx",
    output_file="output.xlsx",
    patient_sheet="patients",
    patient_id_col="patient_id",
    sheet_configs=sheet_configs,
    seed=42,
)
```

The exception date strings are parsed with the same flexible parser used for all date values (supports multiple formats and placeholder strings). Exceptions are matched against the parsed date, so `"2024-12-31"` and `"31-12-2024"` both match the same calendar date.

### Reproducible Shifts with Linking Table

To use the same shifts across multiple runs, save and reuse a linking table:

```python
# First run: generate and save shifts
shift_excel_dates(
    input_file="input.xlsx",
    output_file="output.xlsx",
    patient_sheet="patients",
    patient_id_col="patient_id",
    sheet_configs=sheet_configs,
    linking_table_output="shift_mappings.csv",  # Save shifts
    seed=42,
)

# Subsequent runs: reuse the same shifts
shift_excel_dates(
    input_file="new_input.xlsx",
    output_file="new_output.xlsx",
    patient_sheet="patients",
    patient_id_col="patient_id",
    sheet_configs=sheet_configs,
    linking_table_path="shift_mappings.csv",  # Reuse saved shifts
)
```

### Preserving formatting with `shift_excel_dates_inplace`

If your workbook has rich formatting (cell styles, column widths, conditional formatting, etc.) use `shift_excel_dates_inplace` instead. It copies the input file and modifies date cells directly via openpyxl, so all formatting is preserved exactly.

```python
from nuh_helper import shift_excel_dates_inplace

shift_excel_dates_inplace(
    input_file="input.xlsx",
    output_file="output.xlsx",
    patient_sheet="patients",
    patient_id_col="patient_id",
    sheet_configs=sheet_configs,
    seed=42,
    linking_table_output="shift_mappings.csv",
)
```

The function accepts the same parameters as `shift_excel_dates` except `date_format` (not needed — the original cell format is preserved). External links and named ranges are removed from the output to avoid Excel repair dialogs.

### Passing Non Dates

Studies frequently include data that's not parsable as a date in the date columns.
Rarely is this a typo, it can be text like `Record missing` or `2007-09-UN` to signify that information is only partially available.
This has previously been one or two dozen entries across hundreds of cells that need to be copied to the final output unchanged.
To accommodate this, each column in each sheet can have a fixed set of strings that are passed through as-is

```python
sheet_configs = {
    "page-data": {
        "patient_id_col": "pid",
        "date_columns": ["dob"],
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
- `sheet_configs`: Dictionary mapping sheet names to configuration dicts with:
  - `patient_id_col`: Patient ID column name in that sheet
  - `date_columns`: List of date column names to shift
  - `header_row`: (Optional) Zero-based row index for the row that contains column names
  - `skip_rows_after_header`: (Optional) List of zero-based row indices to exclude from data (e.g. a data-type row immediately below the header)
  - `shift_exceptions`: (Optional) Dict mapping column names to lists of date strings that should never be shifted (e.g. a fixed end-of-study date). Dates are parsed using the same flexible parser as regular date values.
  - `pass_as_is`: (Optional) Dict mapping column names to lists of "non dates" that are passed through without being changed
- `patient_header_row`: (Optional) Zero-based header row for the patient sheet (default: 0). If the patient sheet is in `sheet_configs`, that sheet’s `header_row` is used instead.
- `patient_skip_rows`: (Optional) Zero-based row indices to exclude from patient data (e.g. a data-type row). If the patient sheet is in `sheet_configs`, that sheet’s `skip_rows_after_header` is used instead.
- `min_shift_days` / `max_shift_days`: Range of days to shift (default: -15 to 15)
- `linking_table_path`: (Optional) Path to existing linking table CSV for reproducibility
- `linking_table_output`: (Optional) Path to save the linking table CSV
- `seed`: (Optional) Random seed for generating shifts
- `date_format`: (Optional, `shift_excel_dates` only) Excel date format string (e.g., ‘YYYY-MM-DD’)

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

### Dataset Profile

Profile a dataset and generate a Scan Report.

```python
from nuh_helper import generate_scan_report


csv_files = [
    "patients.csv",
]

generate_scan_report(csv_files, min_cell_count=5)
```

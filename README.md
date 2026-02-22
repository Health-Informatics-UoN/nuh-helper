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
- `patient_header_row`: (Optional) Zero-based header row for the patient sheet (default: 0). If the patient sheet is in `sheet_configs`, that sheet’s `header_row` is used instead.
- `patient_skip_rows`: (Optional) Zero-based row indices to exclude from patient data (e.g. a data-type row). If the patient sheet is in `sheet_configs`, that sheet’s `skip_rows_after_header` is used instead.
- `min_shift_days` / `max_shift_days`: Range of days to shift (default: -15 to 15)
- `linking_table_path`: (Optional) Path to existing linking table CSV for reproducibility
- `linking_table_output`: (Optional) Path to save the linking table CSV
- `seed`: (Optional) Random seed for generating shifts
- `date_format`: (Optional) Excel date format string (e.g., 'YYYY-MM-DD')

### Excel layout (header row and merged cells)

Sheets can have a non-standard layout: e.g. a merged title row, then a description row, then the actual column names, then a data-type row. Configure as follows:

- Set `header_row` to the **zero-based index of the row that contains the column names** (the row you use for config: `patient_id_col`, `date_columns`).
- Set `skip_rows_after_header` to the indices of any rows **below the header** that should not be treated as data (e.g. a data-type row).
- **Merged cells**: The library reads the header row via openpyxl and resolves merged cells (value taken from the top-left of each merge), so column names are correct even when the sheet has merged cells. Merged ranges in the description area (rows above the header) are preserved when writing the output.

### Date shifting features

- Shifts dates consistently across multiple Excel sheets
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

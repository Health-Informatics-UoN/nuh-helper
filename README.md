# nuh-helper

Helper library for enabling data studies: utilities for study enablement such as date shifting, profiling, and related workflows.

## Notebook Installation

```bash
!pip install git+https://github.com/Health-Informatics-UoN/nuh-helper.git
```

## Modules

- **`nuh_helper.date_shift`** — Date shifting for patient data in Excel/DataFrames (consistent shifts per patient ID, reproducible via linking tables).
- **`nuh_helper.profile`** - Profile a dataset into a Scan Report
- **`nuh_helper.book_page`** - a library for manipulating `.xlsx` files and collections of `.csv` (or `.tsv`) files with an identical API

## Usage - date_shift

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

> `shift_excel_dates_inplace()` gets the `patient_id_col` and other `patient_sheet` values from the `sheet_config` rather than duplicate the configuration in the function call

```python
from nuh_helper import shift_excel_dates_inplace

shift_excel_dates_inplace(
    input_file="input.xlsx",
    output_file="output.xlsx",
    patient_sheet="patients",
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

The function accepts the same parameters as `shift_excel_dates` except;

- `date_format`
  - not needed — the original cell format is preserved
- `"text_columns": [???],`
  - a per-sheet list of columns that have text only and shouldn't have dates
  - these need to be explicitly specified to prevent extra columns in the CDM

### Allowing Non Dates (in `date_columns`) and Faux Dates (data that looks like dates in text columns)

Studies frequently include data that's not parsable as a date in the date columns.
This seems to change for each study, it has been text like "`_ missing _ _`" with spaces or incomplete values like `2007-09-UN` to signify that information is only partially available.
Studies may also include things like `120/80 mmHg` which can look like dates to the safety checks.

To accommodate this, teh technician can specify a **shift_ignore** file when invoking `shift_excel_dates_inplace()`

```python
from nuh_helper import shift_excel_dates_inplace

shift_excel_dates_inplace(
    input_file="input.xlsx",
    output_file="output.xlsx",

    shift_ignore = "path/to/shift_ignore.yaml",

    patient_sheet="patients",
    sheet_configs=sheet_configs,
    seed=42,
    linking_table_output="shift_mappings.csv",
)
```

This is a simple `.yaml` file with `page_name: column_name: [ignore_entry, "ignored entry2"]` format like this ...

```yaml
page-desc:
  glitter:
    - can't recall the date but on 12/11/2001 they had an itchy tummy
  foo:
    - agreed on 21st jul 2017
page-data:
  dob:
    - mssing
```

When running the shift, it's expected that the technician will ...

1. execute with an empty file
2. encounter a halt because of a value in the CDM
3. determine if it's safe to ignore or not and
   - add add the value to the ignore list
   - **or** remove, modify, or otherwise handle the value
4. gradually build up the list of approved values per column

While this does require repeated manual intervention ...

- It's faster than searching and restoring the fields manually
- It doesn't sacrifice any control or the ability to inspect/detect problematic values
- It keeps any potential patient data out of the configuration file
  - ... as the ignore list will presumably be deleted or stored with the linking_table.csv files

### "Clamping" the Study End Date

A further feature of `shift_excel_dates_inplace()` allows supplying a "clamp date" to control shifted date values.
At least one previous study collected data which included dates after the study end date.
The procedure for this study tasked the data technician to change all dates back to the study end date.
Setting `clamp_date=???` when invoking `shift_excel_dates_inplace()` will clamp all shifted dates in the same manner removing the need for human effort.
The approach will not change any dates which are not being shifted - even if they extend past this value - with the assumption being that non-shifted dates are not study data.

### Key parameters (date shifting)

- `input_file`: Path to input Excel file
- `output_file`: Path to output Excel file with shifted dates
- `patient_sheet`: Name of the sheet containing patient IDs
- `min_shift_days` / `max_shift_days`: Range of days to shift (default: -15 to 15)
- `linking_table_path`: (Optional) Path to existing linking table CSV for reproducibility
- `linking_table_output`: (Optional) Path to save the linking table CSV
- `seed`: (Optional) Random seed for generating shifts
- `clamp_date`: (Optional) If provided, all shifted dates will be clamped to this value as the latest, after they've been shifted. During previous studies there had been a requirement that any dates be updated so that none were after the study end date. Doing this during date shifting alleviates the need to search the final `.xslx` sheet for dates. Un-shifted dates, presumably things like CDM version history, are not changed.
- `sheet_configs`: Dictionary mapping sheet names to configuration dicts, or, the string 'skip' if that sheet should be skipped but is a valid part of the CDM.
  - `patient_id_col`: Patient ID column name in that sheet
  - `date_columns`: List of date column names to shift
  - `text_columns`: List of non-date columns to ignore
  - `header_row`: (Optional) Zero-based row index for the row that contains column names (default to 0 with `shift_excel_dates_inplace()`)
  - `skip_rows_after_header`: (Optional) List of zero-based row indices to exclude from data (e.g. a data-type row immediately below the header) (default to [] with `shift_excel_dates_inplace()`)
  - `shift_exceptions`: (Optional) Dict mapping column names to lists of date strings that should never be shifted (e.g. a fixed end-of-study date). Dates are parsed using the same flexible parser as regular date values. (not allowed with `shift_excel_dates_inplace()`)
- `sanity_date_latest`, `sanity_date_earliest`: (Optional) The latest and earliest dates allowed in the data. These are used as a sanity check to detect obvious typos like `1208-09-13` and will default to the current day and `1900-01-01` respectively.

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
- `shift_excel_dates_inplace`: can date-shift `.csv` files by streaming the rows rather than loading the whole file.

### Date Shifting `.csv` Files

In most cases, the process for using `shift_excel_dates_inplace()` with `.csv` or `.tsv` files is identical to working with a whole `.xlsx` file.

- `input_file: str | Path | list[Path]` is read as ...
  1. a `str` that's changed to a `Path`
  2. a `Path` (or a `str`) that's checked for `.tsv` and `.csv` files
  3. a `list[Path]` of `.csv` and `.tsv` files that are treated as workbook pages
  - each must have a different stem; `foo.tsv` and `foo.csv` will cause an error
- `output_file =` is a folder where all of the final `.csv` pages will be placed
  - if the folder/parents are missing;l they will be created
  - it **must not** have any of the input files
  - the original `.csv` or `.tsv` format/extension will be used for the output pages

There is technically no requirement that all of the `.csv` files be in the same folder.
If you're date-shifting carrot-transform output, you will want to use the list of files and remove the `person_ids.tsv` and `summary_mapstream.tsv` files.

### Bypassing Data Normalization and Hidden Date Checks

In one instance we've needed to disable data normalization and hidden date checks.
If you need to do this, at the top of your pipeline ...

```python
# this didn't work until we did it at the top of the python file

# import the HiddenDate exception class
from nuh_helper.date_shift.defs import HiddenDate
from nuh_helper.book_page.book import BaseWorkBook

# disable raising this when dates are detected in places they shouldn't be
HiddenDate.perform_checks = True

# disable changing values (mostly CSV) into date or datetime objects
# ... also stops turning "" and "  " into None
BaseWorkBook.normalize_values = False

```

## Usage - profile

Profile a dataset and generate a Scan Report.

```python
from nuh_helper import generate_scan_report


csv_files = [
    "patients.csv",
]

generate_scan_report(csv_files, min_cell_count=5)
```

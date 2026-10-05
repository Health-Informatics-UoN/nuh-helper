"""
Date shifting for patient data in Excel spreadsheets.

Consistently shifts dates for patient IDs across multiple sheets and columns
in an Excel file, with support for reproducible shifts using a linking table.
"""

import logging
import shutil
from datetime import date, datetime
from pathlib import Path
from typing import Any, cast

import datefinder
import pandas as pd
import yaml
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from nuh_helper.date_shift import _excel, _parse, mappings

logger = logging.getLogger(__name__)


class UnknownPatient(Exception):
    def __init__(self, page: str, id: str) -> None:
        message = f"Unknown {id=} on {page=}"
        super().__init__(message)
        self._message = message


class ShiftFoundNonDate(Exception):
    def __init__(
        self, page: str, row: int, col: int, col_name: str, value: str
    ) -> None:
        # the spaces and newline in the message make the value easier to copy/paste back
        message = f"{page=}[{row}, {col} @   {col_name=}  ]\n\t'{value}'"
        super().__init__(message)
        self._message: str = message
        self._page: str = page
        self._row: int = row
        self._col: int = col
        self._col_name: str = col_name
        self._value: str = value


# >> pr 127 Exception goes here
# << end of pr 127

# >> pr 128 Exception goes here
# << end of pr 128


class HiddenDate(Exception):
    def __init__(
        self,
        sheet_name: str,
        row: int,
        col: int,
        col_name: str,
        value: str,
        found: datetime,
    ) -> None:
        message = (
            f"hidden date in [{sheet_name=}, {row}, {col} @ {col_name}] {value=} "
            + f"// {found=}"
        )
        super().__init__(message)
        self._message = message
        self._sheet_name = sheet_name
        self._row = row
        self._col = col
        self._col_name = col_name
        self._value = value
        self._found = found


# >> pr 131 Exception goes here
# << end of pr 131


class DateColumnsMissing(Exception):
    def __init__(self, page_name: str, column_names: list[str]) -> None:
        message = f"date {column_names=} is missing from the cdm {page_name=}"
        super().__init__(message)
        self._message = message

        self._page_name = page_name
        self._column_names = column_names


class TextColumnsMissing(Exception):
    def __init__(self, page_name: str, column_names: list[str]) -> None:
        message = f"text {column_names=} is missing from the cdm {page_name=}"
        super().__init__(message)
        self._message = message

        self._page_name = page_name
        self._column_names = column_names


class BlankColumnHasData(Exception):
    """raised when a column with a blank name has data.

    prevents data being hidden in the wrong part of the CDM"""

    def __init__(self, page: str, row: int, col: int, value: any) -> None:
        message = f"[{page=} @ {row}, {col}] is a blank column with data {value=}"
        super().__init__(message)
        self._message = message
        self._page: str = page
        self._row: int = row
        self._col: int = col
        self._value: any = value


class PatientColumnMissing(Exception):
    """raised when the patient column is missing from any sheet"""

    def __init__(self, page_name: str, column_name: str) -> None:
        message = f"The patient {column_name=} is not present on {page_name=}"
        super().__init__(message)
        self._message = message

        self._page_name = page_name
        self._column_name = column_name


class ExtraColumn(Exception):
    """raised when an unknown column appears in a page we're shifting.

    could mean a column is named wrong, or, that the sheet_config is incomplete"""

    def __init__(self, page_name: str, column_name: str) -> None:
        message = (
            f"{column_name=} is neither ignored or shifted in the cdm {page_name=}"
        )
        super().__init__(message)
        self._message = message

        self._page_name = page_name
        self._column_name = column_name


class PageMissing(Exception):
    def __init__(self, page_name: str) -> None:
        message = f"an expected page was missing from the source cdm {page_name=}"
        super().__init__(message)
        self._message = message

        self._page_name = page_name


class ExtraPage(Exception):
    def __init__(self, page_name: str) -> None:
        message = f"an unexpected page was found in the source cdm {page_name=}"
        super().__init__(message)
        self._message = message
        self._page_name = page_name


def _get_patient_ids_and_shift_mappings(
    input_file: str,
    patient_sheet: str,
    patient_id_col: str,
    sheet_configs: dict[str, dict[str, Any]],
    min_shift_days: int,
    max_shift_days: int,
    linking_table_path: str | None,
    seed: int | None,
    patient_header_row: int,
    patient_skip_rows: list[int] | None,
) -> tuple[list[str], pd.DataFrame]:
    """
    Read patient IDs from the patient sheet and resolve shift mappings
    (load from CSV or generate). Returns (patient_ids, shift_mappings)
    with shift_mappings already normalized and deduplicated.
    """
    effective_patient_header_row = patient_header_row
    effective_patient_skip_rows = patient_skip_rows
    if patient_sheet in sheet_configs:
        cfg = sheet_configs[patient_sheet]
        effective_patient_header_row = cast(
            int, cfg.get("header_row", patient_header_row)
        )
        effective_patient_skip_rows = cfg.get(
            "skip_rows_after_header", patient_skip_rows
        )

    patient_excel = pd.ExcelFile(input_file, engine="openpyxl")
    patient_df, _, _, _ = _excel._read_sheet_with_structure(
        patient_excel,
        sheet_name=patient_sheet,
        header_row=effective_patient_header_row,
        input_file=input_file,
        skip_rows_after_header=effective_patient_skip_rows,
    )
    if patient_id_col not in patient_df.columns:
        raise PatientColumnMissing(patient_sheet, patient_id_col)

    patient_ids = (
        patient_df[patient_id_col]
        .apply(_parse._normalize_patient_id)
        .dropna()
        .unique()
        .tolist()
    )
    logger.info("Found %d patient(s) in sheet '%s'", len(patient_ids), patient_sheet)

    if linking_table_path and Path(linking_table_path).exists():
        logger.info("Loading shift mappings from '%s'", linking_table_path)
        shift_mappings = mappings.load_shift_mappings(linking_table_path)
        shift_mappings = shift_mappings[shift_mappings["patient_id"].isin(patient_ids)]
        existing_ids = set(shift_mappings["patient_id"])
        missing_ids = [pid for pid in patient_ids if pid not in existing_ids]
        if missing_ids:
            logger.warning(
                "%d patient(s) had no entry in the linking table; new shifts generated",
                len(missing_ids),
            )
            new_shifts = mappings.generate_shift_mappings(
                missing_ids, min_shift_days, max_shift_days, seed
            )
            shift_mappings = pd.concat([shift_mappings, new_shifts], ignore_index=True)
    else:
        logger.info("Generating shift mappings for %d patient(s)", len(patient_ids))
        shift_mappings = mappings.generate_shift_mappings(
            patient_ids, min_shift_days, max_shift_days, seed
        )

    shift_mappings["patient_id"] = shift_mappings["patient_id"].apply(
        _parse._normalize_patient_id
    )
    shift_mappings = shift_mappings.dropna(subset=["patient_id"]).drop_duplicates(
        subset=["patient_id"], keep="first"
    )
    return (patient_ids, shift_mappings)


def apply_date_shifts(
    df: pd.DataFrame,
    patient_id_col: str,
    date_columns: list[str],
    shift_mappings: pd.DataFrame,
    date_format: str | None = None,
    shift_exceptions: dict[str, list[str]] | None = None,
) -> pd.DataFrame:
    """
    Apply date shifts to specified columns in a DataFrame.

    Args:
        df: pd.DataFrame containing patient data.
        patient_id_col: Name of the column containing patient IDs.
        date_columns: List of column names containing dates to shift.
        shift_mappings: DataFrame with 'patient_id' and 'shift_days' columns.
        date_format: Optional date format string (e.g., 'YYYY-MM-DD').
                     Note: This parameter is kept for API compatibility but formatting
                     is applied at the Excel cell level, not in the DataFrame.
        shift_exceptions: Optional dict mapping column names to lists of date strings
                          that should never be shifted (e.g. a fixed end-of-study date).

    Returns:
        DataFrame with shifted dates.
    """
    df = df.copy()

    # Normalize patient IDs in the working DataFrame to align with mapping keys
    df[patient_id_col] = df[patient_id_col].apply(_parse._normalize_patient_id)

    shift_dict = dict(
        zip(
            shift_mappings["patient_id"],
            shift_mappings["shift_days"],
            strict=True,
        )
    )

    # Pre-parse exception dates once per column
    parsed_exceptions: dict[str, set[date]] = {}
    if shift_exceptions:
        for col, exc_values in shift_exceptions.items():
            parsed_set: set[date] = set()
            for v in exc_values:
                ts = _parse._parse_date_value(v)
                if ts is not None:
                    parsed_set.add(ts.date())
            if parsed_set:
                parsed_exceptions[col] = parsed_set

    for date_col in date_columns:
        if date_col not in df.columns:
            logger.warning(
                "Date column '%s' not found in DataFrame, skipping", date_col
            )
            continue

        # Parse flexible date strings (handles YYYY-DD-MM and placeholders "Unknown")
        non_null_before = df[date_col].notna().sum()
        df[date_col] = df[date_col].apply(_parse._parse_date_value)
        parse_failures = non_null_before - sum(x is not None for x in df[date_col])
        if parse_failures > 0:
            logger.debug(
                "Column '%s': %d value(s) could not be parsed as dates",
                date_col,
                parse_failures,
            )

        # Apply shifts
        exc_dates = parsed_exceptions.get(date_col, set())
        df[date_col] = df.apply(
            lambda row: (
                row[date_col]  # noqa: B023
                + pd.Timedelta(days=shift_dict.get(row[patient_id_col], 0))
                if row[date_col] is not None  # noqa: B023
                and row[patient_id_col] in shift_dict  # noqa: B023
                and row[date_col].date() not in exc_dates  # noqa: B023
                else row[date_col]  # noqa: B023
            ),
            axis=1,
        )

        # Convert back to date-only format (removes time component)
        df[date_col] = df[date_col].apply(
            lambda x: (
                x.date() if isinstance(x, pd.Timestamp | datetime | date) else None
            ),
        )

    return df


def shift_excel_dates(
    input_file: str,
    output_file: str,
    patient_sheet: str,
    patient_id_col: str,
    sheet_configs: dict[str, dict[str, Any]],
    min_shift_days: int = -15,
    max_shift_days: int = 15,
    linking_table_path: str | None = None,
    linking_table_output: str | None = None,
    seed: int | None = None,
    patient_header_row: int = 0,
    patient_skip_rows: list[int] | None = None,
    date_format: str | None = None,
) -> None:
    """
    Shift dates in an Excel file for patient IDs consistently across sheets.

    Args:
        input_file: Path to input Excel file.
        output_file: Path to output Excel file with shifted dates.
        patient_sheet: Name of the sheet containing patient IDs.
        patient_id_col: Name of the column containing patient IDs in the patient sheet.
        sheet_configs: Dictionary mapping sheet names to configuration dicts.
                      Each config dict should have:
                      - 'patient_id_col': Name of patient ID column in that sheet
                      - 'date_columns': List of date column names to shift
                      Optional per-sheet header handling:
                      - 'header_row': zero-based row index of the column names (default 0)
                      - 'skip_rows_after_header': list of zero-based row indices to exclude
                        from data (e.g. a data-type row immediately below the header)
        min_shift_days: Minimum number of days to shift (default: -15).
        max_shift_days: Maximum number of days to shift (default: 15).
        linking_table_path: Optional path to existing linking table CSV for reproducibility.
        linking_table_output: Path to save the linking table CSV (default: 'shift_mappings.csv').
        seed: Optional random seed for generating shifts.
        patient_header_row: Zero-based header row index for the patient sheet (default: 0).
        patient_skip_rows: Optional zero-based row indices to exclude from patient data
                     (e.g. a data-type row immediately below the header).
        date_format: Optional Excel date format string (e.g., 'YYYY-MM-DD', 'yyyy-mm-dd').
                     If None, Excel's default date format is used.
                     Common formats: 'YYYY-MM-DD', 'MM/DD/YYYY', 'DD-MM-YYYY', etc.
    """  # noqa: E501
    logger.info("Shifting dates: '%s' → '%s'", input_file, output_file)
    logger.debug(
        "Shift range: %d to %d days, seed=%s",
        min_shift_days,
        max_shift_days,
        seed,
    )

    _patient_ids, shift_mappings = _get_patient_ids_and_shift_mappings(
        input_file=input_file,
        patient_sheet=patient_sheet,
        patient_id_col=patient_id_col,
        sheet_configs=sheet_configs,
        min_shift_days=min_shift_days,
        max_shift_days=max_shift_days,
        linking_table_path=linking_table_path,
        seed=seed,
        patient_header_row=patient_header_row,
        patient_skip_rows=patient_skip_rows,
    )

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        excel_file = pd.ExcelFile(input_file, engine="openpyxl")

        for sheet_name in excel_file.sheet_names:
            default_header_row = 0
            header_row = default_header_row
            sheet_date_columns: list[str] | None = None
            skip_rows_after_header: list[int] | None = None
            sheet_shift_exceptions: dict[str, list[str]] | None = None

            if sheet_name in sheet_configs:
                config = sheet_configs[cast(str, sheet_name)]
                sheet_patient_id_col: str = cast(str, config["patient_id_col"])
                date_columns: list[str] = cast(list[str], config["date_columns"])
                header_row = cast(int, config.get("header_row", header_row))
                skip_rows_after_header = config.get("skip_rows_after_header")
                sheet_shift_exceptions = config.get("shift_exceptions")
                sheet_date_columns = date_columns
                logger.info(
                    "Shifting %d date column(s) in sheet '%s'",
                    len(date_columns),
                    sheet_name,
                )

            df, _description_df, description_rows, description_merged_ranges = (
                _excel._read_sheet_with_structure(
                    excel_file,
                    sheet_name=cast(str, sheet_name),
                    header_row=header_row,
                    input_file=input_file,
                    skip_rows_after_header=skip_rows_after_header,
                )
            )

            if sheet_name in sheet_configs:
                if sheet_patient_id_col not in df.columns:
                    raise ValueError(
                        f"Patient ID column '{sheet_patient_id_col}' not found in sheet '{sheet_name}'"  # noqa: E501
                    )

                df = apply_date_shifts(
                    df,
                    sheet_patient_id_col,
                    date_columns,
                    shift_mappings,
                    date_format=None,
                    shift_exceptions=sheet_shift_exceptions,
                )

            _excel._write_sheet_with_structure(
                writer,
                sheet_name=cast(str, sheet_name),
                data_df=df,
                description_rows=description_rows,
                header_row=header_row,
                date_columns=sheet_date_columns,
                date_format=date_format,
                description_merged_ranges=description_merged_ranges,
            )

    logger.info("Output written to '%s'", output_file)

    linking_path = linking_table_output or "shift_mappings.csv"
    shift_mappings.to_csv(linking_path, index=False)
    logger.info("Linking table saved to '%s'", linking_path)


def shift_excel_dates_inplace(
    input_file: str,
    output_file: str,
    patient_sheet: str,
    sheet_configs: dict[str, dict[str, Any]],
    min_shift_days: int = -15,
    max_shift_days: int = 15,
    linking_table_path: str | None = None,
    linking_table_output: str | None = None,
    seed: int | None = None,
    shift_ignore_yaml: None | str | Path = None,
) -> None:
    """
    Shift dates in an Excel file, preserving all cell formatting.

    Unlike shift_excel_dates(), this function copies the input file and then
    modifies date cells directly via openpyxl, so all formatting (cell styles,
    merged cells, column widths, conditional formatting, etc.) is preserved.

    Args:
        input_file: Path to input Excel file.
        output_file: Path for the output file (copy of input with shifted dates).
        patient_sheet: Name of the sheet containing patient IDs.
        sheet_configs:
          Dictionary mapping sheet names to configuration dicts, or, the string
            'skip' if that sheet should be skipped but is a valid part of the
            CDM.
          Each config dict should have:
          - 'patient_id_col': Name of patient ID column in that sheet
          - 'date_columns': List of date column names to shift
          - 'text_columns': List of non-date columns to ignore
          Optional per-sheet header handling:
          - 'header_row': zero-based row index of the column names (default 0)
          - 'skip_rows_after_header': list of zero-based row indices to
            exclude from data (e.g. a data-type row immediately below the
            header)
        min_shift_days: Minimum number of days to shift (default: -15).
        max_shift_days: Maximum number of days to shift (default: 15).
        linking_table_path: Optional path to existing linking table CSV for reproducibility.
        linking_table_output: Path to save the linking table CSV (default: 'shift_mappings.csv').
        seed: Optional random seed for generating shifts.
        shift_ignore_yaml:
            path to a .yaml file holding `page:{column:[values]}` lists of cell
            values that're ignored and passed as-is with no manipulation. None
            and "" are always added, and all values will be .strip()

        (Defaults to linking_table_path.parent / shift_ignore.do-not-commit.yaml)
    """  # noqa: E501
    logger.info("Shifting dates in-place: '%s' → '%s'", input_file, output_file)
    logger.debug(
        "Shift range: %d to %d days, seed=%s",
        min_shift_days,
        max_shift_days,
        seed,
    )

    # TODO; it'd be cool to "normalize" the linking_table_path/linking_table_output here

    # get the shift_ignore: dict[str, dict[str, list[str]]] from shift_ignore_yaml: Path
    if not shift_ignore_yaml:
        if linking_table_path:
            shift_ignore_yaml = (
                Path(linking_table_path).parent / "shift_ignore.do-not-commit.yaml"
            )
        else:
            raise RuntimeError(
                "can't guess shift_ignore_yaml without linking_table_path"
            )
    elif not isinstance(shift_ignore_yaml, Path):
        shift_ignore_yaml = Path(shift_ignore_yaml)

    if not shift_ignore_yaml.is_file():
        shift_ignore: dict[str, dict[str, set[str]]] = {}
    else:
        with open(shift_ignore_yaml) as file:
            data = yaml.safe_load(file)
            shift_ignore: dict[str, dict[str, list[str]]] = {
                page.strip(): {
                    col.strip(): {ignored.strip for ignored in data[page][col]}
                    + {"", None}
                    for col in data[page]
                }
                for page in data
            }

    # read some parameters from the config (rather than asking for them in the invoke)
    if patient_sheet not in sheet_configs:
        raise ValueError(f"{patient_sheet=}")
    patient_id_col: str = sheet_configs[patient_sheet]["patient_id_col"]
    header_row: int = sheet_configs[patient_sheet].get("header_row", 0)
    skip_rows: list[int] = sheet_configs[patient_sheet].get(
        "skip_rows_after_header", []
    )
    if not skip_rows:
        skip_rows = []

    shutil.copy2(input_file, output_file)

    _patient_ids, shift_mappings = _get_patient_ids_and_shift_mappings(
        input_file=input_file,
        patient_sheet=patient_sheet,
        patient_id_col=patient_id_col,
        sheet_configs=sheet_configs,
        min_shift_days=min_shift_days,
        max_shift_days=max_shift_days,
        linking_table_path=linking_table_path,
        seed=seed,
        patient_header_row=header_row,
        patient_skip_rows=skip_rows,
    )

    shift_dict: dict[str, int] = dict(
        zip(
            shift_mappings["patient_id"],
            shift_mappings["shift_days"],
            strict=True,
        )
    )

    wb = load_workbook(output_file, keep_links=False)
    wb.defined_names.clear()

    # check for sheets we didn't have an explanation for
    for sheet_name in wb.sheetnames:
        if sheet_name not in sheet_configs:
            raise ExtraPage(sheet_name)

    # start looping through each sheet in the configuration
    for sheet_name, sheet_config in sheet_configs.items():
        ###
        # load the workbook sheet and configuration
        ##

        if sheet_name not in wb.sheetnames:
            raise PageMissing(sheet_name)

        if isinstance(sheet_config, str) and sheet_config == "skip":
            # it's a skipped sheet
            logger.info(f"Skipping {sheet_name=}")
            continue
        logger.info(f"Shifting {sheet_name=}")

        # these will raise errors if the keys are missing - that's fine
        ws = cast(Worksheet, wb[sheet_name])
        patient_id_col: str = cast(str, sheet_config["patient_id_col"]).strip()
        date_columns: list[str] = [col.strip() for col in sheet_config["date_columns"]]
        text_columns: list[str] = [col.strip() for col in sheet_config["text_columns"]]

        header_row: int = cast(int, sheet_config.get("header_row", 0))
        skip_rows: list[int] = sheet_config.get("skip_rows_after_header", [])

        ###
        # do some checks of the configuration
        ##

        # this is an old field i want to be careful about skipping
        if "shift_exceptions" in sheet_config:
            raise RuntimeError(f"shift_exceptions was removed, update {sheet_name=}")

        # check patient_id_col isn't re-used as other column types
        if patient_id_col in date_columns:
            raise ValueError(
                f"{patient_id_col=} shouldn't be in date_columns of {sheet_name=}"
            )
        if patient_id_col in text_columns:
            raise ValueError(
                f"{patient_id_col=} shouldn't be in text_columns of {sheet_name=}"
            )

        # check that column names don't appear in both
        col_names = [col for col in date_columns if col in text_columns]
        if col_names:
            raise ValueError(
                f"{sheet_name=} has the some columns in both date and text {col_names=}"
            )
        col_names = _excel._get_row_values_resolving_merged(
            ws, header_row + 1, ws.max_column + 1
        )

        # find patient_id_idx and normalize the column name list
        patient_id_idx = None
        for col_idx in range(len(col_names)):
            col_name = col_names[col_idx]

            # if the name is None ... leave it as that (it's fine)
            if col_name is None:
                continue

            # strip the name
            col_name = col_name.strip()

            # if the name is now blank; store None as the column name
            if col_name == "":
                col_names[col_idx] = None
                continue

            # update it to just be the stripped version
            col_names[col_idx] = col_name

            # check the config to see if we know what to do with this column
            if col_name not in ([patient_id_col] + date_columns + text_columns):
                raise ExtraColumn(sheet_name, col_name)

            #
            if col_name == patient_id_col:
                patient_id_idx = col_idx

        # check to be sure we found the patient_id_col/patient_id_idx
        if patient_id_idx is None:
            raise PatientColumnMissing(sheet_name, patient_id_col)

        missing = [col for col in text_columns if col not in col_names]
        if missing:
            raise TextColumnsMissing(sheet_name, missing)
        missing = [col for col in date_columns if col not in col_names]
        if missing:
            raise DateColumnsMissing(sheet_name, missing)

        ###
        # process each row of the workbook
        ##
        for row_idx in range(ws.max_row):
            # skip all rows that happen before the header row
            if row_idx <= header_row:
                continue

            # skip any skip rows
            if row_idx in skip_rows:
                continue

            # write a log message for the user every 40 rows
            if (row_idx % 40) == 0:
                logger.info(f"Shifting {sheet_name=} up to row {row_idx}")

            # get the patient id for this row
            cell_value = ws.cell(row=row_idx + 1, column=patient_id_idx + 1).value

            if cell_value:
                if not isinstance(cell_value, str):
                    raise ValueError(
                        f"bad val {sheet_name=} {patient_id_col=} ;; {cell_value=}"
                    )

                cell_value = cell_value.strip()

            # check for date in the cell_value
            for found in datefinder.find_dates(str(cell_value)):
                raise HiddenDate(
                    sheet_name,
                    row_idx,
                    patient_id_idx,
                    patient_id_col,
                    cell_value,
                    found,
                )

            pid = _parse._normalize_patient_id(cell_value)
            if pid is None:
                # if the pid is None; the rest of the row should be None as well
                non_blank = [
                    cell
                    for cell in [
                        # get all cells
                        (
                            col_idx,
                            col_names[col_idx],
                            ws.cell(row=row_idx + 1, column=col_idx + 1).value,
                        )
                        for col_idx in range(ws.max_column)
                    ]
                    # keep the cells that aren't blank
                    if cell[2] is not None and cell[2].strip() != ""
                ]
                if non_blank:
                    raise ValueError(
                        f"{row_idx=} has no pid, should be blank but has {non_blank}"
                    )

                # we've completed the checks for a None pid
                # ... so we need to skip the rest of this loop
                continue

            # determine how many days to shift
            shift_days = shift_dict.get(pid)
            if shift_days is None:
                # all patients need to have a shift value
                raise UnknownPatient(sheet_name, pid)
            shift_delta = pd.Timedelta(days=shift_days)

            # now scan each column
            # ... even when page has no date_columns; still check for hidden dates
            for col_idx in range(ws.max_column):
                # skip the patient id column (it was already checked anyway)
                if col_idx == patient_id_idx:
                    continue

                # get the cell value. replace it with None if it's just whitespace
                cell = ws.cell(row=row_idx + 1, column=col_idx + 1)
                cell_value = cell.value
                if isinstance(cell_value, str):
                    cell_value = cell_value.strip()
                    if not cell_value:
                        cell_value = None
                        cell.value = None

                # so if the cell is empty; skip the rest of the checks for this cell
                # ... we go tot he next column in the row
                if cell_value is None:
                    continue

                # if there's no column name; this column of the row should be None
                # ... and we win't reach this line if the cell_value was None
                col_name = col_names[col_idx]
                if col_name is None:
                    raise BlankColumnHasData(sheet_name, row_idx, col_idx, cell_value)

                # skip values in shift_ignore
                if "shift_ignore" in sheet_config:
                    raise RuntimeError("move shift_ignore from sheet_configs to a file")
                if sheet_name not in shift_ignore:
                    shift_ignore[sheet_name] = {}
                if col_name not in shift_ignore[sheet_name]:
                    shift_ignore[sheet_name][col_name] = {None, ""}
                if cell_value in shift_ignore[sheet_name][col_name]:
                    continue

                # for text columns, we now *just* need to check for a hidden date
                if col_name in text_columns:
                    # check for dates in non-date columns
                    for found in datefinder.find_dates(str(cell_value)):
                        raise HiddenDate(
                            sheet_name,
                            row_idx,
                            col_idx,
                            col_name,
                            cell_value,
                            found,
                        )
                    cell.value = cell_value
                    continue

                # we've already checked this (sort of)
                # assert col_name in date_columns

                # cell_value can be str | datetime | date
                # ... so parsed might not always succeed
                parsed = _parse._parse_date_value(cell_value)
                if parsed is None:
                    # raise an error; we already handled values we should pass as-is
                    raise ShiftFoundNonDate(
                        sheet_name, row_idx, col_idx, col_name, cell_value
                    )

                # perform the actual shifting
                shifted = parsed + shift_delta

                # controls wether the data is displayed with the 00:00:00 in Excel
                # ... it might be nice to just drop the time if it's 00:00:00
                if isinstance(cell_value, datetime):
                    cell.value = cast(Any, shifted.to_pydatetime())
                else:
                    cell.value = cast(Any, shifted.to_pydatetime().date())
                    cell.number_format = "yyyy-mm-dd"

        logger.info(f"Shifting {sheet_name=} processed {row_idx} rows")

    wb.save(output_file)
    logger.info("Output written to '%s'", output_file)

    linking_path = linking_table_output or "shift_mappings.csv"
    shift_mappings.to_csv(linking_path, index=False)
    logger.info("Linking table saved to '%s'", linking_path)


# Re-export public API
from nuh_helper.date_shift.mappings import (  # noqa: E402
    generate_shift_mappings,
    load_shift_mappings,
)

__all__ = [
    generate_shift_mappings,
    load_shift_mappings,
    shift_excel_dates_inplace,
]

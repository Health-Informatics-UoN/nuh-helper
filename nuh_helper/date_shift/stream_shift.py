import logging
from dataclasses import dataclass, replace
from datetime import datetime
from enum import StrEnum
from pathlib import Path
from typing import Any, Literal, cast

import datefinder
import yaml

from nuh_helper.book_page import Book, book_open
from nuh_helper.date_shift import _parse, mappings


class ColumnType(StrEnum):
    TEXT = "text"
    DATE = "date"
    IGNORE = "skip"


logger = logging.getLogger(__name__)


class PatientColumnMissing(Exception):
    """raised when the patient column is missing from any sheet"""

    def __init__(self, page_name: str, column_name: str) -> None:
        message = f"The patient {column_name=} is not present on {page_name=}"
        super().__init__(message)
        self._message = message

        self._page_name = page_name
        self._column_name = column_name


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


def shift_excel_dates_inplace(
    input_file: str | Path | list[Path],
    output_file: str | Path,
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
    logger.info(
        "Shifting dates in-place: '%s' → '%s'",
        input_file
        if not isinstance(input_file, list)
        else [file.name for file in input_file],
        output_file,
    )
    logger.debug(
        "Shift range: %d to %d days, seed=%s",
        min_shift_days,
        max_shift_days,
        seed,
    )

    book: Book = book_open(input_file)
    logger.info(f"Book has {type(book)=}")

    if not linking_table_path:
        linking_table_path = None
    elif not isinstance(linking_table_path, Path):
        linking_table_path = Path(linking_table_path)

    if not linking_table_output:
        linking_table_output = None
    elif not isinstance(linking_table_output, Path):
        linking_table_output = Path(linking_table_output)

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
        logger.warning(f"didn't find shift_ignore file {shift_ignore_yaml=}")
    else:
        with open(shift_ignore_yaml) as file:
            data = yaml.safe_load(file)
            shift_ignore: dict[str, dict[str, list[str]]] = {
                page.strip(): {
                    col.strip(): [
                        ignored.strip() if isinstance(ignored, str) else ignored
                        for ignored in data[page][col]
                    ]
                    + ["", None]
                    for col in data[page]
                }
                for page in data
            }

    if patient_sheet not in sheet_configs:
        raise PageMissing(patient_sheet)

    with book[patient_sheet] as page:
        scanner = PatientRowStateMachine(sheet_configs)
        shift_deltas = mappings.patient_shift_deltas(
            [
                pid.patient_id
                for pid in [
                    scanner.next_details(page.name, row, data)
                    for row, data in page.stream_rows()
                ]
                if pid
            ],
            linking_table_path,
            linking_table_output or "shift_mappings.csv",
            seed,
            min_shift_days,
            max_shift_days,
        )
        scanner = None
    logger.info(f"loaded deltas {len(shift_deltas)=}")

    # check for missing sheets
    for sheet_name in sheet_configs:
        if sheet_name not in book:
            raise PageMissing(sheet_name)

    ##
    # this calls the `.stream_copy()` method to copy from one book to another
    # WARNING; this works like a callback rather than a loop;
    # - each iteration is a callback from copy with a source row
    # - each row is passed with a sheet name and a row number
    # - modify the values in the row_cells to change what's in the output file(s)
    # - no option to add/remove rows is available
    # - no option to jump/seek through rows is available
    #   - ... which is why this acts a bit weird
    scanner = PatientRowStateMachine(sheet_configs)
    for sheet_name, row_idx, row_cells in book.stream_copy(output_file):
        # check if we're in a patient row
        # ... the method will raise exceptions if something is wrong or mis-configured
        detail = scanner.next_details(sheet_name, row_idx, row_cells)
        if not detail:
            continue

        # determine how many days to shift
        shift_delta = shift_deltas.get(detail.patient_id)
        if shift_delta is None:
            # all patients need to have a shift value
            raise UnknownPatient(sheet_name, detail.patient_id)

        # now scan each column
        # ... even when page has no date_columns; still check for hidden dates
        # for col_idx in range(len(row_cells)):
        for col_idx, cell_value in enumerate(row_cells):
            # skip the patient id column (it was already checked anyway)
            if col_idx == detail.patient_id_idx:
                continue

            # get the cell value. replace it with None if it's just whitespace
            if isinstance(cell_value, str):
                cell_value = cell_value.strip()

            if not cell_value:
                cell_value = None

            col_name = detail.col_names[col_idx]

            # with no column name; this column of the row should have been None
            if cell_value and not col_name:
                raise BlankColumnHasData(sheet_name, row_idx, col_idx, cell_value)

            # choose what to do based on the column/cell type
            match detail.column_type(shift_ignore, col_name, cell_value):
                case ColumnType.IGNORE:
                    # "skip" is sheet-wide, but, ignore is for a single cell of data
                    pass
                case ColumnType.TEXT:
                    # check for dates in non-date columns
                    for found in datefinder.find_dates(str(cell_value)):
                        logger.debug(
                            f"found hidden date {sheet_name=}/{col_name=}"
                            + f":{cell_value}"
                        )
                        ignored = shift_ignore[sheet_name][col_name]
                        logger.debug(
                            f"shift_ignore[{sheet_name=}][{col_name=}]="
                            + f" {len(ignored)}"
                        )
                        for ig in ignored:
                            logger.debug(f"- {ig}")

                        raise HiddenDate(
                            sheet_name,
                            row_idx,
                            col_idx,
                            col_name,
                            cell_value,
                            found,
                        )
                case ColumnType.DATE:
                    # cell_value can be str | datetime | date
                    # ... so parsed might not always succeed
                    parsed = _parse._parse_date_value(cell_value)
                    if parsed is None:
                        # raise an error; we already handled ignored values
                        raise ShiftFoundNonDate(
                            sheet_name, row_idx, col_idx, col_name, cell_value
                        )

                    # perform the actual shifting
                    shifted = parsed + shift_delta

                    # would might be nice to always drop the 00:00:00
                    if isinstance(cell_value, datetime):
                        cell_value = cast(Any, shifted.to_pydatetime())
                    else:
                        cell_value = cast(Any, shifted.to_pydatetime().date())
            row_cells[col_idx] = cell_value

    logger.info("Output written to '%s'", output_file)


@dataclass
class PatientPageDetail:
    page_name: str
    col_names: tuple[str, ...]
    patient_id_idx: int
    patient_id_col: str
    patient_id: str
    date_columns: tuple[str, ...]
    text_columns: tuple[str, ...]

    def column_type(
        self,
        shift_ignore: dict[str, dict[str, list[str]]],
        col_name: str,
        cell_value: str,
    ) -> bool:
        sheet_name: str = self.page_name

        if sheet_name not in shift_ignore:
            shift_ignore[sheet_name] = {}

        if col_name not in shift_ignore[sheet_name]:
            shift_ignore[sheet_name][col_name] = {None, ""}

        return (
            ColumnType.IGNORE
            if cell_value in shift_ignore[sheet_name][col_name]
            else ColumnType.TEXT
            if col_name in self.text_columns
            else ColumnType.DATE
        )


class PatientRowStateMachine:
    """builds upon the callbacks from book_page to handle row/page switching. does all
    the checking for configs and what page we're on. needed because both patient IDs and
    The Big Loop need the same checks"""

    def __init__(self, sheet_configs: dict[str, dict]) -> None:
        self._sheet_configs = sheet_configs
        self._header_row: None | int = None
        self._skip_rows: None | list[int] = None
        self._page_detail: None | PatientPageDetail = None

    def next_details(
        self, sheet_name: str, row_idx: int, row_cells: list
    ) -> PatientPageDetail | Literal[False]:
        """determine if we should do a shift/check based on where we are in the loop."""

        # check that it's in the configs
        if sheet_name not in self._sheet_configs:
            raise ExtraPage(sheet_name)

        sheet_config: dict | str = self._sheet_configs[sheet_name]

        if row_idx == 0:
            ###
            # process the "which row is header and which do we skip" configuration
            ##

            self._page_detail = None

            # do nothing to the skipped ones
            if sheet_config == "skip":
                return False

            self._header_row: int = cast(int, sheet_config.get("header_row", 0))
            self._skip_rows: list[int] = sheet_config.get("skip_rows_after_header", [])

            # check for holes in skip_rows (possible issue)
            if self._skip_rows:
                missing_skipped = [
                    not_skipped + 1
                    for not_skipped in range(self._header_row, max(self._skip_rows))
                    if (not_skipped + 1) not in self._skip_rows
                ]
                if missing_skipped:
                    header_row = self._header_row
                    skip_rows = self._skip_rows
                    raise RuntimeError(
                        f"{sheet_name=} and {missing_skipped=} but"
                        + f" {header_row=} and {skip_rows=}"
                    )

            # check if the config hard-codes patient data to skip
            if "shift_exceptions" in sheet_config:
                raise RuntimeError(
                    f"shift_exceptions was removed, update {sheet_name=}"
                )
            if "shift_ignore" in sheet_config:
                raise RuntimeError("move shift_ignore from sheet_configs to a file")

        if (
            (sheet_config == "skip")
            or (row_idx < self._header_row)
            or (row_idx in self._skip_rows)
        ):
            return False

        if row_idx == self._header_row:
            ###
            # now read the "which column" configuration
            ##

            # grab the configuration
            patient_id_col: str = str(sheet_config["patient_id_col"]).strip()
            date_columns: list[str] = normalize_inplace(sheet_config["date_columns"])
            text_columns: list[str] = normalize_inplace(sheet_config["text_columns"])

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
            overlapping_columns = [col for col in date_columns if col in text_columns]
            if overlapping_columns:
                raise ValueError(
                    f"{sheet_name=} has the some columns in both date and text "
                    + str(overlapping_columns)
                )

            # normalize the column names
            col_names = normalize_inplace(row_cells)

            # find patient_id_idx and check columns
            patient_id_idx = None
            for col_idx, col_name in enumerate(col_names):
                if not col_name:
                    continue

                # check the config to see if we know what to do with this column
                if col_name not in ([patient_id_col] + date_columns + text_columns):
                    raise ExtraColumn(sheet_name, col_name)

                # select the patient id
                if col_name == patient_id_col:
                    if patient_id_idx is not None:
                        raise KeyError(
                            f"{sheet_name=} multiple columns {patient_id_col=}"
                        )
                    patient_id_idx = col_idx

            # check to be sure we found the patient_id_col/patient_id_idx
            if patient_id_idx is None:
                raise PatientColumnMissing(sheet_name, patient_id_col)

            # check for columns in the config and not in the sheet
            missing = [col for col in text_columns if col not in col_names]
            if missing:
                raise TextColumnsMissing(sheet_name, missing)

            missing = [col for col in date_columns if col not in col_names]
            if missing:
                raise DateColumnsMissing(sheet_name, missing)

            # we've completed the header_row; save the archetype of the page detail
            self._page_detail = PatientPageDetail(
                page_name=sheet_name,
                col_names=col_names,
                patient_id_idx=patient_id_idx,
                patient_id_col=patient_id_col,
                patient_id="<header row>",
                date_columns=tuple(date_columns),
                text_columns=tuple(text_columns),
            )

            # we don't shift anything in the header_row
            return False

        # get the patient id for this row
        patient_id = row_cells[self._page_detail.patient_id_idx]

        # check for date in the patient_id
        for found in datefinder.find_dates(str(patient_id)):
            raise HiddenDate(
                sheet_name,
                row_idx,
                self._page_detail.patient_id_idx,
                self._page_detail.patient_id_col,
                patient_id,
                found,
            )

        patient_id = _parse._normalize_patient_id(patient_id)
        if patient_id is None:
            # if the pid is None; the rest of the row should be None as well
            non_blank = [
                cell
                for cell in row_cells
                # keep the cells that aren't blank
                if cell is not None and str(cell).strip() != ""
            ]
            if non_blank:
                raise ValueError(
                    f"{row_idx=} has no pid, should be blank but has {non_blank}"
                )

            # we've completed the checks for a None pid
            # ... so we need to skip the rest of this loop
            return False

        # it's a valid pid - return a page detail object
        # ... so that we don't need to examine this scanner
        return replace(
            self._page_detail,
            patient_id=patient_id,
        )


def normalize_inplace(row: list[str]) -> list[str]:
    for idx, val in enumerate(row):
        if isinstance(val, str):
            val = val.strip()

        if not val:
            val = None

        row[idx] = val
    return row

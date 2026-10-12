import logging
from datetime import datetime
from pathlib import Path
from typing import Any, cast

import datefinder
import yaml

from nuh_helper.book_page import Book, book_open
from nuh_helper.date_shift import _parse, mappings

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
        patient_header_row = sheet_configs[patient_sheet].get("header_row", 0)
        patient_skip_rows = sheet_configs[patient_sheet].get("skip_rows", [])
        patient_id_col_name = sheet_configs[patient_sheet]["patient_id_col"]
        patient_ids = []
        for row, data in page.stream_rows():
            if row < patient_header_row or row in patient_skip_rows:
                continue
            elif row == patient_header_row:
                if patient_id_col_name not in data:
                    raise PatientColumnMissing(page.name, patient_id_col_name)
                patient_id_col = data.index(patient_id_col_name)
            else:
                patient_id = data[patient_id_col]
                if not patient_id:
                    continue
                patient_id = _parse._normalize_patient_id(patient_id.strip())
                if patient_id in patient_ids:
                    continue
                patient_ids.append(patient_id)

        shift_deltas = mappings.patient_shift_deltas(
            patient_ids,
            linking_table_path,
            linking_table_output or "shift_mappings.csv",
            seed,
            min_shift_days,
            max_shift_days,
        )
    logger.info(f"loaded deltas {len(shift_deltas)=}")

    # check for missing sheets
    for sheet_name in sheet_configs:
        if sheet_name not in book:
            raise PageMissing(sheet_name)

    last_name = None
    last_row = None
    skipping_sheet_name = ""

    ##
    # this calls the `.stream_copy()` method to copy from one book to another
    # WARNING; this works like a callback rather than a loop;
    # - each iteration is a callback from copy with a source row
    # - each row is passed with a sheet name and a row number
    # - modify the values in the row_cells to change what's in the output file(s)
    # - no option to add/remove rows is available
    # - no option to jump/seek through rows is available
    #   - ... which is why this acts a bit weird
    steate_tool = row_state(sheet_configs)
    for sheet_name, row_idx, row_cells in book.stream_copy(output_file):
        steate_tool.in_body(sheet_name, row_idx, row_cells)
        #####
        # Part-1; check if we've reached a different sheet
        ##
        assert isinstance(sheet_name, str)
        if row_idx != 0:
            # SAN/safety check in case something tries to skip showing us rows
            if sheet_name != last_name:
                raise RuntimeError(f"didn't see row0 from {sheet_name=}")
            last_row = row_idx
        else:
            if last_name and last_row:
                logger.info(f"Shifting {last_name=} processed {last_row} rows")
            last_name = sheet_name
            last_row = row_idx

            # check that it's in the configs
            if sheet_name not in sheet_configs:
                raise ExtraPage(sheet_name)

            sheet_config: dict | str = sheet_configs[sheet_name]

            # do nothing to the skipped ones
            if sheet_config == "skip":
                skipping_sheet_name = sheet_name
                continue

            # grab the configuration
            patient_id_col: str = cast(str, sheet_config["patient_id_col"]).strip()
            date_columns: list[str] = [
                col.strip() for col in sheet_config["date_columns"]
            ]
            text_columns: list[str] = [
                col.strip() for col in sheet_config["text_columns"]
            ]

            header_row: int = cast(int, sheet_config.get("header_row", 0))
            skip_rows: list[int] = sheet_config.get("skip_rows_after_header", [])

            ###
            # do some checks of the configuration
            ##

            # check for holes in skip_rows (possible issue)
            if skip_rows:
                missing_skipped = [
                    not_skipped + 1
                    for not_skipped in range(header_row, max(skip_rows))
                    if (not_skipped + 1) not in skip_rows
                ]
                if missing_skipped:
                    raise RuntimeError(
                        f"{sheet_name=} and {missing_skipped=} but"
                        + f" {header_row=} and {skip_rows=}"
                    )

            # check if the config hard-codes patient data to skip
            if "shift_exceptions" in sheet_config:
                raise RuntimeError(
                    f"shift_exceptions was removed, update {sheet_name=}"
                )

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
                    f"{sheet_name=} has the some columns in both date and text "
                    + str(col_names)
                )

            col_names = None
            patient_id_idx = None

        #####
        # Part-2; skip skipped sheets, and, skip skipped rows
        ##

        # do nothing to the skipped sheets
        if sheet_config == "skip":
            # "safety" check so that we only skip rows with the correct name
            if skipping_sheet_name != sheet_name:
                raise RuntimeError("this shouldn't happen")
            continue
        logger.info(f"Shifting {sheet_name=}")

        # skip pre-header and skip rows
        if (row_idx < header_row) or (row_idx in skip_rows):
            # there's an earlier check to find "holes" in the skip-rows
            continue

        #####
        # Part-3; handle the header rows
        ##
        if row_idx == header_row:
            # get the real column names now
            col_names = row_cells

            # find patient_id_idx and normalize the column name list
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

                # select the patient id
                if col_name == patient_id_col:
                    patient_id_idx = col_idx

            # check to be sure we found the patient_id_col/patient_id_idx
            if patient_id_idx is None:
                raise mappings.PatientColumnMissing(sheet_name, patient_id_col)

            missing = [col for col in text_columns if col not in col_names]
            if missing:
                raise TextColumnsMissing(sheet_name, missing)
            missing = [col for col in date_columns if col not in col_names]
            if missing:
                raise DateColumnsMissing(sheet_name, missing)

            continue

        #####
        # Part-4; process normal columns
        ##
        if col_names is None:
            raise RuntimeError("does the sheet skip all the skip rows?")
        else:
            # code used to be indented like this

            ###
            # process each row of the workbook
            ##

            # write a log message every 40 rows (show we're not crashed)
            if (row_idx % 40) == 0:
                logger.info(f"Shifting {sheet_name=} up to row {row_idx}")

            # get the patient id for this row
            pid = row_cells[patient_id_idx]

            # check for date in the cell_value
            for found in datefinder.find_dates(str(pid)):
                raise HiddenDate(
                    sheet_name,
                    row_idx,
                    patient_id_idx,
                    patient_id_col,
                    pid,
                    found,
                )

            pid = _parse._normalize_patient_id(pid)
            if pid is None:
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
                continue

            # determine how many days to shift
            shift_delta = shift_deltas.get(pid)
            if shift_delta is None:
                # all patients need to have a shift value
                raise UnknownPatient(sheet_name, pid)

            # now scan each column
            # ... even when page has no date_columns; still check for hidden dates
            for col_idx in range(len(row_cells)):
                # skip the patient id column (it was already checked anyway)
                if col_idx == patient_id_idx:
                    continue

                # get the cell value. replace it with None if it's just whitespace
                cell_value = row_cells[col_idx]
                if isinstance(cell_value, str):
                    cell_value = cell_value.strip()
                    if not cell_value:
                        cell_value = None

                # so if the cell is empty; skip the rest of the checks for this cell
                # ... we go tot he next column in the row
                if cell_value is None:
                    row_cells[col_idx] = None
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
                        logger.debug(
                            f"found hidden date {sheet_name=}/{col_name=}:{cell_value}"
                        )
                        ignored = shift_ignore[sheet_name][col_name]
                        logger.debug(
                            f"shift_ignore[{sheet_name=}][{col_name=}]= {len(ignored)}"
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
                    row_cells[col_idx] = cell_value
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
                    row_cells[col_idx] = cast(Any, shifted.to_pydatetime())
                else:
                    row_cells[col_idx] = cast(Any, shifted.to_pydatetime().date())
                    print('redo cell.number_format = "yyyy-mm-dd"')

    logger.info(
        f"Shifting {last_name=} processed {last_row} rows (and that's the last page)"
    )
    logger.info("Output written to '%s'", output_file)


class row_state:
    def __init__(self, sheet_configs: dict[str, dict]) -> None:
        self._sheet_configs = sheet_configs
        self._last_name: None | str = None
        self._last_row: None | int = None
        self._header_row: None | int = None
        self._skip_rows: None | list[int] = None

        self._col_names: None | list[str] = None
        self._patient_id_idx: None | int = None
        self._patient_id_col: None | str = None
        self._date_columns: None | list[str] = None
        self._text_columns: None | list[str] = None
        self._sheet_config: None | dict | str = None

    def in_body(self, sheet_name: str, row_idx: int, row_cells: list) -> str | bool:
        # check that it's in the configs
        if sheet_name not in self._sheet_configs:
            raise ExtraPage(sheet_name)

        if row_idx != 0:
            # SAN/safety check in case something tries to skip showing us rows
            if sheet_name != self._last_name:
                raise RuntimeError(f"didn't see row0 from {sheet_name=}")
            self._last_row = row_idx

        if row_idx == 0:
            self._last_name = sheet_name
            self._last_row = row_idx

            self._sheet_config: dict | str = self._sheet_configs[sheet_name]

            # do nothing to the skipped ones
            if self._sheet_config == "skip":
                return False

            # grab the configuration
            self.pid_column: str = cast(
                str, self._sheet_config["patient_id_col"]
            ).strip()
            self._date_columns: list[str] = [
                col.strip() for col in self._sheet_config["date_columns"]
            ]
            self._text_columns: list[str] = [
                col.strip() for col in self._sheet_config["text_columns"]
            ]

            self._header_row: int = cast(int, self._sheet_config.get("header_row", 0))
            self._skip_rows: list[int] = self._sheet_config.get(
                "skip_rows_after_header", []
            )

            ###
            # do some checks of the configuration
            ##

            # check for holes in skip_rows (possible issue)
            if self._skip_rows:
                missing_skipped = [
                    not_skipped + 1
                    for not_skipped in range(self._header_row, max(self._skip_rows))
                    if (not_skipped + 1) not in self._skip_rows
                ]
                if missing_skipped:
                    raise RuntimeError(
                        f"{sheet_name=} and {missing_skipped=} but"
                        + f" {self._header_row=} and {self._skip_rows=}"
                    )

            # check if the config hard-codes patient data to skip
            if "shift_exceptions" in self._sheet_configs:
                raise RuntimeError(
                    f"shift_exceptions was removed, update {sheet_name=}"
                )

            # check patient_id_col isn't re-used as other column types
            if self.pid_column in self._date_columns:
                raise ValueError(
                    f"{self.pid_column=} shouldn't be in date_columns of {sheet_name=}"
                )
            if self.pid_column in self._text_columns:
                raise ValueError(
                    f"{self.pid_column=} shouldn't be in text_columns of {sheet_name=}"
                )

            # check that column names don't appear in both
            col_names = [col for col in self._date_columns if col in self._text_columns]
            if col_names:
                raise ValueError(
                    f"{sheet_name=} has the some columns in both date and text "
                    + str(col_names)
                )

            self._col_names = None
            self._patient_id_idx = None

        if self._sheet_config == "skip":
            return False

        if row_idx < self._header_row:
            return False

        if row_idx == self._header_row:
            raise Exception("compute columns and indcides")

        if row_idx in self._skip_rows:
            return False

        # get the patient id for this row
        pid = row_cells[self._patient_id_idx]

        # check for date in the cell_value
        for found in datefinder.find_dates(str(pid)):
            raise HiddenDate(
                sheet_name,
                row_idx,
                self._patient_id_idx,
                self._patient_id_col,
                pid,
                found,
            )

        pid = _parse._normalize_patient_id(pid)
        if pid is None:
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

        # it's a valid pid - return it
        return pid

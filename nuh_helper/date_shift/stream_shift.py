import logging
from datetime import datetime
from pathlib import Path
from typing import Any, cast

import datefinder
from pandas import Timedelta

from nuh_helper.book_page import Book, book_open
from nuh_helper.date_shift import _parse, mappings
from nuh_helper.date_shift.defs import (
    BlankColumnHasData,
    FoundExtraPages,
    FoundMissingPages,
    HiddenDate,
    ShiftFoundNonDate,
    UnknownPatient,
    sheet_configs_t,
)
from nuh_helper.date_shift.ignore import shift_ignore_load, shift_ignore_t
from nuh_helper.date_shift.row_machine import (
    ColumnType,
    PatientRow,
    PatientRowStateMachine,
)

logger = logging.getLogger(__name__)


def shift_excel_dates_inplace(
    input_file: str | Path | list[Path],
    output_file: str | Path,
    patient_sheet: str,
    sheet_configs: sheet_configs_t,
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
        input_file: Path to input Excel file, or, folder of .csv files, or list of .csv files
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
        f"Shifting dates in-place: '{input_file=}' -> '{output_file=}'"
        + f"{min_shift_days=} {max_shift_days=} {seed=}"
    )

    book: Book = book_open(input_file)
    logger.info(f"Book has {type(book)=}")

    # check for missing sheets
    if sheet_names := [name for name in sheet_configs if name not in book]:
        raise FoundMissingPages(sheet_names)

    # check that there aren't any extra pages in config
    if sheet_names := [page.name for page in book if page.name not in sheet_configs]:
        raise FoundExtraPages(sheet_names, list(sheet_configs))

    # load the shift_ignore rules
    shift_ignore: shift_ignore_t = shift_ignore_load(shift_ignore_yaml)

    shift_deltas = mappings.load_shift_deltas(
        book,
        patient_sheet,
        sheet_configs,
        min_shift_days,
        max_shift_days,
        old_linking_table=linking_table_path,
        out_linking_table=linking_table_output,
        seed=seed,
    )

    logger.info("loaded shift_ignore and shift_deltas")

    ##
    # this calls the `.stream_copy()` method to copy from one book to another
    # WARNING; this works like a callback rather than a loop;
    # - each iteration is a callback from copy with a source row
    # - each row is passed with a sheet name and a row number
    # - modify the values in the row_cells to change what's in the output file(s)
    # - no option to add/remove rows is available
    # - no option to jump/seek through rows is available
    #   - ... which is why this acts a bit weird
    for detail in PatientRowStateMachine(sheet_configs).book_copy(book, output_file):
        # determine how many days to shift
        shift_delta: Timedelta = shift_deltas.get(detail.patient_id)
        if shift_delta is None:
            # all patients need to have a shift value
            raise UnknownPatient(detail.page_name, detail.patient_id)

        # process each cell of this row using the subroutine
        for col_idx, cell_value in enumerate(detail.row_cells):
            shifted_cell = _shift_cell(
                cell_value=cell_value,
                col_idx=col_idx,
                detail=detail,
                shift_delta=shift_delta,
                shift_ignore=shift_ignore,
            )

            if str(shifted_cell) != str(cell_value):
                detail.row_cells[col_idx] = shifted_cell

    logger.info(f"Output written {output_file=}")


def _shift_cell(
    cell_value: str | object | None,
    col_idx: int,
    detail: PatientRow,
    shift_delta: Timedelta,
    shift_ignore: shift_ignore_t,
) -> str:
    # skip the patient id column (it was already checked anyway)
    if col_idx == detail.patient_id_idx:
        return detail.patient_id

    # get the cell value. replace it with None if it's just whitespace
    if isinstance(cell_value, str):
        cell_value = cell_value.strip()

    if cell_value is None:
        return None

    col_name = detail.col_names[col_idx]

    # with no column name; this column of the row should have been None
    if not col_name:
        raise BlankColumnHasData(detail.page_name, detail.row_idx, col_idx, cell_value)

    # choose what to do based on the column/cell type
    match detail.column_type(shift_ignore, col_name, cell_value):
        case ColumnType.IGNORE:
            # "skip" is sheet-wide, but, ignore is for a single cell of data
            # ... so we're checking for the ignore command here
            pass
        case ColumnType.TEXT:
            # check for dates in non-date columns
            for found in datefinder.find_dates(str(cell_value)):
                logger.debug(
                    f"found hidden date {detail.page_name=}/{col_name=}"
                    + f":{cell_value}"
                )
                raise HiddenDate(
                    detail.page_name,
                    detail.row_idx,
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
                    detail.page_name, detail.row_idx, col_idx, col_name, cell_value
                )

            # perform the actual shifting
            shifted = parsed + shift_delta

            # TODO; allow column/sheet/workbook format settings
            if isinstance(cell_value, datetime):
                return cast(Any, shifted.to_pydatetime())
            else:
                return cast(Any, shifted.to_pydatetime().date())

    return cell_value

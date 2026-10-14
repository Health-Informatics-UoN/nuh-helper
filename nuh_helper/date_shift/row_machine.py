"""implements a Finite State Machine to handle calls from the streaming.

While the book_page API is meant to be generic, this is certainly not


"""

import logging
from collections.abc import Generator as gen
from dataclasses import dataclass, replace
from enum import StrEnum
from pathlib import Path
from typing import Literal, cast

import datefinder
from book_page import Book

from nuh_helper.date_shift import _parse
from nuh_helper.date_shift.defs import (
    DateColumnsMissing,
    ExtraColumns,
    HiddenDate,
    PatientColumnMissing,
    TextColumnsMissing,
    normalize_inplace,
    sheet_configs_t,
)

logger = logging.getLogger(__name__)


class ColumnType(StrEnum):
    TEXT = "text"
    DATE = "date"
    IGNORE = "skip"


@dataclass
class PatientRow:
    """carries the data needed to process a row from the sheet config and header_row"""

    page_name: str
    col_names: tuple[str, ...]
    patient_id_idx: int
    patient_id_col: str
    patient_id: str
    date_columns: tuple[str, ...]
    text_columns: tuple[str, ...]
    row_idx: int
    row_cells: list[str]

    def column_type(
        self,
        shift_ignore: dict[str, dict[str, set[None | str]]],
        col_name: str,
        cell_value: str,
    ) -> bool:
        """determine what we should do for a given column"""

        ignored: set[None | str] = shift_ignore.get(self.page_name, {}).get(
            col_name, {None, ""}
        )

        return (
            ColumnType.IGNORE
            if cell_value in ignored
            else ColumnType.TEXT
            if col_name in self.text_columns
            else ColumnType.DATE
        )


class PatientRowStateMachine:
    """builds upon the callbacks from book_page to handle row/page switching. does all
    the checking for configs and what page we're on. needed because both patient IDs and
    The Big Loop need the same checks"""

    def __init__(self, sheet_configs: sheet_configs_t) -> None:
        self._sheet_configs: sheet_configs_t = sheet_configs
        self._header_row: None | int = None
        self._skip_rows: None | list[int] = None
        self._detail_archetype: None | PatientRow = None

    def next_details(
        self, sheet_name: str, row_idx: int, row_cells: list
    ) -> PatientRow | Literal[False]:
        """need to call this every step of the loop. it returns False to indicate that
        this row doesn't need further processing, or a page-description that includes
        the located patient id"""

        row_cells = normalize_inplace(row_cells)

        sheet_config: dict | str = self._sheet_configs[sheet_name]

        if row_idx == 0:
            ###
            # process the "which row is header and which do we skip" configuration
            ##

            self._detail_archetype = None

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

            # detect extra columns in this sheet
            known = [patient_id_col] + date_columns + text_columns
            if extra_columns := [col for col in row_cells if col and col not in known]:
                raise ExtraColumns(sheet_name, extra_columns)

            # find patient_id_idx and check columns
            patient_id_idx = None
            for col_idx, col_name in enumerate(row_cells):
                if not col_name:
                    continue

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

            if missing := [col for col in text_columns if col not in row_cells]:
                raise TextColumnsMissing(sheet_name, missing)

            if missing := [col for col in date_columns if col not in row_cells]:
                raise DateColumnsMissing(sheet_name, missing)

            # we've completed the header_row; save the archetype of the page detail
            self._detail_archetype = PatientRow(
                page_name=sheet_name,
                col_names=row_cells,
                patient_id_idx=patient_id_idx,
                patient_id_col=patient_id_col,
                date_columns=tuple(date_columns),
                text_columns=tuple(text_columns),
                # these three need to be replaced
                patient_id="<archetype>",
                row_idx=-1,
                row_cells=[],
            )

            # we don't shift anything in the header_row
            return False

        # get the patient id for this row
        patient_id = row_cells[self._detail_archetype.patient_id_idx]

        # check for date in the patient_id
        for found in datefinder.find_dates(str(patient_id)):
            raise HiddenDate(
                sheet_name,
                row_idx,
                self._detail_archetype.patient_id_idx,
                self._detail_archetype.patient_id_col,
                patient_id,
                found,
            )

        patient_id = _parse._normalize_patient_id(patient_id)
        if patient_id is None:
            # if the pid is None; the rest of the row should be None as well
            if non_blank := [
                cell
                for cell in row_cells
                # keep the cells that aren't blank
                if cell is not None and str(cell).strip() != ""
            ]:
                raise ValueError(
                    f"{row_idx=} has no pid, should be blank but has {non_blank}"
                )

            # we've completed the checks for a None pid
            # ... so we need to skip the rest of this loop
            return False

        # it's a valid pid - return a page detail object
        # ... so that we don't need to examine this scanner
        return replace(
            self._detail_archetype,
            patient_id=patient_id,
            row_idx=row_idx,
            row_cells=row_cells,
        )

    def book_copy(scanner, book: Book, output_file: Path) -> gen[PatientRow]:
        for sheet_name, row_idx, row_cells in book.stream_copy(output_file):
            # check if we're in a patient row, and, yield the row details if so
            detail = scanner.next_details(sheet_name, row_idx, row_cells)
            if detail:
                yield detail

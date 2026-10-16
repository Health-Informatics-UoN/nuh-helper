"""Shift mapping generation and loading for date-shift."""

import csv
import logging
import random
from pathlib import Path

import pandas as pd
from pandas import Timedelta

from nuh_helper.book_page import BaseWorkBook
from nuh_helper.date_shift import _parse
from nuh_helper.date_shift.defs import (
    sheet_configs_t,
)
from nuh_helper.date_shift.row_machine import (
    PatientRowStateMachine,
)

logger = logging.getLogger(__name__)


def load_shift_deltas(
    book: BaseWorkBook,
    patient_sheet: str,
    sheet_configs: sheet_configs_t,
    min_shift_days: int,
    max_shift_days: int,
    old_linking_table: str | Path | None,
    out_linking_table: str | Path | None,
    seed: int | None,
) -> dict[str, Timedelta]:
    """load the day-offsets for patients.

    Args
        - book: Bool = the CSV or XLSX workbook to read the patient ids from
        - patient_sheet: str = the authoritatively page name with a list of patients
            * we assert that it exists in the book
            * rows don't need to be unique
        - sheet_configs: the full sheet configs
        - min_shift_days = dates are shifted up to this far back (probably negative)
        - max_shift_days = dates can be shifted at most this forward (should be > min)
        - linking_table_path: str | None = None,
        - linking_table_output: str | None = None,
        - seed = an optional seed to pass to the shift generation

    Returns:
        dictionary mapping 'patient_id' to a pd.Timedelta 'shift_days' which we use.

    """

    if not old_linking_table:
        old_linking_table = None
    elif not isinstance(old_linking_table, Path):
        old_linking_table = Path(old_linking_table)

    if not out_linking_table:
        out_linking_table = None
    elif not isinstance(out_linking_table, Path):
        out_linking_table = Path(out_linking_table)

    with book[patient_sheet] as page:
        scanner = PatientRowStateMachine(sheet_configs)
        patient_ids: list[str] = [
            pid.patient_id
            for pid in [
                scanner.next_details(page.name, row, data)
                for row, data in page.stream_rows()
            ]
            if pid
        ]

    # load the old ids (should we remove IDs that are "gone"?)
    if not old_linking_table or not old_linking_table.is_file():
        old_mappings: dict[str, int] = {}
    else:
        with open(old_linking_table) as file:
            old_mappings: dict[str, int] = {
                _parse._normalize_patient_id(row[0]): int(row[1])
                for row in csv.reader(file)
                if row and row[1] and row[1].strip() != "shift_days"
            }

    if seed:
        random.seed(seed)

    # create the list of all mappings
    all_mappings: dict[str, int] = {
        id: (
            old_mappings[id]
            if id in old_mappings
            else random.randint(min_shift_days, max_shift_days)
        )
        for id in patient_ids
    }
    removed = len([k for k in all_mappings if k not in patient_ids])
    logger.info(
        f"patient_shift_deltas() ; {len(all_mappings)=} patient(s) (... and {removed=})"
    )

    # store all mappings in the out file
    if out_linking_table:
        out_linking_table.parent.mkdir(parents=True, exist_ok=True)
        with open(out_linking_table, "w") as file:
            csv.writer(file).writerows(
                [["patient_id", "shift_days"]]
                + [[key, all_mappings[key]] for key in sorted(all_mappings)]
            )

    return {k: pd.Timedelta(days=all_mappings[k]) for k in all_mappings}

"""Shift mapping generation and loading for date-shift."""

import csv
import logging
import random
import time
from pathlib import Path

import pandas as pd
from book_page import Page

from nuh_helper.date_shift import _parse

logger = logging.getLogger(__name__)


def generate_shift_mappings(
    patient_ids: list[str],
    min_shift_days: int = -15,
    max_shift_days: int = 15,
    seed: int | None = None,
) -> pd.DataFrame:
    """
    Generate random shift mappings for patient IDs.

    Args:
        patient_ids: List of patient IDs to generate shifts for.
        min_shift_days: Minimum number of days to shift (default: -15).
        max_shift_days: Maximum number of days to shift (default: 15).
        seed: Optional random seed for reproducibility.

    Returns:
        DataFrame with columns 'patient_id' and 'shift_days'.
    """
    if seed is not None:
        random.seed(seed)

    shifts = [random.randint(min_shift_days, max_shift_days) for _ in patient_ids]
    return pd.DataFrame({"patient_id": patient_ids, "shift_days": shifts})


def load_shift_mappings(csv_path: str) -> pd.DataFrame:
    """
    Load shift mappings from a CSV file.

    Args:
        csv_path: Path to the CSV file containing shift mappings.
                  Expected columns: 'patient_id' and 'shift_days'.

    Returns:
        DataFrame with shift mappings.
    """
    df = pd.read_csv(csv_path)
    if "patient_id" not in df.columns or "shift_days" not in df.columns:
        raise ValueError("CSV must contain 'patient_id' and 'shift_days' columns")
    logger.info("Loaded %d shift mapping(s) from '%s'", len(df), csv_path)
    return df


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


def patient_shift_deltas(
    page: Page,
    src_linking_table: str | Path,
    out_linking_table: str | Path,
    patient_id_col_name: str,
    patient_header_row: int,
    patient_skip_rows: list[int] | None,
    seed: int | None,
    min_shift_days: int,
    max_shift_days: int,
) -> dict[str, pd.Timedelta]:

    logger.info("patient_shift_deltas() ; preparing per-patient shifts")

    # find the patient id index
    patient_id_col: None | int = None
    patient_ids = []
    start = time.time()

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
            if (len(patient_ids) % 40) == 0:
                seconds = time.time() - start
                logger.info(
                    f"patient_shift_deltas() ; so far {len(patient_ids)=} in {seconds=}"
                )
    logger.info(f"patient_shift_deltas() ; Found {len(patient_ids)=} in {page.name=}")

    # load the old ids (should we remove IDs that are "gone"?)
    if isinstance(src_linking_table, str):
        src_linking_table = Path(src_linking_table)
    if not src_linking_table.is_file():
        old_mappings: dict[str, int] = {}
    else:
        with open(src_linking_table) as file:
            old_mappings: dict[str, int] = {
                _parse._normalize_patient_id(row[0]): int(row[1])
                for row in csv.reader(file)
                if row and row[1] and row[1].strip() != "shift_days"
                # if row and row[0] and row[0].strip() != "patient_id"
            }
    logger.info(
        f"patient_shift_deltas() ; Loaded {len(old_mappings)=} in {src_linking_table =}"
    )

    # create the list of all mappings

    if seed:
        random.seed(seed)
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
    with open(out_linking_table, "w") as file:
        csv.writer(file).writerows(
            [["patient_id", "shift_days"]]
            + [[key, all_mappings[key]] for key in sorted(all_mappings)]
        )
    logger.info(f"patient_shift_deltas() ; Saved mappings {out_linking_table}")

    return {k: pd.Timedelta(days=all_mappings[k]) for k in all_mappings}

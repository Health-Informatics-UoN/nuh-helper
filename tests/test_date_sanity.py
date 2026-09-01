from pathlib import Path

import pytest

from nuh_helper import shift_excel_dates_inplace
from nuh_helper.date_shift import (
    DateTooFarAhead,
    DateTooFarBack,
)


def test_too_far_ahead(tmp_path: Path) -> None:

    source_file = Path(__file__).parent / "data/date_sanity/too_far_ahead.xlsx"
    output_path = tmp_path / "target.xlsx"
    linking_table_old = tmp_path / "linking_table_old.csv"
    linking_table_out = tmp_path / "linking_table_out.csv"

    sheet_configs = {
        "tofarr": {
            "patient_id_col": "patient",
            "date_columns": [
                "dobirth",
            ],
            "header_row": 0,
            "skip_rows_after_header": [],
        },
    }
    with pytest.raises(DateTooFarAhead) as info:
        shift_excel_dates_inplace(
            input_file=str(source_file),
            output_file=str(output_path),
            patient_sheet="tofarr",
            patient_id_col="patient",
            sheet_configs=sheet_configs,
            min_shift_days=-20,
            max_shift_days=-1,
            seed=14333,
            linking_table_path=str(linking_table_old),
            linking_table_output=str(linking_table_out),
        )

    assert (
        info.value._message == "the date 2072-12-10 00:00:00 is too far in the future"
    )


def test_too_far_back(tmp_path: Path) -> None:

    source_file = Path(__file__).parent / "data/date_sanity/too_far_back.xlsx"
    output_path = tmp_path / "target.xlsx"
    linking_table_old = tmp_path / "linking_table_old.csv"
    linking_table_out = tmp_path / "linking_table_out.csv"

    sheet_configs = {
        "Sheet1": {
            "patient_id_col": "patient",
            "date_columns": [
                "dob",
            ],
            "header_row": 0,
            "skip_rows_after_header": [],
        },
    }
    with pytest.raises(DateTooFarBack) as info:
        shift_excel_dates_inplace(
            input_file=str(source_file),
            output_file=str(output_path),
            patient_sheet="Sheet1",
            patient_id_col="patient",
            sheet_configs=sheet_configs,
            min_shift_days=-20,
            max_shift_days=-1,
            seed=14333,
            linking_table_path=str(linking_table_old),
            linking_table_output=str(linking_table_out),
        )

    assert info.value._message == "the date 1007-02-05 00:00:00 is too far in the past"

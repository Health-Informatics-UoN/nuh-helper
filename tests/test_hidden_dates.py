from pathlib import Path

import pytest

from nuh_helper import shift_excel_dates_inplace
from nuh_helper.date_shift import (
    HiddenDate,
)


def test_iso8601(tmp_path: Path) -> None:

    source_file = Path(__file__).parent / "data/hidden_dates/iso8601.xlsx"
    output_path = tmp_path / "target.xlsx"
    linking_table_old = tmp_path / "linking_table_old.csv"
    linking_table_out = tmp_path / "linking_table_out.csv"

    sheet_configs = {
        "args": {
            "patient_id_col": "ptid",
            "date_columns": [
                "dob",
            ],
            "header_row": 0,
            "skip_rows_after_header": [],
        },
    }
    with pytest.raises(HiddenDate) as info:
        shift_excel_dates_inplace(
            input_file=str(source_file),
            output_file=str(output_path),
            patient_sheet="args",
            patient_id_col="ptid",
            sheet_configs=sheet_configs,
            min_shift_days=-20,
            max_shift_days=-1,
            seed=14333,
            linking_table_path=str(linking_table_old),
            linking_table_output=str(linking_table_out),
        )

    assert info.value._message == "foobar"


def test_us_date(tmp_path: Path) -> None:

    source_file = Path(__file__).parent / "data/hidden_dates/us_date.xlsx"
    output_path = tmp_path / "target.xlsx"
    linking_table_old = tmp_path / "linking_table_old.csv"
    linking_table_out = tmp_path / "linking_table_out.csv"

    sheet_configs = {
        "data": {
            "patient_id_col": "ptid",
            "date_columns": [
                "dob",
            ],
            "header_row": 0,
            "skip_rows_after_header": [],
        },
    }
    with pytest.raises(HiddenDate) as info:
        shift_excel_dates_inplace(
            input_file=str(source_file),
            output_file=str(output_path),
            patient_sheet="data",
            patient_id_col="ptid",
            sheet_configs=sheet_configs,
            min_shift_days=-20,
            max_shift_days=-1,
            seed=14333,
            linking_table_path=str(linking_table_old),
            linking_table_output=str(linking_table_out),
        )

    assert info.value._message == "foobar"


def test_written(tmp_path: Path) -> None:

    source_file = Path(__file__).parent / "data/hidden_dates/written.xlsx"
    output_path = tmp_path / "target.xlsx"
    linking_table_old = tmp_path / "linking_table_old.csv"
    linking_table_out = tmp_path / "linking_table_out.csv"

    sheet_configs = {
        "yeah": {
            "patient_id_col": "ptid",
            "date_columns": [
                "dob",
            ],
            "header_row": 0,
            "skip_rows_after_header": [],
        },
    }
    with pytest.raises(HiddenDate) as info:
        shift_excel_dates_inplace(
            input_file=str(source_file),
            output_file=str(output_path),
            patient_sheet="yeah",
            patient_id_col="ptid",
            sheet_configs=sheet_configs,
            min_shift_days=-20,
            max_shift_days=-1,
            seed=14333,
            linking_table_path=str(linking_table_old),
            linking_table_output=str(linking_table_out),
        )

    assert info.value._message == "foobar"

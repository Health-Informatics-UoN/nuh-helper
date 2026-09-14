from pathlib import Path

import pytest

from nuh_helper import shift_excel_dates_inplace
from nuh_helper.date_shift import UnknownPatient


def test_missing_id(tmp_path: Path) -> None:

    sheet_configs = {
        "patients": {
            "patient_id_col": "ptid",
            "header_row": 1,
            "skip_rows_after_header": [2],
            "date_columns": ["dob"],
        },
        "deceased": {
            "patient_id_col": "patient_id",
            "header_row": 2,
            "skip_rows_after_header": [3, 4],
            "date_columns": [
                "deaddat",
                "diagdat",
            ],
        },
    }

    source_file = str(Path(__file__).parent / "data/missing-id.xlsx")
    output_path = tmp_path / "target.xlsx"
    linking_table_old = tmp_path / "linking_table_old.csv"
    linking_table_out = tmp_path / "linking_table_out.csv"

    with pytest.raises(UnknownPatient) as info:
        shift_excel_dates_inplace(
            input_file=str(source_file),
            output_file=str(output_path),
            patient_sheet="patients",
            patient_id_col="ptid",
            sheet_configs=sheet_configs,
            min_shift_days=-20,
            max_shift_days=-1,
            seed=14333,
            linking_table_path=str(linking_table_old),
            linking_table_output=str(linking_table_out),
        )

    assert str(info.value._message) == "Unknown id='nuh006' on page='deceased'"

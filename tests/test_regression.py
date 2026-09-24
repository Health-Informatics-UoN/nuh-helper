from pathlib import Path

import pytest

from nuh_helper.date_shift import (
    shift_excel_dates_inplace,
)


def test_carrot_dates(tmp_path: Path) -> None:
    """test shifting a sheet from carrot"""

    source = Path(__file__).parent / "data/regression/carrot_dates.xlsx"
    target = tmp_path / source.name
    offsets_old = tmp_path / "user-offsets.old.csv"
    offsets_new = tmp_path / "user-offsets.new.csv"

    shift_excel_dates_inplace(
        input_file=str(source),
        output_file=str(target),
        patient_sheet="observation",
        patient_id_col="person_id",
        sheet_configs={
            "observation": {
                "patient_id_col": "person_id",
                "header_row": 0,
                "date_columns": ["observation_date", "observation_datetime"],
                "text_columns": [
                    "observation_id",
                    "observation_concept_id",
                    "observation_type_concept_id",
                    "qualifier_source_value",
                ],
            },
        },
        min_shift_days=-20,
        max_shift_days=-1,
        seed=14333,
        linking_table_path=offsets_old,
        linking_table_output=offsets_new,
    )

    pytest.fail("???")

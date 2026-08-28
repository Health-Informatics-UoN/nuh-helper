from pathlib import Path

import pytest

from nuh_helper import shift_excel_dates_inplace
from nuh_helper.date_shift import (
    DateColumnMissing,
    ExtraColumn,
    ExtraPage,
    IgnoredColumnMissing,
    PageMissing,
)


def test_column_missing(tmp_path: Path) -> None:

    source_file = Path(__file__).parent / "data/structural/column_missing.xlsx"
    output_path = tmp_path / "target.xlsx"
    linking_table_old = tmp_path / "linking_table_old.csv"
    linking_table_out = tmp_path / "linking_table_out.csv"

    sheet_configs = {
        "miss": {
            "patient_id_col": "ptid",
            "date_columns": [
                "dob",
                "a-missing-column",
            ],
            "header_row": 0,
            "skip_rows_after_header": [],
        },
    }
    with pytest.raises(DateColumnMissing) as info:
        shift_excel_dates_inplace(
            input_file=str(source_file),
            output_file=str(output_path),
            patient_sheet="miss",
            patient_id_col="ptid",
            sheet_configs=sheet_configs,
            min_shift_days=-20,
            max_shift_days=-1,
            seed=14333,
            linking_table_path=str(linking_table_old),
            linking_table_output=str(linking_table_out),
        )

    assert info.value._column_name == "a-missing-column"


def test_extra_column(tmp_path: Path) -> None:
    """test that should fail because there's an unexpected column not the config"""

    source_file = Path(__file__).parent / "data/structural/extra_column.xlsx"
    output_path = tmp_path / "target.xlsx"
    linking_table_old = tmp_path / "linking_table_old.csv"
    linking_table_out = tmp_path / "linking_table_out.csv"

    sheet_configs = {
        "people": {
            "patient_id_col": "ptid",
            "date_columns": ["dob"],
            "header_row": 1,
            "skip_rows_after_header": [],
        },
        "events": {
            "patient_id_col": "patient",
            "date_columns": ["date"],
            "header_row": 0,
            "skip_rows_after_header": [],
        },
    }
    with pytest.raises(ExtraColumn) as info:
        shift_excel_dates_inplace(
            input_file=str(source_file),
            output_file=str(output_path),
            patient_sheet="people",
            patient_id_col="ptid",
            sheet_configs=sheet_configs,
            min_shift_days=-20,
            max_shift_days=-1,
            seed=14333,
            linking_table_path=str(linking_table_old),
            linking_table_output=str(linking_table_out),
        )

    assert info.value._page_name == "an-extra-column"


def test_ignored_column_missing(tmp_path: Path) -> None:

    source_file = Path(__file__).parent / "data/structural/ignored_column_missing.xlsx"
    output_path = tmp_path / "target.xlsx"
    linking_table_old = tmp_path / "linking_table_old.csv"
    linking_table_out = tmp_path / "linking_table_out.csv"

    sheet_configs = {
        "people": {
            "patient_id_col": "ptid",
            "date_columns": ["dob"],
            "header_row": 1,
            "skip_rows_after_header": [],
        },
        "events": {
            "patient_id_col": "patient",
            "date_columns": ["date"],
            "header_row": 0,
            "skip_rows_after_header": [],
        },
    }
    with pytest.raises(IgnoredColumnMissing) as info:
        shift_excel_dates_inplace(
            input_file=str(source_file),
            output_file=str(output_path),
            patient_sheet="people",
            patient_id_col="ptid",
            sheet_configs=sheet_configs,
            min_shift_days=-20,
            max_shift_days=-1,
            seed=14333,
            linking_table_path=str(linking_table_old),
            linking_table_output=str(linking_table_out),
        )

    assert info.value._column_name == "a-missing-column"


def test_page_missing(tmp_path: Path) -> None:

    source_file = Path(__file__).parent / "data/structural/page_missing.xlsx"
    output_path = tmp_path / "target.xlsx"
    linking_table_old = tmp_path / "linking_table_old.csv"
    linking_table_out = tmp_path / "linking_table_out.csv"

    sheet_configs = {
        "people": {
            "patient_id_col": "ptid",
            "date_columns": ["dob"],
            "header_row": 1,
            "skip_rows_after_header": [],
        },
        "events": {
            "patient_id_col": "patient",
            "date_columns": ["date"],
            "header_row": 0,
            "skip_rows_after_header": [],
        },
    }
    with pytest.raises(PageMissing) as info:
        shift_excel_dates_inplace(
            input_file=str(source_file),
            output_file=str(output_path),
            patient_sheet="people",
            patient_id_col="ptid",
            sheet_configs=sheet_configs,
            min_shift_days=-20,
            max_shift_days=-1,
            seed=14333,
            linking_table_path=str(linking_table_old),
            linking_table_output=str(linking_table_out),
        )

    assert info.value._page_name == "a-missing-page"


def test_page_extra(tmp_path: Path) -> None:
    """test that should fail because there's a page in the doc but not the config"""

    source_file = Path(__file__).parent / "data/structural/page_extra.xlsx"
    output_path = tmp_path / "target.xlsx"
    linking_table_old = tmp_path / "linking_table_old.csv"
    linking_table_out = tmp_path / "linking_table_out.csv"

    sheet_configs = {
        "people": {
            "patient_id_col": "ptid",
            "date_columns": ["dob"],
            "header_row": 1,
            "skip_rows_after_header": [],
        },
        "events": {
            "patient_id_col": "patient",
            "date_columns": ["date"],
            "header_row": 0,
            "skip_rows_after_header": [],
        },
    }
    with pytest.raises(ExtraPage) as info:
        shift_excel_dates_inplace(
            input_file=str(source_file),
            output_file=str(output_path),
            patient_sheet="people",
            patient_id_col="ptid",
            sheet_configs=sheet_configs,
            min_shift_days=-20,
            max_shift_days=-1,
            seed=14333,
            linking_table_path=str(linking_table_old),
            linking_table_output=str(linking_table_out),
        )

    assert info.value._page_name == "an-extra-page"

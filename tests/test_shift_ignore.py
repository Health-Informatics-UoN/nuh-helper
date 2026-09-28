from pathlib import Path

import pytest
from openpyxl import load_workbook

from nuh_helper import shift_excel_dates_inplace
from nuh_helper.date_shift import HiddenDate, ShiftFoundNonDate


@pytest.mark.parametrize("shift_ignore", [True, False])
def test_ignore_in_date_columns(shift_ignore: bool, tmp_path: Path) -> None:

    source_file = Path(__file__).parent / "data/shift_ignore/workbook.xlsx"
    output_path = tmp_path / "target.xlsx"
    linking_table_old = Path(__file__).parent / "data/shift_ignore/offsets.csv"
    linking_table_out = tmp_path / "linking_table_out.csv"

    sheet_configs = {
        "page-desc": "skip",
        "page-data": {
            "patient_id_col": "pid",
            "date_columns": ["dob"],
            "text_columns": ["top"],
            "header_row": 1,
            "skip_rows_after_header": [],
        },
    }

    def body() -> None:
        shift_excel_dates_inplace(
            input_file=str(source_file),
            output_file=str(output_path),
            patient_sheet="page-data",
            patient_id_col="pid",
            sheet_configs=sheet_configs,
            min_shift_days=-20,
            max_shift_days=-1,
            seed=14333,
            linking_table_path=str(linking_table_old),
            linking_table_output=str(linking_table_out),
        )

    if not shift_ignore:
        with pytest.raises(ShiftFoundNonDate) as info:
            body()
        assert str(info.value) == "page='page-data'[6, 2 @ col_name='dob'] val='mssing'"
        return
    else:
        sheet_configs["page-data"]["shift_ignore"] = {"dob": ["mssing"]}
        body()

    workbook = load_workbook(output_path)

    worksheet = workbook.worksheets[1]

    # first column
    assert worksheet.cell(1, 1).value == "personal id"
    assert worksheet.cell(2, 1).value == "pid"
    assert worksheet.cell(3, 1).value == "nuh71"
    assert worksheet.cell(4, 1).value == "nuh06"
    assert worksheet.cell(5, 1).value == "nuh23"
    assert worksheet.cell(6, 1).value == "nuh67"
    assert worksheet.cell(7, 1).value == "nuh27"

    # last column
    assert worksheet.cell(1, 3).value == "pizza topping"
    assert worksheet.cell(2, 3).value == "top"
    assert worksheet.cell(3, 3).value == "cheese"
    assert worksheet.cell(4, 3).value == "unknown"
    assert worksheet.cell(5, 3).value == "mushrooms"
    assert worksheet.cell(6, 3).value == "this can't be a date - sorry 2016"
    assert worksheet.cell(7, 3).value == "idk - this can't be a date anymore"

    # the important column to check - the dates
    assert worksheet.cell(1, 2).value == "birthday"
    assert worksheet.cell(2, 2).value == "dob"
    assert str(worksheet.cell(3, 2).value) == "2001-12-17 00:00:00"
    assert str(worksheet.cell(4, 2).value) == "1993-09-20 00:00:00"
    assert worksheet.cell(5, 2).value is None
    assert str(worksheet.cell(6, 2).value) == "mssing"  # change that's under test
    assert str(worksheet.cell(7, 2).value) == "1999-11-30 00:00:00"


@pytest.mark.parametrize("shift_ignore", [True, False])
def test_ignore_in_text_columns(shift_ignore: bool, tmp_path: Path) -> None:

    source_file = Path(__file__).parent / "data/shift_ignore/workbook.xlsx"
    output_path = tmp_path / "target.xlsx"
    linking_table_old = Path(__file__).parent / "data/shift_ignore/offsets.csv"
    linking_table_out = tmp_path / "linking_table_out.csv"

    sheet_configs = {
        "page-desc": {
            "patient_id_col": "pid",
            "date_columns": [],
            "text_columns": ["foo", "glitter"],
            "header_row": 0,
            "skip_rows_after_header": [],
        },
        "page-data": "skip",
    }

    def body() -> None:
        shift_excel_dates_inplace(
            input_file=str(source_file),
            output_file=str(output_path),
            patient_sheet="page-desc",
            patient_id_col="pid",
            sheet_configs=sheet_configs,
            min_shift_days=-20,
            max_shift_days=-1,
            seed=14333,
            linking_table_path=str(linking_table_old),
            linking_table_output=str(linking_table_out),
        )

    if not shift_ignore:
        print(output_path)
        with pytest.raises(HiddenDate) as info:
            body()
        assert str(info.value._message) == (
            "hidden date in [sheet_name='page-desc', 3, 3 @ glitter] "
            + 'value="can\'t recall the date but on 12/11/2001 they had an itchy tummy"'
            + " // found=datetime.datetime(2001, 12, 11, 0, 0)"
        )
        return

    else:
        # add the exception
        sheet_configs["page-desc"]["shift_ignore"] = {
            "glitter": [
                "can't recall the date but on 12/11/2001 they had an itchy tummy",
            ],
            "foo": [
                "agreed on 21st jul 2017",
            ],
        }

        # run the shift
        body()

        workbook = load_workbook(output_path)
        page = workbook["page-desc"]

        obtained = [
            [str(page.cell(row + 1, col + 1).value) for col in range(page.max_column)]
            for row in range(page.max_row)
        ]

        expected = [
            ["foo", "pid", "glitter"],
            ["bar", "nuh71", "road"],
            [
                "shoe",
                "nuh06",
                "can't recall the date but on 12/11/2001 they had an itchy tummy",
            ],
            ["agreed on 21st jul 2017", "nuh23", "grip"],
            ["farm", "nuh67", "grim"],
            ["cake", "nuh27", "cheese"],
        ]

        assert expected == obtained

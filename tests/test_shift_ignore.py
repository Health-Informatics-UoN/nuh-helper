from pathlib import Path

import book_page
import pytest

from nuh_helper import shift_excel_dates_inplace
from nuh_helper.date_shift import HiddenDate, PatientColumnMissing, ShiftFoundNonDate

test_data = Path(__file__).parent / "data/shift_ignore"


@pytest.mark.parametrize("use_csv", [False, True])
def test_shift_exceptions(use_csv: bool, tmp_path: Path) -> None:
    """check that misconfigurations cause an error"""

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
            "shift_exceptions": {"dob": ["mssing"]},
        },
    }

    source_file, output_path = src_and_out_names(use_csv, tmp_path, sheet_configs)

    with pytest.raises(RuntimeError) as info:
        shift_excel_dates_inplace(
            input_file=source_file,
            output_file=output_path,
            patient_sheet="page-data",
            sheet_configs=sheet_configs,
            min_shift_days=-20,
            max_shift_days=-1,
            seed=14333,
            linking_table_path=str(linking_table_old),
            linking_table_output=str(linking_table_out),
        )
    assert str(info.value) == (
        "shift_exceptions was removed, update sheet_name='page-data'"
    )


@pytest.mark.parametrize("use_csv", [False, True])
def test_ptid_in_date_columns(use_csv: bool, tmp_path: Path) -> None:
    """check that misconfigurations cause an error"""

    linking_table_old = Path(__file__).parent / "data/shift_ignore/offsets.csv"
    linking_table_out = tmp_path / "linking_table_out.csv"

    sheet_configs = {
        "page-desc": "skip",
        "page-data": {
            "patient_id_col": "pid",
            "date_columns": ["dob", "pid"],
            "text_columns": ["top"],
            "header_row": 1,
            "skip_rows_after_header": [],
            "shift_ignore": {"dob": ["mssing"]},
        },
    }

    source_file, output_path = src_and_out_names(use_csv, tmp_path, sheet_configs)

    with pytest.raises(ValueError) as info:
        shift_excel_dates_inplace(
            input_file=source_file,
            output_file=str(output_path),
            patient_sheet="page-data",
            sheet_configs=sheet_configs,
            min_shift_days=-20,
            max_shift_days=-1,
            seed=14333,
            linking_table_path=str(linking_table_old),
            linking_table_output=str(linking_table_out),
        )
    assert str(info.value) == (
        "patient_id_col='pid' shouldn't be in date_columns of sheet_name='page-data'"
    )


@pytest.mark.parametrize("use_csv", [False, True])
def test_ptid_in_text_columns(use_csv: bool, tmp_path: Path) -> None:
    """check that misconfigurations cause an error"""

    linking_table_old = Path(__file__).parent / "data/shift_ignore/offsets.csv"
    linking_table_out = tmp_path / "linking_table_out.csv"

    sheet_configs = {
        "page-desc": "skip",
        "page-data": {
            "patient_id_col": "pid",
            "date_columns": ["dob"],
            "text_columns": ["top", "pid"],
            "header_row": 1,
            "skip_rows_after_header": [],
            "shift_ignore": {"dob": ["mssing"]},
        },
    }

    source_file, output_path = src_and_out_names(use_csv, tmp_path, sheet_configs)

    with pytest.raises(ValueError) as info:
        shift_excel_dates_inplace(
            input_file=source_file,
            output_file=str(output_path),
            patient_sheet="page-data",
            sheet_configs=sheet_configs,
            min_shift_days=-20,
            max_shift_days=-1,
            seed=14333,
            linking_table_path=str(linking_table_old),
            linking_table_output=str(linking_table_out),
        )
    assert str(info.value) == (
        "patient_id_col='pid' shouldn't be in text_columns of sheet_name='page-data'"
    )


@pytest.mark.parametrize("use_csv", [False, True])
def test_columns_overlap(use_csv: bool, tmp_path: Path) -> None:
    """check that misconfigurations cause an error"""

    linking_table_old = Path(__file__).parent / "data/shift_ignore/offsets.csv"
    linking_table_out = tmp_path / "linking_table_out.csv"

    sheet_configs = {
        "page-desc": "skip",
        "page-data": {
            "patient_id_col": "pid",
            "date_columns": ["dob"],
            "text_columns": ["top", "dob"],
            "header_row": 1,
            "skip_rows_after_header": [],
            "shift_ignore": {"dob": ["mssing"]},
        },
    }

    source_file, output_path = src_and_out_names(use_csv, tmp_path, sheet_configs)

    with pytest.raises(ValueError) as info:
        shift_excel_dates_inplace(
            input_file=source_file,
            output_file=str(output_path),
            patient_sheet="page-data",
            sheet_configs=sheet_configs,
            min_shift_days=-20,
            max_shift_days=-1,
            seed=14333,
            linking_table_path=str(linking_table_old),
            linking_table_output=str(linking_table_out),
        )
    assert str(info.value) == (
        "sheet_name='page-data' has the some columns in both date and text "
        + "col_names=['dob']"
    )


@pytest.mark.parametrize("use_csv", [False, True])
def test_ptid_column_missing(use_csv: bool, tmp_path: Path) -> None:
    """check that misconfigurations cause an error"""

    linking_table_old = Path(__file__).parent / "data/shift_ignore/offsets.csv"
    linking_table_out = tmp_path / "linking_table_out.csv"

    sheet_configs = {
        "page-desc": "skip",
        "page-data": {
            "patient_id_col": "ptid",
            "date_columns": ["dob"],
            "text_columns": ["top", "pid"],
            "header_row": 1,
            "skip_rows_after_header": [],
            "shift_ignore": {"dob": ["mssing"]},
        },
    }

    source_file, output_path = src_and_out_names(use_csv, tmp_path, sheet_configs)

    with pytest.raises(PatientColumnMissing) as info:
        shift_excel_dates_inplace(
            input_file=source_file,
            output_file=str(output_path),
            patient_sheet="page-data",
            sheet_configs=sheet_configs,
            min_shift_days=-20,
            max_shift_days=-1,
            seed=14333,
            linking_table_path=str(linking_table_old),
            linking_table_output=str(linking_table_out),
        )
    assert info.value._page_name == "page-data"
    assert info.value._column_name == "ptid"


@pytest.mark.parametrize("use_csv", [False, True])
@pytest.mark.parametrize("shift_ignore", [True, False])
def test_ignore_in_date_columns(
    use_csv: bool, shift_ignore: bool, tmp_path: Path
) -> None:

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

    source_file, output_path = src_and_out_names(use_csv, tmp_path, sheet_configs)

    def body(shift_ignore_yaml: None | Path) -> None:
        shift_excel_dates_inplace(
            input_file=source_file,
            output_file=str(output_path),
            patient_sheet="page-data",
            sheet_configs=sheet_configs,
            min_shift_days=-20,
            max_shift_days=-1,
            seed=14333,
            linking_table_path=str(linking_table_old),
            linking_table_output=str(linking_table_out),
            shift_ignore_yaml=shift_ignore_yaml,
        )

    if not shift_ignore:
        with pytest.raises(ShiftFoundNonDate) as info:
            body(None)
        assert info.value._page == "page-data"
        assert info.value._row == 5
        assert info.value._col == 1
        assert info.value._col_name == "dob"
        assert info.value._value == "mssing"
        return
    else:
        body(test_data / "shift-ignore.yaml")

        obtained = load_workbook(output_path)["page-data"]
        expected = [
            ["personal id", "birthday", None, "pizza topping"],
            ["pid", "dob", None, "top"],
            ["nuh71", "2001-12-17", None, "cheese"],
            ["nuh06", "1993-09-20", None, "unknown"],
            ["nuh23", None, None, "mushrooms"],
            ["nuh67", "mssing", None, "this can't be a date - sorry 2016"],
            [
                "nuh27",
                "1999-11-30",
                None,
                "idk - this can't be a date anymore",
            ],
            ["nuh65", "mssing", None, "this one has spaces"],
        ]
        assert obtained == expected


@pytest.mark.parametrize("use_csv", [False, True])
@pytest.mark.parametrize("shift_ignore", [True, False])
def test_ignore_in_text_columns(
    use_csv: bool, shift_ignore: bool, tmp_path: Path
) -> None:

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

    source_file, output_path = src_and_out_names(use_csv, tmp_path, sheet_configs)

    def body(shift_ignore_yaml: None | Path) -> None:
        shift_excel_dates_inplace(
            input_file=source_file,
            output_file=str(output_path),
            patient_sheet="page-desc",
            sheet_configs=sheet_configs,
            min_shift_days=-20,
            max_shift_days=-1,
            seed=14333,
            linking_table_path=str(linking_table_old),
            linking_table_output=str(linking_table_out),
            shift_ignore_yaml=shift_ignore_yaml,
        )

    if not shift_ignore:
        with pytest.raises(HiddenDate) as info:
            body(None)
        assert str(info.value._message) == (
            "hidden date in [sheet_name='page-desc', 2, 2 @ glitter] "
            + 'value="can\'t recall the date but on 12/11/2001 they had an itchy tummy"'
            + " // found=datetime.datetime(2001, 12, 11, 0, 0)"
        )
        return

    else:
        # run the shift
        body(test_data / "shift-ignore.yaml")

        obtained = load_workbook(output_path)["page-desc"]

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
            ["glip", "nuh65", "jkl"],
        ]

        assert expected == obtained


def load_workbook(path: str | Path) -> dict[str, list[list[str | None]]]:

    book = {}

    for page in book_page.book_open(path):
        data = []

        for row, cells in page.stream_rows():
            assert row == len(data)
            for idx, val in enumerate(cells):
                if val is None:
                    continue
                val = str(val).strip()
                if not val:
                    cells[idx] = None
                    continue

                import re

                pattern = r"\d{4}-\d{2}-\d{2} 00:00:00"
                if re.fullmatch(pattern, val):
                    val = val[:10]
                cells[idx] = val
            data.append(cells)

        book[page.name] = data

    return book


@pytest.mark.parametrize("use_csv", [False, True])
def test_move_to_file(use_csv: bool, tmp_path: Path) -> None:

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

    source_file, output_path = src_and_out_names(use_csv, tmp_path, sheet_configs)

    # add the old config we want to remove
    sheet_configs["page-desc"]["shift_ignore"] = {
        "glitter": [
            "can't recall the date but on 12/11/2001 they had an itchy tummy",
        ],
        "foo": [
            "agreed on 21st jul 2017",
        ],
    }

    with pytest.raises(RuntimeError) as error:
        shift_excel_dates_inplace(
            input_file=source_file,
            output_file=str(output_path),
            patient_sheet="page-desc",
            sheet_configs=sheet_configs,
            min_shift_days=-20,
            max_shift_days=-1,
            seed=14333,
            linking_table_path=str(linking_table_old),
            linking_table_output=str(linking_table_out),
        )
        pytest.fail("execution should throw an exception before now")

    assert str(error.value) == "move shift_ignore from sheet_configs to a file"


def src_and_out_names(
    use_csv: bool, tmp_path: Path, sheet_configs: dict[str, any]
) -> tuple[Path | list[Path], Path]:
    # change the source file to point to the csv files
    if use_csv:
        source_file = [test_data / (name + ".csv") for name in sheet_configs]
        output_path = tmp_path / "output"
        output_path.mkdir()
    else:
        source_file = test_data / "workbook.xlsx"
        output_path = tmp_path / "target.xlsx"
    return source_file, output_path

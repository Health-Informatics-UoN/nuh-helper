import textwrap
from pathlib import Path
from typing import Literal, Never

import pytest

from nuh_helper.book_page import (
    BaseWorkBook,
    CSVWorkBook,
    XLSXWorkBook,
    book_open,
)

test_data = (
    # i feel like test data should be in a folder following this pattern
    Path(__file__).parent / "data" / __name__
)
test_book = (
    # keep the .xlsx and .csv in their own folder
    test_data / "book"
)

data_3x4 = [
    ["id", "name", "age", "score"],
    ["1", "Alice", "28", "87.5"],
    ["2", "Bob", "34", "92.1"],
    ["3", "Carol", "22", "78.3"],
]
data_3x6 = [
    ["id", "name", "age", "score", "city", "active"],
    ["1", "Alice", "28", "87.5", "London", "True"],
    ["2", "Bob", "34", "92.1", "Paris", "False"],
    ["3", "Carol", "22", "78.3", "Tokyo", "True"],
]


def test_csv_name_collision() -> None:
    with pytest.raises(ValueError) as error:
        CSVWorkBook([test_book / "data_3x4.csv", test_book / "data_3x4.tsv"])
    assert (
        textwrap.dedent(
            f"""
            the following csv files have name collisions
                data_3x4
                    {test_book}/data_3x4.csv
                data_3x4
                    {test_book}/data_3x4.tsv
            """
        ).strip()
        == str(error.value).strip()
    )


@pytest.mark.parametrize(
    "kind, path",
    [
        (CSVWorkBook, test_book),
        (XLSXWorkBook, test_book / "data.xlsx"),
    ],
)
@pytest.mark.parametrize(
    "to_str",
    [False, True],
)
@pytest.mark.parametrize(
    "generic",
    [False, True],
)
@pytest.mark.parametrize(
    "name, data",
    [
        ("data_3x4", data_3x4),
        ("data_3x6", data_3x6),
    ],
)
def test_read_data(
    kind: CSVWorkBook | XLSXWorkBook,
    path: Path,
    to_str: bool,
    generic: bool,
    name: Literal["data_3x4", "data_3x6"],
    data: list[list[str]],
) -> None:

    # arrange
    book: BaseWorkBook

    # act
    if generic:
        book = book_open(str(path) if to_str else path)
    else:
        book = kind(str(path) if to_str else path)

    # assert
    assert isinstance(book, kind)
    assert name in book
    page = book[name]
    copy = []
    for _, row in page.stream_rows():
        copy.append([str(c) for c in row])

    assert data == copy


@pytest.mark.parametrize(
    "kind, path",
    [
        (CSVWorkBook, test_book),
        (XLSXWorkBook, test_book / "data.xlsx"),
    ],
)
@pytest.mark.parametrize(
    "src_str",
    [False, True],
)
@pytest.mark.parametrize(
    "out_str",
    [False, True],
)
@pytest.mark.parametrize(
    "end_str",
    [False, True],
)
@pytest.mark.parametrize(
    "generic_open",
    [False, True],
)
@pytest.mark.parametrize(
    "name, data",
    [
        ("data_3x4", data_3x4),
        ("data_3x6", data_3x6),
    ],
)
def test_copy_book(
    src_str: bool,
    out_str: bool,
    end_str: bool,
    kind: type,
    path: Path,
    tmp_path: Path,
    generic_open: bool,
    name: str,
    data: list[list[str]],
) -> None:
    # arrange
    source: BaseWorkBook
    changed_data = [
        [f"edit[{r},{c}]{data[r][c]}" for c in range(len(data[0]))]
        for r in range(len(data))
    ]
    src = str(path) if src_str else path
    out = str(tmp_path) if out_str else tmp_path
    end: str | Path = (tmp_path / path.name) if path.is_file() else tmp_path
    if end_str:
        end = str(end)

    # act
    source = kind(src)

    for page, row, cells in source.stream_copy(out):
        assert page in ("data_3x4", "data_3x6"), f"bad page name {page}"
        for col, val in enumerate(cells):
            cells[col] = f"edit[{row},{col}]{val}"

    # assert
    target: BaseWorkBook
    target = book_open(end) if generic_open else kind(end)

    # assert
    assert name in source
    assert name in target
    assert isinstance(source, kind)
    assert isinstance(target, kind)

    page = target[name]
    copy = []
    for row, cell in page.stream_rows():
        assert len(copy) == row
        copy.append(cell)
    assert copy == changed_data


def test_stream_csv_rows() -> None:
    book = CSVWorkBook([test_book / "data_3x4.csv", test_book / "data_3x6.tsv"])

    copy = []
    page = book["data_3x4"]
    for _, row in page.stream_rows():
        copy.append(row.copy())

    assert copy == data_3x4


def test_csv_stream_wont_overwrite_file(tmp_path: Path) -> None:
    book = CSVWorkBook([test_book / "data_3x4.csv", test_book / "data_3x6.tsv"])
    into = tmp_path / "foo"
    into.write_text("foobar")

    with pytest.raises(RuntimeError) as error:
        for _ in book.stream_copy(into):
            pytest.fail("this should have raised an exception")

    assert str(error.value) == f"can't write csvs to {into} because it's a file"


@pytest.mark.parametrize(
    "kind, path",
    [
        (CSVWorkBook, [test_book / "data_3x4.csv", test_book / "data_3x6.tsv"]),
        (XLSXWorkBook, test_book / "data.xlsx"),
    ],
)
def test_stream_rows(kind: type, path: Path) -> None:
    book = kind(path)

    copy = {}

    # scan each row
    for page in book:
        copy[page.name] = []
        for idx, row in page.stream_rows():
            assert idx == len(copy[page.name])
            copy[page.name].append([str(c) for c in row])

    assert copy == {
        "data_3x4": data_3x4,
        "data_3x6": data_3x6,
    }


def test_non_file(tmp_path: Path) -> None:
    with pytest.raises(RuntimeError) as error:
        XLSXWorkBook(tmp_path / "foo.xlsx")
    assert f"xlsx file not found {tmp_path}/foo.xlsx" == str(error.value)
    with pytest.raises(RuntimeError) as error:
        for _ in XLSXWorkBook(tmp_path / "bar.xlsx").stream_copy(tmp_path / "foo.xlsx"):
            pytest.fail("this should have raised an exception")
    assert f"xlsx file not found {tmp_path}/bar.xlsx" == str(error.value)


def test_csv_file() -> None:
    path = test_book / "data.xlsx"

    with pytest.raises(RuntimeError) as error:
        CSVWorkBook(path)
    assert f"need a dir or list of files, but, got a file {path=}" == str(error.value)


def test_xlsx_name_cap() -> Never:
    book = XLSXWorkBook(test_book / "data.xlsx")
    long_name = "name which exceeds 31 characters"
    with (
        pytest.raises(KeyError) as error,
        book[long_name] as page,
    ):
        pytest.fail(f"{long_name=} should have failed {page.name=}")
    assert str(error.value) == "'excel limits you to 31 character page names'"


def test_bom_is_removed() -> None:
    book = CSVWorkBook([test_data / "page-desc.csv"])

    obtained = []
    page = book["page-desc"]
    for idx, row in page.stream_rows():
        assert idx == len(obtained)
        obtained.append(row.copy())

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
        ["clip", "nuh65", "jkl"],
    ]
    assert obtained == expected

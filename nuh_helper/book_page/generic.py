from pathlib import Path

from .book import BasePage
from .csv_book import CSVWorkBook
from .xlsx_book import XLSXWorkBook

# use this or-type instead of the "real" base class since we only have/support these two
Book = CSVWorkBook | XLSXWorkBook


def book_open(src: str | Path | list[Path]) -> Book:
    if isinstance(src, str):
        src = Path(src)

    if isinstance(src, Path) and src.name.endswith(".xlsx"):
        return XLSXWorkBook(src)
    else:
        return CSVWorkBook(src)


__all__ = [  # noqa: PLE0604
    Book,
    BasePage,
    book_open,
]

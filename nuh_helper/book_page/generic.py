from pathlib import Path

from .book import Page
from .csv_book import CSVBook
from .xlsx_book import XLSXBook

# use this or-type instead of the "real" base class since we only have/support these two
Book = CSVBook | XLSXBook


def book_open(src: str | Path | list[Path]) -> Book:
    if isinstance(src, str):
        src = Path(src)

    if isinstance(src, Path) and src.name.endswith(".xlsx"):
        return XLSXBook(src)
    else:
        return CSVBook(src)


__all__ = [  # noqa: PLE0604
    Book,
    Page,
    book_open,
]

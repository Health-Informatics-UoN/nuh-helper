import datetime
from typing import Never

import pandas as pd
import pytest

from nuh_helper.date_shift._parse import _normalize_patient_id, _parse_date_value


@pytest.mark.parametrize(
    "source, expected",
    [
        ("12-31-2004", "2004-12-31 00:00:00"),
        ("12/03/2016", "2016-03-12 00:00:00"),
        ("2016-12-03", "2016-12-03 00:00:00"),
        ("2016-13-03", "2016-03-13 00:00:00"),
        ("31-12-2004", "2004-12-31 00:00:00"),
        ("not-a-date", "None"),
        (3.9, "None"),
        (None, "None"),
        (datetime.date(2026, 9, 12), "2026-09-12 00:00:00"),
        (datetime.datetime(2026, 9, 12, 13, 14, 15), "2026-09-12 13:14:15"),
        (float("nan"), "None"),
        (pd.NaT, "None"),
        (pd.Timestamp("2026-09-12 13:14:15"), "2026-09-12 13:14:15"),
    ],
)
def test_date_value(source: str, expected: str) -> Never:
    obtained: str = str(_parse_date_value(source))

    assert obtained == expected


@pytest.mark.parametrize(
    "source, expected",
    [
        ("   ", "None"),
        ("  nuh007 ", "nuh007"),
        ("", "None"),
        ("2016-13-03", "2016-13-03"),
        (0, "None"),
        (3.9, "3.9"),
        (None, "None"),
        (datetime.date(2026, 9, 12), "2026-09-12"),
        (float("nan"), "None"),
    ],
)
def test_patient_id(source: str, expected: str) -> Never:
    obtained: str = str(_normalize_patient_id(source))

    assert obtained == expected

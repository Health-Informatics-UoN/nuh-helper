from nuh_helper.date_shift.validation import (
    Error as Error,
)
from nuh_helper.date_shift.validation import (
    ExcessRows as ExcessRows,
)
from nuh_helper.date_shift.validation import (
    Path as Path,
)
from nuh_helper.date_shift.validation import (
    PatientColumnMissing as PatientColumnMissing,
)
from nuh_helper.date_shift.validation import (
    UnlabeledColumns as UnlabeledColumns,
)
from nuh_helper.date_shift.validation import (
    Worksheet as Worksheet,
)
from nuh_helper.date_shift.validation import (
    blank_cell as blank_cell,
)
from nuh_helper.date_shift.validation import (
    blank_row as blank_row,
)
from nuh_helper.date_shift.validation import (
    format_errors as format_errors,
)
from nuh_helper.date_shift.validation import (
    inspect as inspect,
)


def test_inspect() -> None:
    """

    https://github.com/Health-Informatics-UoN/nuh-helper/issues/78

    https://github.com/Health-Informatics-UoN/nuh-helper/issues/8
    """

    patients_src = Path(__file__).parent / "data/patients2with-extra-data.xlsx"

    errors = inspect(patients_src, sheet_configs)

    message = format_errors(errors)
    print(">>>")
    print(message)
    print("<<<")

    assert ExcessRows("measurements", [14]) in errors
    assert UnlabeledColumns("measurements", [3, 4]) in errors

    assert len(errors) == 2


sheet_configs = {
    "patients": {
        "patient_id_col": "patient_id",
        "header_row": 0,
        "skip_rows_after_header": [],
        "date_columns": [
            "dob",
            "last_alive",
        ],
    },
    "results": {
        "patient_id_col": "patient_id",
        "header_row": 0,
        "skip_rows_after_header": [],
        "date_columns": [
            "date_result",
        ],
    },
    "measurements": {
        "patient_id_col": "p_id",
        "header_row": 1,
        "skip_rows_after_header": [2, 3],
        "date_columns": [
            "date8061",
        ],
    },
}

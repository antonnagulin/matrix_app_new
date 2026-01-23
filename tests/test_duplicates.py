from src.filters import normalize_fio


def test_no_duplicates():
    data = [
        {"ФИО": "Иванов Иван Иванович"},
        {"ФИО": "Петров Петр Петрович"},
    ]

    fio_dict = {}
    for row in data:
        fio = normalize_fio(row["ФИО"])
        fio_dict.setdefault(fio, []).append(row)

    assert all(len(v) == 1 for v in fio_dict.values())

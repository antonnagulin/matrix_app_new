import pandas as pd
import pytest
from openpyxl import Workbook


@pytest.fixture
def sprav_df():
    return pd.DataFrame(
        [
            {
                "Таб.№": "1",
                "ФИО": "Иванов Иван Иванович",
                "Штатная должность": "Инженер",
                "Статус занятости": "Работает",
                "Увольн.": "31.12.9999",
            },
            {
                "Таб.№": "2",
                "ФИО": "Иванов Иван Иванович",
                "Штатная должность": "Интеграция: штатная должность по умолч.",
                "Статус занятости": "",
                "Увольн.": "01.01.2020",
            },
        ]
    )


@pytest.fixture
def check_wb():
    wb = Workbook()
    ws = wb.active
    ws.append(["ФИО"])
    ws.append(["Иванов Иван Иванович"])
    ws.append(["Петров Петр Петрович"])
    return wb, ws

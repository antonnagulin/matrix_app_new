from datetime import datetime
import sys
from .constants import ACTIVE_DATE
import pandas as pd


def normalize_fio(fio: str) -> str:
    if pd.isna(fio) or fio is None:
        return ""
        
    fio = str(fio).replace(".", "").strip()  # убираем точки и пробелы
    fio = fio.replace("Ё", "Е").replace("ё", "е")  # нормализуем ё
    parts = fio.split()
    
    if len(parts) >= 3:
        # Фамилия + инициалы (первые буквы имени и отчества)
        return f"{parts[0]} {parts[1][0]}{parts[2][0]}".upper()
    elif len(parts) == 2:
        # Фамилия + имя полностью
        return f"{parts[0]} {parts[1]}".upper()
    elif len(parts) == 1:
        # Только фамилия
        return parts[0].upper()

    return ""



def normalize_spr_fio(fio: str) -> str:
    if pd.isna(fio) or fio is None:
        return ""

    fio = fio.replace("Ё", "Е").replace("ё", "е")  # нормализуем ё
    parts = fio.split()
    if len(parts) >= 3:
        return f"{parts[0]} {parts[1]} {parts[2]}".upper()
    
    return ""


def is_full_fio(fio: str) -> bool:
    """
    Проверяет, что ФИО в формате 'Иванов Иван Иванович'
    """
    parts = fio.split()
    if len(parts) != 3:
        return False
    return all(len(p) > 1 for p in parts)


def is_active(dismiss_value) -> bool:
    """Возвращает True, если сотрудник активен"""
    if pd.isna(dismiss_value):
        return False
    if isinstance(dismiss_value, str):
        try:
            if "-" in dismiss_value:
                d = pd.to_datetime(dismiss_value).date()
            else:
                d = datetime.strptime(dismiss_value, "%d.%m.%Y").date()
            return d == datetime.strptime(ACTIVE_DATE, "%d.%m.%Y").date()
        except:
            return False
    if isinstance(dismiss_value, (pd.Timestamp, datetime)):
        return dismiss_value.date() == datetime.strptime(ACTIVE_DATE, "%d.%m.%Y").date()
    return False


def is_dismissed(record: dict) -> bool:
    position = str(record.get("Штатная должность", "")).strip()
    status = str(record.get("Статус занятости", "")).strip().lower()
    dismiss = record.get("Увольн.")
    return (
        position == "Интеграция: штатная должность по умолч."
        or status == "уволился"
        or not is_active(dismiss)
    )


def join_with_slash(row):
    parts = [
        row.get("Орг. единица уровня 1", ""),
        row.get("Орг. единица уровня 2", ""),
        row.get("Орг. единица уровня 3", ""),
        row.get("Орг. единица уровня 4", ""),
    ]
    parts = [str(p).strip() for p in parts if pd.notna(p) and str(p).strip()]
    return " / ".join(parts)
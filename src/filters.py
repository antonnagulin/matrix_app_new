import pandas as pd
from datetime import datetime
from .constants import ACTIVE_DATE

import pandas as pd

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



def normalize_spr_fio(fio: str) -> str:
    if pd.isna(fio) or fio is None:
        return ""

    fio = fio.replace("Ё", "Е").replace("ё", "е")  # нормализуем ё
    parts = fio.split()
    if len(parts) >= 3:
        return f"{parts[0]} {parts[1]} {parts[2]}".upper()
    
    return ""



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

def is_dismissed(record) -> bool:
    position = str(record.get("Штатная должность", "")).strip()
    status = str(record.get("Статус занятости", "")).strip().lower()
    dismiss = record.get("Увольн.")
    # return position == "Интеграция: штатная должность по умолч." or status == "уволился" or not is_active(dismiss)
    return status == "уволился" or not is_active(dismiss)

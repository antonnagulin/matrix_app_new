"""Многоуровневый индекс ФИО справочника: полный / инициалы / фамилия+имя / фамилия+И."""

from __future__ import annotations

from dataclasses import dataclass, field

import pandas as pd

from .constants import FIO_COL
from .filters import is_dismissed


def _parts(fio) -> list[str]:
    if fio is None or (isinstance(fio, float) and pd.isna(fio)):
        return []
    text = str(fio).strip()
    # Тире/запятые → пробелы/единый дефис, точки → пробелы (инициалы)
    text = (
        text.replace("\u2013", "-")
        .replace("\u2014", "-")
        .replace(",", " ")
        .replace(".", " ")
    )
    text = text.replace("Ё", "Е").replace("ё", "е")
    return [p for p in text.split() if p]


def full_fio_key(fio) -> str:
    """Полный ключ: все слова верхним регистром (ИВАНОВ ИВАН ИВАНОВИЧ / ИВАНОВ И И)."""
    parts = _parts(fio)
    return " ".join(parts).upper() if parts else ""


def initials_fio_key(fio) -> str:
    """Ключ инициалов: ФАМИЛИЯ + ИО (для 3+ частей)."""
    parts = _parts(fio)
    if len(parts) >= 3:
        return f"{parts[0]} {parts[1][0]}{parts[2][0]}".upper()
    return ""


def surname_name_key(fio) -> str:
    """Ключ фамилия+имя (первые два токена)."""
    parts = _parts(fio)
    if len(parts) >= 2:
        return f"{parts[0]} {parts[1]}".upper()
    return ""


def surname_initial_key(fio) -> str:
    """Ключ фамилия + первая буква имени («Иванов И» / полное ФИО)."""
    parts = _parts(fio)
    if len(parts) >= 2 and parts[1]:
        return f"{parts[0]} {parts[1][0]}".upper()
    return ""


def person_group_key(fio) -> str:
    """Ключ «человек»: полный ключ, иначе инициалы, иначе как есть."""
    return full_fio_key(fio) or initials_fio_key(fio) or str(fio or "").strip().upper()


@dataclass
class FioIndex:
    by_full: dict[str, list] = field(default_factory=dict)
    # initials -> {person_key: [records]}
    by_initials: dict[str, dict[str, list]] = field(default_factory=dict)
    # surname+name -> {person_key: [records]}
    by_surname_name: dict[str, dict[str, list]] = field(default_factory=dict)
    # surname+first letter of name -> {person_key: [records]}
    by_surname_initial: dict[str, dict[str, list]] = field(default_factory=dict)

    def person_count(self) -> int:
        return len(self.by_full)


def build_fio_index(spr_df: pd.DataFrame) -> FioIndex:
    index = FioIndex()
    # to_dict("records") заметно быстрее iterrows на больших справочниках
    for row in spr_df.to_dict("records"):
        fio = row.get(FIO_COL, "")
        parts = _parts(fio)
        full = full_fio_key(fio)
        init = initials_fio_key(fio)
        sn = surname_name_key(fio)
        si = surname_initial_key(fio)
        person = person_group_key(fio)
        if not person:
            continue

        if full:
            index.by_full.setdefault(full, []).append(row)
        if init:
            index.by_initials.setdefault(init, {}).setdefault(person, []).append(row)
        if sn and len(parts) >= 2:
            # Для «Иванов И» от инициалов не засоряем индекс фамилия+имя
            # однобуквенным именем — только если имя длиннее 1 символа
            name_part = parts[1]
            if len(name_part) > 1:
                index.by_surname_name.setdefault(sn, {}).setdefault(person, []).append(row)
        if si:
            index.by_surname_initial.setdefault(si, {}).setdefault(person, []).append(row)
    return index


def _flatten_groups(groups: dict[str, list]) -> list:
    rows: list = []
    for recs in groups.values():
        rows.extend(recs)
    return rows


def lookup_fio(index: FioIndex, fio_value) -> list:
    """
    Найти записи справочника для ФИО из файла проверки.

    Если на уровне инициалов / фамилия+имя несколько разных людей —
    возвращаются записи всех кандидатов (дальше логика пометит дубликат).
    """
    full = full_fio_key(fio_value)
    parts = _parts(fio_value)
    is_initials_query = len(parts) >= 3 and all(len(p) == 1 for p in parts[1:3])
    is_two_token_initial = len(parts) == 2 and len(parts[1]) == 1

    # 1) Точное полное совпадение (не для запроса-инициалов «И И»)
    if full and not is_initials_query and not is_two_token_initial and full in index.by_full:
        return list(index.by_full[full])

    # 2) Инициалы «Фамилия И О» / «Фамилия И. О.»
    init = initials_fio_key(fio_value)
    if init and init in index.by_initials:
        return _flatten_groups(index.by_initials[init])

    # 3) Два токена с однобуквенным именем: «Фамилия И» / «Фамилия И.»
    if is_two_token_initial:
        si = surname_initial_key(fio_value)
        if si and si in index.by_surname_initial:
            return _flatten_groups(index.by_surname_initial[si])

    # 4) Фамилия + имя (2 слова в проверке ↔ полное ФИО в справочнике)
    sn = surname_name_key(fio_value)
    if sn and not is_two_token_initial and sn in index.by_surname_name:
        return _flatten_groups(index.by_surname_name[sn])

    # 5) Last resort: полный ключ (в т.ч. запрос-инициалы, если так записано в справочнике)
    if full and full in index.by_full:
        return list(index.by_full[full])

    return []


def pick_default_duplicate_record(records: list):
    """
    Запись по умолчанию, если разбор дубликатов выключен.

    Returns:
        активная не-совместная запись или None, если таких нет.
    """
    filtered = [
        rec
        for rec in records
        if str(_get(rec, "Группа сотрудников", "")).strip() != "Внутренние совмест."
        and not is_dismissed(rec)
    ]
    if filtered:
        return filtered[0]
    return None


def _get(record, key, default=""):
    if hasattr(record, "get"):
        return record.get(key, default)
    return default


def distinct_person_count(records: list) -> int:
    keys = {person_group_key(r.get("ФИО", "") if hasattr(r, "get") else r["ФИО"]) for r in records}
    keys.discard("")
    return len(keys)

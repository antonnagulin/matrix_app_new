from .filters import is_dismissed, is_full_fio

JOINT_GROUP = "Внутренние совмест."


def is_joint(record) -> bool:
    return str(record.get("Группа сотрудников", "")).strip() == JOINT_GROUP


def _full_fio_key(fio) -> str:
    """Ключ сравнения полного ФИО: регистр и ё/е, без схлопывания в инициалы."""
    text = str(fio or "").strip()
    text = text.replace("Ё", "Е").replace("ё", "е")
    return " ".join(text.split()).casefold()


def check_duplicate(yellow_duplicates: dict):
    """
    Returns:
        unresolved, resolved, dismissed, joints_only
    """
    unresolved = {}
    resolved = {}
    dismissed = {}
    joints_only = {}

    for key, value in yellow_duplicates.items():
        filtered = [
            inf
            for inf in value
            if not is_joint(inf) and not is_dismissed(inf)
        ]

        if not filtered:
            non_joint = [inf for inf in value if not is_joint(inf)]
            if not non_joint:
                joints_only[key] = value
            else:
                # Остались только уволенные (возможно вместе с совместниками)
                dismissed[key] = value
            continue

        if len(filtered) == 1:
            resolved[key] = filtered
        else:
            unresolved[key] = filtered

    return unresolved, resolved, dismissed, joints_only


def second_check_duplicate(yellow_duplicates: dict):
    """
    Проверяет полные ФИО и делит на:
    - remaining (неоднозначные / сокращённые → ручной выбор)
    - resolved (ровно одно совпадение по полному ФИО)
    """
    remaining = {}
    resolved = {}

    for key, value in yellow_duplicates.items():
        _row, _fio_norm, fio_value, _pos = key

        if not is_full_fio(fio_value):
            remaining[key] = value
            continue

        key_fio = _full_fio_key(fio_value)
        filtered = [
            item for item in value if _full_fio_key(item.get("ФИО", "")) == key_fio
        ]

        if not filtered:
            remaining[key] = value
            continue

        if len(filtered) == 1:
            resolved[key] = filtered
        else:
            remaining[key] = filtered

    return remaining, resolved

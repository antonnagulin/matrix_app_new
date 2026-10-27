"""Фильтрация недавно уволенных (по году в дате увольнения)."""


def check_dismissed(orange_dict: dict) -> dict:
    """Оставить записи с увольнением в 2025–2026."""
    kept = {}
    for key, value in orange_dict.items():
        date = str(value[0].get("Увольн.", "") or "")
        if date.startswith("2025") or date.startswith("2026"):
            kept[key] = value
    return kept


# Совместимость со старым именем
check_dissmised = check_dismissed

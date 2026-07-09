from .filters import is_full_fio

def check_duplicate(elloyy_dublicate: dict):
    new_elloyy_dublicate = {}
    not_elloyy_dublicate = {}
    is_dismissed = {}

    for key, value in elloyy_dublicate.items():
        filtered = [
            inf for inf in value
            if inf["Штатная должность"] != "Интеграция: штатная должность по умолч."
            and inf["Группа сотрудников"] != "Внутренние совмест."
            and inf["Статус занятости"] == "активный"
        ]

        if not filtered:
            is_dismissed[key] = value
            continue

        if len(filtered) == 1:
            not_elloyy_dublicate[key] = filtered
        else:
            new_elloyy_dublicate[key] = filtered

    return new_elloyy_dublicate, not_elloyy_dublicate, is_dismissed


def second_check_duplicate(elloyy_dublicate: dict):
    """
    Проверяет полные ФИО и делит на:
    - new_elloyy_dublicate (два и более совпадения)
    - not_elloyy_dublicate (ровно одно совпадение)
    """
    new_elloyy_dublicate = {}
    not_elloyy_dublicate = {}

    for key, value in elloyy_dublicate.items():
        row, fio_norm, fio_value, pos_or_eararchy_value = key

        if not is_full_fio(fio_value):
            continue  # пропускаем сокращённые ФИО

        # фильтруем по ФИО
        filtered = [item for item in value if fio_value == item["ФИО"]]

        if not filtered:
            continue

        if len(filtered) == 1:
            not_elloyy_dublicate[key] = filtered
        else:
            new_elloyy_dublicate[key] = filtered
            
    # print(f'new_elloyy_dublicate {new_elloyy_dublicate}')
    # print(f'not_elloyy_dublicate {not_elloyy_dublicate}')
    return new_elloyy_dublicate, not_elloyy_dublicate
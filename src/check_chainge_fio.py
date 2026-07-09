import pandas as pd
from .file_utils import read_sprav_file
from .filters import normalize_fio


def chainge_fio(red_not_found, chainge_fio_file):
    
    # print("Не нашлись в справочнике")
    # print(red_not_found)

    chainge_fio_df = read_sprav_file(chainge_fio_file)

    # Нормализуем ФИО один раз
    chainge_fio_df["ФИО_NORM"] = chainge_fio_df["Смена ФИО"].apply(normalize_fio)

    not_in_all = []

    for item in red_not_found:
        item_norm = normalize_fio(item)

        matched_rows = chainge_fio_df[
            chainge_fio_df["ФИО_NORM"] == item_norm
        ]

        if matched_rows.empty:
            not_in_all.append(item)
            continue

        unique_values_count = pd.concat(
            [matched_rows["Табельный номер"], matched_rows["Смена ФИО"]]
        ).nunique()

        if unique_values_count > 1:  # есть хотя бы одно отличие
            print(f"\nНайдено расхождение для: {item}")
            print(matched_rows.drop(columns=["ФИО_NORM"]).iloc[0:10])

    # if not_in_all:
    #     print("\nНе нашлись вообще")
        # print(not_in_all)
    
    return 

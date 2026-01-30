import pandas as pd

# Загружаем файлы
main_df = pd.read_excel("check.xlsx")
dict_df = pd.read_excel("spravochnik.xlsx")

# Функция: убрать первые два слова
def remove_first_two_words(text):
    if pd.isna(text):
        return ""
    words = str(text).split()
    return " ".join(words[2:]) if len(words) > 2 else ""

# Обрабатываем "Раздел персонала"
dict_df["Раздел персонала"] = (
    dict_df["Раздел персонала"]
    .apply(remove_first_two_words)
)


# Сцепляем с "Орг. единица уровня 2"
def join_with_slash(row):
    parts = [
        'ООО "Газпром добыча Краснодар"',
        row["Раздел персонала"],
        row["Орг. единица уровня 2"],
        row["Орг. единица уровня 3"],
        row["Орг. единица уровня 4"],
    ]

    # оставляем только непустые значения
    parts = [str(p).strip() for p in parts if pd.notna(p) and str(p).strip()]

    return " / ".join(parts)

dict_df["Раздел персонала"] = dict_df.apply(join_with_slash, axis=1)

# Чистим пробелы
dict_df["Раздел персонала"] = dict_df["Раздел персонала"].str.strip()

# Добавляем столбец в основной файл (по строкам)
main_df["Раздел персонала"] = dict_df["Раздел персонала"]

# Сохраняем результат
main_df.to_excel("result.xlsx", index=False)

print("Готово! Столбец 'Раздел персонала' перезаписан и сохранён в result.xlsx")

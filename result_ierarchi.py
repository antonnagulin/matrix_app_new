import pandas as pd
import os
from src.settings.result_ierarchi import target_columns, folder_path, normalize_department_name


def analyze_departments_pivot(file_paths):
    """
    Анализирует файлы и собирает данные для сводной таблицы с нормализацией названий подразделений.
    Возвращает список словарей с данными по каждому файлу и подразделению.
    """
    results = []

    for file_path in file_paths:
        try:
            # Определяем расширение файла и читаем его
            if file_path.lower().endswith('.xlsx') or file_path.lower().endswith('.xls'):
                df = pd.read_excel(file_path)
            elif file_path.lower().endswith('.csv'):
                df = pd.read_csv(file_path, encoding='utf-8')
            else:
                print(f"⛔ Пропускаем файл {file_path} — неподдерживаемый формат")
                continue

            # Ищем нужный столбец
            found_column = None
            for col in target_columns:
                if col in df.columns:
                    found_column = col
                    break

            if not found_column:
                print(f"⚠️ В файле {file_path} не найдены столбцы: {target_columns}")
                continue

            # Получаем подсчёт уникальных значений с их количеством
            value_counts = df[found_column].value_counts(dropna=True)

            # Нормализуем названия подразделений и добавляем результаты для этого файла
            for department, count in value_counts.items():
                normalized_dept = normalize_department_name(department)
                
                if str(normalized_dept).strip() == '' or str(normalized_dept).strip().replace('_', '') == '':
                    continue  
                                
                results.append({
                    'Файл': os.path.basename(file_path),
                    'Подразделение': normalized_dept,
                    'Количество записей': count
                })

        except Exception as e:
            print(f"❌ Ошибка при обработке файла {file_path}: {e}")

    return results

def save_pivot_table_to_excel(results, output_file='ИтоговаяМатрица1.xlsx'):
    """
    Сохраняет результаты в виде сводной таблицы в Excel-файл с группировкой и статистикой.
    """
    if not results:
        print("❌ Нет данных для сохранения")
        return

    # Создаём DataFrame из результатов
    df_result = pd.DataFrame(results)

    # Группируем по файлу и нормализованному подразделению, суммируем количество записей
    df_grouped = df_result.groupby(['Файл', 'Подразделение'], as_index=False)['Количество записей'].sum()

    # Создаём сводную таблицу: файлы — строки, подразделения — столбцы
    pivot_table = df_grouped.pivot_table(
        index='Файл',
        columns='Подразделение',
        values='Количество записей',
        aggfunc='sum',
        fill_value=0
    )

    # Сортируем столбцы по алфавиту для удобства
    pivot_table = pivot_table.reindex(sorted(pivot_table.columns), axis=1)

    # Добавляем сумму по строкам (общее количество записей для каждого файла)
    pivot_table['Сумма по файлу'] = pivot_table.sum(axis=1)

    # Добавляем строку с суммой по столбцам (общее количество записей для каждого подразделения)
    pivot_table.loc['Сумма по подразделениям'] = pivot_table.sum(numeric_only=True, axis=0)

    # Рассчитываем общее число нахождений (сумма всех ячеек, кроме итоговых строк/столбцов)
    total_findings = pivot_table.iloc[:-1, :-1].sum().sum()

    # Заполняем ячейку пересечения итоговой строки и столбца общим числом нахождений
    pivot_table.loc['Сумма по подразделениям', 'Сумма по файлу'] = total_findings

    # Записываем в Excel
    try:
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            pivot_table.to_excel(writer, sheet_name='Сводная таблица', index=True)


            # Настраиваем ширину колонок
            worksheet = writer.sheets['Сводная таблица']
            # Ширина для колонки «Файл»
            worksheet.column_dimensions['A'].width = 35
            # Ширина для остальных колонок (подразделений)
            for col in range(2, pivot_table.shape[1] + 2):
                worksheet.column_dimensions[worksheet.cell(row=1, column=col).column_letter].width = 15

        print(f"✅ Результаты успешно сохранены в файл: {output_file}")
        print(f"📊 Сводная таблица: {pivot_table.shape[0]} файлов × {pivot_table.shape[1]} подразделений")
        print(f"🔎 Обработано записей: {len(df_result)}")
        print(f"🧮 Общее число нахождений: {total_findings}")

    except Exception as e:
        print(f"❌ Ошибка при сохранении файла: {e}")

def main():
    file_extensions = ['.xlsx', '.xls', '.csv']
    file_paths = [
        os.path.join(folder_path, f)
        for f in os.listdir(folder_path)
        if any(f.lower().endswith(ext) for ext in file_extensions)
    ]

    if not file_paths:
        print("❌ Файлы не найдены в указанной папке")
        return

    print(f"🔎 Найдено файлов для обработки: {len(file_paths)}")

    # Анализируем файлы
    results = analyze_departments_pivot(file_paths)

    # Сохраняем сводную таблицу в Excel
    save_pivot_table_to_excel(results)

if __name__ == "__main__":
    main()

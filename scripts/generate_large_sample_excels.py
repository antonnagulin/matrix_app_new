"""Генерация больших тестовых Excel-файлов (stdlib only).

По умолчанию:
  sample_data/sprav_large.xlsx          — 8 000 записей справочника
  sample_data/check_large.xlsx          — 15 000 строк проверки
  sample_data/sprav_renamed_fio_large.xlsx — 2 000 смен ФИО

Запуск:
  python scripts/generate_large_sample_excels.py
  python scripts/generate_large_sample_excels.py --sprav 20000 --check 50000 --renamed 5000
  python scripts/generate_large_sample_excels.py --renamed-only --renamed 3000
"""

from __future__ import annotations

import argparse
import random
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from generate_sample_excels import OUT, write_xlsx  # noqa: E402

ACTIVE_DATE = "31.12.9999"
INTEGRATION_POSITION = "Интеграция: штатная должность по умолч."

SPRAV_HEADERS = [
    "Таб.№",
    "ФИО",
    "Штатная должность",
    "Статус занятости",
    "Увольн.",
    "Подразделение",
]

CHECK_HEADERS = ["ФИО", "Должность", "Подразделение", "Таб.№", "Примечание"]

RENAMED_HEADERS = [
    "Старое ФИО",
    "Новое ФИО",
    "Таб.№",
    "Дата изменения",
    "Примечание",
]

RENAMED_NOTES = {
    "marriage": "Смена фамилии после замужества",
    "divorce": "Возврат девичьей фамилии",
    "yo": "Замена ё/е в документах",
    "initials_to_full": "Раскрытие инициалов",
    "full_to_initials": "Сокращение до инициалов в старой выгрузке",
    "hyphen": "Смена двойной фамилии",
    "typo": "Исправление опечатки в фамилии",
    "transliteration": "Унификация написания",
}

SURNAMES = [
    "Иванов", "Петров", "Сидоров", "Смирнов", "Кузнецов", "Попов", "Васильев",
    "Соколов", "Михайлов", "Новиков", "Фёдоров", "Морозов", "Волков", "Алексеев",
    "Лебедев", "Семёнов", "Егоров", "Павлов", "Козлов", "Степанов", "Николаев",
    "Орлов", "Андреев", "Макаров", "Никитин", "Захаров", "Зайцев", "Соловьёв",
    "Борисов", "Яковлев", "Григорьев", "Романов", "Воробьёв", "Сергеев", "Кузьмин",
    "Фролов", "Александров", "Дмитриев", "Королёв", "Гусев", "Киселёв", "Ильин",
    "Максимов", "Поляков", "Сорокин", "Виноградов", "Ковалёв", "Белов", "Медведев",
    "Ершов", "Ёлкин", "Антонов", "Тарасов", "Жуков", "Баранов", "Филиппов", "Комаров",
    "Давыдов", "Беляев", "Герасимов", "Богданов", "Осипов", "Сидоренко", "Матвеев",
    "Титов", "Марков", "Миронов", "Крылов", "Куликов", "Карпов", "Власов", "Мельников",
    "Денисов", "Гаврилов", "Тихонов", "Казаков", "Афанасьев", "Данилов", "Савельев",
    "Тимофеев", "Фомин", "Чернов", "Абрамов", "Мартынов", "Ефимов", "Федотов", "Щербаков",
    "Назаров", "Калинин", "Исаев", "Чернышёв", "Быков", "Маслов", "Родионов", "Коновалов",
    "Лазарев", "Воронин", "Климов", "Филатов", "Пономарёв", "Голубев", "Кудрявцев",
    "Прохоров", "Наумов", "Потапов", "Журавлёв", "Уваров", "Трофимов", "Аксёнов",
]

NAMES_M = [
    "Александр", "Алексей", "Андрей", "Антон", "Артём", "Борис", "Вадим", "Валентин",
    "Василий", "Виктор", "Владимир", "Владислав", "Геннадий", "Георгий", "Григорий",
    "Даниил", "Денис", "Дмитрий", "Евгений", "Егор", "Иван", "Игорь", "Илья", "Кирилл",
    "Константин", "Леонид", "Максим", "Михаил", "Никита", "Николай", "Олег", "Павел",
    "Пётр", "Роман", "Руслан", "Сергей", "Станислав", "Степан", "Тимофей", "Фёдор",
    "Юрий", "Ярослав", "Арсений", "Матвей", "Платон", "Семён", "Тарас", "Эдуард",
]

NAMES_F = [
    "Анна", "Алина", "Валентина", "Вера", "Виктория", "Галина", "Дарья", "Евгения",
    "Екатерина", "Елена", "Жанна", "Зинаида", "Ирина", "Ксения", "Лариса", "Людмила",
    "Марина", "Мария", "Надежда", "Наталья", "Оксана", "Ольга", "Полина", "Светлана",
    "Татьяна", "Юлия", "Яна", "Вероника", "Диана", "Кристина", "Милана", "София",
    "Алёна", "Василиса", "Елизавета", "Карина", "Лидия", "Нина", "Раиса", "Ульяна",
]

PATRONYMICS_M = [
    "Иванович", "Петрович", "Сергеевич", "Александрович", "Дмитриевич", "Андреевич",
    "Николаевич", "Владимирович", "Алексеевич", "Михайлович", "Егорович", "Олегович",
    "Павлович", "Романович", "Юрьевич", "Борисович", "Геннадьевич", "Фёдорович",
    "Антонович", "Викторович", "Константинович", "Степанович", "Тимофеевич", "Ярославович",
]

PATRONYMICS_F = [
    "Ивановна", "Петровна", "Сергеевна", "Александровна", "Дмитриевна", "Андреевна",
    "Николаевна", "Владимировна", "Алексеевна", "Михайловна", "Егоровна", "Олеговна",
    "Павловна", "Романовна", "Юрьевна", "Борисовна", "Геннадьевна", "Фёдоровна",
    "Антоновна", "Викторовна", "Константиновна", "Степановна", "Тимофеевна", "Ярославовна",
]

DEPARTMENTS = [
    "Отдел АСУ", "Отдел кадров", "Бухгалтерия", "Склад", "Отдел продаж", "Сервис",
    "IT", "Юр. отдел", "Логистика", "HR", "Аналитика", "Производство", "Закупки",
    "Маркетинг", "Охрана труда", "Транспорт", "Контроль качества", "Планово-экономический",
]

POSITIONS = [
    "Инженер", "Менеджер", "Аналитик", "Бухгалтер", "Специалист", "Техник",
    "Программист", "Юрист", "Курьер", "Кладовщик", "Оператор", "Мастер", "Диспетчер",
    "Экономист", "Секретарь", "Водитель", "Контролёр", "Начальник смены",
]


def _person(rng: random.Random, tab_no: int) -> dict:
    female = rng.random() < 0.45
    surname = rng.choice(SURNAMES)
    if female:
        name = rng.choice(NAMES_F)
        patronymic = rng.choice(PATRONYMICS_F)
        if surname.endswith("ов"):
            surname = surname[:-2] + "ова"
        elif surname.endswith("ев"):
            surname = surname[:-2] + "ева"
        elif surname.endswith("ин"):
            surname = surname + "а"
    else:
        name = rng.choice(NAMES_M)
        patronymic = rng.choice(PATRONYMICS_M)

    fio = f"{surname} {name} {patronymic}"
    dept = rng.choice(DEPARTMENTS)
    position = rng.choice(POSITIONS)
    return {
        "tab": str(tab_no),
        "fio": fio,
        "position": position,
        "status": "работает",
        "dismiss": ACTIVE_DATE,
        "dept": dept,
    }


def _dismiss_record(rec: dict, rng: random.Random, mode: str) -> dict:
    out = dict(rec)
    if mode == "integration":
        out["position"] = INTEGRATION_POSITION
        out["status"] = "работает"
        out["dismiss"] = ACTIVE_DATE
    elif mode == "status":
        out["status"] = "уволился"
        out["dismiss"] = f"{rng.randint(1, 28):02d}.{rng.randint(1, 12):02d}.{rng.randint(2018, 2025)}"
    else:
        out["status"] = "работает"
        out["dismiss"] = f"{rng.randint(1, 28):02d}.{rng.randint(1, 12):02d}.{rng.randint(2015, 2024)}"
    return out


def _to_sprav_row(rec: dict) -> list[str]:
    return [rec["tab"], rec["fio"], rec["position"], rec["status"], rec["dismiss"], rec["dept"]]


def _format_fio_for_check(fio: str, rng: random.Random) -> str:
    parts = fio.split()
    if len(parts) < 3:
        return fio
    surname, name, patronymic = parts[0], parts[1], parts[2]
    roll = rng.random()
    if roll < 0.55:
        return f"{surname} {name[0]}. {patronymic[0]}."
    if roll < 0.85:
        return fio
    # вариант с ё для теста normalize_fio
    return fio.replace("Е", "Ё").replace("е", "ё")


def generate_sprav(rng: random.Random, count: int, duplicate_pairs: int) -> list[list[str]]:
    rows: list[list[str]] = []
    records: list[dict] = []
    tab = 1000

    active_target = count - duplicate_pairs
    dismissed_target = max(1, int(active_target * 0.12))
    active_target -= dismissed_target

    for _ in range(active_target):
        tab += 1
        records.append(_person(rng, tab))

    dismiss_modes = ["integration", "status", "date"]
    for i in range(dismissed_target):
        rec = _person(rng, tab + 1 + i)
        rec = _dismiss_record(rec, rng, dismiss_modes[i % 3])
        records.append(rec)

    # дубликаты: вторая запись уволена, первая активна
    dup_sources = rng.sample(records, k=min(duplicate_pairs, len(records)))
    for src in dup_sources:
        dup = _dismiss_record(dict(src), rng, rng.choice(dismiss_modes))
        dup["tab"] = src["tab"]
        records.append(dup)

    rng.shuffle(records)
    for rec in records:
        rows.append(_to_sprav_row(rec))
    return rows


def _is_female_patronymic(patronymic: str) -> bool:
    return patronymic.endswith("на") or patronymic.endswith("вна")


def _female_surname(surname: str) -> str:
    if surname.endswith("ов"):
        return surname[:-2] + "ова"
    if surname.endswith("ев"):
        return surname[:-2] + "ева"
    if surname.endswith("ин") and not surname.endswith("ина"):
        return surname + "а"
    if surname.endswith("ый"):
        return surname[:-2] + "ая"
    return surname


def _random_date(rng: random.Random) -> str:
    return f"{rng.randint(1, 28):02d}.{rng.randint(1, 12):02d}.{rng.randint(2018, 2025)}"


def _to_initials(fio: str) -> str:
    parts = fio.split()
    if len(parts) < 3:
        return fio
    return f"{parts[0]} {parts[1][0]}. {parts[2][0]}."


def _apply_yo_variant(text: str, rng: random.Random, *, to_yo: bool) -> str:
    if to_yo:
        return text.replace("Е", "Ё").replace("е", "ё")
    return text.replace("Ё", "Е").replace("ё", "е")


def _old_fio_for_change(new_fio: str, change: str, rng: random.Random) -> str | None:
    parts = new_fio.split()
    if len(parts) < 3 and change not in ("yo", "typo"):
        return None

    surname, name, patronymic = parts[0], parts[1], parts[2]

    if change == "marriage":
        if not _is_female_patronymic(patronymic):
            return None
        old_surname = _female_surname(rng.choice(SURNAMES))
        while old_surname == surname:
            old_surname = _female_surname(rng.choice(SURNAMES))
        return f"{old_surname} {name} {patronymic}"

    if change == "divorce":
        if not _is_female_patronymic(patronymic):
            return None
        maiden = _female_surname(rng.choice(SURNAMES))
        while maiden == surname:
            maiden = _female_surname(rng.choice(SURNAMES))
        return f"{maiden} {name} {patronymic}"

    if change == "yo":
        return _apply_yo_variant(new_fio, rng, to_yo=rng.random() < 0.5)

    if change == "initials_to_full":
        return _to_initials(new_fio)

    if change == "full_to_initials":
        return new_fio

    if change == "hyphen":
        if not _is_female_patronymic(patronymic):
            return None
        husband = rng.choice(SURNAMES)
        return f"{_female_surname(surname)}-{husband} {name} {patronymic}"

    if change == "typo":
        if len(surname) < 4:
            return None
        pos = rng.randint(1, len(surname) - 2)
        old_surname = surname[:pos] + rng.choice("аоиеуюя") + surname[pos + 1 :]
        if old_surname == surname:
            return None
        return f"{old_surname} {name} {patronymic}"

    if change == "transliteration":
        variants = {
            "ий": "ый",
            "ый": "ий",
            "ия": "ья",
            "ья": "ия",
        }
        for src, dst in variants.items():
            if surname.endswith(src):
                return f"{surname[:-len(src)]}{dst} {name} {patronymic}"
        return None

    return None


def _synthetic_renamed_row(rng: random.Random, tab: int) -> list[str]:
    rec = _person(rng, tab)
    new_fio = rec["fio"]
    change = rng.choice(list(RENAMED_NOTES.keys()))
    old_fio = _old_fio_for_change(new_fio, change, rng)
    if not old_fio or old_fio == new_fio:
        old_fio = _to_initials(new_fio)
        change = "initials_to_full"
    if change == "full_to_initials":
        old_fio, new_fio = new_fio, _to_initials(new_fio)
    return [old_fio, new_fio, str(tab), _random_date(rng), RENAMED_NOTES[change]]


def generate_renamed_fio(
    rng: random.Random,
    sprav_records: list[list[str]],
    count: int,
) -> list[list[str]]:
    """Справочник смены ФИО: старое имя (в выгрузке) → новое (в основном справочнике)."""
    rows: list[list[str]] = []
    used_tabs: set[str] = set()

    sprav_sample = list(sprav_records)
    rng.shuffle(sprav_sample)
    from_sprav = min(count, max(1, int(count * 0.75)), len(sprav_sample))
    sprav_changes = ["marriage", "yo", "initials_to_full", "hyphen", "typo", "transliteration"]

    for rec in sprav_sample[:from_sprav]:
        tab, new_fio = rec[0], rec[1]
        if tab in used_tabs:
            continue
        change = rng.choice(sprav_changes)
        old_fio = _old_fio_for_change(new_fio, change, rng)
        if not old_fio or old_fio == new_fio:
            change = "initials_to_full"
            old_fio = _to_initials(new_fio)
            if old_fio == new_fio:
                continue
        rows.append([old_fio, new_fio, tab, _random_date(rng), RENAMED_NOTES[change]])
        used_tabs.add(tab)

    tab = 500_000
    while len(rows) < count:
        tab += 1
        row = _synthetic_renamed_row(rng, tab)
        if row[1] in {r[1] for r in rows}:
            continue
        rows.append(row)

    rng.shuffle(rows)
    return rows[:count]


def _renamed_stats(rows: list[list[str]]) -> str:
    by_note: dict[str, int] = {}
    for r in rows:
        by_note[r[4]] = by_note.get(r[4], 0) + 1
    parts = ", ".join(f"{k}: {v}" for k, v in sorted(by_note.items(), key=lambda x: -x[1]))
    return f"Смена ФИО: {len(rows)} записей ({parts})"


def generate_check(
    rng: random.Random,
    sprav_records: list[list[str]],
    count: int,
    *,
    missing_ratio: float = 0.12,
    empty_ratio: float = 0.02,
    repeat_missing_ratio: float = 0.03,
) -> list[list[str]]:
    sprav_by_tab = {row[0]: row for row in sprav_records}
    sprav_list = list(sprav_records)

    rows: list[list[str]] = []
    missing_pool: list[str] = []

    def add_from_sprav(rec: list[str], note: str) -> None:
        fio = _format_fio_for_check(rec[1], rng)
        rows.append([fio, rec[2], rec[5], rec[0], note])

    def add_missing() -> None:
        tab = str(rng.randint(900_000, 999_999))
        surname = rng.choice(SURNAMES)
        name = rng.choice(NAMES_M + NAMES_F)
        patronymic = rng.choice(PATRONYMICS_M + PATRONYMICS_F)
        fio = _format_fio_for_check(f"{surname} {name} {patronymic}", rng)
        missing_pool.append(fio)
        rows.append([fio, rng.choice(POSITIONS), rng.choice(DEPARTMENTS), tab, "нет в справочнике"])

    def add_empty() -> None:
        rows.append(["", "", "", "", "пустая строка"])

    targets = {
        "sprav": int(count * (1 - missing_ratio - empty_ratio - repeat_missing_ratio)),
        "missing": int(count * missing_ratio),
        "empty": int(count * empty_ratio),
        "repeat_missing": int(count * repeat_missing_ratio),
    }
    # добиваем до точного count
    while sum(targets.values()) < count:
        targets["sprav"] += 1
    while sum(targets.values()) > count:
        targets["sprav"] -= 1

    for _ in range(targets["sprav"]):
        add_from_sprav(rng.choice(sprav_list), "из справочника")

    for _ in range(targets["missing"]):
        add_missing()

    for _ in range(targets["empty"]):
        add_empty()

    for _ in range(targets["repeat_missing"]):
        if not missing_pool:
            add_missing()
        fio = rng.choice(missing_pool)
        rows.append([fio, rng.choice(POSITIONS), rng.choice(DEPARTMENTS), "", "повтор отсутствующего"])

    rng.shuffle(rows)
    return rows


def _stats(sprav_rows: list[list[str]], check_rows: list[list[str]]) -> str:
    dismissed = sum(
        1
        for r in sprav_rows
        if r[2] == INTEGRATION_POSITION or r[3].lower() == "уволился" or r[4] != ACTIVE_DATE
    )
    fio_counts: dict[str, int] = {}
    for r in sprav_rows:
        fio_counts[r[1]] = fio_counts.get(r[1], 0) + 1
    duplicates = sum(1 for c in fio_counts.values() if c > 1)

    empty_check = sum(1 for r in check_rows if not (r[0] or "").strip())
    return (
        f"Справочник: {len(sprav_rows)} строк, уволенных/неактивных ~{dismissed}, "
        f"ФИО с дубликатами: {duplicates}\n"
        f"Проверка: {len(check_rows)} строк, пустых ФИО: {empty_check}"
    )


def main() -> None:
    parser = argparse.ArgumentParser(description="Генерация больших тестовых Excel-файлов")
    parser.add_argument("--sprav", type=int, default=8000, help="Число записей в справочнике")
    parser.add_argument("--check", type=int, default=15000, help="Число строк в файле проверки")
    parser.add_argument("--renamed", type=int, default=2000, help="Число записей смены ФИО")
    parser.add_argument("--seed", type=int, default=42, help="Seed для воспроизводимости")
    parser.add_argument(
        "--duplicates",
        type=int,
        default=None,
        help="Число пар дубликатов в справочнике (по умолчанию ~5%% от sprav)",
    )
    parser.add_argument(
        "--renamed-only",
        action="store_true",
        help="Сгенерировать только sprav_renamed_fio_large.xlsx (читает sprav_large.xlsx если есть)",
    )
    args = parser.parse_args()

    rng = random.Random(args.seed)
    renamed_path = OUT / "sprav_renamed_fio_large.xlsx"

    if args.renamed_only:
        dup_pairs = args.duplicates if args.duplicates is not None else max(50, args.sprav // 20)
        sprav_rows = generate_sprav(rng, args.sprav, dup_pairs)
        renamed_rows = generate_renamed_fio(rng, sprav_rows, args.renamed)
        write_xlsx(renamed_path, RENAMED_HEADERS, renamed_rows)
        print(f"OK: {renamed_path}")
        print(_renamed_stats(renamed_rows))
        return

    dup_pairs = args.duplicates if args.duplicates is not None else max(50, args.sprav // 20)

    sprav_rows = generate_sprav(rng, args.sprav, dup_pairs)
    check_rows = generate_check(rng, sprav_rows, args.check)
    renamed_rows = generate_renamed_fio(rng, sprav_rows, args.renamed)

    sprav_path = OUT / "sprav_large.xlsx"
    check_path = OUT / "check_large.xlsx"

    write_xlsx(sprav_path, SPRAV_HEADERS, sprav_rows)
    write_xlsx(check_path, CHECK_HEADERS, check_rows)
    write_xlsx(renamed_path, RENAMED_HEADERS, renamed_rows)

    print(f"OK: {sprav_path}")
    print(f"OK: {check_path}")
    print(f"OK: {renamed_path}")
    print(_stats(sprav_rows, check_rows))
    print(_renamed_stats(renamed_rows))


if __name__ == "__main__":
    main()

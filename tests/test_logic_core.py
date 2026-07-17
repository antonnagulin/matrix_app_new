from src.logic import run_check
from src.constants import RED_FILL_COLOR, ELOYY_FILL_COLOR, GREEN_FILL_COLOR


class DummyLog:
    def __init__(self):
        self.lines = []

    def insert(self, *_args):
        self.lines.append(_args[1])

    def see(self, *_):
        pass


def test_run_check_red_and_yellow_without_duplicate_resolve(sprav_df, check_wb, monkeypatch):
    wb, ws = check_wb
    log = DummyLog()

    monkeypatch.setattr("src.logic.read_sprav_file", lambda _: sprav_df)
    monkeypatch.setattr("src.logic.load_check_file", lambda _: (wb, ws))
    monkeypatch.setattr("src.logic.save_file", lambda *_: None)

    run_check(
        sprav_file="sprav.xlsx",
        check_file="check.xlsx",
        output_file="out.xlsx",
        check_dismissed=True,
        check_duplicates=False,
        log=log,
    )

    # Иванов — дубликат в справочнике, без разбора → жёлтый
    assert ELOYY_FILL_COLOR in (ws["A2"].fill.start_color.rgb or "")
    # Петров отсутствует → красный
    assert RED_FILL_COLOR in (ws["A3"].fill.start_color.rgb or "")
    assert "Петров Петр Петрович" in "".join(log.lines)


def test_run_check_resolves_duplicate_to_green(sprav_df, check_wb, monkeypatch):
    wb, ws = check_wb
    log = DummyLog()

    monkeypatch.setattr("src.logic.read_sprav_file", lambda _: sprav_df)
    monkeypatch.setattr("src.logic.load_check_file", lambda _: (wb, ws))
    monkeypatch.setattr("src.logic.save_file", lambda *_: None)

    run_check(
        sprav_file="sprav.xlsx",
        check_file="check.xlsx",
        output_file="out.xlsx",
        check_dismissed=True,
        check_duplicates=True,
        prompt_duplicate_selection=False,
        log=log,
    )

    # После фильтрации: ФИО жёлтый, следующая колонка — зелёная, комментарий в «Комментарий»
    assert ELOYY_FILL_COLOR in (ws["A2"].fill.start_color.rgb or "")
    assert GREEN_FILL_COLOR in (ws["B2"].fill.start_color.rgb or "")
    assert ws["D2"].value == "Дубликат разрешён после 1-й проверки (шаг 2)"
    assert RED_FILL_COLOR in (ws["A3"].fill.start_color.rgb or "")

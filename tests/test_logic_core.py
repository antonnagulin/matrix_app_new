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

    monkeypatch.setattr("src.logic.read_sprav_file", lambda *a, **k: sprav_df)
    monkeypatch.setattr("src.logic.load_check_file", lambda *a, **k: (wb, ws, None))
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

    monkeypatch.setattr("src.logic.read_sprav_file", lambda *a, **k: sprav_df)
    monkeypatch.setattr("src.logic.load_check_file", lambda *a, **k: (wb, ws, None))
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

    # После фильтрации: ФИО жёлтый, «Дубль шаг 2» — зелёная (не пользовательские колонки)
    def _col(header: str) -> int:
        for col in range(1, (ws.max_column or 1) + 1):
            if ws.cell(row=1, column=col).value == header:
                return col
        raise AssertionError(header)

    step2 = _col("Дубль шаг 2")
    comment = _col("Комментарий")
    assert ELOYY_FILL_COLOR in (ws["A2"].fill.start_color.rgb or "")
    assert GREEN_FILL_COLOR in (ws.cell(row=2, column=step2).fill.start_color.rgb or "")
    assert ELOYY_FILL_COLOR not in (ws["B2"].fill.start_color.rgb or "")
    assert ws.cell(row=2, column=comment).value == (
        "Дубликат разрешён после 1-й проверки (шаг 2)"
    )
    assert RED_FILL_COLOR in (ws["A3"].fill.start_color.rgb or "")

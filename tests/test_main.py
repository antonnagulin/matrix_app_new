"""Тесты точки входа."""

from unittest.mock import MagicMock, patch


def test_main_starts_app_and_runs_loop():
    mock_root = MagicMock()
    with patch("src.main.tk.Tk", return_value=mock_root), patch(
        "src.main.EmployeeCheckApp"
    ) as mock_app_cls:
        from src.main import main

        main()

    mock_app_cls.assert_called_once_with(mock_root)
    mock_root.mainloop.assert_called_once()

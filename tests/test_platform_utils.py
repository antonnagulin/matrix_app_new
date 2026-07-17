"""Тесты platform_utils."""

from pathlib import Path
from unittest.mock import MagicMock, patch

from src.platform_utils import open_folder


def test_open_folder_returns_false_for_missing_dir(tmp_path):
    assert open_folder(str(tmp_path / "missing")) is False


def test_open_folder_uses_xdg_open_on_linux(tmp_path):
    folder = tmp_path / "out"
    folder.mkdir()
    with patch("src.platform_utils.sys.platform", "linux"), patch(
        "src.platform_utils.subprocess.run"
    ) as mock_run:
        mock_run.return_value = MagicMock(returncode=0)
        assert open_folder(str(folder)) is True
        mock_run.assert_called_once()
        assert mock_run.call_args.args[0][0] == "xdg-open"


def test_open_folder_fallback_message_on_failure(tmp_path):
    folder = tmp_path / "out"
    folder.mkdir()
    with patch("src.platform_utils.sys.platform", "linux"), patch(
        "src.platform_utils.subprocess.run", side_effect=FileNotFoundError
    ):
        assert open_folder(str(folder)) is False

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from .constants import RED_FILL_COLOR, ORANGE_FILL_COLOR

def read_sprav_file(file_path):
    """Чтение справочника в DataFrame"""
    return pd.read_excel(file_path, dtype=str).fillna("")

def load_check_file(file_path):
    """Открытие файла для проверки Excel"""
    wb = load_workbook(file_path)
    ws = wb.active
    return wb, ws

def save_file(wb, file_path):
    wb.save(file_path)

def get_fill(color_hex):
    return PatternFill(start_color=color_hex, end_color=color_hex, fill_type="solid")

import os 
from array import array
from openpyxl import Workbook
import pandas as pd
import numpy as np
import matplotlib as mp

def create_xlsx(filename: str, cols: int):
    try:
        wb = Workbook()
        xlsx_sheet = wb.active
        for col in range(1, cols + 100):
            xlsx_sheet.cell(row=1, column=col, value=col)
        file_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), f"{filename}.xlsx")
        wb.save(file_path)
        print(f"Файл создан и сохранен: {file_path}")
    
    except Exception as e:
        print(f"Ошибка создания файла: {e}")
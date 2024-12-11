import os 
import openpyxl
from array import array
from openpyxl import Workbook
import pandas as pd
import numpy as np
import matplotlib as mp

import numxl

def init_xlsx(filename: str, cols: int):
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

def generate_gaus_matrix(xlsx_filename:str, rows:int, cols:int, k_student_num:int):
    try:
        book = openpyxl.open(f"{os.path.dirname(os.path.realpath(__file__))}/{xlsx_filename}.xlsx", read_only=False)
        xlsx_sheet = book.active
    except:
        print(f"ошибка чтения файла {os.path.dirname(os.path.realpath(__file__))}/{xlsx_filename}.xlsx")

    try:
        for row in range(0, rows):
            xlsx_sheet[row + 2][rows].value = (20 - k_student_num)/(row + 2)
            for col in range(0, cols):
                xlsx_sheet[row + 2][col].value = (k_student_num + 1)/(row + col + 22 - k_student_num)
        book.save(f"{os.path.dirname(os.path.realpath(__file__))}/{xlsx_filename}.xlsx")
        print(f"Матрица {rows}x{cols} сгенерирована и сохранена в {os.path.dirname(os.path.realpath(__file__))}/{xlsx_filename}.xlsx")
    except Exception as e:
        print(f"ошибка записи в файл {os.path.dirname(os.path.realpath(__file__))}/{xlsx_filename}.xlsx")
        print(e)

def read_xlsx_to_matrix(xlsx_filename:str, rows:int, cols:int) -> list:
    try:
        book = openpyxl.load_workbook(f"{os.path.dirname(os.path.realpath(__file__))}/{xlsx_filename}.xlsx")
        book_sheet = book.active
    except Exception as e:
        print(f"ошибка чтения файла {os.path.dirname(os.path.realpath(__file__))}/{xlsx_filename}.xlsx")
        print(f"{e}")

    matrix = []

    for row in range(2, rows + 2):
        row_list = []
        for col in range(2, cols + 2):
            cell = book_sheet.cell(row=row, column=col)
            row_list.append(cell.value)
        matrix.append(row_list)
    print(f"матрица успешно считана")

    return matrix

def forward_elimination(matrix: list, rows: int = 6, cols: int = 6) -> list:
    """
    Выполняет прямой ход метода Гаусса для преобразования матрицы к верхнетреугольной форме.
    
    :param matrix: Двумерный список, представляющий расширенную матрицу системы линейных уравнений.
    :param rows: Количество строк в матрице.
    :param cols: Количество столбцов в матрице.
    :return: Преобразованная матрица в верхнетреугольной форме.
    """
    for row in range(0, rows - 1):
        # Находим ведущий элемент и проверяем его ненулевое значение
        pivot = matrix[row][row]
        if pivot == 0:
            # Если ведущий элемент равен 0, ищем строку для перестановки
            for swap_row in range(row + 1, rows - 1):
                if matrix[swap_row][row] != 0:
                    matrix[row], matrix[swap_row] = matrix[swap_row], matrix[row]
                    pivot = matrix[row][row]
                    break
            else:
                # Если подходящей строки не найдено, система может быть несовместимой или иметь бесконечное множество решений
                raise ValueError("Матрица вырожденная, решение не может быть найдено.")
        
        # Приводим текущую строку так, чтобы ведущий элемент стал равен 1 (опционально)
        for col in range(row, cols):
            matrix[row][col] /= pivot

        # Обнуляем элементы ниже ведущего
        for i in range(row + 1, rows - 1):
            factor = matrix[i][row]
            for col in range(row, cols - 1):
                matrix[i][col] -= factor * matrix[row][col]
    
    return matrix
    
def back_substitution(matrix: list, rows: int = 6, cols: int = 6) -> list:
    """
    Выполняет обратный ход метода Гаусса для нахождения решений.
    
    :param matrix: Верхнетреугольная матрица (расширенная матрица системы).
    :param rows: Количество строк в матрице.
    :param cols: Количество столбцов в матрице.
    :return: Список решений для системы уравнений.
    """
    # Инициализируем список для решений
    solutions = [0] * rows
    
    # Обратный ход: идем от последней строки к первой
    for i in range(rows - 1, -1, -1):
        # Свободный член b_i
        solutions[i] = matrix[i][cols - 1]
        
        # Вычитаем уже найденные x_j
        for j in range(i + 1, rows):
            solutions[i] -= matrix[i][j] * solutions[j]
        
        # Делим на ведущий элемент
        solutions[i] /= matrix[i][i]
    
    return solutions


def main():
    numxl.create_xlsx('gaus', 6)
    generate_gaus_matrix('gaus', 6, 6, 10)
    matrix = read_xlsx_to_matrix('gaus', 6, 6)

    print("[")
    print(*[str(row) for row in matrix], sep=",\n")
    print("]")

    matrix = forward_elimination(matrix=matrix, rows=6, cols=6)
    print("прямой ход [")
    print(*[str(row) for row in matrix], sep=",\n")
    print("]")

    solutions = back_substitution(matrix=matrix, rows=6, cols=6)
    print()
    print(solutions)

main()
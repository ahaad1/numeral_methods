import os 
import openpyxl
from array import array
from openpyxl import Workbook
import pandas as pd
import numpy as np
import matplotlib as mp

import numxl
    
xlsx_sheet = None
workbook = None

step_row = 2
step_col = 1

def generate_gauss_matrix(student_num: int, rows: int, cols: int) -> list:
    """
    Генерирует матрицу как список списков для метода Гаусса.

    :param student_num: Номер студента.
    :param rows: Количество строк в матрице.
    :param cols: Количество столбцов в матрице.
    :return: Матрица (список списков).
    """
    global step_row
    global xlsx_sheet

    matrix = []  # Инициализируем пустую матрицу

    for row in range(rows):  # Генерируем строки
        current_row = []  # Инициализируем текущую строку
        for col in range(cols):  # Генерируем столбцы
            current_row.append((student_num + 1) / (row + col + 22 - student_num))  # Добавляем элемент в строку
        # Добавляем значение для правой части уравнения
        current_row.append((20 - student_num) / (row + 2))
        matrix.append(current_row)  # Добавляем строку в матрицу

    numxl.write_xlsx(arr=matrix, row=step_row, col=1, sheet=xlsx_sheet, message=f"Инициализация матрицы для студента с номером {student_num}")
    step_row += 9
    return matrix

def forward_elimination(matrix: list):
    """
    Выполняет прямой ход метода Гаусса (зануление нижнего левого угла матрицы),
    включая подсчёт шагов и их сохранение.
    
    :param matrix: Двумерный список (список списков), представляющий матрицу.
    :return: Преобразованная матрица (после прямого хода).
    """

    global step_row
    global xlsx_sheet

    rows = len(matrix)         # Количество строк
    columns = len(matrix[0])   # Количество столбцов (включая столбец свободных членов)

    # Копируем исходную матрицу для работы
    matrix_clone = [row[:] for row in matrix]

    step_count = 1  # Счётчик шагов

    for k in range(rows):  # k - номер текущей строки
        # Деление k-строки на первый ненулевой элемент
        divisor = matrix_clone[k][k]
        for j in range(columns):
            matrix_clone[k][j] /= divisor

        # Запись текущего состояния после нормализации строки
        numxl.write_xlsx(
            arr=matrix_clone,
            row=step_row,
            col=1,
            sheet=xlsx_sheet,
            message=f"Прямой ход: Нормализация строки {k+1} (Шаг {step_count})"
        )
        step_row +=8
        step_count += 1

        # Зануление элементов ниже текущей строки
        for i in range(k + 1, rows):  # i - индекс строки ниже k
            factor = matrix_clone[i][k]
            for j in range(columns):  # j - индекс столбца
                matrix_clone[i][j] -= matrix_clone[k][j] * factor

            # Запись текущего состояния после зануления строки
            numxl.write_xlsx(
                arr=matrix_clone,
                row=step_row,
                col=1,
                sheet=xlsx_sheet,
                message=f"Прямой ход: Зануление строки {i+1}, столбец {k+1} (Шаг {step_count})"
            )
            step_row += 8
            step_count += 1

    return matrix_clone
    
def backward_substitution(matrix: list):
    """
    Выполняет обратный ход метода Гаусса (зануление верхнего правого угла матрицы),
    включая логирование шагов.
    
    :param matrix: Двумерный список (список списков), представляющий матрицу.
    :return: Преобразованная матрица.
    """

    global step_row
    global xlsx_sheet

    rows = len(matrix)         # Количество строк
    columns = len(matrix[0])   # Количество столбцов (включая свободные члены)

    # Копируем исходную матрицу для работы
    matrix_clone = [row[:] for row in matrix]

    step_count = 1  # Счётчик шагов

    # Обратный ход
    for k in range(rows - 1, -1, -1):  # k - номер строки (с конца к началу)
        # Нормализация строки: деление на диагональный элемент
        divisor = matrix_clone[k][k]
        for j in range(columns - 1, -1, -1):  # i - номер столбца (с конца к началу)
            matrix_clone[k][j] /= divisor

        # Логирование состояния после нормализации строки
        numxl.write_xlsx(
            arr=matrix_clone,
            row=step_row,
            col=1,
            sheet=xlsx_sheet,
            message=f"Обратный ход: Нормализация строки {k+1} (Шаг {step_count})"
        )
        step_row += 8
        step_count += 1

        # Зануление элементов выше текущей строки
        for i in range(k - 1, -1, -1):  # i - индекс строки выше k
            factor = matrix_clone[i][k]
            for j in range(columns - 1, -1, -1):  # j - индекс столбца (с конца к началу)
                matrix_clone[i][j] -= matrix_clone[k][j] * factor

            # Логирование состояния после зануления строки
            numxl.write_xlsx(
                arr=matrix_clone,
                row=step_row,
                col=1,
                sheet=xlsx_sheet,
                message=f"Обратный ход: Зануление строки {i+1}, столбец {k+1} (Шаг {step_count})"
            )
            step_row += 8
            step_count += 1

    return matrix_clone

def extract_answers(matrix: list) -> list:
    """
    Извлекает столбец ответов из расширенной матрицы (последний столбец).
    :param matrix: Двумерный список (список списков), представляющий расширенную матрицу.
    :return: Список ответов (последний столбец).
    """
    rows = len(matrix)  # Количество строк
    answers = [matrix[i][-1] for i in range(rows)]  # Извлечение последнего столбца
    numxl.write_xlsx(
                arr=answers,
                row=2,
                col=10,
                sheet=xlsx_sheet,
                message=f"Ответ:"
            )
    return answers

def gaus_solution(student_num: int = 1, rows:int = 6, cols:int = 6):
    global step_row
    global xlsx_sheet
    global workbook

    xlsx_sheet, workbook = numxl.create_xlsx(filename='gaus', cols=cols)
    extract_answers(backward_substitution(forward_elimination(generate_gauss_matrix(student_num=student_num, rows=rows, cols=cols))))
    workbook.save('gaus.xlsx')
import os 
import openpyxl
from array import array
from openpyxl import Workbook
import pandas as pd
import numpy as np
import matplotlib as mp

import numxl
    

def generate_gauss_matrix(student_num: int, rows: int, cols: int) -> list:
    """
    Генерирует матрицу как список списков для метода Гаусса.

    :param student_num: Номер студента.
    :param rows: Количество строк в матрице.
    :param cols: Количество столбцов в матрице.
    :return: Матрица (список списков).
    """
    matrix = []  # Инициализируем пустую матрицу

    for row in range(rows):  # Генерируем строки
        current_row = []  # Инициализируем текущую строку
        for col in range(cols):  # Генерируем столбцы
            current_row.append((student_num + 1) / (row + col + 22 - student_num))  # Добавляем элемент в строку
        # Добавляем значение для правой части уравнения
        current_row.append((20 - student_num) / (row + 2))
        matrix.append(current_row)  # Добавляем строку в матрицу

    

    return matrix

def forward_elimination(matrix: list):
    """
    Выполняет прямой ход метода Гаусса (зануление нижнего левого угла матрицы).
    :param matrix: Двумерный список (список списков), представляющий матрицу.
    :return: Преобразованная матрица (после прямого хода).
    """
    rows = len(matrix)         # Количество строк
    columns = len(matrix[0])   # Количество столбцов (включая столбец свободных членов)

    # Копируем исходную матрицу для работы
    matrix_clone = [row[:] for row in matrix]

    for k in range(rows):  # k - номер текущей строки
        # Деление k-строки на первый ненулевой элемент
        divisor = matrix_clone[k][k]
        for j in range(columns):
            matrix_clone[k][j] /= divisor

        # Зануление элементов ниже текущей строки
        for i in range(k + 1, rows):  # i - индекс строки ниже k
            factor = matrix_clone[i][k] / matrix_clone[k][k]
            for j in range(columns):  # j - индекс столбца
                matrix_clone[i][j] -= matrix_clone[k][j] * factor
        

    return matrix_clone
    
def backward_substitution(matrix: list):
    """
    Выполняет обратный ход метода Гаусса (зануление верхнего правого угла матрицы).
    :param matrix: Двумерный список (список списков), представляющий матрицу.
    :return: Преобразованная матрица.
    """
    rows = len(matrix)         # Количество строк
    columns = len(matrix[0])   # Количество столбцов (включая свободные члены)

    # Копируем исходную матрицу для работы
    matrix_clone = [row[:] for row in matrix]

    # Обратный ход
    for k in range(rows - 1, -1, -1):  # k - номер строки (с конца к началу)
        # Нормализация строки: деление на диагональный элемент
        for i in range(columns - 1, -1, -1):  # i - номер столбца (с конца к началу)
            matrix_clone[k][i] /= matrix[k][k]

        # Зануление элементов выше текущей строки
        for i in range(k - 1, -1, -1):  # i - индекс строки выше k
            factor = matrix_clone[i][k] / matrix_clone[k][k]
            for j in range(columns - 1, -1, -1):  # j - индекс столбца (с конца к началу)
                matrix_clone[i][j] -= matrix_clone[k][j] * factor

    return matrix_clone

def extract_answers(matrix: list) -> list:
    """
    Извлекает столбец ответов из расширенной матрицы (последний столбец).
    :param matrix: Двумерный список (список списков), представляющий расширенную матрицу.
    :return: Список ответов (последний столбец).
    """
    rows = len(matrix)  # Количество строк
    answers = [matrix[i][-1] for i in range(rows)]  # Извлечение последнего столбца
    return answers

def main():
    numxl.create_xlsx('gaus', 6)
    matrix = generate_gauss_matrix(10, 6, 6)

    print("[")
    print(*[str(row) for row in matrix], sep=",\n")
    print("]")
    forward_elimination_matrix = forward_elimination(matrix)
    
    print("forward_elimination_matrix[")
    print(*[str(row) for row in forward_elimination_matrix], sep=",\n")
    print("]")

    backward_substitution_matrix = backward_substitution(forward_elimination_matrix)
    print("backward_substitution_matrix[")
    print(*[str(row) for row in backward_substitution_matrix], sep=",\n")
    print("]")

    answer_matrix = extract_answers(backward_substitution_matrix)
    print("backward_substitution_matrix[")
    print(*[str(row) for row in answer_matrix], sep=",\n")
    print("]")

main()
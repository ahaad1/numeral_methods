import random
from array import array
import copy
import pandas as pd
import numpy as np
import matplotlib as mp
import openpyxl
from openpyxl import load_workbook
import numxl



def generate_random_system_for_iteration(size: int):

    matrix = np.zeros((size, size+1) , dtype=float)

    for row in range(size):
        for col in range(size):
            matrix[row][col] = np.random.uniform(-10, 10)  # Генерация случайного числа от -10 до 10
        
        # Вычисление суммы модулей всех элементов строки, кроме диагонального
        non_diagonal_sum = sum(abs(matrix[row][col]) for col in range(size) if col != row)
        
        # Установка диагонального элемента
        matrix[row][row] = non_diagonal_sum + np.random.uniform(1, 5)  # Диагональный элемент больше суммы модулей остальных
        
        # Генерация элемента столбца правых частей
        matrix[row][size] = np.random.uniform(-10, 10)

    return matrix     



def solve_iteration_method(size:int , epsilon: float):

    numxl.create_xlsx("iteration_metod" , 100)
    book = openpyxl.open("iteration_metod.xlsx", read_only=False)
    sheet = book.active

    position = 1;
    matrix = generate_random_system_for_iteration(size)
    numxl.write_xlsx(matrix , position , 1 , sheet , "Инициализированная матрица:")
    position += size + 2


    iter_solve = np.zeros(size , dtype=float)

    for row in range (0 , size):
        iter_solve[row] = matrix[row][size]/matrix[row][row]

    numxl.write_xlsx(iter_solve , position , 1 , sheet , "Нулевое приближение:" )
    position += size + 2
    
    err = np.ones(size , dtype=float)
    iter = 1 
    while any(err > epsilon):
        solve_clone = copy.deepcopy(iter_solve)  # Копируем текущее приближение

        for row in range(size):
            buff = 0

            for col in range(size):
                if col != row:  # Исключаем диагональный элемент
                    buff += solve_clone[col] * matrix[row][col]

            # Обновляем текущее значение для строки
            iter_solve[row] = (matrix[row][size] - buff) / matrix[row][row]
            err[row] = abs(iter_solve[row] - solve_clone[row])/abs(iter_solve[row])
        
        numxl.write_xlsx(iter_solve , position , 1 , sheet , f"{iter}-oe приближение:" )
        numxl.write_xlsx(err , position , 4 , sheet , f"погрешность {iter}-го приближения:" )
        position += size + 2
        iter = iter + 1
        
        numxl.write_xlsx(iter_solve , 1 , size+3 , sheet , "Решение:" )
       
    book.save("iteration_metod.xlsx")




solve_iteration_method(10 , 1E-12)




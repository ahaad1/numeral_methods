import random
from array import array
import copy
import pandas as pd
import numpy as np
import matplotlib as mp
import openpyxl
from openpyxl import load_workbook
import numxl




def generate_symetrical_matrix(size: int , random_max_value: int):

    generated_matrix = np.zeros((size , size) , dtype=float)

    for row in range (size):
        for col in range (row , size):
            value = np.random.uniform(- random_max_value , random_max_value)
            generated_matrix[row][col] = value 
            generated_matrix[col][row] = value 

    return generated_matrix


def check_nondiagonal_elements(matrix: np.ndarray, epsilon: float) -> bool:
    """
    Проверяет, являются ли все недиагональные элементы квадратной матрицы меньше заданного значения epsilon.
    
    :param matrix: Квадратная матрица (NumPy массив).
    :param epsilon: Пороговое значение.
    :return: True, если все недиагональные элементы меньше epsilon, иначе False.
    """ 
    size = matrix.shape[0]

    for i in range(size):
        for j in range(size):
            if i != j and abs(matrix[i, j]) >= epsilon:  # Проверка недиагонального элемента
                return False  # Найден элемент, который >= epsilon

    return True


def yakobi_method(size: int , epsilon: float ,  random_max_value: int):

    numxl.create_xlsx("7_yakobi_metod" , 100)
    book = openpyxl.open("7_yakobi_metod.xlsx", read_only=False)
    sheet = book.active

    position = 1;
    matrix = generate_symetrical_matrix(size , random_max_value)

    numxl.write_xlsx(matrix , position , 1 , sheet , "Инициализированная матрица:")
    position += size + 2
    iteration = 1
    while not check_nondiagonal_elements(matrix , epsilon):
        max_value = float('-inf')  # Инициализация минимальным возможным значением
        max_i = 0;
        max_j = 0;

        for i in range(size):
            for j in range(size):
                if i != j:  # Условие для пропуска диагональных элементов
                    if abs(matrix[i, j]) > max_value:
                        max_value = abs(matrix[i, j])
                        max_i = i;
                        max_j = j;
   
        phi = np.arctan2(2*matrix[max_i][max_j] , matrix[max_i][max_i]-matrix[max_j][max_j])/2
        
        rotation_matrix = np.eye(size , dtype=float)
        rotation_matrix[max_i][max_i] = rotation_matrix[max_j][max_j] = np.cos(phi)
        rotation_matrix[max_i][max_j] = -np.sin(phi)
        rotation_matrix[max_j][max_i] = np.sin(phi)

        matrix = rotation_matrix.T @ matrix @ rotation_matrix
        numxl.write_xlsx(matrix , position , 1 , sheet , f"{iteration}-я иттерация")
        iteration +=1 
        position += size + 2

    result = np.zeros((size) , dtype=float)
    for lamb in range(size):
        result[lamb] = matrix[lamb][lamb]

    numxl.write_xlsx(result , 1 , size+2 , sheet , "Собственные значения:")
    book.save("7_yakobi_metod.xlsx")



yakobi_method(5, 1E-12 , 10)


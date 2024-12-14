import random
from array import array

import pandas as pd
import numpy as np
import matplotlib as mp
import openpyxl
from openpyxl import load_workbook
import numxl

def rounding ( number : float , rounding: float):
    if abs(number) < rounding :
        return 0
    else:
        return number 


def print_rotation_variant( students_var : int , size : int):
    matrix = np.zeros((size , size+1) , dtype=float)
    try:
        for row in range (0,  size):
            matrix[row][size] = (students_var + 2)/(row + 1)
            for col in range (0 , size):
                matrix[row][col] = (20 - students_var)/(row + col + 19 - students_var)
        return matrix
     
    except Exception as e:
        print(f"Ошибка генерации матрицы: {e}")




def rotation_method( students_var : int  , size : int ):

    numxl.create_xlsx("rotation_metod" , 100)
    book = openpyxl.open("rotation_metod.xlsx", read_only=False)
    sheet = book.active

    position = 1;
    matrix = print_rotation_variant(students_var , size)
    numxl.write_xlsx(matrix , position , 1 , sheet , "Инициализированная матрица:")
    position += size + 2
    book.save("rotation_metod.xlsx")

    cos_and_sin = np.array((1 , 2) , dtype=float)
    j = 0;
    buf1 = 0;
    buf2 = 0;
    for col  in range (0 , size - 1):
        j+=1;
        for row in range (j , size ):
            cos_and_sin [0] = matrix[col][col] / np.sqrt(pow(matrix[col][col] , 2) + pow(matrix[row][col] , 2))
            cos_and_sin [1] = matrix[row][col] / np.sqrt(pow(matrix[col][col] , 2) + pow(matrix[row][col] , 2))
            numxl.write_xlsx(cos_and_sin , position , 1 , sheet , " Значения косинуса и синуса:" )
            position += 4
            for i in range ( col , size + 1):
                buf1 = matrix[col][i]
                buf2 = matrix[row][i]
                matrix[col][i] = rounding( cos_and_sin[0]*buf1 + cos_and_sin[1]*buf2 , 1E-17)
                matrix[row][i] =rounding(-cos_and_sin[1]*buf1 + cos_and_sin[0]*buf2 , 1E-17)
            
            numxl.write_xlsx(matrix , position , 1 ,sheet , f"Зануляем {row+1 , col+1} элемент:" )
            position = position + size + 2


    result = np.zeros((size , 1) , dtype=float)
    result[size-1][0] = matrix[size-1][size]/matrix[size-1][size-1]
    for xi in reversed(range(0 , size-1)):
        x = 0
        for j in range(1 , size-xi):
            x = x + matrix[xi][size-j]*result[size-j][0]
        result[xi][0] = (matrix[xi][size] - x)/matrix[xi][xi]


    numxl.write_xlsx(result , 1 , size+3 , sheet , "Результат:")       
    book.save("rotation_metod.xlsx")
    

rotation_method(10 , 6)
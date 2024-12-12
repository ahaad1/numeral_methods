import random
from array import array

import pandas as pd
import numpy as np
import matplotlib as mp
import openpyxl
from numpy.ma.core import shape
from openpyxl import load_workbook
import numxl



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
    
    for col  in range (0 , size + 1):
        for row in range (1 , size ):
            cos_and_sin [0] = matrix[row-1][col] / np.sqrt(pow(matrix[row-1][col] , 2) + pow(matrix[row][col] , 2))
            cos_and_sin [1] = matrix[row][col] / np.sqrt(pow(matrix[row-1][col] , 2) + pow(matrix[row][col] , 2))
            numxl.write_xlsx(cos_and_sin , position , 1 , sheet , " Значения косинуса и синуса:" )
            position += 5
    
    book.save("rotation_metod.xlsx")
    

rotation_method(5 , 6)
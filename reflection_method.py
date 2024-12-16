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


def print_reflection_variant( students_var : int , size : int):
    matrix = np.zeros((size , size+1) , dtype=float)
    try:
        for row in range (0,  size):
            matrix[row][size] = (students_var + 3)/(3*row + 1)
            for col in range (0 , size):
                matrix[row][col] = (18 + row - students_var)/(2*row+col+29-students_var)
        return matrix
     
    except Exception as e:
        print(f"Ошибка генерации матрицы: {e}")




def reflection_method( students_var : int  , size : int ):

    numxl.create_xlsx("reflection_metod" , 100)
    book = openpyxl.open("reflection_metod.xlsx", read_only=False)
    sheet = book.active

    position = 1;
    matrix = print_reflection_variant(students_var , size)
    numxl.write_xlsx(matrix , position , 1 , sheet , "Инициализированная матрица:")
    position += size + 2
    book.save("reflection_metod.xlsx")

    


reflection_method(10 , 6)
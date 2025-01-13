import numpy as np
import openpyxl
from openpyxl import Workbook
import numxl


def generate_symetrical_matrix(size: int , random_max_value: int):

    generated_matrix = np.zeros((size , size) , dtype=float)

    for row in range (size):
        for col in range (row , size):
            value = np.random.uniform(- random_max_value , random_max_value)
            generated_matrix[row][col] = value 
            generated_matrix[col][row] = value 

    return generated_matrix



def yakobi_lr_method(size: int , epsilon: float ,  random_max_value: int):

    numxl.create_xlsx("8_yakobi_lr_metod" , 100)
    book = openpyxl.open("8_yakobi_lr_metod.xlsx", read_only=False)
    sheet = book.active

    position = 1;
    matrix = generate_symetrical_matrix(size , random_max_value)

    numxl.write_xlsx(matrix , position , 1 , sheet , "Инициализированная матрица:")
    position += size + 2
    iteration = 1



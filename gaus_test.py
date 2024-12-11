import random
from array import array

import pandas as pd
import numpy as np
import matplotlib as mp
import openpyxl
from numpy.ma.core import shape
from openpyxl import load_workbook
import numxl



def approx_machine(number):
    if abs(number) < 1E-13 :
        return 0
    return number

def get_matrix(start_position , shape , sheet):
    '''
    Принимает два параметра:
    start_position - начальная строка для считывания
    shape - пара i;j размера матрицы
    Возвращает:
    Матрицу i на j , заполненную элементами из .xlsx
    '''
    i = shape[0]
    j = shape[1]

    arr = np.zeros((i , j))
    for row in range (start_position , start_position+i+1):
        for col in range (0 , j):
            arr[row-2][col] = sheet[row][col].value
    return arr




def print_matrix(start_position , arr , sheet):
    shape_arr = shape(arr)
    print(shape_arr)
    for row in range (0 , shape_arr[0]):
        for col in range(0 , shape_arr[1]):
            sheet[row+start_position][col].value = approx_machine(arr[row][col])
    return



def gauss_method(arr , sheet):
    mx = arr[0]
    mx_size = arr[1]
    write_row = mx_size+3

    for step in range (1 , mx_size):
        for i in range (step , mx_size):
            coef = -mx[i][step-1]/mx[step-1][step-1]
            for j in range (0 , mx_size+1):
              mx[i][j] = mx[i][j] + coef*mx[step-1][j]
        print_matrix(write_row , mx , sheet)
        write_row = write_row + mx_size + 2

    result = np.zeros((mx_size , 1) , dtype=float)
    result[mx_size-1][0] = mx[mx_size-1][mx_size]/mx[mx_size-1][mx_size-1]
    for xi in reversed(range(0 , mx_size-1)):
        x = 0

        for j in range(1 , mx_size-xi):
            x = x + mx[xi][mx_size-j]*result[mx_size-j][0]

        result[xi][0] = (mx[xi][mx_size] - x)/mx[xi][xi]

    print_matrix(write_row, result , sheet)
    return result

def print_newton_variant(k , sheet):
    for row in range (0 , 6):
        sheet[row + 2][6].value = (20 - k)/( row + 2)
        for col in range (0 , 6):
            sheet[row + 2][col].value = (k+1)/(row+col+22-k)
    return 1


def generate_random_system(start_position , size , sheet):

    for row in range(0, size):
        for col in range(0, size+1):
            print(col)
            sheet[row + start_position][col].value = random.random()*random.randint(0,100)
    return 1


def main():
    numxl.create_xlsx("new_table" , 100)
    book = openpyxl.open("new_table.xlsx", read_only=False)
    sheet = book.active
    #generate_random_system(2 , 19 , sheet)
    print_newton_variant(15 , sheet)
    book.save("new_table.xlsx")
    arr = get_matrix(1 , (6 , 7) ,sheet)
    res = gauss_method((arr , 6) , sheet)

    book.save("new_table.xlsx")

main()
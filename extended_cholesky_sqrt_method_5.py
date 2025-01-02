import math
import numxl

xlsx_sheet = None
workbook = None
step_row = 2

def generate_matrix_and_vector(k: int):
    global step_row, xlsx_sheet

    A = [[0.0] * 6 for _ in range(6)]
    b = [0.0] * 6

    A[0] = [40-3*k, 35-2*k, 50-3*k, 19-k, 44-7*k, 35]
    A[1][1:] = [k+4, 19-k, 47-2*k, 29-4*k, 46]
    A[2][2:] = [k+2, 58, 64-7*k, 53-2*k]
    A[3][3:] = [43-5*k, 4*k + 69, 67-6*k]
    A[4][4:] = [27, 71-10*k]
    A[5][5] = k+7

    for i in range(6):
        for j in range(i + 1, 6):
            A[j][i] = A[i][j]

    b[0] = 218-7*k
    b[1] = 1118-15*k
    b[2] = 11719-7*k
    b[3] = 57849-14*k
    b[4] = 12954-3*k
    b[5] = 57164-11*k

    numxl.write_xlsx(arr=A, row=step_row, col=1, sheet=xlsx_sheet, message="Матрица A:")
    step_row += len(A) + 2
    numxl.write_xlsx(arr=[[bi] for bi in b], row=step_row, col=1, sheet=xlsx_sheet, message="Вектор b:")
    step_row += len(b) + 2

    return A, b

def ldl_decomposition(A):
    global step_row, xlsx_sheet

    n = len(A)
    L = [[0.0] * n for _ in range(n)]
    D = [0.0] * n

    for i in range(n):
        for j in range(i):
            sum_LDL = sum(L[i][k] * D[k] * L[j][k] for k in range(j))
            L[i][j] = (A[i][j] - sum_LDL) / D[j]
        sum_LDL = sum(L[i][k] * D[k] * L[i][k] for k in range(i))
        D[i] = A[i][i] - sum_LDL
        L[i][i] = 1.0
        numxl.write_xlsx(arr=L, row=step_row, col=1, sheet=xlsx_sheet, message=f"Шаг {i+1}: L")
        numxl.write_xlsx(arr=[[di] for di in D], row=step_row, col=8, sheet=xlsx_sheet, message=f"Шаг {i+1}: D")
        step_row += n + 2
    return L, D

def forward_substitution(L, b):
    n = len(b)
    y = [0.0] * n
    for i in range(n):
        sum_Ly = sum(L[i][j] * y[j] for j in range(i))
        y[i] = b[i] - sum_Ly
    return y

def diagonal_substitution(D, y):
    return [y[i] / D[i] for i in range(len(y))]

def backward_substitution(L, y):
    n = len(y)
    x = [0.0] * n
    for i in reversed(range(n)):
        sum_Lx = sum(L[j][i] * x[j] for j in range(i + 1, n))
        x[i] = y[i] - sum_Lx
    return x

def cholesky_solution(k: int = 2):
    global xlsx_sheet, workbook, step_row
    xlsx_sheet, workbook = numxl.create_xlsx(filename='cholesky_manual_5', cols=10)

    A, b = generate_matrix_and_vector(k)

    print("Матрица A:")
    for row in A:
        print(row)

    L, D = ldl_decomposition(A)

    y = forward_substitution(L, b)
    numxl.write_xlsx(arr=[[yi] for yi in y], row=2, col=8, sheet=xlsx_sheet, message="Решение Ly = b (y):")
    step_row += len(y) + 3

    z = diagonal_substitution(D, y)
    numxl.write_xlsx(arr=[[zi] for zi in z], row=2, col=10, sheet=xlsx_sheet, message="Решение Dz = y (z):")
    step_row += len(z) + 3

    x = backward_substitution(L, z)
    numxl.write_xlsx(arr=[[xi] for xi in x], row=2, col=12, sheet=xlsx_sheet, message="Решение L^T x = z (x):")

    workbook.save('cholesky_manual_5.xlsx')
    print("Решение сохранено в 'cholesky_manual_5.xlsx'")
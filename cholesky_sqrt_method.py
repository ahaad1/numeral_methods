import math
import numxl

xlsx_sheet = None
workbook = None
step_row = 2


def generate_matrix_and_vector(k: int):
    """
    Генерирует симметричную матрицу A и вектор b вручную.
    """
    global step_row, xlsx_sheet

    # Инициализируем пустую матрицу A и вектор b
    A = [[0.0] * 6 for _ in range(6)]
    b = [0.0] * 6

    # Заполняем верхний треугольник матрицы A
    A[0] = [5*k**2 + 14*k + 40, 12*k + 35, 19*k + 35, 11*k + 35, 9*k + 44, 16*k + 35]
    A[1][1:] = [16*k**2 + 40*k + 47, 13*k + 39, 10*k + 47, 4*k + 44, 16*k + 46]
    A[2][2:] = [9*k**2 + 24*k + 55, 15*k + 58, 9*k + 64, 11*k + 53]
    A[3][3:] = [10*k**2 + 12*k + 82, 4*k + 69, 6*k + 67]
    A[4][4:] = [80, 10*k + 71]
    A[5][5] = 4*k**2 + 12*k + 75

    # Симметрично заполняем нижний треугольник
    for i in range(6):
        for j in range(i + 1, 6):
            A[j][i] = A[i][j]

    # Формируем вектор b
    b[0] = 5*k**2 + 8*k + 218
    b[1] = -32*k**2 + 16*k + 9818
    b[2] = 27*k**2 + 7*k + 20323
    b[3] = -3*k**2 + 9979*k + 97327
    b[4] = 3*k + 20414
    b[5] = 8*k**2 + 67*k + 339

    # Печать для проверки
    print("Матрица A:")
    for row in A:
        print(row)
    print("Вектор b:", b)

    # Записываем A и b в Excel
    numxl.write_xlsx(arr=A, row=step_row, col=1, sheet=xlsx_sheet, message="Матрица A:")
    step_row += len(A) + 2  # Смещение строки для записи вектора b
    numxl.write_xlsx(arr=[[bi] for bi in b], row=step_row, col=1, sheet=xlsx_sheet, message="Вектор b:")
    step_row += len(b) + 2  # Смещение строки для следующей записи

    return A, b


def cholesky_decomposition(A):
    """
    Выполняет разложение Холецкого A = LL^T вручную.
    """
    global step_row, xlsx_sheet

    n = len(A)
    L = [[0.0] * n for _ in range(n)]  # Пустая нижняя треугольная матрица

    for i in range(n):
        for j in range(i + 1):
            sum_L = sum(L[i][k] * L[j][k] for k in range(j))
            if i == j:
                L[i][j] = math.sqrt(A[i][i] - sum_L)  # Диагональные элементы
            else:
                L[i][j] = (A[i][j] - sum_L) / L[j][j]  # Вне диагонали
        # Записываем текущую матрицу L
        numxl.write_xlsx(arr=L, row=step_row, col=1, sheet=xlsx_sheet, message=f"Шаг {i+1}: L")
        step_row += n + 2
    return L


def forward_substitution(L, b):
    """
    Решает Ly = b вручную (прямой ход).
    """
    n = len(b)
    y = [0.0] * n
    for i in range(n):
        sum_Ly = sum(L[i][j] * y[j] for j in range(i))
        y[i] = (b[i] - sum_Ly) / L[i][i]
    return y


def backward_substitution(L, y):
    """
    Решает L^T x = y вручную (обратный ход).
    """
    n = len(y)
    x = [0.0] * n
    for i in reversed(range(n)):
        sum_Lx = sum(L[j][i] * x[j] for j in range(i + 1, n))
        x[i] = (y[i] - sum_Lx) / L[i][i]
    return x


def cholesky_solution(k: int = 2):
    """
    Решение системы методом квадратного корня без NumPy.
    """
    global xlsx_sheet, workbook, step_row
    xlsx_sheet, workbook = numxl.create_xlsx(filename='cholesky_manual', cols=10)

    # Генерируем матрицу A и вектор b
    A, b = generate_matrix_and_vector(k)

    # Разложение Холецкого
    L = cholesky_decomposition(A)

    # Решаем Ly = b
    y = forward_substitution(L, b)
    numxl.write_xlsx(arr=[[yi] for yi in y], row=2, col=8, sheet=xlsx_sheet, message="Решение Ly = b (y):")
    step_row += len(y) + 3

    # Решаем L^T x = y
    x = backward_substitution(L, y)
    numxl.write_xlsx(arr=[[xi] for xi in x], row=2, col=10, sheet=xlsx_sheet, message="Решение L^T x = y (x):")

    # Сохраняем результат
    workbook.save('cholesky_manual.xlsx')
    print("Решение сохранено в 'cholesky_manual.xlsx'")
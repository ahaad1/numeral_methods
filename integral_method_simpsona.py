import numpy as np

def f(x):
    return np.exp(x) * np.cos(10 * x)

def simpson_integral(func, a, b, n):
    h = (b - a) / n
    x = np.linspace(a, b, n + 1)
    y = func(x)
    integral = h / 3 * (y[0] + 2 * np.sum(y[2:n:2]) + 4 * np.sum(y[1:n:2]) + y[n])
    return integral

def find_optimal_n(func, a, b, epsilon):
    # Начальное количество разбиений
    n_low = 2  # Минимальное возможное число разбиений
    n_high = 4  # Начальное большее число разбиений
    
    # Определяем начальное значение интегралов
    integral_low = simpson_integral(func, a, b, n_low)
    integral_high = simpson_integral(func, a, b, n_high)
    
    # Увеличиваем верхнюю границу, пока не найдем интервал, содержащий решение
    while abs(integral_high - integral_low) >= epsilon:
        n_low = n_high
        n_high *= 2
        integral_low = integral_high
        integral_high = simpson_integral(func, a, b, n_high)
    
    # Дихотомический поиск
    while n_high - n_low > 2:
        n_mid = (n_low + n_high) // 2
        integral_mid = simpson_integral(func, a, b, n_mid)
        
        if abs(integral_mid - integral_low) < epsilon:
            n_high = n_mid
            integral_high = integral_mid
        else:
            n_low = n_mid
            integral_low = integral_mid
    
    # Финальное уточнение
    return n_high, simpson_integral(func, a, b, n_high)

# Задаем параметры
a = 0
b = np.pi
epsilon = 1e-07  # Заданная погрешность

# Вычисляем минимальное n и значение интеграла
n, integral_value = find_optimal_n(f, a, b, epsilon)

print(f"Минимальное количество разбиений: {n}")
print(f"Значение интеграла: {integral_value}")
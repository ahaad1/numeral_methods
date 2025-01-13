import numpy as np

def func(x):
    return np.sin(10*x)


def solution():
    Ih = 0
    a = 0
    b = np.pi/2
    n = 12
    h = (b-a)/n
    x = 0
    xh2 = 0
    n2 = 24
    sum = 0

    for step in range (n):
        xh2 = x + h/2
        Ih = func(xh2)
        x = x + h
        sum += Ih

    print(sum*h)


    sum = 0 
    x = 0 
    h = (b-a)/n2

    for step in range (n2):
        xh2 = x + h/2
        Ih = func(xh2)
        x = x + h
        sum+= Ih

    print(Ih)




solution()
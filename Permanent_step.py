import numpy as np
import xlrd, xlwt
from xlutils.copy import copy
book = xlrd.open_workbook('test.xls')
sheet = book.sheet_by_index(0)
write_book = copy(book)
sheet1 = write_book.get_sheet(0)
sheet1.write(0,18,'Градиентный спуск с постоянным шагом')
sheet1.write(1,18,'Номер итерации')
sheet1.write(1,19,'x')
sheet1.write(1,20,'y')
sheet1.write(1,21,'f(x,y)')
sheet1.write(1, 22, 'Шаг')
sheet1.write(1,23,'Количество вызовов функции')
sheet1.write(1,24,'Количество вызовов градиента')
sheet1.write(1,25,'Действительная часть градиента')
sheet1.write(1,26,'Мнимая часть градиента')
li = complex(0, 1)
polynomial = np.array([1, complex(8, 5), complex(5, 8)])


def pol_value(z, polynomial):
    value = complex(0, 0)
    for i in range(len(polynomial)):
        value = value * z + polynomial[i]
    return value


def abs_pol_value(z, polynomial):
    pol = pol_value(z, polynomial)
    return abs(pol) * abs(pol)


def deriv(z, polynomial):
    deriv_value = complex(0, 0)
    n = len(polynomial) - 1
    for i in range(n):
        deriv_value = deriv_value * z + polynomial[i] * (n - i)
    return deriv_value


def gradient(z, polynomial):
    global f_prev
    pol = pol_value(z, polynomial)
    dpol = deriv(z, polynomial)
    grad = 2.0 * ((dpol * np.conj(pol)).real - li * (dpol * np.conj(pol)).imag)
    f_prev = (abs(pol) * abs(pol)).real
    return grad


f_prev = 0.0


def grad_permanent_step(z, poly, epsilon):
    z_prev = z
    grad = gradient(z, poly)
    a = 0.0009
    iters = 0
    grad_calls = 1
    func_calls = 1
    sheet1.write(iters + 2, 18, iters)
    sheet1.write(iters + 2, 19, z_prev.real)
    sheet1.write(iters + 2, 20, z_prev.imag)
    sheet1.write(iters + 2, 21, f_prev)
    sheet1.write(iters + 2, 22, a)
    sheet1.write(iters + 2, 23, func_calls)
    sheet1.write(iters + 2, 24, grad_calls)
    sheet1.write(iters + 2, 25, grad.real)
    sheet1.write(iters + 2, 26, grad.imag)
    while (f_prev > epsilon):
        iters += 1
        z_prev -= a*grad
        grad = gradient(z_prev, poly)
        grad_calls += 1
        func_calls +=1
        sheet1.write(iters + 2, 18, iters)
        sheet1.write(iters + 2, 19, z_prev.real)
        sheet1.write(iters + 2, 20, z_prev.imag)
        sheet1.write(iters + 2, 21, f_prev)
        sheet1.write(iters + 2, 22, a)
        sheet1.write(iters + 2, 23, func_calls)
        sheet1.write(iters + 2, 24, grad_calls)
        sheet1.write(iters + 2, 25, grad.real)
        sheet1.write(iters + 2, 26, grad.imag)
    return z_prev.real, z_prev.imag
print(grad_permanent_step(complex(0, 0), polynomial , 0.00000000001))
write_book.save("test.xls")
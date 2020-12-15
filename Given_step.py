import numpy as np
import xlrd, xlwt
from xlutils.copy import copy
book = xlrd.open_workbook('test.xls')
sheet = book.sheet_by_index(0)
write_book = copy(book)
sheet1 = write_book.get_sheet(0)
sheet1.write(0,28,'Градиентный спуск с постоянным шагом')
sheet1.write(1,28,'Номер итерации')
sheet1.write(1,29,'x')
sheet1.write(1,30,'y')
sheet1.write(1,31,'f(x,y)')
sheet1.write(1,32, 'Шаг')
sheet1.write(1,33,'Количество вызовов функции')
sheet1.write(1,34,'Количество вызовов градиента')
sheet1.write(1,35,'Действительная часть градиента')
sheet1.write(1,36,'Мнимая часть градиента')
li = complex(0, 1)
polynomial = np.array([1, complex(8, 5), complex(5, 8)])

def pol_value(z, polynomial):
    value = complex(0, 0)
    for i in range(len(polynomial)):
        value = value*z+polynomial[i]
    return value

def abs_pol_value(z, polynomial):
    pol = pol_value(z, polynomial)
    return abs(pol)*abs(pol)

def deriv(z, polynomial):
    deriv_value = complex(0, 0)
    n = len(polynomial) - 1
    for i in range(n):
        deriv_value = deriv_value*z + polynomial[i] * (n - i)
    return deriv_value

def gradient(z, polynomial):
    global f_prev
    pol = pol_value(z, polynomial)
    dpol = deriv(z, polynomial)
    grad = 2.0*((dpol*np.conj(pol)).real - li*(dpol*np.conj(pol)).imag)
    f_prev = (abs(pol)*abs(pol)).real
    return grad
f_prev = 0.0
def grad_given_step(z, poly, epsilon):
    z_prev = z
    grad = gradient(z_prev, poly)
    a = 1000.0
    iters = 0
    grad_calls = 1
    func_calls = 1
    sheet1.write(iters + 2, 28, iters)
    sheet1.write(iters + 2, 29, z_prev.real)
    sheet1.write(iters + 2, 30, z_prev.imag)
    sheet1.write(iters + 2, 31, f_prev)
    sheet1.write(iters + 2, 32, a)
    sheet1.write(iters + 2, 33, func_calls)
    sheet1.write(iters + 2, 34, grad_calls)
    sheet1.write(iters + 2, 35, grad.real)
    sheet1.write(iters + 2, 36, grad.imag)
    while (f_prev > epsilon):
        iters += 1
        a += 1
        z_prev -= 1 / a*grad
        grad = gradient(z_prev, poly)
        grad_calls += 1
        func_calls += 1
        sheet1.write(iters + 2, 28, iters)
        sheet1.write(iters + 2, 29, z_prev.real)
        sheet1.write(iters + 2, 30, z_prev.imag)
        sheet1.write(iters + 2, 31, f_prev)
        sheet1.write(iters + 2, 32, a)
        sheet1.write(iters + 2, 33, func_calls)
        sheet1.write(iters + 2, 34, grad_calls)
        sheet1.write(iters + 2, 35, grad.real)
        sheet1.write(iters + 2, 36, grad.imag)
    return z_prev.real, z_prev.imag
print(grad_given_step(complex(0, 0), polynomial , 0.00000000001))
write_book.save("test.xls")
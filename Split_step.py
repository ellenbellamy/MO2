import numpy as np
import xlrd, xlwt
from xlutils.copy import copy
book = xlrd.open_workbook('test.xls')
sheet = book.sheet_by_index(0)
write_book = copy(book)
sheet1 = write_book.get_sheet(0)
sheet1.write(0,8,'Градиентный спуск с дроблением')
sheet1.write(1,8,'Номер итерации')
sheet1.write(1,9,'x')
sheet1.write(1,10,'y')
sheet1.write(1,11,'f(x,y)')
sheet1.write(1, 12, 'Шаг')
sheet1.write(1,13,'Количество вызовов функции')
sheet1.write(1,14,'Количество вызовов градиента')
sheet1.write(1,15,'Действительная часть градиента')
sheet1.write(1,16,'Мнимая часть градиента')
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

def grad_split_step(z, pol, epsilon):
    global f_prev
    a = 0.5
    delta = 0.5
    z_prev = z
    grad = gradient(z, pol)
    func_calls = 1
    grad_calls = 1
    iters = 0
    sheet1.write(iters+2, 8, iters)
    sheet1.write(iters+2, 9, z_prev.real)
    sheet1.write(iters+2, 10, z_prev.imag)
    sheet1.write(iters+2, 11, f_prev)
    sheet1.write(iters + 2, 12, a)
    sheet1.write(iters+2, 13, func_calls)
    sheet1.write(iters+2, 14, grad_calls)
    sheet1.write(iters+2, 15, grad.real)
    sheet1.write(iters + 2, 16, grad.imag)
    while (f_prev > epsilon):
        while (True):
            iters += 1
            z_cur = z_prev - a * grad
            f_cur = abs_pol_value(z_cur, pol)
            func_calls += 1
            if ((f_cur - f_prev) <= -a * delta * ((grad).real **2 + (grad).imag **2)):
                z_prev = z_cur
                f_prev = f_cur
                sheet1.write(iters + 2, 8, iters)
                sheet1.write(iters + 2, 9, z_cur.real)
                sheet1.write(iters + 2, 10, z_cur.imag)
                sheet1.write(iters + 2, 11, f_cur)
                sheet1.write(iters + 2, 12, a)
                sheet1.write(iters + 2, 13, func_calls)
                sheet1.write(iters + 2, 14, grad_calls)
                sheet1.write(iters+2, 15, grad.real)
                sheet1.write(iters + 2, 16, grad.imag)
                break
            else:
                a /= 2
                sheet1.write(iters + 2, 8, iters)
                sheet1.write(iters + 2, 9, z_cur.real)
                sheet1.write(iters + 2, 10, z_cur.imag)
                sheet1.write(iters + 2, 11, f_cur)
                sheet1.write(iters + 2, 12, a)
                sheet1.write(iters + 2, 13, func_calls)
                sheet1.write(iters + 2, 14, grad_calls)
                sheet1.write(iters + 2, 15, grad.real)
                sheet1.write(iters + 2, 16, grad.imag)

        grad = gradient(z_prev, pol)
        grad_calls += 1
        func_calls += 1
        a *= 2

    return (z_cur).real, (z_cur).imag
print(grad_split_step(complex(0, 0), polynomial , 0.000000001))
write_book.save("test.xls")
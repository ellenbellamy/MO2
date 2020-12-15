from numpy import *
from functools import partial
from scipy.optimize import minimize
import xlrd, xlwt
from xlutils.copy import copy
book = xlrd.open_workbook('test.xls')
sheet = book.sheet_by_index(0)
write_book = copy(book)
sheet1 = write_book.get_sheet(0)
sheet1.write(0,0,'Координатный спуск')
sheet1.write(1,0,'Номер итерации')
sheet1.write(1,1,'x')
sheet1.write(1,2,'y')
sheet1.write(1,3,'f(x,y)')
sheet1.write(1,4,'Количество вызовов функции')

def h():
  return poly1d([1, 8 + 5j, 5 + 8j])

def h_conj():
  return poly1d(conj(h()))

count_p = 0
def p(x, y):
  global count_p
  count_p += 1
  return h()(x + y*1j)

count_p_conj = 0
def p_conj(x, y):
  global count_p_conj
  count_p_conj += 1
  return h_conj()(x - y*1j)

def h():
  return poly1d([1, 8 + 5j, 5 + 8j])

def h_conj():
  return poly1d(conj(h()))

def function(x, y):
    return real(p(x,y) * p_conj(x,y))

def my_minimize(function):

    eps = 0.000000001

    y = 10000
    x = -10000

    x_prev = 9
    y_prev = 9
    iters = 0
    sheet1.write(iters + 2, 0, iters)
    sheet1.write(iters + 2, 1, x)
    sheet1.write(iters + 2, 2, y)
    sheet1.write(iters + 2, 3, function(x,y))
    sheet1.write(iters + 2, 4, count_p)
    while sqrt((x - x_prev)**2 + (y - y_prev)**2) > eps:
        x_prev = x
        y_prev = y
        iters += 1
        x = minimize(partial(function, y = y), 0, method='nelder-mead').x[0]
        y = minimize(partial(function, x), 0, method='nelder-mead').x[0]
        sheet1.write(iters + 2, 0, iters)
        sheet1.write(iters + 2, 1, x)
        sheet1.write(iters + 2, 2, y)
        sheet1.write(iters + 2, 3, function(x, y))
        sheet1.write(iters + 2, 4, count_p)
    return x, y, iters

min = my_minimize(function)

print(min)
write_book.save("test.xls")

from deptControl import *
#from columnRevise import *

def main():
    print()
    print('Функционал программы:')
    print("1. проверка даты выполнения паботы")
    print("2. сверка столбцев по численными дынными и перенос больших значений")
    print("Введите номер подходящей функции")
    choise = 1#int(input())

    choise_unit = {
        1: deptControl,
        #2: columnRevise
    }
    if choise in choise_unit:
        choise_unit[choise]()
    else:
        print('указанный вариант отсутствует')
        raise SystemExit(11)
main()
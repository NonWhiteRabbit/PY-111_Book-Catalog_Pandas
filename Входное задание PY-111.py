"""
Библиотека для хранения данных книг, поиска, добавления, удаления, а также редактирования
информации о книге в каталоге.

Библиотек реализована с использованием библиотеки Pandas.
Каталог книг хранится в файле "Каталог книг.xls". Все манипуляции с каталогом отражаются в этом файле.

Для вывода каталога на экран использован модуль PrettyTable по причине того, что при выводе
на печать в консоли отображаются не все столбцы каталога"""


import pandas as pd
import xlwt
from prettytable import PrettyTable


def main():
    """
    Функция, запрашивающая у пользователя информацию о том, что он хочет сделать с каталогом книг.

    :return: None
    """
    valid = False
    while not valid:
        print('\nЧто вы хотите сделать', '1 - осуществить поиск книги в каталоге', '2 - добавить книгу в каталог',
              '3 - редактировать информацию о существующей книге', '4 - удалить книгу из каталога',
              '5 - посмотреть каталог', '0 - выход', sep='\n')
        vvod = int(input())
        try:
            vvod = int(vvod)
        except ValueError:
            print("\nНекорректный ввод")
            continue
        if vvod == 1:
            input_for_search()
            continue
        elif vvod == 2:
            input_for_add()
            continue
        elif vvod == 3:
            input_for_replace()
            continue
        elif vvod == 4:
            input_for_delete()
            continue
        elif vvod == 5:
            catalog_print()
            continue
        elif vvod == 0:
            valid = True
        else:
            print("\nНекорректный ввод")
            continue


def input_for_search():
    """
    Функция для получения от пользователя данных для поиска в каталоге

    :return: Вывод на экран результатов поиска
    """
    with pd.ExcelFile("Каталог книг.xls") as xls:
        df = pd.read_excel(xls, usecols=['Название', 'Автор', 'Год выпуска', 'Жанр'])
    search_data = input('Что ищем (название, автор, год, жанр)?\nЕсли перечисление, то через запятую\n')
    search_data = search_data.split(',')
    search_data = list(map(lambda x: x.strip(), search_data))
    search_data = list(map(lambda x: x.lower(), search_data))
    for i in search_data:
        print(search(df, i))


def search(catalog, str_for_search=''):
    """
    Функция для поиска введенных пользователем данных в каталоге

    :param catalog: каталог книг
    :param str_for_search: данные, полученные от пользователя
    :return: Вывод на экран результатов поиска
    """
    if len(str_for_search) == 0:
        print('Вы ничего не ввели\n')
        input_for_search()
    else:
        pos = 0
        counter = 0
        while pos in catalog.index:
            for i in catalog.iloc[pos]:
                j = str(i)
                j = j.lower()
                if j.find(str_for_search) != -1:
                    counter += 1
                    print(f'\n{catalog.iloc[pos]}')
            pos += 1
        if counter == 0:
            print('\nНичего не найдено')
    return str("_"*48)


def input_for_add():
    """
    Функция для получения от пользователя данных для добавления книги в каталог

    Проверяет полноту ввода данных, а также наличие книги в каталоге
    :return: Вывод на экран результата добавления книги
    """
    with pd.ExcelFile("Каталог книг.xls") as xls:
        df = pd.read_excel(xls, usecols=['Название', 'Автор', 'Год выпуска', 'Жанр'])
    valid = False
    while not valid:
        new_data = input('Введите через запятую информацию для добавляемой книги (название, автор, год, жанр):\n')
        new_data = new_data.split(',')
        if len(new_data) != 4:
            print('Данные неполные\n')
            continue
        new_data = list(map(lambda x: x.strip(), new_data))
        new_data = list(map(lambda x: x.title(), new_data))
        n = new_data[0]
        a = new_data[1]
        y = new_data[2]
        g = new_data[3]
        counter = 0
        for i in df['Название']:
            if i.lower() == n.lower():
                counter += 1
        if counter != 0:
            print('Такая книга уже есть в каталоге\n')
            continue
        else:
            valid = True
        add(df, n, a, y, g)


def add(catalog, name=None, author=None, year=None, genre=None):
    """
    Функция для добавления книги в каталог и сохранения изменений в файл

    :param catalog: каталог книг, импортированый из файла
    :param name: название книги для добавления
    :param author: автор книги для добавления
    :param year: год выпуска книги для добавления
    :param genre: жанр книги для добавления
    :return: Вывод на экран результата добавления книги
    """
    df = pd.DataFrame({
        'Название': [name],
        'Автор': [author],
        'Год выпуска': [year],
        'Жанр': [genre]}, index=[len(catalog.index)])
    result = catalog.append(df)
    with pd.ExcelWriter("Каталог книг.xls") as writer:
        result.to_excel(writer)
    print('\nКнига добавлена')


def input_for_replace():
    """
    Функция для получения от пользователя названия книги для редактирования информации о ней в каталоге

    :return: None
    """
    with pd.ExcelFile("Каталог книг.xls") as xls:
        df = pd.read_excel(xls, usecols=['Название', 'Автор', 'Год выпуска', 'Жанр'])
    print(df['Название'])
    valid = False
    while not valid:
        rep_data = input('\nВведите название книги для редактирования в каталоге:\n')
        if len(rep_data) == 0:
            print('Вы ничего не ввели\n')
            continue
        else:
            valid = True
        rep_data = rep_data.strip()
        replace(df, rep_data)


def replace(catalog, rep_name):
    """
    Функция для редактирования информации о книге в каталоге и записи этой информации в файл с каталогом

    :param catalog: каталог книг, импортированый из файла
    :param rep_name: название книги для редактирования
    :return: None
    """
    counter = 0
    valid = False
    key = 0
    for k in catalog.index:
        for i in catalog.iloc[k]:
            j = str(i)
            j = j.lower()
            if j == rep_name.lower():
                key = k
                counter += 1
                print(f'\n{catalog.iloc[k]}\n')
    while not valid:
        rep_answer = int(input(
            "Что будем редактировать?\nНазвание-1/Автор-2/"
            "Год выпуска-3/Жанр-4/Выход в главное меню-0)?\n"))
        try:
            rep_answer = int(rep_answer)
        except ValueError:
            print("\nНекорректный ввод")
            continue
        if rep_answer == 1:
            rep_book = input("Введите новое название книги\n")
            rep_book = rep_book.strip()
            rep_book = rep_book.title()
            catalog.loc[key, 'Название'] = rep_book
        elif rep_answer == 2:
            rep_author = input("Введите нового автора книги\n")
            rep_author = rep_author.strip()
            rep_author = rep_author.title()
            catalog.loc[key, 'Автор'] = rep_author
        elif rep_answer == 3:
            rep_year = input("Введите новый год выпуска книги\n")
            rep_year = rep_year.strip()
            rep_year = rep_year.title()
            catalog.loc[key, 'Год выпуска'] = rep_year
        elif rep_answer == 4:
            rep_genre = input("Введите новый жанр книги\n")
            rep_genre = rep_genre.strip()
            rep_genre = rep_genre.strip()
            catalog.loc[key]['Жанр'] = rep_genre
        elif rep_answer == 0:
            valid = True
        else:
            print("\nНекорректный ввод")
            continue
    if counter == 0:
        print('\nТакая книга не существует')
    else:
        with pd.ExcelWriter("Каталог книг.xls") as writer:
            catalog.to_excel(writer)


def input_for_delete():
    """
    Функция для получения от пользователя данных для удаления книги из каталога

    :return: Вывод на экран результата удаления книги
    """
    with pd.ExcelFile("Каталог книг.xls") as xls:
        df = pd.read_excel(xls, usecols=['Название', 'Автор', 'Год выпуска', 'Жанр'])
    print(df['Название'])
    valid = False
    while not valid:
        del_data = input('Введите название книги для удаления из каталога:\n')
        if len(del_data) == 0:
            print('Вы ничего не ввели\n')
            continue
        else:
            valid = True
        del_data = del_data.strip()
        delete(df, del_data)


def delete(catalog, del_name):
    """
    Функция для удаления книги из каталога и сохранения изменений в файл

    :param catalog: каталог книг, импортированый из файла
    :param del_name: название книги для удаления
    :return: Вывод на экран результата удаления книги
    """
    counter = 0
    for k in catalog.index:
        for i in catalog.iloc[k]:
            j = str(i)
            j = j.lower()
            if j == del_name.lower():
                counter += 1
                print(f'\n{catalog.iloc[k]}\n')
                print('Удаляем?', '1 - да', '2 - нет', sep='\n')
                del_answer = int(input())
                try:
                    del_answer = int(del_answer)
                except ValueError:
                    print("\nНекорректный ввод")
                    continue
                if del_answer == 1:
                    catalog1 = catalog.drop([k])
                    catalog1.index = [i for i in range(len(catalog1.index))]
                    print('\nКнига удалена')
                    with pd.ExcelWriter("Каталог книг.xls") as writer:
                        catalog1.to_excel(writer)
                else:
                    print('Выход')
    if counter == 0:
        print('\nТакая книга не существует')


def catalog_print():
    with pd.ExcelFile("Каталог книг.xls") as xls:
        df = pd.read_excel(xls, usecols=['Название', 'Автор', 'Год выпуска', 'Жанр'])
    col = [i for i in df.columns]
    rows = []
    for i in df.index:
        for j in df.iloc[i]:
            rows.append(j)
    table = PrettyTable(col)
    while rows:
        table.add_row(rows[:len(col)])
        rows = rows[len(col):]
    print(table)


if __name__ == "__main__":
    main()

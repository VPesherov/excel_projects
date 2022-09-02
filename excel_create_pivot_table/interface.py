import os
import tkinter.messagebox as mb
from textwrap import wrap
from tkinter import *
from tkinter.filedialog import askopenfilename
from excel_script import work_with_excel

output_file_name = 'Сводная таблица результат.xlsx'

error_dict = {
    'КОД-1': ['Ошибка открытия файла КОД-1', 'Файл не найден или не был выбран.',
              'При нажатии на кнопку не был выбран итоговый файл.\nРешение: Выбрать исходный файл заново.'],
    'КОД-2': ['Ошибка открытия файла КОД-2', 'Файл не найден или был выбран не тот файл.',
              'При выполнении программы файл был не выбран или выбран, но не того типа с которым работает программа'
              '\nРешение: Выбрать исходный файл заноно.'],
    'КОД-3': ['Ошибка запущенного процесса КОД-3', f'Требуется закрыть файл вывода \'{output_file_name}\'.',
              f'При выполнении программы, файл в который записывается исходный результат был открыт.\
               \nРешение: Закрыть выходной файл {output_file_name}.']
}

help_dict = {
    '1': 'Для начала работы нажмите на кнопку \'Выбрать файл\' \n при успешном выборе'
         'файла высветится соответсвующая надпись.',
    '2': 'После успешного выбора файла, нажмите кнопку \'Начать работу\' \n и ожидайте'
         'завершения программы, если ошибки отсутсвуют программу выведет \n соответствующее'
         'окно и название выходного файла.'
}


def start_excel():
    msg = error_dict.get('КОД-2')[1]
    try:
        if filepath == "":
            return mb.showerror(error_dict.get('КОД-2')[0], msg)
        try:
            work_with_excel(filepath, output_file_name)
            msg = f"Успешно! Результат хранится в файле \'{output_file_name}\'"
            mb.showinfo("Результат выполнения", msg)
        except PermissionError:
            msg = error_dict.get('КОД-3')[1]
            return mb.showerror(error_dict.get('КОД-3')[0], msg)
    except NameError:
        return mb.showerror(error_dict.get('КОД-2')[0], msg)


def open_file():
    global filepath
    filepath = ""

    filepath = askopenfilename(
        filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
    )
    if filepath == "":
        msg = error_dict.get('КОД-1')[1]
        mb.showerror(error_dict.get('КОД-1')[0], msg)
    else:
        msg = f"Был выбран файл:\n\n\'{os.path.basename(filepath)}\'"
        mb.showinfo("Информация", msg)
        my_text = f'Выбран файл: {os.path.basename(filepath)}'
        lbl1.config(text=my_text)
        window.update()
        width = lbl1.winfo_width()
        if width > 290:
            char_width = width / len(my_text)
            wrapped_text = '\n'.join(wrap(my_text, int(290 / char_width)))
            lbl1['text'] = wrapped_text
            window.geometry('310x120')
        else:
            window.geometry('300x100')


def create_error_menu():
    window2 = Tk()
    window2.geometry('650x220')
    window2.title('Сводка об ошибках')

    lbl = Label(master=window2, text='Коды ошибок: ', justify=LEFT, font='20')
    lbl.place(x=10, y=10)
    my_y = 50
    for i, v in enumerate(error_dict):
        lbl = Label(master=window2, text=f'{v}: {error_dict.get(v)[1]}\n{error_dict.get(v)[2]}\n\n', justify=LEFT)
        lbl.place(x=10, y=my_y)
        my_y += 50


def create_help_menu():
    window1 = Tk()
    window1.geometry('490x180')
    window1.title('Окно помощь')
    lbl = Label(master=window1, text='Инструкция по использованию программы.', justify=LEFT, font='20')
    lbl.place(x=10, y=10)
    my_y = 50
    for i, v in enumerate(help_dict):
        lbl = Label(master=window1, text=f'{i + 1}. {help_dict.get(v)}', justify=LEFT)
        lbl.place(x=10, y=my_y)
        my_y += 50


def create_interface():
    global window, lbl1
    window = Tk()
    window.geometry('300x100')
    window.title('Результат по регионам')

    main_menu = Menu(window)
    window.config(menu=main_menu)

    file_menu = Menu(main_menu, tearoff=0)
    file_menu.add_command(label='Выбрать файл', command=open_file)

    main_menu.add_cascade(label='Файл', menu=file_menu)

    help_menu = Menu(main_menu, tearoff=0)
    help_menu.add_command(label="Помощь", command=create_help_menu)
    help_menu.add_command(label="Ошибки", command=create_error_menu)
    main_menu.add_cascade(label='Справка', menu=help_menu)

    main_menu.add_command(label='Выход', command=window.destroy)

    btn = Button(master=window, text="Выбрать файл", command=open_file)
    btn.grid(row=0, column=0, ipadx=95, padx=10, pady=5)
    btn1 = Button(master=window, text="Начать работу", command=start_excel)
    btn1.grid(row=1, column=0, ipadx=95, padx=10, pady=5)

    lbl1 = Label(master=window, text="")
    lbl1.grid(row=2, column=0, sticky=W, padx=10)

    window.mainloop()

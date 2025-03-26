'''
создание главного окна
'''
import tkinter as tk
from tkinter import ttk
from update_table import update_table_main

def get_file_name():
    '''
    get file name from entry
    '''
    result = update_table_main(entry1.get(), entry2.get(), entry3.get(), entry4.get())
    label_out.config(text = result)
if __name__ == '__main__':
# Создание основного окна
    root = tk.Tk()
    root.title("Обновить данные")
    root.geometry("600x600")  # Задаем размеры окна

    # Создание метки основного фаила
    label1 = tk.Label(root, text="Название основного фаила")
    label1.pack(pady=20)  # Добавляем метку в окно и задаем отступы
    entry1 = ttk.Entry(width=40)
    entry1.pack(padx=6, pady=6)
    # Создание метки листа основного фаила
    label2 = tk.Label(root, text="Название листа основного фаила")
    label2.pack(pady=20)  # Добавляем метку в окно и задаем отступы
    entry2 = ttk.Entry(width=40)
    entry2.pack(padx=6, pady=6)
    # Создание метки дополнительного фаила
    label3 = tk.Label(root, text="Название дополнительного фаила")
    label3.pack(pady=20)  # Добавляем метку в окно и задаем отступы
    entry3 = ttk.Entry(width=40)
    entry3.pack(padx=6, pady=6)
    # Создание метки листа дополнительного фаила
    label4 = tk.Label(root, text="Название листа дополнительного фаила")
    label4.pack(pady=20)  # Добавляем метку в окно и задаем отступы
    entry4 = ttk.Entry(width=40)
    entry4.pack(padx=6, pady=6)
    # Создание вывода
    label_out = tk.Label(root)
    label_out.pack(pady=20)
    # Создание кнопки
    button = tk.Button(root, text="Обновить", command=get_file_name)
    button.pack(pady=10)  # Добавляем кнопку в окно и задаем отступы

    # Запуск главного цикла приложения
    root.mainloop()

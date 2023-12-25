# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, filedialog
import os
import pandas as pd
import sys
from tkcalendar import DateEntry
import datetime as dt

try: script_path = os.path.realpath(os.path.dirname(__file__))
except: script_path = os.path.abspath(os.path.join(os.getcwd()))

sys.path.append(script_path)

import db_upload_module
import report_download_module
from db_upload_module import df_to_db
from report_download_module import save_excel

def update_status_text(status_field, message):
    status_field.config(state='normal')
    status_field.insert(0.0, f'{message}\n')
    status_field.config(state='disabled')

def choose_file():
    path_entry.delete(0, tk.END)
    file = filedialog.askopenfile(
        mode='r', 
        filetypes=[('Excel binary', '*.xlsb'), ('Все файлы', '*.*')])
    if file:
       filepath = os.path.abspath(file.name)
       path_entry.delete(0, tk.END)
       path_entry.insert(0, str(filepath))
       update_status_text(status_field, f'Выбран файл {filepath}')
       import_button['state'] = 'normal'
       
def choose_folder():
    report_path_entry.delete(0, tk.END)
    folder = filedialog.askdirectory()
    if folder:
       folderpath = os.path.abspath(folder)
       report_path_entry.delete(0, tk.END)
       report_path_entry.insert(0, str(folderpath))
       update_status_text(status_field, f'Выбрана папка {folderpath}')
       report_execute_button['state'] = 'normal'

import_df = pd.DataFrame()
def import_excel():
    global import_df
    filepath = path_entry.get()
    import_df = pd.read_excel(filepath)
    label_text = len(import_df)
    update_status_text(status_field, f'Прочтено линий: {label_text}')
    upload_button['state']='normal'
     
    
def upload_df_to_db():
    df_to_db(import_df)
    update_status_text(status_field, f'Перенос в БД завершен')
    upload_button['state']='disabled'

def report_download():
    global eff_from, eff_to
    eff_from_string = dt.datetime.strptime(eff_from.get(), '%d-%m-%Y').strftime('%Y-%m-%d')
    eff_to_string = dt.datetime.strptime(eff_to.get(), '%d-%m-%Y').strftime('%Y-%m-%d')
    folderpath = report_path_entry.get()
    save_excel(eff_from_string, eff_to_string, folderpath)
    file_date = dt.datetime.now().strftime('_%Y%m%d')
    update_status_text(status_field, f'Отчет report_{eff_from_string}_{eff_to_string}{file_date}.xlsx сформирован в {folderpath}')
    
def db_script_reset():
    import db_scripts_create
    update_status_text(status_field, f'Хранимые процедуры в БД обновлены/созданы')
    
win = tk.Tk()
# h = 500
# v = 300
# h_offset = 500
# v_offset = 50 
win.title('Проект 2312')
# win.geometry(f"{h}x{v}+{h_offset}+{v_offset}")
win.resizable(False, False)


# разметка фреймов

upload_frame = tk.Frame(win, borderwidth=2, bg=None, highlightbackground="black", highlightthickness=1)
report_frame = tk.Frame(win, borderwidth=1, bg=None, highlightbackground="black", highlightthickness=1)
status_frame = tk.Frame(win, borderwidth=1, bg=None, highlightbackground="black", highlightthickness=1)

upload_frame.grid(row=0, column=0, sticky='wens', pady=5, padx=5)
report_frame.grid(row=1, column=0, sticky='wens', pady=5, padx=5)
status_frame.grid(row=2, column=0, sticky='wens', pady=5, padx=5)
''

upload_frame_label = tk.Label(upload_frame, text='Загрузка данных из файла в БД', bg=None)
upload_frame_label.grid(row=0,
                        column=0, 
                        columnspan=3, 
                        padx=5, 
                        pady=5, 
                        sticky='we')

# 

browse_button_text = 'Выбрать файл'
browse_button = tk.Button(
    upload_frame, 
    text=browse_button_text,
    command=choose_file
    )

browse_button.grid(row=1,column=0, padx=5, pady=5)

path_entry = tk.Entry(upload_frame)
path_entry.grid(row=1, 
                column=1, 
                columnspan=2, 
                padx=5, 
                pady=5, 
                sticky='we' )


import_button_text = 'Чтение данных'
import_button = tk.Button(
    upload_frame, 
    text=import_button_text,
    command=import_excel,
    state='disabled'
    )
import_button.grid(row=2,column=0, padx=5, pady=5)

upload_button_text = 'Загрузить в БД'
upload_button = tk.Button(
    upload_frame,
    text=upload_button_text,
    command=upload_df_to_db,
    state='disabled'
    )
upload_button.grid(row=2,column=1, padx=5, pady=5)

upload_frame.columnconfigure(2, weight=1)


# секция выгрузки отчета

report_frame_label = tk.Label(report_frame, text='Выгрузка файла отчета за период из БД', bg=None)
report_frame_label.grid(row=0,column=0, columnspan=4, padx=5, pady=5, sticky='we')

eff_from, eff_to = tk.StringVar(), tk.StringVar()

dt1, dt2 = '',''
def my_upd1(*args): # triggered when value of string varaible changes
    global dt1,dt2
    if(len(eff_from.get())>3):
        dt1 = cal1.get_date()
        cal2.config(mindate = dt1)
        dt1 = dt1.strftime("%Y-%m-%d") # create date type
        query.set("Выбран диапазон между " + dt1 + " и " + dt2)
def my_upd2(*args):
    if(len(eff_to.get())>3):
        dt2 = cal2.get_date()
        cal1.config(maxdate = dt2)
        dt2 = dt2.strftime("%Y-%m-%d")
        query.set("Выбран диапазон между " + dt1 + " и " + dt2)
def my_reset():
    dt1, dt2= '',''
    dt_today = dt.date.today().strftime('%d-%m-%Y')
    cal1.set_date(dt_today) # todays date 
    cal2.set_date(dt_today) # todays date 
    cal1.config(maxdate=None)
    cal2.config(mindate=None)    
    query.set('Диапазон не выбран')

selection_label = tk.Label(report_frame, text=f'Выбор\nдиапазона')
selection_label.grid(row=1, column=0, rowspan=2)

from_label = tk.Label(report_frame, text='Начало периода')
from_label.grid(row=1, column=1)
to_label = tk.Label(report_frame, text='Конец периода')
to_label.grid(row=1, column=2)

cal1=DateEntry(report_frame, selectmode='day', textvariable = eff_from, date_pattern='dd-MM-yyyy')
cal1.grid(row=2,column=1, padx=5, pady=5)
cal2=DateEntry(report_frame,selectmode='day', textvariable = eff_to, date_pattern='dd-MM-yyyy')
cal2.grid(row=2,column=2, padx=5, pady=5)
b1=tk.Button(report_frame,text='Сброс'
    ,command=lambda:my_reset())
b1.grid(row=2,column=3, padx=5)

query=tk.StringVar(value = 'Диапазон не выбран')
l1=tk.Label(report_frame, textvariable = query, bg= None)
l1.grid(row=3, column=0, columnspan=4, pady=5)
eff_from.trace('w',my_upd1) # on change of string variable 
eff_to.trace('w',my_upd2) # on change of string variable 

report_path_button_text = 'Выбрать папку'
report_path_button = tk.Button(
    report_frame, 
    text=report_path_button_text,
    command=choose_folder
    )

report_path_button.grid(row=4,column=0, padx=5, pady=5, sticky='w')

report_path_entry = tk.Entry(report_frame)
report_path_entry.grid(row=4, column=1, columnspan=3, padx=5, pady=5, sticky='we' )

report_execute_button_text = 'Сформировать отчет'
report_execute_button = tk.Button(
    report_frame, 
    text=report_execute_button_text,
    command=report_download,
    state='disabled'
    )
report_execute_button.grid(row=5, column=0, columnspan=2, padx=2, pady=5, sticky='w')

report_frame.columnconfigure(0, weight=2)
report_frame.columnconfigure(1, weight=3)
report_frame.columnconfigure(2, weight=3)

# секция поля статуса

temp_s_frame_name = tk.Label(status_frame, text='Статус')

temp_s_frame_name.pack(expand=True, padx=3, pady=5)

fontObj = tk.font.Font(size=8)

status_field = tk.Text(status_frame, 
                       width = 60, 
                       height=4, 
                       wrap='word', 
                       font=fontObj,
                       # state='disabled'
                       )
status_field.pack(pady=5)

# секция меню

main_menu = tk.Menu(win)
win.config(menu=main_menu)

file_menu = tk.Menu(main_menu, tearoff=0)
file_menu.add_command(label='Выход', command=lambda: win.destroy())
control_menu = tk.Menu(main_menu, tearoff=0)
control_menu.add_command(label='Записать/обновить скрипты БД', command=db_script_reset)

main_menu.add_cascade(label='Файл', menu=file_menu)
main_menu.add_cascade(label='Управление', menu=control_menu)

win.mainloop()

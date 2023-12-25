# -*- coding: utf-8 -*-

import datetime as dt
import os
import yaml
import pyodbc
import re
import pandas as pd

#Обращаемся на сервер БД за выгрузкой и отчетом за период
try: script_path = os.path.realpath(os.path.dirname(__file__))
except: script_path = os.path.abspath(os.path.join(os.getcwd()))

with open(script_path+'\\config.yml','r') as f:
    config_data = yaml.full_load(f)


connection_parameters = config_data['sql_server']['connection_parameters']
connection_string = f'''
    Driver={connection_parameters['db_driver']};
    Server={connection_parameters['db_server']};
    Database={connection_parameters['db_name']};
    Trusted_Connection=yes; 
    '''

cnxn = pyodbc.connect(connection_string)
cursor = cnxn.cursor()

def db_query_exec(query):
    try:
        connection = pyodbc.connect(connection_string)
        first_word = next(m.group() for m in re.finditer(r'\w+', query)) # первое слово в запросе
        print(f'Executing {first_word} query')
        connection.autocommit = True
        cursor = connection.cursor()
        cursor.execute(query)

    except (Exception, pyodbc.Error) as error:
        print("Error while accesing", error)

    finally:
        #закрываем БД
        if connection:
            cursor.close()
            connection.close()
            print("DB connection closed")

default_schema_name = config_data['sql_server']['default_parameters']['schema']
default_table_name =  config_data['sql_server']['default_parameters']['import_data_table']
default_raw_procedure =  config_data['sql_server']['default_parameters']['raw_data_procedure']
default_report_procedure =  config_data['sql_server']['default_parameters']['report_data_procedure']

def report_db_request(eff_from, eff_to, **kwargs):
    target_schema = kwargs.get('schema', default_schema_name)
    raw_procedure = kwargs.get('raw_procedure', default_raw_procedure)
    report_procedure = kwargs.get('report_procedure', default_report_procedure)
    query = f''' 
    exec {target_schema}.{raw_procedure} '{eff_from}', '{eff_to}'
    '''
    out_df_raw = pd.read_sql_query(query, con = cnxn, parse_dates='dt')
    query = f''' 
    exec {target_schema}.{report_procedure} '{eff_from}', '{eff_to}'
    '''
    out_df_report = pd.read_sql_query(query, con = cnxn, parse_dates='dt') 
    return out_df_raw, out_df_report

# Словари форматов
sheet_title_format_dict = {
    'bold': True,
    'font_size': 12,
    'text_wrap': False
}
date_format_dict = {
    'num_format':'dd.mm.yyyy'
}
percent_format_dict = {
    'num_format':'0.00%'
}
rounded_format_dict =  {
    'num_format':'#,##0.00'
}

def save_excel (eff_from, eff_to, folderpath, **kwargs):
    out_df_raw, out_df_report = report_db_request(eff_from, eff_to, **kwargs)
    file_date = dt.datetime.now().strftime('_%Y%m%d')
    out_file_name = f'{"report_"+eff_from+"_"+eff_to+file_date}.xlsx'
    out_path = folderpath + '\\'
    writer = pd.ExcelWriter(f'{out_path+out_file_name}', engine='xlsxwriter')
    workbook = writer.book
    sheet_title_format =  workbook.add_format(sheet_title_format_dict) # Формат заголовка на листе
    date_format = workbook.add_format(date_format_dict) # Формат даты
    cell_format = workbook.add_format() # общий формат
    percent_format = workbook.add_format(percent_format_dict) # процентный формат
    rounded_format = workbook.add_format(rounded_format_dict) # округленный формат
    column_data_dict = {
        0: {'source_col': 'dt',
            'header': 'Дата',
            'format': date_format,
            'width': 13
              },
        1: {'source_col': 'article',
            'header': 'Артикул',
            'width': 18
           },
        2: {'source_col': 'kg',
            'header': 'Количество',
            'width': 12
        }
    }
    worksheet_name = 'raw_data'
    worksheet = workbook.add_worksheet(worksheet_name)
    sheet_title_text = f'Выгрузка за период {eff_from} - {eff_to}'
    active_row = 0
    active_col = 0
    # создаем функцию записи датафрейма на лист по словарю выходной таблицы
    def write_df_table(
        active_row,
        active_col,
        column_data_dict,
        out_df
    ):
        worksheet.merge_range(
            active_row, 
            active_col, 
            active_row, 
            active_col + len(column_data_dict) - 1, 
            sheet_title_text,
            sheet_title_format
        )
        active_row += 1
        for col in column_data_dict:
            data = [column_data_dict[col]['header']] + out_df[column_data_dict[col]['source_col']].fillna('').to_list()
            col_cell_format = column_data_dict[col].get('format', cell_format)
            worksheet.write_column(
                active_row,
                active_col + col,
                data,
                col_cell_format
            )
            worksheet.set_column(active_col + col, active_col + col, column_data_dict[col].get('width', 10))
        worksheet.freeze_panes(active_row + 1, 0)

    # выполняем функцию для таблицы сырых данных
    write_df_table(
        active_row,
        active_col,
        column_data_dict, 
        out_df_raw)
        
    # запись листа с отчетными данными

    column_data_dict = {
        0: {
            'source_col': 'r_year',
            'header': 'Год',
            'width': 10
              },
        1: {
            'source_col': 'r_month',
            'header': 'Месяц',
            'width': 10
           },
        2: {
            'source_col': 'article',
            'header': 'Артикул',
            'width': 12
           },
        3: {
            'source_col': 'annual_avg',
            'header': 'Сред. за год',
            'format': rounded_format,
            'width': 12
           },
        4: {
            'source_col': 'month_avg',
            'header': 'Сред. за месяц',
            'format': rounded_format,
            'width': 12
        },
        5: {
            'source_col': 'article_share',
            'header': 'Доля за период',
            'format': percent_format,
            'width': 12
        }
    }

    worksheet_name = 'report'
    worksheet = workbook.add_worksheet(worksheet_name)
    sheet_title_text = f'Отчет за период {eff_from} - {eff_to}'
    active_row = 0
    active_col = 0

    # выполняем функцию для таблицы отчетных данных
    write_df_table(
        active_row,
        active_col,
        column_data_dict, 
        out_df_report)

    # завершаем запись файла
    workbook.close()


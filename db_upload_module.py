
# -*- coding: utf-8 -*-
import pandas as pd
import os
import yaml
import datetime as dt
import pyodbc
import re

if __name__ == '__main__':
    print('internal call')
else: print('external call')

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

def df_preprocessing(df):
    # подготовка датафрейма к переносу в БД
    df['dt'] = pd.to_datetime(df['dt'], unit='D', origin='1899-12-30')
    df['article'] = df['article'].astype('str')
    return df
    

def df_to_db(df, **kwargs):
    target_schema = kwargs.get('schema', default_schema_name)
    target_table = kwargs.get('table', default_table_name)
    print(target_schema, target_table)
    df = df_preprocessing(df)
    article_max_length = int(df['article'].str.len().max())
    query = f'''
        drop table if exists {target_schema}.{target_table};
	    create table {target_schema}.{target_table} (
            record_id int IDENTITY PRIMARY KEY,
		    dt date,
		    article varchar({article_max_length}),
		    kg numeric
		    )
    '''
    print(f'Создание таблицы {target_table}')
    db_query_exec(query)
    print(f'Внесение данных ({len(df)} строк) в таблицу {target_table}')
    cursor = cnxn.cursor()
    for index, row in df.iterrows():
          cursor.execute(
              f'INSERT INTO {target_schema}.{target_table} (dt, article, kg) values (?,?,?)', 
                        row['dt'],
                        row['article'],
                        row['kg']
                        )
    cnxn.commit()
    cursor.close()
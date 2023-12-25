

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

#функция выполнения запроса в БД
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

#создаем процедуру выгрузки данных в заданном периоде
query = f'''
create or alter procedure {default_schema_name}.{default_raw_procedure}
    @eff_from date = '1900-01-01',   
    @eff_to date = '9999-12-31' 
AS   
	SELECT 
		dt,
		article,
		kg
	FROM {default_schema_name}.{default_table_name} rd 
	WHERE dt BETWEEN  @eff_from AND @eff_to 
'''
print('Coздание процедуры отчета')
db_query_exec(query)

#создаем процедуру получения отчетных данных в заданном периоде
query = f'''
create or alter procedure {default_schema_name}.{default_report_procedure}
    @eff_from date = '1900-01-01',   
    @eff_to date = '9999-12-31' 
    AS
	with total_data AS (
		SELECT distinct year(dt) AS r_year,
		month(dt) AS r_month,
		rd.article,
		round(avg(kg) over (partition by year(dt), month(dt), article), 2) AS month_avg,
		round(avg(kg) over (partition by year(dt), article),2) AS annual_avg
		FROM {default_schema_name}.{default_table_name} rd 
		)
	SELECT DISTINCT  
	year(dt) AS r_year,
	month(dt) AS r_month,
	rd.article,
	td.annual_avg,
	td.month_avg,
	round((sum(kg) over (partition by rd.article))/(sum(kg) over ()),3) AS article_share
	FROM {default_schema_name}.{default_table_name} rd
	left join total_data td on month(dt) = r_month AND year(dt)= r_year AND td.article=rd.article
	where dt between @eff_from and @eff_to
    order by rd.article, year(dt), month(dt)
'''
print('Coздание процедуры отчета')
db_query_exec(query)
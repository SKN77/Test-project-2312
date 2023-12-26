# Test project 2312

Tестовое задание MS SQL и Python 

Test_project_2312.py основной скрипт с пользовательским интерфейсом

config.yml файл с конфигурацией. По умолчанию прописано название заранее созданной БД test_task

В меню Управления интерфейсного окна пункт 'Записать/обновить скрипты БД' выолняет модуль db_scripts_create.py. Модуль создает две хранимые процедуры в БД - выгрузки сырых данных в заданном диапазоне и выгрузки отчетных данных.

Остальной функционал реализован в основном поле интерфейсного окна.

Для работы используются дополнительные библиотеки

* pandas
* pyodbc
* pyxlsb
* pyyaml
* tkcalendar
* xlsxwriter

## Модули
* db_scripts_create.py - модуль создания/перезаписи существующих хранимых процедур в БД (названия в config.yml)
* 

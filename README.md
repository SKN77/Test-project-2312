# Test project 2312

Tестовое задание MS SQL и Python 

```Test_project_2312.py``` основной скрипт с пользовательским интерфейсом

По умолчанию прописано название заранее созданной БД ```test_task``` (можно изменить на название другой уже созданной БД в ```config.yml```)

В меню Управления интерфейсного окна пункт 'Записать/обновить скрипты БД' выолняет модуль ```db_scripts_create.py```. Модуль создает две хранимые процедуры в БД - выгрузки сырых данных в заданном диапазоне и выгрузки отчетных данных.

Остальной функционал реализован в основном поле интерфейсного окна.

Для работы используются дополнительные библиотеки

* pandas
* pyodbc
* pyxlsb
* pyyaml
* tkcalendar
* xlsxwriter

## Модули
* ```Test_project_2312.py``` - основной скрипт с формой пользовательского интерфейса
* ```db_scripts_create.py``` - модуль создания/перезаписи существующих хранимых процедур в БД (названия в config.yml)
* ```db_upload_module.py``` - модуль передачи чтения Excel binary файла условленного формата и загрузки прочитанных данных в БД
* ```report_download_module.py``` - модуль экспорта "сырых" и отчетных данных из БД на заданный период в Excel файл в установленной форме
* ```config.yml``` файл с конфигурацией. 

# ECS2XLS конвертор XML в XLSX для файлов экспорта трендов FLS ECS8

## Назначение
Преобразование xml файла экспорта трендов SCADA FLS ECS8 в EXCEL.
![Demo](resources/xml2xls.gif)

## Использование
1. Экспортировать данные трендов в xml:
  ```ECS8 -> Trend -> Trend Export -> Logged values```
  По умолчанию имя файла ~ "Trend on 2024-08-30T21.23.17.xml"

2. В консоли windows выполнить:
  ```.\xml2xls.exe '.\Trend on 2024-08-30T21.23.17.xml'```
  Будет создан файл ~ "Trend on 2024-08-30T21.23.17.xml.xlsx"


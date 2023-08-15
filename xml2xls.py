import logging
import os
import time
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
from pathlib import Path

# import numpy as np
import pandas as pd
from alive_progress import alive_bar, config_handler
import argparse

from colorama import init, Fore
from colorama import Back
from colorama import Style

init(autoreset=True)

DEBUG = False
DATA = []
QTY_TAGS = 0  # количество тегов
QTY_VALS = 0  # количество значений
_PRG_DIR = Path("./").absolute()
_TRND_DIR = _PRG_DIR  # _PRG_DIR / "trends/"
_LOG_FILE = _PRG_DIR / "xml2xls.log"
MAX_ROWS_EXL = 1000000  # макс. кол-во строк в Excel 1048000
_MAGADAN_UTC = 11  # Магаданское время +11 часов к UTC
log_format = f"%(asctime)s - %(levelname)s -(%(funcName)s(%(lineno)d) - %(message)s"
logging.basicConfig(
    format=log_format,
    level=logging.DEBUG if DEBUG else logging.INFO,
    # filename=_LOG_FILE,
)
logger = logging.getLogger()


def get_trends_files_name(sortby="time", rev=True) -> list:
    """Получение списка XML файлов в папке Trends"""
    sort = {"time": os.path.getmtime, "size": os.path.getsize, "name": None}
    files = []
    for file in sorted(_TRND_DIR.glob("*.xml"), reverse=rev, key=sort[sortby]):
        files.append(file.name)
    return files


def open_xml(xml_file_name: str):
    """"""
    tree = ET.parse(_TRND_DIR / xml_file_name)
    root = tree.getroot()
    return root


def get_trends_info(xml_root) -> dict:
    """Извлечение из xml информации по трендам"""
    trends_info = {}
    date_format = "%Y-%m-%dT%H:%M:%S"
    # период выборки
    trends_info["start_interval"] = datetime.strptime(
        xml_root[1].text[:19], date_format
    )
    trends_info["end_interval"] = datetime.strptime(xml_root[0].text[:19], date_format)
    # указатель на теги в xml
    trends_info["xml_tags"] = xml_root[3]
    # количество тегов
    trends_info["qty_tags"] = len(xml_root[3])
    return trends_info


def get_tag_meta(tag):
    pass


def make_list_for_df(enteries) -> list:
    """"""
    data = []

    with alive_bar(len(enteries), force_tty=True, length=30) as bar:
        for enterie in enteries:
            str_time = enterie[1].text[:25].replace("Z", "")
            time_v = datetime.strptime(str_time, "%Y-%m-%dT%H:%M:%S.%f") + timedelta(
                hours=_MAGADAN_UTC
            )  # enterie.find('{Fls.Core.Value.CI.Export.M1}Time').text
            value = enterie[2].text
            # добавление строк в Pandas dataframe
            # df_tag.loc[len(df_tag)] = [time_v, value]
            data.append((time_v, value))
            bar()
    return data


def save_as_xlsx(dataframes: list, filename):
    """Сохраняет переданные datafaremes в xlsx"""
    writer = pd.ExcelWriter(_TRND_DIR / filename, engine="xlsxwriter")
    workbook = writer.book
    # перебор тегов
    with alive_bar(len(dataframes), force_tty=True, length=3) as bar:
        for tag in dataframes:
            df = tag[2]  #
            tag_name = tag[0]
            # колл-во строк в DF
            count_row = df.shape[0]
            # колличество вкладок excel
            qty_sheets = count_row // MAX_ROWS_EXL + 1 if (count_row > MAX_ROWS_EXL) else 1
            logger.debug(f"Тег: {tag[0]}; count row: {count_row}; pages:{qty_sheets} ")
            for i in range(qty_sheets):
                logger.debug(f"Перебор листов i {i}")
                sheet_name = f"{tag_name} {i + 1}"
                first_row = i * MAX_ROWS_EXL
                end_row = (
                    ((i + 1) * MAX_ROWS_EXL)
                    if count_row >= ((i + 1) * MAX_ROWS_EXL)
                    else count_row
                )
                logger.debug(f"sheet_name: {sheet_name}; first_row: {first_row}; end row: {end_row}")
                df_page = df.iloc[first_row:end_row]
                print(
                    f"Сохраняется лист: {Fore.GREEN}{sheet_name}{Fore.WHITE}; "
                    f"колл записей:{Fore.GREEN} {df_page.shape[0]}{Fore.WHITE};\n  "
                    f"дата от {df_page.iat[0, 0]} до {df_page.iat[df_page.shape[0] - 1, 0]} \n")

                df_page.to_excel(writer, sheet_name=sheet_name, startrow=2, startcol=0, index=False)

                worksheet = writer.sheets[sheet_name]
                worksheet.write(0, 0, f"Значения тега {tag_name} за период c {df_page.iat[0, 0]} по "
                                      f"{df_page.iat[df_page.shape[0] - 1, 0]}")
            bar()
    writer.close()


def convert_xml2xls(xml_file_name):
    """"""
    start_time = time.time()
    data = []  #
    data2 = []

    file = _TRND_DIR / xml_file_name
    file_size = file.stat().st_size / 1048576

    print(f"{Fore.YELLOW}Открывается файл: {Fore.GREEN + xml_file_name + Style.RESET_ALL}; "
          f"Размер: {Fore.GREEN + str(round(file_size)) + Style.RESET_ALL} Мб")
    with alive_bar(1, force_tty=True, length=3) as bar:
        # открытие xml файла
        xml_root = open_xml(xml_file_name)
        bar()
    # Проверка, что файл export трендов
    tmp = xml_root.tag
    if xml_root.tag != '{Fls.Core.Value.CI.Export.M1}ValueHistorianExport':
        err_msg = f"{Fore.RED} Это не экспорт трендов ECS"
        raise Exception(err_msg)
    # получение информации о трендах
    trends_info = get_trends_info(xml_root)
    print(f"Интервал: {Fore.GREEN}{trends_info['start_interval']} - {trends_info['end_interval']}")
    print(f"Количество тегов: {Fore.GREEN}{trends_info['qty_tags']}")
    print(f"\n{Fore.YELLOW}Формируется датафрейм")
    s_time = time.time()
    qnty_tags = 0
    qnty_vals = 0
    qnty_vals_l = 0
    # Перебор тегов
    for tag in trends_info["xml_tags"]:  # xml_root[3]:
        qnty_tags += 1
        # if qnty_tags > 2: break
        # имя тега
        designation = tag.find("{Fls.Core.Value.CI.Export.M1}Designation").text
        unit = tag.find("{Fls.Core.Value.CI.Export.M1}Unit").text
        value_entries = []
        # date_format = '%Y-%m-%dT%H:%M:%S.%f'
        df_tag = pd.DataFrame(columns=[f"{designation}_dt", f"{designation}_value"])
        enteries = tag.find("{Fls.Core.Value.CI.Export.M1}valueEntries")
        print(f"Обрабатывается тег:{Fore.GREEN}{designation}")
        qnty_vals_l = 0
        data2 = make_list_for_df(enteries)
        qnty_vals += len(data2)
        df_tag = pd.DataFrame(data2, columns=["dt", "value"])
        data2.clear()
        # df_tag = df_tag.iloc[:1048576] # макс. колл-во строк в EXEL <= 1048576
        data.append([designation, unit, df_tag])

    print(f"{Fore.YELLOW}Парсинг закончен")
    print(f"За время: {Fore.GREEN}{time.time() - start_time}")
    print(f"Обработано: Тегов {Fore.GREEN}{qnty_tags}{Fore.WHITE}; Значений: {Fore.GREEN}{qnty_vals}{Fore.WHITE};"
          f" Тег/сек: {Fore.GREEN}{qnty_vals / (time.time() - start_time)}")
    print(f"\n{Fore.YELLOW}Начало сохранения в XLSX {Fore.GREEN}{xml_file_name}.xlsx")
    save_as_xlsx(data, xml_file_name + ".xlsx")
    print(f"\n{Fore.YELLOW}Exсel сохранён: {Fore.MAGENTA}{xml_file_name}.xlsx")
    print(f"Весь процесс занял времени:{Fore.GREEN}{round(time.time() - start_time)}сек")


def check_xml_file(file_name: str) -> bool:
    """Проверяет, что файл существует и XML"""
    file_xml = Path(file_name)
    if not file_xml.suffix.lower() == '.xml':
        err_msg = f"Нет расширения XML {Fore.RED}{file_xml}"
        raise Exception(err_msg)

    if not file_xml.is_file():
        err_msg = f"Файл не найден {Fore.RED}{file_xml}"
        raise Exception(err_msg)



def main():
    global MAX_ROWS_EXL
    parser = argparse.ArgumentParser(
        prog='ECS2XLS',
        description='ECS2XLS конвертор экспорта трендов ECS8 XML в XLSX',
        epilog='2023 7Art'
    )
    try:
        parser.add_argument('file', type=str, help='Имя файла экспорта трендов xml')
        parser.add_argument('-max_row', type=int, default=1000000, help='Макс количество строк на листе Excel (Def:1M)')
        args = parser.parse_args()
        MAX_ROWS_EXL = args.max_row
        file_xml = Path(args.file)
        check_xml_file(args.file)
        convert_xml2xls(args.file)

    except Exception as e:
        print(e)
        logger.info(e)
    finally:
        print("Press Enter to continue ...")
        input()




if __name__ == '__main__':
    main()

import logging
import os
import time
import xml.etree.ElementTree as ET
import pandas as pd
import argparse
import re
from datetime import datetime, timedelta
from pathlib import Path
from alive_progress import alive_bar, config_handler
from colorama import init, Fore
from colorama import Style

_VERSION = 0.9
_PRG_DIR = Path("./").absolute()
_TRND_DIR = _PRG_DIR  # _PRG_DIR / "trends/"
_LOG_FILE = _PRG_DIR / "xml2xls.log"
_MAGADAN_UTC = 11  # Магаданское время +11 часов к UTC
DEBUG = False
DATA = []
QTY_TAGS = 0  # количество тегов
QTY_VALS = 0  # количество значений
MAX_ROWS_EXL = 1000000  # макс. кол-во строк в Excel 1048000

init(autoreset=True)
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


def open_xml(xml_file_name: str) -> ET.parse:
    """Открытие xml файла"""
    with open(_TRND_DIR / xml_file_name, 'rb') as xml_file:
        tree = ET.parse(xml_file)
    root = tree.getroot()
    return root


def get_trends_info(xml_root: ET.parse) -> dict:
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


def make_list_for_df(enteries) -> list:
    """"""
    data = []
    pattern = r'(\d{4}-\d{2}-\d+T\d+:\d+:\d{2})(?!\.\d)'
    repl = r'\1.0'
    with alive_bar(len(enteries), force_tty=True, length=30) as bar:
        for enterie in enteries:
            str_time = enterie[1].text[:25].replace("Z", "")
            str_time = re.sub(pattern, repl, str_time)
            time_v = datetime.strptime(str_time, "%Y-%m-%dT%H:%M:%S.%f") + timedelta(
                hours=_MAGADAN_UTC
            )  # enterie.find('{Fls.Core.Value.CI.Export.M1}Time').text
            value = enterie[2].text
            # добавление строк в Pandas dataframe
            # df_tag.loc[len(df_tag)] = [time_v, value]
            data.append((time_v, value))
            bar()
    return data


def save_as_xlsx(dataframes: list, filename: str):
    """Сохраняет переданные dataframes в xlsx"""
    with alive_bar(len(dataframes) + 1, force_tty=True, length=30) as bar:
        with pd.ExcelWriter(_TRND_DIR / filename, engine="xlsxwriter") as writer:
            # перебор тегов
            for tag in dataframes:
                df = tag[2]  #
                tag_name = tag[0]
                # колл-во строк в DF
                count_row = df.shape[0]
                # количество вкладок excel
                qty_sheets = count_row // MAX_ROWS_EXL + 1 if (count_row > MAX_ROWS_EXL) else 1
                logger.debug(f"Tag: {tag[0]}; count row: {count_row}; pages:{qty_sheets} ")
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
                        f"Creating sheet: {Fore.GREEN}{sheet_name}{Fore.WHITE}; "
                        f"number of records:{Fore.GREEN} {df_page.shape[0]}{Fore.WHITE};\n  "
                        f"Data from {df_page.iat[0, 0]} to {df_page.iat[df_page.shape[0] - 1, 0]} \n")

                    df_page.to_excel(writer, sheet_name=sheet_name, startrow=2, startcol=0, index=False)

                    worksheet = writer.sheets[sheet_name]
                    worksheet.write(0, 0, f"Tag values {tag_name} from date {df_page.iat[0, 0]} to "
                                          f"{df_page.iat[df_page.shape[0] - 1, 0]}")
                bar()

            print(f"File saving...")
            bar()


def convert_xml2xls(xml_file_name: str):
    """Преобразование XML файла экспорта данных трена ECS в XLSX"""
    start_time = time.time()
    data = []  #

    file = _TRND_DIR / xml_file_name
    file_size = file.stat().st_size / 1048576
    xls_file_name = xml_file_name + ".xlsx"

    print(f"{Fore.YELLOW}File opens: {Fore.MAGENTA + xml_file_name + Style.RESET_ALL}; "
          f"Size: {Fore.GREEN + str(round(file_size)) + Style.RESET_ALL} Mb")
    with alive_bar(1, force_tty=True, length=3) as bar:
        # открытие xml файла
        xml_root = open_xml(xml_file_name)
        bar()
    # Проверка, что файл export трендов
    if xml_root.tag != '{Fls.Core.Value.CI.Export.M1}ValueHistorianExport':
        err_msg = f"{Fore.RED} This file is not an export of ECS8 trends"
        raise Exception(err_msg)
    # получение информации о трендах
    trends_info = get_trends_info(xml_root)
    print(f"Interval: {Fore.GREEN}{trends_info['start_interval']} - {trends_info['end_interval']}")
    print(f"Number of tags: {Fore.GREEN}{trends_info['qty_tags']}")
    print(f"\n{Fore.YELLOW}Dataframe is being ")
    qnty_tags = 0
    qnty_vals = 0
    # Перебор тегов
    for tag in trends_info["xml_tags"]:  # xml_root[3]:
        qnty_tags += 1
        # if qnty_tags > 2: break
        # имя тега
        designation = tag.find("{Fls.Core.Value.CI.Export.M1}Designation").text
        unit = tag.find("{Fls.Core.Value.CI.Export.M1}Unit").text
        # date_format = '%Y-%m-%dT%H:%M:%S.%f'
        enteries = tag.find("{Fls.Core.Value.CI.Export.M1}valueEntries")
        print(f"Tag processed:{Fore.GREEN}{designation}")
        data2 = make_list_for_df(enteries)
        qnty_vals += len(data2)
        df_tag = pd.DataFrame(data2, columns=["dt", "value"])
        data2.clear()
        # df_tag = df_tag.iloc[:1048576] # макс. колл-во строк в EXEL <= 1048576
        data.append([designation, unit, df_tag])

    del xml_root

    print(f"{Fore.YELLOW}Parsing completed")
    print(f"At time: {Fore.GREEN}{time.time() - start_time}")

    print(f"\n{Fore.YELLOW}Start saving: {Fore.MAGENTA}{xls_file_name}")
    # Сохранение в Excell
    save_as_xlsx(data, xls_file_name)

    print(f"\n{Fore.YELLOW}Exсel saved: {Fore.MAGENTA}{xls_file_name}")
    print(f"Full time: {Fore.GREEN}{round(time.time() - start_time)} {Fore.WHITE}sec;")
    print(
        f"Processed: {Fore.WHITE}Tags: {Fore.GREEN}{qnty_tags}{Fore.WHITE}; Values: {Fore.GREEN}{qnty_vals}{Fore.WHITE};")


def check_xml_file(file_name: str) -> bool:
    """Проверяет, что файл существует и XML"""
    file_xml = Path(file_name)
    if not file_xml.suffix.lower() == '.xml':
        err_msg = f"Not XML file {Fore.RED}{file_xml}"
        raise Exception(err_msg)

    if not file_xml.is_file():
        err_msg = f"File not found {Fore.RED}{file_xml}"
        raise Exception(err_msg)


def main():
    global MAX_ROWS_EXL
    parser = argparse.ArgumentParser(
        prog=f'ECS2XLS',
        description='ECS2XLS -  Utility to convert XML file (FLS ECS Trend) to XLSX',
        epilog=f'2023 7Art v{_VERSION}'
    )
    try:
        parser.add_argument('file', type=str, help='Trend export xml file name')
        parser.add_argument('-max_row', type=int, default=1000000, help='Maximum number of rows in an Excel sheet ('
                                                                        'default:1M)')
        args = parser.parse_args()
        MAX_ROWS_EXL = args.max_row
        file_xml = Path(args.file)
        check_xml_file(args.file)
        convert_xml2xls(args.file)


    except Exception as e:
        print(e)
        logger.info(e)
    finally:
        print("Press Enter to continue...")
        input()


if __name__ == '__main__':
    main()

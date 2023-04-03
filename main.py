import os
import urllib.request
import ssl
import openpyxl

import re
import openpyxl


def download_file(url, folder_path, file_name):
    # Создание папки, если она не существует
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

    # Полный путь к файлу
    file_path = os.path.join(folder_path, file_name)

    # Создание безопасного контекста SSL
    context = ssl.create_default_context()
    context.check_hostname = False
    context.verify_mode = ssl.CERT_NONE

    # Загрузка файла с помощью безопасного контекста SSL
    with urllib.request.urlopen(url, context=context) as u, open(file_path, 'wb') as f:
        f.write(u.read())

    print(f'Файл {file_name} загружен и сохранен в папке {folder_path}')


def read_zip_band_from_file(file_path, sheet_name):
    # Открытие файла Excel
    workbook = openpyxl.load_workbook(file_path, data_only=True)  # data_only=True - получение значений, а не формул

    # Получение листа
    sheet = workbook[sheet_name]

    # Получение всех строк
    rows = sheet.iter_rows()

    # {zip_start: [zip_start, zip_end], ...} = '00500': ['00500', '00599']
    zip_band_dict = {row[0].value.split("-")[0]:row[0].value.split("-") for row in rows if row[0].value}

    # [[zip_start, zip_end], [zip_start, zip_end], ...] = ['00500', '00599'], ['01000', '01099'], ['01100', '01199'],
    zip_band_list = [value for key, value in zip_band_dict.items()]

    # for row in rows:
    #     if row[0].value:
    #         print(row[0].value.split('-'))
    #

    print(zip_band_dict)
    print(zip_band_list)
    return zip_band_list, zip_band_dict

def read_excel_file(file_path):
    # Открытие файла Excel
    workbook = openpyxl.load_workbook(file_path)

    # Получение активного листа
    sheet = workbook.active

    # Получение 5 строки
    row = sheet[5]

    # Получение текста из ячеек строки
    row_text = [str(cell.value) for cell in row]

    # Поиск чисел в тексте строки
    pattern = r'\d{3}-\d{2}'
    matches = re.findall(pattern, ' '.join(row_text))

    # Вывод чисел
    if len(matches) == 2:
        print(matches[0], matches[1])
    else:
        print('Найдено неверное количество чисел в строке')



if __name__ == '__main__':
    ## 1) Загрузка файла
    url = 'https://www.ups.com/media/us/currentrates/zone-csv/011.xls'
    folder_path = 'zip_code'
    # Сразу переименовываем файл в *.xlsx
    file_name = '011.xlsx'

    download_file(url, folder_path, file_name)

    ## 2) Чтение файла Excel
    # file_path = 'Inbox Data/Carriers zone ranges.xlsx'
    # sheet_name = 'UPS zip ranges'
    #
    # read_zip_band_from_file(file_path, sheet_name)

    ## 3) Чтение полученого файла Excel и поиск чисел
    file_path = 'zip_code/011.xlsx'
    read_excel_file(file_path)
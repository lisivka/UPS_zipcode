import os
import urllib.request
import ssl
import openpyxl
import re
from openpyxl import Workbook


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
    zip_band_dict = {row[0].value.split("-")[0]: row[0].value.split("-") for row in rows if row[0].value}

    # [[zip_start, zip_end], [zip_start, zip_end], ...] = ['00500', '00599'], ['01000', '01099'], ['01100', '01199'],
    zip_band_list = [value for key, value in zip_band_dict.items()]
    # print(zip_band_dict)
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
        ref_start = matches[0].replace("-", "")  # Удаляем символ "-" 012-34 -> 01234 Референсный диапазон zip кодов
        ref_end = matches[1].replace("-", "")
        print(ref_start, ref_end)
    else:
        print('Найдено неверное количество чисел в строке')

    return ref_start, ref_end


def download_all_files(zip_band_list, url, folder_path, count_files=None):
    count = 0
    for index, zip_band in enumerate(zip_band_list[1:]):  # Пропускаем первый элемент - это заголовок
        zip_start = zip_band[0]
        zip_end = zip_band[1]
        name = zip_start[:-2]  # Удаляем последние 2 символа для загрузки файла например 011.xls

        url_file = url + f'{name}.xls'  ## url_file = f'https://www.ups.com/media/us/currentrates/zone-csv/{name}.xls'

        file_name = f'{name}.xlsx'  # Сразу переименовываем файл в *.xlsx
        download_file(url_file, folder_path, file_name)
        check, ref_start, ref_end = check_zip_code_from_load_file(folder_path, file_name, zip_start, zip_end)
        if check == False:
            expand_zip_band_list(zip_band_list, index, ref_start, ref_end, zip_start, zip_end)


        # Отключить после тестирования
        count += 1
        if count == count_files:
            break
        # ----------------------------

def expand_zip_band_list(zip_band_list, index, ref_start, ref_end, zip_start, zip_end):
    len_r_start = len(ref_start)-len(str(int(ref_start)))
    len_r_end = len(ref_end) - len(str(int(ref_end)))
    len_z_end = len(zip_end) - len(str(int(zip_end)))

    zip_band_list[index+1] = ["0"*len_r_start+str(int(ref_start) - 1),"0"*len_r_end+str( int(ref_end))]
    zip_band_list.insert(index + 2, ["0"*len_r_end+str(int(ref_end)+1), "0"*len_z_end+str(int(zip_end))])



    pass

def check_zip_code_from_load_file(folder_path, file_name, zip_start, zip_end):
    # Чтение полученого файла Excel и поиск диапазона чисел
    file_path = folder_path + '/' + file_name
    ref_start, ref_end = read_excel_file(file_path)
    # 00501-1 <= 00500 <= 00599 and 00599 >= 00599
    if int(ref_start) - 1 <= int(zip_start) <= int(ref_end) and int(zip_end) == int(ref_end):

        print('Диапазоны совпадают')
        answer = True
    else:
        print('Диапазоны не совпадают')

        print(f'Диапазон zip кодов из файла {file_name} = {ref_start} - {ref_end}')
        print(f'Диапазон zip кодов из файла Carriers zone ranges.xlsx = {zip_start} - {zip_end}')
        answer = False
    return answer, ref_start, ref_end

def write_data_to_txt(file_name):
    with open(file_name, "w") as f:
        f.write(f"UPS zone ranges,	zip from,	zip to\n")
        for item in zip_band_list[1:]:
            f.write(f"{item[0]}-{item[1]},{item[0]}, {item[1]}\n")



def txt_to_xlsx(input_file, output_file):

    wb = Workbook()
    ws = wb.active # New sheet
    ws.title = "UPS zip ranges"
    with open(input_file, 'r') as f:
        lines = f.readlines()
        for row_idx, line in enumerate(lines):

            # Split line into elements
            elements = line.strip().split(',')
            # Save elements to Excel cells
            for col_idx, element in enumerate(elements):
                ws.cell(row=row_idx+1, column=col_idx+1, value=element)
    wb.save(output_file)
    print("File saved!")




if __name__ == '__main__':
    ## 1) Чтение входящего файла Excel получение  диапазонов zip кодов
    file_path = 'Inbox Data/Carriers zone ranges.xlsx'
    sheet_name = 'UPS zip ranges'
    zip_band_list, zip_band_dict = read_zip_band_from_file(file_path, sheet_name)


    # ## 2) Загрузка файлов и проверка диапазонов zip кодов из загруженных файлов
    url = r'https://www.ups.com/media/us/currentrates/zone-csv/'
    folder_path = 'Output Data'

    # # После тестирования установить None
    count_files = 20  # Количество файлов для загрузки (для тестирования) None - все файлы
    download_all_files(zip_band_list, url, folder_path, count_files)

    ## 3) Запись в файл исправленого диапазона Для простоты проверки в IDE запись в txt, а затем в xlsx
    file_txt = f"{folder_path}/output.txt"
    write_data_to_txt(file_txt)
    txt_to_xlsx(file_txt, f"{folder_path}/NEW Carriers zone ranges.xlsx")


    print(zip_band_list)




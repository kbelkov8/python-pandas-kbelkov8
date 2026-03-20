from excel import (
    Excel_read,
)


def run():

    file_name = input('Введите название файла без ".xlsx": ')

    try:
        Excel_read(file_name).new_master_table()
        print(Excel_read(file_name).calibration_report())
    except FileNotFoundError:
        print('Данный файл не найден\n')

        run()

run()
import os
import sys
import pandas as pd

import satu


tilda_heads = {'Tilda UID': "UID",
               'Brand': "Бренд",
               'SKU': "Артикль",
               'Mark': "Значение",
               'Category': "Категория",
               'Title': "Заголовок",
               'Description': "Описание",
               'Text': "Текст",
               'Photo': "Ссылкана фото",
               'Price': "Цена",
               'Quantity': "Количество",
               'Price Old': "Старая цена",
               'Editions': "Варинт",
               'Modifications': "Модификация",
               'External ID': "Внутренний ID",
               'Parent UID': "Родительский UID"}


def show_help(param):
    if param == '-h' or param == '--help' or param == '/?' or param == '?':
        print("Для выполнения скрипта добавьте название обрабатываемого фалйа:")
        print("Например:")
        print("\t\txls2msrktplss file.xls")
        print("\t\txls2msrktplss file.xlsx")
        print()
        print('------------------')
        print('Для получения справки (этого экрана) введите ключ один из ключей:')
        print('\t-h, --help, /?, ?')
        print('либо запустите приложение без дополнительных аргументов')
        exit()


def main():
    if os.path.exists('tilda') == False:
        os.mkdir('tilda')
    if os.path.exists('satu') == False:
        os.mkdir('satu')

    param = []
    for params in sys.argv:
        param.append(params)
    if (len(param) < 2):
        param.append('/?')
    if (param[1].split('.')[-1] != 'xls') and (param[1].split('.')[-1] != 'xlsx'):
        param[1] = '/?'
    show_help(param[1])
    try:
        data_set = pd.ExcelFile(param[1])
    except:
        print('ОШИБКА! Фаил ', param[1],
              ' поврежден или не является таблицей. Проверти целостность файла или его версию.')
        print()
        print(param[1])
        print('------------------')
        show_help('/?')

    for sheet_name in data_set.sheet_names:

        if (sheet_name == "Цены") or (sheet_name == "Оглавление"):
            continue
        income_data_sheet = data_set.parse(sheet_name)
        print(sheet_name)

        k = 0
        for column in income_data_sheet.columns:
            income_data_sheet = income_data_sheet.rename(columns={column: income_data_sheet.iat[0, k]})
            k = k + 1
        income_data_sheet = income_data_sheet.drop(0)
        empty_list = []

        income_data_sheet = income_data_sheet.drop(columns='Подкатегория 1')
        income_data_sheet = income_data_sheet.drop(columns='Подкатегория 2')

        k = 1
        while k <= len(income_data_sheet):
            empty_list.append('')
            k = k + 1

        tilda_data_set = pd.DataFrame({})
        for new_columns in tilda_heads:
            income = new_columns
            outcome = tilda_heads[new_columns]
            try:
                tilda_data_set[income] = income_data_sheet[outcome]
                income_data_sheet = income_data_sheet.drop(columns=[outcome])
            except:
                tilda_data_set[income] = empty_list

        for col in income_data_sheet.head():
            tilda_data_set[col] = income_data_sheet[col]
            new_col = 'Characteristics:' + col
            tilda_data_set = tilda_data_set.rename(columns={col: new_col})
            # print(col)

        # income_data_sheet.to_excel((sheet_names[8]+'.xlsx'), index=False)
        print(tilda_data_set.head())

        tilda_data_set.to_csv(('tilda/tilda - ' + sheet_name + '.csv'), index=False, sep=";")


        """
        Следующая строка собирают и создает файла для импорта в маркетплейс satu.
        Если такой необходимости нет, сделайте ее комментарием, либо удалите.         
        """
        products_out = satu.satu(tilda_data_set, sheet_name)   # строка вызыакт сбор таблиц для сату

        print("ВНИМАНИЕ!!! Не забудьте добавить или удалить обязательные для каждого сервиса поля с вашими техническими данными")


if __name__ == '__main__':
    main()

from typing import List, Any
import sys
import pandas as pd
import logging


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

def fulling(len_products_in, fuller):
    raw = fuller
    k = 0
    while k < len_products_in:
        raw.append(fuller + 'a')
    return raw


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
    global data_set
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

        products = tilda_data_set
        products_in = products.head(0)
        products_out = pd.DataFrame()
        len_products_in = len(products_in.head())

        products_out["Код_товара"] = []
        products_out["Название_позиции"] = []
        products_out["Поисковые_запросы"] = []
        products_out["Номер_группы"] = []
        products_out["Название_группы"] = []
        products_out["Адрес_подраздела"] = []
        products_out["Идентификатор_подраздела"] = []
        products_out["ID_группы_разновидностей"] = []
        products_out["Описание"] = []
        products_out["Тип_товара"] = []
        products_out["Наличие"] = []
        products_out["Цена"] = []
        products_out["Валюта"] = []
        products_out["Единица_измерения"] = []
        products_out["Ссылка_изображения"] = []

        products_out_col_number = 6
        products_in_col_number = 0

        for all in products_in.columns:
            all = all.replace('Characteristics: ', '')
            all = all.replace('Characteristics:', '')

            if all == "Category":
                products_out["Название_группы"] = products["Category"]

            elif all == "Price":
                products_out["Цена"] = products["Price"]
                products_out["Валюта"] = fulling(len_products_in, "KZT")
                products_out["Единица_измерения"] = fulling(len_products_in, "шт.")


            elif all == "Photo":
                products_out["Ссылка_изображения"] = products[all]


            elif all == "Title":
                products_out["Название_позиции"] = products[all]


            elif all == "Description":
                products_out["Описание"] = products[all]


            elif all == "Brand":
                products_out["Производитель"] = products[all]


            elif all == "Страна производства":
                try:
                    try:
                        products_out["Страна_производитель"] = products['Characteristics:' + all]
                    except:
                        products_out["Страна_производитель"] = products['Characteristics: ' + all]
                except:
                    products_out["Страна_производитель"] = products[all]


            else:
                products_out["A"] = fulling(len_products_in, all)
                products_out["B"] = ''
                try:
                    try:
                        products_out["C"] = products['Characteristics:' + all]
                    except:
                        products_out["C"] = products['Characteristics: ' + all]
                except:
                    products_out["C"] = products[all]
                products_out = products_out.rename(
                    columns={'A': 'Название_характеристики', 'B': 'Измерение_характеристики',
                             'C': 'Значение_характеристики'})
        """        products_out["Название_характеристики"]
                products_out["Измерение_характеристики"]
                products_out["Значение_характеристики"]
                """

        products_out["Наличие"] = fulling(len_products_in, "+")
        products_out["Тип_товара"] = fulling(len_products_in, "r")

        tilda_data_set.to_csv(('tilda - ' + sheet_name + '.csv'), index=False, sep=";")
        products_out.to_excel(('satu - ' + sheet_name + '.xlsx'), index=False)
        
        print("ВНИМАНИЕ!!! Не забудьте добавить или удалить обязательные для каждого сервиса поля с вашими техническими данными")


if __name__ == '__main__':
    main()

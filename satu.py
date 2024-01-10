import pandas as pd

def fulling(len_products_in, fuller):
    raw = fuller
    k = 0
    while k < len_products_in:
        raw.append(fuller + 'a')
    return raw

def satu (products, sheet_name):
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
        all = all.replace('Characteristics:', '')
        except_list = ('SKU', 'Mark', 'Text', 'Quantity', 'Price Old', 'Editions', 'Modifications', 'External ID', 'Parent UID', 'Скадкав %')
        if all in except_list:
            continue

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

        elif all == "Страна" or all == "Страна ":
            products_out["Страна_производитель"] = products['Characteristics:' + all]

        elif all == "Номер_группы" or all == "Номер_группы ":
            products_out["Номер_группы"] = products['Characteristics:' + all]

        elif all == "Название_группы" or all == "Название_группы ":
            products_out["Название_группы"] = products['Characteristics:' + all]

        elif all == "Адрес_подраздела" or all == "Адрес_подраздела ":
            products_out["Адрес_подраздела"] = products['Characteristics:' + all]

        elif all == "Идентификатор_подраздела" or all == "Идентификатор_подраздела ":
            products_out["Идентификатор_подраздела"] = products['Characteristics:' + all]

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

    products_out.to_excel(('satu/satu - ' + sheet_name + '.xlsx'), index=False)

    return (products_out)
# -*- coding: utf-8 -*-
import os
import sqlite3
from sqlite3 import Error
import pandas as pd
import openpyxl as op


def get_tables():
    os.chdir('./input')
    tables_list = os.listdir()
    print(tables_list)
    print('tables list was generated')
    return tables_list


def parsing_tables(tables_list):
    list_of_tabs = []
    # список заголовков большой общей таблицы
    list_of_headers = ["Номер заказа", "Статус", "Время создания заказа", "Время оплаты", "Стоимость товаров, Руб",
                       "Стоимость доставки, Руб", "Сумма заказа, Руб", "Скидка магазина, Руб", "Оплачено клиентом, Руб",
                       "Сумма возврата, Руб", "Возвраты", "Товары", "Артикул", "Id товаров",
                       "Примечания к заказу (покупателя)", "Примечания к заказу (продавца)", "Имя получателя", "Страна",
                       "Штат/провинция", "Город", "Адрес", "Индекс", "Телефон", "Способ доставки", "Отгрузка истекает",
                       "Трекинг номер", "Время отправки", "Время подтверждения покупателем"]
    # temlate_of_dataframe переменная для инициализации переменной frame_of_tables в формате DataFrame
    template_of_dataframe = {'Номер заказа': [], 'Статус': [], 'Время создания заказа': [], 'Время оплаты': [],
                             'Стоимость товаров, Руб': [], 'Стоимость доставки, Руб': [],
                             'Сумма заказа, Руб': [], 'Скидка магазина, Руб': [], 'Оплачено клиентом, Руб': [],
                             'Сумма возврата, Руб': [], 'Возвраты': [],
                             'Товары': [], 'Артикул': [], 'Id товаров': [], 'Примечания к заказу (покупателя)': [],
                             'Примечания к заказу (продавца)': [],
                             'Имя получателя': [], 'Страна': [], 'Штат/провинция': [], 'Город': [], 'Адрес': [],
                             'Индекс': [], 'Телефон': [],
                             'Способ доставки': [], 'Отгрузка истекает': [], 'Трекинг номер': [], 'Время отправки': [],
                             'Время подтверждения покупателем': []}
    # dictianary_to_write словарь используемый для построчной записи данных в DataFrame
    dictionary_to_write = {'Номер заказа': 0, 'Статус': 0, 'Время создания заказа': 0, 'Время оплаты': 0,
                           'Стоимость товаров, Руб': 0, 'Стоимость доставки, Руб': 0,
                           'Сумма заказа, Руб': 0, 'Скидка магазина, Руб': 0, 'Оплачено клиентом, Руб': 0,
                           'Сумма возврата, Руб': 0, 'Возвраты': 0,
                           'Товары': 0, 'Артикул': 0, 'Id товаров': 0, 'Примечания к заказу (покупателя)': 0,
                           'Примечания к заказу (продавца)': 0,
                           'Имя получателя': 0, 'Страна': 0, 'Штат/провинция': 0, 'Город': 0, 'Адрес': 0,
                           'Индекс': 0, 'Телефон': 0,
                           'Способ доставки': 0, 'Отгрузка истекает': 0, 'Трекинг номер': 0, 'Время отправки': 0,
                           'Время подтверждения покупателем': 0}

    frame_of_tables = pd.DataFrame(template_of_dataframe)
    print("start generating a large shared table")
    for table in tables_list:
        work_book = op.load_workbook(table)
        for name in work_book.sheetnames:
            sheet = work_book[name]
            maxrow = sheet.max_row  # Получение максимального значения строк данного листа книги
            maxcol = sheet.max_column  # Получение максимального столбцов строк данного листа книги
            # Номер заказа|Статус|Время создания заказа|Время оплаты|Стоимость товаров, Руб|Стоимость доставки, Руб|Сумма заказа, Руб|Скидка магазина, Руб|Оплачено клиентом, Руб|Сумма возврата, Руб|Возвраты|Товары|Артикул|Id товаров|Примечания к заказу (покупателя)|Примечания к заказу (продавца)|Имя получателя|Страна|Штат/провинция|Город|Адрес|Индекс|Телефон|Способ доставки|Отгрузка истекает|Трекинг номер|Время отправки|Время подтверждения покупателем|
            for row in sheet.iter_rows(min_row=5, min_col=1, max_row=maxrow, max_col=maxcol):
                index_of_headers = 0
                list_of_tabs.append(table)
                for cell in row:
                    if cell.value != None:
                        if cell.value == "":
                            dictionary_to_write[list_of_headers[index_of_headers]] = "Void"
                        else:
                            dictionary_to_write[list_of_headers[index_of_headers]] = cell.value
                    # else:
                    #     print("None")
                    #     dictionary_to_write[list_of_headers[index_of_headers]] = "Void"
                    index_of_headers += 1
                frame_of_tables = frame_of_tables.append(dictionary_to_write, ignore_index=True)
    frame_of_tables["Имя таблицы"] = list_of_tabs
    print("large shared table was generated :")
    print(frame_of_tables)
    return frame_of_tables


def generate_products_xlsx(frame_of_tables):
    # Поля которые необходимо включить в таблицу Товары.xlsx
    # Товары|Артикул|Количество
    data_frame_tovar = pd.DataFrame({'Товары': [], 'Количество': []})
    frame_of_tables = frame_of_tables[['Товары']]
    df_to_save = pd.DataFrame({'Товары': [], 'Количество': []})
    names = []
    nums = []
    dictionary_to_write = {'Товары': 0, 'Количество': 0}
    name_of_table = "Товары.xlsx"

    os.chdir('../output')
    list_of_names_tabs = os.listdir()
    print(list_of_names_tabs)

    if (name_of_table in list_of_names_tabs) == False:
        print(name_of_table+" не найдена в директории /output")
        print("Создание файла " + name_of_table)
        wb = op.Workbook()
        Sheet_name = wb.sheetnames
        wb.save(filename=name_of_table)
        print("Таблица " + name_of_table + " успешно создана")
    else:
        print(name_of_table + " уже создана")

    new_string = ""
    list_of_names = frame_of_tables['Товары'].values.tolist()

    for i in list_of_names:
        new_string = new_string + i + "\n"

    list_of_names = new_string.split('\n')

    while '' in list_of_names:
        list_of_names.remove('')

    list_of_names = [name[4:] for name in list_of_names]


    for name in list_of_names:
        last_index_of_name = name.index("- Количество:")
        name_of_position = name[:last_index_of_name]
        number = name[last_index_of_name:]
        names.append(name_of_position)
        number = number.replace('- Количество: ','')
        number = number[:number.index(' ')]
        nums.append(number)

    for index in range(len(names)):
        dictionary_to_write['Товары'] = names[index]
        dictionary_to_write['Количество'] = int(nums[index])
        data_frame_tovar = data_frame_tovar.append(dictionary_to_write, ignore_index=True)
    set_names = set(names)


    for index in range(len(set_names)):
        df = data_frame_tovar[data_frame_tovar['Товары'] == list(set_names)[index]]
        dictionary_to_write['Товары'] = list(set_names)[index]
        dictionary_to_write['Количество'] = df['Количество'].sum()
        df_to_save = df_to_save.append(dictionary_to_write, ignore_index=True)

    print(df_to_save)
    print("table goods (Tovar) was generated")

    df_to_save[['Товары', 'Количество']].to_excel(name_of_table, index=False)

    print("data was saved to " + name_of_table)


def generate_parcels_xlsx(frame_of_tables):
    # Поля таблицы Посылки.xlsx
    # Номер заказа | Статус | Время создания заказа | Время оплаты | Стоимость Товаров | Стоимость доставки
    # | Сумма заказа| Скидка магазина | Оплачено клиентом | Сумма возврата | Возвраты | Товары
    # | Артикул | Id товара | Примечяние заказа (покупателя) | Примечание заказа (продавца) | Имя получателя
    # | Штат/провинция | Страна | Город | Телефон | Номер трекинга
    list_of_headers = ["Номер заказа", "Статус", "Время создания заказа", "Время оплаты", "Стоимость товаров, Руб",
                       "Стоимость доставки, Руб", "Сумма заказа, Руб", "Скидка магазина, Руб", "Оплачено клиентом, Руб",
                       "Сумма возврата, Руб", "Возвраты", "Товары", "Артикул", "Id товаров",
                       "Примечания к заказу (покупателя)", "Примечания к заказу (продавца)", "Имя получателя",
                       "Страна", "Штат/провинция", "Город", "Адрес", "Индекс", "Телефон",
                       "Способ доставки", "Отгрузка истекает", "Трекинг номер"]

    template_dict = {"Номер заказа": [], "Статус": [], "Время создания заказа": [], "Время оплаты": [],
                     "Стоимость товаров, Руб": [], "Стоимость доставки, Руб": [], "Сумма заказа, Руб": [],
                     "Скидка магазина, Руб": [], "Оплачено клиентом, Руб": [], "Сумма возврата, Руб": [],
                     "Возвраты": [], "Товары": [], "Артикул": [], "Id товаров": [],
                     "Примечания к заказу (покупателя)": [],"Примечания к заказу (продавца)": [], "Имя получателя": [],
                     "Страна": [], "Штат/провинция": [], "Город": [], "Адрес": [], "Индекс": [], "Телефон": [],
                     "Способ доставки": [],"Отгрузка истекает": [], "Трекинг номер": []}

    dictionary_to_write = {"Номер заказа": 0, "Статус": 0, "Время создания заказа": 0, "Время оплаты": 0,
                     "Стоимость товаров, Руб": 0, "Стоимость доставки, Руб": 0, "Сумма заказа, Руб": 0,
                     "Скидка магазина, Руб": 0,"Оплачено клиентом, Руб": 0, "Сумма возврата, Руб": 0,
                     "Возвраты": 0, "Товары": 0, "Артикул": 0,
                     "Id товаров": 0, "Примечания к заказу (покупателя)": 0,
                     "Примечания к заказу (продавца)": 0, "Имя получателя": 0, "Страна": 0,
                     "Штат/провинция": 0, "Город": 0, "Адрес": 0, "Индекс": 0,
                     "Телефон": 0, "Способ доставки": 0, "Отгрузка истекает": 0, "Трекинг номер": 0}

    data_frame_parcels = pd.DataFrame(template_dict)
    name_of_table = "Посылки.xlsx"

    os.chdir('../output')
    list_of_names_tabs = os.listdir()
    print(list_of_names_tabs)

    if (name_of_table in list_of_names_tabs) == False:
        print(name_of_table + " не найдена в директории /output")
        print("Создание файла " + name_of_table)
        wb = op.Workbook()
        Sheet_name = wb.sheetnames
        wb.save(filename=name_of_table)
        print("Таблица " + name_of_table + " успешно создана")
    else:
        print(name_of_table + " уже создана")


    shape = frame_of_tables.shape
    len_of_dataframe = shape[0]
    print("Table" + name_of_table + "generate...")

    for index in range(len_of_dataframe):
        #print("----------------------------------------")
        for header_index in range(len(list_of_headers)):
            #dictionary_to_write["Номер заказа"] = frame_of_tables.loc[index, "Номер заказа"]
            #print(list_of_headers[header_index])
            dictionary_to_write[list_of_headers[header_index]] = frame_of_tables.loc[index, list_of_headers[header_index]]
        #print(dictionary_to_write)
        data_frame_parcels = data_frame_parcels.append(dictionary_to_write, ignore_index=True)
    #print(data_frame_parcels)

    print("Save table as Посылки.xlsx")
    data_frame_parcels.to_excel(name_of_table, index=False)

    print("data was saved to " + name_of_table)


def create_connection(path):
    connection = None
    try:
        connection = sqlite3.connect(path)
        print("Connection to SQLite DB successful")
    except Error as e:
        print(f"The error '{e}' occurred")

    return connection


def execute_query(connection, query):
    cursor = connection.cursor()
    try:
        cursor.execute(query)
        connection.commit()
        print("query executed successfully")
    except Error as e:
        print(f"The error '{e}' occurred")


def generate_SQLite():
    db_name = "out.db"
    os.chdir("../output")
    connection = create_connection(db_name)
    return connection, db_name


def dataframe_to_SQLite(frame_of_tables, connection, db_name):
    lsit_of_headers = ["order_id", "status", "order_credit_dt", "pay_dt", "goods_cost", "delivery_cost",
                       "order_cost", "shop_discount", "customer_paid", "refund", "returns", "products", "products_articles",
                       "product_ids", "buyer_note", "seller_note", "delivery_name", "delivery_country",
                       "delivery_state", "delivery_city", "delivery_address", "delivery_zip", "delivery_phone",
                       "delivery_type", "delivery_send_until_date", "delivery_tracking_number",
                       "delivery_send_dt", "delivery_receive_dt", "table_name"]

    rus_headers = ["Номер заказа", "Статус", "Время создания заказа", "Время оплаты", "Стоимость товаров, Руб",
                   "Стоимость доставки, Руб", "Сумма заказа, Руб", "Скидка магазина, Руб", "Оплачено клиентом, Руб",
                   "Сумма возврата, Руб", "Возвраты", "Товары", "Артикул", "Id товаров",
                   "Примечания к заказу (покупателя)", "Примечания к заказу (продавца)", "Имя получателя", "Страна",
                   "Штат/провинция", "Город", "Адрес", "Индекс", "Телефон", "Способ доставки", "Отгрузка истекает",
                   "Трекинг номер", "Время отправки", "Время подтверждения покупателем", "Имя таблицы"]

    print("write to database...")
    for index in range(len(lsit_of_headers)):

        frame_of_tables = frame_of_tables.rename(columns={rus_headers[index]: lsit_of_headers[index]})

    frame_of_tables.to_sql(db_name, con=connection, schema='dbo', if_exists='replace')
    print("successfully")


if __name__ == '__main__':

    list_of_names = get_tables()
    if len(list_of_names) == 0:
        print("/input пуста")
    else:
        frame_of_tables = parsing_tables(list_of_names)
        generate_products_xlsx(frame_of_tables)
        generate_parcels_xlsx(frame_of_tables)
        connection, db_name = generate_SQLite()
        dataframe_to_SQLite(frame_of_tables, connection, db_name)
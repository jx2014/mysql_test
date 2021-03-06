#!/usr/bin/python

import xlrd
#import MySQLdb
#import _mysql
import mysql.connector
import _mysql_connector
import time
import pandas as pd
import matplotlib.pyplot as plt

plt.style.use('ggplot')



#database = MySQLdb.connect(host='localhost', user='root', passwd='NewPass123!', db='mysqlPython', auth_plugin='mysql_native_password')
#database = _mysql.connect(host='localhost', user='root', passwd='NewPass123!', db='mysqlPython', auth_plugin='mysql_native_password')
#database = mysql.connector.connect(host='localhost', user='root', passwd='NewPass123!', db='mysqlPython')
#cursor = database.cursor()
#cursor.execute(create_table)



def col_to_num(col_letter):
    total_num = 0
    for i, c in enumerate(col_letter.upper()[::-1]):
        num = (ord(c) - ord('A') + 1) * (26 ** i)
        total_num = num + total_num
    return total_num - 1


def display_row(list_items):
    i = 0
    p = ''
    while i < len(list_items):
        p = p + "{%s:15} " % i
        i = i + 1
    print p.format(*list_items)


def read_mysql_result():
    col_names= []
    for i in database.fetch_fields():
        col_names.append(i[4])
    display_row(col_names)
    if database.have_result_set:
        row = database.fetch_row()
        while row:
            display_row(row)
            row = database.fetch_row()
    database.consume_result()


def upload_main_table(excel_file_path, sheet_name):
    book = xlrd.open_workbook(excel_file_path)
    sheet = book.sheet_by_name(sheet_name)
    create_table = """CREATE TABLE main_table 
                    (rma            int         NOT NULL,
                    macid           char(16)    NOT NULL,
                    customer        text        NOT NULL,
                    prod_type       char(255)   NOT NULL,
                    receive_date    datetime    NOT NULL,
                    fault_report    text   NOT NULL);
                    """
    # need to upload new table every time we have a new service report.
    try:
        database.query("""DROP TABLE main_table;""")
    except _mysql_connector.MySQLInterfaceError as e:
        print e
    finally:
        database.query(create_table)
    # upload cell values one by one.
    query = """INSERT INTO main_table (rma, macid, customer, prod_type, receive_date, fault_report)
            VALUES ('%s', '%s', '%s', '%s', '%s', '%s');"""
    for r in range(1, sheet.nrows):
        rma = sheet.cell(r, col_to_num('A')).value
        macid = sheet.cell(r, col_to_num('B')).value
        customer = sheet.cell(r, col_to_num('C')).value
        prod_type = sheet.cell(r, col_to_num('D')).value
        # receive_date = sheet.cell(r, col_to_num('E')).value
        receive_date = time.strftime('%Y/%m/%d',
                                     time.localtime((sheet.cell(r, col_to_num('E')).value - 25568) * 86400))
        fault_report = sheet.cell(r, col_to_num('F')).value
        values = (rma, macid, customer, prod_type, receive_date, fault_report)
        database.query(query % values)
    # consolidate the product names
    database.query("""ALTER TABLE main_table ADD prod_type_corrected VARCHAR(30);""")
    database.query("""SET SQL_SAFE_UPDATES = 0;""")
    database.query("""UPDATE main_table
                      SET prod_type_corrected = 'router 1.0' 
                      WHERE prod_type LIKE '%router%' AND prod_type LIKE '%1.0%';""")
    database.query("""UPDATE main_table
                      SET prod_type_corrected = 'router 2.0' 
                      WHERE prod_type LIKE '%router%' AND prod_type LIKE '%2.0%';""")
    database.query("""UPDATE main_table
                      SET prod_type_corrected = 'printer 2.0' 
                      WHERE prod_type LIKE '%printer%' AND prod_type LIKE '%2.0%';""")


def update_ship_date_table(excel_file_path, sheet_name):
    book = xlrd.open_workbook(excel_file_path)
    sheet = book.sheet_by_name(sheet_name)
    create_table = """CREATE TABLE ship_date_table 
                        (macid   char(16)    NOT NULL,
                        ship_date   datetime    NOT NULL);
                        """
    try:
        database.query("""SELECT * FROM ship_date_table LIMIT 5""")
        database.consume_result()
    except _mysql_connector.MySQLInterfaceError as e:
        print e
        database.query(create_table)
    query = """INSERT INTO ship_date_table (macid, ship_date)
            VALUES ('%s', '%s');"""
    for r in range(1, sheet.nrows):
        macid = sheet.cell(r, col_to_num('A')).value
        ship_date = time.strftime('%Y/%m/%d',
                                     time.localtime((sheet.cell(r, col_to_num('B')).value - 25568) * 86400))
        values = (macid, ship_date)
        database.query(query % values)



def upload_total_shippment_table(excel_file_path, sheet_name):
    book = xlrd.open_workbook(excel_file_path)
    sheet = book.sheet_by_name(sheet_name)
    create_table = """CREATE TABLE total_shipment_table 
                    (prod_type   char(16)    NOT NULL,
                    ship_year    int    NOT NULL,                    
                    ship_qty    int NOT NULL);
                    """
    # need to upload new table every time we have a new service report.
    try:
        database.query("""DROP TABLE total_shipment_table;""")
    except _mysql_connector.MySQLInterfaceError as e:
        print e
    finally:
        database.query(create_table)
    # upload cell values one by one.
    query = """INSERT INTO total_shipment_table (prod_type, ship_year, ship_qty)
                VALUES ('%s', '%s', '%s');"""
    for r in range(1, sheet.nrows):
        prod_type = sheet.cell(r, col_to_num('A')).value
        ship_year = sheet.cell(r, col_to_num('B')).value
        ship_qty = sheet.cell(r, col_to_num('C')).value
        values = (prod_type, ship_year, ship_qty)
        database.query(query % values)

def query_DOM(product_name):
    query_DOM = """SELECT DATE_FORMAT(receive_date, "%Y") as Ship_Years, COUNT(*) as Counts FROM main_table WHERE main_table.prod_type_corrected='%s' GROUP BY DATE_FORMAT(receive_date, "%Y");""" % product_name

if __name__ == "__main__":
    # connect to database using c API
    database = _mysql_connector.MySQL()
    database.connect(host='localhost', user='root', password='NewPass123!', database='mysqlPython')

    # assigning spread sheet location
    main_spreadsheet = r"test_data/main_spreadsheet.xlsx"
    main_worksheet = "Sheet1"
    ship_date_spreadsheet = r"test_data/ship_date_spreadsheet.xlsx"
    ship_date_worksheet = "Sheet1"
    total_shipment_spreadsheet = r"test_data/total_shipment_spreadsheet.xlsx"
    total_shipment_worksheet = "Sheet1"

    # upload spreadsheets
    upload_main_table(main_spreadsheet, main_worksheet)
    update_ship_date_table(ship_date_spreadsheet, ship_date_worksheet)
    upload_total_shippment_table(total_shipment_spreadsheet, total_shipment_worksheet)

    # not needed, for debug only.
    database.query("""SELECT * FROM main_table;""")
    read_mysql_result()

    # finish up database.
    database.commit()
    database.close()

    # connect to database using mysql connector for easy processing.
    database = mysql.connector.connect(host='localhost', user='root', passwd='NewPass123!', db='mysqlPython', auth_plugin='mysql_native_password')
    cursor = database.cursor()

    df = pd.read_sql("SELECT * FROM main_table", con=database)
    # read database into panda data frame for plotting purposes.
    df = pd.read_sql("SELECT COUNT(*) as Counts, prod_type_corrected as Products FROM main_table GROUP BY prod_type_corrected", con=database)
    df.set_index('Products', inplace=True)
    df.sort_values(by=['Counts'], ascending=False, inplace=True)
    df.plot.bar()

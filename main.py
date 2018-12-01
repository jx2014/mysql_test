#!/usr/bin/python

import xlrd
#import MySQLdb
#import _mysql
#import mysql.connector
import _mysql_connector
import time

excel_spreadsheet = r"test_data/main_spreadsheet.xlsx"
excel_worksheet = "Sheet1"

book = xlrd.open_workbook(excel_spreadsheet)
sheet = book.sheet_by_name(excel_worksheet)

#database = MySQLdb.connect(host='localhost', user='root', passwd='NewPass123!', db='mysqlPython', auth_plugin='mysql_native_password')
#database = _mysql.connect(host='localhost', user='root', passwd='NewPass123!', db='mysqlPython', auth_plugin='mysql_native_password')
#database = mysql.connector.connect(host='localhost', user='root', passwd='NewPass123!', db='mysqlPython')
database = _mysql_connector.MySQL()
database.connect(host='localhost', user='root', password='NewPass123!', database='mysqlPython')

#cursor = database.cursor()

create_table = """CREATE TABLE main_spreadsheet 
                (rma            int         NOT NULL,
                macid           char(16)    NOT NULL,
                customer        text        NOT NULL,
                prod_type       char(255)   NOT NULL,
                receive_date    datetime    NOT NULL,
                fault_report    text   NOT NULL);
                """
database.query(create_table)
#cursor.execute(create_table)

query = """INSERT INTO main_spreadsheet (rma, macid, customer, prod_type, receive_date, fault_report)
        VALUES ('%s', '%s', '%s', '%s', '%s', '%s');"""


def col_to_num(col_letter):
    total_num = 0
    for i, c in enumerate(col_letter.upper()[::-1]):
        num = (ord(c) - ord('A') + 1) * (26 ** i)
        total_num = num + total_num
    return total_num - 1


for r in range(1, sheet.nrows):
    rma = sheet.cell(r, col_to_num('A')).value
    macid = sheet.cell(r, col_to_num('B')).value
    customer = sheet.cell(r, col_to_num('C')).value
    prod_type = sheet.cell(r, col_to_num('D')).value
    #receive_date = sheet.cell(r, col_to_num('E')).value
    receive_date = time.strftime('%Y/%m/%d', time.localtime((sheet.cell(10, col_to_num('E')).value - 25568) * 86400))
    fault_report = sheet.cell(r, col_to_num('F')).value
    values = (rma, macid, customer, prod_type, receive_date, fault_report)
    database.query(query % values)
    #cursor.execute(query, values)


#cursor.close()

database.commit()

database.close()
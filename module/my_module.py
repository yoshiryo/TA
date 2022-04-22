import openpyxl
import pprint
import datetime

def read_xlsx():
    wb = openpyxl.load_workbook('C:\\Users\\ryota\\Desktop\\TA\\input\\TA.xlsx')
    sheet = wb['出勤報告書']
    return sheet, wb

def make_xlsx():
    sheet, wb = read_xlsx()
    dt = datetime.datetime.now()
    year = str(dt.year)[-2:]
    month = str(dt.month).zfill(2)
    num = "0000000"
    name_kata = "ダイダイ　ダイ"
    name_kan = "大大大"
    phone_num = "000-0000-0000"
    for x, s in enumerate(num):
        sheet.cell(row = 6, column = 5 + x, value = s)
    sheet['E8'] = name_kata
    sheet['E9'] = name_kan
    sheet['V8'] = phone_num

    sheet['G11'] = year[0]
    sheet['H11'] = year[1]

    sheet['J11'] = month[0]
    sheet['K11'] = month[1]
    wb.save('C:\\Users\\ryota\\Desktop\\TA\\output\\test.xlsx')
"""
def write_number(sheet, num, start_row, start_col):
    for x,  in enumerate(num):]
"""
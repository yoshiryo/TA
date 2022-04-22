import openpyxl
import pprint

def read_xlsx():
    wb = openpyxl.load_workbook('C:\\Users\\ryota\\Desktop\\TA\\input\\TA.xlsx')
    sheet = wb['出勤報告書']
    print(sheet)
import glob
import os

from openpyxl import Workbook
from simplified_scrapy import SimplifiedDoc, utils
import pandas as pd
import xlwt

till_file = "Order Application Project/Put File Here/"
till_file_dir = "Order Application Project/Put File Here/"
xlsx_file = 'Order Application Project/Converted Files/xml_to_xlsx.xlsx'
xls_file = 'Order Application Project/Converted Files/xlsx_to_xls.xls'


def get_xml_file():
    global till_file
    path = till_file_dir

    for (root, dirs, files) in os.walk(path):
        for f in files:
            if '.xml' in f:
                till_file = till_file_dir + f


def read_file(filename):
    xml = utils.getFileContent(filename)
    doc = SimplifiedDoc(xml)
    tables = doc.selects('Worksheet').selects('Row').selects('Cell').text  # Get all data
    sheetNames = doc.selects('Worksheet>ss:Name()')  # Get sheet name
    return sheetNames, tables


def to_excel(sheetNames, tables):
    wb = Workbook()  # Create Workbook

    for i in range(len(sheetNames)):
        worksheet = wb.create_sheet(sheetNames[i])  # Create sheet
        for row in tables[i]:
            worksheet.append(row)

    wb.remove(wb['Sheet'])
    wb.save(xlsx_file)  # Save file


def to_xls():
    cata = pd.read_excel(xlsx_file, sheet_name='PLUGRP')
    plu = pd.read_excel(xlsx_file, sheet_name='PLU')

    with pd.ExcelWriter(xls_file) as writer:
        cata.to_excel(writer, sheet_name='PLUGRP', engine='xlsxwriter', index=False)
        plu.to_excel(writer, sheet_name='PLU', engine='xlsxwriter', index=False)


def convert_file():
    get_xml_file()
    to_excel(*read_file(till_file))
    to_xls()

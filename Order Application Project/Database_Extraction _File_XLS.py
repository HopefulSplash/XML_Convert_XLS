# Program to extract number
# of rows using Python
import os
from json import JSONEncoder
import Database_Extraction_File_XML as def_xml
import xlrd
import json
import numpy

file_path = "Order Application Project/Converted Files/xlsx_to_xls.xls"
json_file = "Order Application Project/JSON_File_Here/JSON_DATA.json"

global plu_group_index
global shop_product_index
global wb
global cata_ws
global prod_ws

global cata_id_name
global prod_plu_name_cata_price_mod

clear_console = lambda: os.system('cls' if os.name in ('nt', 'dos') else 'clear')


def load_workbook(file_find):
    # Give the location of the file
    return xlrd.open_workbook(file_find)


def get_worksheets(workbook_search):
    global plu_group_index
    global shop_product_index

    for index, sheet in enumerate(workbook_search.sheets()):
        if wb.sheet_by_index(index).name == 'PLUGRP':
            plu_group_index = index
        elif wb.sheet_by_index(index).name == 'PLU':
            shop_product_index = index


def get_worksheet(cata_index, prod_index):
    global cata_ws
    global prod_ws

    cata_ws = wb.sheet_by_index(cata_index)
    prod_ws = wb.sheet_by_index(prod_index)


def get_prod_list():
    prod_plu_index = 0
    prod_name_index = 0
    prod_cata_index = 0
    prod_price_index = 0
    prod_price_mod_index = 0

    for col in range(prod_ws.ncols):
        if "PLU Number:PLU" in prod_ws.cell_value(1, col):
            prod_plu_index = col
        elif "Display Text:DYT" in prod_ws.cell_value(1, col):
            prod_name_index = col
        elif "Group ID:LGID" in prod_ws.cell_value(1, col):
            prod_cata_index = col
        elif "Standard Price:P1" in prod_ws.cell_value(1, col):
            prod_price_index = col
        elif "Price Modifier Divider:PMD" in prod_ws.cell_value(1, col):
            prod_price_mod_index = col

    for row in range(prod_ws.nrows):
        if row > 2:
            plu_val = int(prod_ws.cell_value(row, prod_plu_index))
            name_val = (prod_ws.cell_value(row, prod_name_index))
            cata_val = int(prod_ws.cell_value(row, prod_cata_index))
            price_val = round(float(prod_ws.cell_value(row, prod_price_index)), 2)
            mod_val = int(prod_ws.cell_value(row, prod_price_mod_index))
            prod_plu_name_cata_price_mod.append([plu_val, name_val, cata_val, price_val, mod_val])


def get_cata_list():
    cata_name_index = 0
    cata_id_index = 0

    for col in range(cata_ws.ncols):
        if "Group ID:LGID" in cata_ws.cell_value(1, col):
            cata_id_index = col
        elif "Description:DESC" in cata_ws.cell_value(1, col):
            cata_name_index = col

    for row in range(cata_ws.nrows):
        if row > 1:
            cata_name_val = (cata_ws.cell_value(row, cata_name_index))
            cata_id_val = int(cata_ws.cell_value(row, cata_id_index))
            cata_id_name.append([cata_id_val, cata_name_val])


class NumpyArrayEncoder(JSONEncoder):
    def default(self, obj):
        if isinstance(obj, numpy.ndarray):
            return obj.tolist()
        return JSONEncoder.default(self, obj)


def setup_json_data():
    json_directory_prod = [{}]
    json_directory_prod.pop()
    json_directory_cata = [{}]
    json_directory_cata.pop()

    for p_array in prod_plu_name_cata_price_mod:
        json_directory_prod.append([{
            "prod_plu": p_array[0],
            "prod_name": p_array[1],
            "prod_cata_id": p_array[2],
            "prod_price": p_array[3],
            "prod_price_mod": p_array[4]

        }])

    for c_array in cata_id_name:
        json_directory_cata.append([{
            "category": {
                "cata_id": c_array[0],
                "cata_name": c_array[1],
            }
        }])


def write_json():
    numpy_array_cata = numpy.array(cata_id_name)
    numpy_array_prod = numpy.array(prod_plu_name_cata_price_mod)
    numpyData = {"category": numpy_array_cata, "product": numpy_array_prod}
    with open(json_file, "w") as write_file:
        json.dump(numpyData, write_file, cls=NumpyArrayEncoder)


def read_json():
    with open(json_file, "r") as read_file:
        decodedArray = json.load(read_file)

        finalNumpyArrayOne = numpy.asarray(decodedArray["product"])
        finalNumpyArrayTwo = numpy.asarray(decodedArray["category"])


if __name__ == "__main__":
    clear_console()
    print("Processing................")
    def_xml.convert_file()
    clear_console()
    print("Processing................")
    cata_id_name = [[]]
    prod_plu_name_cata_price_mod = [[]]
    cata_id_name.pop()
    prod_plu_name_cata_price_mod.pop()
    clear_console()
    print("Processing................")
    wb = load_workbook(file_path)
    get_worksheets(wb)
    get_worksheet(plu_group_index, shop_product_index)
    clear_console()
    print("Processing................")
    get_cata_list()
    get_prod_list()
    setup_json_data()
    clear_console()
    print("Processing................")
    write_json()
    read_json()
    clear_console()
    print("*** Completed ***")


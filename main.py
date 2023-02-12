import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Border, Side
from Functions import get_files, map_files_with_keys, extracting_files_names_from_excel, find_incorrect_file_name, create_modify_excel_sheet, import_data_in_excel
files = get_files()

files_dict = map_files_with_keys(files)

file_number = int(input('Please select file: >'))

list_files_names = extracting_files_names_from_excel(files_dict, file_number)


# name_format = {1:4, 2:3, 3:2, 4:2, 5:2, 6:3, 7:4, 8:range(1, 100)}
incorrect_names = []
incorrect_items = []
for file_name in list_files_names:
    find_incorrect_file_name(file_name, incorrect_names, incorrect_items)

# print(incorrect_names)
# print(incorrect_items)

wb, ws = create_modify_excel_sheet()

import_data_in_excel(ws, incorrect_names, incorrect_items)

wb.save('Incorrect_Naming.xlsx')
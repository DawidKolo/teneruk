import re
import os

import xlsxwriter
from openpyxl import load_workbook



path = "C:\\test\\000"
path_out = "C:\\test\\001"

filelist = os.listdir(path)
xlsx_name = "\\new_app.xlsx"
pattern = ["P1", "P2", "P3", "P4", "P5", "P6", "P7", "P8"]

def check_xls_file():
    if not os.path.isfile(path_out + xlsx_name):
        excel = xlsxwriter.Workbook(path_out + xlsx_name)
        excel.close()
    else:
        pass
    return

def collect_txt_filenames():
    txt_files = []
    for file in filelist:
        if file.endswith(".txt"):
            txt_files.append(file)
    return txt_files

def create_sheetnames():
    txt_files = collect_txt_filenames()
    sheet_names = []
    for txt_file in txt_files:
        x = re.findall("[0-9]", txt_file)
        x.insert(4, "-")
        x.insert(7, "-")

        date = "".join(x)
        sheet_name = date[0:7]
        sheet_names.append(sheet_name)

    return sheet_names, date

def search_for_p():
    txt_files = collect_txt_filenames()
    temp_list = []
    for txt_file in txt_files:
        f = open(path + "\\" + txt_file, "r")
        files = f.read()

        p_table = []

        for p in pattern:
            match_p = re.findall(p, files)
            p_table.append(len(match_p))
        temp_list.append(p_table)
        f.close()

    return temp_list

def write_sheetname_to_wb():

    sheetname = list(dict.fromkeys(create_sheetnames()[0]))
    for c in range(len(sheetname)):
        myFilename = path_out + xlsx_name

        wb = load_workbook(filename=myFilename)

        if not sheetname in wb.sheetnames:
            wb.create_sheet(sheetname[c])

        wb.save(filename=myFilename)
        wb.close()


def insert_values_to_spreadsheet():
    sheetname = list(dict.fromkeys(create_sheetnames()[0]))
    myFilename = path_out + xlsx_name
    workbook = load_workbook(filename=myFilename)



    for w in range(len(sheetname)):
        ws = workbook[sheetname[w]]

    newRowLocation = ws.max_row + 1

    for n in range(0, 8):
        ws.cell(column=2+n, row=2, value=pattern[n])

    workbook.save(filename=myFilename)
    workbook.close()
    return ws

write_sheetname_to_wb()
insert_values_to_spreadsheet()








#print(create_sheetnames())
#write_sheetname_to_wb()

import re
import os
import datetime

import xlsxwriter
from openpyxl import load_workbook


# variables
path = "C:\\test\\000"
path_out = "C:\\test\\001"

filelist = os.listdir(path)
xlsx_name = "\\new_app.xlsx"
pattern = ["P1", "P2", "P3", "P4", "P5", "P6", "P7", "P8"]


# checks if the xlsx file exists -> if not creates a new one
def check_xls_file():
    if not os.path.isfile(path_out + xlsx_name):
        excel = xlsxwriter.Workbook(path_out + xlsx_name)
        excel.close()
        return True


# gets filenames od txt files and adds them to the list, returns the list of strings
def collect_txt_filenames():
    txt_files = []
    for file in filelist:
        if file.endswith(".txt"):
            txt_files.append(file)
    return txt_files


# gets dates from the filenames and adds them to the list, returns list of integers
def get_week_of_year():
    date_for_week = create_sheetnames()[1]
    week_of_year = []
    for i in range(len(date_for_week)):

        year = int(date_for_week[i][:4])
        month = int(date_for_week[i][5:7])
        day = int(date_for_week[i][8:10])
        week = datetime.date(year, month, day).isocalendar()[1]
        week_of_year.append(week)

    return week_of_year


# creates a sheetnames out of filenames, returns a tuple of lists of strings
def create_sheetnames():
    txt_files = collect_txt_filenames()
    sheet_names = []
    timestamp = []
    for txt_file in txt_files:
        x = re.findall("[0-9]", txt_file)
        x.insert(4, "-")
        x.insert(7, "-")

        date = "".join(x)
        sheet_name = date[0:7]
        sheet_names.append(sheet_name)
        timestamp.append(date)

    return sheet_names, timestamp


# searches for P1,P2 etc. from the pattern list in txt files, returns a list
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


# writes sheetname to the workbook
def write_sheetname_to_wb():

    sheetname = list(dict.fromkeys(create_sheetnames()[0]))
    for c in range(len(sheetname)):
        myFilename = path_out + xlsx_name

        wb = load_workbook(filename=myFilename)

        if not sheetname in wb.sheetnames:
            wb.create_sheet(sheetname[c])

        wb.save(filename=myFilename)
        wb.close()

# Writes the P values to the spreadsheets
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









get_week_of_year()
print(create_sheetnames())
#print(create_sheetnames())
#write_sheetname_to_wb()

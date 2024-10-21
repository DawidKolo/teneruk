import re
import os
from openpyxl import load_workbook

# Define Input and output paths below
path = "C:\\Users\\Dawid.Kolodziej1\\OneDrive - Precisely Inc\\Desktop\\000"
path_out = "C:\\Users\\Dawid.Kolodziej1\\OneDrive - Precisely Inc\\Desktop\\001"

for File in os.listdir(path + "\\" + "."):
    if File.endswith(".txt"):

        fileList = os.listdir(path)[0]

        x = re.findall("[0-9]", fileList)
        x.insert(4, "-")
        x.insert(7, "-")


        date = "".join(x)
        sheet_name = date[0:7]


        f = open(path + "\\" + fileList, "r")
        files = f.read()
        pattern = ["P1", "P2", "P3", "P4"]
        p_table = []
        for p in pattern:
            match_p = re.findall(p, files)
            p_table.append(len(match_p))

        print(p_table)

        f.close()

        myFileName = path_out + "\\" + "DemoFile2.xlsx"

        #load the workbook, and put the sheet into a variable
        wb = load_workbook(filename=myFileName)

        if not sheet_name in wb.sheetnames:
            wb.create_sheet(sheet_name)

        ws = wb[sheet_name]

        #max_row is a sheet function that gets the last row in a sheet.
        newRowLocation = ws.max_row + 1

        #write to the cell you want, specifying row and column, and value :-)
        ws.cell(column=1, row=newRowLocation, value=date)
        ws.cell(column=2, row=newRowLocation, value=p_table[0])
        ws.cell(column=3, row=newRowLocation, value=p_table[1])
        ws.cell(column=4, row=newRowLocation, value=p_table[2])
        ws.cell(column=5, row=newRowLocation, value=p_table[3])


        wb.save(filename=myFileName)
        wb.close()
    else:
        print("No TXT file")

    os.rename(path + "\\" + fileList, path_out + "\\" + fileList)

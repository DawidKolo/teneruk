import re
import os
from openpyxl import load_workbook

path = ""
path_out = ""

for File in os.listdir(""):
    if File.endswith(".txt"):
        print("T")

        fileList = os.listdir(path)[0]
        print(fileList)
        x = re.findall("[0-9]", fileList)
        date = "".join(x)
        print(date)

        f = open(path+"\\"+fileList, "r")
        files = f.read()


        match_p1 = re.findall("P1", files)
        match_p2 = re.findall("P2", files)
        match_p3 = re.findall("P3", files)
        match_p4 = re.findall("P4", files)

        p_1 = len(match_p1)
        p_2 = len(match_p2)
        p_3 = len(match_p3)
        p_4 = len(match_p4)

        print("p1: "+ str(p_1))
        print("p2: "+ str(p_2))
        print("p3: "+ str(p_3))
        print("p4: "+ str(p_4))
        f.close()

        myFileName=""
    #load the workbook, and put the sheet into a variable
        wb = load_workbook(filename=myFileName)
        ws = wb['Sheet1']

    #max_row is a sheet function that gets the last row in a sheet.
        newRowLocation = ws.max_row +1

    #write to the cell you want, specifying row and column, and value :-)
        ws.cell(column=1,row=newRowLocation, value=date)
        ws.cell(column=2, row=newRowLocation, value=p_1 )
        ws.cell(column=3, row=newRowLocation, value=p_2 )
        ws.cell(column=4, row=newRowLocation, value=p_3 )
        ws.cell(column=5, row=newRowLocation, value=p_4 )
        wb.save(filename=myFileName)
        wb.close()
    else:
        print("No TXT file")
#os.rename(path+"\\"+fileList, path_out+"\\"+fileList)
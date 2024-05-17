import openpyxl
import  wget

path = "C:\\Users\\user\\PycharmProjects\\download\\data.xlsx"
wb_obj = openpyxl.load_workbook(path)
sheet_obj = wb_obj.active
for i in range(1, 1047):
    filename = sheet_obj.cell(row = i, column = 1).value
    downloadUrl = sheet_obj.cell(row=i, column=2).value
    response = wget.download(downloadUrl, filename)
    print(i)

import openpyxl
from docxtpl import DocxTemplate 
import datetime

path = "S:\DATA ENTRY\data.xlsx"
workbook=openpyxl.load_workbook(path)
sheet=workbook.active

list_values=list(sheet.values)
print(list_values)
doc = DocxTemplate("reportcard.docx")

for i in range(1,len(list_values),4):
    doc.render({"fn" : list_values[i][1],
                "ln" : list_values[i][2],
                "std" : list_values[i][3],
                "nat":list_values[i][4],
                "trm":list_values[i][5],
                "s1":list_values[i][6],
                "s2":list_values[i+1][6],
                "s3":list_values[i+2][6],
                "m1":list_values[i][7],
                "m2":list_values[i+1][7],
                "m3":list_values[i+2][7],
                "mm1":list_values[i][8],
                "mm2":list_values[i+1][8],
                "mm3":list_values[i+2][8],
                "p1":list_values[i][9],
                "p2":list_values[i+1][9],
                "p3":list_values[i+2][9],
                "g1":list_values[i][10],
                "g2":list_values[i+1][10],
                "g3":list_values[i+2][10],
                "mt":list_values[i+3][7],
                "mmt":list_values[i+3][8],
                "pt":list_values[i+3][9],
                "gt":list_values[i+3][10],})
    doc_name="reportcard"+ list_values[i][1] + list_values[i][2] +".docx"
    doc.save(doc_name)
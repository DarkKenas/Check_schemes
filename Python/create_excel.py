import os
import openpyxl as px
from openpyxl.styles import Side, Border, Alignment
def ex_write(path, mass):
    path_shab = os.path.abspath(r"shabl.xlsx")
    wb = px.load_workbook(path_shab)
    ws = wb.active
    i=1
    for m in mass:
        ws.append([i,*m])
        i+=1
    for j in ws[ws.calculate_dimension()]:
        for k in j:
            k.alignment = Alignment(horizontal="center", vertical="center")
            medium = Side(border_style="thin", color ="000000")
            k.border = Border(top=medium, bottom=medium, left=medium, right=medium)
    wb.save(path)
    wb.close()

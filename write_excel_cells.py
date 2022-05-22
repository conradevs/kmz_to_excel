from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
from openpyxl.utils import get_column_letter
from openpyxl.styles import numbers, is_date_format
from openpyxl.styles.styleable import StyleableObject
from openpyxl.worksheet.hyperlink import Hyperlink

from helpers import letter

def excel_fill(ws,points,lines,shapes):
    index=1
    ws['A'+str(index)].value = 'Points'
    index=+1
    for i in range(0,len(points)-1):
        ws['A'+str(index+i)].value = points[i].name
        ws['B'+str(index+i)].value = points[i].coord_x
        ws['c'+str(index+i)].value = points[i].coord_y
        ws['D'+str(index+i)].value = points[i].coord_z
    index=index+len(points)
    ws['A'+str(index)].value = 'Lines'
    for i in range(0,len(lines)-1):
        ws['A'+str(index+i)].value = lines[i].name
        for j in range(1,len(lines[i].points)):
            ws[letter(j)+str(index+i)].value = lines[i].points[j]
            ws[letter(j+1)+str(index+i)].value = lines[i].points[j+1]
            ws[letter(j+2)+str(index+i)].value = lines[i].points[j+2]
            j=+3
            if j>=len(lines[i].points)-1: break
    index=index+len(lines)
    ws['A'+str(index)].value = 'Shapes'
    for i in range(0,len(shapes)-1):
        ws['A'+str(index+i)].value = shapes[i].name
        for j in range(1,len(lines[i].points)):
            ws[letter(j)+str(index+i)].value = shapes[i].points[j]
            ws[letter(j+1)+str(index+i)].value = shapes[i].points[j+1]
            ws[letter(j+2)+str(index+i)].value = shapes[i].points[j+2]
            j=+3
            if j>=len(lines[i].points)-1: break
            
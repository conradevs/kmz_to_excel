from zipfile import ZipFile

filename = 'test.kmz'

kmz = ZipFile(filename, 'r')
kml = kmz.open('doc.kml', 'r')

import xml.sax, xml.sax.handler
from kmzHandler import PlacemarkHandler
from Line import Point, Line
from write_excel_cells import excel_fill

parser = xml.sax.make_parser()
handler = PlacemarkHandler()
parser.setContentHandler(handler)
parser.parse(kml)
kmz.close()

def points_lines_shapes(mapping):
    sep = ','
       
    output = 'Name' + sep + 'Coordinates\n'
    #points = ''
    #lines = ''
    #shapes = '' 
    Points = []
    Lines = []
    Shapes = []
    for key in mapping:
        coord_str = mapping[key]['coordinates'] + sep
        coord_split = coord_str.split(',')
        if 'LookAt' in mapping[key]: #points
            #points += key + sep + coord_str + "\n"           
            Points.append(Point(key,coord_split[0],coord_split[1],coord_split[2]))
        elif 'LineString' in mapping[key]: #lines
            #lines += key + sep + coord_str + "\n"
            line_points = []
            for i in range(1,len(coord_split)-1):
                if i>len(coord_split)-3: break
                line_points.append(Point(key+'_'+str(i),coord_split[i],coord_split[i+1],coord_split[i+2]))
                i= i+2
            Lines.append(Line(key,line_points))
        else: #shapes
            #shapes += key + sep + coord_str + "\n"
            shape_points = []
            for i in range(1,len(coord_split)-1):
                shape_points.append(Point(key+'_'+str(i),coord_split[i],coord_split[i+1],coord_split[i+2]))
                i= i+2
            Shapes.append(Line(key,shape_points))
    output = [Points,Lines,Shapes]
    return output
 
output = points_lines_shapes(handler.mapping)

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
from openpyxl.utils import get_column_letter
from openpyxl.styles import numbers, is_date_format
from openpyxl.styles.styleable import StyleableObject
from openpyxl.worksheet.hyperlink import Hyperlink

wb = Workbook()
ws =  wb.active
ws.title = "COORDENADAS"
excel_fill(ws,output[0],output[1],output[2])
out_filename = filename[:-3] + "xlsx" #output filename same as input plus .csv
wb.save(filename = out_filename)



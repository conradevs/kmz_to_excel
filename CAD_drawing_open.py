import sys
sys.path.append('../')
from pyautocad import Autocad, APoint, aDouble
from comtypes import *
from comtypes.client import CreateObject
from comtypes import COMError
from openpyxl import Workbook
from helpers import letter
from Line import Point
from write_excel_cells import bdi_ws_fill

from convert_calculations import utmToLatLng

def select_text_inside_polygon(acad,points):
    print('polygon ')

def coordinates_from_Autocad(huso,is_north_hemisphere):
    acad = Autocad(create_if_not_exists=True)
    print('Working at... '+acad.doc.Name)

    set = acad.get_selection(text='Select some objects')
    coord_list = list()
    try:
        for obj in set:
            if(obj.ObjectName == 'AcDbPolyline'):
                print ('there is a Polyline in the selection')
            if(obj.ObjectName == 'AcDbLine'):
                    print ('there is a Line in the selection')
            if(obj.ObjectName == 'AcDbCircle'):
                    print ('there is a Circle in the selection')
                    position = obj.Center
                    print(position)
            if(obj.ObjectName == 'AcDbBlockReference'):
                position = obj.InsertionPoint
                coord = utmToLatLng(huso,position[0],position[1],is_north_hemisphere)
                coord_list.append(Point("coord",str(coord[0]),str(coord[1]),0))
        return(coord_list)
    except COMError as ce:
        target_error = ce.args # this is a tuple
        if target_error[1] == 'Call was rejected by callee.':
            print(target_error[1])
        return("Error extracting data: "+target_error[1])  

def save_in_excel_file(output_file_name,output_directory,coord_list):
    if(len(coord_list)==0): return
    # Create xlsx file
    wb = Workbook()
    #select active tab
    ws =  wb.active
    # name it "BDI_LMT"
    ws.title = "BDI_LMT"
    # call excel bdi format writing function
    bdi_ws_fill(ws,coord_list)
    # output xlsx file same name as kmz file
    out_filename = output_directory + "/" +output_file_name + ".xlsx"
    print(out_filename)
    # save output file
    wb.save(filename = out_filename)

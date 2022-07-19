import sys
from numpy import array
sys.path.append('../')
from pyautocad import ACAD, Autocad, APoint, aDouble
from comtypes import *
from comtypes.client import CreateObject
from comtypes import COMError
from openpyxl import Workbook
import array

acad = Autocad(create_if_not_exists=True)
doc = acad.ActiveDocument

print('Working at... '+acad.doc.Name)
#doc.Utility.Prompt("Hello AutoCAD\n")

vbCrLf = str
prompt1 = vbCrLf and 'Enter the start point'
point_1 = APoint(0.0,0.0,0.0)
point_2 = acad.ActiveDocument.Utility.GetPoint(point_1,prompt1)
print(point_2)
#point_1 = APoint(0.0,0.0,0.0)
ssetObj = acad.ActiveDocument.SelectionSets.Item("New_")
print(ssetObj.Name)
print(ssetObj.__dict__)
point_base = APoint(point_2)
polygon = list()
polygon.append(point_base + APoint(-20.0,20.0,0.0))
polygon.append(point_base + APoint(20.0,20.0,0.0))
polygon.append(point_base + APoint(20.0,-20.0,0.0))
polygon.append(point_base + APoint(-20.0,-20.0,0.0))
Points_list = list()
index = 0
for point in polygon:
    Points_list.append(point[0])
    Points_list.append(point[1])
    Points_list.append(point[2])
    index = index+3
print(polygon[0])
print(polygon[1])
print(polygon[2])
print(polygon[3])
for point in Points_list: print(str(point))
FilterType = array.array('h',[1])
FilterData = ['Text']
mode = ACAD.acSelectionSetCrossingPolygon
print(mode)
ssetObj.SelectByPolygon(mode,Points_list,FilterType,FilterData)
#for obj in ssetObj:
#    if(obj.ObjectName == 'AcDbText'):
#        print ('Text')
#ssetObj.Delete()

#coord_points=list()
#for obj in set:
#    if(obj.ObjectName == 'AcDbCircle'):
#            print ('there is a Circle in the selection')
#            position = APoint(obj.Center)
#            coord_points.append(position)
#            print(position)
#    if(obj.ObjectName == 'AcDbBlockReference'):
#        position = APoint(obj.InsertionPoint)
#        coord_points.append(position)
#vbCrLf = str
#prompt1 = vbCrLf and 'Enter the start point'
#point_1 = APoint(0.0,0.0,0.0)
#ssetObj = doc.SelectionSets.Add("New_SelectionSet")
#polygon = list()
#polygon.append(set[0] + APoint(-20.0,20.0,0.0))
#polygon.append(set[0] + APoint(20.0,20.0,0.0))
#polygon.append(set[0] + APoint(20.0,-20.0,0.0))
#polygon.append(set[0] + APoint(-20.0,-20.0,0.0))
#mode = acad.acSelectionSetFence
#ssetObj.SelectByPolygon(mode,polygon)


#acad.ActiveDocument.Utility.GetPoint(point_1,prompt1)
#coord_list = list()
#for obj in set:
#    if(obj.ObjectName == 'AcDbPolyline'):
#        print ('there is a Polyline in the selection')
#    if(obj.ObjectName == 'AcDbLine'):
#            print ('there is a Line in the selection')
#    if(obj.ObjectName == 'AcDbCircle'):
#            print ('there is a Circle in the selection')
#            position = obj.Center
#            coord_list.append(position)
#            print(position)
#    if(obj.ObjectName == 'AcDbBlockReference'):
#        position = APoint(obj.InsertionPoint)
#        coord_list.append(position)
#        print(position)
#    set2 = acad.get_selection(text='Select for name')
    
#for center in coord_list:
#    #selection = utility.AcadSelectionSets.Add
#    points = [center + APoint(-20.0,+20.0,0.0),center + APoint(+20.0,+20.0,0.0),center + APoint(+20.0,-20.0,0.0),center + APoint(-20.0,-20.0,0.0)]
#    print(points[0]+', '+points[1]+', '+points[2]+', '+points[3]+', '+points[4])
#    #selection.SelectByPolygon(points, ['text'])
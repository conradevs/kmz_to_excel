from pyautocad import Autocad, APoint, aDouble

acad = Autocad(create_if_not_exists=True)
# pR = aDouble(0,0,0,2,0,0,2,4,0,0,4,0,0,0,0)
# acad.model.AddPolyline(pR)
print('Working at... '+acad.doc.Name)
for obj in acad.iter_objects():
    print (obj.ObjectName)


# print('Texts and lines')
# for obj in acad.iter_objects(['Text', 'Line']):
#     print (obj.ObjectName)

set = acad.get_selection(text='Select some objects')
#for set in acad.doc.SelectionSets:
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
        caca=0
        position = obj.InsertionPoint
        print(position)
    
prompt1 = "Enter the start point of the line: "
#point_1 = APoint(0,0,1)
#point_2 = acad.doc.Utility.GetPoint(point_1,prompt1)

#print(point_2)
import array

import comtypes.client

app = comtypes.client.GetActiveObject("AutoCAD.Application")
ms = app.ActiveDocument
FilterType = array.array('h',list_of_dxf_group_codes)
SSobj = ms.SelectionSets.Add("SS")
point1 = [0,0,0]
point2 = [10, 10, 0]
mode = 2
SSobj.select(mode, point1, point2, 0)
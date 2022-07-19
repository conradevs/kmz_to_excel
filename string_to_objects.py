
from Line import Point, Line

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
        coord_str = mapping[key]['coordinates']
        coord_split = coord_str.split(',')
        if 'LookAt' in mapping[key] or 'Camera' in mapping[key]:  #points         
            Points.append(Point(key,coord_split[0],coord_split[1],coord_split[2]))
        elif 'LineString' in mapping[key]: #lines
            coord_split = coord_str.split(' ')
            line_points = []
            i=1
            for point in coord_split:
                point_split = point.split(',')
                if(i>=len(coord_split)): break
                line_points.append(Point(key+'_'+str(i),point_split[0],point_split[1],point_split[2]))
                i=i+1
            Lines.append(Line(key,line_points))
        else: #shapes
            coord_split = coord_str.split(' ')
            shape_points = []
            i=1
            for point in coord_split:
                point_split = point.split(',')
                #print(point)
                if(i>=len(coord_split)): break
                shape_points.append(Point(key+'_'+str(i),point_split[0],point_split[1],point_split[2]))
                i=i+1
            Shapes.append(Line(key,shape_points))
    output = [Points,Lines,Shapes]
    return output
class Point:
    def __init__(self,name,coord_x,coord_y,coord_z):
        self.name=name
        self.coord_x = coord_x
        self.coord_y = coord_y
        self.coord_z = coord_z
    
class Line:
    def __init__(self,name,points):
        self.name = name
        self.points = points

class Shape:
    def __init__(self,name,points):
        self.name = name
        self.points = points
import sys
import copy
from zipfile import ZipFile
import xml.sax, xml.sax.handler
from kmzHandler import PlacemarkHandler
from string_to_objects import points_lines_shapes
from openpyxl import load_workbook
from Line import Point

BDI_CONFIGURACIONES_INIT = 9

class BDI_Header:
    def __init__(self):
        self.station = ''
        self.section = ''
        self.SB_init = ''
        self.SB_end = ''
        self.work_code = ''

class Column:
    def __init__(self,name):
        self.name = name
    def load_configuration_from_template(self,index,ws):
        self.coordinates = ws['B'+str(index)].value
        self.conductors = ws['C'+str(index)].value
        self.line_material = ws['D'+str(index)].value
        self.section = ws['E'+str(index)].value
        self.cable_isulation = ws['F'+str(index)].value
        self.column_material = ws['G'+str(index)].value
        self.column_height = ws['H'+str(index)].value
        self.strength = ws['I'+str(index)].value
        self.column_role = ws['J'+str(index)].value
        self.length = ws['K'+str(index)].value
        self.configuration = ws['L'+str(index)].value
        self.cross_material = ws['M'+str(index)].value
        self.insulator_type = ws['N'+str(index)].value
        self.insulator_device = ws['O'+str(index)].value
        self.earth = ws['P'+str(index)].value
        self.state = ws['Q'+str(index)].value
        self.rein = ws['R'+str(index)].value
        self.instalation_date = ws['S'+str(index)].value
        self.bird_device = ws['T'+str(index)].value
        self.other_configuration = ws['U'+str(index)].value
        self.other_cross_material = ws['V'+str(index)].value
        self.other_insulator_type = ws['W'+str(index)].value
        self.other_insulator_device = ws['X'+str(index)].value
        self.contract_number = ws['Y'+str(index)].value
        self.observations = ws['Z'+str(index)].value
        self.elements = ws['AA'+str(index)].value

def BDI_configurations_load():
    wb = load_workbook(filename = 'BDI_configuraciones.xlsx')
    ws = wb.active
    row = BDI_CONFIGURACIONES_INIT
    output = []
    while (ws['A'+str(row)].value != None):
        #print(ws['A'+str(row)].value)
        output.append(Column(ws['A'+str(row)].value))
        output[-1].load_configuration_from_template(row,ws)
        row = row + 1
    wb.close
    return output

def column_class_constructor(point: Point,configurations):
    post_config = point.name.split(' ')
    post = None
    # Apoyo suspension por defecto
    if(len(post_config) == 1):
        post = copy.deepcopy(configurations[0]) # Cargo parametros desde template (poste C4)
        post.name = post_config[0]
        return post
    if(post_config[1]) == 'SU':
        post = configurations[0] # Cargo parametros desde template (poste C4)
        if(len(post_config)>2 and post_config[2] == 'col'):
            post.column_material = 'Columna de hormigón con orificios'
            post.strength = '3000 N (300 Kg)'
            post.cross_material = 'Metálica'
            post.earth = 'Si'
        if(len(post_config)>2 and post_config[2] == 'C5'): # Esfuerzo C5
            post.strength = 'Clase 5'
    # Apoyo suspension en angulo
    if(post_config[1] == 'SA' and post_config[2] == 'BA'):
        post = copy.deepcopy(configurations[1])
    if(post_config[1] == 'SA' and post_config[2] == 'RS'):
        post = copy.deepcopy(configurations[2])
    if(post_config[1] == 'SA' and post_config[2] == 'RD'):
        post = copy.deepcopy(configurations[3])
    if(post_config[1] == 'SD'):
        post = copy.deepcopy(configurations[4])
    # Apoyo amarre en angulo    
    if(post_config[1] == 'AA' and post_config[2] == 'BA'):
        post = copy.deepcopy(configurations[5])
    if(post_config[1] == 'AA' and post_config[2] == 'DE'):
        post = copy.deepcopy(configurations[6])
    # Apoyo amarre en angulo derivacion
    if(post_config[1] == 'AAD' and post_config[2] == 'BA'):
        post = copy.deepcopy(configurations[7])
    if(post_config[1] == 'AAD' and post_config[2] == 'DE'):
        post = copy.deepcopy(configurations[8])
    if(post_config[1] == 'SAD' and post_config[2] == 'BA'):
        post = copy.deepcopy(configurations[9])
    if(post_config[1] == 'SAD' and post_config[2] == 'RS'):
        post = copy.deepcopy(configurations[10])
    if(post_config[1] == 'SAD' and post_config[2] == 'RD'):
        post = copy.deepcopy(configurations[11])
    if(post_config[1] == 'TEA'):
        post = copy.deepcopy(configurations[12])
    if(post_config[1] == 'TE'):
        post = copy.deepcopy(configurations[13])
    if(post != None): post.name = post_config[0]
    return post

def BDI_write_row(point: Column, ws, row):
       ws['A'+str(row)].value = point.name
       ws['B'+str(row)].value = point.coordinates
       ws['C'+str(row)].value = point.conductors
       ws['D'+str(row)].value = point.line_material
       ws['E'+str(row)].value = point.section
       ws['F'+str(row)].value = point.cable_isulation
       ws['G'+str(row)].value = point.column_material
       ws['H'+str(row)].value = point.column_height
       ws['I'+str(row)].value = point.strength
       ws['J'+str(row)].value = point.column_role
       ws['K'+str(row)].value = point.length
       ws['L'+str(row)].value = point.configuration
       ws['M'+str(row)].value = point.cross_material
       ws['N'+str(row)].value = point.insulator_type
       ws['O'+str(row)].value = point.insulator_device
       ws['P'+str(row)].value = point.earth
       ws['Q'+str(row)].value = point.state
       ws['R'+str(row)].value = point.rein
       ws['S'+str(row)].value = point.instalation_date
       ws['T'+str(row)].value = point.bird_device
       ws['U'+str(row)].value = point.other_configuration
       ws['V'+str(row)].value = point.other_cross_material
       ws['W'+str(row)].value = point.other_insulator_type   
       ws['X'+str(row)].value = point.other_insulator_device
       ws['Y'+str(row)].value = point.contract_number
       ws['Z'+str(row)].value = point.observations
       ws['AA'+str(row)].value = point.elements

def test_function(file_input,directory_output):

    kmz = ZipFile(file_input, 'r')
    kml = kmz.open('doc.kml', 'r')

    parser = xml.sax.make_parser()
    handler = PlacemarkHandler()
    parser.setContentHandler(handler)
    parser.parse(kml)
    kmz.close()

    points = points_lines_shapes(handler.mapping)[0]

    config = BDI_configurations_load()
    wb = load_workbook(filename = 'BDI_template.xlsx')
    ws = wb.active
    postes = []
    for point in points:
        post = column_class_constructor(point,config)
        if(post != None):
            post.coordinates = "S" + point.coord_y[1:10] + "° " + "W"+point.coord_x[1:10] + "°"
            postes.append(post)
    for row in range(BDI_CONFIGURACIONES_INIT,len(postes)+BDI_CONFIGURACIONES_INIT):
        BDI_write_row(postes[row-BDI_CONFIGURACIONES_INIT], ws, row)

    output = points_lines_shapes(handler.mapping)
    wb.save(filename = directory_output+'/BDI_output.xlsx')
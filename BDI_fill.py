from openpyxl import load_workbook
from Line import Point

BDI_CONFIGURACIONES_INIT = 9

class Column:
    def __init__(self,name):
        self.name = name
    def fill_from_BDI_configuraciones(self,index,ws):
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

def BDI_configurations_read():
    wb = load_workbook(filename = 'BDI_configuraciones.xlsx')
    ws = wb.active
    row = BDI_CONFIGURACIONES_INIT
    output = []
    while (ws['A'+str(row)].value != None):
        #print(ws['A'+str(row)].value)
        output.append(Column(ws['A'+str(row)].value))
        output[-1].fill_from_BDI_configuraciones(row,ws)
        row = row + 1
    wb.close
    return output

def column_class_constructor(point: Point,configurations):
    config_code = point.name.split('_')
    if(config_code[0]) == 'SU':
        post = configurations[0]
        if(len(config_code)>1 and config_code[1] == 'col'):
            post.column_material = 'Columna de hormigón con orificios'
        if(len(config_code)>2 and config_code[2] == 'C5'):
            post.strength = 'Clase 5'

def test_function():
    config = BDI_configurations_read()
    print(config[0].name)
    print(config[0].coordinates)
    print(config[0].conductors)
    print(config[0].line_material)
    print(config[0].section)
    print(config[0].cable_isulation)
    wb = load_workbook(filename = 'BDI_template.xlsx')
    ws = wb.active
    print('From BDI button')
    ws['A'+str(BDI_CONFIGURACIONES_INIT)].value = config[0].name
    ws['B'+str(BDI_CONFIGURACIONES_INIT)].value = config[0].coordinates
    ws['C'+str(BDI_CONFIGURACIONES_INIT)].value = config[0].conductors
    ws['D'+str(BDI_CONFIGURACIONES_INIT)].value = config[0].line_material
    ws['E'+str(BDI_CONFIGURACIONES_INIT)].value = config[0].section
    ws['F'+str(BDI_CONFIGURACIONES_INIT)].value = config[0].cable_isulation
    ws['G'+str(BDI_CONFIGURACIONES_INIT)].value = config[0].column_material
    
    wb.save(filename = 'BDI_output_test.xlsx')

from tkinter import *
from tkinter import filedialog as fd
from write_excel_cells import convert_file, bdi_ws_fill
from CAD_drawing_open import coordinates_from_Autocad, save_in_excel_file
from BDI_fill import test_function, BDI_Header
#main window
root =Tk()
root.title('kmz a excel')
root.iconbitmap('Logo.ico')
#root.geometry("400x250")
main_labelframe = LabelFrame(root,text='Extrae coordenadas de un .kmz a un libro de excel',labelanchor='n',padx=10,pady=10)
main_labelframe.pack(ipadx=10,ipady=10,expand=True)

# choose kmz file for convertion
def pick_kmz_file(label_show: Label):
    kmz_file = fd.askopenfile(filetypes = (('kmz files', '*.kmz'),('kml', '*.kml')))
    kmz_file = kmz_file.name
    print(kmz_file)
    label_show.config(text = kmz_file)

kmz_file_select_container = LabelFrame(main_labelframe,text="Elije archivo a convertir",pady=5,width=120)
label_file_select = Label(kmz_file_select_container, width=50, text = "Elegir archivo kmz...")
button_file_select = Button(kmz_file_select_container,width=15,text="Buscar", command=lambda:pick_kmz_file(label_file_select))

kmz_file_select_container.pack(expand=True)
label_file_select.pack()
button_file_select.pack()

# path for new directory
def pick_directory(label_show: Label):
    path = fd.askdirectory()
    label_show.config(text = path)

path_select_container = LabelFrame(main_labelframe,text="Directorio donde guardar archivo excel",pady=5,width=120)
label_path_select = Label(path_select_container, width=50, text = "Elegir directorio...")
button_path_select = Button(path_select_container,width=15,text="Buscar", command=lambda:pick_directory(label_path_select))

path_select_container.pack()
label_path_select.pack()
button_path_select.pack()

print(label_file_select.cget('text'))
print(label_path_select.cget('text'))
label_pick_output = Label(main_labelframe,text='Elige como guardar resultados')
excel_select_container = LabelFrame(main_labelframe,text='En planilla excel',pady=5,padx=5,width=120)
excel_select_container.pack()
button_convert = Button(excel_select_container,text="Convertir a planilla excel",command=lambda:convert_file(label_file_select,label_path_select))
button_convert.grid(row=0,column=0)
BDI_select_container = LabelFrame(main_labelframe,text='En planilla BDI',pady=2,padx=2)
BDI_select_container.pack()
# General BDI info
    # select BDI template    
def select_template_file(header_info: BDI_Header, show_selection: Entry):
    excel_template = fd.askopenfile(filetypes=[('Excel Files', ['xlsx'])])
    excel_template = excel_template.name
    print(excel_template)
    show_selection.delete(0, 'end')
    show_selection.insert(0,excel_template)
    header_info.load_from_template(header_info,excel_template)
    header_info.fill_form(header_info,carpeta_get,estacion_get,salida_get,tension_get,SB_1_get,SB_2_get)

    # button find file
template_select = Entry(BDI_select_container)
template_select.insert(0,'Cargar encabezado de BDI desde un template...')
template_select.grid(row=0,columnspan=3,sticky = 'ew')
header = BDI_Header
button_template_select = Button(BDI_select_container,width=15,text="Buscar", command=lambda:select_template_file(header,template_select))
button_template_select.grid(row=0,column=3)
    
    # fill form by hand
carpeta_label = Label(BDI_select_container,text= 'Nro de carpeta :')
carpeta_label.grid(row=1,column=0)
carpeta_get = Entry(BDI_select_container)
carpeta_get.grid(row=1,column=1,columnspan=2,sticky = 'ew')
    # ESTACION
estacion_label = Label(BDI_select_container,text= 'Estacion :')
estacion_label.grid(row=2,column=0)
estacion_get = Entry(BDI_select_container)
estacion_get.grid(row=2,column=1)

    # SALIDA
salida_label = Label(BDI_select_container,text= 'Salida :')
salida_label.grid(row=2,column=2)
salida_get = Entry(BDI_select_container)
salida_get.grid(row=2,column=3)

    # Tension
tension_label = Label(BDI_select_container,text= 'Tensi??n = ')
tension_label.grid(row=3,column=0)
tension_get = Entry(BDI_select_container)
tension_get.grid(row=3,column=1)
tension_V = Label(BDI_select_container,text= ' (V)')
tension_V.grid(row=3,column=2,sticky = 'w')

    # DESDE SB
SB_1_label = Label(BDI_select_container,text= 'Desde SB :')
SB_1_label.grid(row=4,column=0)
SB_1_get = Entry(BDI_select_container)
SB_1_get.grid(row=4,column=1)

    # HASTA SB
SB_2_label = Label(BDI_select_container,text= 'Hasta SB :')
SB_2_label.grid(row=4,column=2)
SB_2_get = Entry(BDI_select_container)
SB_2_get.grid(row=4,column=3)


button_to_BDI_file = Button(BDI_select_container,
    text="Convertir a planilla BDI",
    command=lambda: test_function(header,
        label_file_select.cget('text'),
        label_path_select.cget('text'),
        carpeta_get,
        estacion_get,
        salida_get,
        tension_get,
        SB_1_get,
        SB_2_get
    ))
button_to_BDI_file.grid(row=5,column=2)

#def on_click_button_acad(excel_book_path):
#    if excel_book_path == 'Elegir directorio...': 
#        print('Elige directorio donde guardar resultados')
#        return
#    coord_list = coordinates_from_Autocad(21,False)
#    save_in_excel_file("ACAD_Output",excel_book_path,coord_list)


#button_convert_from_ACAD = Button(main_labelframe,text="Extraer desde ACAD",command=lambda:on_click_button_acad(label_path_select.cget('text')))
#button_convert_from_ACAD.pack(expand=True)
root.mainloop()


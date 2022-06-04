from tkinter import *
from tkinter import filedialog as fd
from write_excel_cells import convert_file

#main window
root =Tk()
root.title('kmz a excel')
root.iconbitmap('Logo.ico')
root.geometry("400x250")
main_labelframe = LabelFrame(root,text='Extrae coordenadas de un .kmz a un libro de excel',labelanchor='n',padx=10,pady=10)
main_labelframe.pack(ipadx=10,ipady=10,expand=True)

# choose kmz file for convertion
def pick_kmz_file(label_show: Label):
    kmz_file = fd.askopenfile(filetypes = (('kmz files', '*.kmz'),('kml', '*.kml')))
    kmz_file = kmz_file.name
    print(kmz_file)
    label_show.config(text = kmz_file)

kmz_file_picker_container = LabelFrame(main_labelframe,text="Elije archivo a convertir",pady=5,width=120)
label_file_picker = Label(kmz_file_picker_container, width=50, text = "Elegir archivo kmz...")
button_file_picker = Button(kmz_file_picker_container,width=15,text="Buscar", command=lambda:pick_kmz_file(label_file_picker))

kmz_file_picker_container.pack(expand=True)
label_file_picker.pack()
button_file_picker.pack()

# path for new directory
def pick_directory(label_show: Label):
    path = fd.askdirectory()
    label_show.config(text = path)

path_picker_container = LabelFrame(main_labelframe,text="Directorio donde guardar archivo excel",pady=5,width=120)
label_path_picker = Label(path_picker_container, width=50, text = "Elegir directorio...")
button_path_picker = Button(path_picker_container,width=15,text="Buscar", command=lambda:pick_directory(label_path_picker))

path_picker_container.pack()
label_path_picker.pack()
button_path_picker.pack()

print(label_file_picker.cget('text'))
print(label_path_picker.cget('text'))

button_convert = Button(main_labelframe,text="Convertir",command=lambda:convert_file(label_file_picker,label_path_picker))
button_convert.pack(expand=True)
button_convert = Button(main_labelframe,text="Convertir",command=lambda:coordinates_from_Autocad())
root.mainloop()





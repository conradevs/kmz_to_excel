from tkinter import *
from tkinter import filedialog as fd
from write_excel_cells import convert_file

#main window
root =Tk()
root.title('kmz a excel')
root.iconbitmap('Logo.ico')
root.geometry("600x500")

# choose kmz file for convertion
def pick_kmz_file(label_show: Label):
    kmz_file = fd.askopenfile(filetypes = (('kmz files', '*.kmz'),('kml', '*.kml')))
    kmz_file = kmz_file.name
    print(kmz_file)
    label_show.config(text = kmz_file)

kmz_file_picker_container = LabelFrame(root,text="Elije archivo a convertir")
label_file_picker = Label(kmz_file_picker_container, width=30, text = "Elegir archivo kmz...")
button_file_picker = Button(kmz_file_picker_container,text="Buscar", command=lambda:pick_kmz_file(label_file_picker))

kmz_file_picker_container.pack()
label_file_picker.pack()
button_file_picker.pack()

# path for new directory
def pick_directory(label_show: Label):
    path = fd.askdirectory()
    label_show.config(text = path)

path_picker_container = LabelFrame(root,text="Directorio donde guardar archivo excel")
label_path_picker = Label(path_picker_container, width=30, text = "Elegir directorio...")
button_path_picker = Button(path_picker_container,text="Buscar", command=lambda:pick_directory(label_path_picker))

path_picker_container.pack()
label_path_picker.pack()
button_path_picker.pack()

print(label_file_picker.cget('text'))
print(label_path_picker.cget('text'))

button_convert = Button(root,text="Convertir",command=lambda:convert_file(label_file_picker,label_path_picker))
button_convert.pack()
root.mainloop()





# -*- coding: utf-8 -*-
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import os
import pandas as pd
import hashlib
import codecs
from pdfrw import PdfReader, PdfWriter

global window
window = Tk()
window.title("Planeamiento")
window.geometry('300x300')
def clicked():
    folder_selected = filedialog.askdirectory()
    global path
    path = folder_selected
    #return path

def listaExcel ():
    ListaArchivos=os.listdir(path)
    ListaPdf = []
    ListaNombres = []
    ListaCsv = []
    ListaHash = []
    ListaHash64 = []
    for x in ListaArchivos:
        if x.endswith(".pdf"):
            ListaNombres.append(x)
            ListaPdf.append(path + "\\" + x)
    for pdf in ListaPdf:
        ListaCsv.append(PdfReader(pdf).Info.CSV[1:-1])
        with open(pdf,"rb") as f:
            bytes = f.read()
            ReadableHash = hashlib.sha256(bytes).hexdigest()
            ReadableHash64 = codecs.encode(codecs.decode(ReadableHash,'hex'),'base64').decode()
            ListaHash.append(ReadableHash)
            ListaHash64.append(ReadableHash64[0:-2])
    diccionario = {"NombreFichero":ListaNombres,  "Hash64": ListaHash64, "CSV": ListaCsv}
    df1 = pd.DataFrame(diccionario)
    df1.to_excel(path + "\\"+"ListaCSV.xlsx")
    messagebox.showinfo('Exportación', 'Lista Exportada')
#Texto introducción
lbl = Label(window, text=" 1. Selecciona la carpeta donde \nse encuentran los pdfs.",anchor='w',justify=LEFT )
lbl.grid(column=0, row=0,sticky=W)
btn = Button(window, text="Carpeta",command=clicked)
btn.grid(column=0, row=1)
lbl2 = Label(window, text=" \n2. Pulsa para exportar la lista \nen xls.",anchor='w',justify=LEFT )
lbl2.grid(column=0, row=2,sticky=W)
btn2 = Button(window, text="Exportar",command=listaExcel)
btn2.grid(column=0, row=3)
window.mainloop()
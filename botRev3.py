import os
import fitz  # PyMuPDF
import pandas as pd

import tkinter as tk
from tkinter import StringVar, messagebox
from tkinter import filedialog
#archivo_excel = "C:/Users/Asus/Documents/Proyectos/BotRevisorDocs/Docs/Listado.xlsx"
#carpeta_pdf = "C:/Users/Asus/Documents/Proyectos/BotRevisorDocs/Docs/certificados"
archivo_excel = "" 
carpeta_pdf = ""

def definirExcel():
    global archivo_excel 
    archivo_excel = filedialog.askopenfilename()
    texto = StringVar()
    texto.set(archivo_excel)
    ele.config(textvariable=texto)
def definirCarpeta():
    global carpeta_pdf 
    carpeta_pdf = filedialog.askdirectory()
    texto = StringVar()
    texto.set(carpeta_pdf+"/")
    ecarp.config(textvariable=texto)

def revisarCertificados():
    try:

        # Cargar el archivo Excel
        df = pd.read_excel(archivo_excel)
        nombre_columna = "Resultado"
        # Directorio que contiene los archivos PDF
        #carpeta_pdf = "Docs/Certificados"

        # Columna a evaluar
        c1 = "Apellido1"
        c2 = "Apellido2"
        c3 = "Nombre1"
        c4 = "Nombre2"

        # Recorrer todas las filas del DataFrame
        for index, row in df.iterrows():
            # Evaluar el valor de la celda
            #if isinstance(row[c1], str) and "texto que buscas" in row[c1]:
            #   print(f"Texto encontrado en la fila {index + 1}")
                
            # Texto a buscar
            a1 = row[c1];
            a2 = row[c2];
            n1 = row[c3];
            n2 = row[c4];

            # Recorrer todos los archivos en el directorio
            founded = 0
            for nombre_archivo in os.listdir(carpeta_pdf):
                if nombre_archivo.endswith(".pdf") and founded == 0:
                    ruta_archivo = os.path.join(carpeta_pdf,nombre_archivo)

                    # Abrir el archivo PDF
                    documento = fitz.open(ruta_archivo)

                    # Recorrer todas las páginas del documento
                    pagina = documento[0]
                    texto = pagina.get_text()

                    # Verificar si el texto está en la página
                    
                    if a1.lower() in texto.lower():
                        if a2.lower() in texto.lower():
                            if n1.lower() in texto.lower():
                                if n2.lower() in texto.lower():
                                    df.loc[index, nombre_columna] = "Validado nombres completos"
                                    founded = 1
                                else:
                                    df.loc[index, nombre_columna] = "Segundo nombre errado"
                            else:
                                df.loc[index, nombre_columna] = "Primer nombre errado"
                        else:
                            df.loc[index, nombre_columna] = "Segundo apellido errado"
                    else:
                        df.loc[index, nombre_columna] = "Primer apellido errado"
                    df.to_excel(archivo_excel, index=False)
                    documento.close()
        texto = StringVar()
        texto.set("Verificación finalizada correctamente")
        errm.config(textvariable=texto)
    except OSError as error:
        print(error)
        texto = StringVar()
        texto.set("Cierre los archivos, excel y certificados")
        errm.config(textvariable=texto)
    #finaliza manejo excel

#fin función revisar Certificados

# Crear la ventana principal
root = tk.Tk()
root.title("Formulario de Datos")

# Crear y colocar las etiquetas y campos de texto
le = tk.Label(root, text="Listado Excel:")
le.grid(row=0, column=0, padx=10, pady=5)

ele = tk.Label(root, text="")
ele.grid(row=0, column=1, padx=10, pady=5)

# Crear y colocar el botón de guardar
btnle = tk.Button(root, text="Listado Excel", command=definirExcel)
btnle.grid(row=0, column=2, columnspan=2, pady=5)

carp = tk.Label(root, text="Carpeta:")
carp.grid(row=1, column=0, padx=10, pady=5)

ecarp = tk.Label(root, text="")
ecarp.grid(row=1, column=1, padx=10, pady=5)

errm = tk.Label(root, text="")
errm.grid(row=2, column=0, padx=10, pady=5)

# Crear y colocar el botón de guardar
btncarp = tk.Button(root, text="Carpeta con Certificados", command=definirCarpeta)
btncarp.grid(row=1, column=2, columnspan=2, pady=5)

# Crear y colocar el botón de guardar
boton_guardar = tk.Button(root, text="Revisar", command=revisarCertificados)
boton_guardar.grid(row=2, column=1, columnspan=2, pady=10)
# Ejecutar el bucle principal de la aplicación
root.mainloop()



import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import zipfile
import os
import re
from PyPDF2 import PdfMerger
import unicodedata

# Función para seleccionar un archivo ZIP
def seleccionar_zip():
    archivo_zip = filedialog.askopenfilename(title="Seleccionar archivo ZIP", filetypes=[("ZIP files", "*.zip")])
    if archivo_zip:
        lista_archivos.delete(0, tk.END)  # Limpiar lista
        archivos_pdf = extraer_pdfs_de_zip(archivo_zip)
        if archivos_pdf:
            # Ordenar los archivos por apellido extraído del nombre
            archivos_pdf_ordenados = ordenar_por_apellido(archivos_pdf)
            for archivo in archivos_pdf_ordenados:
                lista_archivos.insert(tk.END, archivo)
        else:
            messagebox.showwarning("Advertencia", "No se encontraron archivos PDF en el ZIP.")

# Función para extraer los archivos PDF del ZIP
def extraer_pdfs_de_zip(archivo_zip):
    archivos_pdf = []
    # Crear una carpeta temporal para extraer los archivos PDF
    with zipfile.ZipFile(archivo_zip, 'r') as zip_ref:
        # Extraer solo los archivos PDF
        archivos = zip_ref.namelist()
        for archivo in archivos:
            if archivo.lower().endswith('.pdf'):
                # Extraer el archivo PDF
                zip_ref.extract(archivo, "temp_pdf_folder")
                archivos_pdf.append(os.path.join("temp_pdf_folder", archivo))
    return archivos_pdf

# Función para normalizar y comparar los apellidos
def normalizar_apellido(apellido):
    # Eliminar acentos y convertir a minúsculas
    return unicodedata.normalize('NFD', apellido).encode('ascii', 'ignore').decode('ascii').lower()

# Función para ordenar los archivos por apellido
def ordenar_por_apellido(archivos):
    def extraer_apellido(archivo):
        match = re.search(r'_MENSUAL_(.*?)(?:_\d+)?\.\w+', archivo)
        if match:
            apellidos = match.group(1).split('_')
            # Acceder al elemento del medio (índice 1)
            apellido = apellidos[1]
            return apellido
        return ""

    # Ordenar los archivos alfabéticamente por el apellido extraído
    def comparar_apellidos(archivo):
        apellido = extraer_apellido(archivo)
        # Eliminar todos los caracteres no alfanuméricos y convertir a minúsculas
        apellido_normalizado = re.sub(r'[^a-zA-Z0-9]', '', apellido).lower()
        return apellido_normalizado

    archivos_ordenados = sorted(archivos, key=comparar_apellidos)
    return archivos_ordenados

# Función para unificar los PDFs seleccionados
def unificar_pdfs():
    archivos = lista_archivos.get(0, tk.END)
    if archivos:
        merger = PdfMerger()
        # Primero ordenamos los archivos por apellido
        archivos_ordenados = ordenar_por_apellido(list(archivos))
        
        # Agregar los archivos PDF ordenados
        for archivo in archivos_ordenados:
            merger.append(archivo)
        
        archivo_salida = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if archivo_salida:
            merger.write(archivo_salida)
            merger.close()
            messagebox.showinfo("Éxito", "Los archivos PDF han sido unificados exitosamente.")
            limpiar_temporary_files()  # Limpiar los archivos temporales
    else:
        messagebox.showwarning("Advertencia", "No se han seleccionado archivos PDF.")


# Limpiar los archivos temporales extraídos
def limpiar_temporary_files():
    folder = "temp_pdf_folder"
    if os.path.exists(folder):
        for archivo in os.listdir(folder):
            archivo_path = os.path.join(folder, archivo)
            if os.path.isfile(archivo_path):
                os.remove(archivo_path)
        os.rmdir(folder)

# Crear la ventana principal
ventana = tk.Tk()
ventana.title("Unificador de PDFs desde ZIP")
ventana.geometry("800x400")

# Estilo para los widgets
style = ttk.Style()
style.configure("TButton",
                font=("Helvetica", 12),
                padding=10,
                relief="flat",
                background="#4CAF50",
                foreground="black")
style.map("TButton",
          background=[("active", "black"), ("pressed", "#388e3c")],
          foreground=[("pressed", "white"), ("active", "white")])

# Crear la lista donde se mostrarán los archivos PDF seleccionados
lista_archivos = tk.Listbox(ventana, width=80, height=10, font=("Helvetica", 12), bg="#f0f0f0", selectmode=tk.SINGLE, fg="black")
lista_archivos.pack(pady=20)

# Botón para seleccionar el archivo ZIP
boton_seleccionar_zip = ttk.Button(ventana, text="Seleccionar archivo ZIP", command=seleccionar_zip)
boton_seleccionar_zip.pack(pady=10)

# Botón para unificar los PDFs
boton_unificar = ttk.Button(ventana, text="Unificar PDFs", command=unificar_pdfs)
boton_unificar.pack(pady=10)

# Iniciar la interfaz gráfica
ventana.mainloop()

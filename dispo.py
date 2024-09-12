import pandas as pd
from tkinter import Tk, Label, Button, filedialog, StringVar, OptionMenu, messagebox
from docx import Document
import os

# Función para cargar y agrupar el archivo CSV
def cargar_datos(file_path):
    df = pd.read_csv(file_path, sep=';')
    df_grouped = df.groupby('ID').first()
    return df_grouped

# Función para obtener la lista de coordinadores, MD asistenciales o investigadores
def obtener_coordinadores(df_grouped, columna):
    coordinadores = df_grouped[columna].dropna().unique()
    return coordinadores

# Función para filtrar estudios por el coordinador, MD asistencial o investigador
def estudios_por_coordinador(df_grouped, coordinador_seleccionado, columna):
    estudios_filtrados = df_grouped[df_grouped[columna] == coordinador_seleccionado]
    
    tabla_estudios = pd.DataFrame({
        'Acrónimo': estudios_filtrados['Acrónimo Estudio'],
        'Número del Comité': estudios_filtrados['Número IRB'],
        'Fase del Estudio': estudios_filtrados.apply(
            lambda row: f"Estado: {'Activo' if '1. Activo' in str(row['Estado general del estudio']) else 'Inactivo'}, "
                        f"Reclutamiento: {'Si' if '1. Si' in str(row['Reclutamiento activo']) else 'No'}", axis=1),
        'Sujetos Tamizados': estudios_filtrados['Total fallas de tamizaje'],
        'Sujetos Activos': estudios_filtrados['Total de activos']
    }).reset_index(drop=True)
    
    return tabla_estudios

# Función para generar el documento en Word
def generar_documento(nombre_persona, categoria, tabla_estudios):
    doc = Document()
    doc.add_heading(f"Reporte para {categoria}", 0)
    doc.add_paragraph(f"Nombre: {nombre_persona}")
    
    if not tabla_estudios.empty:
        doc.add_paragraph("Tabla de estudios asociados:")
        table = doc.add_table(rows=1, cols=len(tabla_estudios.columns))
        hdr_cells = table.rows[0].cells
        for i, col_name in enumerate(tabla_estudios.columns):
            hdr_cells[i].text = col_name
        
        for index, row in tabla_estudios.iterrows():
            row_cells = table.add_row().cells
            for i, value in enumerate(row):
                row_cells[i].text = str(value)
    else:
        doc.add_paragraph("No hay estudios asociados.")
    
    filename = f"Reporte_{nombre_persona.replace(' ', '_')}.docx"
    doc.save(filename)
    messagebox.showinfo("Éxito", f"Documento guardado como {filename}")

# Función para actualizar la lista de personas según la categoría seleccionada
def actualizar_personas():
    if not os.path.exists(file_path.get()):
        messagebox.showerror("Error", "Por favor, selecciona un archivo CSV válido.")
        return

    df_grouped = cargar_datos(file_path.get())

    if seleccion_categoria.get() == "Seleccionar":
        messagebox.showerror("Error", "Por favor selecciona una categoría.")
        return

    categoria_opciones = {
        "Coordinador Principal": 'Coordinador Principal',
        "Coordinador backup principal 1": 'Coordinador backup principal 1',
        "Coordinador backup principal 2": 'Coordinador backup principal 2',
        "Coordinador backup principal 3": 'Coordinador backup principal 3',
        "Coordinador backup principal 4": 'Coordinador backup principal 4',
        "Coordinador backup principal 5": 'Coordinador backup principal 5',
        "MD asistencial 1": 'MD asistencial 1',
        "MD asistencial 2": 'MD asistencial 2',
        "MD asistencial 3": 'MD asistencial 3',
        "MD asistencial 4": 'MD asistencial 4',
        "MD asistencial 5": 'MD asistencial 5',
        "MD asistencial 6": 'MD asistencial 6',
        "MD asistencial 7": 'MD asistencial 7',
        "MD asistencial 8": 'MD asistencial 8',
        "Investigador Principal": 'Investigador Principal',
        "Co-Investigador 1": 'Co-Investigador 1',
        "Co-Investigador 2": 'Co-Investigador 2',
        "Co-Investigador 3": 'Co-Investigador 3',
        "Co-Investigador 4": 'Co-Investigador 4',
        "Co-Investigador 5": 'Co-Investigador 5',
        "Co-Investigador 6": 'Co-Investigador 6',
        "Co-Investigador 7": 'Co-Investigador 7'
    }
    
    categoria = categoria_opciones.get(seleccion_categoria.get())
    coordinadores = obtener_coordinadores(df_grouped, categoria)

    # Corrección del error - Verificar si la lista está vacía con len()
    if len(coordinadores) == 0:
        messagebox.showerror("Error", f"No hay personal registrado en la categoría {seleccion_categoria.get()}.")
        return
    
    # Actualizamos la lista de personas en el OptionMenu
    seleccion_persona.set("Seleccionar Persona")
    persona_menu['menu'].delete(0, 'end')  # Limpiar el menú actual
    for nombre in coordinadores:
        persona_menu['menu'].add_command(label=nombre, command=lambda value=nombre: seleccion_persona.set(value))

# Función para generar el informe
def generar_reporte():
    if seleccion_persona.get() == "Seleccionar Persona":
        messagebox.showerror("Error", "Por favor selecciona una persona.")
        return

    df_grouped = cargar_datos(file_path.get())
    
    categoria_opciones = {
        "Coordinador Principal": 'Coordinador Principal',
        "Coordinador backup principal 1": 'Coordinador backup principal 1',
        "Coordinador backup principal 2": 'Coordinador backup principal 2',
        "Coordinador backup principal 3": 'Coordinador backup principal 3',
        "Coordinador backup principal 4": 'Coordinador backup principal 4',
        "Coordinador backup principal 5": 'Coordinador backup principal 5',
        "MD asistencial 1": 'MD asistencial 1',
        "MD asistencial 2": 'MD asistencial 2',
        "MD asistencial 3": 'MD asistencial 3',
        "MD asistencial 4": 'MD asistencial 4',
        "MD asistencial 5": 'MD asistencial 5',
        "MD asistencial 6": 'MD asistencial 6',
        "MD asistencial 7": 'MD asistencial 7',
        "MD asistencial 8": 'MD asistencial 8',
        "Investigador Principal": 'Investigador Principal',
        "Co-Investigador 1": 'Co-Investigador 1',
        "Co-Investigador 2": 'Co-Investigador 2',
        "Co-Investigador 3": 'Co-Investigador 3',
        "Co-Investigador 4": 'Co-Investigador 4',
        "Co-Investigador 5": 'Co-Investigador 5',
        "Co-Investigador 6": 'Co-Investigador 6',
        "Co-Investigador 7": 'Co-Investigador 7'
    }
    
    categoria = categoria_opciones.get(seleccion_categoria.get())
    tabla_resultante = estudios_por_coordinador(df_grouped, seleccion_persona.get(), categoria)

    # Generar el documento Word
    generar_documento(seleccion_persona.get(), categoria, tabla_resultante)

# Función para abrir un cuadro de diálogo para seleccionar archivo CSV
def seleccionar_archivo():
    archivo = filedialog.askopenfilename(title="Seleccionar archivo", filetypes=(("CSV files", "*.csv"),))
    file_path.set(archivo)
    if archivo:
        messagebox.showinfo("Archivo seleccionado", f"Archivo seleccionado: {archivo}")

# Configuración de la ventana principal con Tkinter
root = Tk()
root.title("Generador de informes")
root.geometry("400x500")

# Variable para el archivo CSV seleccionado
file_path = StringVar()

# Variable para la selección de categoría
seleccion_categoria = StringVar(value="Seleccionar")

# Variable para la selección de persona
seleccion_persona = StringVar(value="Seleccionar Persona")

# Título
titulo = Label(root, text="Generador de Informes CIC", font=("Arial", 16))
titulo.pack(pady=10)

# Botón para cargar el archivo CSV
boton_cargar = Button(root, text="Cargar archivo CSV", command=seleccionar_archivo)
boton_cargar.pack(pady=10)

# Etiqueta para mostrar el archivo seleccionado
archivo_label = Label(root, textvariable=file_path, wraplength=300)
archivo_label.pack(pady=5)

# Cuadro de selección de categoría
# Cuadro de selección de categoría
categorias = ["Seleccionar", "Coordinador Principal", "Coordinador backup principal 1", "Coordinador backup principal 2",
              "Coordinador backup principal 3", "Coordinador backup principal 4", "Coordinador backup principal 5", 
              "MD asistencial 1", "MD asistencial 2", "MD asistencial 3", "MD asistencial 4", "MD asistencial 5", 
              "MD asistencial 6", "MD asistencial 7", "MD asistencial 8", "Investigador Principal", 
              "Co-Investigador 1", "Co-Investigador 2", "Co-Investigador 3", "Co-Investigador 4", "Co-Investigador 5",
              "Co-Investigador 6", "Co-Investigador 7"]

categoria_menu = OptionMenu(root, seleccion_categoria, *categorias, command=lambda _: actualizar_personas())
categoria_menu.pack(pady=10)

# Cuadro de selección de persona
persona_menu = OptionMenu(root, seleccion_persona, "Seleccionar Persona")
persona_menu.pack(pady=10)

# Botón para generar el informe
boton_generar = Button(root, text="Generar Informe", command=generar_reporte)
boton_generar.pack(pady=20)

# Ejecutar la aplicación
root.mainloop()

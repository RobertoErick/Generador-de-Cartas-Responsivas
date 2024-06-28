import tkinter as tk
from tkinter import filedialog, messagebox
from docx import Document
import os

# Función para reemplazar texto en los marcadores
def reemplazar_marcadores(doc, respuestas):
    for marcador, respuesta in respuestas.items():
        for p in doc.paragraphs:
            if f'{{{{{marcador}}}}}' in p.text:
                p.text = p.text.replace(f'{{{{{marcador}}}}}', respuesta)

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        if f'{{{{{marcador}}}}}' in p.text:
                            p.text = p.text.replace(f'{{{{{marcador}}}}}', respuesta)

# Función para crear los directorios si no existen
def crear_directorios(base_path):
    responsivas_dir = os.path.join(base_path, 'Cartas Responsivas')
    recepcion_dir = os.path.join(base_path, 'Cartas Recepción')
    
    os.makedirs(responsivas_dir, exist_ok=True)
    os.makedirs(recepcion_dir, exist_ok=True)
    
    return responsivas_dir, recepcion_dir

# Función para crear el documento basado en una plantilla
def crear_carta_responsiva_desde_plantilla(plantilla_path, respuestas, output_dir):
    # Cargar la plantilla
    doc = Document(plantilla_path)

    # Reemplazar los marcadores con las respuestas
    reemplazar_marcadores(doc, respuestas)

    # Formatear el nombre del archivo
    nombre_archivo = f"{respuestas['Nombre']}_{respuestas['Puesto']}_{respuestas['Equipo']}_Carta_Responsiva.docx"
    output_path = os.path.join(output_dir, nombre_archivo)

    # Guardar el documento con las respuestas reemplazadas
    doc.save(output_path)
    messagebox.showinfo("Éxito", f"Documento '{nombre_archivo}' creado con éxito en {output_path}.")

# Función para crear el documento basado en una plantilla
def crear_carta_recepcion_desde_plantilla(plantilla_path, respuestas, output_dir):
    # Cargar la plantilla
    doc = Document(plantilla_path)

    # Reemplazar los marcadores con las respuestas
    reemplazar_marcadores(doc, respuestas)

    # Formatear el nombre del archivo
    nombre_archivo = f"{respuestas['Nombre']}_{respuestas['Puesto']}_{respuestas['Equipo']}_Carta_Recepcion.docx"
    output_path = os.path.join(output_dir, nombre_archivo)

    # Guardar el documento con las respuestas reemplazadas
    doc.save(output_path)
    messagebox.showinfo("Éxito", f"Documento '{nombre_archivo}' creado con éxito en {output_path}.")

# Función para abrir un archivo de plantilla
def abrir_plantilla():
    ruta = filedialog.askopenfilename(filetypes=[("Documentos de Word", "*.docx")])
    return ruta

# Función para manejar el cuestionario desde la interfaz gráfica
def cuestionario():
    respuestas = {
        "Equipo": entry_equipo.get(),
        "Modelo": entry_modelo.get(),
        "No_Serie": entry_no_serie.get(),
        "IMEI": entry_IMEI.get(),
        "Procesador": entry_procesador.get(),
        "RAM": entry_ram.get(),
        "SistemaOperativo": entry_sistema_operativo.get(),
        "Almacenamiento": entry_almacenamiento.get(),
        "Clave": entry_clave.get(),
        "Cargador": entry_cargador.get(),
        "Mouse": entry_mouse.get(),
        "Puesto": entry_puesto.get(),
        "Nombre": entry_nombre.get()
    }

    plantilla_responsiva_path = abrir_plantilla()
    plantilla_recepcion_path = abrir_plantilla()

    if not plantilla_responsiva_path or not plantilla_recepcion_path:
        messagebox.showwarning("Advertencia", "Debes seleccionar ambas plantillas.")
        return

    base_path = os.path.dirname(plantilla_responsiva_path)
    responsivas_dir, recepcion_dir = crear_directorios(base_path)

    crear_carta_responsiva_desde_plantilla(plantilla_responsiva_path, respuestas, responsivas_dir)
    crear_carta_recepcion_desde_plantilla(plantilla_recepcion_path, respuestas, recepcion_dir)

# Crear la interfaz gráfica
root = tk.Tk()
root.title("Generador de Cartas")

# Crear y posicionar los widgets
labels_entries = [
    ("Equipo", "entry_equipo"),
    ("Modelo", "entry_modelo"),
    ("No. Serie", "entry_no_serie"),
    ("IMEI", "entry_IMEI"),
    ("Procesador", "entry_procesador"),
    ("RAM", "entry_ram"),
    ("Sistema Operativo", "entry_sistema_operativo"),
    ("Almacenamiento", "entry_almacenamiento"),
    ("Clave", "entry_clave"),
    ("Cargador", "entry_cargador"),
    ("Mouse", "entry_mouse"),
    ("Puesto", "entry_puesto"),
    ("Nombre", "entry_nombre")
]

for i, (label_text, entry_var) in enumerate(labels_entries):
    label = tk.Label(root, text=label_text)
    label.grid(row=i, column=0, padx=10, pady=5)
    entry = tk.Entry(root)
    entry.grid(row=i, column=1, padx=10, pady=5)
    globals()[entry_var] = entry

submit_button = tk.Button(root, text="Generar Cartas", command=cuestionario)
submit_button.grid(row=len(labels_entries), columnspan=2, pady=10)

# Ejecutar la aplicación
root.mainloop()

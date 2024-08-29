import ifcopenshell
import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo
from datetime import datetime
import re
from tqdm import trange
from time import sleep
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import threading
import logging

# FUNCIONES ORIGINALES

def get_entities_filtered(ifcschema_entities, get_types):
    ents_not_collect = ["IfcGeometricRepresentationItem", "IfcObject",
                        "IfcObjectDefinition", "IfcProduct", "IfcRelationship",
                        "IfcRepresentationItem", "IfcRoot", "IfcCarPoiLisD"]

    if get_types:
        ents = [e for e in ifcschema_entities if "type" in e.lower()]
    else:
        ents = [e for e in ifcschema_entities if "type" not in e.lower() and not e in ents_not_collect]
    return (len(ents), ents)

def get_ents_info_to_df(ifc_entName):
    entity_info = [en.get_info() for en in ifc_file.by_type(ifc_entName)]
    df = pd.DataFrame(entity_info)
    return df

def contract_entName(entName, trunc_n=3):
    entName_split = re.findall('[A-Z][^A-Z]*', entName)
    return "".join([s[:trunc_n] if len(s) >= trunc_n else s for s in entName_split])

def create_ws(wb, ws_name):
    if ws_name in wb.sheetnames:
        ws = wb.create_sheet(title=ws_name + "_1")
    else:
        ws = wb.create_sheet(title=ws_name)
    return ws

def create_ws_and_table(wb, ifc_entName, table_suffix_counter={}):
    df = get_ents_info_to_df(ifc_entName)
    if len(ifc_entName) > 20:
        ifc_entName = contract_entName(ifc_entName, trunc_n=3)
    ws = create_ws(wb, ifc_entName)
    for row in dataframe_to_rows(df, index=False, header=True):
        ws.append([str(v) if str(v).startswith('#') or '(' in str(v) else v for v in row])
    df_shape = df.shape
    last_column_letter = get_excel_column_letter(df_shape[1] - 1)
    tbl_xl_rangeAddress = f"A1:{last_column_letter}{df_shape[0] + 1}"
    if not re.match(r'^[A-Z]{1,3}\d+:[A-Z]{1,3}\d+$', tbl_xl_rangeAddress):
        raise ValueError(f"Invalid range address generated: {tbl_xl_rangeAddress}")
    base_table_name = ifc_entName
    suffix = table_suffix_counter.get(base_table_name, 0)
    unique_table_name = f"{base_table_name}_{suffix}"
    table_suffix_counter[base_table_name] = suffix + 1
    tbl = Table(displayName=unique_table_name, ref=tbl_xl_rangeAddress)
    style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=True)
    tbl.tableStyleInfo = style
    ws.add_table(tbl)
    return True

def remove_ws(wb, ws_name="Sheet"):
    if ws_name in wb.sheetnames:
        ws_to_delete = wb[ws_name]
        wb.remove(ws_to_delete)

def purge_wb(wb):
    for ws in wb.sheetnames:
        ws = wb[ws]
    for tbl in ws.tables.values(): ws._tables.remove(tbl)
    remove_ws(wb, ws_name=ws)
    return True

def get_excel_column_letter(n):
    string = ""
    while n >= 0:
        n, remainder = divmod(n, 26)
        string = chr(65 + remainder) + string
        n -= 1
    return string

# FUNCIONES PARA LA INTERFAZ GRÁFICA

def select_ifc_file():
    global ifc_filename, ifc_basepath, ifc_file
    file_path = filedialog.askopenfilename(filetypes=[("IFC files", "*.ifc")])
    if file_path:
        ifc_filename = os.path.basename(file_path).split('.')[0]
        ifc_basepath = os.path.dirname(file_path)
        ifc_file = ifcopenshell.open(file_path)
        file_entry.delete(0, tk.END)
        file_entry.insert(0, file_path)
        console_output.insert(tk.END, f"Archivo IFC seleccionado: {file_path}\n")

def select_destination_folder():
    global destination_folder
    destination_folder = filedialog.askdirectory()
    folder_entry.delete(0, tk.END)
    folder_entry.insert(0, destination_folder)
    console_output.insert(tk.END, f"Carpeta de destino seleccionada: {destination_folder}\n")

def start_processing_thread():
    # Deshabilitar botones
    btn_select_file.config(state=tk.DISABLED)
    btn_select_folder.config(state=tk.DISABLED)
    btn_process.config(state=tk.DISABLED)
    # Crear un hilo separado para ejecutar el procesamiento
    thread = threading.Thread(target=process_ifc_file)
    thread.start()

logging.basicConfig(filename='app.log', level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

def process_ifc_file():
    logging.info("Inicio del procesamiento")
    try:
        if not ifc_file or not destination_folder:
            messagebox.showerror("Error", "Seleccione un archivo IFC y una carpeta de destino antes de continuar.")
            return

        global ifc_getTypes
        ifc_getTypes = False

        # Carga del modelo y procesamiento
        sch_entities_names = [e.name() for e in ifcopenshell.schema_by_name(ifc_file.schema).entities()]
        entities_names_in_use = []
        for en in sch_entities_names:
            try:
                if len(ifc_file.by_type(en)) != 0:
                    entities_names_in_use.append(en)
            except RuntimeError as e:
                console_output.insert(tk.END, f"Advertencia: {e} - Entidad '{en}' ignorada.\n")

        ents = get_entities_filtered(entities_names_in_use, get_types=ifc_getTypes)
        xls_filename = os.path.join(destination_folder, datetime.now().strftime("%Y-%m-%d") + f"_{ifc_filename}_entTypes_{str(ifc_getTypes)}_2.xlsx")

        wb = Workbook()
        purge_wb(wb)

        table_suffix_counter = {}
        t = trange(len(ents[1]), desc='Entity: ', leave=True)
        progress_bar["maximum"] = len(ents[1])  # Ajusta la barra de progreso
        for i, en in zip(t, ents[1]):
            t.set_description(f"Entity: {en}")
            t.refresh()
            create_ws_and_table(wb, en, table_suffix_counter)
            progress_bar["value"] = i + 1  # Actualiza la barra de progreso
            progress_bar.update_idletasks()

        remove_ws(wb, ws_name="Sheet")
        wb.save(xls_filename)
        console_output.insert(tk.END, f"Archivo {xls_filename} ha sido guardado exitosamente!\n")
    except Exception as e:
        logging.error(f"Error durante el procesamiento: {str(e)}")
    finally:
        enable_buttons()  # Habilitar botones al finalizar el procesamiento

def enable_buttons():
    """Función para habilitar los botones después de finalizar el procesamiento."""
    btn_select_file.config(state=tk.NORMAL)
    btn_select_folder.config(state=tk.NORMAL)
    btn_process.config(state=tk.NORMAL)

# CONFIGURACIÓN DE LA INTERFAZ GRÁFICA

app = tk.Tk()
app.title("Interfaz para Procesar Archivos IFC")
app.geometry("700x500")  # Ajusta el tamaño de la ventana

# Frame para seleccionar archivo IFC
frame_file = tk.Frame(app)
frame_file.pack(fill=tk.X, padx=20, pady=5)

file_entry = tk.Entry(frame_file)
file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True,padx=(0, 5))

btn_select_file = tk.Button(frame_file, text="Seleccionar Archivo IFC", command=select_ifc_file, width=25)
btn_select_file.pack(side=tk.RIGHT)

# Frame para seleccionar carpeta de destino
frame_folder = tk.Frame(app)
frame_folder.pack(fill=tk.X, padx=20, pady=5)

folder_entry = tk.Entry(frame_folder)
folder_entry.pack(side=tk.LEFT, fill=tk.X, expand=True,padx=(0, 5))

btn_select_folder = tk.Button(frame_folder, text="Seleccionar Carpeta de Destino", command=select_destination_folder, width=25)
btn_select_folder.pack(side=tk.RIGHT)

# Botón para procesar el archivo
btn_process = tk.Button(app, text="Procesar Archivo",  command=start_processing_thread)
btn_process.pack(pady=10)

# Barra de progreso
progress_bar = ttk.Progressbar(app, orient="horizontal", length=100, mode="determinate")
progress_bar.pack(fill=tk.X, padx=20, pady=10)

# Consola de salida
console_output = scrolledtext.ScrolledText(app, wrap=tk.WORD, width=80, height=10)
console_output.pack(padx=20, pady=10, fill=tk.BOTH, expand=True)

app.mainloop()
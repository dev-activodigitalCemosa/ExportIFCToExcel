import ifcopenshell
import ifcopenshell.util
import ifcopenshell.geom
import ifcopenshell.util.element
import ifcopenshell.util.placement
import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo
from datetime import datetime
import re
from tqdm import trange
from time import sleep
import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import logging

# CONFIGURACIÓN DE CustomTkinter
ctk.set_appearance_mode("System")  # "System", "Dark", "Light"
ctk.set_default_color_theme("blue")  # "blue", "green", "dark-blue"

# FUNCIONES ORIGINALES

def get_entities_filtered(ifcschema_entities, get_types):
    # ents_not_collect = ["IfcGeometricRepresentationItem", "IfcObject",
    #                     "IfcObjectDefinition", "IfcProduct", "IfcRelationship",
    #                     "IfcRepresentationItem", "IfcRoot", "IfcCarPoiLisD"]
    ents_not_collect = ["IfcColumn","IfcElement"]

    if get_types:
        ents = [e for e in ifcschema_entities if "type" in e.lower()]
    else:
        ents = [e for e in ifcschema_entities if "type" not in e.lower() and e in ents_not_collect]
    return (len(ents), ents)



def get_ents_info_to_df(ifc_entName):
    #campos a filtrar
    fields_to_exclude = ['Description','ObjectPlacement','Representation']
    # Obtener la información básica
    entities = ifc_file.by_type(ifc_entName)
    entity_info = [en.get_info() for en in entities]
    entity_container = [ifcopenshell.util.element.get_container(en) for en in entities]
    entity_coord = [ifcopenshell.util.placement.get_local_placement(en.ObjectPlacement)[:,3][:3] for en in entities]

    entity_properties = []
    for en in entities:
        props = ifcopenshell.util.element.get_psets(en)
        filtered_props = {key: value for key, value in props.items() if 'Pset' not in key}
        entity_properties.append(filtered_props)
    

    # Crear un DataFrame con MultiIndex en columnas

    # if not df_props_expanded.empty:
    #     # Obtener los niveles del MultiIndex a partir de las columnas existentes
    #     columns = df_props_expanded.columns
    #     tuples = [col.split('.', 1) for col in columns]
    #     multi_index = pd.MultiIndex.from_tuples(tuples, names=['Category', 'Property'])
    #     df_props_expanded.columns = multi_index
    
    # Crear DataFrames
    df_info = pd.DataFrame(entity_info)
    df_props_expanded = pd.json_normalize(entity_properties)
    df_container = pd.DataFrame(entity_container, columns=['Location'])
    df_coord = pd.DataFrame(entity_coord, columns=['X', 'Y', 'Z'])  # Nombrar las columnas XYZ

     # Filtrar campos específicos en df_info
    if fields_to_exclude:
        existing_fields = [col for col in df_info.columns if col not in fields_to_exclude]
        df_info = df_info[existing_fields]

    # Concatenar los DataFrames
    combined_df = pd.concat([df_info, df_props_expanded, df_container, df_coord], axis=1)
    
    return combined_df

def contract_entName(entName, trunc_n=3):
    entName_split = re.findall('[A-Z][^A-Z]*', entName)
    return "".join([s[:trunc_n] if len(s) >= trunc_n else s for s in entName_split])

def create_ws(wb, ws_name):
    if ws_name in wb.sheetnames:
        ws = wb.create_sheet(title=ws_name + "_1")
    else:
        ws = wb.create_sheet(title=ws_name)
    return ws

# def create_ws_and_table(wb, ifc_entName, table_suffix_counter={}):
#     df = get_ents_info_to_df(ifc_entName)
#     if len(ifc_entName) > 20:
#         ifc_entName = contract_entName(ifc_entName, trunc_n=3)
#     ws = create_ws(wb, ifc_entName)
#     for row in dataframe_to_rows(df, index=False, header=True):
#         ws.append([str(v) if str(v).startswith('#') or '(' in str(v) else v for v in row])
#     df_shape = df.shape
#     last_column_letter = get_excel_column_letter(df_shape[1] - 1)
#     tbl_xl_rangeAddress = f"A1:{last_column_letter}{df_shape[0] + 1}"
#     if not re.match(r'^[A-Z]{1,3}\d+:[A-Z]{1,3}\d+$', tbl_xl_rangeAddress):
#         raise ValueError(f"Invalid range address generated: {tbl_xl_rangeAddress}")
#     base_table_name = ifc_entName
#     suffix = table_suffix_counter.get(base_table_name, 0)
#     unique_table_name = f"{base_table_name}_{suffix}"
#     table_suffix_counter[base_table_name] = suffix + 1
#     tbl = Table(displayName=unique_table_name, ref=tbl_xl_rangeAddress)
#     style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=True)
#     tbl.tableStyleInfo = style
#     ws.add_table(tbl)
#     return True

def create_ws_and_table(wb, ifc_entName, table_suffix_counter={}):
    df = get_ents_info_to_df(ifc_entName)
    
    if len(ifc_entName) > 20:
        ifc_entName = contract_entName(ifc_entName, trunc_n=3)
    
    ws = create_ws(wb, ifc_entName)
    
    # Escribir encabezados jerárquicos
    if not df.empty:
        if isinstance(df.columns, pd.MultiIndex):
            # Escribir encabezados de nivel 1
            for col_num, (level1, _) in enumerate(df.columns, start=1):
                ws.cell(row=1, column=col_num, value=str(level1) if level1 else '')
            # Escribir encabezados de nivel 2
            for col_num, (_, level2) in enumerate(df.columns, start=1):
                ws.cell(row=2, column=col_num, value=str(level2) if level2 else '')
        else:
            # Para columnas simples, solo escribir el encabezado en la primera fila
            for col_num, col_name in enumerate(df.columns, start=1):
                ws.cell(row=1, column=col_num, value=str(col_name))

        # Escribir los datos del DataFrame
        for row in dataframe_to_rows(df, index=False, header=False):
            ws.append([str(v) if str(v).startswith('#') or '(' in str(v) else v for v in row])

        df_shape = df.shape
        last_column_letter = get_excel_column_letter(df_shape[1]-1)
        tbl_xl_rangeAddress = f"A1:{last_column_letter}{df_shape[0] + 2}"

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
        file_entry.delete(0, ctk.END)
        file_entry.insert(0, file_path)
        console_output.insert(ctk.END, f"Archivo IFC seleccionado: {file_path}\n")

def select_destination_folder():
    global destination_folder
    destination_folder = filedialog.askdirectory()
    folder_entry.delete(0, ctk.END)
    folder_entry.insert(0, destination_folder)
    console_output.insert(ctk.END, f"Carpeta de destino seleccionada: {destination_folder}\n")

def start_processing_thread():
    # Deshabilitar botones
    btn_select_file.configure(state=ctk.DISABLED)
    btn_select_folder.configure(state=ctk.DISABLED)
    btn_process.configure(state=ctk.DISABLED)
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
                console_output.insert(ctk.END, f"Advertencia: {e} - Entidad '{en}' ignorada.\n")
        ents = get_entities_filtered(entities_names_in_use, get_types=ifc_getTypes)
        xls_filename = os.path.join(destination_folder, datetime.now().strftime("%Y-%m-%d") + f"_{ifc_filename}_entTypes_{str(ifc_getTypes)}_2.xlsx")

        wb = Workbook()
        purge_wb(wb)

        table_suffix_counter = {}
        t = trange(len(ents[1]), desc='Entity: ', leave=True)
        progress_bar.set(0)
        total_ents = len(ents[1])
        for i, en in zip(t, ents[1]):
            t.set_description(f"Entity: {en}")
            t.refresh()
            create_ws_and_table(wb, en, table_suffix_counter)
            # Calcula el progreso actual y actualiza la barra de progreso
            progress = (i + 1) / total_ents
            progress_bar.set(progress)  # Actualiza la barra de progreso
            progress_bar.update_idletasks()

        remove_ws(wb, ws_name="Sheet")
        wb.save(xls_filename)
        console_output.insert(ctk.END, f"Archivo {xls_filename} ha sido guardado exitosamente!\n")
    except Exception as e:
        logging.error(f"Error durante el procesamiento: {str(e)}")
    finally:
        enable_buttons()  # Habilitar botones al finalizar el procesamiento

def enable_buttons():
    """Función para habilitar los botones después de finalizar el procesamiento."""
    btn_select_file.configure(state=ctk.NORMAL)
    btn_select_folder.configure(state=ctk.NORMAL)
    btn_process.configure(state=ctk.NORMAL)

# CONFIGURACIÓN DE LA INTERFAZ GRÁFICA

app = ctk.CTk()  # Cambiar tk.Tk() por ctk.CTk()
app.title("Exportador de propiedades IFC")
app.geometry("700x500")  # Ajusta el tamaño de la ventana

# Frame para seleccionar archivo IFC
frame_file = ctk.CTkFrame(app)
frame_file.pack(fill=ctk.X, padx=20, pady=5)

file_entry = ctk.CTkEntry(frame_file)
file_entry.pack(side=ctk.LEFT, fill=ctk.X, expand=True, padx=(0, 5))

btn_select_file = ctk.CTkButton(frame_file, text="Seleccionar Archivo IFC", command=select_ifc_file, width=25)
btn_select_file.pack(side=ctk.RIGHT)

# Frame para seleccionar carpeta de destino
frame_folder = ctk.CTkFrame(app)
frame_folder.pack(fill=ctk.X, padx=20, pady=5)

folder_entry = ctk.CTkEntry(frame_folder)
folder_entry.pack(side=ctk.LEFT, fill=ctk.X, expand=True, padx=(0, 5))

btn_select_folder = ctk.CTkButton(frame_folder, text="Seleccionar Carpeta de Destino", command=select_destination_folder, width=25)
btn_select_folder.pack(side=ctk.RIGHT)

# Botón para procesar el archivo
btn_process = ctk.CTkButton(app, text="Procesar Archivo", command=start_processing_thread)
btn_process.pack(pady=10)

# Barra de progreso
progress_bar = ctk.CTkProgressBar(app, orientation="horizontal", width=100, mode="determinate")
progress_bar.pack(fill=ctk.X, padx=20, pady=10)

# Consola de salida
console_output = ctk.CTkTextbox(app, wrap=ctk.WORD, width=80, height=10)
console_output.pack(padx=20, pady=10, fill=ctk.BOTH, expand=True)

app.mainloop()

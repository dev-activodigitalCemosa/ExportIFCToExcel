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


#FUNCTIONS

def get_entities_filtered(ifcschema_entities, get_types):
    ents_not_collect = ["IfcGeometricRepresentationItem", "IfcObject", \
                    "IfcObjectDefinition", "IfcProduct", "IfcRelationship", \
                    "IfcRepresentationItem", "IfcRoot"]

    if get_types:
        ents = [e for e in ifcschema_entities \
           if "type" in e.lower()]
    else:
        ents = [e for e in ifcschema_entities \
               if "type" not in e.lower() and not e in ents_not_collect]
    return (len(ents),ents)


def get_ents_info_to_df(ifc_entName):
    entity_info = [en.get_info() for en in ifc_file.by_type(ifc_entName)]
#print(entity_info)
#convert data to a pandas dataframe
    df = pd.DataFrame(entity_info)
    return df

def contract_entName (entName, trunc_n = 3):
#takes a Ifc entity name, split it at each uppercase occurrence, then slice each portion taking its truncon leftmost chars, finally concatenates hack the contracted Ifc entity name in a string
    entName_split= re.findall('[A-Z][^A-Z]*', entName)
    return "".join([s[:trunc_n] if len(s) >= trunc_n else s for s in entName_split])

def create_ws(wb, ws_name):
#Check if the sheet exists
    if ws_name in wb.sheetnames:
        ws = wb.create_sheet(title=ws_name + "_1")
    else: 
        ws = wb.create_sheet(title=ws_name)
    return ws

def create_ws_and_table(wb,ifc_entName):
#Get entity info DataFrame 
    df= get_ents_info_to_df(ifc_entName)
#Reduce Lengthy entity names in contract verston
    if len(ifc_entName) > 20: 
        ifc_entName = contract_entName(ifc_entName, trunc_n = 3)
#Create a new worksheet named as per Ifc entName
    ws= create_ws(wb,ifc_entName)
#write the DataFrame to the Excel file
    for row in dataframe_to_rows(df, index=False, header=True):
        ws.append([str(v) if str(v).startswith('#') or '(' in str(v) else v for v in row])
#Get dataframe shape 
    df_shape = df.shape
#Set table Excel range
    tbl_xl_rangeAddress = f"A1:{cols_dict[df_shape[1] - 1]}{df_shape[0] + 1}"
    
    # Verifica que el rango esté bien formado
    if not re.match(r'^[A-Z]{1,3}\d+:[A-Z]{1,3}\d+$', tbl_xl_rangeAddress):
        raise ValueError(f"Invalid range address generated: {tbl_xl_rangeAddress}")
#Create an Excel table 
    tbl = Table(displayName=ifc_entName, ref=tbl_xl_rangeAddress)
#Add a default style wit…
    style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=True)
    tbl.tableStyleInfo = style
#commit table to worksheet
    ws.add_table(tbl)
    return True

def remove_ws(wb, ws_name="Sheet"):
# check if the sheet exists
    if ws_name in wb.sheetnames:
#Get the sheet to delete
        ws_to_delete = wb[ws_name]
#Delete the sheet
        wb.remove(ws_to_delete)
        print(f"Sheet {ws_name} has been deleted.")

def purge_wb(wb):
# Iterate through each sheet
    for ws in wb.sheetnames:
        ws = wb[ws]
#Remove all tables
    for tbl in ws.tables.values(): ws._tables.remove(tbl)
#Remove Ws
    remove_ws(wb, ws_name=ws)
    return True

 #INPUTS

ifc_basepath = "C:\\Users\\Trabajo\\Desktop\\ModelosIfc" 
ifc_filename = "SPAHOTEL"
ifc_getTypes = False

#load model
ifc_file = ifcopenshell.open(f"{ifc_basepath}\\{ifc_filename}.ifc")
# get all used entities in a model
sch_entities_names = [e.name() for e in ifcopenshell.schema_by_name(ifc_file.schema).entities()]
entities_names_in_use = [en for en in sch_entities_names if len(ifc_file.by_type(en))!=0]
 #entitles nomes_in_use filtered entities names
ents = get_entities_filtered(entities_names_in_use, get_types=ifc_getTypes)
ents

# construct output Excel filename
xls_filename= os.getcwd() + "\\exports2XLS\\" + datetime.now().strftime("%Y-%m-%d") + \
            "_"+ ifc_filename + "_entTypes_" + str(ifc_getTypes)+ "_2.xlsx"
xls_filename

#create a dictionary
cols_dict = dict(enumerate([chr(i) for i in range(65, 91)]))
cols_dict

# Create a new Excel workbook
wb = Workbook()

#Remove all sheets and tables 
purge_wb(wb)

#Save the Excel file
wb.save(xls_filename)

#Add a new worksheet and a table with all entities data, showing a progress bar 
t= trange(len(ents[1]), desc='Entity: ', leave=True)
for i, en in zip(t,ents[1]):
    t.set_description(f"Entity: (en)")
    t.refresh() #to show immediately the update
    create_ws_and_table(wb,en)

#Remove default worksheet named "Sheet" if present
remove_ws(wb,ws_name="Sheet")

#Save the Excel file 
wb.save(xls_filename)
print(f"File (xls_filename) has been saved successfully!")

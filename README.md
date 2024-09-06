# Exportador de Propiedades IFC

Este proyecto permite seleccionar archivos IFC y exportar información relevante sobre las entidades del modelo IFC a un archivo Excel. La aplicación está construida utilizando Python, las librerías `ifcopenshell`, `openpyxl`, y `customtkinter`.

## Requisitos

Para que el proyecto funcione correctamente, asegúrate de instalar las siguientes dependencias:

```bash
pip install ifcopenshell openpyxl customtkinter tqdm
```
<<<<<<< HEAD
```markdown
Librerías utilizadas:
- ifcopenshell: Permite leer, escribir y procesar archivos IFC.
- openpyxl: Manejo de archivos Excel (.xlsx).
pandas: Manipulación de datos en DataFrames.
customtkinter: Creación de una interfaz gráfica más moderna con Tkinter.
tqdm: Proporciona barras de progreso.
threading: Permite ejecutar procesos en hilos separados.
logging: Para registrar errores e información del proceso.
tkinter: Para gestión de cuadros de diálogo y mensajes en la GUI.
datetime: Para crear nombres de archivo con marcas de tiempo.
Funcionalidad
El programa realiza las siguientes operaciones:
```
Selección del archivo IFC: A través de un cuadro de diálogo, el usuario puede seleccionar un archivo IFC.
Selección de la carpeta de destino: El usuario elige una carpeta donde se guardará el archivo Excel generado.
Procesamiento del archivo IFC: El sistema analiza el archivo IFC seleccionado y exporta la información de las entidades al archivo Excel.
Barra de progreso: El usuario puede observar el progreso del procesamiento mediante una barra de progreso visual.
Salida de consola: Se muestran mensajes e información sobre el procesamiento en una consola dentro de la interfaz gráfica.
Estructura del Código
Funcionalidades de procesamiento:

get_entities_filtered(ifcschema_entities, get_types): Filtra las entidades IFC en función de si se desea obtener tipos (get_types=True) o no. Ignora algunas entidades que no son relevantes.

get_ents_info_to_df(ifc_entName): Obtiene información detallada sobre una entidad IFC y la organiza en un DataFrame. Extrae información básica, propiedades, contenedores y coordenadas de las entidades.

contract_entName(entName, trunc_n=3): Contrae el nombre de una entidad para que no exceda un número de caracteres, lo cual es útil para nombres largos en Excel.

create_ws(wb, ws_name): Crea una hoja de trabajo (worksheet) en el archivo Excel con el nombre proporcionado. Si el nombre ya existe, le añade un sufijo para hacerlo único.

create_ws_and_table(wb, ifc_entName, table_suffix_counter={}): Crea una hoja de trabajo y una tabla en el archivo Excel para una entidad IFC dada. Inserta los datos de la entidad y ajusta el formato de la tabla.

remove_ws(wb, ws_name="Sheet"): Elimina una hoja específica del archivo Excel si existe.

purge_wb(wb): Elimina todas las hojas y tablas de un archivo Excel para comenzar con una plantilla limpia.

get_excel_column_letter(n): Convierte un número de columna en su correspondiente letra de Excel (por ejemplo, 1 -> 'A', 26 -> 'Z').

Interfaz Gráfica:

select_ifc_file(): Abre un cuadro de diálogo para seleccionar un archivo IFC y muestra el nombre del archivo seleccionado en la interfaz.

select_destination_folder(): Abre un cuadro de diálogo para seleccionar una carpeta de destino donde se guardará el archivo Excel generado.

start_processing_thread(): Inicia un nuevo hilo para procesar el archivo IFC y deshabilita los botones durante la ejecución.

process_ifc_file(): Ejecuta el proceso de lectura del archivo IFC, filtra las entidades y crea el archivo Excel correspondiente. Incluye manejo de errores y muestra el progreso en la barra de progreso.

enable_buttons(): Habilita los botones de la interfaz después de que el procesamiento ha terminado.

Configuración de la interfaz gráfica (GUI): Se utiliza customtkinter para construir una interfaz moderna que incluye:

Un cuadro de texto para mostrar el archivo IFC seleccionado.
Un cuadro de texto para mostrar la carpeta de destino seleccionada.
Un botón para iniciar el procesamiento.
Una barra de progreso que muestra el avance durante la ejecución.
Una consola de salida que muestra mensajes del proceso.
Detalles de Implementación
Multi-hilo: Para evitar que la interfaz se congele durante el procesamiento del archivo IFC, el proceso se ejecuta en un hilo separado.

Registros: Se utilizan logging y messagebox para registrar errores e información en tiempo de ejecución.

Formateo de Excel: Utiliza openpyxl para crear archivos Excel con tablas formateadas y escribe encabezados jerárquicos (en caso de que el DataFrame tenga MultiIndex).

Estructura de la Aplicación
Flujo del Programa:
El usuario selecciona un archivo IFC y una carpeta de destino.
El usuario inicia el procesamiento del archivo.
El programa lee y filtra las entidades del archivo IFC.
Los datos de las entidades se exportan a un archivo Excel.
El archivo Excel es guardado en la carpeta de destino seleccionada.
La interfaz muestra el progreso y notifica al usuario cuando el proceso ha terminado.
Configuración de la Interfaz Gráfica (Tkinter)
Apariencia: Se utiliza el modo oscuro de CustomTkinter para mejorar la legibilidad.
Widgets:
Entradas de texto: Para mostrar el archivo y la carpeta seleccionados.
Botones: Para seleccionar archivo/carpeta y para iniciar el procesamiento.
Barra de progreso: Indica el avance del procesamiento.
Consola: Muestra mensajes de salida en tiempo real.
Cómo usar la aplicación
Ejecuta el script.
Selecciona un archivo IFC usando el botón "Seleccionar Archivo IFC".
Selecciona una carpeta de destino para el archivo Excel.
Haz clic en "Procesar Archivo" para comenzar la exportación.
Observa el progreso en la barra de progreso y en la consola de salida.
Al finalizar, el archivo Excel estará disponible en la carpeta seleccionada.
Posibles Mejoras
Mejor manejo de excepciones: Actualmente se registra cualquier error que ocurra, pero podría implementarse un manejo más específico para mejorar la robustez del sistema.
Soporte para diferentes formatos: Extender la aplicación para exportar a otros formatos como CSV o JSON.
Interfaz más personalizada: Incluir más opciones de configuración visual y manejo de eventos para mejorar la experiencia del usuario.
=======

## Librerías utilizadas:
- __ifcopenshell__: Permite leer, escribir y procesar archivos IFC.
- __openpyxl__: Manejo de archivos Excel (.xlsx).
- __pandas__: Manipulación de datos en DataFrames.
- __customtkinter__: Creación de una interfaz gráfica más moderna con Tkinter.
- __tqdm__: Proporciona barras de progreso.
- __threading__: Permite ejecutar procesos en hilos separados.
- __logging__: Para registrar errores e información del proceso.
- __tkinter__: Para gestión de cuadros de diálogo y mensajes en la GUI.
- __datetime__: Para crear nombres de archivo con marcas de tiempo.

## Funcionalidad

- __Selección del archivo IFC__: A través de un cuadro de diálogo, el usuario puede seleccionar un archivo IFC.
- __Selección de la carpeta de destino__: El usuario elige una carpeta donde se guardará el archivo Excel generado.
- __Procesamiento del archivo IFC__: El sistema analiza el archivo IFC seleccionado y exporta la información de las entidades al archivo Excel.
- __Barra de progreso__: El usuario puede observar el progreso del procesamiento mediante una barra de progreso visual.
- __Salida de consola__: Se muestran mensajes e información sobre el procesamiento en una consola dentro de la interfaz gráfica.
  
## Funcionalidades de procesamiento:

```phyton
get_entities_filtered(ifcschema_entities, get_types): #Filtra las entidades IFC en función de si se desea obtener tipos (get_types=True) o no. Ignora algunas entidades que no son relevantes.

get_ents_info_to_df(ifc_entName): #Obtiene información detallada sobre una entidad IFC y la organiza en un DataFrame. Extrae información básica, propiedades, contenedores y coordenadas de las entidades.

contract_entName(entName, trunc_n=3): #Contrae el nombre de una entidad para que no exceda un número de caracteres, lo cual es útil para nombres largos en Excel.

create_ws(wb, ws_name): #Crea una hoja de trabajo (worksheet) en el archivo Excel con el nombre proporcionado. Si el nombre ya existe, le añade un sufijo para hacerlo único.

create_ws_and_table(wb, ifc_entName, table_suffix_counter={}): #Crea una hoja de trabajo y una tabla en el archivo Excel para una entidad IFC dada. Inserta los datos de la entidad y ajusta el formato de la tabla.

remove_ws(wb, ws_name="Sheet"): #Elimina una hoja específica del archivo Excel si existe.

purge_wb(wb): #Elimina todas las hojas y tablas de un archivo Excel para comenzar con una plantilla limpia.

get_excel_column_letter(n): #Convierte un número de columna en su correspondiente letra de Excel (por ejemplo, 1 -> 'A', 26 -> 'Z').
```

## Interfaz Gráfica:
```phyton
select_ifc_file(): #Abre un cuadro de diálogo para seleccionar un archivo IFC y muestra el nombre del archivo seleccionado en la interfaz.

select_destination_folder(): #Abre un cuadro de diálogo para seleccionar una carpeta de destino donde se guardará el archivo Excel generado.

start_processing_thread(): #Inicia un nuevo hilo para procesar el archivo IFC y deshabilita los botones durante la ejecución.

process_ifc_file(): #Ejecuta el proceso de lectura del archivo IFC, filtra las entidades y crea el archivo Excel correspondiente. Incluye manejo de errores y muestra el progreso en la barra de progreso.

enable_buttons(): #Habilita los botones de la interfaz después de que el procesamiento ha terminado.

```

## Flujo del Programa:

- El usuario selecciona un archivo IFC y una carpeta de destino.
- El usuario inicia el procesamiento del archivo.
- El programa lee y filtra las entidades del archivo IFC.
- Los datos de las entidades se exportan a un archivo Excel.
- El archivo Excel es guardado en la carpeta de destino seleccionada.
- La interfaz muestra el progreso y notifica al usuario cuando el proceso ha terminado.
  
Configuración de la Interfaz Gráfica (Tkinter)
- Apariencia: Se utiliza el modo predeterminado de tu sistema.
- Entradas de texto: Para mostrar el archivo y la carpeta seleccionados.
- Botones: Para seleccionar archivo/carpeta y para iniciar el procesamiento.
- Barra de progreso: Indica el avance del procesamiento.
- Consola: Muestra mensajes de salida en tiempo real.

## Cómo usar la aplicación

1. Ejecuta el script.
2. Selecciona un archivo IFC usando el botón "Seleccionar Archivo IFC".
3. Selecciona una carpeta de destino para el archivo Excel.
4. Haz clic en "Procesar Archivo" para comenzar la exportación.
5. Observa el progreso en la barra de progreso y en la consola de salida.
6. Al finalizar, el archivo Excel estará disponible en la carpeta seleccionada.

>>>>>>> develop

# Registro de Llaves de Apartamentos
# Función: Este programa permite registrar la entrega y devolución de llaves de apartamentos y zonas comunes.
# Autor: Tomás Rodríguez
# Fecha: 25/06/2025
# Versión: 1.1.1

import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import os
from datetime import datetime, time, timedelta
import traceback
import sys # Importar para sys.platform
import glob # Importar para buscar archivos por patrón
import calendar # Importar para obtener el número de días en un mes
import shutil # Importar shutil para operaciones de copia de archivos
import fpdf
from fpdf import FPDF, XPos, YPos

# --- Configuración y Datos Globales ---
NOMBRE_ARCHIVO_EXCEL = 'registro_llaves.xlsx'
# Prefijo para los archivos CSV de impresión temporal para facilitar su limpieza
PREFIJO_CSV_TEMPORAL = 'reporte_llaves_imprimir_'
# Nombre del archivo para cargar los nombres de los conserjes
NOMBRE_ARCHIVO_CONSERJES = 'conserjes.txt'

# Configuración para hacer un backup del archivo Excel registro_llaves.xlsx
CARPETA_BACKUPS = 'backup_llaves' # Nnombre de la carpeta de backups
DIAS_ENTRE_BACKUPS = 15 # Días entre cada backup, CADA 15 DÍAS HARA UN BACKUP


# Datos del programa para la sección "Acerca de"
VERSION_PROGRAMA = "1.1.1"
AUTOR_PROGRAMA = "Tomás Rodríguez" 
ANIO_COPY0RIGHT = "2025"

# Definición de puertas por Edificio y Planta
VALIDACION_PUERTAS = {
    'C1': {
        '0': ['1', '2', '3', '8'],
        '1': ['1', '2', '3', '4', '12', '13'],
        '2': ['1', '2', '3', '4', '12', '13'],
        '3': ['1', '2', '3', '4', '12', '13'],
        '4': ['1', '2', '3', '8'],
        '5': ['1', '2', '3', '8'],
        '6': ['1', '2'],
        'Local': ['1'],
        '-': ['Cont Luz', 'Cancela', 'Acceso Peatonal Garaje', 'Cont Agua', 'Cubierta', 'Registro Electrico + Centralita Telefónica 1ª Planta', 'Cuarto Maquinas Ascensores', 'Antena Tv', 'Acceso a Jardines']
    },
    'C2': {
        '0': ['4', '5', '6', '7'],
        '1': ['5', '6', '7', '8', '9', '10', '11'],
        '2': ['5', '6', '7', '8', '9', '10', '11'],
        '3': ['5', '6', '7', '8', '9', '10', '11'],
        '4': ['4', '5', '6', '7'],
        '5': ['4', '5', '6', '7'],
        '6': ['3', '4'],
        'Local': ['Fundación Banus', 'Banus Property Ofi', 'Banus Property Comedor'],
        '-': ['Cont Luz', 'Cancela', 'Acceso Peatonal Garaje', 'Cont Agua', 'Cubierta', 'Centralita Telefónica', 'Cuarto Maquinas Ascensores', 'Patio C2, 0-6']
    },
    'D1': {
        '0': ['1', '2', '3', '4'],
        '1': ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12'],
        '2': ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12'],
        '3': ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12'],
        '4': ['1', '2', '3', '4', '5', '6', '7', '8'],
        '5': ['1', '2', '3', '4', '5'],
        '6': ['1', '2'],
        'Local': ['Pyomar SL', 'Ilse Jodts', 'Ilse Jodts', 'Face clinic (Uricma)'],
        '-': ['Cont Luz', 'Cancela', 'Acceso Peatonal Garaje', 'Llave Brazo Desbloqueo', 'Cont Agua Planta Baja', 'RITI', 'Cuarto -1 (Jardineros)', 'Cubierta + Parabólica', 'Centralita Telefónica', 'Cuarto Maquinas Ascensores', 'Armario TV 2ª Planta']
    },
    'D2': {
        '0': ['1', '2'],
        '1': ['1', '2', '3', '4', '5', '6', '7'],
        '2': ['1', '2', '3', '4', '5', '6', '7'],
        '3': ['1', '2', '3', '4', '5', '6', '7'],
        '4': ['1', '2', '3', '4', '5'],
        '5': ['1', '2', '3'],
        '6': ['1', '2'],
        'Local': ['Banus Rental', 'ID Latinia', 'Farmacia', 'Femme Beaty'],
        '-': ['Cont Luz', 'Acceso a Patio (puerta crital)', 'Acceso Peatonal Garaje', 'Cont Agua Planta Baja', 'Centralita Telefónica', 'Armario Telefónica 4ª Planta', 'Cubierta', 'Cuarto de Limpieza 6ª Planta', 'Cuarto Maquinas Ascensores']
    }
}

# Lista única de todas las llaves válidas
LLAVES_VALIDAS = []
llaves_edificios = []
llaves_generales = []

# Generar llaves combinadas de los edificios 'normales'
for edificio in VALIDACION_PUERTAS:
    for planta, puertas in VALIDACION_PUERTAS[edificio].items():
        if planta == '-':
            for puerta in puertas:
                llaves_edificios.append(f"{edificio}, {planta}-{puerta}")
        elif planta == 'Local':
            for puerta in puertas:
                llaves_edificios.append(f"{edificio}, Local-{puerta}")
        else:
            for puerta in puertas:
                llaves_edificios.append(f"{edificio}, {planta}-{puerta}")

# Añadir las llaves generales/especiales directamente a su propia lista
llaves_generales.extend(['Grupo de Presión', 'Anexo D-E', 'Puerta Garaje', 'Magnetica', 'Puerta peatonal Garaje', 'Entrada Garaje','Salida Garaje','Puerta Mágnetica Puerto Banús', 'Puerta Luces Garaje', 'Puerta Magnetica Playa', 'Acceso a Patios C1 y C2', 'Cuadro Piscina', 'Cuarto Motores Piscina',
                        'Pozo', 'Caseta Jardineros', 'Candado Cancelas Anexo a Oficinas', 'Cuarto Desayuno Limpiadoras', 'Caseja Garaje CD', 'Armario Electrico Puertas Garaje', 'Caseta de Riego Piscinas Cádiz', 'Acceso nº 6', 'Acceso nº 5', 'Acceso nº 4', 'Cuadro de Luz Jardin Acceso Gym',
                        'Cuadro Eléctrico Jardin', 'Zonas comunes de uso para Jardinero', 'Caseta Vigilante Parking', 'Acceso a Casa Cádiz por Córdoba', 'Centralización Telefónica Garaje', 'Casetones cuadro de luz Piscinas', 'Caseta Piscinas', 'Cuadro Piscina Nuevo', 'Puerta de Emergencia',
                        'Arqueta Zona Pozo','Mastiles','Maestra Candados', 'Caseta Transformador Garaje', 'Central de Incendios', 'llave Ascensores','Cuarto de Luz Avenida Princial', 'Cajetines Ascensores', 'Candado Cuadro Patio', 'Armario Registro Eléctrico Casas Córdoba'])

# Ordenar cada lista por separado
llaves_edificios.sort()
llaves_generales.sort()

# Concatenar las listas para LLAVES_VALIDAS
LLAVES_VALIDAS = llaves_edificios + llaves_generales


# Nombres de las columnas para las tablas y el DataFrame
COLUMNAS_REGISTRO = [
    'Conserje', 'Llave', 'Hora de Entrega',
    'Hora de Devolución', 'Motivo de la Entrega', 'Nombre de la Persona',
    'Fecha de Entrega'
]

# --- Constantes para Comboboxes de Hora y Fecha ---
HORAS_VALIDAS = [str(i).zfill(2) for i in range(24)] # 00 a 23
MINUTOS_VALIDOS = [str(i).zfill(2) for i in range(60)] # 00 a 59
DIAS_VALIDOS = [str(i).zfill(2) for i in range(1, 32)] # 01 a 31 (Máximo, se ajustará dinámicamente)

# Meses válidos numéricos (para la lógica interna)
MESES_NUMERICOS_VALIDOS = [str(i).zfill(2) for i in range(1, 13)] # 01 a 12

# Mapeo de números de mes a nombres de mes (para la visualización)
MESES_MAP_NUMERO_A_NOMBRE = {
    '01': 'Enero', '02': 'Febrero', '03': 'Marzo', '04': 'Abril',
    '05': 'Mayo', '06': 'Junio', '07': 'Julio', '08': 'Agosto',
    '09': 'Septiembre', '10': 'Octubre', '11': 'Noviembre', '12': 'Diciembre'
}
# Mapeo inverso de nombres de mes a números de mes
MESES_MAP_NOMBRE_A_NUMERO = {v: k for k, v in MESES_MAP_NUMERO_A_NOMBRE.items()}

# Lista de nombres de meses para mostrar en el combobox de la interfaz
MESES_VALIDOS_DISPLAY = [MESES_MAP_NUMERO_A_NOMBRE[num] for num in MESES_NUMERICOS_VALIDOS]

CURRENT_YEAR = datetime.now().year
# Años para mostrar en el combobox, usaremos el año completo
ANOS_VALIDOS = [str(i) for i in range(CURRENT_YEAR - 7, CURRENT_YEAR + 7)]


# --- Funciones de Datos (Back-end) ---

def realizar_backup_automatico(nombre_archivo_excel, carpeta_backups, dias_entre_backups):    

    # Realiza un backup automático del archivo Excel si han pasado los días especificados desde el último backup.    
    if not os.path.exists(carpeta_backups):
        os.makedirs(carpeta_backups)
        print(f"Carpeta de backups '{carpeta_backups}' creada.")

    # Buscar el último archivo de backup
    archivos_backup = [f for f in os.listdir(carpeta_backups) if f.startswith('llaves_backup_') and f.endswith('.xlsx')]
    
    ultimo_backup_fecha = None
    if archivos_backup:
        # Ordenar por fecha para encontrar el más reciente
        archivos_backup.sort(key=lambda x: datetime.strptime(x, 'llaves_backup_%d%m%Y_%H%M%S.xlsx'), reverse=True)
        ultimo_backup_nombre = archivos_backup[0]
        try:
            # Extraer la fecha y hora del nombre del archivo
            ultimo_backup_fecha = datetime.strptime(ultimo_backup_nombre, 'llaves_backup_%d%m%Y_%H%M%S.xlsx')
        except ValueError:
            # Si el formato del nombre del archivo no coincide, ignorar y proceder como si no hubiera backup
            ultimo_backup_fecha = None

    necesita_backup = False
    if ultimo_backup_fecha is None:
        #print("No se encontró un backup válido o con formato correcto. Se realizará uno ahora.")
        necesita_backup = True
    else:
        dias_pasados = (datetime.now() - ultimo_backup_fecha).days
        if dias_pasados >= dias_entre_backups:
            necesita_backup = True
            print(f"Han pasado {dias_pasados} días desde el último backup. Se realizará uno ahora.")
        else:
            print(f"Último backup fue hace {dias_pasados} días. No se necesita backup aún.")

    if necesita_backup:
        try:
            timestamp = datetime.now().strftime("%d%m%Y_%H%M%S")
            nombre_backup = f"llaves_backup_{timestamp}.xlsx"
            ruta_backup = os.path.join(carpeta_backups, nombre_backup)

            if os.path.exists(nombre_archivo_excel):
                shutil.copy2(nombre_archivo_excel, ruta_backup)
                print(f"Backup creado exitosamente: {ruta_backup}")
            else:
                print(f"Error: El archivo original '{nombre_archivo_excel}' no existe para hacer el backup.")
        except Exception as e:
            print(f"Error al realizar el backup: {e}")

def inicializar_excel():
    """
    Crea el archivo Excel con las columnas si no existe.
    """
    if not os.path.exists(NOMBRE_ARCHIVO_EXCEL):
        df = pd.DataFrame(columns=COLUMNAS_REGISTRO)
        df.to_excel(NOMBRE_ARCHIVO_EXCEL, index=False)

def validar_fecha(fecha_str):
    """
    Valida que la fecha tenga el formato dd/mm/aa o dd/mm/aaaa.
    Esta función es para validar el string 'dd/mm/yy' o 'dd/mm/yyyy'.
    (La validación de fechas inexistentes como '31 de febrero' se maneja
     directamente al intentar crear un objeto datetime).
    """
    if not fecha_str: return False
    try:
        # Intenta con formato dd/mm/yy
        datetime.strptime(fecha_str, "%d/%m/%y")
        return True
    except ValueError:
        try:
            # Si falla, intenta con dd/mm/yyyy
            datetime.strptime(fecha_str, "%d/%m/%Y")
            return True
        except ValueError:
            return False

def obtener_todos_los_registros():
    """
    Carga y devuelve todos los registros del archivo Excel.
    """
    try:
        inicializar_excel()
        # Intentar leer el Excel, forzando 'Hora de Entrega' y 'Hora de Devolución' como objetos string
        # ya que pueden contener texto mixto (horas y descripciones).
        df = pd.read_excel(NOMBRE_ARCHIVO_EXCEL, dtype={
            'Hora de Entrega': str,
            'Hora de Devolución': str
        })
        # Asegurarse de que las columnas están en el orden correcto y no faltan
        for col in COLUMNAS_REGISTRO:
            if col not in df.columns:
                df[col] = ''  # Añadir columnas faltantes como vacías
        df = df[COLUMNAS_REGISTRO]  # Reordenar columnas

        # Rellenar cualquier NaN con una cadena vacía y también la cadena "nan"
        df['Hora de Devolución'] = df['Hora de Devolución'].fillna('')
        df['Hora de Devolución'] = df['Hora de Devolución'].replace('nan', '', regex=False)

        return df
    except FileNotFoundError:
        return pd.DataFrame(columns=COLUMNAS_REGISTRO)
    except Exception as e:
        messagebox.showerror("Error de lectura", f"No se pudo leer el archivo de registro: {e}")
        return pd.DataFrame(columns=COLUMNAS_REGISTRO)

def guardar_multiples_registros(nuevos_registros_df):
    """
    Guarda múltiples nuevas entradas de datos en el archivo Excel.
    'nuevos_registros_df' es un DataFrame de pandas.
    """
    try:
        inicializar_excel()
        df_existente = obtener_todos_los_registros() # Usar la función que maneja dtypes
        df_final = pd.concat([df_existente, nuevos_registros_df], ignore_index=True)
        df_final.to_excel(NOMBRE_ARCHIVO_EXCEL, index=False)
        return True
    except Exception as e:
        messagebox.showerror("Error al guardar", f"No se pudo guardar los registros en el archivo Excel: {e}")
        return False

def actualizar_registro(old_data_pandas_index, new_data):
    """
    Actualiza un registro específico en el archivo Excel.
    old_data_pandas_index: El índice original del registro a modificar en el DataFrame completo.
    new_data: Un diccionario con los datos actualizados para el registro.
    """
    try:
        df = obtener_todos_los_registros() # Cargar el DataFrame manteniendo los dtypes

        if isinstance(old_data_pandas_index, str):
            old_data_pandas_index = int(old_data_pandas_index)

        if old_data_pandas_index in df.index:
            for col, value in new_data.items():
                df.loc[old_data_pandas_index, col] = value
            df.to_excel(NOMBRE_ARCHIVO_EXCEL, index=False)
            return True
        else:
            messagebox.showwarning("Error de modificación", "Índice de registro no válido o no encontrado para modificar.")
            return False

    except Exception as e:
        messagebox.showerror("Error al actualizar", f"No se pudo actualizar el registro: {e}")
        return False

def _cargar_conserjes_desde_txt():
    """
    Carga los nombres de los conserjes desde un archivo de texto.
    Cada línea del archivo es un nombre de conserje.
    """
    conserjes = []
    try:
        with open(NOMBRE_ARCHIVO_CONSERJES, 'r', encoding='utf-8') as f:
            for line in f:
                name = line.strip()
                if name:  # Solo añadir si la línea no está vacía
                    conserjes.append(name)
    except FileNotFoundError:
        print(f"Advertencia: Archivo '{NOMBRE_ARCHIVO_CONSERJES}' no encontrado. No se cargarán conserjes predefinidos.")
        messagebox.showwarning(
            "Archivo no encontrado",
            f"El archivo '{NOMBRE_ARCHIVO_CONSERJES}' no se encontró.\n"
            "Los nombres de los conserjes deberán introducirse manualmente."
        )
    except Exception as e:
        messagebox.showerror("Error de lectura", f"Error al leer el archivo de conserjes: {e}")
    return conserjes

# --- Clases de la Interfaz Gráfica (Front-end) ---

class AppRegistroLlaves:
    def __init__(self, root):
        self.root = root
        self.root.title("Gestión de Llaves de Apartamentos")
        self.root.geometry("1200x800")
        self.root.state('zoomed')

        self.root.configure(bg="lightgray")

        try:
            self.root.iconbitmap('llave.ico')
        except tk.TclError:
            print("Advertencia: No se pudo cargar el archivo 'llave.ico'. Usando el icono predeterminado.")

        # Realizar la limpieza de archivos CSV antiguos al inicio
        self._cleanup_old_csv_files()

        # Cargar la lista de conserjes al inicio de la aplicación
        self.conserjes_list = _cargar_conserjes_desde_txt()


        style = ttk.Style()
        style.theme_use("clam")
        
        style.configure("Treeview",
                        foreground="#000000",
                        rowheight=28,
                        bordercolor="#A0A0A0",
                        borderwidth=2,
                        relief="solid"
                       )
        
        style.map("Treeview",
                  background=[('selected', '#347083')],
                  foreground=[('selected', '#FFFFFF')])

        style.configure("Treeview.Heading",
                        font=("Helvetica", 10, "bold"),
                        background="#D0D0D0",
                        foreground="#000000",
                        relief="raised")
        style.map("Treeview.Heading",
                  background=[('active', '#C0C0C0')])

        style.configure("Accent.TButton",
                        font=("Helvetica", 10, "bold"),
                        background="#007BFF",
                        foreground="white",
                        relief="raised",
                        borderwidth=2
                       )
        style.map("Accent.TButton",
                  background=[('active', '#0056B3')],
                  foreground=[('disabled', 'gray')])


        self.notebook = ttk.Notebook(root)
        self.notebook.pack(pady=10, expand=True, fill="both")

        self.pagina_registro = ttk.Frame(self.notebook)
        self.pagina_buscador = ttk.Frame(self.notebook)

        self.notebook.add(self.pagina_registro, text="Agregar Registros")
        self.notebook.add(self.pagina_buscador, text="Buscador de Llaves")

        # Inicialización de todas las variables StringVar de la clase
        self._last_search_llave = tk.StringVar(value='')
        self._last_search_año = tk.StringVar(value='')
        self._last_search_mes = tk.StringVar(value='') # Variable para el último mes buscado (numérico)
        self._last_search_dia = tk.StringVar(value='') # Variable para el último día buscado
        
        self.search_var_llave = tk.StringVar(value='')
        self.search_var_ano = tk.StringVar(value='')
        self.search_var_mes = tk.StringVar(value='') # Almacenará el valor numérico del mes seleccionado (01-12)
        self.search_var_dia = tk.StringVar(value='') # Variable para el combobox de día en el buscador
        self.search_mes_display_var = tk.StringVar(value='') # NUEVO: Almacenará el nombre del mes para la UI

        # Variable para controlar el orden actual de la columna
        self._orden_actual = {}

        # Variables para controlar el editor activo en la tabla de registro
        self._editor_widget = None
        self._active_editor_row_id = None
        self._active_editor_col_index = None

        self._crear_pagina_registro()
        self._crear_pagina_buscador()

        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_change)
        
        info_frame = ttk.Frame(self.root, style="Footer.TFrame")
        info_frame.pack(side=tk.BOTTOM, fill="x", padx=0, pady=0)
        
        style.configure("Footer.TFrame", background="lightgray")

        info_text = f"Acerca de: Gestión de Llaves v{VERSION_PROGRAMA} | © {ANIO_COPY0RIGHT} {AUTOR_PROGRAMA}"
        ttk.Label(info_frame, text=info_text, font=("Helvetica", 8, "italic"),
                  anchor="center", background="lightgray").pack(fill="x", pady=5)


    def _cleanup_old_csv_files(self):
        """
        Busca y elimina archivos CSV temporales generados por la función de impresión.
        Se ejecuta al iniciar la aplicación.
        """
        current_dir = os.getcwd()
        # El patrón ahora considera que el sufijo de filename puede ser cualquier cosa después del PREFIJO_CSV_TEMPORAL
        pattern = os.path.join(current_dir, f"{PREFIJO_CSV_TEMPORAL}*.csv")
        
        files_to_delete = glob.glob(pattern)
        
        if files_to_delete:
            print(f"Limpiando {len(files_to_delete)} archivos CSV temporales antiguos...")
            for file_path in files_to_delete:
                try:
                    os.remove(file_path)
                    print(f"Eliminado: {os.path.basename(file_path)}")
                except OSError as e:
                    print(f"Error al eliminar {os.path.basename(file_path)}: {e}")
                    # Podría ser que el archivo esté aún abierto por Excel si la app se cerró abruptamente.
                    # No es crítico si no se puede eliminar inmediatamente, se intentará de nuevo la próxima vez.
            print("Limpieza de archivos CSV temporales completada.")

    def _on_tab_change(self, event):
        selected_tab = self.notebook.tab(self.notebook.select(), "text")
        if selected_tab == "Buscador de Llaves":
            self._set_default_search_year() # Establece el año actual por defecto al cambiar de pestaña
            self._on_mes_seleccionado() # Actualiza el estado del combobox de día
            self._re_ejecutar_ultima_busqueda()
            self._cargar_años_disponibles() # Asegura que la lista de años esté actualizada
        elif selected_tab == "Agregar Registros":
            self._check_entry_tree_for_data() # Recargar datos al volver a la pestaña de registro

    def buscar_registros(self, llave=None, año=None, mes=None, dia=None, pendientes_devolucion=False):
        """
        Busca registros en el DataFrame según los criterios proporcionados
        (llave, año, mes, día y/o pendientes de devolución).
        """
        df = obtener_todos_los_registros()
        if df.empty:
            return pd.DataFrame()

        resultados = df.copy()

        if llave:
            resultados = resultados[resultados['Llave'].astype(str).str.contains(llave, case=False, na=False)]

        # Función auxiliar para extraer año, mes o día de la fecha
        def get_date_part_from_string(date_str, part_type):
            try:
                if isinstance(date_str, str):
                    if len(date_str.split('/')[2]) == 2:
                        dt_obj = datetime.strptime(date_str, "%d/%m/%y")
                    elif len(date_str.split('/')[2]) == 4:
                        dt_obj = datetime.strptime(date_str, "%d/%m/%Y")
                    else:
                        return None
                    
                    if part_type == 'year':
                        return dt_obj.year
                    elif part_type == 'month':
                        return dt_obj.month
                    elif part_type == 'day':
                        return dt_obj.day
                return None
            except (ValueError, IndexError):
                return None

        if año:
            resultados['Año_Extr'] = resultados['Fecha de Entrega'].apply(lambda x: get_date_part_from_string(x, 'year'))
            # Asegurarse de que el año a comparar también sea int
            resultados = resultados[resultados['Año_Extr'] == (int(año) if año.isdigit() else None)]
            resultados = resultados.drop(columns=['Año_Extr'])

        if mes: # Mes ya viene como número (01-12) desde self.search_var_mes
            resultados['Mes_Extr'] = resultados['Fecha de Entrega'].apply(lambda x: get_date_part_from_string(x, 'month'))
            # Asegurarse de que el mes a comparar también sea int
            resultados = resultados[resultados['Mes_Extr'] == (int(mes) if mes.isdigit() else None)]
            resultados = resultados.drop(columns=['Mes_Extr'])

        if dia:
            resultados['Dia_Extr'] = resultados['Fecha de Entrega'].apply(lambda x: get_date_part_from_string(x, 'day'))
            # Asegurarse de que el día a comparar también sea int
            resultados = resultados[resultados['Dia_Extr'] == (int(dia) if dia.isdigit() else None)]
            resultados = resultados.drop(columns=['Dia_Extr'])

        # Filtrar por "pendientes de devolución" si el flag está activado
        if pendientes_devolucion:
            resultados = resultados[
                (resultados['Hora de Devolución'].isna()) | 
                (resultados['Hora de Devolución'] == '') | 
                (resultados['Hora de Devolución'] == 'nan')
            ]

        return resultados

    def _save_time_from_popup(self, row_id, col_index, hour_val, minute_val, popup_window):
        """Guarda la hora seleccionada desde el popup en la celda del Treeview."""
        new_time = f"{hour_val}:{minute_val}"
        self._update_treeview_cell_value(row_id, col_index, new_time)
        popup_window.destroy()
        # After saving from popup, ensure active editor state is reset to allow next edit
        self._editor_widget = None
        self._active_editor_col_index = None
        self._active_editor_row_id = None # Reiniciar también el row_id

    def _save_date_from_popup(self, row_id, col_index, day_val, month_val, year_val, popup_window):
        """Guarda la fecha seleccionada desde el popup en la celda del Treeview, validando que sea una fecha válida."""
        try:
            # Intenta crear un objeto datetime para validar que la fecha sea real (ej. no 31 de febrero)
            temp_date = datetime(int(year_val), int(month_val), int(day_val))
            new_date_str = temp_date.strftime("%d/%m/%y") # Usar %y para año de dos dígitos según tu formato

            self._update_treeview_cell_value(row_id, col_index, new_date_str)
            popup_window.destroy()
            # After saving from popup, ensure active editor state is reset to allow next edit
            self._editor_widget = None
            self._active_editor_col_index = None
            self._active_editor_row_id = None # Reiniciar también el row_id
        except ValueError:
            messagebox.showwarning("Fecha Inválida", "La fecha seleccionada no es válida (ej. 31 de febrero). Por favor, revisa día, mes y año.")

    def _crear_pagina_registro(self):
        ttk.Label(self.pagina_registro, text="Agregar nuevos registros de Llaves de Apartamentos",
                  font=("Helvetica", 16, "bold")).pack(pady=10)

        input_table_frame = ttk.Frame(self.pagina_registro)
        input_table_frame.pack(pady=10, padx=10, expand=True, fill="both")

        self.entry_tree = ttk.Treeview(input_table_frame, columns=COLUMNAS_REGISTRO, show='headings',
                                       selectmode='browse')

        for col in COLUMNAS_REGISTRO:
            self.entry_tree.heading(col, text=col, anchor=tk.W)
            self.entry_tree.column(col, width=120, stretch=tk.YES)

        self.entry_tree.pack(side=tk.LEFT, fill="both", expand=True)

        yscrollbar = ttk.Scrollbar(input_table_frame, orient="vertical", command=self.entry_tree.yview)
        yscrollbar.pack(side=tk.RIGHT, fill="y")
        self.entry_tree.configure(yscrollcommand=yscrollbar.set)

        # Bind para un solo clic para activar edición
        self.entry_tree.bind("<Button-1>", self._on_entry_tree_click)
        # Bind para navegación por teclado en el Treeview cuando NO hay un editor activo
        self.entry_tree.bind("<Return>", lambda e: self._move_and_edit_next_cell(event=e, direction='right'))
        self.entry_tree.bind("<Tab>", lambda e: self._move_and_edit_next_cell(event=e, direction='right'))
        self.entry_tree.bind("<Right>", lambda e: self._move_and_edit_next_cell(event=e, direction='right'))
        self.entry_tree.bind("<Left>", lambda e: self._move_and_edit_next_cell(event=e, direction='left'))
        self.entry_tree.bind("<Down>", lambda e: self._move_and_edit_next_cell(event=e, direction='down'))
        self.entry_tree.bind("<Up>", lambda e: self._move_and_edit_next_cell(event=e, direction='up'))


        self.entry_tree.tag_configure("evenrow", background="#f0f0f0")
        self.entry_tree.tag_configure("oddrow", background="#ffffff")


        button_frame = ttk.Frame(self.pagina_registro)
        button_frame.pack(pady=(10, 0), fill="x")

        self.btn_agregar_registros = ttk.Button(button_frame, text="Agregar Registros",
                                                command=self._guardar_multiples_registros_gui, state=tk.DISABLED, style="Accent.TButton")
        self.btn_agregar_registros.pack(side=tk.RIGHT, padx=5)

        ttk.Button(button_frame, text="Añadir Fila Vacía", command=lambda: self._anadir_fila_vacia(1)).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Duplicar Fila Seleccionada", command=self._duplicar_fila_seleccionada).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Eliminar Fila Seleccionada", command=self._eliminar_fila_seleccionada).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Limpiar Todas las Filas", command=self._confirmar_limpiar_todas_las_filas).pack(side=tk.LEFT, padx=5)        

        self._anadir_fila_vacia(num_rows=13)

    def _confirmar_limpiar_todas_las_filas(self):
        confirm = messagebox.askyesno(
            "Confirmar Limpiar Filas",
            "¿Estás seguro de que deseas eliminar TODOS los registros de la tabla? Esta acción no se puede deshacer."
        )
        if confirm:
            self._limpiar_todas_las_filas()

    def _anadir_fila_vacia(self, num_rows=1):
        current_children_count = len(self.entry_tree.get_children())
        for i in range(num_rows):
            tag = "evenrow" if (current_children_count + i) % 2 == 0 else "oddrow"
            self.entry_tree.insert("", "end", values=["" for _ in COLUMNAS_REGISTRO], tags=(tag,))
        self._check_entry_tree_for_data()

    def _eliminar_fila_seleccionada(self):
        selected_items = self.entry_tree.selection()
        if not selected_items:
            messagebox.showwarning("Selección", "Por favor, selecciona una fila para eliminar.")
            return
        
        # Si hay un editor activo, finalízalo antes de eliminar la fila
        if self._editor_widget and self._editor_widget.winfo_exists():
            self._finalize_active_editor()

        for item in selected_items:
            self.entry_tree.delete(item)
        self._reaplicar_tags_zebra(self.entry_tree)
        self._check_entry_tree_for_data()

    def _duplicar_fila_seleccionada(self):
        """
        Duplica la fila actualmente seleccionada y la inserta justo debajo.
        """
        selected_items = self.entry_tree.selection()
        if not selected_items:
            messagebox.showwarning("Selección", "Por favor, selecciona una fila para duplicar.")
            return

        # Finalizar el editor si está activo antes de duplicar para asegurar que el valor esté guardado.
        if self._editor_widget and self._editor_widget.winfo_exists():
            self._finalize_active_editor()
            # Si al finalizar el editor hubo un error de validación y no se destruyó, no avanzamos.
            # Esto evita duplicar una fila con datos inválidos que el usuario aún debe corregir.
            if self._editor_widget and self._editor_widget.winfo_exists():
                return

        item_id = selected_items[0]
        current_values = self.entry_tree.item(item_id, 'values')

        all_children = self.entry_tree.get_children()
        try:
            # Encontrar el índice de la fila seleccionada y añadir 1 para insertar justo debajo
            insert_index = all_children.index(item_id) + 1
        except ValueError:
            # Esto no debería ocurrir si item_id proviene de get_children(),
            # pero como fallback, inserta al final.
            insert_index = len(all_children) 

        # Insertar la fila duplicada
        new_item_id = self.entry_tree.insert("", insert_index, values=current_values)

        self._reaplicar_tags_zebra(self.entry_tree)
        self._check_entry_tree_for_data()

        # Opcional: seleccionar y hacer visible la fila recién duplicada
        self.entry_tree.selection_set(new_item_id)
        self.entry_tree.focus(new_item_id)
        self.entry_tree.see(new_item_id)


    def _limpiar_todas_las_filas(self):
        """
        Limpiay resetea todas las filas de la tabla de entrada. Esta función es llamada tras confirmación.
        """
        # Finalizar el editor si está activo antes de limpiar
        if self._editor_widget and self._editor_widget.winfo_exists():
            self._finalize_active_editor()

        for item in self.entry_tree.get_children():
            self.entry_tree.delete(item)
        # Añadir de nuevo filas vacías después de limpiar
        self._anadir_fila_vacia(num_rows=13)
        self._check_entry_tree_for_data()

    def _reaplicar_tags_zebra(self, tree):
        """
        Reaplica las etiquetas de cebra a las filas del Treeview después de una eliminación.
        """
        for i, item_id in enumerate(tree.get_children()):
            tag = "evenrow" if i % 2 == 0 else "oddrow"
            tree.item(item_id, tags=(tag,))

    def _check_entry_tree_for_data(self, event=None):
        has_data = False
        for item_id in self.entry_tree.get_children():
            values = self.entry_tree.item(item_id, 'values')
            # Comprueba si algun valor es no vacio o un espacio en blanco, excepto 'Hora de Devolución'
            if any(val.strip() for i, val in enumerate(values) if COLUMNAS_REGISTRO[i] != 'Hora de Devolución'):
                has_data = True
                break

        if has_data:
            self.btn_agregar_registros.config(state=tk.NORMAL)
        else:
            self.btn_agregar_registros.config(state=tk.DISABLED)

    def _on_entry_tree_click(self, event):
        """
        Activa la edición de una celda de la tabla al hacer un solo clic.
        """
        region = self.entry_tree.identify("region", event.x, event.y)
        if region == "heading": return # No editar cabeceras

        new_row_id = self.entry_tree.identify_row(event.y)
        new_column_id = self.entry_tree.identify_column(event.x)

        if not new_row_id or not new_column_id: return # No hay fila o columna

        new_col_index = int(new_column_id[1:]) - 1

        # Si hay un editor activo y el clic es en una celda diferente, finaliza el editor anterior
        if self._editor_widget and self._editor_widget.winfo_exists():
            if (new_row_id != self._active_editor_row_id or
                new_col_index != self._active_editor_col_index):
                self._finalize_active_editor() # Esto guardará y destruirá el editor anterior
            else:
                # El clic es en la misma celda que se está editando, no hacer nada.
                return

        # Establecer la selección y el foco en la celda clickeada
        self.entry_tree.focus(new_row_id)
        self.entry_tree.selection_set(new_row_id)

        # Activar el editor para la nueva celda
        self._activate_cell_editor(new_row_id, new_col_index)

    def _activate_cell_editor(self, row_id, col_index):
        """
        Crea y posiciona el widget de edición (Entry o Combobox) en la celda especificada.
        """
        # Eliminar cualquier editor existente antes de crear uno nuevo (medida de seguridad)
        if self._editor_widget and self._editor_widget.winfo_exists():
            self._editor_widget.destroy()
            self._editor_widget = None

        current_values_in_row = list(self.entry_tree.item(row_id, 'values'))
        current_value = current_values_in_row[col_index]
        col_name = COLUMNAS_REGISTRO[col_index]
        
        x, y, width, height = self.entry_tree.bbox(row_id, col_index)
        if x is None or y is None: # La celda no es visible o no se pudo obtener su bbox
            # Intentar hacer visible la celda y reintentar obtener el bbox
            self.entry_tree.see(row_id)
            x, y, width, height = self.entry_tree.bbox(row_id, col_index)
            if x is None or y is None: # Si todavía no es visible, salir
                return 

        editor_widget = None

        # Almacenar la ubicación del editor activo
        self._active_editor_row_id = row_id
        self._active_editor_col_index = col_index
        
        # Asegurarse de que la fila esté enfocada y seleccionada (podría ser redundante si ya se hizo en _on_entry_tree_click o _move_and_edit_next_cell)
        self.entry_tree.focus(row_id) 
        self.entry_tree.selection_set(row_id) 


        if col_name == 'Conserje' or col_name == 'Llave':
            editor_widget = ttk.Combobox(self.entry_tree, values=self.conserjes_list if col_name == 'Conserje' else LLAVES_VALIDAS, state="readonly")
            editor_widget.set(current_value)
            editor_widget.bind("<<ComboboxSelected>>", lambda e, r=row_id, c=col_index: self._move_and_edit_next_cell(event=e, direction='right', current_row_id=r, current_col_index_from_event=c)) 
            editor_widget.bind("<FocusOut>", lambda e: self._finalize_active_editor())
            # Pasar row_id y col_index a _move_and_edit_next_cell para una navegación precisa
            editor_widget.bind("<Return>", lambda e, r=row_id, c=col_index: self._move_and_edit_next_cell(event=e, direction='right', current_row_id=r, current_col_index_from_event=c))
            editor_widget.bind("<Tab>", lambda e, r=row_id, c=col_index: self._move_and_edit_next_cell(event=e, direction='right', current_row_id=r, current_col_index_from_event=c))
            editor_widget.bind("<Right>", lambda e, r=row_id, c=col_index: self._move_and_edit_next_cell(event=e, direction='right', current_row_id=r, current_col_index_from_event=c))
            editor_widget.bind("<Left>", lambda e, r=row_id, c=col_index: self._move_and_edit_next_cell(event=e, direction='left', current_row_id=r, current_col_index_from_event=c))
            editor_widget.bind("<Down>", lambda e, r=row_id, c=col_index: self._move_and_edit_next_cell(event=e, direction='down', current_row_id=r, current_col_index_from_event=c))
            editor_widget.bind("<Up>", lambda e, r=row_id, c=col_index: self._move_and_edit_next_cell(event=e, direction='up', current_row_id=r, current_col_index_from_event=c))

        elif col_name == 'Hora de Entrega' or col_name == 'Fecha de Entrega':
            # Para Hora de Entrega y Fecha de Entrega, siempre abrimos el popup. No hay editor inline.
            if col_name == 'Hora de Entrega':
                self._open_hora_editor_popup(row_id, col_index, current_value)
            else: # col_name == 'Fecha de Entrega'
                self._open_fecha_editor_popup(row_id, col_index, current_value)
            
            # Limpiamos el estado del editor inline ya que el popup lo reemplaza conceptualmente
            self._editor_widget = None 
            self._active_editor_col_index = None
            self._active_editor_row_id = None # Reiniciar también el row_id
            return # Salir ya que el popup se encarga de la interacción y de llamar a _save_time_from_popup/_save_date_from_popup

        else: # Todos los demás campos (Entry)
            editor_widget = ttk.Entry(self.entry_tree)
            editor_widget.insert(0, current_value)
            editor_widget.bind("<FocusOut>", lambda e: self._finalize_active_editor()) # Al perder el foco, guarda
            # Al pulsar Enter o Tab, guarda y mueve a la siguiente celda
            # Pasar row_id y col_index a _move_and_edit_next_cell para una navegación precisa
            editor_widget.bind("<Return>", lambda e, r=row_id, c=col_index: self._move_and_edit_next_cell(event=e, direction='right', current_row_id=r, current_col_index_from_event=c))
            editor_widget.bind("<Tab>", lambda e, r=row_id, c=col_index: self._move_and_edit_next_cell(event=e, direction='right', current_row_id=r, current_col_index_from_event=c))
            # Para las flechas, mueve y activa el editor en la nueva celda
            editor_widget.bind("<Right>", lambda e, r=row_id, c=col_index: self._move_and_edit_next_cell(event=e, direction='right', current_row_id=r, current_col_index_from_event=c))
            editor_widget.bind("<Left>", lambda e, r=row_id, c=col_index: self._move_and_edit_next_cell(event=e, direction='left', current_row_id=r, current_col_index_from_event=c))
            editor_widget.bind("<Down>", lambda e, r=row_id, c=col_index: self._move_and_edit_next_cell(event=e, direction='down', current_row_id=r, current_col_index_from_event=c))
            editor_widget.bind("<Up>", lambda e, r=row_id, c=col_index: self._move_and_edit_next_cell(event=e, direction='up', current_row_id=r, current_col_index_from_event=c))
            

        if editor_widget:
            editor_widget.place(x=x, y=y, width=width, height=height)
            editor_widget.focus_set() # Esto es crucial para que el cursor aparezca
            # Seleccionar todo el texto en el Entry al activarlo
            if isinstance(editor_widget, ttk.Entry):
                editor_widget.selection_range(0, tk.END)
                editor_widget.icursor(tk.END) # Colocar el cursor al final
            elif isinstance(editor_widget, ttk.Combobox):
                # Generar un evento de clic para intentar abrir el dropdown de inmediato
                editor_widget.event_generate('<Button-1>')


            self._editor_widget = editor_widget # Almacenar el editor activo

    def _finalize_active_editor(self):
        """
        Finaliza la edición del widget activo, guardando su valor en el Treeview
        y luego destruyendo el widget. Esta es la función central de finalización.
        """
        if self._editor_widget and self._editor_widget.winfo_exists() and \
           self._active_editor_row_id is not None and self._active_editor_col_index is not None:
            
            new_value = self._editor_widget.get()
            col_name = COLUMNAS_REGISTRO[self._active_editor_col_index]

            # Actualizar la celda en el Treeview
            self._update_treeview_cell_value(
                self._active_editor_row_id,
                self._active_editor_col_index,
                new_value
            )
            
            # Destruir el widget editor después de guardar
            self._editor_widget.destroy()

        # Asegurarse de que el estado del editor se limpie
        self._editor_widget = None
        self._active_editor_col_index = None
        self._active_editor_row_id = None # Reiniciar también el row_id
        self._check_entry_tree_for_data()

    def _update_treeview_cell_value(self, row_id, col_index, new_value):
        """
        Actualiza el valor de una celda específica en el Treeview.
        Esta función solo se encarga de la actualización del modelo de datos del Treeview.
        """
        current_values = list(self.entry_tree.item(row_id, 'values'))
        current_values[col_index] = new_value
        self.entry_tree.item(row_id, values=current_values)
        self._check_entry_tree_for_data() # Revisa el estado de los datos después de una actualización

    def _parse_time_string(self, time_str):
        """AYUDA a convertir 'HH:MM' string en (hour, minute) tupla."""
        try:
            # Try to parse as time object first
            dt_obj = datetime.strptime(time_str, "%H:%M")
            return dt_obj.strftime("%H"), dt_obj.strftime("%M")
        except ValueError:
            # Si no es valida la hora, trata de convertir al formato comun flotante de Excel
            pass
                
        try:
            # Covierte float a datetime.time
            float_val = float(time_str)
            if 0 <= float_val < 1:
                total_minutes = float_val * 1440 # 24 hours * 60 minutes
                hours = int(total_minutes // 60)
                minutes = int(total_minutes % 60)
                return str(hours).zfill(2), str(minutes).zfill(2)
        except ValueError:
            pass

        return "00", "00"

    def _parse_date_string(self, date_str):
        """Ayuda a convertir 'DD/MM/YY' o 'DD/MM/YYYY' string a (day, month, full_year) tupla."""
        if isinstance(date_str, datetime): # Si ya es un objeto datetime
            return date_str.strftime("%d"), date_str.strftime("%m"), str(date_str.year)
        if not date_str: # Si está vacío, usar fecha actual
            now = datetime.now()
            return now.strftime("%d"), now.strftime("%m"), str(now.year)
        try:
            dt_obj = datetime.strptime(date_str, "%d/%m/%y")
            return dt_obj.strftime("%d"), dt_obj.strftime("%m"), str(dt_obj.year)
        except ValueError:
            try:
                dt_obj = datetime.strptime(date_str, "%d/%m/%Y")
                return dt_obj.strftime("%d"), dt_obj.strftime("%m"), str(dt_obj.year)
            except ValueError:
                now = datetime.now()
                return now.strftime("%d"), now.strftime("%m"), str(now.year)

    def _open_hora_editor_popup(self, row_id, col_index, current_value):
        popup = tk.Toplevel(self.root)
        popup.title("Seleccionar Hora")
        try:
            popup.iconbitmap('llave.ico') # Set icon for popup
        except tk.TclError:
            print("Advertencia: No se pudo cargar el archivo 'llave.ico' para el popup de hora.")
        popup.transient(self.root)
        popup.grab_set()
        popup.geometry("250x120")
        popup.resizable(False, False)

        frame = ttk.Frame(popup, padding="10")
        frame.pack(expand=True, fill="both")

        # Asegurarse de que current_value sea un string antes de intentar parsearlo
        initial_hr, initial_min = self._parse_time_string(str(current_value))

        hr_var = tk.StringVar(value=initial_hr)
        min_var = tk.StringVar(value=initial_min)

        ttk.Label(frame, text="Hora:").grid(row=0, column=0, padx=5, pady=5)
        hora_combo = ttk.Combobox(frame, textvariable=hr_var, values=HORAS_VALIDAS, state="readonly", width=5)
        hora_combo.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(frame, text="Minuto:").grid(row=0, column=2, padx=5, pady=5)
        minuto_combo = ttk.Combobox(frame, textvariable=min_var, values=MINUTOS_VALIDOS, state="readonly", width=5)
        minuto_combo.grid(row=0, column=3, padx=5, pady=5)

        button_frame = ttk.Frame(popup)
        button_frame.pack(pady=5)
        ttk.Button(button_frame, text="Guardar", command=lambda: self._save_time_from_popup(row_id, col_index, hr_var.get(), min_var.get(), popup)).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancelar", command=popup.destroy).pack(side=tk.LEFT, padx=5)

        popup.protocol("WM_DELETE_WINDOW", popup.destroy)

    def _open_fecha_editor_popup(self, row_id, col_index, current_value):
        popup = tk.Toplevel(self.root)
        popup.title("Seleccionar Fecha")
        try:
            popup.iconbitmap('llave.ico') # Set icon for popup
        except tk.TclError:
            print("Advertencia: No se pudo cargar el archivo 'llave.ico' para el popup de fecha.")
        popup.transient(self.root)
        popup.grab_set()
        popup.geometry("350x120")
        popup.resizable(False, False)

        frame = ttk.Frame(popup, padding="10")
        frame.pack(expand=True, fill="both")

        # Asegurarse de que current_value sea un string antes de intentar parsearlo
        initial_day, initial_month, initial_year = self._parse_date_string(str(current_value))

        dia_var = tk.StringVar(value=initial_day)
        mes_var = tk.StringVar(value=initial_month)
        ano_var = tk.StringVar(value=initial_year)


        ttk.Label(frame, text="Día:").grid(row=0, column=0, padx=5, pady=5)
        dia_combo = ttk.Combobox(frame, textvariable=dia_var, values=DIAS_VALIDOS, state="readonly", width=5)
        dia_combo.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(frame, text="Mes:").grid(row=0, column=2, padx=5, pady=5)
        mes_combo = ttk.Combobox(frame, textvariable=mes_var, values=MESES_NUMERICOS_VALIDOS, state="readonly", width=5)
        mes_combo.grid(row=0, column=3, padx=5, pady=5)

        ttk.Label(frame, text="Año:").grid(row=0, column=4, padx=5, pady=5)
        ano_combo = ttk.Combobox(frame, textvariable=ano_var, values=ANOS_VALIDOS, state="readonly", width=8)
        ano_combo.grid(row=0, column=5, padx=5, pady=5)

        # Vincular la actualización de días al cambiar mes o año
        mes_combo.bind("<<ComboboxSelected>>", lambda e: self._actualizar_dias_en_popup(dia_combo, mes_var, ano_var))
        ano_combo.bind("<<ComboboxSelected>>", lambda e: self._actualizar_dias_en_popup(dia_combo, mes_var, ano_var))
        # Llamar una vez para inicializar los días correctamente
        self._actualizar_dias_en_popup(dia_combo, mes_var, ano_var)

        button_frame = ttk.Frame(popup)
        button_frame.pack(pady=5)
        ttk.Button(button_frame, text="Guardar", command=lambda: self._save_date_from_popup(row_id, col_index, dia_var.get(), mes_var.get(), ano_var.get(), popup)).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancelar", command=popup.destroy).pack(side=tk.LEFT, padx=5)

        popup.protocol("WM_DELETE_WINDOW", popup.destroy)

    def _actualizar_dias_en_popup(self, dia_combo_widget, mes_var, ano_var):
        """Actualiza los días válidos en un combobox de día de un popup de fecha."""
        try:
            mes = int(mes_var.get())
            año = int(ano_var.get())
            _, num_days = calendar.monthrange(año, mes)
            valid_days = [f"{d:02d}" for d in range(1, num_days + 1)]
            dia_combo_widget['values'] = valid_days
            # Si el día actual es mayor que el número de días del mes, ajustarlo al último día del mes
            current_day = int(dia_combo_widget.get()) if dia_combo_widget.get() else 0
            if current_day > num_days:
                dia_combo_widget.set(f"{num_days:02d}")
            elif not dia_combo_widget.get() and valid_days: # Si no hay día seleccionado, poner el primer día válido
                dia_combo_widget.set(valid_days[0])
        except ValueError:
            dia_combo_widget['values'] = []
            dia_combo_widget.set('')

    def _guardar_multiples_registros_gui(self):
        # Finalizar el editor si está activo antes de guardar
        if self._editor_widget and self._editor_widget.winfo_exists():
            self._finalize_active_editor()

        registros_a_guardar = []
        filas_con_errores = []

        # Recorrer todas las filas en el Treeview
        for i, item_id in enumerate(self.entry_tree.get_children()):
            values = self.entry_tree.item(item_id, 'values')
            row_data = dict(zip(COLUMNAS_REGISTRO, values))

            # Comprobar si la fila está completamente vacía (excepto 'Hora de Devolución' que puede estar vacía)
            is_empty_row = all(row_data[k].strip() == '' for k in COLUMNAS_REGISTRO if k != 'Hora de Devolución')
            if is_empty_row:
                continue # Saltar filas completamente vacías

            is_valid_row = True
            error_message = []

            # Validar campos obligatorios (todos excepto 'Hora de Devolución')
            required_fields = [k for k in COLUMNAS_REGISTRO if k != 'Hora de Devolución']
            for field in required_fields:
                if not row_data[field].strip():
                    error_message.append(f"Campo '{field}' vacío.")
                    is_valid_row = False

            # Validar formato de fecha y existencia si el campo no está vacío
            if row_data['Fecha de Entrega'].strip():
                try:
                    # Intenta crear un objeto datetime para validar la existencia de la fecha
                    # Asumiendo que row_data['Fecha de Entrega'] está en formato dd/mm/yy o dd/mm/yyyy
                    date_parts = row_data['Fecha de Entrega'].split('/')
                    if len(date_parts[2]) == 2:
                        datetime.strptime(row_data['Fecha de Entrega'], "%d/%m/%y")
                    elif len(date_parts[2]) == 4:
                        datetime.strptime(row_data['Fecha de Entrega'], "%d/%m/%Y")
                    else:
                        error_message.append("Formato de 'Fecha de Entrega' incorrecto (dd/mm/aa o dd/mm/aaaa).")
                        is_valid_row = False
                except ValueError:
                    error_message.append("Fecha de 'Fecha de Entrega' inexistente o inválida (ej. 31/02).")
                    is_valid_row = False
            else:
                # Si la fecha de entrega está vacía y es un campo requerido, ya se habrá capturado
                pass


            if is_valid_row:
                registros_a_guardar.append(row_data)
            else:
                filas_con_errores.append(f"Fila {i+1}: {', '.join(error_message)}")

        if not registros_a_guardar and not filas_con_errores:
            messagebox.showinfo("Guardar Registros", "No hay registros para guardar (todas las filas están vacías).")
            return

        if filas_con_errores:
            error_details = "\n".join(filas_con_errores)
            messagebox.showwarning("Errores de Validación",
                                   f"Los siguientes registros tienen errores y NO se guardarán hasta que se corrijan:\n\n{error_details}\n\n"
                                   "Por favor, corrige todas las filas con errores. Ningún registro se guardará hasta que todas las filas tengan datos válidos.")
            # Si hay errores, no se guarda NADA y se retorna.
            return 
        
        # Si llegamos aquí, significa que no hay filas_con_errores y al menos hay un registro para guardar
        if registros_a_guardar: 
            try:
                # Asegurarse de que los datos del Treeview se conviertan a un DataFrame con tipos de datos consistentes.
                df_temp_nuevos_registros = pd.DataFrame(registros_a_guardar, columns=COLUMNAS_REGISTRO)
                for col in COLUMNAS_REGISTRO:
                    df_temp_nuevos_registros[col] = df_temp_nuevos_registros[col].astype(str)

                # Llamar a la función de guardado que ya tienes y funciona correctamente
                if guardar_multiples_registros(df_temp_nuevos_registros):
                    messagebox.showinfo("Guardar Registros", "Todos los registros se guardaron exitosamente.")
                    # Eliminar todas las filas, ya que todos los registros fueron guardados
                    for item_id in self.entry_tree.get_children():
                        self.entry_tree.delete(item_id)
                    
                    # Añadir nuevas filas vacías para continuar
                    current_rows = len(self.entry_tree.get_children())
                    if current_rows < 13: # Mantener un mínimo de 13 filas visibles
                        self._anadir_fila_vacia(num_rows=(13 - current_rows))

                    self._reaplicar_tags_zebra(self.entry_tree) # Reaplica los estilos de cebra
                else:
                    messagebox.showerror("Error al Guardar", "Ocurrió un error al intentar guardar los registros. Las filas no se eliminaron de la tabla.")
            except Exception as e:
                # Esto captura errores durante la preparación del DataFrame, no el guardado en sí
                messagebox.showerror("Error al Guardar", f"Ocurrió un error inesperado al preparar los datos para guardar: {e}\n{traceback.format_exc()}")
        else: 
            messagebox.showinfo("Guardar Registros", "No hay registros completos para guardar.")

        self._check_entry_tree_for_data()

    def _crear_pagina_buscador(self):
        ttk.Label(self.pagina_buscador, text="Buscador de Registros de Llaves",
                  font=("Helvetica", 16, "bold")).pack(pady=10)

        search_frame = ttk.LabelFrame(self.pagina_buscador, text="Criterios de Búsqueda", padding="10")
        search_frame.pack(pady=10, padx=10, fill="x")

        ttk.Label(search_frame, text="Llave:", font=("Helvetica", 11, "bold")).grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.search_llave_combo = ttk.Combobox(search_frame, textvariable=self.search_var_llave,
                                               values=[''] + LLAVES_VALIDAS)
        self.search_llave_combo.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        # COMBOBOX PARA AÑO, MES Y DÍA
        ttk.Label(search_frame, text="Año:", font=("Helvetica", 10, "bold")).grid(row=0, column=2, padx=10, pady=5, sticky="w")
        self.search_ano_combo = ttk.Combobox(search_frame, textvariable=self.search_var_ano,
                                             values=[''] + ANOS_VALIDOS, state="readonly", width=8) # Usar ANOS_VALIDOS directamente
        self.search_ano_combo.grid(row=0, column=3, padx=5, pady=5, sticky="ew")

        ttk.Label(search_frame, text="Mes:", font=("Helvetica", 10, "bold")).grid(row=0, column=4, padx=10, pady=5, sticky="w")
        self.search_mes_combo = ttk.Combobox(search_frame, textvariable=self.search_mes_display_var, # Usa la nueva variable para el display
                                             values=[''] + MESES_VALIDOS_DISPLAY, state="readonly", width=8) # Usa nombres de mes
        self.search_mes_combo.grid(row=0, column=5, padx=5, pady=5, sticky="ew")

        # Vincular para habilitar/deshabilitar el día y convertir a número
        self.search_mes_combo.bind("<<ComboboxSelected>>", self._on_mes_seleccionado)

        ttk.Label(search_frame, text="Día:", font=("Helvetica", 10, "bold")).grid(row=0, column=6, padx=10, pady=5, sticky="w")
        self.search_dia_combo = ttk.Combobox(search_frame, textvariable=self.search_var_dia,
                                             values=[''], state="disabled", width=8) # Inicia deshabilitado y solo con la opción vacía
        self.search_dia_combo.grid(row=0, column=7, padx=5, pady=5, sticky="ew")

        # Vincular para recalcular días válidos si el año cambia
        self.search_ano_combo.bind("<<ComboboxSelected>>", self._on_mes_seleccionado)

        # BOTONES 
        ttk.Button(search_frame, text="Buscar", command=self._ejecutar_busqueda, style="Accent.TButton").grid(row=0, column=8, padx=10, pady=5)
        ttk.Button(search_frame, text="Limpiar Búsqueda", command=self._limpiar_y_mostrar_todo).grid(row=0, column=9, padx=5, pady=5)
        ttk.Button(search_frame, text="Imprimir Tabla", command=lambda: self._imprimir_tabla_gui(self.tree_buscador, "Tabla de Búsqueda")).grid(row=0, column=10, padx=5, pady=5)


        search_frame.grid_columnconfigure(1, weight=3) # Llave
        search_frame.grid_columnconfigure(3, weight=1) # Año
        search_frame.grid_columnconfigure(5, weight=1) # Mes (NUEVO)
        search_frame.grid_columnconfigure(7, weight=1) # Día (NUEVO)

        # Ajuste para que los botones no se expandan
        for i in [0, 2, 4, 6, 8, 9, 10]: # Ajustar índices de columna
            search_frame.grid_columnconfigure(i, weight=0)


        table_frame = ttk.Frame(self.pagina_buscador)
        table_frame.pack(pady=10, padx=10, expand=True, fill="both")

        self.tree_buscador = ttk.Treeview(table_frame, columns=COLUMNAS_REGISTRO, show='headings',
                                          selectmode='browse')

        for col_name in COLUMNAS_REGISTRO:
            self.tree_buscador.heading(col_name, text=col_name, anchor=tk.W, 
                                       command=lambda c=col_name: self._ordenar_treeview_por_columna(c))
            self.tree_buscador.column(col_name, width=100, stretch=tk.YES)

        self.tree_buscador.pack(side=tk.LEFT, fill="both", expand=True)

        yscrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree_buscador.yview)
        yscrollbar.pack(side=tk.RIGHT, fill="y")
        self.tree_buscador.configure(yscrollcommand=yscrollbar.set)

        self.tree_buscador.tag_configure("evenrow", background="#f0f0f0")
        self.tree_buscador.tag_configure("oddrow", background="#ffffff")

        
        modify_button_frame = ttk.Frame(self.pagina_buscador)
        modify_button_frame.pack(pady=5)
        
        self.btn_pendientes = ttk.Button(modify_button_frame, text="Llaves Pendientes", command=self._mostrar_pendientes, style="Accent.TButton").pack(side=tk.LEFT, padx=5)
        self.btn_modificar = ttk.Button(modify_button_frame, text="Modificar Registro Seleccionado", command=self._modificar_registro_gui, state=tk.DISABLED)
        self.btn_modificar.pack()
        



        self.tree_buscador.bind("<<TreeviewSelect>>", self._on_tree_select_for_modify)

        self._set_default_search_year() # Establecer el año actual por defecto al inicio
        self._on_mes_seleccionado() # Inicializar el estado del combobox de día
        self._re_ejecutar_ultima_busqueda()
        self._cargar_años_disponibles() # Asegura que los valores de años, meses y días estén cargados

    def _set_default_search_year(self):
        """Establece el año actual como valor por defecto en el combobox de búsqueda."""
        current_year_str = str(datetime.now().year)
        if current_year_str in self.search_ano_combo['values']:
            self.search_var_ano.set(current_year_str)
        else:
            self.search_var_ano.set('') # Si por alguna razón el año actual no está en la lista

    def _on_mes_seleccionado(self, event=None):
        """
        Habilita o deshabilita el combobox de día y actualiza sus valores según la selección del mes y año.
        Ahora incluye una opción en blanco para el día por defecto.
        """
        selected_month_name = self.search_mes_display_var.get()
        if selected_month_name: # Si se ha seleccionado un nombre de mes (no la opción en blanco)
            try:
                # Convertir el nombre del mes a su número para la lógica interna
                mes_numero = MESES_MAP_NOMBRE_A_NUMERO.get(selected_month_name)
                if mes_numero is None: # Si el nombre seleccionado no se encuentra en el mapeo (ej. si se borró manualmente)
                    self.search_dia_combo.config(state="disabled")
                    self.search_var_dia.set('')
                    self.search_dia_combo['values'] = ['']
                    self.search_var_mes.set('') # Limpiar también el valor numérico interno
                    return

                self.search_var_mes.set(mes_numero) # Almacenar el valor numérico en la variable interna

                año = int(self.search_var_ano.get()) if self.search_var_ano.get() else datetime.now().year
                
                _, num_days = calendar.monthrange(año, int(mes_numero)) # Usar el número del mes para calendar
                
                new_days_validos = [f"{i:02d}" for i in range(1, num_days + 1)]
                self.search_dia_combo['values'] = [''] + new_days_validos # Añadir opción en blanco al inicio
                self.search_dia_combo.config(state="readonly")

                # Si el día actualmente seleccionado es numérico y excede el número de días del nuevo mes,
                # o si el día no es un dígito, se limpia la selección de día.
                current_selected_day = self.search_var_dia.get()
                if current_selected_day.isdigit() and int(current_selected_day) > num_days:
                    self.search_var_dia.set('') # Limpiar la selección si es inválida
                # Si el día estaba vacío, se mantiene vacío para permitir la búsqueda solo por mes.
                # No se establece un día por defecto (como '01').
                
            except ValueError:
                # Si el año no es un número válido aún, o si hay un error en monthrange
                self.search_dia_combo.config(state="disabled")
                self.search_var_dia.set('')
                self.search_dia_combo['values'] = [''] # Mantener la opción en blanco incluso deshabilitado
                self.search_var_mes.set('') # Limpiar el valor numérico interno también
        else:
            # Si no hay mes seleccionado (es la opción en blanco), el día se deshabilita y se limpia
            self.search_dia_combo.config(state="disabled")
            self.search_var_dia.set('')
            self.search_dia_combo['values'] = [''] # Mantener la opción en blanco
            self.search_var_mes.set('') # Limpiar el valor numérico interno

    def _cargar_años_disponibles(self):
        """
        Carga los años únicos de la columna 'Fecha de Entrega' del Excel y los pone en el combobox de años.
        También asegura que los comboboxes de mes y día se inicialicen con los valores válidos.
        """
        df = obtener_todos_los_registros()
        if df.empty or 'Fecha de Entrega' not in df.columns:
            self.search_ano_combo['values'] = [''] + ANOS_VALIDOS
            self.search_mes_combo['values'] = [''] + MESES_VALIDOS_DISPLAY # Usar nombres para el display
            self.search_dia_combo['values'] = [''] # Iniciar el día solo con la opción en blanco
            return

        años = set()
        for fecha_str in df['Fecha de Entrega'].dropna().astype(str):
            try:
                # Intentar parsear como dd/mm/yy o dd/mm/yyyy
                if isinstance(fecha_str, str):
                    parts = fecha_str.split('/')
                    if len(parts[2]) == 2:
                        año_completo = datetime.strptime(fecha_str, "%d/%m/%y").year
                    elif len(parts[2]) == 4:
                        año_completo = datetime.strptime(fecha_str, "%d/%m/%Y").year
                    else:
                        continue
                else:
                    continue
                años.add(str(año_completo))
            except (ValueError, IndexError):
                continue # Ignorar fechas con formato incorrecto

        sorted_años = sorted(list(años.union(set(ANOS_VALIDOS))), reverse=True) # Combinar años de registros con ANOS_VALIDOS
        self.search_ano_combo['values'] = [''] + sorted_años
        
        # Restaurar la última búsqueda de año si el valor todavía es válido
        if self._last_search_año.get() in self.search_ano_combo['values']:
             self.search_var_ano.set(self._last_search_año.get())
        else:
             self.search_var_ano.set(str(datetime.now().year)) # Establecer año actual si no hay valor o es inválido
        
        # Asegurarse de que los comboboxes de mes y día siempre tengan los valores estáticos iniciales
        self.search_mes_combo['values'] = [''] + MESES_VALIDOS_DISPLAY # Usar nombres para el display
        # Restaurar el último mes buscado (numérico) y convertirlo a nombre para el display
        mes_num_val = self._last_search_mes.get() # Obtener el valor numérico del mes
        if mes_num_val and mes_num_val in MESES_MAP_NUMERO_A_NOMBRE:
            self.search_mes_display_var.set(MESES_MAP_NUMERO_A_NOMBRE[mes_num_val])
            self.search_var_mes.set(mes_num_val) # Mantener el valor numérico interno
        else:
            self.search_mes_display_var.set('')
            self.search_var_mes.set('') # Asegurar que el valor numérico interno también se borre

        # Los valores del día se actualizarán al seleccionar el mes.
        # Restaura el último día buscado (si existe y el mes está seleccionado)
        if self._last_search_dia.get() and self.search_var_mes.get():
             self.search_var_dia.set(self._last_search_dia.get())
        else:
             self.search_var_dia.set('')
        self._on_mes_seleccionado() # Llamar para inicializar el combobox de día correctamente

    def _on_tree_select_for_modify(self, event=None):
        selected_items = self.tree_buscador.selection()
        if len(selected_items) == 1:
            self.btn_modificar.config(state=tk.NORMAL)
        else:
            self.btn_modificar.config(state=tk.DISABLED)

    def _modificar_registro_gui(self):
        selected_item_id = self.tree_buscador.selection()
        if not selected_item_id:
            messagebox.showwarning("Modificar", "Por favor, selecciona un registro para modificar.")
            return

        item_id = selected_item_id[0]
        # Cuando se insertan items en tree_buscador, se usa el índice de pandas como id.
        # Aseguramos que el original_pandas_index sea el entero.
        original_pandas_index = int(item_id)
        current_values = self.tree_buscador.item(item_id, 'values')

        self._abrir_ventana_modificar(original_pandas_index, current_values)

    def _abrir_ventana_modificar(self, pandas_index_to_modify, current_record_values):
        modify_window = tk.Toplevel(self.root)
        modify_window.title("Modificar Registro de Llave")
        try:
            modify_window.iconbitmap('llave.ico')
        except tk.TclError:
            print("Advertencia: No se pudo cargar el archivo 'llave.ico' para la ventana de modificación.")

        modify_window.geometry("840x360") # Aumentado el ancho de la ventana
        modify_window.transient(self.root)
        modify_window.grab_set()

        modify_vars = {}
        # Referencias a los comboboxes de día, mes, año dentro de la ventana de modificación
        self.mod_dia_combo = None 
        self.mod_mes_combo = None
        self.mod_ano_combo = None

        for i, col in enumerate(COLUMNAS_REGISTRO):
            # Asegurarse de que el valor inicial sea un string para los StringVars
            modify_vars[col] = tk.StringVar(value=str(current_record_values[i]))

        input_frame = ttk.LabelFrame(modify_window, text="Datos del Registro", padding="5")
        input_frame.pack(padx=10, pady=10, fill="both", expand=True)

        row_idx = 0
        for i, col in enumerate(COLUMNAS_REGISTRO):
            ttk.Label(input_frame, text=f"{col}:").grid(row=row_idx, column=0, sticky="w", pady=2, padx=5)

            if col == 'Conserje':
                combo = ttk.Combobox(input_frame, textvariable=modify_vars[col], values=self.conserjes_list, state="readonly")
                combo.grid(row=row_idx, column=1, sticky="ew", pady=2, padx=5)
            elif col == 'Llave':
                combo = ttk.Combobox(input_frame, textvariable=modify_vars[col], values=LLAVES_VALIDAS, state="readonly")
                combo.grid(row=row_idx, column=1, sticky="ew", pady=2, padx=5)
            elif col == 'Hora de Entrega': # Sólo Hora de Entrega usa el selector de hora
                # Convertir a string para _parse_time_string para evitar errores si es datetime.time o float
                initial_hr, initial_min = self._parse_time_string(str(current_record_values[i]))
                hr_mod_var = tk.StringVar(value=initial_hr)
                min_mod_var = tk.StringVar(value=initial_min)
                
                hora_frame = ttk.Frame(input_frame)
                hora_frame.grid(row=row_idx, column=1, sticky="ew", pady=2, padx=5)
                # Crear comboboxes con los valores parseados, pero el textvariable se asigna después
                # directamente a modify_vars para poder recuperar el valor combinado
                ttk.Combobox(hora_frame, textvariable=hr_mod_var, values=HORAS_VALIDAS, state="readonly", width=5).pack(side=tk.LEFT, padx=2)
                ttk.Label(hora_frame, text=":").pack(side=tk.LEFT)
                ttk.Combobox(hora_frame, textvariable=min_mod_var, values=MINUTOS_VALIDOS, state="readonly", width=5).pack(side=tk.LEFT, padx=2)
                
                # Almacenar las variables individuales para recuperar el valor combinado más tarde
                modify_vars[col] = (hr_mod_var, min_mod_var)
            elif col == 'Fecha de Entrega':
                # Convertir a string para _parse_date_string para evitar errores si es datetime.date o float
                initial_day, initial_month, initial_year = self._parse_date_string(str(current_record_values[i]))
                dia_mod_var = tk.StringVar(value=initial_day)
                mes_mod_var = tk.StringVar(value=initial_month)
                ano_mod_var = tk.StringVar(value=initial_year)
                
                fecha_frame = ttk.Frame(input_frame)
                fecha_frame.grid(row=row_idx, column=1, sticky="ew", pady=2, padx=5)
                
                self.mod_dia_combo = ttk.Combobox(fecha_frame, textvariable=dia_mod_var, values=DIAS_VALIDOS, state="readonly", width=5) # Store for later use
                self.mod_dia_combo.pack(side=tk.LEFT, padx=2)
                ttk.Label(fecha_frame, text="/").pack(side=tk.LEFT)
                self.mod_mes_combo = ttk.Combobox(fecha_frame, textvariable=mes_mod_var, values=MESES_NUMERICOS_VALIDOS, state="readonly", width=5) 
                self.mod_mes_combo.pack(side=tk.LEFT, padx=2)
                ttk.Label(fecha_frame, text="/").pack(side=tk.LEFT)
                self.mod_ano_combo = ttk.Combobox(fecha_frame, textvariable=ano_mod_var, values=ANOS_VALIDOS, state="readonly", width=8) # Store for later use
                self.mod_ano_combo.pack(side=tk.LEFT, padx=2)
                
                # Bind para actualizar los días al cambiar mes o año
                self.mod_mes_combo.bind("<<ComboboxSelected>>", lambda event: self._actualizar_dias_en_popup(self.mod_dia_combo, mes_mod_var, ano_mod_var))
                self.mod_ano_combo.bind("<<ComboboxSelected>>", lambda event: self._actualizar_dias_en_popup(self.mod_dia_combo, mes_mod_var, ano_mod_var))
                self._actualizar_dias_en_popup(self.mod_dia_combo, mes_mod_var, ano_mod_var) # Llamar para inicializar los días correctamente
                
                # Almacenar las variables individuales para recuperar el valor combinado más tarde
                modify_vars[col] = (dia_mod_var, mes_mod_var, ano_mod_var)
            else: # Todos los demás campos, incluida 'Hora de Devolución', serán Entry
                entry = ttk.Entry(input_frame, textvariable=modify_vars[col])
                entry.grid(row=row_idx, column=1, sticky="ew", pady=2, padx=5)

            row_idx += 1

        input_frame.grid_columnconfigure(1, weight=1)

        def _guardar_modificacion():
            new_data = {}
            for col in COLUMNAS_REGISTRO:
                if col == 'Hora de Entrega': # Sólo Hora de Entrega usa la tupla de variables de hora/minuto
                    # Si la variable de modificación es una tupla (comboboxes de hora/minuto)
                    if isinstance(modify_vars.get(col), tuple):
                        hr_var, min_var = modify_vars[col]
                        new_value = f"{hr_var.get()}:{min_var.get()}"
                    else: # Si no, es un StringVar normal de entrada de texto
                        new_value = modify_vars[col].get()
                    
                    new_data[col] = new_value.strip() # Almacenar como string, quitando espacios
                elif col == 'Fecha de Entrega':
                    dia_var, mes_var, ano_var = modify_vars[col]
                    selected_day = dia_var.get()
                    selected_month = mes_var.get()
                    selected_year_full = ano_var.get()

                    if selected_day and selected_month and selected_year_full:
                        try:
                            # Valida la fecha antes de guardar
                            temp_date = datetime(int(selected_year_full), int(selected_month), int(selected_day))
                            new_date_str = temp_date.strftime("%d/%m/%y")
                            new_data[col] = new_date_str
                        except ValueError:
                            messagebox.showwarning("Formato de Fecha", "La fecha seleccionada no es válida (ej. 31 de febrero). El formato debe ser dd/mm/aa.")
                            return # Detener el proceso de guardado si la fecha es inválida
                    else:
                        new_data[col] = "" # Guardar como vacío si alguna parte está vacía
                else: # Todos los demás campos, incluida 'Hora de Devolución', se recuperan como StringVar
                    new_data[col] = modify_vars[col].get().strip()


            for field in [c for c in COLUMNAS_REGISTRO if c != 'Hora de Devolución']:
                if not new_data[field].strip():
                    messagebox.showwarning("Campos obligatorios", f"El campo '{field}' no puede estar vacío.")
                    return

            llave_val = new_data['Llave']
            if llave_val.strip() and llave_val not in LLAVES_VALIDAS:
                messagebox.showwarning("Validación de Llave", f"Llave '{llave_val}' no válida.")
                return

            if actualizar_registro(pandas_index_to_modify, new_data):
                messagebox.showinfo("Éxito", "Registro modificado correctamente.")
                modify_window.destroy()
                self._re_ejecutar_ultima_busqueda() # Re-ejecutar la última búsqueda para actualizar la tabla
            else:
                messagebox.showerror("Error", "Fallo al modificar el registro.")

        button_frame = ttk.Frame(modify_window)
        button_frame.pack(pady=10)
        ttk.Button(button_frame, text="Guardar Cambios", command=_guardar_modificacion).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancelar", command=modify_window.destroy).pack(side=tk.LEFT, padx=5)

        modify_window.protocol("WM_DELETE_WINDOW", modify_window.destroy)

    def _ejecutar_busqueda(self):
        llave_busqueda = self.search_var_llave.get().strip()
        año_busqueda = self.search_var_ano.get().strip()
        mes_busqueda = self.search_var_mes.get().strip() # Se usa el valor numérico interno
        dia_busqueda = self.search_var_dia.get().strip() 
        

        self._last_search_llave.set(llave_busqueda)
        self._last_search_año.set(año_busqueda)
        self._last_search_mes.set(mes_busqueda) # Guardar último mes buscado (numérico)
        self._last_search_dia.set(dia_busqueda) # Guardar último día buscado

        if not llave_busqueda and not año_busqueda and not mes_busqueda and not dia_busqueda: # Actualizado
            self._limpiar_y_mostrar_todo(should_clear_last_search=False)
            return

        resultados_df = self.buscar_registros(llave=llave_busqueda, año=año_busqueda, mes=mes_busqueda, dia=dia_busqueda) # Actualizado
        self._mostrar_resultados_en_treeview(resultados_df)

    def _mostrar_pendientes(self):        
        # Busqueda de registros con 'Hora de Devolución' vacía.

        llave_busqueda = self.search_var_llave.get().strip()
        año_busqueda = self.search_var_ano.get().strip()
        mes_busqueda = self.search_var_mes.get().strip() # Se usa el valor numérico interno
        dia_busqueda = self.search_var_dia.get().strip() 
    

        self._last_search_llave.set(llave_busqueda)
        self._last_search_año.set(año_busqueda)
        self._last_search_mes.set(mes_busqueda) # Guardar último mes buscado (numérico)
        self._last_search_dia.set(dia_busqueda) # Guardar último día buscado

        
        resultados_df_raw = self.buscar_registros(
            llave=llave_busqueda if llave_busqueda else None,
            año=año_busqueda if año_busqueda else None,
            mes=mes_busqueda if mes_busqueda else None,
            dia=dia_busqueda if dia_busqueda else None
        )

        resultados_pendientes_df = resultados_df_raw[
            resultados_df_raw['Hora de Devolución'].isna() | 
            (resultados_df_raw['Hora de Devolución'] == '')
        ]

        self._mostrar_resultados_en_treeview(resultados_pendientes_df)

    def _re_ejecutar_ultima_busqueda(self):
        llave = self._last_search_llave.get()
        año = self._last_search_año.get()
        mes_num_val = self._last_search_mes.get() # Obtener el valor numérico del mes
        dia = self._last_search_dia.get()

        self.search_var_llave.set(llave)
        self.search_var_ano.set(año)
        self.search_var_dia.set(dia)

        # Establecer el nombre del mes en la variable de display para el combobox
        if mes_num_val and mes_num_val in MESES_MAP_NUMERO_A_NOMBRE:
            self.search_mes_display_var.set(MESES_MAP_NUMERO_A_NOMBRE[mes_num_val])
            self.search_var_mes.set(mes_num_val) # Asegurar que el valor numérico interno esté sincronizado
        else:
            self.search_mes_display_var.set('')
            self.search_var_mes.set('') # Limpiar también el valor numérico interno

        self._on_mes_seleccionado() # Llamar para asegurar que el día está en el estado correcto según el mes y año

        if not llave and not año and not mes_num_val and not dia: # Usar mes_num_val aquí
            todos_los_registros = obtener_todos_los_registros()
            self._mostrar_resultados_en_treeview(todos_los_registros)
        else:
            resultados_df = self.buscar_registros(llave=llave, año=año, mes=mes_num_val, dia=dia) # Usar mes_num_val para la búsqueda
            self._mostrar_resultados_en_treeview(resultados_df)

    def _limpiar_y_mostrar_todo(self, should_clear_last_search=True):
        self.search_var_llave.set('')
        self.search_var_ano.set('')
        self.search_var_mes.set('') # Limpiar valor numérico
        self.search_mes_display_var.set('') # Limpiar valor de display
        self.search_var_dia.set('')

        if should_clear_last_search:
            self._last_search_llave.set('')
            self._last_search_año.set('')
            self._last_search_mes.set('') # Limpiar el último valor numérico
            self._last_search_dia.set('')

        self.search_llave_combo.config(state="normal")
        self.search_llave_combo['values'] = [''] + LLAVES_VALIDAS
        
        self._cargar_años_disponibles() # Esta función también maneja los combobox de mes y día
        self._on_mes_seleccionado() # Asegurar que el día se deshabilita al limpiar todo

        todos_los_registros = obtener_todos_los_registros()
        self._mostrar_resultados_en_treeview(todos_los_registros)

    def _mostrar_resultados_en_treeview(self, df):
        for item in self.tree_buscador.get_children():
            self.tree_buscador.delete(item)

        if df.empty:
            return

        for i, (index, row) in enumerate(df.iterrows()):
            tag = "evenrow" if i % 2 == 0 else "oddrow"
            self.tree_buscador.insert("", "end", iid=str(index), values=list(row), tags=(tag,))

    def _reaplicar_tags_zebra(self, treeview_widget):
        """
        Reaplica los tags de estilo de cebra a todas las filas del Treeview.
        """
        for i, item_id in enumerate(treeview_widget.get_children()):
            tag = "evenrow" if i % 2 == 0 else "oddrow"
            treeview_widget.item(item_id, tags=(tag,))

    def _ordenar_treeview_por_columna(self, col):
        """
        Ordena los datos del treeview por la columna especificada.
        Alterna entre orden ascendente y descendente.
        """
        data = []
        for item_id in self.tree_buscador.get_children():
            values = list(self.tree_buscador.item(item_id, 'values'))
            data.append((values, item_id))

        if not data:
            return

        col_index = COLUMNAS_REGISTRO.index(col)

        reverse = self._orden_actual.get(col, False)

        def get_sort_key(item):
            value = item[0][col_index]
            if col == 'Fecha de Entrega' and isinstance(value, str) and value.strip():
                try:
                    if len(value.split('/')[2]) == 2:
                        return datetime.strptime(value, "%d/%m/%y")
                    elif len(value.split('/')[2]) == 4:
                        return datetime.strptime(value, "%d/%m/%Y")
                    else:
                        return value
                except (ValueError, IndexError):
                    return value
            return value

        data.sort(key=get_sort_key, reverse=not reverse)

        for item_id in self.tree_buscador.get_children():
            self.tree_buscador.delete(item_id)

        for i, (values, original_item_id) in enumerate(data):
            tag = "evenrow" if i % 2 == 0 else "oddrow"
            self.tree_buscador.insert("", "end", iid=original_item_id, values=values, tags=(tag,))
        
        self._orden_actual[col] = not reverse
   
    def _cleanup_all_existing_temp_pdfs(self):
        """
        Busca y elimina cualquier archivo PDF temporal existente
        """
        current_dir = os.getcwd()
        for filename in os.listdir(current_dir):
            if filename.startswith(PREFIJO_CSV_TEMPORAL) and filename.endswith(".pdf"):
                file_path = os.path.join(current_dir, filename)
                try:
                    os.remove(file_path)
                    print(f"Eliminado PDF temporal existente: {file_path}")
                except Exception as e:
                    print(f"Advertencia: No se pudo eliminar el PDF temporal {file_path}. Podría estar en uso. Detalle: {e}")

    def _imprimir_tabla_gui(self, treeview_widget, title="Tabla de Registros"):
        """
        Exporta los resultados actualmente mostrados en un Treeview dado a un archivo PDF
        con formato optimizado para impresión, manejando el ajuste de texto en celdas
        y encabezados anidados.
        """
        items = treeview_widget.get_children()
        if not items:
            messagebox.showinfo("Información", f"No hay registros en la '{title}' para imprimir.")
            return

        data_to_export = []
        for item_id in items:
            values = treeview_widget.item(item_id, 'values')
            data_to_export.append([str(v) for v in values])

        # --- Borrar PDFs temporales existentes antes de generar uno nuevo ---
        self._cleanup_all_existing_temp_pdfs()
        # -------------------------------------------------------------------

        # --- DEFINICIÓN DE ENCABEZADOS PARA EL PDF ---
        pdf_headers_config = [
            {"text": "Conserje", "cols_span": 1, "sub_headers": None},
            {"text": "Llave", "cols_span": 1, "sub_headers": None},
            {"text": "Hora de", "cols_span": 2, "sub_headers": ["Entrega", "Devolución"]}, 
            {"text": "Motivo de la Entrega", "cols_span": 1, "sub_headers": None},
            {"text": "Nombre de la Persona", "cols_span": 1, "sub_headers": None},
            {"text": "Fecha de Entrega", "cols_span": 1, "sub_headers": None}
        ]
        
        total_span = sum(h["cols_span"] for h in pdf_headers_config)
        if total_span != len(COLUMNAS_REGISTRO):
            messagebox.showerror("Error de Configuración", 
                                 f"El número total de columnas abarcadas en la configuración de encabezados ({total_span}) no coincide con el número de COLUMNAS_REGISTRO ({len(COLUMNAS_REGISTRO)}).")
            return

        safe_title = title.replace(" ", "_").replace("__", "_").replace("de_", "").lower()
        timestamp = datetime.now().strftime("%d%m%Y_%H%M%S")
        filename = f"{PREFIJO_CSV_TEMPORAL}{safe_title}_{timestamp}.pdf"
        full_path = os.path.abspath(filename)

        try:
            pdf = FPDF(orientation='L', unit='mm', format='A4')
            pdf.set_auto_page_break(auto=True, margin=15) 
            pdf.add_page()
            
            # --- Título del documento ---
            pdf.set_font('Arial', 'B', 16)
            pdf.cell(0, 10, title, border=0, align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT) 
            pdf.ln(5) 

            # --- Definición de Anchos de Columna (CRÍTICO - AJUSTA ESTOS VALORES) ---
            col_widths_mm = [
                30,  # Conserje
                45,  # Llave 
                25,  # Hora de Entrega 
                30,  # Hora de Devolución 
                50,  # Motivo de la Entrega 
                40,  # Nombre de la Persona
                30   # Fecha de Entrega
            ]
            
            total_current_width = sum(col_widths_mm)
            page_available_width = pdf.w - 2 * pdf.l_margin
            
            if total_current_width < page_available_width:
                extra_space_per_column = (page_available_width - total_current_width) / len(col_widths_mm)
                col_widths_mm = [width + extra_space_per_column for width in col_widths_mm]


            # --- Función auxiliar para dibujar una fila de celdas ---
            def draw_cells_in_row(pdf_obj, cells_data, col_widths, cell_height_base, align='L', border=1):
                x_start_row = pdf_obj.get_x()
                y_start_row = pdf_obj.get_y() 
                
                max_row_height = cell_height_base 

                current_x = x_start_row
                
                for i, cell_text in enumerate(cells_data):
                    pdf_obj.set_xy(current_x, y_start_row) 
                    
                    y_before_cell_draw = pdf_obj.get_y()
                    
                    current_border = border
                    if align == 'C' and cell_text.strip() == "": 
                        current_border = 0 
                    
                    pdf_obj.multi_cell(col_widths[i], cell_height_base, cell_text, border=current_border, align=align, new_x=XPos.RIGHT)
                    
                    height_occupied_by_cell = pdf_obj.get_y() - y_before_cell_draw
                    if height_occupied_by_cell > max_row_height:
                        max_row_height = height_occupied_by_cell
                    
                    current_x += col_widths[i] 

                pdf_obj.set_xy(pdf_obj.l_margin, y_start_row + max_row_height)
                
                return max_row_height 


            # --- DIBUJAR ENCABEZADOS DE TABLA (DOS NIVELES) ---
            pdf.set_font('Arial', 'B', 10) 
            
            x_current = pdf.l_margin
            y_start_top_headers = pdf.get_y()
            max_top_header_height = 0
            
            col_index_counter = 0 
            for header_info in pdf_headers_config:
                text = header_info["text"]
                cols_span = header_info["cols_span"]
                
                span_width = sum(col_widths_mm[col_index_counter : col_index_counter + cols_span])
                
                pdf.set_xy(x_current, y_start_top_headers)
                y_before_cell = pdf.get_y()
                pdf.multi_cell(span_width, 10, text, border=1, align='C', new_x=XPos.RIGHT) 
                
                height_occupied = pdf.get_y() - y_before_cell
                if height_occupied > max_top_header_height:
                    max_top_header_height = height_occupied
                
                x_current += span_width
                col_index_counter += cols_span

            pdf.set_xy(pdf.l_margin, y_start_top_headers + max_top_header_height)
            
            bottom_row_texts_for_drawing = []
            for header_info in pdf_headers_config:
                if header_info["sub_headers"]: 
                    bottom_row_texts_for_drawing.extend(header_info["sub_headers"])
                else: 
                    for _ in range(header_info["cols_span"]):
                        bottom_row_texts_for_drawing.append("") 
            
            bottom_header_height = draw_cells_in_row(pdf, bottom_row_texts_for_drawing, col_widths_mm, 10, align='C', border=1)
            
            total_header_height = max_top_header_height + bottom_header_height

            # --- Imprimir Datos de la Tabla ---
            pdf.set_font('Arial', '', 9) 
            
            for row_data in data_to_export:
                estimated_row_height = 8 

                if pdf.get_y() + estimated_row_height > pdf.h - pdf.b_margin: 
                    pdf.add_page()
                    pdf.set_font('Arial', 'B', 10)
                    
                    x_current = pdf.l_margin
                    y_start_top_headers_new_page = pdf.get_y()
                    max_top_header_height_new_page = 0
                    col_index_counter = 0
                    for header_info in pdf_headers_config:
                        text = header_info["text"]
                        cols_span = header_info["cols_span"]
                        span_width = sum(col_widths_mm[col_index_counter : col_index_counter + cols_span])
                        
                        pdf.set_xy(x_current, y_start_top_headers_new_page)
                        y_before_cell = pdf.get_y()
                        pdf.multi_cell(span_width, 10, text, border=1, align='C', new_x=XPos.RIGHT)
                        
                        height_occupied = pdf.get_y() - y_before_cell
                        if height_occupied > max_top_header_height_new_page:
                            max_top_header_height_new_page = height_occupied
                        
                        x_current += span_width
                        col_index_counter += cols_span
                    
                    pdf.set_xy(pdf.l_margin, y_start_top_headers_new_page + max_top_header_height_new_page)
                    
                    draw_cells_in_row(pdf, bottom_row_texts_for_drawing, col_widths_mm, 10, align='C', border=1)
                    
                    pdf.set_font('Arial', '', 9) 
                
                draw_cells_in_row(pdf, row_data, col_widths_mm, 5, align='L', border=1) 


            # --- Guardar y Abrir el PDF ---
            pdf.output(full_path)
            messagebox.showinfo("Imprimir", f"Archivo PDF generado: {os.path.basename(full_path)}")

            try:
                if sys.platform == "win32":
                    os.startfile(full_path) 
                elif sys.platform == "darwin": 
                    os.system(f'open "{full_path}"')
                else: 
                    os.system(f'xdg-open "{full_path}"')
                messagebox.showinfo("Impresión", "Se ha abierto el archivo PDF. Por favor, imprímelo desde el visor de PDF.")
            except Exception as e:
                messagebox.showerror("Error al Abrir PDF", f"No se pudo abrir el archivo PDF automáticamente: {e}\n\nPor favor, abre el archivo manualmente en:\n{full_path}")

        except Exception as e:
            error_trace = traceback.format_exc()
            messagebox.showerror(
                "Error al Generar PDF",
                f"No se pudo generar el archivo PDF para impresión: {e}\n\nDetalles:\n{error_trace}"
            )
             
    def _move_and_edit_next_cell(self, event=None, direction=None, current_row_id=None, current_col_index_from_event=None):
        """
        Finaliza la edición de la celda actual y mueve el foco al siguiente
        campo para edición, activando el editor automáticamente.
        """
        # Finalizar el editor actual si existe antes de mover
        if self._editor_widget and self._editor_widget.winfo_exists():
            self._finalize_active_editor()
            # Si al finalizar el editor hubo un error de validación y no se destruyó, no avanzamos.
            if self._editor_widget and self._editor_widget.winfo_exists(): 
                return 

        # Determinar la celda actual. Priorizar los argumentos pasados.
        current_item = current_row_id if current_row_id is not None else self.entry_tree.focus()
        current_col_index = current_col_index_from_event if current_col_index_from_event is not None else (self._active_editor_col_index if self._active_editor_col_index is not None else 0)
        
        if not current_item:
            selected_items = self.entry_tree.selection()
            if selected_items:
                current_item = selected_items[0]
            else:
                all_items = self.entry_tree.get_children()
                if all_items:
                    current_item = all_items[0] # Seleccionar la primera fila si no hay foco ni selección
                else:
                    return # No hay filas en la tabla, no se puede mover

        next_item = current_item
        next_col_index = current_col_index

        num_cols = len(COLUMNAS_REGISTRO)
        all_items = self.entry_tree.get_children()
        num_rows = len(all_items)
        current_row_index = all_items.index(current_item) if current_item in all_items else -1

        # Lógica de movimiento basada en la dirección o tipo de evento
        if direction == 'right' or (event and event.keysym in ['Return', 'Tab']):
            next_col_index += 1
            if next_col_index >= num_cols:
                next_col_index = 0
                next_row_index = current_row_index + 1
                if next_row_index >= num_rows:
                    # Al final de la tabla, añadir una nueva fila
                    self._anadir_fila_vacia(num_rows=1)
                    all_items = self.entry_tree.get_children() # Actualizar la lista de items
                    next_item = all_items[-1] # La nueva última fila
                else:
                    next_item = all_items[next_row_index]
        elif direction == 'left' or (event and event.keysym == 'Left'):
            next_col_index -= 1
            if next_col_index < 0:
                next_col_index = num_cols - 1 # Ir a la última columna de la fila anterior
                next_row_index = current_row_index - 1
                if next_row_index < 0:
                    next_row_index = 0 # Mantener en la primera fila
                    next_col_index = 0 # Y primera columna
                next_item = all_items[next_row_index]
        elif direction == 'down' or (event and event.keysym == 'Down'):
            next_row_index = current_row_index + 1
            if next_row_index >= num_rows:
                self._anadir_fila_vacia(num_rows=1)
                all_items = self.entry_tree.get_children()
                next_item = all_items[-1]
            else:
                next_item = all_items[next_row_index]
        elif direction == 'up' or (event and event.keysym == 'Up'):
            next_row_index = current_row_index - 1
            if next_row_index < 0:
                next_row_index = 0 # Mantener en la primera fila
            next_item = all_items[next_row_index]
        else:
            # Si no hay dirección específica (ej. click simple en una celda sin movimiento)
            # Solo queremos activar el editor en la celda clickada, no mover.
            # En este caso, current_item y current_col_index ya deberían ser los del click.
            pass 
            
        # Asegurarse de que next_item es visible (desplazarse si es necesario)
        self.entry_tree.see(next_item)
        self.entry_tree.selection_set(next_item) # Seleccionar la nueva fila
        
        # Activar el editor para la nueva celda
        self._activate_cell_editor(next_item, next_col_index)


if __name__ == "__main__":

    # realizamos la copia de seguridad de los registros actuales
    realizar_backup_automatico(NOMBRE_ARCHIVO_EXCEL, CARPETA_BACKUPS, DIAS_ENTRE_BACKUPS)
    root = tk.Tk()
    app = AppRegistroLlaves(root)
    root.mainloop()


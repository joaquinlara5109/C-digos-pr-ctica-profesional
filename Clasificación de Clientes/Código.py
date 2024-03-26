#Inc Vendedor

#Librerías
import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkcalendar import DateEntry
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import subprocess
import pandas as pd
import numpy as np
import csv
import sys
import locale
import openpyxl


def detectar_delimitador(archivo):
    codificaciones = ['ISO-8859-1', 'latin-1', 'utf-8']
    for encoding in codificaciones:
        try:
            with open(archivo, 'r', encoding=encoding) as f:
                linea = f.readline()
            delimitadores = [',', ';', '\t']
            delimitador_frecuencia = {delim: linea.count(delim) for delim in delimitadores}
            delimitador = max(delimitador_frecuencia, key=delimitador_frecuencia.get)
            return delimitador, encoding
        except UnicodeDecodeError:
            continue
    return None, None  

def encontrar_skip_rows_csv(archivo, delimitador, encoding):
    with open(archivo, 'r', encoding=encoding) as f:
        for i, linea in enumerate(f):
            if not linea.strip().startswith('Unnamed'):
                valores = linea.split(delimitador)
                valores_no_vacios = [v for v in valores if v.strip()]
                if len(valores_no_vacios) >= 6:
                    return i
    return 0

def encontrar_skip_rows_excel(archivo):
    df_temp = pd.read_excel(archivo, engine='openpyxl', nrows=20)
    for i, row in df_temp.iterrows():
        celdas_no_vacias = sum(not pd.isna(cell) for cell in row[:6])
        if celdas_no_vacias >= 6:
            return i
    return 0

def leer_archivo(archivo):
    try:
        if archivo.lower().endswith('.csv'):
            delimitador, encoding = detectar_delimitador(archivo)
            if delimitador is None:
                raise ValueError("No se pudo determinar el delimitador.")
            skip_rows = encontrar_skip_rows_csv(archivo, delimitador, encoding)
            df = pd.read_csv(archivo, delimiter=delimitador, skiprows=skip_rows, encoding=encoding, on_bad_lines='skip')
        elif archivo.lower().endswith('.xlsx'):
            skip_rows = encontrar_skip_rows_excel(archivo)
            df = pd.read_excel(archivo, engine='openpyxl', skiprows=skip_rows)
            if skip_rows > 0:
                df.columns = df.iloc[0] 
                df = df[1:]
            mask = df.columns.astype(str).str.contains('^Unnamed')
            df = df.loc[:, ~mask]
            columns_with_nan_title = df.columns[pd.isna(df.columns)]
            for col in columns_with_nan_title:
                df = df[pd.isna(df[col]) | (df[col] == '')]
                df = df.dropna(how='all')
        else:
            raise ValueError("Formato de archivo no soportado.")
    except Exception as e:
        print(f"Error al leer el archivo: {e}")
        raise
    return df       

# Funciones para limpiar los datos
def mapear_meses(dfnuevo):
    meses_mapeo = {
        'ENERO': 1, 'FEBRE': 2, 'MARZO': 3, 'ABRIL': 4, 'MAYO': 5, 'JUNIO': 6,
        'JULIO': 7, 'AGOST': 8, 'SEPTI': 9, 'OCTUB': 10, 'NOVIE': 11, 'DICIE': 12
    }
    dfnuevo['MES'] = dfnuevo['MES'].str.strip().map(meses_mapeo)
    dfnuevo['FECHA'] = pd.to_datetime(dfnuevo.assign(DAY=dfnuevo['DIA'], MONTH=dfnuevo['MES'], YEAR=dfnuevo['AÑO']).loc[:, ['YEAR','MONTH','DAY']])
    dfnuevo.sort_values(by='FECHA', inplace=True)

def eliminar_columnas1(df):

    df['DESCRIPCION CODIGO'] = df['DESCRIPCION CODIGO'].str.strip().str.upper()
    df['LINEA'] = df['LINEA'].str.strip().str.upper()
    df['FAMILIA'] = df['FAMILIA'].str.strip().str.upper()
    df['COND.PAGO'] = df['COND.PAGO'].str.strip().str.upper()
    df['VENDEDOR'] = df['VENDEDOR'].str.strip().str.upper()


    df_filtrado = df[
        (df['DOCUMENTO'] == 'FACTURA ELECTRONICA') &
        (df['COND.PAGO'].str.strip() != 'TRASLADOS') &
        (df['ESTADO'] != 'NULO') &
        (df['CODIGO'] != 99999) 
    ].copy()

    columnas_para_conservar = ['DOCUMENTO', 'NUMERO', 'DIA', 'MES', 'AÑO', 'ESTADO', 'RUT', 'DESCRIPCION RUT', 'VENDEDOR', 'CODIGO', 'DESCRIPCION CODIGO', 'CANTIDAD','NETO', 'COSTO FINAL']
    df_filtrado = df_filtrado[columnas_para_conservar]
    return df_filtrado

def eliminar_columnas2(df):

    df['DESCRIPCION CODIGO'] = df['DESCRIPCION CODIGO'].str.strip().str.upper()
    df['LINEA'] = df['LINEA'].str.strip().str.upper()
    df['FAMILIA'] = df['FAMILIA'].str.strip().str.upper()
    df['COND.PAGO'] = df['COND.PAGO'].str.strip().str.upper()
    df['VENDEDOR'] = df['VENDEDOR'].str.strip().str.upper()
    df['DESCRIPCION RUT'] = df['DESCRIPCION RUT'].str.strip().str.upper()


    df_filtrado = df[
        (df['DOCUMENTO'] == 'NOTA DE CREDITO ELECTRONICA') &
        (df['COND.PAGO'].str.strip() != 'TRASLADOS') &
        (df['ESTADO'] != 'NULO') &
        (df['CODIGO'] != 99999) 
    ].copy()

    columnas_para_conservar = ['DOCUMENTO', 'NUMERO', 'DIA', 'MES', 'AÑO', 'ESTADO', 'RUT', 'DESCRIPCION RUT', 'VENDEDOR', 'CODIGO', 'DESCRIPCION CODIGO', 'CANTIDAD','NETO', 'COSTO FINAL']
    df_filtrado = df_filtrado[columnas_para_conservar]
    return df_filtrado

# Funciones para clasifiación de clientes
def analisis_compras_clientes(df):

    df = df.sort_values(by=['NUMERO', 'NETO'], ascending=[True, False])
    df_suma_neto = df.groupby('RUT')['NETO'].sum().reset_index().rename(columns={'NETO': 'Acum. Vent. Netas'})
    df_unicos = df.drop_duplicates(subset=['NUMERO'], keep='first')
    df_unicos['TRANSACCION'] = df_unicos['CANTIDAD'].apply(lambda x: 0 if pd.isna(x) else (1 if x > 0 else (0 if x == 0 else -1)))
    df_conteo_ajustado = df_unicos.groupby('RUT')['TRANSACCION'].sum().reset_index().rename(columns={'TRANSACCION': 'Num. Compras'})
    df_resultado = pd.merge(df_suma_neto, df_conteo_ajustado, on='RUT', how='outer')

    return df_resultado

def combinar_analisis_y_devoluciones(df_analisis_clientes, df_devoluciones_clientes, df_razon_social):

    df_devoluciones_clientes = df_devoluciones_clientes.rename(columns={'Acum. Vent. Netas': 'Monto Devoluciones', 'Num. Compras': 'Num. Devoluciones'})
    df_devoluciones_clientes['Monto Devoluciones'] = df_devoluciones_clientes['Monto Devoluciones'].fillna(0)
    df_devoluciones_clientes['Num. Devoluciones'] = df_devoluciones_clientes['Num. Devoluciones'].fillna(0)
    df_combinado = pd.merge(df_analisis_clientes, df_devoluciones_clientes, on='RUT', how='inner')
    df_combinado['Neto Final'] = df_combinado['Acum. Vent. Netas'] + df_combinado['Monto Devoluciones']
    df_combinado['Compras finales'] = df_combinado['Num. Compras'] + df_combinado['Num. Devoluciones']
    df_combinado = pd.merge(df_combinado, df_razon_social, on='RUT', how='left')
    df_combinado['Num. Devoluciones'] = df_combinado['Num. Devoluciones'] * -1
    columnas_interes = ['DESCRIPCION RUT', 'RUT', 'Acum. Vent. Netas', 'Num. Compras', 'Monto Devoluciones', 'Num. Devoluciones', 'Neto Final', 'Compras finales']
    df_resultado_final = df_combinado[columnas_interes]

    return df_resultado_final

# Funciones para realizar el concurso para todos los vendedores y los de retail
def procesar_concurso(df_filtrado_final, df_devoluciones_final):

    def calcular_neto_y_costo_final(df, df_relaciones):
        df_agrupado = df.groupby(['VENDEDOR', 'RUT']).agg(
        NETO_FINAL=('NETO', 'sum'),
        COSTO_FINAL=('COSTO FINAL', 'sum')
        ).reset_index()
        df_final = pd.merge(df_relaciones, df_agrupado, on=['VENDEDOR', 'RUT'], how='left')
        return df_final

    def union_neto_y_costo_final(df_final_ventas, df_final_devoluciones):
        df_final_devoluciones = df_final_devoluciones.rename(columns={'NETO_FINAL':'NETO D', 'COSTO_FINAL':'COSTO FINAL D'})
        df_unificado = pd.merge(df_final_ventas, df_final_devoluciones, on=['VENDEDOR', 'RUT'], how='left')
        df_unificado['NETO F'] = df_unificado['NETO_FINAL'] + df_unificado['NETO D']
        df_unificado['COSTO F'] = df_unificado['COSTO_FINAL'] + df_unificado['COSTO FINAL D']
        df_unificado['NETO F'] = df_unificado['NETO F'].fillna(0) 
        df_unificado['COSTO F'] = df_unificado['COSTO F'].fillna(0)

        df_unificado['MARGENP'] = np.where(df_unificado['NETO F'] != 0, (df_unificado['NETO F'] - df_unificado['COSTO F']) / df_unificado['NETO F'], 0)

        df_unificado['Num. ClientesN'] = (df_unificado['NETO F'] > 0).astype(int)

        df_resumen = df_unificado.groupby('VENDEDOR').agg(
            NETO_F_TOTAL=('NETO F', 'sum'),
            Num_ClientesN_Total=('Num. ClientesN', 'sum'),
            MARGENP_Promedio=('MARGENP', 'mean')
        ).reset_index()

        return df_resumen
  
    def agregar_puntuaciones(df_resumen):
        def normalizar(column):
            maximo = column.max()
            minimo = column.min()

            if maximo == minimo:
                return pd.Series(1, index=column.index)
            else:
                return ((column - minimo) / (maximo - minimo)).round(3)

        df_resumen['Puntuación Cantidad C'] = normalizar(df_resumen['Num_ClientesN_Total'])
        df_resumen['Puntuación Vent.Netas'] = normalizar(df_resumen['NETO_F_TOTAL'])
        df_resumen['Puntuación Margen'] = normalizar(df_resumen['MARGENP_Promedio'])
        
        peso_clientes_nuevos = 1/3
        peso_ingresos = 1/3
        peso_margenp_promedio = 1/3
        
        df_resumen['Puntuación Total'] = (
            df_resumen['Puntuación Cantidad C'] * peso_clientes_nuevos +
            df_resumen['Puntuación Vent.Netas'] * peso_ingresos +
            df_resumen['Puntuación Margen'] * peso_margenp_promedio
        ).round(3)
        
        df_resumen = df_resumen.rename(columns={
            'Num_ClientesN_Total': 'Cantidad ClientesN', 
            'NETO_F_TOTAL': 'Acum. Vent.Netas',
            'MARGENP_Promedio': 'Margen Prom'
        })

        return df_resumen

    def filtrar_por_vendedor(df):

        vendedores_especial = [
            "RODRIGO OSSANDON",
            "FRANCISCO LARCO",
            "JUAN C.JULIA",
            "MARJORIE JULIO",
            "JUAN SEGURA",
            "RODRIGO PAREDES",
            "FRANCISCO GREBE"
        ]

        df_final = df[df['VENDEDOR'].isin(vendedores_especial)]

        return df_final

    df_ventas_filtrado = df_filtrado_final.copy()
    df_devoluciones_filtrado = df_devoluciones_final[df_devoluciones_final['CODIGO'] != 901].copy()
    df_relaciones = df_ventas_filtrado.drop_duplicates(subset=['RUT'], keep='first')[['VENDEDOR', 'RUT']].copy()
    df_final_ventas = calcular_neto_y_costo_final(df_ventas_filtrado, df_relaciones)
    df_final_devoluciones = calcular_neto_y_costo_final(df_devoluciones_filtrado, df_relaciones)
    df_final_ventas.fillna(0, inplace=True)
    df_final_devoluciones.fillna(0, inplace=True)
    df_resumen = union_neto_y_costo_final(df_final_ventas, df_final_devoluciones)
    df_concurso_vendedores_totales = agregar_puntuaciones(df_resumen)
    df_concurso_especial = filtrar_por_vendedor(df_resumen)
    df_concurso_vendedores_especial = agregar_puntuaciones(df_concurso_especial)

    return df_concurso_vendedores_totales, df_concurso_vendedores_especial, df_relaciones

def guardar_excel_con_formato(archivo_salida, dataframes_con_nombres):
    with pd.ExcelWriter(archivo_salida, engine='xlsxwriter') as writer:
        workbook = writer.book
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#D9D9D9',
            'border': 1
        })
        
        formato_dinero = workbook.add_format({'num_format': '$#,##0.00'})
        formato_porcentaje = workbook.add_format({'num_format': '0.00%'})

        for df, sheet_name in dataframes_con_nombres:        
            column_widths = [max([len(str(x)) for x in df[col].values] + [len(str(col))]) for col in df.columns]    
            df.to_excel(writer, sheet_name=sheet_name, startrow=2, header=False, index=False)          
            worksheet = writer.sheets[sheet_name]
            
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
                
            worksheet.autofilter(1, 0, 1, len(df.columns) - 1)
            
            for i, width in enumerate(column_widths):
                worksheet.set_column(i, i, width)

                if sheet_name in ['Clientes Nuevos']:
                    for col_name in ['Acum. Vent. Netas', 'Monto Devoluciones', 'Neto Final']:
                        if col_name in df.columns:
                            col_idx = df.columns.get_loc(col_name) 
                            for row_num in range(len(df)):
                                worksheet.write(row_num + 2, col_idx, df[col_name].iloc[row_num], formato_dinero)
                
                elif sheet_name in ['Concurso General', 'Concurso Retail']:
                    for col_name in ['Acum. Vent.Netas']:
                        if col_name in df.columns:
                            col_idx = df.columns.get_loc(col_name)  
                            for row_num in range(len(df)):
                                worksheet.write(row_num + 2, col_idx, df[col_name].iloc[row_num], formato_dinero)
                    
                    for col_name in ['Margen Prom']:
                        if col_name in df.columns:
                            col_idx = df.columns.get_loc(col_name) 
                            for row_num in range(len(df)):
                                worksheet.write(row_num + 2, col_idx, df[col_name].iloc[row_num], formato_porcentaje)

# Función Principal, se hacen las llamadas a las demás funciones
def procesar(archivo_concurso, archivo_salida, archivo_antiguos, fecha_inicio_nuevos, fecha_fin_nuevos, fecha_inicio_antiguos, fecha_fin_antiguos, activar_concurso):
    df_ventas = leer_archivo(archivo_concurso)
    df_filtrado = eliminar_columnas1(df_ventas)
    df_antiguos = leer_archivo(archivo_antiguos)
    df_filtrado_antiguos = eliminar_columnas1(df_antiguos)
    df_devoluciones = eliminar_columnas2(df_ventas)

    fecha_inicio_nuevos = pd.to_datetime(fecha_inicio_nuevos, format='%Y-%m-%d')
    fecha_fin_nuevos = pd.to_datetime(fecha_fin_nuevos, format='%Y-%m-%d')
    fecha_inicio_antiguos = pd.to_datetime(fecha_inicio_antiguos, format='%Y-%m-%d')
    fecha_fin_antiguos = pd.to_datetime(fecha_fin_antiguos, format='%Y-%m-%d')

    mapear_meses(df_filtrado_antiguos)
    mapear_meses(df_filtrado)
    mapear_meses(df_devoluciones)

    df_filtrado = df_filtrado[(df_filtrado['FECHA'] >= fecha_inicio_nuevos) & (df_filtrado['FECHA'] <= fecha_fin_nuevos)].copy()
    df_devoluciones = df_devoluciones[(df_devoluciones['FECHA'] >= fecha_inicio_nuevos) & (df_devoluciones['FECHA'] <= fecha_fin_nuevos)].copy()

    df_filtrado_antiguos = df_filtrado_antiguos[(df_filtrado_antiguos['FECHA'] >= fecha_inicio_antiguos) & (df_filtrado_antiguos['FECHA'] <= fecha_fin_antiguos)].copy()

    df_filtrado_antiguos_sin_duplicados = df_filtrado_antiguos.drop_duplicates(subset=['RUT'], keep='first')

    ruts_no_deseados = df_filtrado_antiguos_sin_duplicados['RUT'].unique()
    df_filtrado_final = df_filtrado[~df_filtrado['RUT'].isin(ruts_no_deseados)]
    df_filtrado_final = df_filtrado[~df_filtrado['RUT'].isin(ruts_no_deseados)]
    ruts_filtrados = df_filtrado_final[['RUT']].drop_duplicates()
    df_devoluciones_final = pd.merge(ruts_filtrados, df_devoluciones, on='RUT', how='left')
    df_analisis_clientes = analisis_compras_clientes(df_filtrado_final)
    df_devoluciones_clientes =analisis_compras_clientes(df_devoluciones_final)
    df_razon_social = df_ventas[['RUT', 'DESCRIPCION RUT']]
    df_razon_social = df_razon_social.drop_duplicates(subset=['RUT'], keep='first')
    df_clientes_finales = combinar_analisis_y_devoluciones(df_analisis_clientes, df_devoluciones_clientes, df_razon_social)
    df_devolucion_to_excel = df_devoluciones_final[df_devoluciones_final['FECHA'].notna()]   
    columnas_interes = ['DOCUMENTO', 'NUMERO', 'CODIGO', 'DESCRIPCION CODIGO','DIA', 'MES', 'AÑO', 'ESTADO', 'RUT', 'DESCRIPCION RUT', 'VENDEDOR', 'CANTIDAD', 'NETO', 'COSTO FINAL', 'FECHA']
    df_devolucion_to_excel = df_devolucion_to_excel[columnas_interes]
    df_filtrado_final = df_filtrado_final[columnas_interes]

    if activar_concurso:

        df_concurso_vendedores_totales, df_concurso_vendedores_especial, df_relaciones= procesar_concurso(df_filtrado_final, df_devoluciones_final)
        df_relaciones = pd.merge(df_relaciones, df_razon_social, on='RUT', how= 'left')
        dataframes_con_nombres = [
            (df_concurso_vendedores_totales, 'Concurso General'),
            (df_concurso_vendedores_especial, 'Concurso Retail'),
            (df_relaciones, 'Relaciones'),
            (df_clientes_finales, 'Clientes Nuevos'),
            (df_filtrado_final, 'Ventas'),
            (df_devolucion_to_excel, 'Devoluciones'),
        ]
            #df_filtrado_antiguos_sin_duplicados[['RUT']].to_excel(writer, sheet_name='Clientes Antiguos', index=False, header=True, startrow=0)

    else:
        
        dataframes_con_nombres = [
            (df_clientes_finales, 'Clientes Nuevos'),
            (df_filtrado_final, 'Ventas'),
            (df_devolucion_to_excel, 'Devoluciones'),
        ]           
            #df_filtrado_antiguos_sin_duplicados[['RUT']].to_excel(writer, sheet_name='Clientes Antiguos', index=False, header=True, startrow=0)
    guardar_excel_con_formato(archivo_salida, dataframes_con_nombres)



#archivo_concurso= 'Concurso.xlsx'
#archivo_antiguos = 'VENTAS AÑO-2023.xlsx'
#archivo_salida = 'Clientes Nuevosz.xlsx'
#procesar(archivo_concurso, archivo_salida, archivo_antiguos)

# Inicio del código para la interfaz
def abrir_pdf():
    nombre_pdf = "Gestión de Clientes1.pdf"
    if getattr(sys, 'frozen', False):
        basedir = sys._MEIPASS  
    else:
        basedir = os.path.dirname(__file__)
    pdf_path = os.path.join(basedir, nombre_pdf)
    
    if os.name == 'nt':
        os.startfile(pdf_path, 'open')
    elif os.name == 'posix':
        subprocess.call(['open', pdf_path])
    else:
        subprocess.call(['xdg-open', pdf_path])

def abrir_pdf1():
    nombre_pdf = "Gestión de Clientes2.pdf"
    if getattr(sys, 'frozen', False):
        basedir = sys._MEIPASS  
    else:
        basedir = os.path.dirname(__file__)
    pdf_path = os.path.join(basedir, nombre_pdf)
    
    if os.name == 'nt':
        os.startfile(pdf_path, 'open')
    elif os.name == 'posix':
        subprocess.call(['open', pdf_path])
    else:
        subprocess.call(['xdg-open', pdf_path])

def cargar_imagen(nombre_imagen="Perno1.png"):
    if getattr(sys, 'frozen', False):
        basedir = sys._MEIPASS
    else:
        basedir = os.path.dirname(__file__)
    imagen_path = os.path.join(basedir, nombre_imagen)
    return tk.PhotoImage(file=imagen_path)

def gui():

    try:
        locale.setlocale(locale.LC_TIME, 'es_ES')  
    except locale.Error:
        print("La localización específica no pudo ser establecida.")

    def seleccionar_archivo_concurso():
        archivo = filedialog.askopenfilename(title="Seleccionar archivo de CONCURSO", filetypes=[("Excel files", "*.xlsx *.xls *.csv")])
        entry_archivo_concurso.delete(0, tk.END)
        entry_archivo_concurso.insert(0, archivo)

    def seleccionar_archivo_antiguos():
        archivo = filedialog.askopenfilename(title="Seleccionar archivo de antiguos", filetypes=[("Excel files", "*.xlsx *.xls *.csv")])
        entry_archivo_antiguos.delete(0, tk.END)
        entry_archivo_antiguos.insert(0, archivo)

    def seleccionar_archivo_salida():
        archivo = filedialog.asksaveasfilename(title="Especificar archivo de salida", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        entry_archivo_salida.delete(0, tk.END)
        entry_archivo_salida.insert(0, archivo)  

    def obtener_fecha_formato(calendar_widget):
        return calendar_widget.get_date().strftime('%Y-%m-%d')

    def procesar_archivos():
        archivo_antiguos = entry_archivo_antiguos.get()
        archivo_concurso = entry_archivo_concurso.get()
        archivo_salida = entry_archivo_salida.get()
        fecha_inicio_nuevos_str = obtener_fecha_formato(fecha_inicio_nuevos)
        fecha_fin_nuevos_str = obtener_fecha_formato(fecha_fin_nuevos)
        fecha_inicio_antiguos_str = obtener_fecha_formato(fecha_inicio_antiguos)
        fecha_fin_antiguos_str = obtener_fecha_formato(fecha_fin_antiguos)
        activar_concurso = generar_vendedores_var.get()  


        if archivo_concurso and archivo_antiguos and archivo_salida:
            try:
                procesar(archivo_concurso, archivo_salida, archivo_antiguos, fecha_inicio_nuevos_str, fecha_fin_nuevos_str, fecha_inicio_antiguos_str, fecha_fin_antiguos_str, activar_concurso)
                messagebox.showinfo("Éxito", "Los archivos han sido procesados correctamente.")
            except Exception as e:
                messagebox.showerror("Error", f"Ocurrió un error al procesar los archivos: {e}")
        else:
            messagebox.showwarning("Advertencia", "Por favor, complete todos los campos.")

    root = tk.Tk()
    root.title("Gestión y clasificación de Clientes")
    root.geometry('640x435')
    root.configure(bg='#2D2D2D')



    style = ttk.Style()
    style.configure('TLabel', font=('Arial', 11), background='#2D2D2D', foreground='white')
    style.configure('TButton', font=('Arial', 11))
    style.configure('TEntry', font=('Arial', 11))
    style.configure('TCheckbutton', font=('Arial', 11))

    fondo_columna = tk.Frame(root, bg='white', width= 2)
    fondo_columna.grid(row=1, column=3, rowspan=10, sticky='nswe', pady=0)

    tk.Label(root, text="Selección de Archivos", anchor='center', 
        font=('Arial', 12, 'bold'), background='light coral', foreground='white').grid(row=1, column=0, padx=0, pady=0, sticky='we', columnspan=3)

    tk.Label(root, text="Guías", anchor='center', 
        font=('Arial', 12, 'bold'), background='light coral', foreground='white').grid(row=1, column=4, padx=0, pady=0, sticky='we', columnspan=5)
    tk.Label(root, text="Función extra", anchor='center', 
        font=('Arial', 12, 'bold'), background='light coral', foreground='white').grid(row=5, column=4, padx=0, pady=0, sticky='we', columnspan=5)




    logo = cargar_imagen()  
    logo_label = ttk.Label(root, image=logo, background='#2D2D2D')
    logo_label.grid(row=0, column=0, padx=5, pady=5, sticky='we', columnspan=3)  

    ttk.Label(root, text="Archivo con Clientes Nuevos:", anchor='w').grid(row=2, column=0, padx=5, pady=5, sticky='we')
    entry_archivo_concurso = ttk.Entry(root, width=25)
    entry_archivo_concurso.grid(row=2, column=1, padx=5, pady=5, sticky='we')
    ttk.Button(root, text="Seleccionar", command=seleccionar_archivo_concurso).grid(row=2, column=2, padx=5, pady=5, sticky='we')

    ttk.Label(root, text="Archivo con Clientes Antiguos:", anchor='w').grid(row=3, column=0, padx=5, pady=5, sticky='we')
    entry_archivo_antiguos = ttk.Entry(root, width=25)
    entry_archivo_antiguos.grid(row=3, column=1, padx=5, pady=5, sticky='we')
    ttk.Button(root, text="Seleccionar", command=seleccionar_archivo_antiguos).grid(row=3, column=2, padx=5, pady=5, sticky='we')

    ttk.Label(root, text="Archivo de Salida:", anchor='w').grid(row=4, column=0, padx=5, pady=5, sticky='we')
    entry_archivo_salida = ttk.Entry(root, width=25)
    entry_archivo_salida.grid(row=4, column=1, padx=5, pady=5, sticky='we')
    ttk.Button(root, text="Especificar", command=seleccionar_archivo_salida).grid(row=4, column=2, padx=5, pady=5, sticky='we')

    tk.Label(root, text="Período de evaluación para clientes nuevos", anchor='center', 
        font=('Arial', 12, 'bold'), background='light coral', foreground='white').grid(row=5, column=0, padx=0, pady=5, sticky='we', columnspan=3)

    fecha_inicio_nuevos = DateEntry(root, date_pattern='dd/mm/yyyy', locale='es_ES')
    fecha_inicio_nuevos.grid(row=6, column=1, padx=5, pady=5, sticky='we')
    ttk.Label(root, text="Fecha de Inicio:", anchor='w').grid(row=6, column=0, padx=5, pady=5, sticky='we')

    fecha_fin_nuevos = DateEntry(root, date_pattern='dd/mm/yyyy', locale='es_ES')
    fecha_fin_nuevos.grid(row=7, column=1, padx=5, pady=5, sticky='we')
    ttk.Label(root, text="Fecha de Término:", anchor='w').grid(row=7, column=0, padx=5, pady=5, sticky='we')

    tk.Label(root, text="Período en el que un cliente se considera antiguo", anchor='center', 
        font=('Arial', 12, 'bold'), background='light coral', foreground='white').grid(row=8, column=0, padx=0, pady=5, sticky='we', columnspan=3)

    fecha_inicio_antiguos = DateEntry(root, date_pattern='dd/mm/yyyy', locale='es_ES')
    fecha_inicio_antiguos.grid(row=9, column=1, padx=5, pady=5, sticky='we')
    ttk.Label(root, text="Fecha de Inicio:", anchor='w').grid(row=9, column=0, padx=5, pady=5, sticky='we')

    fecha_fin_antiguos = DateEntry(root, date_pattern='dd/mm/yyyy', locale='es_ES')
    fecha_fin_antiguos.grid(row=10, column=1, padx=5, pady=5, sticky='we')
    ttk.Label(root, text="Fecha de Término:", anchor='w').grid(row=10, column=0, padx=5, pady=5, sticky='we')

    ttk.Button(root, text="Interpretación", command=abrir_pdf1).grid(row=3, column=4, padx=10, pady=10, sticky='we')

    ttk.Button(root, text="Manual de Usuario", command=abrir_pdf).grid(row=2, column=4, padx=10, pady=5, sticky='we')


    generar_vendedores_var = tk.BooleanVar()
    generar_vendedores_var.set(False)  

    checkbutton_generar_vendedores = ttk.Checkbutton(root, text="Generar Concurso", variable=generar_vendedores_var, style='TCheckbutton')
    checkbutton_generar_vendedores.grid(row=6, column=4, padx=5, pady=5, sticky='w')

    ttk.Button(root, text="Generar Reporte", command=procesar_archivos).grid(row=11, column=4, padx=15, pady=20, sticky='we')


    root.mainloop()

if __name__ == "__main__":
    gui()


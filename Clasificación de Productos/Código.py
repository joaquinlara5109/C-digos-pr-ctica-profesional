#Clasifiación de productos

import csv
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import subprocess
import pandas as pd
import numpy as np
import csv
import sys  
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

def distribuir_total_negativo(df, df_901):
    for index, row in df_901.iterrows():
        numero = row['NUMERO']
        total_negativo = row['Ingresos']
        indices_relacionados = df[df['NUMERO'] == numero].index
        if len(indices_relacionados) > 1:
            df.loc[indices_relacionados, 'Ingresos'] += total_negativo / len(indices_relacionados)
    return df

def obtener_primera_descripcion(df):
    primera_descripcion = df.groupby('CODIGO')['DESCRIPCION CODIGO'].first().reset_index()
    return primera_descripcion.set_index('CODIGO')['DESCRIPCION CODIGO']

def obtener_primera_linea(df):
    primera_linea = df.groupby('CODIGO')['LINEA'].first().reset_index()
    return primera_linea.set_index('CODIGO')['LINEA']

def asignar_etiquetas(df):
    df_filtrado = df[(df['CANTIDAD'] > 0) & (df['UTILIDAD'] > 0)].copy()

    total_cantidad = df_filtrado['CANTIDAD'].sum()
    total_utilidad = df_filtrado['UTILIDAD'].sum()

    df_filtrado['% CONTR. CANT'] = (df_filtrado['CANTIDAD'] / total_cantidad) * 100
    df_filtrado['% CONTR. UTIL'] = (df_filtrado['UTILIDAD'] / total_utilidad) * 100

    df_filtrado.sort_values(by='% CONTR. CANT', ascending=False, inplace=True)
    df_filtrado.reset_index(drop=True, inplace=True)
    df_filtrado['% ACUM. CANT'] = df_filtrado['% CONTR. CANT'].cumsum()

    bins_cant = [-np.inf, 80, 95, np.inf]
    labels_cant = ['A', 'B', 'C']

    df_filtrado.sort_values(by='% CONTR. UTIL', ascending=False, inplace=True)
    df_filtrado.reset_index(drop=True, inplace=True)
    df_filtrado['% ACUM. UTIL'] = df_filtrado['% CONTR. UTIL'].cumsum()

    bins_util = [-np.inf, 80, 95, np.inf]
    labels_util = ['A', 'B', 'C']

    df_filtrado['GRUPO UTIL'] = pd.cut(df_filtrado['% ACUM. UTIL'], bins=bins_util, labels=labels_util)
    df_filtrado['GRUPO CANT'] = pd.cut(df_filtrado['% ACUM. CANT'], bins=bins_cant, labels=labels_cant)

    df_filtrado.drop(columns=['% CONTR. CANT', '% CONTR. UTIL', '% ACUM. CANT', '% ACUM. UTIL'], inplace=True)

    return df_filtrado

def renombrar_columnas(df):
    mapeo_nombres = {
        'CODIGO': 'Código',
        'PRIMERA DESCRIPCION': 'Descripción',
        'PRIMERA LINEA': 'Línea',
        'FAMILIA': 'Familia',
        'CANTIDAD': 'Cantidad',
        'COSTO FINAL': 'Costos',
        'UTILIDAD': 'Utilidad',
        'GRUPO UTIL': 'G. Utilidad',
        'GRUPO CANT': 'G. Cantidad'
    }
    return df.rename(columns=mapeo_nombres)

def filtrar_datos_negativos(df):
    return df[(df['UTILIDAD'] < 0) | (df['CANTIDAD'] < 0)]

def calcular_Transacciones(row):
    if row['CANTIDAD'] > 0:
        return 1
    elif row['CANTIDAD'] <= 0:
        return -1
    else:
        return 0


def agregar_desglose_documentos(df_filtrado):
    df_filtrado = df_filtrado[df_filtrado['CODIGO'] != 901].copy()

    df_filtrado['PRIMERA DESCRIPCION'] = df_filtrado['CODIGO'].map(obtener_primera_descripcion(df_filtrado))
    df_filtrado['PRIMERA LINEA'] = df_filtrado['CODIGO'].map(obtener_primera_linea(df_filtrado))

    df_filtrado['TRANS_FACT'] = df_filtrado.apply(
        lambda x: 1 if 'FACTURA ELECTRONICA' in x['DOCUMENTO'] else 
                  (-1 if 'NOTA DE CREDITO ELECTRONICA' in x['DOCUMENTO'] and x['RUT'] != 19 else 0),
        axis=1
    )
    df_filtrado['TRANS_BOLETA'] = df_filtrado.apply(
        lambda x: 1 if 'BOLETA' in x['DOCUMENTO'] else 
                  (-1 if 'NOTA DE CREDITO ELECTRONICA' in x['DOCUMENTO'] and x['RUT'] == 19 else 0),
        axis=1
    )

    df_filtrado['CANT_FACT'] = df_filtrado.apply(lambda x: x['CANTIDAD'] if 'FACTURA ELECTRONICA' in x['DOCUMENTO'] else 0, axis=1)
    df_filtrado['CANT_BOLETA'] = df_filtrado.apply(lambda x: x['CANTIDAD'] if 'BOLETA' in x['DOCUMENTO'] else 0, axis=1)
    df_filtrado['DEVOL_FACT'] = df_filtrado.apply(lambda x: x['CANTIDAD'] if 'NOTA DE CREDITO ELECTRONICA' in x['DOCUMENTO'] and x['RUT'] != 19 else 0, axis=1)
    df_filtrado['DEVOL_BOLETA'] = df_filtrado.apply(lambda x: x['CANTIDAD'] if 'NOTA DE CREDITO ELECTRONICA' in x['DOCUMENTO'] and x['RUT'] == 19 else 0, axis=1)

    df_agrupado = df_filtrado.groupby(['CODIGO', 'PRIMERA DESCRIPCION', 'PRIMERA LINEA', 'FAMILIA']).agg({
        'TRANS_FACT': 'sum',
        'TRANS_BOLETA': 'sum',
        'CANT_FACT': 'sum',
        'CANT_BOLETA': 'sum',
        'DEVOL_FACT': 'sum',
        'DEVOL_BOLETA': 'sum',
    }).reset_index()

    df_agrupado['TOTAL_FACT'] = df_agrupado['CANT_FACT'] + df_agrupado['DEVOL_FACT']
    df_agrupado['TOTAL_BOLETA'] = df_agrupado['CANT_BOLETA'] + df_agrupado['DEVOL_BOLETA']

    df_agrupado.rename(columns={
        'CODIGO': 'Código',
        'PRIMERA DESCRIPCION': 'Descripción',
        'PRIMERA LINEA': 'Línea',
        'FAMILIA': 'Familia',
        'TRANS_FACT': 'Transacciones Factura',
        'TRANS_BOLETA': 'Transacciones Boleta',
        'TOTAL_FACT': 'Total de Ventas Factura',
        'TOTAL_BOLETA': 'Total de Ventas Boleta',
    }, inplace=True)

    df_agrupado.drop(['CANT_FACT', 'CANT_BOLETA', 'DEVOL_FACT', 'DEVOL_BOLETA'], axis=1, inplace=True)

    return df_agrupado

def procesar(archivo, archivo_destino, archivo_destino_negativos):
    df = leer_archivo(archivo)
    df['DOCUMENTO'] = df['DOCUMENTO'].str.strip().str.upper()
    df['DESCRIPCION CODIGO'] = df['DESCRIPCION CODIGO'].str.strip().str.upper()
    df['LINEA'] = df['LINEA'].str.strip().str.upper()
    df['FAMILIA'] = df['FAMILIA'].str.strip().str.upper()
    df['COND.PAGO'] = df['COND.PAGO'].str.strip().str.upper()

    df_filtrado = df[
        (df['ESTADO'] != 'NULO') &
        (df['COND.PAGO'].str.strip() != 'TRASLADOS') &
        (df['CODIGO'] != 99999)
    ].copy()
    df_filtrado.rename(columns={'NETO': 'Ingresos'}, inplace=True)
    df_desglose = agregar_desglose_documentos(df_filtrado)

    df_901 = df_filtrado[df_filtrado['CODIGO'] == 901]
    df_filtrado = df_filtrado[df_filtrado['CODIGO'] != 901]
    df_filtrado = distribuir_total_negativo(df_filtrado, df_901)
    
    df_filtrado['PRIMERA DESCRIPCION'] = df_filtrado['CODIGO'].map(obtener_primera_descripcion(df_filtrado))
    df_filtrado['PRIMERA LINEA'] = df_filtrado['CODIGO'].map(obtener_primera_linea(df_filtrado))
    
    df_filtrado['Ingresos'] = df_filtrado['Ingresos'].round(0).astype(int)
    df_filtrado['COSTO FINAL'] = df_filtrado['COSTO FINAL'].round(0).astype(int)

    df_filtrado['UTILIDAD'] = df_filtrado['Ingresos'] - df_filtrado['COSTO FINAL']
    df_filtrado['Transacciones'] = df_filtrado.apply(calcular_Transacciones, axis=1)

    df_negativos = filtrar_datos_negativos(df_filtrado)
    
    df_productos = df_filtrado.groupby(['CODIGO', 'PRIMERA DESCRIPCION', 'PRIMERA LINEA', 'FAMILIA']).agg({
        'CANTIDAD': 'sum',
        'Ingresos': 'sum',
        'COSTO FINAL': 'sum',
        'UTILIDAD': 'sum',
        'Transacciones': 'sum' 
    }).reset_index()
    
    df_negativos_productos = filtrar_datos_negativos(df_productos)
    df_productos = asignar_etiquetas(df_productos)
    
    df_lineas = df_productos.groupby(['PRIMERA LINEA','FAMILIA']).agg({
        'CANTIDAD': 'sum',
        'Ingresos': 'sum',
        'COSTO FINAL': 'sum',
        'UTILIDAD': 'sum',
        'Transacciones': 'sum' 
    }).reset_index()
    
    df_negativos_lineas = filtrar_datos_negativos(df_lineas)
    df_lineas = asignar_etiquetas(df_lineas)
    
    df_familias = df_productos.groupby('FAMILIA').agg({
        'CANTIDAD': 'sum',
        'Ingresos': 'sum',
        'COSTO FINAL': 'sum',
        'UTILIDAD': 'sum',
        'Transacciones': 'sum' 
    }).reset_index()
    
    df_negativos_familias = filtrar_datos_negativos(df_familias)
    df_familias = asignar_etiquetas(df_familias)

    df_productos = renombrar_columnas(df_productos)
    df_lineas = renombrar_columnas(df_lineas)
    df_familias = renombrar_columnas(df_familias)
    df_negativos_productos = renombrar_columnas(df_negativos_productos)
    df_negativos_lineas = renombrar_columnas(df_negativos_lineas)
    df_negativos_familias = renombrar_columnas(df_negativos_familias)


    with pd.ExcelWriter(archivo_destino, engine='xlsxwriter') as writer:
        workbook = writer.book
        format_dollar = workbook.add_format({'num_format': '$#,##0.00'})  
        bold_format = workbook.add_format({'bold': True})

        sheets_info = [
            (df_productos, 'Productos', ['F', 'G', 'H']),
            (df_lineas, 'Lineas', ['D', 'E', 'F']),
            (df_familias, 'Familias', ['C', 'D', 'E'])
        ]
        
        for df, sheet_name, cols_format in sheets_info:
            worksheet = workbook.add_worksheet(sheet_name)
            writer.sheets[sheet_name] = worksheet
            
            for col_num, column_title in enumerate(df.columns):
                worksheet.write(0, col_num, column_title, bold_format)
            
            worksheet.autofilter(1, 0, 1, len(df.columns) - 1)
            
            for row_num, row in enumerate(df.values, start=2):  
                for col_num, value in enumerate(row):
                    col_letter = chr(col_num + 65)  
                    if col_letter in cols_format:  
                        worksheet.write(row_num, col_num, value, format_dollar)
                    else:
                        worksheet.write(row_num, col_num, value)
            for i, column in enumerate(df.columns):
                max_len = max(max([len(str(x)) for x in df[column]] + [len(column)]), 10)  
                worksheet.set_column(i, i, max_len + 1)  
    
    archivo_destino_negativos = archivo_destino_negativos.replace('.xlsx', '') + " - Negativos.xlsx"
    with pd.ExcelWriter(archivo_destino_negativos, engine='xlsxwriter') as writer:
        workbook = writer.book
        format_dollar = workbook.add_format({'num_format': '$#,##0.00'})  
        bold_format = workbook.add_format({'bold': True})

        for df, sheet_name, cols_format in [
            (df_negativos_productos, 'Productos Negativos', ['F', 'G', 'H']),
            (df_negativos_lineas, 'Lineas Negativas', ['D', 'E', 'F']),
            (df_negativos_familias, 'Familias Negativas', ['C', 'D', 'E'])
        ]:
            worksheet = workbook.add_worksheet(sheet_name)
            writer.sheets[sheet_name] = worksheet
            
            for col_num, column_title in enumerate(df.columns):
                worksheet.write(0, col_num, column_title, bold_format)  
            worksheet.autofilter(1, 0, 1, len(df.columns) - 1)  
            
            for row_num, row in enumerate(df.values):
                for col_num, value in enumerate(row):
                    if df.columns[col_num] in cols_format:  
                        worksheet.write(row_num + 2, col_num, value, format_dollar)
                    else:
                        worksheet.write(row_num + 2, col_num, value)
            for i, column in enumerate(df.columns):
                max_len = max(max([len(str(x)) for x in df[column]] + [len(column)]), 10)  
                worksheet.set_column(i, i, max_len + 1)  


    archivo_destino_desglose = archivo_destino.replace('.xlsx', '') + " - Desglose Ventas.xlsx"
    with pd.ExcelWriter(archivo_destino_desglose, engine='xlsxwriter') as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet('Desglose de Ventas')
        writer.sheets['Desglose de Ventas'] = worksheet

        for col_num, column_title in enumerate(df_desglose.columns):
            worksheet.write(0, col_num, column_title)  
        worksheet.autofilter(1, 0, 1, len(df_desglose.columns) - 1)  
        
        for row_num, row in enumerate(df_desglose.values, start=2):
            for col_num, value in enumerate(row):
                worksheet.write(row_num, col_num, value)
        
        for i, column in enumerate(df_desglose.columns):
            max_len = max(df_desglose[column].astype(str).map(len).max(), len(column))
            worksheet.set_column(i, i, max_len + 1)

    print("Procesamiento completado y guardado en:", archivo_destino)
    print("Datos negativos guardados en:", archivo_destino_negativos)
    print("Ventas desglosadas guardados en:", archivo_destino_desglose)

# INICIO DE PROGRAMA
#archivo = 'VENTAS AÑO-2023.xlsx'
#archivo_destino = 'Ventas_totales año-2023.xlsx'
#archivo_destino_negativos = 'Prod_negativos_2023.xlsx'
#procesar(archivo, archivo_destino, archivo_destino_negativos)

def abrir_pdf():
    nombre_pdf = "Análisis de productos1.pdf"
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

def abrir_pdf1(nombre_pdf):
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

def cargar_imagen(nombre_imagen):
    if getattr(sys, 'frozen', False):
        basedir = sys._MEIPASS
    else:
        basedir = os.path.dirname(__file__)
    imagen_path = os.path.join(basedir, nombre_imagen)
    return tk.PhotoImage(file=imagen_path)

def gui():
    def seleccionar_archivo():
        archivo = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls *.csv")])
        entry_archivo.delete(0, tk.END)
        entry_archivo.insert(0, archivo)

    def seleccionar_archivo_destino():
        archivo_destino = filedialog.asksaveasfilename(title="Especificar archivo de salida", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if archivo_destino:  
            entry_archivo_destino.delete(0, tk.END)
            entry_archivo_destino.insert(0, archivo_destino)

    def procesar_archivo():
        archivo = entry_archivo.get()
        archivo_destino = entry_archivo_destino.get()
        
        if archivo and archivo_destino:
            try:
                procesar(archivo, archivo_destino, archivo_destino)  
                messagebox.showinfo("Éxito", "El archivo ha sido procesado correctamente.")

            except Exception as e:
                messagebox.showerror("Error", f"Ocurrió un error al procesar el archivo: {e}")
        else:
            messagebox.showwarning("Advertencia", "Por favor, complete todos los campos.")

    root = tk.Tk()
    root.title("Análisis de productos")
    root.geometry('590x210')  
    root.configure(bg='#2D2D2D')

    
    fondo_columna = tk.Frame(root, bg='white', width= 2)
    fondo_columna.grid(row=3, column=3, rowspan=3, sticky='nswe', pady=0)

    tk.Label(root, text="Selección de Archivos", anchor='center', 
        font=('Arial', 12, 'bold'), background='light coral', foreground='white').grid(row=3, column=0, padx=0, pady=0, sticky='we', columnspan=3)

    tk.Label(root, text="Guías", anchor='center', 
        font=('Arial', 12, 'bold'), background='light coral', foreground='white').grid(row=3, column=4, padx=0, pady=0, sticky='we', columnspan=5)



    logo = cargar_imagen('Perno1.png')  
    logo_label = ttk.Label(root, image=logo, background='#2D2D2D')
    logo_label.grid(row=0, column=0, padx=5, pady=10, rowspan=3, sticky='we')

    style = ttk.Style()
    style.configure('TLabel', font=('Arial', 11), background='#2D2D2D', foreground='white')
    style.configure('TButton', font=('Arial', 11))
    style.configure('TEntry', font=('Arial', 11))
    style.configure('TCheckbutton', font=('Arial', 11))

    for i in range(3):
        root.grid_rowconfigure(i, minsize=10)
        root.grid_columnconfigure(0, minsize=10)

    ttk.Label(root, text="Archivo de Ventas:", background='#2D2D2D', foreground='white').grid(row=4, column=0, sticky='w', padx=10)
    entry_archivo = ttk.Entry(root, width=25)
    entry_archivo.grid(row=4, column=1, sticky='we', padx=10, pady=8)

    ttk.Button(root, text="Seleccionar", command=seleccionar_archivo).grid(row=4, column=2, padx=10, pady=8)

    ttk.Label(root, text="Archivo de Salida:", background='#2D2D2D', foreground='white').grid(row=5, column=0, sticky='w', padx=10)
    entry_archivo_destino = ttk.Entry(root, width=25)
    entry_archivo_destino.grid(row=5, column=1, sticky='we', padx=10, pady=8)

    ttk.Button(root, text="Especificar", command=seleccionar_archivo_destino).grid(row=5, column=2, padx=10, pady=8)

    def abrir_interpretacion():
        abrir_pdf1("Análisis de productos2.pdf")
        abrir_pdf1("Análisis de productos3.pdf")

    ttk.Button(root, text="Interpretación", command=abrir_interpretacion).grid(row=5, column=5, padx=10, pady=5, sticky='we')

    ttk.Button(root, text="Manual de Usuario", command=abrir_pdf).grid(row=4, column=5, padx=10, pady=5, sticky='we')

    ttk.Button(root, text="Generar Reportes", command=procesar_archivo).grid(row=7, column=5, pady=15)

    root.mainloop()

if __name__ == "__main__":
    gui()
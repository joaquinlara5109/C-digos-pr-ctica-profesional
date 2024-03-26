#xpx2
#STOCK-VENTAS
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

def procesar_stocks(archivo_stocks):
    df_stock = leer_archivo(archivo_stocks)
    df_stock.columns = df_stock.columns.str.strip().str.replace(' ', '').str.replace('\n', '').str.replace('.', '')
    df_stock.rename(columns={
        'SCoq': 'St. Matriz',
        'SL/S': 'St. Serena',
        'SBI': 'St. BI',
        'SBod': 'Stock Bodega',
        'StockEmprea': 'Stock Empresa'
    }, inplace=True)
    columnas_interes = ['Grupo', 'Subgrupo', 'Codigo', 'Descripcion', 'St. Matriz', 'St. Serena', 'St. BI', 'Stock Bodega', 'Stock Empresa']
    df_reducido = df_stock[columnas_interes].copy()
    codigos_unicos_stock = df_reducido['Codigo'].unique()
    return df_reducido, codigos_unicos_stock

def procesar_ventas(archivo_ventas, codigos_validos):
    df_ventas = leer_archivo(archivo_ventas)

    df_ventas['DESCRIPCION CODIGO'] = df_ventas['DESCRIPCION CODIGO'].str.strip().str.upper()
    df_ventas['LINEA'] = df_ventas['LINEA'].str.strip().str.upper()
    df_ventas['FAMILIA'] = df_ventas['FAMILIA'].str.strip().str.upper()
    df_ventas['COND.PAGO'] = df_ventas['COND.PAGO'].str.strip().str.upper()

    df_filtrado = df_ventas[df_ventas['CODIGO'].isin(codigos_validos)]
    df_filtrado = df_filtrado[(df_ventas['ESTADO'] != 'NULO') & (df_ventas['COND.PAGO'] != 'TRASLADOS') & (df_ventas['CODIGO'] != 99999) & (df_ventas['CODIGO'] != 901) & (df_ventas['SUCURSAL'] != 'TRANSITO CASA MATRIZ') & (df_ventas['SUCURSAL'] != 'TRANSITO BODEGA CENTRAL')]
    df_agrupado = df_filtrado.groupby(['CODIGO', 'SUCURSAL'])['CANTIDAD'].sum().reset_index()
    df_pivot = df_agrupado.pivot(index='CODIGO', columns='SUCURSAL', values='CANTIDAD').reset_index().fillna(0)
    df_pivot.rename(columns={'CASA MATRIZ': 'Vent. Matriz', 'STOCK SERENA': 'Vent. Serena', 'BARRIO INDUSTRIAL': 'Vent. BI'}, inplace=True)
    return df_pivot

def determinar_accion(stock_actual, stock_bodega, promedio_ventas, ventas_actuales):
    if stock_actual > promedio_ventas and stock_actual > ventas_actuales:
        return 'Buen Estado'
    elif stock_actual > promedio_ventas and stock_actual <= ventas_actuales and stock_actual > 0:
        return 'Decidir'
    elif stock_actual > 0 and stock_actual <= promedio_ventas:
        return 'Decidir'
    elif stock_actual <= 0 and stock_bodega > promedio_ventas:
        return 'Trasladar'
    elif stock_actual <= 0 and stock_bodega <= promedio_ventas:
        return 'Comprar'
    elif stock_actual <= 0 and stock_bodega <= 0:
        return 'Comprar'
    #elif stock_actual > 0 and stock_bodega <= 0:
    #   return 'Bod sin St.'
    else:
        return 'Ninguna'

def combinar_ventas_stock1(df_stock, df_ventas, df_ventas_promedio):
    df_stock.rename(columns={'Codigo': 'CODIGO'}, inplace=True)
    
    df_combinado = pd.merge(df_ventas, df_stock, on='CODIGO', how='left')
    
    df_combinado['Compra M'] = np.nan
    df_combinado['Compra LS'] = np.nan
    df_combinado['Compra BI'] = np.nan


    df_ventas_promedio.rename(columns={
        'Promedio Vent.M': 'Prom. Vent.M',
        'Promedio Vent.S': 'Prom. Vent.LS',
        'Promedio Vent.BI': 'Prom. Vent.BI'
    }, inplace=True)

    df_combinado_final = pd.merge(df_combinado, df_ventas_promedio, on='CODIGO', how='left')

    df_combinado_final['Prom. Vent.M'].fillna(0, inplace=True)
    df_combinado_final['Prom. Vent.LS'].fillna(0, inplace=True)
    df_combinado_final['Prom. Vent.BI'].fillna(0, inplace=True)
    df_combinado_final['Acción M'] = df_combinado_final.apply(
        lambda x: determinar_accion(x['St. Matriz'], x['Stock Bodega'], x['Prom. Vent.M'], x['Vent. Matriz']), axis=1
    )
    df_combinado_final['Acción LS'] = df_combinado_final.apply(
        lambda x: determinar_accion(x['St. Serena'], x['Stock Bodega'], x['Prom. Vent.LS'], x['Vent. Serena']), axis=1
    )
    df_combinado_final['Acción BI'] = df_combinado_final.apply(
        lambda x: determinar_accion(x['St. BI'], x['Stock Bodega'], x['Prom. Vent.BI'], x['Vent. BI']), axis=1
    )    

    columnas_finales = [
        'Grupo', 'Subgrupo', 'CODIGO', 'Descripcion',
        'Prom. Vent.M', 'Vent. Matriz', 'St. Matriz', 'Acción M', 'Compra M',
        'Prom. Vent.LS', 'Vent. Serena', 'St. Serena', 'Acción LS', 'Compra LS',
        'Prom. Vent.BI', 'Vent. BI', 'St. BI', 'Acción BI', 'Compra BI',
        'Stock Bodega'
    ]
    
    df_final = df_combinado_final[columnas_finales]
    
    df_final.rename(columns={
        'Grupo': 'Familia', 
        'Subgrupo': 'Línea', 
        'CODIGO': 'Código', 
        'Descripcion': 'Descripción', 
        'Stock Bodega': 'St. Bod'
    }, inplace=True)
    
    return df_final

def combinar_ventas_stock2(df_stock, df_ventas, df_ventas_promedio):
    df_stock.rename(columns={'Codigo': 'CODIGO'}, inplace=True)
    df_combinado = pd.merge(df_stock, df_ventas, on='CODIGO', how='left').fillna(0)
    
    df_combinado['Compra M'] = np.nan
    df_combinado['Compra LS'] = np.nan
    df_combinado['Compra BI'] = np.nan

    df_combinado_final = pd.merge(df_combinado, df_ventas_promedio, on='CODIGO', how='left')

    df_combinado_final.rename(columns={
        'Promedio Vent.M': 'Prom. Vent.M',
        'Promedio Vent.S': 'Prom. Vent.LS',
        'Promedio Vent.BI': 'Prom. Vent.BI'
    }, inplace=True)

    df_combinado_final['Prom. Vent.M'].fillna(0, inplace=True)
    df_combinado_final['Prom. Vent.LS'].fillna(0, inplace=True)
    df_combinado_final['Prom. Vent.BI'].fillna(0, inplace=True)

    df_combinado_final['Acción M'] = df_combinado_final.apply(
        lambda x: determinar_accion(x['St. Matriz'], x['Stock Bodega'], x['Prom. Vent.M'], x['Vent. Matriz']), axis=1
    )
    df_combinado_final['Acción LS'] = df_combinado_final.apply(
        lambda x: determinar_accion(x['St. Serena'], x['Stock Bodega'], x['Prom. Vent.LS'], x['Vent. Serena']), axis=1
    )
    df_combinado_final['Acción BI'] = df_combinado_final.apply(
        lambda x: determinar_accion(x['St. BI'], x['Stock Bodega'], x['Prom. Vent.BI'], x['Vent. BI']), axis=1
    )

    columnas_finales = [
        'Grupo', 'Subgrupo', 'CODIGO', 'Descripcion',
        'Prom. Vent.M', 'Vent. Matriz', 'St. Matriz', 'Acción M', 'Compra M',
        'Prom. Vent.LS', 'Vent. Serena', 'St. Serena', 'Acción LS', 'Compra LS',
        'Prom. Vent.BI', 'Vent. BI', 'St. BI', 'Acción BI', 'Compra BI',
        'Stock Bodega'
    ]

    df_final = df_combinado_final[columnas_finales]
    
    df_final.rename(columns={
        'Grupo': 'Familia',
        'Subgrupo': 'Línea',
        'CODIGO': 'Código',
        'Descripcion': 'Descripción',
        'Stock Bodega': 'St. Bod'
    }, inplace=True)
    
    return df_final

def contar_meses_unicos(archivo_ventas):
    df_ventas = leer_archivo(archivo_ventas)
    
    df_ventas['MES'] = df_ventas['MES'].astype(str).str.strip()
    
    meses_mapeo = {
        'ENERO': 1, 'FEBRE': 2, 'MARZO': 3, 'ABRIL': 4, 'MAYO': 5, 'JUNIO': 6,
        'JULIO': 7, 'AGOST': 8, 'SEPTI': 9, 'OCTUB': 10, 'NOVIE': 11, 'DICIE': 12
    }
    df_ventas['MES'] = df_ventas['MES'].map(meses_mapeo)
    
    df_ventas.dropna(subset=['DIA', 'MES', 'AÑO'], inplace=True)
    
    df_ventas['DIA'] = df_ventas['DIA'].astype(int)
    df_ventas['MES'] = df_ventas['MES'].astype(int)
    df_ventas['AÑO'] = df_ventas['AÑO'].astype(int)

    df_ventas['FECHA'] = pd.to_datetime({'year': df_ventas['AÑO'], 'month': df_ventas['MES'], 'day': df_ventas['DIA']})
    
    numero_meses = df_ventas['FECHA'].dt.to_period('M').nunique()
    print("Número de meses",numero_meses)
    return numero_meses

def procesar_ventas_y_calcular_promedio(archivo_ventas_promedio, codigos_validos):
    df_pivot = procesar_ventas(archivo_ventas_promedio, codigos_validos)
    numero_meses = contar_meses_unicos(archivo_ventas_promedio)
    nombres_originales = ['Vent. Matriz', 'Vent. Serena', 'Vent. BI']

    nombres_nuevos = {
        'Vent. Matriz': 'Prom. Vent.M', 
        'Vent. Serena': 'Prom. Vent.LS', 
        'Vent. BI': 'Prom. Vent.BI'
    }
    
    df_pivot.rename(columns=nombres_nuevos, inplace=True)

    for col in nombres_nuevos.values():
        df_pivot[col] = (df_pivot[col] / numero_meses).round(0)    
    columnas_finales = ['CODIGO'] + list(nombres_nuevos.values())
    df_final = df_pivot[columnas_finales]
    
    return df_final

def generar_reporte(archivo_ventas, archivo_stocks, archivo_salida,archivo_ventas_promedio):
    df_stock, codigos_validos = procesar_stocks(archivo_stocks)
    df_ventas = procesar_ventas(archivo_ventas, codigos_validos)
    df_ventas_promedio = procesar_ventas_y_calcular_promedio(archivo_ventas_promedio, codigos_validos)
    df_final1 = combinar_ventas_stock1(df_stock, df_ventas, df_ventas_promedio)
    df_final2 = combinar_ventas_stock2(df_stock, df_ventas, df_ventas_promedio)
    
    with pd.ExcelWriter(archivo_salida, engine='xlsxwriter') as writer:
        df_final1.to_excel(writer, sheet_name='Venta-Stock Mes', index=False, header=False, startrow=2)
        df_final2.to_excel(writer, sheet_name='Venta-Stock General', index=False, header=False, startrow=2)

        workbook = writer.book
        worksheet1 = writer.sheets['Venta-Stock Mes']
        worksheet2 = writer.sheets['Venta-Stock General']

        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#D9D9D9',
            'border': 1
        })

        for col_num, value in enumerate(df_final1.columns.values):
            worksheet1.write(0, col_num, value, header_format)
            worksheet2.write(0, col_num, value, header_format)

        worksheet1.autofilter(1, 0, 1, len(df_final1.columns) - 1)
        worksheet2.autofilter(1, 0, 1, len(df_final2.columns) - 1)

        for worksheet in [worksheet1, worksheet2]:
            for i, col in enumerate(df_final1.columns):
                worksheet.set_column(i, i, max(len(col), 15))

        format_red = workbook.add_format({'bg_color': '#FF0000', 'font_color': '#FFFFFF'})
        format_yellow = workbook.add_format({'bg_color': '#FFFF00', 'font_color': '#000000'})

        max_row1 = len(df_final1) + 2
        max_row2 = len(df_final2) + 2

        st_bod_range1 = f'T3:T{max_row1}'
        st_bod_range2 = f'T3:T{max_row2}'
        worksheet1.conditional_format(st_bod_range1, {'type': 'cell', 'criteria': '==', 'value': 0, 'format': format_red})
        worksheet2.conditional_format(st_bod_range2, {'type': 'cell', 'criteria': '==', 'value': 0, 'format': format_red})

        for col, next_col in [('F', 'G'), ('K', 'L'), ('P', 'Q')]:
            next_col_range1 = f'{next_col}3:{next_col}{max_row1}'
            next_col_range2 = f'{next_col}3:{next_col}{max_row2}'

            worksheet1.conditional_format(next_col_range1, {'type': 'cell', 'criteria': '==', 'value': 0, 'format': format_red})
            worksheet2.conditional_format(next_col_range2, {'type': 'cell', 'criteria': '==', 'value': 0, 'format': format_red})

            formula1 = f'AND({col}3={next_col}3, {col}3<>0)'
            formula2 = f'AND({col}3={next_col}3, {col}3<>0)'

            worksheet1.conditional_format(next_col_range1, {'type': 'formula', 'criteria': formula1, 'format': format_yellow})
            worksheet2.conditional_format(next_col_range2, {'type': 'formula', 'criteria': formula2, 'format': format_yellow})

#archivo_ventas = 'documentos de ventas ENERO.xlsx'
#archivo_ventas_promedio = 'VENTAS AÑO-2023.xlsx'
#archivo_stocks = 'stock_por_sucursal_T0.csv'
#archivo_salida = 'Ventas-Stock_Decisión1z1x12xq.xlsx'
#generar_reporte(archivo_ventas, archivo_stocks, archivo_salida, archivo_ventas_promedio)
#print("Archivo:", archivo_salida, "generado con éxito")

def abrir_pdf():
    nombre_pdf = "Gestión de inventarios0.1.pdf"
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
    nombre_pdf = "Gestión de inventarios0.2.pdf"
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

    def seleccionar_archivo_ventas_promedio():
        archivo = filedialog.askopenfilename(title="Seleccionar archivo de ventas promedio", filetypes=[("Excel files", "*.xlsx *.xls *.csv")])
        entry_archivo_ventas_promedio.delete(0, tk.END)
        entry_archivo_ventas_promedio.insert(0, archivo)

    def seleccionar_archivo_stock():
        archivo = filedialog.askopenfilename(title="Seleccionar archivo de stock", filetypes=[("Excel files", "*.xlsx *.xls *.csv")])
        entry_archivo_stocks.delete(0, tk.END)
        entry_archivo_stocks.insert(0, archivo)

    def seleccionar_archivo_ventas():
        archivo = filedialog.askopenfilename(title="Seleccionar archivo de ventas último mes", filetypes=[("Excel files", "*.xlsx *.xls *.csv")])
        entry_archivo_ventas.delete(0, tk.END)
        entry_archivo_ventas.insert(0, archivo)

    def seleccionar_archivo_salida():
        archivo = filedialog.asksaveasfilename(title="Especificar archivo de salida", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        entry_archivo_salida.delete(0, tk.END)
        entry_archivo_salida.insert(0, archivo)

    def procesar_archivos():
        archivo_ventas_promedio = entry_archivo_ventas_promedio.get()
        archivo_stocks = entry_archivo_stocks.get()
        archivo_salida = entry_archivo_salida.get()
        archivo_ventas = entry_archivo_ventas.get()


        if archivo_ventas_promedio and archivo_stocks and archivo_salida:
            try:
                generar_reporte(archivo_ventas, archivo_stocks, archivo_salida,archivo_ventas_promedio)
                messagebox.showinfo("Éxito", "Los archivos han sido procesados correctamente.")

            except Exception as e:
                messagebox.showerror("Error", f"Ocurrió un error al procesar los archivos: {e}")
        else:
            messagebox.showwarning("Advertencia", "Por favor, complete todos los campos.")

    root = tk.Tk()
    root.title("Gestión de Inventarios")
    root.geometry('628x280')
    root.configure(bg='#2D2D2D')

    style = ttk.Style()
    style.configure('TLabel', font=('Arial', 11), background='#2D2D2D', foreground='white')
    style.configure('TButton', font=('Arial', 11))
    style.configure('TEntry', font=('Arial', 11))
    style.configure('TCheckbutton', font=('Arial', 11))


    fondo_columna = tk.Frame(root, bg='white', width= 2)
    fondo_columna.grid(row=3, column=3, rowspan=5, sticky='nswe', pady=0)

    tk.Label(root, text="Selección de Archivos", anchor='center', 
        font=('Arial', 12, 'bold'), background='light coral', foreground='white').grid(row=3, column=0, padx=0, pady=0, sticky='we', columnspan=3)

    tk.Label(root, text="Guías", anchor='center', 
        font=('Arial', 12, 'bold'), background='light coral', foreground='white').grid(row=3, column=4, padx=0, pady=0, sticky='we', columnspan=5)



    logo = cargar_imagen()  
    logo_label = ttk.Label(root, image=logo, background='#2D2D2D')
    logo_label.grid(row=0, column=0, padx=5, pady=5, rowspan=3, sticky='we')

    ttk.Label(root, text="Archivo de Ventas Promedio:", background='#2D2D2D', foreground='white').grid(row=4, column=0, padx=5, pady=5, sticky='w')
    entry_archivo_ventas_promedio = ttk.Entry(root, width=25)
    entry_archivo_ventas_promedio.grid(row=4, column=1, padx=5, pady=5, sticky='we')
    ttk.Button(root, text="Seleccionar", command=seleccionar_archivo_ventas_promedio).grid(row=4, column=2, padx=5, pady=5)

    ttk.Label(root, text="Archivo de Stock:", background='#2D2D2D', foreground='white').grid(row=5, column=0, padx=5, pady=5, sticky='w')
    entry_archivo_stocks = ttk.Entry(root, width=25)
    entry_archivo_stocks.grid(row=5, column=1, padx=5, pady=5, sticky='we')
    ttk.Button(root, text="Seleccionar", command=seleccionar_archivo_stock).grid(row=5, column=2, padx=5, pady=5)

    ttk.Label(root, text="Archivo ventas último mes:", background='#2D2D2D', foreground='white').grid(row=6, column=0, padx=5, pady=5, sticky='w')
    entry_archivo_ventas = ttk.Entry(root, width=25)
    entry_archivo_ventas.grid(row=6, column=1, padx=5, pady=5, sticky='we')
    ttk.Button(root, text="Seleccionar", command=seleccionar_archivo_ventas).grid(row=6, column=2, padx=5, pady=5)

    ttk.Label(root, text="Archivo de Salida:", background='#2D2D2D', foreground='white').grid(row=7, column=0, padx=5, pady=5, sticky='w')
    entry_archivo_salida = ttk.Entry(root, width=25)
    entry_archivo_salida.grid(row=7, column=1, padx=5, pady=5, sticky='we')
    ttk.Button(root, text="Especificar", command=seleccionar_archivo_salida).grid(row=7, column=2, padx=5, pady=5)

    ttk.Button(root, text="Interpretación", command=abrir_pdf1).grid(row=5, column=4, padx=10, pady=10, sticky='we')

    ttk.Button(root, text="Manual de Usuario", command=abrir_pdf).grid(row=4, column=4, padx=10, pady=5, sticky='we')

    ttk.Button(root, text="Generar Reporte", command=procesar_archivos).grid(row=8, column=4, padx=5, pady=20)

    root.mainloop()

if __name__ == "__main__":
    gui()

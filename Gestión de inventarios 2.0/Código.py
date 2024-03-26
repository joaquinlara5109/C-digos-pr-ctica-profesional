#New xpx2

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

def limpiar_datos(archivo_ventas_stock):
    df_filtrado = leer_archivo(archivo_ventas_stock)
    df_filtrado.columns = df_filtrado.columns.str.strip()

    columnas_nuevas = ['Acción M', 'Acción LS', 'Acción BI', 'Estado BodC', 'Compra M', 'Compra LS', 'Compra BI']
    for columna in columnas_nuevas:
        df_filtrado[columna] = np.nan

    df_filtrado1 = df_filtrado.rename(columns={
        'Linea': 'Línea',
        'Codigo': 'Código',
        'Descripcion': 'Descripción',
        'Rotacion': 'Rotación',
        'CM.Ventas 30': 'Prom. Vent.M',
        'Stock Cmatriz': 'St. Matriz',
        'LS.Ventas 30': 'Prom. Vent.LS',
        'Stock La Serena': 'St. LS',
        'BI.Ventas 30': 'Prom. Vent.BI',
        'Stock Barrio': 'St. BI',
        'Stock Bodega Central': 'St. BodC',
        'Ventas 30': 'Ventas promedio'
    })
    columnas_interes = [
        'Familia', 'Línea', 'Código', 'Descripción', 'Proveedor', 'Rotación',
        'Prom. Vent.M', 'St. Matriz', 'Acción M', 'Compra M',
        'Prom. Vent.LS', 'St. LS', 'Acción LS', 'Compra LS',
        'Prom. Vent.BI', 'St. BI', 'Acción BI', 'Compra BI',
        'St. BodC', 'Estado BodC', 'Ventas promedio'
    ]

    df_limpio = df_filtrado1[columnas_interes]

    return df_limpio

def rellenar_columnas_accion(df_filtrado, buen_estado):
    
    def determinar_accion(stock_actual, stock_bodega, stock_otras_sucursales, promedio_ventas, buen_estado):
        # Condiciones cuando está en buen estado
        if stock_actual > promedio_ventas * buen_estado:
            return 'Buen Estado'
        # Condiciones cuando debe decidirse una acción
        elif stock_actual > 0 and stock_actual <= promedio_ventas * buen_estado:
            return 'Decidir'
        # Condiciones para traslados desde la bodega central
        elif stock_actual <= 0 and stock_bodega >= promedio_ventas and stock_bodega > 0:
            return 'Tras. Bd'
        # Condiciones para evaluar traslados desde otras sucursales o comprar
        elif stock_actual <= 0:
            sucursales_candidatas = []
            for sucursal, info in stock_otras_sucursales.items():
                stock_sucursal = info['stock']
                # Asegurar que el promedio de ventas no sea cero para el cálculo, se reemplaza por 1 en ese caso
                promedio_ventas_sucursal = max(info['promedio_ventas'], 1)
                if stock_sucursal > promedio_ventas_sucursal * buen_estado and (stock_sucursal - (promedio_ventas_sucursal * buen_estado)) > (promedio_ventas - max(stock_actual, 0)):
                    # Calcular el excedente y su proporción
                    excedente = stock_sucursal - promedio_ventas_sucursal * buen_estado
                    proporcion_excedente = excedente / promedio_ventas_sucursal
                    sucursales_candidatas.append((sucursal, proporcion_excedente))

            if sucursales_candidatas:
                # Elegir la sucursal con mayor proporción de excedente
                mejor_sucursal = max(sucursales_candidatas, key=lambda item: item[1])[0]
                return f'Tras. {mejor_sucursal}'

            return 'Comprar'
        else:
            return 'Ninguna'

    def determinar_estado(stock_bodega, ventas_promedio, buen_estado):
        if stock_bodega > ventas_promedio * buen_estado:
            return 'Buen Estado'
        elif stock_bodega <= ventas_promedio * buen_estado and stock_bodega > ventas_promedio * 0.5:
            return 'Medio Mes'
        elif stock_bodega <= ventas_promedio * 0.5:
            return 'Estado Crítico'
        else:
            return 'Ninguno'

    df_filtrado['Acción M'] = df_filtrado.apply(
        lambda x: determinar_accion(
            x['St. Matriz'], x['St. BodC'], {'LS': {'stock': x['St. LS'], 'promedio_ventas': x['Prom. Vent.LS']}, 'BI': {'stock': x['St. BI'], 'promedio_ventas': x['Prom. Vent.BI']}}, x['Prom. Vent.M'],
        buen_estado), axis=1
    )
    df_filtrado['Acción LS'] = df_filtrado.apply(
        lambda x: determinar_accion(
            x['St. LS'], x['St. BodC'], {'M': {'stock': x['St. Matriz'], 'promedio_ventas': x['Prom. Vent.M']}, 'BI': {'stock': x['St. BI'], 'promedio_ventas': x['Prom. Vent.BI']}}, x['Prom. Vent.LS'],
        buen_estado), axis=1
    )
    df_filtrado['Acción BI'] = df_filtrado.apply(
        lambda x: determinar_accion(
            x['St. BI'], x['St. BodC'], {'M': {'stock': x['St. Matriz'], 'promedio_ventas': x['Prom. Vent.M']}, 'LS': {'stock': x['St. LS'], 'promedio_ventas': x['Prom. Vent.LS']}}, x['Prom. Vent.BI'],
        buen_estado), axis=1
    )

    df_filtrado['Estado BodC'] = df_filtrado.apply(
        lambda x: determinar_estado(
            x['St. BodC'], x['Ventas promedio'],
        buen_estado), axis=1
    )

    return df_filtrado

def generar_reporte(archivo_ventas_stock, archivo_destino, buen_estado):
    df_filtrado = limpiar_datos(archivo_ventas_stock)
    df_acciones = rellenar_columnas_accion(df_filtrado, buen_estado)

    df_acciones = df_acciones.drop(columns=['Ventas promedio'])

    df_acciones.insert(20, 'Suma Compras', None)

    with pd.ExcelWriter(archivo_destino, engine='xlsxwriter') as writer:
        df_acciones.to_excel(writer, sheet_name='Venta-Stock', startrow=2, header=False, index=False)
        
        workbook = writer.book
        worksheet = writer.sheets['Venta-Stock']
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#D9D9D9',
            'border': 1
        })

        # Formatos condicionales
        red_format = workbook.add_format({'bg_color': '#FFC7CE'})
        orange_format = workbook.add_format({'bg_color': '#E89907'})

        green_format = workbook.add_format({'bg_color': '#C6EFCE'})
        yellow_format = workbook.add_format({'bg_color': '#FFEB9C'})

        for col_num, value in enumerate(df_acciones.columns.values):
            worksheet.write(0, col_num, value, header_format)

        worksheet.autofilter(1, 0, 1, len(df_acciones.columns) - 1)

        for i, col in enumerate(df_acciones.columns):
            width = max([len(str(x)) for x in df_acciones[col].values] + [len(col)]) + 1
            worksheet.set_column(i, i, width)


        for row_num in range(3, len(df_acciones) + 3):
            worksheet.write_formula(row_num - 1, 20, f'=SUM(J{row_num+1},N{row_num+1},R{row_num+1})')

        column_indexes = [7, 11, 15, 18]  
        for col_idx in column_indexes:
            worksheet.conditional_format(2, col_idx, len(df_acciones) + 2, col_idx, 
                                        {'type': 'cell', 
                                        'criteria': '<=', 
                                        'value': 0, 
                                        'format': red_format})


        col_t_index = 19  
        worksheet.conditional_format(2, col_t_index, len(df_acciones) + 2, col_t_index, 
                                     {'type': 'cell', 'criteria': 'equal to', 'value': '"Buen Estado"', 'format': green_format})
        worksheet.conditional_format(2, col_t_index, len(df_acciones) + 2, col_t_index, 
                                     {'type': 'cell', 'criteria': 'equal to', 'value': '"Medio Mes"', 'format': yellow_format})
        worksheet.conditional_format(2, col_t_index, len(df_acciones) + 2, col_t_index, 
                                     {'type': 'cell', 'criteria': 'equal to', 'value': '"Estado Crítico"', 'format': orange_format})

#archivo_ventas_stock = 'ventas_stock.xlsx'
#archivo_destino = 'ASX7.xlsx'
#generar_reporte(archivo_ventas_stock, archivo_destino)

def abrir_pdf():
    nombre_pdf = "Gestión de inventarios3.pdf"
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
        buen_estado = float(entry_nivel_estado.get()) if entry_nivel_estado.get() else 1.65

        
        if archivo and archivo_destino:
            try:
                generar_reporte(archivo, archivo_destino, buen_estado)  
                messagebox.showinfo("Éxito", "El archivo ha sido procesado correctamente.")

            except Exception as e:
                messagebox.showerror("Error", f"Ocurrió un error al procesar el archivo: {e}")
        else:
            messagebox.showwarning("Advertencia", "Por favor, complete todos los campos.")

    root = tk.Tk()
    root.title("Gestión de Inventarios")
    root.geometry('651x265')  
    root.configure(bg='#2D2D2D')

    
    fondo_columna = tk.Frame(root, bg='white', width= 2)
    fondo_columna.grid(row=3, column=3, rowspan=6, sticky='nswe', pady=0)

    tk.Label(root, text="Selección de Archivos", anchor='center', 
        font=('Arial', 12, 'bold'), background='light coral', foreground='white').grid(row=3, column=0, padx=0, pady=0, sticky='we', columnspan=3)

    tk.Label(root, text="Guías", anchor='center', 
        font=('Arial', 12, 'bold'), background='light coral', foreground='white').grid(row=3, column=4, padx=0, pady=0, sticky='we', columnspan=5)

    
    tk.Label(root, text="Elección de parámetro", anchor='center', 
        font=('Arial', 12, 'bold'), background='light coral', foreground='white').grid(row=7, column=0, padx=0, pady=0, sticky='we', columnspan=3)



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

    ttk.Label(root, text="Archivo de Stocks y Ventas:", background='#2D2D2D', foreground='white').grid(row=4, column=0, sticky='w', padx=10)
    entry_archivo = ttk.Entry(root, width=25)
    entry_archivo.grid(row=4, column=1, sticky='we', padx=10, pady=8)

    ttk.Button(root, text="Seleccionar", command=seleccionar_archivo).grid(row=4, column=2, padx=10, pady=8)

    ttk.Label(root, text="Archivo de Salida:", background='#2D2D2D', foreground='white').grid(row=5, column=0, sticky='w', padx=10)
    entry_archivo_destino = ttk.Entry(root, width=25)
    entry_archivo_destino.grid(row=5, column=1, sticky='we', padx=10, pady=8)

    ttk.Button(root, text="Especificar", command=seleccionar_archivo_destino).grid(row=5, column=2, padx=10, pady=8)

    def abrir_interpretacion():
        abrir_pdf1("Gestión de inventarios4.pdf")

    ttk.Button(root, text="Interpretación", command=abrir_interpretacion).grid(row=5, column=5, padx=10, pady=5, sticky='we')

    ttk.Button(root, text="Manual de Usuario", command=abrir_pdf).grid(row=4, column=5, padx=10, pady=5, sticky='we')

    ttk.Button(root, text="Generar Reporte", command=procesar_archivo).grid(row=9, column=5, pady=10)

    ttk.Label(root, text="Buen estado:", background='#2D2D2D', foreground='white').grid(row=8, column=0, padx=10, pady=5, sticky='w')
    entry_nivel_estado = ttk.Entry(root, width=25)
    entry_nivel_estado.insert(0, "1.05")  
    entry_nivel_estado.grid(row=8, column=1, padx=10, pady=8, sticky='we')

    root.mainloop()

if __name__ == "__main__":
    gui()
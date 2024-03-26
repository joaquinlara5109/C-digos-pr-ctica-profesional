#Test Batería_
import numpy as np
import pandas as pd
from sklearn.preprocessing import MinMaxScaler
from sklearn.model_selection import train_test_split
from tensorflow.keras.models import Sequential
from tensorflow.keras.layers import Dense, LSTM
from tensorflow.keras.callbacks import EarlyStopping
import matplotlib.pyplot as plt
import tensorflow as tf
from datetime import datetime, timedelta
import csv
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import subprocess
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

def eliminar_columnas(df):
    df['DESCRIPCION CODIGO'] = df['DESCRIPCION CODIGO'].str.strip().str.upper()
    df['LINEA'] = df['LINEA'].str.strip().str.upper()
    df['FAMILIA'] = df['FAMILIA'].str.strip().str.upper()
    df['COND.PAGO'] = df['COND.PAGO'].str.strip().str.upper()
    df_filtrado = df[
        (df['COND.PAGO'].str.strip() != 'TRASLADOS') &
        (df['ESTADO'] != 'NULO') &
        (df['CODIGO'] != 901) &
        (df['CODIGO'] != 99999)
    ].copy()
    columnas_para_conservar = ['DIA', 'MES', 'AÑO', 'ESTADO', 'COND.PAGO', 'CODIGO', 'DESCRIPCION CODIGO', 'CANTIDAD', 'FAMILIA', 'LINEA']
    df_filtrado = df_filtrado[columnas_para_conservar]
    
    df_filtrado['CANTIDAD'] = pd.to_numeric(df_filtrado['CANTIDAD'], errors='coerce').fillna(0)
    
    return df_filtrado

def generar_semanas_y_mapear_meses(dfnuevo):
    meses_mapeo = {
        'ENERO': 1, 'FEBRE': 2, 'MARZO': 3, 'ABRIL': 4, 'MAYO': 5, 'JUNIO': 6,
        'JULIO': 7, 'AGOST': 8, 'SEPTI': 9, 'OCTUB': 10, 'NOVIE': 11, 'DICIE': 12
    }
    dfnuevo['MES'] = dfnuevo['MES'].str.strip().map(meses_mapeo)
    dfnuevo['FECHA'] = pd.to_datetime(dfnuevo.assign(DAY=dfnuevo['DIA'], MONTH=dfnuevo['MES'], YEAR=dfnuevo['AÑO']).loc[:, ['YEAR','MONTH','DAY']])
    dfnuevo.sort_values(by='FECHA', inplace=True)

"""def encontrar_codigos_por_familia_y_linea(df, nombre_familia, linea_producto):
    df_filtrado = df[
        (df['FAMILIA'].str.strip().str.upper() == nombre_familia) &
        (df['LINEA'].str.strip().str.upper() == linea_producto)
    ]
    return df_filtrado['CODIGO'].unique()
"""
#
def sumar_cantidades_por_semana_y_mes(df_unificado, codigos_validos):
    df_filtrado = df_unificado[df_unificado['CODIGO'].isin(codigos_validos)]
    
    df_filtrado['Semana'] = df_filtrado['FECHA'].dt.isocalendar().week
    df_filtrado['Mes-Año'] = df_filtrado['FECHA'].dt.strftime('%Y-%m')
    df_semanal = df_filtrado.groupby(['Semana', 'CODIGO'])['CANTIDAD'].sum().unstack(fill_value=0)
    df_mensual = df_filtrado.groupby(['Mes-Año', 'CODIGO'])['CANTIDAD'].sum().unstack(fill_value=0)
    
    return df_semanal, df_mensual

def crear_df_descripciones(df_unificado, codigos_validos):

    df_descripciones = df_unificado[df_unificado['CODIGO'].isin(codigos_validos)][['CODIGO', 'DESCRIPCION CODIGO']]
    df_descripciones = df_descripciones.drop_duplicates().reset_index(drop=True)
    df_descripciones.rename(columns={'CODIGO': 'Codigo', 'DESCRIPCION CODIGO': 'Descripcion'}, inplace=True)
    return df_descripciones

#### POST SERIES DE TIEMPO

def detectar_y_reemplazar_outliers(serie_tiempo, method='linear'):
    z_scores = (serie_tiempo - serie_tiempo.mean()) / serie_tiempo.std()
    threshold = 3
    outliers_idx = np.where(np.abs(z_scores) > threshold)[0]
    tenth_percentile = np.percentile(serie_tiempo, 5)
    ninetieth_percentile = np.percentile(serie_tiempo, 95)
    for idx in outliers_idx:
        if serie_tiempo[idx] < tenth_percentile:
            serie_tiempo[idx] = tenth_percentile
        elif serie_tiempo[idx] > ninetieth_percentile:
            serie_tiempo[idx] = ninetieth_percentile
    serie_interpolada = serie_tiempo.interpolate(method=method)
    return serie_interpolada

def preparar_datos(datos):
    data = np.array(datos).reshape(-1, 1)
    scaler = MinMaxScaler(feature_range=(0, 1))
    data_scaled = scaler.fit_transform(data)
    return data_scaled, scaler

def crear_conjunto_datos(dataset, look_back):
    X, Y = [], []
    for i in range(len(dataset)-look_back-1):
        a = dataset[i:(i+look_back), 0]
        X.append(a)
        Y.append(dataset[i + look_back, 0])
    return np.array(X), np.array(Y)

def entrenar_modelo(X, Y, look_back):
    X = np.reshape(X, (X.shape[0], 1, X.shape[1]))
    X_train, X_test, Y_train, Y_test = train_test_split(X, Y, test_size=0.2, random_state=12)
    model = Sequential()
    model.add(LSTM(200, input_shape=(1, look_back)))
    model.add(Dense(1))
    model.compile(optimizer='adam', loss='mean_squared_error')
    early_stop = EarlyStopping(monitor='val_loss', patience=25, verbose=1)
    history = model.fit(X_train, Y_train, epochs=200, batch_size=24, verbose=1, validation_split=0.2, callbacks=[early_stop])
    return model, history

"""def graficar_perdida(history):
    plt.figure(figsize=(10, 6))
    plt.plot(history.history['loss'], label='Pérdida de Entrenamiento')
    plt.plot(history.history['val_loss'], label='Pérdida de Validación')
    plt.title('Pérdida de Entrenamiento y Validación a lo Largo de las Épocas')
    plt.xlabel('Épocas')
    plt.ylabel('Pérdida')
    plt.legend()
    plt.show()
"""
def predecir_futuro_dataframe(model, data, future_periods, scaler, look_back, codigo):
    input_seq = data[-look_back:]
    future_predictions = []
    for _ in range(future_periods):
        prediction = model.predict(input_seq.reshape(1, 1, look_back))
        future_predictions.append(prediction[0,0])
        input_seq = np.roll(input_seq, -1)
        input_seq[-1] = prediction
    future_predictions = scaler.inverse_transform(np.array(future_predictions).reshape(-1, 1)).flatten()
    return pd.DataFrame({'Codigo': codigo, 'Prediccion': future_predictions})

"""def graficar_predicciones(datos_reales, future_predictions, future_periods):
    plt.figure(figsize=(10,6))
    plt.plot(np.arange(len(datos_reales)), datos_reales, label="Datos reales")
    plt.plot(np.arange(len(datos_reales), len(datos_reales) + future_periods), future_predictions.reshape(-1), label="Predicciones futuras", marker='o')
    plt.legend()
    plt.show()
"""
def exportar_predicciones_y_series_de_tiempo_con_xlsxwriter(predicciones, df_mensual, df_descripciones, nombre_archivo):
    predicciones['Codigo'] = predicciones['Codigo'].astype(int)
    df_descripciones['Codigo'] = df_descripciones['Codigo'].astype(int)
    df_descripciones = df_descripciones.drop_duplicates(subset='Codigo', keep='first')
    
    with pd.ExcelWriter(nombre_archivo, engine='xlsxwriter') as writer:
        predicciones_organizadas = predicciones.pivot(index='Codigo', columns='Fecha', values='Prediccion').reset_index()
        predicciones_organizadas.columns = ['Codigo'] +[f"Periodo {i+1}" for i in range(len(predicciones_organizadas.columns)-1)]
        series_de_tiempo = pd.DataFrame()
        for codigo in df_mensual.columns:
            serie = df_mensual[codigo]
            serie_df = pd.DataFrame(serie)
            serie_df['Codigo'] = codigo
            serie_df.rename(columns={codigo: 'Valor'}, inplace=True)
            serie_df['Fecha'] = serie_df.index.strftime('%m-%Y')
            series_de_tiempo = pd.concat([series_de_tiempo, serie_df])

        series_de_tiempo = series_de_tiempo.pivot(index='Codigo', columns='Fecha', values='Valor').reset_index()
        fechas_columnas = series_de_tiempo.columns[1:]  

        fechas_columnas_dt = pd.to_datetime(fechas_columnas, format='%m-%Y')  
        fechas_columnas_dt_sorted = fechas_columnas_dt.sort_values()
        fechas_columnas_sorted_str = fechas_columnas_dt_sorted.strftime('%m-%Y')
        columnas_finales = ['Codigo'] + list(fechas_columnas_sorted_str)
        series_de_tiempo= series_de_tiempo.reindex(columns=columnas_finales)
        series_de_tiempo.columns.name = None
        ultimo_mes = series_de_tiempo.columns[-1]  
        ultimo_mes = datetime.strptime(ultimo_mes, '%m-%Y')
        nombres_periodos_futuros = [(ultimo_mes + timedelta(days=32*i)).strftime('%m-%Y') for i in range(1, 5)]

        columnas_meses = series_de_tiempo.columns[1:]  
        media_12_meses = series_de_tiempo[columnas_meses[-12:]].mean(axis=1)
        media_6_meses = series_de_tiempo[columnas_meses[-6:]].mean(axis=1)

        resultado_12_meses = pd.DataFrame({'Codigo': series_de_tiempo['Codigo'],'Media 12 Meses': media_12_meses})
        resultado_6_meses = pd.DataFrame({'Codigo': series_de_tiempo['Codigo'],'Media 6 Meses': media_6_meses})
        resultado_media = pd.merge(resultado_12_meses, resultado_6_meses, on='Codigo', how='left')
        resultado_media = pd.merge(resultado_media, df_descripciones, on='Codigo', how='left')

        predicciones_organizadas_final = pd.merge(resultado_media, predicciones_organizadas, on='Codigo', how='left')
        print(predicciones_organizadas_final)


        columnas_interes = ['Codigo', 'Descripcion', 'Media 12 Meses', 'Media 6 Meses','Periodo 1', 'Periodo 2', 'Periodo 3', 'Periodo 4']
        predicciones_organizadas_final = predicciones_organizadas_final[columnas_interes]


        predicciones_organizadas_final.rename(
            columns={
        'Codigo': 'Código', 
        'Descripcion': 'Descripción',
        'Periodo 1': nombres_periodos_futuros[0],
        'Periodo 2': nombres_periodos_futuros[1],
        'Periodo 3': nombres_periodos_futuros[2],
        'Periodo 4': nombres_periodos_futuros[3]
        }, inplace=True)

        series_de_tiempo['Codigo'] = series_de_tiempo['Codigo'].astype(int)
        series_de_tiempo_con_descripciones = pd.merge(series_de_tiempo, df_descripciones[['Codigo', 'Descripcion']], on='Codigo', how='left')

        columnas_principales = ['Codigo', 'Descripcion']
        columnas_fechas = [col for col in series_de_tiempo_con_descripciones.columns if col not in columnas_principales]
        series_de_tiempo_final = series_de_tiempo_con_descripciones[columnas_principales + columnas_fechas]
        series_de_tiempo_final.rename(columns={'Codigo':'Código', 'Descripcion':'Descripción'}, inplace=True)


        predicciones_organizadas_final.to_excel(writer, sheet_name='Predicciones', index=False)
        series_de_tiempo_final.to_excel(writer, sheet_name='Últimas Observaciones', index=False)

        workbook  = writer.book
        worksheet_predicciones = writer.sheets['Predicciones']
        formato_un_decimal = workbook.add_format({'num_format': '0.0'})
        worksheet_predicciones.set_column('C:H', None, formato_un_decimal)

    print(f"Archivo '{nombre_archivo}' creado con éxito con XlsxWriter.")

def funcion_principal(nombres_archivos, archivo_destino, codigos_adicionales=[]):

    codigos_validos = [
    58012, 68746, 68747, 68748, 68749, 68750, 68751, 69517, 69518, 69520,
    69521, 69523, 69700, 70205, 70206, 73202, 75515, 75516, 75523, 76042,
    78139, 78140, 78141, 78142, 78143, 79133, 79134, 81707, 81708, 81709,
    84631, 88391, 88394, 98864, 98865, 98866, 98867
    ]


    codigos_validos.extend(codigos_adicionales)

    dfs_procesados = []
    for nombre_base in nombres_archivos:
            print(nombre_base)
            df = leer_archivo(nombre_base)
            df_procesado = eliminar_columnas(df)
            dfs_procesados.append(df_procesado)

    if dfs_procesados:
        df_unificado = pd.concat(dfs_procesados, ignore_index=True)
        generar_semanas_y_mapear_meses(df_unificado) 
        df_descripciones = crear_df_descripciones(df_unificado, codigos_validos)
        if len(codigos_validos) > 0:
            df_semanal, df_mensual = sumar_cantidades_por_semana_y_mes(df_unificado, codigos_validos)
        else:
            print("No se encontraron productos para la familia y línea ingresadas.")
    else:
        print("Archivo no encontrado. Por favor, reintenta.")

    df_mensual.index = pd.to_datetime(df_mensual.index + "-01")  
    look_back = 4
    predicciones_totales = pd.DataFrame()  
    for codigo in df_mensual.columns:  
        serie_tiempo = df_mensual[codigo].dropna()
        
        if len(serie_tiempo) > 4:
            serie_tiempo_sin_outliers = detectar_y_reemplazar_outliers(serie_tiempo, method='linear')
            datos = serie_tiempo_sin_outliers.values
            data_scaled, scaler = preparar_datos(datos)
            X, Y = crear_conjunto_datos(data_scaled, look_back)  
    
            if len(X) > 0:
                model, history = entrenar_modelo(X, Y, look_back)
                predicciones_df = predecir_futuro_dataframe(model, data_scaled[-(look_back+1):], 4, scaler, look_back, codigo)
                
                ultima_fecha = df_mensual.index[-1]
                fecha_inicio_predicciones = ultima_fecha + pd.DateOffset(months=1)  
                fechas_futuras = pd.date_range(start=fecha_inicio_predicciones, periods=4, freq='MS')  
                predicciones_df['Fecha'] = fechas_futuras
                predicciones_totales = pd.concat([predicciones_totales, predicciones_df], ignore_index=True)

    if not predicciones_totales.empty:
        print("Predicciones para los códigos:")
        print(predicciones_totales)
        exportar_predicciones_y_series_de_tiempo_con_xlsxwriter(predicciones_totales, df_mensual, df_descripciones, archivo_destino)

    else:
        print("No se encontraron suficientes datos para generar predicciones.")

    return

#funcion_principal('VENTAS AÑO-2021.xlsx', 'VENTAS AÑO-2022.xlsx', 'VENTAS AÑO-2023.xlsx')

#funcion_principal('ventas año-2024.xlsx', 'VENTAS AÑO-2023.xlsx')

def abrir_pdf():
    nombre_pdf = "Pronóstico de baterías1.pdf"
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
    nombre_pdf = "Pronóstico de baterías2.pdf"
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
    def seleccionar_archivos():
        archivos_nuevos = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls *.csv")])
        archivos_actuales = set(lista_archivos.get(0, tk.END))  
        archivos_totales = archivos_actuales.union(set(archivos_nuevos))  
        lista_archivos.delete(0, tk.END)  
        for archivo in archivos_totales:
            lista_archivos.insert(tk.END, archivo)  
            print(archivo)

    def seleccionar_archivo_destino():
        archivo_destino = filedialog.asksaveasfilename(title="Especificar archivo de salida", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if archivo_destino:  
            entry_archivo_destino.delete(0, tk.END)
            entry_archivo_destino.insert(0, archivo_destino)

    codigos_adicionales = []        
    def agregar_nuevo_codigo():
        entrada = entry_nuevo_codigo.get().strip()  
        codigos_str = entrada.split(' ')  
        codigos_invalidos = []  
        
        for codigo_str in codigos_str:
            try:
                codigo = int(codigo_str)  
                if codigo not in codigos_adicionales:  
                    codigos_adicionales.append(codigo)
            except ValueError:
                if codigo_str.strip() != "":  
                    codigos_invalidos.append(codigo_str)

        if codigos_invalidos:
            messagebox.showerror("Error", f"Valores inválidos: {', '.join(codigos_invalidos)}. Ingrese solo números enteros separados por espacios.")
        else:
            messagebox.showinfo("Éxito", "Códigos agregados correctamente.")
        
        entry_nuevo_codigo.delete(0, tk.END)  

    def limpiar_todo():
        lista_archivos.delete(0, tk.END)
        entry_archivo_destino.delete(0, tk.END)
        entry_nuevo_codigo.delete(0, tk.END)
        codigos_adicionales.clear()
        messagebox.showinfo("Limpiar", "Se han limpiado los datos ingresados.")

    def procesar_archivo():
        archivos = lista_archivos.get(0, tk.END)
        archivo_destino = entry_archivo_destino.get()
        if archivos and archivo_destino:
            try:
                print(codigos_adicionales)
                funcion_principal(list(archivos), archivo_destino, codigos_adicionales)  
                messagebox.showinfo("Éxito", "Los archivos han sido procesados correctamente.")
                codigos_adicionales.clear()
            except Exception as e:
                messagebox.showerror("Error", f"Ocurrió un error al procesar los archivos: {e}")
        else:
            messagebox.showwarning("Advertencia", "Por favor, seleccione los archivos y especifique el archivo de salida.")

    root = tk.Tk()
    root.title("Pronóstico de baterías")
    root.geometry('680x295')  
    root.configure(bg='#2D2D2D')

    # Logo
    logo = cargar_imagen('Perno1.png')  
    logo_label = ttk.Label(root, image=logo, background='#2D2D2D')
    logo_label.grid(row=0, column=0, padx=5, pady=5, sticky='we')

    style = ttk.Style()
    style.configure('TLabel', font=('Arial', 11), background='#2D2D2D', foreground='white')
    style.configure('TButton', font=('Arial', 11))
    style.configure('TEntry', font=('Arial', 11))
    style.configure('TCheckbutton', font=('Arial', 11))

    fondo_columna = tk.Frame(root, bg='white', width=2)
    fondo_columna.grid(row=0, column=3, rowspan=6, sticky='nswe', pady=0)

    ttk.Label(root, text="Selección de Archivos", anchor='center', 
        font=('Arial', 12, 'bold'), background='light coral', foreground='white').grid(row=1, column=0, padx=0, pady=0, sticky='we', columnspan=3)

    ttk.Label(root, text="Guías", anchor='center', 
        font=('Arial', 12, 'bold'), background='light coral', foreground='white').grid(row=1, column=4, padx=0, pady=0, sticky='we', columnspan=5)

    ttk.Label(root, text="Ingreso de otras baterías al análisis", anchor='center', 
        font=('Arial', 12, 'bold'), background='light coral', foreground='white').grid(row=4, column=0, padx=0, pady=0, sticky='we', columnspan=3)


    ttk.Button(root, text="Seleccionar Archivos de Ventas", command=seleccionar_archivos).grid(row=2, column=2, padx=10, pady=5, sticky='we')

    lista_archivos = tk.Listbox(root, width=45, height=2)
    lista_archivos.grid(row=2, column=0, columnspan=2, padx=10, pady=10, sticky='we')

    entry_archivo_destino = ttk.Entry(root, width=45)
    entry_archivo_destino.grid(row=3,column=0, columnspan=2, padx=10, pady=5, sticky='we')
    ttk.Button(root, text="Especificar Destino", command=seleccionar_archivo_destino).grid(row=3, column=2, padx=10, pady=5, sticky='we')

    ttk.Button(root, text="Manual de Usuario", command=abrir_pdf).grid(row=2, column=5, padx=10, pady=5, sticky='we')

    ttk.Button(root, text="Interpretación", command=abrir_pdf1).grid(row=3, column=5, padx=10, pady=5, sticky='we')

    ttk.Button(root, text="Generar Reporte", command=procesar_archivo).grid(row=6, column=4, columnspan=3, padx=10, pady=30)

    ttk.Label(root, text="Agregar Códigos:").grid(row=5, column=0, padx=10, pady=5, sticky='w')
    entry_nuevo_codigo = ttk.Entry(root)
    entry_nuevo_codigo.grid(row=5, column=1, padx=10, pady=5, sticky='we')
    ttk.Button(root, text="Añadir", command=agregar_nuevo_codigo).grid(row=5, column=2, padx=10, pady=10, sticky='e')
    ttk.Button(root, text="Limpiar Todo", command=limpiar_todo).grid(row=6, column=0, padx=10, pady=5, sticky='we')

    root.mainloop()

if __name__ == "__main__":
    gui()

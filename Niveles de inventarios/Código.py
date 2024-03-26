#xpx
#Stocks

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import subprocess
import pandas as pd
import numpy as np
import csv
import sys  
import openpyxl

# Funciones Comunes
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
    motores = ['openpyxl', 'xlrd']  
    for engine in motores:
        try:
            df_temp = pd.read_excel(archivo, engine=engine, nrows=20)
            for i, row in df_temp.iterrows():
                celdas_no_vacias = sum(not pd.isna(cell) for cell in row[:6])
                if celdas_no_vacias >= 6:
                    return i, engine
            break
        except Exception as e:
            print(f"Error con motor {engine}: {e}")
            continue
    return 0, None

def leer_archivo(archivo):
    try:
        if archivo.lower().endswith('.csv'):
            delimitador, encoding = detectar_delimitador(archivo)
            if delimitador is None:
                raise ValueError("No se pudo determinar el delimitador.")
            skip_rows = encontrar_skip_rows_csv(archivo, delimitador, encoding)
            df = pd.read_csv(archivo, delimiter=delimitador, skiprows=skip_rows, encoding=encoding, low_memory=False)
        elif archivo.lower().endswith('.xlsx'):
            skip_rows, engine = encontrar_skip_rows_excel(archivo)
            if engine is None:
                raise ValueError("No se pudo leer el archivo Excel con los motores disponibles.")
            df = pd.read_excel(archivo, engine=engine, skiprows=skip_rows)
            if skip_rows > 0:
                df.columns = df.iloc[0]
                df = df[1:]
            mask = df.columns.astype(str).str.contains('^Unnamed')
            df = df.loc[:, ~mask]
        else:
            raise ValueError("Formato de archivo no soportado.")
    except Exception as e:
        print(f"Error al leer el archivo: {e}")
        raise
    return df

def mapear_meses(dfnuevo):
    meses_mapeo = {
        'ENERO': 1, 'FEBRE': 2, 'MARZO': 3, 'ABRIL': 4, 'MAYO': 5, 'JUNIO': 6,
        'JULIO': 7, 'AGOST': 8, 'SEPTI': 9, 'OCTUB': 10, 'NOVIE': 11, 'DICIE': 12
    }
    dfnuevo['MES'] = dfnuevo['MES'].str.strip().map(meses_mapeo)
    dfnuevo['FECHA'] = pd.to_datetime(dfnuevo.assign(DAY=dfnuevo['DIA'], MONTH=dfnuevo['MES'], YEAR=dfnuevo['AÑO']).loc[:, ['YEAR','MONTH','DAY']])
    dfnuevo.sort_values(by='FECHA', inplace=True)

def procesar_lead_time(archivo_lead_time, df_promedio):

    df_lead_time = leer_archivo(archivo_lead_time)
    df_lead_time['FAMILIA'] = df_lead_time['FAMILIA'].str.strip().str.upper()

    df_merge = pd.merge(df_promedio, df_lead_time, on='FAMILIA', how='left')

    df_merge['LT Normal'].fillna(0, inplace=True)
    df_merge['LT Demora'].fillna(0, inplace=True)
    return df_merge

def contar_meses_unicos2(df_filtrado):
    df_filtrado.dropna(subset=['DIA', 'MES', 'AÑO'], inplace=True)

    df_filtrado['DIA'] = df_filtrado['DIA'].astype(int)
    df_filtrado['MES'] = df_filtrado['MES'].astype(int)
    df_filtrado['AÑO'] = df_filtrado['AÑO'].astype(int)

    if 'FECHA' not in df_filtrado.columns:
        df_filtrado['FECHA'] = pd.to_datetime(df_filtrado[['AÑO', 'MES', 'DIA']])

    numero_meses = df_filtrado['FECHA'].dt.to_period('M').nunique()
    print("número meses: ",numero_meses)
    return numero_meses

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
    columnas_para_conservar = ['SUCURSAL', 'DIA', 'MES', 'AÑO', 'ESTADO', 'COND.PAGO', 'CODIGO', 'DESCRIPCION CODIGO', 'CANTIDAD', 'FAMILIA', 'LINEA']
    df_filtrado = df_filtrado[columnas_para_conservar]
    
    df_filtrado['CANTIDAD'] = pd.to_numeric(df_filtrado['CANTIDAD'], errors='coerce').fillna(0)
    
    return df_filtrado

# FUNCIONES PARA SUCURSALES
def calculo_de_stock_sucursal(df, z_valor, factor_r):
    # Z = 1.65   Nivel de servicio del 95%
    #el valor de 0.5 puede variar dependiendo de la incertidumbre
    #que esté dispuesta a asumir la empresa
    for idx in df.index:

        #Casa matriz
        promedio_ventas_matriz = df.at[idx, 'Prom. Vent.M']
        desviacion_std_matriz = df.at[idx, 'Desv M']
        ratio_matriz = desviacion_std_matriz / promedio_ventas_matriz if promedio_ventas_matriz > 0 else 0
        #
        factor_reduccion_matriz = max(factor_r, 1 - ratio_matriz)
        stock_seguridad_matriz = z_valor * np.sqrt((df.at[idx, 'LT Normal'] + df.at[idx, 'LT Demora']) / 2) * desviacion_std_matriz * factor_reduccion_matriz
        
        df.at[idx, 'St. Seg M'] = np.ceil(stock_seguridad_matriz)
        
        stock_minimo_matriz = promedio_ventas_matriz * (df.at[idx, 'LT Normal'])
        df.at[idx, 'St. Min M'] = np.ceil(stock_minimo_matriz)
        
        punto_pedido_matriz = stock_minimo_matriz + stock_seguridad_matriz
        df.at[idx, 'P. Pedido M'] = np.ceil(punto_pedido_matriz)


        #La Serena
        promedio_ventas_serena = df.at[idx, 'Prom. Vent.LS']
        desviacion_std_serena = df.at[idx, 'Desv LS']
        ratio_serena = desviacion_std_serena / promedio_ventas_serena if promedio_ventas_serena > 0 else 0
        #
        factor_reduccion_serena = max(factor_r, 1 - ratio_serena)
        stock_seguridad_serena = z_valor * np.sqrt((df.at[idx, 'LT Normal'] + df.at[idx, 'LT Demora']) / 2) * desviacion_std_serena * factor_reduccion_serena
        
        df.at[idx, 'St. Seg LS'] = np.ceil(stock_seguridad_serena)
        
        stock_minimo_serena = promedio_ventas_serena * (df.at[idx, 'LT Normal'])
        df.at[idx, 'St. Min LS'] = np.ceil(stock_minimo_serena)
        
        punto_pedido_serena = stock_minimo_serena + stock_seguridad_serena
        df.at[idx, 'P. Pedido LS'] = np.ceil(punto_pedido_serena)

        #Barrio Industrial
        promedio_ventas_barrio = df.at[idx, 'Prom. Vent.BI']
        desviacion_std_barrio= df.at[idx, 'Desv BI']
        ratio_barrio = desviacion_std_barrio / promedio_ventas_barrio if promedio_ventas_barrio > 0 else 0
        #
        factor_reduccion_barrio = max(factor_r, 1 - ratio_barrio)
        stock_seguridad_barrio = z_valor * np.sqrt((df.at[idx, 'LT Normal'] + df.at[idx, 'LT Demora']) / 2) * desviacion_std_barrio * factor_reduccion_barrio
        
        df.at[idx, 'St. Seg BI'] = np.ceil(stock_seguridad_barrio)
        
        stock_minimo_barrio = promedio_ventas_barrio * (df.at[idx, 'LT Normal'])
        df.at[idx, 'St. Min BI'] = np.ceil(stock_minimo_barrio)
        
        punto_pedido_barrio = stock_minimo_barrio + stock_seguridad_barrio
        df.at[idx, 'P. Pedido BI'] = np.ceil(punto_pedido_barrio)

    return df

def procesar_ventas_y_promedio_sucursal(archivo_ventas_promedio):
    df_pivot, df_ventas = procesar_ventas_sucursal(archivo_ventas_promedio)
    mapear_meses(df_ventas)
    numero_meses = contar_meses_unicos2(df_ventas)
    
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

def procesar_ventas_sucursal(archivo_ventas):
    df_ventas = leer_archivo(archivo_ventas)

    df_ventas['DESCRIPCION CODIGO'] = df_ventas['DESCRIPCION CODIGO'].str.strip().str.upper()
    df_ventas['LINEA'] = df_ventas['LINEA'].str.strip().str.upper()
    df_ventas['FAMILIA'] = df_ventas['FAMILIA'].str.strip().str.upper()
    df_ventas['COND.PAGO'] = df_ventas['COND.PAGO'].str.strip().str.upper()

    df_filtrado = df_ventas[(df_ventas['ESTADO'] != 'NULO') & (df_ventas['COND.PAGO'] != 'TRASLADOS') & (df_ventas['CODIGO'] != 99999) & (df_ventas['CODIGO'] != 901) & (df_ventas['SUCURSAL'] != 'TRANSITO CASA MATRIZ') & (df_ventas['SUCURSAL'] != 'TRANSITO BODEGA CENTRAL')]
    
    df_agrupado = df_filtrado.groupby(['CODIGO', 'SUCURSAL'])['CANTIDAD'].sum().reset_index()
    df_pivot = df_agrupado.pivot(index='CODIGO', columns='SUCURSAL', values='CANTIDAD').reset_index().fillna(0)
    df_pivot.rename(columns={'CASA MATRIZ': 'Vent. Matriz', 'STOCK SERENA': 'Vent. Serena', 'BARRIO INDUSTRIAL': 'Vent. BI'}, inplace=True)
    return df_pivot, df_ventas

def ajustar_formato_series_tiempo_con_desviacion(archivo_ventas_promedio):
    ventas_promedio = leer_archivo(archivo_ventas_promedio)
    df_ventas_promedio = eliminar_columnas(ventas_promedio)
    mapear_meses(df_ventas_promedio)
    
    df_ventas_promedio['Mes-Año'] = df_ventas_promedio['FECHA'].dt.strftime('%Y-%m')
    df_agrupado = df_ventas_promedio.groupby(['CODIGO', 'SUCURSAL', 'Mes-Año'])['CANTIDAD'].sum().reset_index()
    df_pivoteado = df_agrupado.pivot_table(index=['CODIGO', 'SUCURSAL'], columns='Mes-Año', values='CANTIDAD', fill_value=0)
    df_pivoteado['Desv. M.'] = df_pivoteado.std(axis=1, ddof=1)
    df_final = df_pivoteado.reset_index()
    df_pivot = df_pivoteado.pivot_table(index='CODIGO', columns='SUCURSAL', values='Desv. M.').reset_index()
    df_pivot.columns.name = None

    df_pivot.rename(columns={
        'BARRIO INDUSTRIAL': 'Desv BI',
        'BODEGA CENTRAL': 'Desv BC',
        'CASA MATRIZ': 'Desv M',
        'STOCK SERENA': 'Desv LS'
    }, inplace=True)

    df_pivot.fillna(0, inplace=True)

    return df_pivot

def procesar_stock_sucursal(archivo_stocks):
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

# FUNCIONES PARA EMPRESA
def calculo_de_stock_empresa(df, z_valor, factor_r):
    # Z = 1.65  # Nivel de confianza del 95%
    #el valor de 0.5 puede variar dependiendo de la incertidumbre
    #que esté dispuesta a asumir la empresa
    for idx in df.index:
        promedio_ventas = df.at[idx, 'Vent. Promedio']
        desviacion_std = df.at[idx, 'Desviación Estándar']
        ratio = desviacion_std / promedio_ventas if promedio_ventas > 0 else 0
        
        # Ajustar el factor de reducción en función del ratio
        factor_reduccion = max(factor_r, 1 - ratio)
        stock_seguridad_ajustado = z_valor * np.sqrt((df.at[idx, 'LT Normal'] + df.at[idx, 'LT Demora']) / 2) * desviacion_std * factor_reduccion
        
        df.at[idx, 'Stock Seg'] = np.ceil(stock_seguridad_ajustado)
        
        stock_minimo = promedio_ventas * (df.at[idx, 'LT Normal'])
        df.at[idx, 'Stock Min'] = np.ceil(stock_minimo)
        
        punto_pedido = stock_minimo + stock_seguridad_ajustado
        df.at[idx, 'Punto Pedido'] = np.ceil(punto_pedido)
    return df

def sumar_cantidades_por_mes_y_calcular_desviacion(df_filtrado):
    df_filtrado['FECHA'] = pd.to_datetime(df_filtrado['FECHA'])
    df_filtrado['Mes-Año'] = df_filtrado['FECHA'].dt.strftime('%Y-%m')
    df_mensual = df_filtrado.groupby(['Mes-Año', 'CODIGO'])['CANTIDAD'].sum().reset_index()
    df_desviacion = df_mensual.groupby('CODIGO')['CANTIDAD'].std(ddof=0).reset_index()
    df_desviacion.rename(columns={'CANTIDAD': 'Desviación Estándar'}, inplace=True)
    # Unión de la desviación estándar a la suma mensual de cantidades
    df_mensual_con_desviacion = pd.merge(df_mensual, df_desviacion, on='CODIGO', how='left')

    return df_mensual_con_desviacion

def sumar_cantidades_por_mes(df_filtrado):
    # Códigos como filas
    df_filtrado['Mes-Año'] = df_filtrado['FECHA'].dt.strftime('%Y-%m')
    df_mensual = df_filtrado.groupby(['Mes-Año', 'CODIGO'])['CANTIDAD'].sum().reset_index()
    return df_mensual

def procesar_stock_empresa(df_lead_time, archivo_stocks):
    df_stock_empresa = leer_archivo(archivo_stocks)
    df_stock_empresa.rename(columns={'Stock Emprea':'Stock Empresa', 'Codigo':'CODIGO'}, inplace = True)
    columnas_interes= ['CODIGO', 'Stock Empresa']
    df_stock_empresa = df_stock_empresa[columnas_interes]

    df_merge = pd.merge(df_lead_time, df_stock_empresa, on='CODIGO', how= 'left')
    df_merge['Stock Empresa'].fillna(0, inplace=True)
    return df_merge

# FUNCIONES PRINCIPALES
def procesar_sucursales(archivo_ventas_promedio, archivo_salida, archivo_lead_time, archivo_stocks, z_valor, factor_reduccion):
    df_filtrado = procesar_ventas_y_promedio_sucursal(archivo_ventas_promedio)
    df_stock_sucursal, c = procesar_stock_sucursal(archivo_stocks)
    df_stock_sucursal.rename(columns = {'Grupo':'FAMILIA', 'Codigo':'CODIGO', 'Descripcion':'DESCRIPCION'}, inplace=True)
    df_promedio = pd.merge(df_filtrado, df_stock_sucursal,on = 'CODIGO', how = 'left')

    df_promedio['St. Min M'] = np.nan
    df_promedio['St. Seg M'] = np.nan
    df_promedio['P. Pedido M'] = np.nan

    df_promedio['St. Min LS'] = np.nan
    df_promedio['St. Seg LS'] = np.nan
    df_promedio['P. Pedido LS'] = np.nan

    df_promedio['St. Min BI'] = np.nan
    df_promedio['St. Seg BI'] = np.nan
    df_promedio['P. Pedido BI'] = np.nan
    df_lead_time = procesar_lead_time(archivo_lead_time, df_promedio)
    columnas_interes = [
        'CODIGO', 'DESCRIPCION', 'FAMILIA', 
        'St. Matriz', 'St. Serena', 'St. BI', 'Stock Bodega', 
        'Prom. Vent.M', 'Prom. Vent.LS', 'Prom. Vent.BI', 'LT Normal',
        'LT Demora','St. Min M', 'St. Seg M', 
        'P. Pedido M','St. Min LS', 'St. Seg LS', 'P. Pedido LS', 
        'St. Min BI', 'St. Seg BI', 'P. Pedido BI'
    ]

    df_lead_time = df_lead_time[columnas_interes]

    df_desviacion = ajustar_formato_series_tiempo_con_desviacion(archivo_ventas_promedio)

    df_final = pd.merge(df_lead_time, df_desviacion,on='CODIGO', how='left')
    df_final.rename(columns={
        'CODIGO': 'Código',
        'DESCRIPCION': 'Descripción',
        'FAMILIA': 'Familia',
        'St. Serena': 'St. LS',
        'Stock Bodega': 'St. Bod'
    }, inplace=True)
    df_final2 = calculo_de_stock_sucursal(df_final, z_valor, factor_reduccion)

    return df_final2

def procesar_empresa(archivo_ventas_promedio, archivo_salida, archivo_lead_time, archivo_stocks, z_valor, factor_reduccion):
    df_ventas = leer_archivo(archivo_ventas_promedio)
    df_lineas = df_ventas[['CODIGO','LINEA']]
    df_lineas.rename(columns={'LINEA': 'Línea', 'CODIGO': 'Código'}, inplace=True)
    df_lineas = df_lineas.drop_duplicates(subset='Código', keep='first')


    print("Ventas leido")
    df_filtrado = eliminar_columnas(df_ventas)
    mapear_meses(df_filtrado)
    meses_unicos = contar_meses_unicos2(df_filtrado)
    df_mensual = sumar_cantidades_por_mes_y_calcular_desviacion(df_filtrado)
    print(df_mensual['Desviación Estándar'])
    columnas_interes = ['CODIGO', 'Desviación Estándar']

    df_estandar = df_mensual[columnas_interes]
    df_estandar = df_estandar.drop_duplicates('CODIGO')
    df_estandar['Desviación Estándar'] = df_estandar['Desviación Estándar'].round(1).astype(float)
    # Se divide la suma total de cantidades por el número de meses únicos para cada código
    df_promedio_mensual = df_mensual.groupby('CODIGO').agg({'CANTIDAD': lambda x: x.sum() / meses_unicos}).reset_index()

    df_promedio_mensual.rename(columns={'CANTIDAD':'Vent. Promedio'}, inplace= True)
    df_promedio_mensual['Vent. Promedio'] = df_promedio_mensual['Vent. Promedio'].round(0).astype(int)
    
    #Código para crear las columnas de Stock Min, Seg, etc.
    df_primera_aparicion= df_filtrado.drop_duplicates('CODIGO')
    df_promedio = pd.merge(df_primera_aparicion, df_promedio_mensual, on = 'CODIGO', how = 'left')
    df_promedio = pd.merge(df_promedio, df_estandar, on= 'CODIGO', how= 'left')
    
    print(df_promedio['Desviación Estándar'])

    df_promedio['Stock Min'] = np.nan
    df_promedio['Stock Seg'] = np.nan
    df_promedio['Punto Pedido'] = np.nan

    #Asociar cada Producto con su Lead time (tiempo que tarda una orden), con base a la familia
    columnas_interes = ['CODIGO', 'DESCRIPCION CODIGO', 'FAMILIA','Vent. Promedio',
    'Stock Min', 'Stock Seg', 'Punto Pedido', 'Desviación Estándar'
    ]
    df_promedio = df_promedio[columnas_interes]
    df_lead_time = procesar_lead_time(archivo_lead_time, df_promedio)
    df_stock_empresa = procesar_stock_empresa(df_lead_time, archivo_stocks)
    df_final = calculo_de_stock_empresa(df_stock_empresa, z_valor, factor_reduccion)
    columnas_interes = ['CODIGO', 'DESCRIPCION CODIGO', 'FAMILIA', 'Stock Empresa','Vent. Promedio', 'LT Normal', 'LT Demora',
    'Stock Min', 'Stock Seg', 'Punto Pedido', 'Desviación Estándar'
    ]
    df_final = df_final[columnas_interes]
    df_final['Stock Min'] = np.ceil(df_final['Stock Min']).astype(int)
    df_final['Stock Seg'] = np.ceil(df_final['Stock Seg']).astype(int)
    df_final['Punto Pedido'] = np.ceil(df_final['Punto Pedido']).astype(int)
    df_final.rename(columns = {'CODIGO':'Código', 'DESCRIPCION CODIGO': 'Descripción','FAMILIA':'Familia'}, inplace=True)
    return df_final, df_lineas

# FUNCIÓN PARA CREAR EXCEL
def crear_reporte(archivo_ventas_promedio, archivo_salida, archivo_lead_time, archivo_stocks, z_valor, factor_reduccion, buen_estado):
    print("*******************************")
    print('Inicio del proceso a nivel empresa')
    df_final1, df_lineas = procesar_empresa(archivo_ventas_promedio, archivo_salida, archivo_lead_time, archivo_stocks, z_valor, factor_reduccion)
    print("*******************************")
    print('Inicio del proceso a nivel sucursales') 
    print("*******************************")
    df_final2 = procesar_sucursales(archivo_ventas_promedio, archivo_salida, archivo_lead_time, archivo_stocks, z_valor, factor_reduccion)
    print('Inicio de la creación del reporte')


    # Nivel Empresa
    df_final1 = pd.merge(df_final1, df_lineas, on= 'Código', how='left')
    columnas_interes1 = ['Familia', 'Línea', 'Código', 'Descripción', 'Stock Empresa', 'Vent. Promedio',
     'LT Normal', 'LT Demora', 'Stock Min', 'Stock Seg', 'Punto Pedido', 'Desviación Estándar']
    df_final1 = df_final1[columnas_interes1]
    df_final1['Desviación Estándar'] = df_final1['Desviación Estándar'].round(1)


    # Nivel Sucursales
    df_final2 = pd.merge(df_final2, df_lineas, on= 'Código', how='left')
    columnas_interes2 = ['Familia', 'Línea', 'Código', 'Descripción', 'St. Matriz', 'St. LS', 'St. BI', 'St. Bod',
     'Prom. Vent.M', 'Prom. Vent.LS', 'Prom. Vent.BI', 'LT Normal', 'LT Demora', 'St. Min M', 'St. Seg M', 'P. Pedido M',
      'St. Min LS', 'St. Seg LS', 'P. Pedido LS', 'St. Min BI', 'St. Seg BI', 'P. Pedido BI', 'Desv BI', 'Desv BC', 'Desv M',
       'Desv LS']
    df_final2 = df_final2[columnas_interes2]
    columnas_desviacion = ['Desv BI', 'Desv BC', 'Desv M', 'Desv LS']
    for columna in columnas_desviacion:
        df_final2[columna] = df_final2[columna].round(1)
    with pd.ExcelWriter(archivo_salida, engine='xlsxwriter') as writer:
        df_final1.to_excel(writer, sheet_name='Stock Empresa', index=False, header= False, startrow=2)
        print(buen_estado)
        workbook = writer.book
        worksheet = writer.sheets['Stock Empresa']
        bold_format = workbook.add_format({'bold': True})
        for col_num, value in enumerate(df_final1.columns.values):
            worksheet.write(0, col_num, value, bold_format)

        worksheet.autofilter(1, 0, 1, len(df_final1.columns) - 1)

        formato_rojo = workbook.add_format({'bg_color': '#FFC7CE'})
        formato_amarillo = workbook.add_format({'bg_color': '#FFEB9C'})
        formato_naranjo = workbook.add_format({'bg_color': '#FFC000'}) 
        formato_celeste = workbook.add_format({'bg_color': '#4FC3F7'})
        formato_verde = workbook.add_format({'bg_color': '#A9D08E'})


        # Stock menor o igual al de seguridad
        worksheet.conditional_format('E3:E1048576', {'type': 'formula',
                                                    'criteria': '=AND(E3<=J3)',
                                                    'format': formato_rojo})
        # Stock menor o igual al min y mayor al seg                                    
        worksheet.conditional_format('E3:E1048576', {'type': 'formula',
                                                    'criteria': '=AND(E3<K3, E3>J3)',
                                                    'format': formato_amarillo})
        """# Stock mayor al min y menor al de pedido                                            
        worksheet.conditional_format('E3:E1048576', {'type': 'formula',
                                                    'criteria': '=AND(E3>I3, E3<K3)', 
                                                    'format': formato_amarillo}) 
        """                                            
        # Stock mayor o igual a PP y menor o igual a buen_estado veces PP
        worksheet.conditional_format('E3:E1048576', {'type': 'formula',
                                                    'criteria': f'=AND(E3>=K3, E3<=K3*{buen_estado})',
                                                    'format': formato_celeste})
        # Stock mayor a buen_estado veces PP
        worksheet.conditional_format('E3:E1048576', {'type': 'formula',
                                                    'criteria': f'=E3>K3*{buen_estado}',
                                                    'format': formato_verde})


        print("Información guardada en", archivo_salida)

        df_final2.to_excel(writer, sheet_name='Stock Sucursal', index=False, header=False, startrow=2)
 
        workbook1 = writer.book
        worksheet1 = writer.sheets['Stock Sucursal']
        bold_format = workbook1.add_format({'bold': True})

        for col_num, value in enumerate(df_final2.columns.values):
            worksheet1.write(0, col_num, value, bold_format)
        worksheet1.autofilter(1, 0, 1, len(df_final2.columns) - 1)

        formato_rojo = workbook1.add_format({'bg_color': '#FFC7CE'})
        formato_amarillo = workbook1.add_format({'bg_color': '#FFEB9C'})
        formato_naranjo = workbook1.add_format({'bg_color': '#FFC000'}) 
        formato_celeste = workbook1.add_format({'bg_color': '#4FC3F7'})
        formato_verde = workbook1.add_format({'bg_color': '#A9D08E'})

        for col in ['O', 'R', 'U']:
            worksheet1.conditional_format(f'{col}3:{col}1048576', {'type': 'cell', 'criteria': '!=', 'value': '""', 'format': formato_rojo})

        for col in ['N', 'Q', 'T']:
            worksheet1.conditional_format(f'{col}3:{col}1048576', {'type': 'cell', 'criteria': '!=', 'value': '""', 'format': formato_amarillo})

        for col in ['P', 'S', 'V']:
            worksheet1.conditional_format(f'{col}3:{col}1048576', {'type': 'cell', 'criteria': '!=', 'value': '""', 'format': formato_celeste})

        #Casa Matriz
        # Stock menor o igual al de seguridad
        worksheet1.conditional_format('E3:E1048576', {'type': 'formula',
                                                        'criteria': '=AND(E3<=O3)',
                                                        'format': formato_rojo})
        # Stock menor o igual al min y mayor al seg                                    
        worksheet1.conditional_format('E3:E1048576', {'type': 'formula',
                                                        'criteria': '=AND(E3<P3, E3>O3)',
                                                        'format': formato_amarillo})
        """# Stock mayor al min y menor al de pedido                                            
        worksheet1.conditional_format('E3:E1048576', {'type': 'formula',
                                                        'criteria': '=AND(E3>N3, E3<P3)', 
                                                        'format': formato_amarillo}) 
        """
        # Stock mayor o igual a PP y menor o igual a buen_estado veces PP
        worksheet1.conditional_format('E3:E1048576', {'type': 'formula',
                                                        'criteria': f'=AND(E3>=P3, E3<=P3*{buen_estado})',
                                                        'format': formato_celeste})
        # Stock mayor a buen_estado veces PP
        worksheet1.conditional_format('E3:E1048576', {'type': 'formula',
                                                        'criteria': f'=E3>P3*{buen_estado}',
                                                        'format': formato_verde})
        #La Serena
        # Stock menor o igual al de seguridad
        worksheet1.conditional_format('F3:F1048576', {'type': 'formula',
                                                        'criteria': '=AND(F3<=R3)',
                                                        'format': formato_rojo})
        # Stock menor o igual al min y mayor al seg                                    
        worksheet1.conditional_format('F3:F1048576', {'type': 'formula',
                                                        'criteria': '=AND(F3<S3, F3>R3)',
                                                        'format': formato_amarillo})
        """# Stock mayor al min y menor al de pedido                                            
        worksheet1.conditional_format('F3:F1048576', {'type': 'formula',
                                                        'criteria': '=AND(F3>Q3, F3<S3)', 
                                                        'format': formato_amarillo}) 
        """
        # Stock mayor o igual a PP y menor o igual a buen_estado veces PP
        worksheet1.conditional_format('F3:F1048576', {'type': 'formula',
                                                        'criteria': f'=AND(F3>=S3, F3<=S3*{buen_estado})',
                                                        'format': formato_celeste})
        # Stock mayor a buen_estado veces PP
        worksheet1.conditional_format('F3:F1048576', {'type': 'formula',
                                                        'criteria': f'=F3>S3*{buen_estado}',
                                                        'format': formato_verde})
        #Barrio Industrial
        # Stock menor o igual al de seguridad
        worksheet1.conditional_format('G3:G1048576', {'type': 'formula',
                                                        'criteria': '=AND(G3<=U3)',
                                                        'format': formato_rojo})
        # Stock menor o igual al min y mayor al seg                                    
        worksheet1.conditional_format('G3:G1048576', {'type': 'formula',
                                                        'criteria': '=AND(G3<V3, G3>U3)',
                                                        'format': formato_amarillo})
        """# Stock mayor al min y menor al de pedido                                            
        worksheet1.conditional_format('G3:G1048576', {'type': 'formula',
                                                        'criteria': '=AND(G3>T3, G3<V3)', 
                                                        'format': formato_amarillo}) 
        """
        # Stock mayor o igual a PP y menor o igual a buen_estado veces PP
        worksheet1.conditional_format('G3:G1048576', {'type': 'formula',
                                                        'criteria': f'=AND(G3>=V3, G3<=V3*{buen_estado})',
                                                        'format': formato_celeste})
        # Stock mayor a buen_estado veces PP
        worksheet1.conditional_format('G3:G1048576', {'type': 'formula',
                                                        'criteria': f'=G3>V3*{buen_estado}',
                                                        'format': formato_verde})

                                                                                                        

        print("Información guardada en", archivo_salida)

#archivo_ventas_promedio = 'VENTAS AÑO-2023.xlsx'
#archivo_lead_time = 'Lead Time.xlsx'
#archivo_salida = 'Combinado2.xlsx'
#archivo_stocks = 'stock_por_sucursal_T0.csv'
#crear_reporte(archivo_ventas_promedio, archivo_salida, archivo_lead_time, archivo_stocks)

def abrir_pdf():
    nombre_pdf = "Gestión de inventarios1.pdf"
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
    nombre_pdf = "Gestión de inventarios2.pdf"
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

    def seleccionar_archivo_ventas_promedio():
        archivo = filedialog.askopenfilename(title="Seleccionar archivo de ventas", filetypes=[("Excel files", "*.xlsx *.xls *.csv")])
        entry_archivo_ventas_promedio.delete(0, tk.END)
        entry_archivo_ventas_promedio.insert(0, archivo)

    def seleccionar_archivo_stock():
        archivo = filedialog.askopenfilename(title="Seleccionar archivo de stock", filetypes=[("Excel files", "*.xlsx *.xls *.csv")])
        entry_archivo_stocks.delete(0, tk.END)
        entry_archivo_stocks.insert(0, archivo)

    def seleccionar_archivo_lead_time():
        archivo = filedialog.askopenfilename(title="Seleccionar archivo lead time", filetypes=[("Excel files", "*.xlsx *.xls *.csv")])
        entry_archivo_lead_time.delete(0, tk.END)
        entry_archivo_lead_time.insert(0, archivo)

    def seleccionar_archivo_salida():
        archivo = filedialog.asksaveasfilename(title="Especificar archivo de salida", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        entry_archivo_salida.delete(0, tk.END)
        entry_archivo_salida.insert(0, archivo)

    def procesar_archivos():
        archivo_ventas_promedio = entry_archivo_ventas_promedio.get()
        archivo_stocks = entry_archivo_stocks.get()
        archivo_lead_time = entry_archivo_lead_time.get()
        archivo_salida = entry_archivo_salida.get()
        z_valor = float(entry_nivel_servicio.get()) if entry_nivel_servicio.get() else 1.65
        factor_reduccion = float(entry_factor_reduccion.get()) if entry_factor_reduccion.get() else 0.5
        buen_estado = float(entry_buen_estado.get()) if entry_buen_estado.get() else 1.05

       
        if archivo_ventas_promedio and archivo_stocks and archivo_lead_time and archivo_salida:
            try:
                crear_reporte(archivo_ventas_promedio, archivo_salida, archivo_lead_time, archivo_stocks, z_valor, factor_reduccion, buen_estado)
                messagebox.showinfo("Éxito", "Los archivos han sido procesados correctamente.")

            except Exception as e:
                messagebox.showerror("Error", f"Ocurrió un error al procesar los archivos: {e}")
        else:
            messagebox.showwarning("Advertencia", "Por favor, complete todos los campos.")




    root = tk.Tk()
    root.title("Gestión de Inventarios")
    root.geometry('580x380')  
    root.configure(bg='#2D2D2D')

    fondo_columna = tk.Frame(root, bg='white', width= 2)
    fondo_columna.grid(row=3, column=3, rowspan=9, sticky='nswe', pady=0)

    tk.Label(root, text="Selección de Archivos", anchor='center', 
        font=('Arial', 12, 'bold'), background='light coral', foreground='white').grid(row=3, column=0, padx=0, pady=0, sticky='we', columnspan=3)

    tk.Label(root, text="Guías", anchor='center', 
        font=('Arial', 12, 'bold'), background='light coral', foreground='white').grid(row=3, column=4, padx=0, pady=0, sticky='we', columnspan=5)

    tk.Label(root, text="Parámetros para el cálculo del Stock de Seguridad", anchor='center', 
        font=('Arial', 12, 'bold'), background='light coral', foreground='white').grid(row=8, column=0, padx=0, pady=0, sticky='we', columnspan=3)

    style = ttk.Style()
    style.configure('TLabel', font=('Arial', 11), background='#2D2D2D', foreground='white')
    style.configure('TButton', font=('Arial', 11))
    style.configure('TEntry', font=('Arial', 11))
    style.configure('TCheckbutton', font=('Arial', 11))  


    for i in range(3):
        root.grid_rowconfigure(i, minsize=10)  
        root.grid_columnconfigure(0, minsize=10)  

    logo = cargar_imagen('Perno1.png')
    logo_label = ttk.Label(root, image=logo, background='#2D2D2D')
    logo_label.grid(row=0, column=0, padx=5, pady=5, rowspan=3, sticky='we')

    ttk.Label(root, text="Archivo de Ventas:", background='#2D2D2D', foreground='white').grid(row=4, column=0, padx=5, pady=5, sticky='w')
    entry_archivo_ventas_promedio = ttk.Entry(root, width=25)
    entry_archivo_ventas_promedio.grid(row=4, column=1, padx=5, pady=5, sticky='we')
    ttk.Button(root, text="Seleccionar", command=seleccionar_archivo_ventas_promedio).grid(row=4, column=2, padx=5, pady=5)

    ttk.Label(root, text="Archivo de Stock:", background='#2D2D2D', foreground='white').grid(row=5, column=0, padx=5, pady=5, sticky='w')
    entry_archivo_stocks = ttk.Entry(root, width=25)
    entry_archivo_stocks.grid(row=5, column=1, padx=5, pady=5, sticky='we')
    ttk.Button(root, text="Seleccionar", command=seleccionar_archivo_stock).grid(row=5, column=2, padx=5, pady=5)

    ttk.Label(root, text="Archivo Lead Time:", background='#2D2D2D', foreground='white').grid(row=6, column=0, padx=5, pady=5, sticky='w')
    entry_archivo_lead_time = ttk.Entry(root, width=25)
    entry_archivo_lead_time.grid(row=6, column=1, padx=5, pady=5, sticky='we')
    ttk.Button(root, text="Seleccionar", command=seleccionar_archivo_lead_time).grid(row=6, column=2, padx=5, pady=5)

    ttk.Label(root, text="Archivo de Salida:", background='#2D2D2D', foreground='white').grid(row=7, column=0, padx=5, pady=5, sticky='w')
    entry_archivo_salida = ttk.Entry(root, width=25)
    entry_archivo_salida.grid(row=7, column=1, padx=5, pady=5, sticky='we')
    ttk.Button(root, text="Especificar", command=seleccionar_archivo_salida).grid(row=7, column=2, padx=5, pady=5)

    ttk.Button(root, text="Manual de Usuario", command=abrir_pdf).grid(row=4, column=4, padx=10, pady=5, sticky='we')

    ttk.Button(root, text="Interpretación", command=abrir_pdf1).grid(row=5, column=4, padx=10, pady=5, sticky='we')


    ttk.Button(root, text="Generar Reporte", command=procesar_archivos).grid(row=12, column=4, padx=5, pady=15)

    ttk.Label(root, text="Nivel de Servicio:", background='#2D2D2D', foreground='white').grid(row=9, column=0, padx=5, pady=5, sticky='w')
    entry_nivel_servicio = ttk.Entry(root, width=10)
    entry_nivel_servicio.insert(0, "1.65")  # Valor por defecto para Nivel de Servicio
    entry_nivel_servicio.grid(row=9, column=1, padx=5, pady=5, sticky='we')
    
    ttk.Label(root, text="Factor de Reducción:", background='#2D2D2D', foreground='white').grid(row=10, column=0, padx=5, pady=5, sticky='w')
    entry_factor_reduccion = ttk.Entry(root, width=10)
    entry_factor_reduccion.insert(0, "0.5")  # Valor por defecto para Factor de Reducción
    entry_factor_reduccion.grid(row=10, column=1, padx=5, pady=5, sticky='we')

    ttk.Label(root, text="Buen estado:", background='#2D2D2D', foreground='white').grid(row=11, column=0, padx=5, pady=5, sticky='w')
    entry_buen_estado = ttk.Entry(root, width=10)
    entry_buen_estado.insert(0, "1.05")  
    entry_buen_estado.grid(row=11, column=1, padx=5, pady=5, sticky='we')


    root.mainloop()

if __name__ == "__main__":
    gui()

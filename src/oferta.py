import os
import pandas as pd
import numpy as np
import openpyxl  # Importa openpyxl para trabajar con archivos Excel
from openpyxl.utils.dataframe import dataframe_to_rows  # Función para convertir un DataFrame de pandas en filas para Excel
from openpyxl.drawing.image import Image  # Importa Image para añadir imágenes a los archivos Excel
from datetime import datetime

TRM_ACTUAL_EURO=4300
TRM_ACTUAL_USD=3800

script_actual = os.path.realpath(__file__)  # Obtiene la ruta absoluta del script en ejecución
script_directory = os.path.dirname(script_actual)  # Obtiene el directorio donde se encuentra el script



def limpiar_dataframe(df):
    """
    Limpia el DataFrame proporcionado eliminando filas y columnas completamente vacías y ajustando los encabezados.
    
    :param df: DataFrame de pandas que será limpiado.
    :return: Tupla que contiene el DataFrame limpio y un diccionario con los totales de algunas columnas específicas.
    """
    df.dropna(axis=0, inplace=True, how='all')  # Elimina filas completamente vacías
    df.dropna(axis=1, inplace=True, how='all')  # Elimina columnas completamente vacías

    # Ajustar el DataFrame seleccionando desde la segunda fila hasta la penúltima
    nuevos_encabezados = df.iloc[0]  # Toma la primera fila para el encabezado
    df = df[1:]  # Elimina la primera fila, manteniendo temporalmente la última fila

    df.columns = nuevos_encabezados  # Establece la primera fila como el encabezado del DataFrame
    fila_TOTALES = df.iloc[-1]  # Ahora fila_TOTALES debería tener los índices correctos

    df = df[:-1]  # Ahora elimina la última fila después de capturar los totales

    # Resetear el índice para reflejar los cambios
    df.reset_index(drop=True, inplace=True)

    # Crear un diccionario con los totales específicos
    totales = {
        'SUBTOTAL': fila_TOTALES['SUBTOTAL'],
        'IVA': fila_TOTALES['IVA'],
        'TOTAL INCLUIDO IVA': fila_TOTALES['TOTAL INCLUIDO IVA']
    }

    return df, totales
 
def convertir_moneda(fila, moneda_objetivo):
    """
    Convierte el valor unitario de compra de una fila del DataFrame a la moneda objetivo.
    
    :param fila: Un DataFrame que representa la fila a convertir. Debe contener las columnas 'MONEDA' y 'VALOR UNITARIO COMPRA'.
    :param moneda_objetivo: La moneda a la que se desea convertir el valor ('COP', 'EUR', 'USD').
    :return: El valor convertido a la moneda objetivo.
    """
    tasas_cambio = {
        'COP': {'COP': 1, 'USD': TRM_ACTUAL_USD, 'EUR': TRM_ACTUAL_EURO},
        'USD': {'COP': 1/TRM_ACTUAL_USD, 'USD': 1, 'EUR': TRM_ACTUAL_EURO/TRM_ACTUAL_USD},
        'EUR': {'COP': 1/TRM_ACTUAL_EURO, 'USD': TRM_ACTUAL_USD/TRM_ACTUAL_EURO, 'EUR': 1},
    }
    
    moneda_origen = fila['MONEDA']
    precio_origen = fila['VALOR UNITARIO COMPRA']
    tasa_cambio = tasas_cambio[moneda_objetivo].get(moneda_origen, 1)  # Default a 1 si la moneda no se encuentra
    
    return precio_origen * tasa_cambio

def dataframe_pvp(directorio):
    """
    Busca y lee el archivo Excel PVP en el directorio especificado, limpiando el DataFrame resultante.
    
    :param directorio: El directorio donde se buscarán los archivos Excel que comienzan con 'PVP'.
    :return: Un DataFrame limpio y un diccionario con los totales de ciertas columnas.
    """
    archivo_pvp = next((archivo for archivo in os.listdir(directorio) if archivo.startswith('PVP')), None)
    ruta_pvp = os.path.join(directorio, archivo_pvp)
    pvp = pd.read_excel(ruta_pvp, header=None)  # Leer sin encabezado
    return limpiar_dataframe(pvp)

def dataframe_sp(directorio):
    """
    Busca y lee el archivo Excel SP en el directorio especificado.
    
    :param directorio: El directorio donde se buscarán los archivos Excel que comienzan con 'SP'.
    :return: Un DataFrame creado a partir del archivo Excel SP encontrado.
    """
    archivo_sp = next((archivo for archivo in os.listdir(directorio) if archivo.startswith('SP')), None)

    ruta_sp = os.path.join(directorio, archivo_sp)

    return pd.read_excel(ruta_sp, skiprows=18, sheet_name='SOLICITUD')

def generar_tabla_comparativa(SP, PVP, moneda_pvp):
    """
    Genera una tabla comparativa entre dos DataFrames (SP y PVP) según la moneda especificada.
    
    :param SP: DataFrame de SP que contiene información sobre productos.
    :param PVP: DataFrame de PVP que contiene información sobre precios de venta.
    :param moneda_pvp: Un entero que indica la moneda del PVP (1: COP, 2: EURO, 3: USD).
    :return: Un DataFrame que contiene la tabla comparativa generada.
    """
    df_comparacion = pd.DataFrame()
    # Comparaciones directas entre columnas de ambos DataFrames
    df_comparacion['NOMBRE'] = SP['DESCRIPCION'] == PVP['DESCRIPCION']
    df_comparacion['REFERENCIA'] = SP['REFERENCIA'] == PVP['REFERENCIA']
    df_comparacion['CANTIDAD'] = SP['CANTIDAD'] == PVP['CANTIDAD']

    # Uso de la función apply con lambda para convertir los precios según la moneda elegida
    if moneda_pvp == 'COP':
        df_comparacion['PRECIO SP EN COP'] = SP.apply(lambda fila: convertir_moneda(fila, 'COP'), axis=1)
    elif moneda_pvp == 'EUR':
        df_comparacion['PRECIO SP EN EUR'] = SP.apply(lambda fila: convertir_moneda(fila, 'EUR'), axis=1)
    elif moneda_pvp == 'USD':
        df_comparacion['PRECIO SP EN USD'] = SP.apply(lambda fila: convertir_moneda(fila, 'USD'), axis=1)

    df_comparacion['SUBTOTAL PVP'] = PVP['SUBTOTAL UNITARIO']
    columna_precio_sp = 'PRECIO SP EN ' + ['COP', 'EUR', 'USD'][moneda_pvp-1]
    df_comparacion[columna_precio_sp] = df_comparacion[columna_precio_sp].replace(0, np.nan)

    df_comparacion['PRECIO VENTA/PRECIO COMPRA'] = df_comparacion['SUBTOTAL PVP'] / df_comparacion[columna_precio_sp]
    df_comparacion['PRECIO VENTA/PRECIO COMPRA'].fillna(0, inplace=True)

    return df_comparacion

def llenar_oferta(directorio, dfpvp):
    """
    Llena y actualiza un archivo Excel de oferta con los datos proporcionados en el DataFrame dfpvp.
    
    :param directorio: Directorio donde se encuentra el archivo de oferta a actualizar.
    :param dfpvp: DataFrame que contiene los datos a insertar en el archivo de oferta.
    """

    archivo_of = next((archivo for archivo in os.listdir(directorio) if archivo.startswith('OFERTA')), None)
    ruta_of = os.path.join(directorio, archivo_of)
    wb_OF = openpyxl.load_workbook(ruta_of)  # Loads the Excel workbook.
    hoja_destino= wb_OF.worksheets[0]

    # Iterates through the rows of the DataFrame and inserts them into the Excel sheet.
    for r_idx, row in enumerate(dataframe_to_rows(dfpvp.loc[:,['DESCRIPCION','REFERENCIA','CANTIDAD','SUBTOTAL UNITARIO']], index=False, header=False), 21):
        for c_idx, value in enumerate(row, 3):
            hoja_destino.cell(row=r_idx, column=c_idx, value=value)  # Inserts each value into the corresponding cell.

    # Obtener la fecha actual en el formato día/mes/año
    fecha_actual = datetime.now().strftime("%d/%m/%Y")
    hoja_destino['I17'] = f"FECHA: {fecha_actual}"
    img = Image(os.path.join(script_directory, 'second.png'))
    hoja_destino.add_image(img, 'B1')

    wb_OF.save(ruta_of)

def main():
    # Carga los datos del SP (Stock Presente) y PVP (Precio de Venta al Público) desde archivos Excel ubicados en el mismo directorio que este script.
    sp = dataframe_sp(script_directory)

    pvp, totales = dataframe_pvp(script_directory)


    # Muestra información relevante de los DataFrames cargados para verificar su correcta lectura y limpieza.
    print("----------------------------------------------------------------------------------------------------------")
    print("+++++++++++++++++++++++++++++++++++++++++++++TABLA SP+++++++++++++++++++++++++++++++++++++++++++++++++++\n")
    print(sp.iloc[:, [5, 0, 8, 10]])
    print("----------------------------------------------------------------------------------------------------------")
    print("+++++++++++++++++++++++++++++++++++++++++++++TABLA PVP++++++++++++++++++++++++++++++++++++++++++++++++++\n")
    print(pvp.iloc[:, [1, 0, 3, 4]])
    print(totales)
    print("----------------------------------------------------------------------------------------------------------")
    # Solicita al usuario que indique la moneda en la que está expresado el PVP, con opciones predefinidas.
    moneda = solicitar_entrada_numerica('En qué moneda está el PVP? (1. COP, 2. EURO, 3. USD): ', [1, 2, 3])

    # Genera una tabla comparativa entre los datos de SP y PVP, ajustando los precios al tipo de moneda seleccionado.
    df_COMPARATIVO = generar_tabla_comparativa(sp, pvp, moneda)

    # Imprime la tabla comparativa para revisión del usuario.
    print("----------------------------------------------------------------------------------------------------------")
    print("++++++++++++++++++++++++++++++++++++++++++TABLA COMPARATIVA+++++++++++++++++++++++++++++++++++++++++++++\n")
    print(df_COMPARATIVO)
    print("----------------------------------------------------------------------------------------------------------")

    # Solicita confirmación al usuario sobre si está de acuerdo con los precios de venta al público (PVP) calculados.
    confirmacion = solicitar_entrada_numerica("Está de acuerdo con el PVP? Proceder a llenar oferta? (1. Si, 0. No): ", [0, 1])

    # Si el usuario confirma (1), procede a llenar el archivo de oferta con los datos de PVP.
    if confirmacion == 1:

        llenar_oferta(script_directory, pvp)
       
        # Limpieza final: Elimina archivos temporales o no necesarios, como la imagen del encabezado y el script actual.
        os.remove(os.path.join(script_directory, 'second.png'))  # Elimina la imagen de encabezado.
        os.remove(script_actual)  # Elimina el script actual para evitar ejecuciones o modificaciones accidentales.


if __name__ == "__main__":
    main()
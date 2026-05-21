import os
import zipfile
import shutil
import pandas as pd
from config import DOWNLOAD_DIR, COLUMNAS_REQUERIDAS, MAPEO_BANCOS

def preparar_directorio_descargas():
    """Crea la carpeta de descargas o la vacía si ya tenía archivos viejos."""
    if os.path.exists(DOWNLOAD_DIR):
        shutil.rmtree(DOWNLOAD_DIR)
    os.makedirs(DOWNLOAD_DIR)

def validar_columnas_excel(columnas_del_excel):
    """Verifica que el Excel subido tenga exactamente las columnas configuradas."""
    return all(columna in columnas_del_excel for columna in COLUMNAS_REQUERIDAS)

def limpiar_datos_excel(df: pd.DataFrame) -> pd.DataFrame:
    """Aplica las reglas de negocio para limpiar el Excel antes de procesar."""
    
    # 1. Eliminar filas que contengan "TOTAL CONSOLIDADO" en cualquier parte
    mask_consolidado = df.astype(str).apply(
        lambda x: x.str.contains('TOTAL CONSOLIDADO', case=False, na=False)
    ).any(axis=1)
    df_limpio = df[~mask_consolidado].copy()
    
    # 2. Convertir textos 'N/A' o celdas vacías en valores nulos reales (NaN)
    df_limpio.replace(['N/A', 'n/a', 'N/a', '', ' '], pd.NA, inplace=True)
    
    # 3. Eliminar filas que tengan TODAS sus columnas vacías
    df_limpio.dropna(how='all', inplace=True)
    
    return df_limpio

def homologar_banco(nombre_banco_excel: str) -> str:
    """Traduce el nombre del banco del Excel al nombre de Banxico."""
    if pd.isna(nombre_banco_excel):
        return ""
    
    nombre_limpio = str(nombre_banco_excel).strip()
    # Busca en el diccionario; si no lo encuentra, devuelve el nombre original
    return MAPEO_BANCOS.get(nombre_limpio, nombre_limpio)

def crear_archivo_zip(nombre_zip="CEPs_Descargados.zip"):
    """Toma todos los PDFs y el Word descargados y los comprime en un archivo ZIP."""
    ruta_zip = os.path.join(os.getcwd(), nombre_zip)
    
    with zipfile.ZipFile(ruta_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for raiz, _, archivos in os.walk(DOWNLOAD_DIR):
            for archivo in archivos:
                ruta_completa = os.path.join(raiz, archivo)
                zipf.write(ruta_completa, arcname=archivo)
                
    return ruta_zip
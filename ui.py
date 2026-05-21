import streamlit as st
import pandas as pd
import os

from config import COLUMNAS_REQUERIDAS, DOWNLOAD_DIR
# Agregamos limpiar_datos_excel a las importaciones
from utils import preparar_directorio_descargas, crear_archivo_zip, validar_columnas_excel, limpiar_datos_excel
from scraper import ValidadorBanxico

def renderizar_interfaz():
    """Dibuja toda la interfaz de usuario en Streamlit."""
    
    # Configuración básica
    st.set_page_config(page_title="Validador de CEP - Banxico", page_icon="🏦", layout="centered")

    st.title("🏦 Validador Masivo de CEP - Banxico")
    st.markdown("Sube tu archivo Excel con los datos de las transferencias para automatizar la descarga de comprobantes.")

    # Carga de archivos
    archivo_excel = st.file_uploader("Sube tu archivo Excel (.xlsx)", type=["xlsx"])

    if archivo_excel is not None:
        df_datos = pd.read_excel(archivo_excel, dtype=str)
        
        if not validar_columnas_excel(df_datos.columns):
            st.error("❌ El Excel no tiene el formato correcto. Faltan columnas o están mal escritas.")
            st.info(f"Las columnas obligatorias son exactamente: **{', '.join(COLUMNAS_REQUERIDAS)}**")
        else:
            # LIMPIAMOS LOS DATOS AQUÍ (Quita TOTAL CONSOLIDADO y acomoda N/A)
            df_datos = limpiar_datos_excel(df_datos)
            
            st.success("✅ Archivo cargado, validado y limpiado correctamente.")
            st.write(f"Total de registros a procesar: **{len(df_datos)}**")
            st.dataframe(df_datos.head())
            
            if st.button("🚀 Iniciar Validación y Descarga"):
                _ejecutar_proceso(df_datos)

def _ejecutar_proceso(df_datos):
    """Maneja la lógica cuando el usuario hace clic en el botón de inicio."""
    preparar_directorio_descargas()
    
    st.divider()
    st.markdown("### Estado del Proceso")
    barra_progreso = st.progress(0)
    texto_estado = st.empty()
    
    def actualizar_progreso(actual, total):
        barra_progreso.progress(actual / total)
        texto_estado.text(f"Procesando registro {actual} de {total}...")
    
    with st.spinner("Iniciando el motor de validación..."):
        validador = ValidadorBanxico()
        try:
            df_resultados = validador.procesar_lote(df_datos, funcion_progreso=actualizar_progreso)
        finally:
            validador.cerrar()
    
    st.success("¡Proceso terminado!")
    st.write("Resumen de operaciones:")
    st.dataframe(df_resultados)
    
    # Generar botón de descarga si hay archivos
    archivos_descargados = os.listdir(DOWNLOAD_DIR)
    if len(archivos_descargados) > 0:
        ruta_zip = crear_archivo_zip()
        with open(ruta_zip, "rb") as fp:
            st.download_button(
                label="📥 Descargar todos los CEPs en ZIP",
                data=fp,
                file_name="CEPs_Validados.zip",
                mime="application/zip"
            )
    else:
        st.warning("⚠️ El proceso finalizó, pero no se encontraron archivos PDF descargados (Esto es normal en la prueba simulada).")
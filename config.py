import os

# --- RUTAS Y DIRECTORIOS ---
BASE_DIR = os.getcwd()
DOWNLOAD_DIR = os.path.join(BASE_DIR, "descargas_cep")

# --- URLs ---
URL_BANXICO_CEP = "https://www.banxico.org.mx/cep/"

# --- COLUMNAS EXACTAS DE TU EXCEL ---
# Agregamos las columnas necesarias para el título del Word
COLUMNAS_REQUERIDAS = [
    "Número",
    "Archivo Origen",
    "Fecha de pago", 
    "Importe", 
    "Cuenta bancaria emisora", 
    "Clave de rastreo", 
    "Institución emisora", 
    "Institución receptora", 
    "Cuenta beneficiaria"
]

# --- DICCIONARIO DE HOMOLOGACIÓN DE BANCOS ---
# Llave (Izquierda): Como viene en tu Excel
# Valor (Derecha): Como aparece EXACTAMENTE en la lista de Banxico
MAPEO_BANCOS = {
    "BANBAJÍO": "BAJIO",
    "BBVA Bancomer": "BBVA MEXICO",
    "BBVA": "BBVA MEXICO",
    "Santander Serfin": "SANTANDER",
    "Banorte": "BANORTE",
    "Citibanamex": "BANAMEX"
    # Nota: Podrás ir agregando más bancos a esta lista conforme los vayas detectando.
}
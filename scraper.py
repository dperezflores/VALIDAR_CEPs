import time
import random
import os
import glob
import pandas as pd
import fitz
from docx import Document
from docx.shared import Inches

import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select

from config import DOWNLOAD_DIR, URL_BANXICO_CEP
from utils import homologar_banco

class ValidadorBanxico:
    def __init__(self):
        self.doc = Document()
        self.ruta_word = os.path.join(DOWNLOAD_DIR, "Capturas_CEP.docx")
        
        self.driver = self._configurar_navegador()
        # Aumentamos a 120 segundos la paciencia total para el CAPTCHA
        self.wait = WebDriverWait(self.driver, 120) 

    def _configurar_navegador(self):
        options = uc.ChromeOptions()
        prefs = {
            "download.default_directory": DOWNLOAD_DIR,
            "download.prompt_for_download": False,
            "plugins.always_open_pdf_externally": True
        }
        options.add_experimental_option("prefs", prefs)
        
        driver = uc.Chrome(options=options, version_main=148) 
        driver.maximize_window()
        return driver

    def _escribir_como_humano(self, elemento, texto: str):
        """Simula a una persona tecleando letra por letra con pausas irregulares."""
        for letra in texto:
            elemento.send_keys(letra)
            # Pausa aleatoria entre 0.05 y 0.25 segundos por cada tecla
            time.sleep(random.uniform(0.05, 0.25))

    def procesar_registro(self, row: pd.Series) -> str:
        if pd.isna(row['Clave de rastreo']) or str(row['Clave de rastreo']).strip().upper() == 'N/A':
            return "Omitido: Sin Clave de rastreo"
            
        try:
            self.driver.delete_all_cookies() 
            self.driver.get(URL_BANXICO_CEP)
            
            # Pausa inicial humana para "leer" la página
            time.sleep(random.uniform(3.0, 5.0)) 
            
            # --- 1. FECHA (Mantenemos inyección para evitar el candado del calendario) ---
            input_fecha = self.wait.until(EC.presence_of_element_located((By.ID, "input_fecha")))
            fecha_raw = str(row['Fecha de pago']).strip().lower()
            
            meses = {
                'ene': '01', 'feb': '02', 'mar': '03', 'abr': '04', 'may': '05', 'jun': '06',
                'jul': '07', 'ago': '08', 'sep': '09', 'oct': '10', 'nov': '11', 'dic': '12'
            }
            for letras, numero in meses.items():
                if letras in fecha_raw:
                    fecha_raw = fecha_raw.replace(letras, numero)
                    
            try:
                fecha_limpia = pd.to_datetime(fecha_raw).strftime('%d-%m-%Y')
            except:
                fecha_limpia = fecha_raw
            
            script_js = """
                var input = arguments[0];
                input.value = arguments[1];
                input.dispatchEvent(new Event('input', { bubbles: true }));
                input.dispatchEvent(new Event('change', { bubbles: true }));
            """
            self.driver.execute_script(script_js, input_fecha, fecha_limpia)
            time.sleep(random.uniform(0.5, 1.5))
            
            # --- 2. CRITERIO Y CLAVE DE RASTREO (Tecleo humano) ---
            Select(self.driver.find_element(By.ID, "input_tipoCriterio")).select_by_visible_text("Clave de rastreo")
            time.sleep(random.uniform(0.5, 1.0))
            
            input_criterio = self.driver.find_element(By.ID, "input_criterio")
            self._escribir_como_humano(input_criterio, str(row['Clave de rastreo']))
            time.sleep(random.uniform(0.5, 1.0))
            
            # --- 3. BANCOS ---
            banco_emisor = homologar_banco(row['Institución emisora'])
            banco_receptor = homologar_banco(row['Institución receptora'])
            
            Select(self.driver.find_element(By.ID, "input_emisor")).select_by_visible_text(banco_emisor)
            time.sleep(random.uniform(0.5, 1.0))
            Select(self.driver.find_element(By.ID, "input_receptor")).select_by_visible_text(banco_receptor)
            time.sleep(random.uniform(0.5, 1.0))
            
            # --- 4. CUENTA Y MONTO (Tecleo humano) ---
            input_cuenta = self.driver.find_element(By.ID, "input_cuenta")
            self._escribir_como_humano(input_cuenta, str(row['Cuenta beneficiaria']))
            time.sleep(random.uniform(0.5, 1.0))
            
            monto_limpio = str(row['Importe']).replace(',', '').replace('$', '').strip()
            input_monto = self.driver.find_element(By.ID, "input_monto")
            self._escribir_como_humano(input_monto, monto_limpio)
            time.sleep(random.uniform(1.0, 2.0)) # Pausa antes de hacer clic
            
            # --- REGISTRAR PDFs ANTES DE DAR CLIC ---
            pdfs_antes = set(glob.glob(os.path.join(DOWNLOAD_DIR, '*.pdf')))
            
            # --- 5. CLIC FORZADO ---
            try:
                btn_principal = self.driver.find_element(By.ID, "btn_Descargar")
            except:
                btn_principal = self.driver.find_element(By.ID, "btn_Consultar")
                
            self.driver.execute_script("arguments[0].click();", btn_principal)
            
            # --- 6. ESPERAR RESULTADO / RESOLVER CAPTCHA ---
            # Aquí es donde el robot te da hasta 120 segundos para que resuelvas el CAPTCHA si sale.
            try:
                btn_pdf_modal = self.wait.until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, ".boton.boton-descarga-pdf"))
                )
                time.sleep(random.uniform(1.0, 2.0))
                self.driver.execute_script("arguments[0].click();", btn_pdf_modal)
            except:
                return "Rechazado por Banxico / Timeout de CAPTCHA"
            
            # --- 7. CONVERTIR PDF A IMAGEN Y PEGAR EN WORD ---
            pdf_nuevo = None
            for _ in range(15):
                pdfs_despues = set(glob.glob(os.path.join(DOWNLOAD_DIR, '*.pdf')))
                archivos_nuevos = pdfs_despues - pdfs_antes
                if archivos_nuevos:
                    pdf_nuevo = archivos_nuevos.pop()
                    break
                time.sleep(1)
                
            if not pdf_nuevo:
                return "Error: Falla en la descarga del archivo"
            
            documento_pdf = fitz.open(pdf_nuevo)
            pagina = documento_pdf.load_page(0)
            matriz_zoom = fitz.Matrix(2.0, 2.0)
            mapa_bits = pagina.get_pixmap(matrix=matriz_zoom)
            
            ruta_imagen_temp = "temp_pdf_render.png"
            mapa_bits.save(ruta_imagen_temp)
            documento_pdf.close()
            
            titulo = f"{row.get('Número', '')} - {row.get('Archivo Origen', '')}"
            self.doc.add_heading(titulo, level=1)
            self.doc.add_picture(ruta_imagen_temp, width=Inches(6.0))
            
            if os.path.exists(ruta_imagen_temp):
                os.remove(ruta_imagen_temp)
                
            return "Éxito (PDF Descargado)"
            
        except Exception as e:
            tipo_error = type(e).__name__
            return f"Error técnico: {tipo_error}"

    def procesar_lote(self, df: pd.DataFrame, funcion_progreso=None) -> pd.DataFrame:
        resultados = []
        total = len(df)
        
        for index, row in df.iterrows():
            estado = self.procesar_registro(row)
            resultados.append({
                "Clave de rastreo": row.get('Clave de rastreo', 'N/A'), 
                "Estado": estado
            })
            
            if funcion_progreso:
                funcion_progreso(index + 1, total)
                
            # --- ENFRIAMIENTO PROFESIONAL (COOLDOWN) ---
            # Si no es el último registro, descansa entre 10 y 15 segundos antes de la siguiente consulta.
            if index < total - 1:
                tiempo_descanso = random.uniform(10.0, 15.0)
                time.sleep(tiempo_descanso)
                
        return pd.DataFrame(resultados)

    def cerrar(self) -> None:
        if len(self.doc.paragraphs) > 0:
            self.doc.save(self.ruta_word)
            
        if self.driver:
            self.driver.quit()
"""
DCIC - Interfaz Unificada v1.0
==================================
Interfaz gráfica para automatización de:
- Falabella
- Mercadolibre Flex
- Walmart
- Ripley
- Paris
- Paginas

Requiere: pip install customtkinter pdfplumber selenium webdriver-manager
"""

import sys
import os
import re
import time
import threading
import winsound
import csv
from datetime import datetime, timedelta
from queue import Queue

# Instalar dependencias si no existen
def install_deps():
    deps = ['customtkinter', 'pdfplumber', 'selenium', 'webdriver-manager', 'pytesseract', 'pdf2image', 'Pillow']
    for dep in deps:
        try:
            __import__(dep.replace('-', '_'))
        except ImportError:
            os.system(f'pip install {dep}')

install_deps()

import customtkinter as ctk
from tkinter import filedialog, messagebox
import pdfplumber

# OCR para PDFs que son imágenes
try:
    import pytesseract
    from pdf2image import convert_from_path
    OCR_AVAILABLE = True
    
    # Configurar Tesseract
    TESSERACT_PATHS = [
        r'C:\Program Files\Tesseract-OCR\tesseract.exe',
        r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
    ]
    for path in TESSERACT_PATHS:
        if os.path.exists(path):
            pytesseract.pytesseract.tesseract_cmd = path
            break
    
    # Configurar Poppler
    POPPLER_PATHS = [
        r'C:\poppler\poppler-24.07.0\Library\bin',
        r'C:\poppler\Library\bin',
        r'C:\poppler\bin',
    ]
    POPPLER_PATH = None
    for path in POPPLER_PATHS:
        if os.path.exists(path):
            POPPLER_PATH = path
            break
except:
    OCR_AVAILABLE = False
    POPPLER_PATH = None

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import StaleElementReferenceException
from webdriver_manager.chrome import ChromeDriverManager


# ============== CONFIGURACIÓN ==============

CANALES = {
    "Falabella": {
        "patron": r'^32\d{8}$',
        "patron_busqueda": r'\b(32\d{8})\b',
        "ubicacion": "ZDESP-FALA-01",
        "color": "#28a745",  # Verde
        "keywords": ["falabella", "fala", "32"]
    },
    "Mercadolibre": {
        "patron": r'^2000\d{12,14}$',
        "patron_busqueda": r'\b(2000\d{12,14})\b',  # Simplificado para detectar en manifiestos
        "ubicacion": "ZDESP-FLEXMELI-01",
        "color": "#FFE600",  # Amarillo
        "keywords": ["mercadolibre", "meli", "flex", "marketcenter", "mkc"]
    },
    "Walmart": {
        "patron": r'^\d{13}$',
        "patron_busqueda": r'\b(\d{13})\b',
        "ubicacion": "ZDESP-WALMAT-01",  # Nota: es WALMAT sin R
        "color": "#17a2b8",  # Celeste
        "keywords": ["walmart", "wmt"]
    },
    "Paris": {
        "patron": r'^307\d{7}$|^308\d{7}$',
        "patron_busqueda": r'\b(30[78]\d{7})\b',
        "ubicacion": "ZDESP-PARIS-01",
        "color": "#001f5b",  # Azul marino
        "keywords": ["paris", "cencosud", "mkc", "marketcenter"]
    },
    "Ripley": {
        "patron": r'^24\d{9}-[A-Z]?$',
        "patron_busqueda": r'\b(24\d{9})\b',
        "ubicacion": "ZDESP-RIPLEY-01",
        "color": "#dc3545",  # Rojo
        "keywords": ["ripley", "rpl"]
    },
    "Paginas": {
        # Formatos: Vincenzi.cl-1369, GlowUp.cl-1700, Miglu-1004, Acqui-1017
        "patron": r'^[A-Za-z]+\.cl-\d+$|^[A-Za-z]+-\d+$',
        "patron_busqueda": r'\b([A-Za-z]+\.cl-\d+|[A-Za-z]+-\d{3,4})\b',
        "ubicacion": "ZDESP-01-01",
        "color": "#9C27B0",  # Morado
        "keywords": ["starken", "paginas", "homeclaf", "vincenzi", "glowup", "miglu", "acqui"]
    }
}

OPERADORES = ["Rafa", "Alejo", "Nicolas", "Tomas", "Contanza"]
OT_AUDIT_CSV = "historial_ots.csv"
GIT_UPDATE_STATUS_FILE = "git_update_status.txt"
USER_PINS = {
    "Rafa": "9115",
    "Alejo": "2026",
    "Nicolas": "6020",
    "Tomas": "1234",
    "Contanza": "1234",
}
USER_COLORS = {
    "Rafa": "#111111",      # Negro
    "Alejo": "#ff69b4",     # Rosado
    "Nicolas": "#1e90ff",   # Azul
    "Tomas": "#ff3b30",     # Rojo
    "Contanza": "#ffffff",  # Blanco
}
OPERADOR_UBICACION_MAP = {
    "Nicolas": "ZDESP-01-01",
    "Alejo": "ZDESP-02-02",
    "Rafa": "ZDESP-03-03",
    "Tomas": "ZDESP-BULKYMELI-01",
    "Contanza": "ZDESP-PARIS-01",
}

WMS_URL = "https://checkweb-prd-checkwms.azurewebsites.net/"
MONITOR_URL = "https://checkweb-prd-checkwms.azurewebsites.net/DocumentoDespacho/monitorsalida"
WMS_USER = "18539597"
WMS_PASS = "185395"

WAIT_TIMEOUT = 60
DELAY_STEP = 0.4      # Entre pasos del wizard
DELAY_SEARCH = 0.5   # Tras escribir en búsqueda
DELAY_PAGE = 0.8     # Tras click Siguiente (la página carga)
MAX_RETRIES = 2
OT_CAPTURE_TIME_TOLERANCE_SEC = 5
OT_CAPTURE_MAX_DELAY_SEC = 180
OT_REQUIRE_UBICACION_MATCH = True

# ─── ChromeDriver precargado en background ───
_preloaded_driver = None   # webdriver.Chrome listo para usar
_preload_lock = threading.Lock()

def preload_driver():
    """Pre-descarga y cachea el binario de ChromeDriver en background.
    NO abre una ventana Chrome — solo prepara el ejecutable para que
    al presionar EJECUTAR el driver se cree instantáneamente."""
    try:
        ChromeDriverManager().install()  # Descarga/cachea el binario
    except:
        pass  # Si falla, setup_driver() lo descargará en ese momento


# ============== EXTRACCIÓN PDF ==============

def detect_canal_from_pdf(pdf_path):
    """Detecta automáticamente el canal basado en el contenido del PDF."""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            # Leer primera página
            text = ""
            text_lower = ""
            if pdf.pages:
                text = pdf.pages[0].extract_text() or ""
                text_lower = text.lower()
            
            # PASO 1: Buscar primero por PATRÓN DE REFERENCIAS (más específico)
            # Orden de búsqueda: del más específico al menos específico
            
            # Paginas: Texto.cl-XXXX o Texto-XXXX (ej: Vincenzi.cl-1369, Miglu-1004)
            if re.search(r'\b[A-Za-z]+\.cl-\d+\b', text) or re.search(r'\b[A-Za-z]+-\d{3,4}\b', text):
                # Verificar que no sea Ripley (que también tiene guión)
                if not re.search(r'\b24\d{9}\b', text):
                    return "Paginas"
            
            # Ripley: 24XXXXXXXXX (11 dígitos que empiezan con 24, con o sin guion-letra al final)
            if re.search(r'\b24\d{9}\b', text):
                return "Ripley"
            
            # Mercadolibre: 2000 + 12-14 dígitos
            if re.search(r'\b2000\d{12,14}\b', text):
                return "Mercadolibre"
            
            # Paris: 307 o 308 + 7 dígitos
            if re.search(r'\b30[78]\d{7}\b', text):
                return "Paris"
            
            # Falabella: 32 + 8 dígitos
            if re.search(r'\b32\d{8}\b', text):
                return "Falabella"
            
            # Walmart: exactamente 13 dígitos (menos específico)
            if re.search(r'\b\d{13}\b', text):
                return "Walmart"
            
            # PASO 2: Buscar por KEYWORDS específicos
            # Paginas / Starken
            if any(kw in text_lower for kw in ["starken", "homeclaf", "vincenzi", "glowup", "miglu", "acqui", "paginas"]):
                return "Paginas"
            
            # Ripley
            if "ripley" in text_lower:
                return "Ripley"
            
            # Falabella
            if "falabella" in text_lower or "fala" in text_lower:
                return "Falabella"
            
            # Mercadolibre
            if "mercadolibre" in text_lower or "meli" in text_lower or "flex" in text_lower:
                return "Mercadolibre"
            
            # Paris (cuidado: marketcenter puede ser Meli o Paris)
            if "paris" in text_lower:
                return "Paris"
            
            # Walmart
            if "walmart" in text_lower:
                return "Walmart"
            
            # Si tiene MARKETCENTER pero no detectó antes, verificar referencias
            if "marketcenter" in text_lower:
                # Verificar si tiene referencias de Paris (307/308)
                if re.search(r'\b30[78]\d{7}\b', text):
                    return "Paris"
                # Si no, asumir Mercadolibre
                return "Mercadolibre"
                
    except Exception as e:
        print(f"Error detectando canal: {e}")
    
    return None


def extract_with_ocr(pdf_path, patron_busqueda):
    """Extrae referencias usando OCR para PDFs que son imágenes."""
    refs = []
    
    if not OCR_AVAILABLE:
        return refs
    
    try:
        # Convertir PDF a imágenes
        if POPPLER_PATH:
            pages = convert_from_path(pdf_path, poppler_path=POPPLER_PATH, dpi=300)
        else:
            pages = convert_from_path(pdf_path, dpi=300)
        
        for page in pages:
            # OCR
            text = pytesseract.image_to_string(page, lang='eng')
            
            # Buscar referencias
            found = re.findall(patron_busqueda, text)
            refs.extend(found)
            
            # Limpiar errores OCR comunes
            text_cleaned = text.replace('O', '0').replace('o', '0').replace('l', '1').replace('I', '1')
            found_cleaned = re.findall(patron_busqueda, text_cleaned)
            refs.extend(found_cleaned)
    except Exception as e:
        print(f"Error OCR: {e}")
    
    # Eliminar duplicados
    return list(dict.fromkeys(refs))


def extract_references(pdf_paths, canal):
    """Extrae referencias de los PDFs según el canal."""
    config = CANALES[canal]
    patron = config["patron"]
    patron_busqueda = config["patron_busqueda"]
    
    all_refs = []
    
    for pdf_path in pdf_paths:
        refs_from_pdf = []
        
        # Primer intento: pdfplumber
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    # De tablas
                    tables = page.extract_tables()
                    for table in tables:
                        for row in table:
                            if row:
                                for cell in row:
                                    if cell:
                                        cell_str = str(cell).strip()
                                        if re.match(patron, cell_str):
                                            if cell_str not in refs_from_pdf:
                                                refs_from_pdf.append(cell_str)
                    
                    # De texto
                    text = page.extract_text() or ""
                    found = re.findall(patron_busqueda, text)
                    for ref in found:
                        if ref not in refs_from_pdf:
                            refs_from_pdf.append(ref)
        except Exception as e:
            print(f"Error pdfplumber {pdf_path}: {e}")
        
        # Si no encontró nada, intentar OCR
        if not refs_from_pdf and OCR_AVAILABLE:
            print(f"Intentando OCR para {os.path.basename(pdf_path)}...")
            refs_from_pdf = extract_with_ocr(pdf_path, patron_busqueda)
        
        # Agregar al total
        for ref in refs_from_pdf:
            if ref not in all_refs:
                all_refs.append(ref)
    
    return all_refs


# ============== AUTOMATIZACIÓN WMS ==============

class WMSAutomation:
    def __init__(self, canal, log_callback=None, operador="Sin definir"):
        self.canal = canal
        self.operador = operador
        self.config = dict(CANALES[canal])
        ubicacion_forzada = OPERADOR_UBICACION_MAP.get(self.operador)
        if ubicacion_forzada:
            self.config["ubicacion"] = ubicacion_forzada
        self.driver = None
        self.wait = None
        self.orders_selected = []
        self.orders_not_found = []
        self.skus_sin_stock = []  # SKUs con error de stock (filas rojas)
        self.ot_generada = None   # Número de OT generada
        self.log_callback = log_callback or print
        self.running = True
    
    def log(self, message):
        self.log_callback(message)
    
    def _create_chrome_driver(self):
        """Crea un nuevo ChromeDriver con las opciones estándar."""
        options = Options()
        options.add_argument("--start-maximized")
        options.add_argument("--disable-notifications")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_experimental_option(
            "excludeSwitches", ["enable-logging", "enable-automation"])
        options.add_experimental_option("useAutomationExtension", False)
        service = Service(ChromeDriverManager().install())
        drv = webdriver.Chrome(service=service, options=options)
        drv.execute_script(
            "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        return drv

    def setup_driver(self):
        global _preloaded_driver
        # El preload ahora solo cachea el binario, no abre Chrome
        # Por lo tanto, _preloaded_driver siempre será None — creamos directo
        # (Dejo la lógica por si en el futuro se reutiliza)
        preloaded = None
        with _preload_lock:
            preloaded = _preloaded_driver
            _preloaded_driver = None

        if preloaded is not None:
            # Verificar que el driver precargado sigue vivo (ping)
            try:
                _ = preloaded.current_url  # Lanza excepción si está muerto
                self.driver = preloaded
                self.log("  ⚡ ChromeDriver precargado utilizado")
            except Exception:
                self.log("  (Driver precargado invalido, creando nuevo...)")
                try:
                    preloaded.quit()
                except:
                    pass
                preloaded = None

        if preloaded is None:
            self.driver = self._create_chrome_driver()

        self.wait = WebDriverWait(self.driver, WAIT_TIMEOUT)
    
    def dismiss_alerts(self):
        """Cierra cualquier alert/confirm/prompt JS inesperado."""
        try:
            alert = self.driver.switch_to.alert
            alert.dismiss()  # Equivale a cancelar/cerrar el popup
        except:
            pass  # No había alert, todo bien

    def js_click(self, element):
        self.dismiss_alerts()  # Limpiar popups antes de hacer click
        self.driver.execute_script("arguments[0].click();", element)
        self.dismiss_alerts()  # Limpiar popups que el click pudo haber generado

    def is_session_alive(self):
        """Verifica que la sesión del WMS sigue activa. Si expiró, hace re-login."""
        try:
            body = self.driver.find_element(By.TAG_NAME, "body").text.lower()
            # El WMS redirige al login cuando expira la sesión
            if "ingresar" in body and "contraseña" in body:
                self.log("  ⚠️ Sesión expirada, reintentando login...")
                return self.login()
            return True
        except:
            return False
    
    def login(self):
        self.log(f"Navegando a {WMS_URL}")
        self.driver.get(WMS_URL)
        time.sleep(3)
        
        self.log("Iniciando sesión...")
        
        try:
            inputs = self.driver.find_elements(By.TAG_NAME, "input")
            visible_inputs = [i for i in inputs if i.is_displayed() and i.get_attribute("type") not in ["hidden", "submit", "button"]]
            
            if len(visible_inputs) >= 2:
                visible_inputs[0].click()
                visible_inputs[0].clear()
                visible_inputs[0].send_keys(WMS_USER)
                time.sleep(0.3)
                
                visible_inputs[1].click()
                visible_inputs[1].clear()
                visible_inputs[1].send_keys(WMS_PASS)
                time.sleep(0.3)
                
                buttons = self.driver.find_elements(By.TAG_NAME, "button")
                for btn in buttons:
                    if "ingresar" in btn.text.lower():
                        self.js_click(btn)
                        time.sleep(4)
                        self.log("Login OK")
                        return True
                
                visible_inputs[1].send_keys(Keys.ENTER)
                time.sleep(4)
                self.log("Login OK")
                return True
        except Exception as e:
            self.log(f"Error login: {e}")
            return False
        
        return False
    
    def wait_for_table_data(self):
        self.log("Esperando carga de datos...")
        
        max_wait = 60
        start_time = time.time()
        
        while time.time() - start_time < max_wait and self.running:
            try:
                body_text = self.driver.find_element(By.TAG_NAME, "body").text
                
                if "Cargando" in body_text:
                    time.sleep(1)
                    continue
                
                info_elements = self.driver.find_elements(By.CSS_SELECTOR, ".dataTables_info")
                if info_elements:
                    info_text = info_elements[0].text
                    if "0 to 0" in info_text or "0 of 0" in info_text or "0 a 0" in info_text:
                        time.sleep(1)
                        continue
                
                rows = self.driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
                data_rows = [r for r in rows if r.is_displayed() and r.text.strip() and "Cargando" not in r.text]
                
                if len(data_rows) > 0:
                    self.log(f"Tabla cargada ({len(data_rows)} filas)")
                    time.sleep(1)
                    return True
                
                time.sleep(1)
            except:
                time.sleep(1)
        
        return False
    
    def navigate_to_monitor(self):
        self.log("Abriendo Monitor de salida...")
        self.driver.get(MONITOR_URL)
        time.sleep(2)
        return self.wait_for_table_data()
    
    def find_search_box(self):
        selectors = ["input[type='search']", ".dataTables_filter input", "input[aria-controls]"]
        for selector in selectors:
            try:
                elements = self.driver.find_elements(By.CSS_SELECTOR, selector)
                for el in elements:
                    if el.is_displayed() and el.is_enabled():
                        return el
            except:
                continue
        return None
    
    def wait_for_search_results(self):
        time.sleep(0.5)
        max_wait = 10
        start_time = time.time()
        
        while time.time() - start_time < max_wait:
            try:
                processing = self.driver.find_elements(By.CSS_SELECTOR, ".dataTables_processing")
                if any(p.is_displayed() for p in processing):
                    time.sleep(0.3)
                    continue
                
                if "Cargando" in self.driver.find_element(By.TAG_NAME, "body").text:
                    time.sleep(0.3)
                    continue
                
                return True
            except:
                time.sleep(0.3)
        return True
    
    def clear_and_type_search(self, text):
        for attempt in range(3):
            try:
                search_box = self.find_search_box()
                if not search_box:
                    return False
                
                search_box.click()
                time.sleep(0.2)
                search_box.send_keys(Keys.CONTROL + "a")
                time.sleep(0.1)
                search_box.send_keys(Keys.DELETE)
                time.sleep(0.5)
                search_box.send_keys(text)
                
                self.wait_for_search_results()
                time.sleep(DELAY_SEARCH)
                return True
            except StaleElementReferenceException:
                time.sleep(0.5)
            except:
                time.sleep(0.5)
        return False
    
    def get_visible_rows(self):
        try:
            rows = self.driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
            return [r for r in rows if r.is_displayed() and r.text.strip() and "Cargando" not in r.text]
        except:
            return []
    
    def try_select_checkbox(self, row):
        try:
            cbs = row.find_elements(By.CSS_SELECTOR, "input[type='checkbox']")
            if cbs:
                if not cbs[0].is_selected():
                    try:
                        cbs[0].click()
                    except:
                        self.js_click(cbs[0])
                return True
        except:
            pass
        return False
    
    def search_and_select(self, reference):
        for attempt in range(MAX_RETRIES):
            if not self.running:
                return False
            
            try:
                if not self.clear_and_type_search(reference):
                    if attempt < MAX_RETRIES - 1:
                        time.sleep(1)
                        continue
                    return False
                
                visible_rows = self.get_visible_rows()
                
                if not visible_rows:
                    if attempt < MAX_RETRIES - 1:
                        time.sleep(1.5)
                        continue
                    return False
                
                for row in visible_rows:
                    try:
                        if reference in row.text:
                            if self.try_select_checkbox(row):
                                time.sleep(0.5)
                                return True
                    except:
                        continue
                
                if attempt < MAX_RETRIES - 1:
                    time.sleep(1.5)
            except:
                if attempt < MAX_RETRIES - 1:
                    time.sleep(1)
        
        return False
    
    def clear_search(self):
        try:
            search_box = self.find_search_box()
            if search_box:
                search_box.click()
                search_box.send_keys(Keys.CONTROL + "a")
                search_box.send_keys(Keys.DELETE)
                time.sleep(0.3)
                self.wait_for_search_results()
        except:
            pass
    
    def click_next(self):
        selectors = [
            "//button[contains(text(), 'Siguiente paso')]",
            "//button[contains(text(), 'Siguiente Paso')]",
        ]
        
        for selector in selectors:
            try:
                btn = self.driver.find_element(By.XPATH, selector)
                if btn.is_displayed():
                    self.js_click(btn)
                    return True
            except:
                continue
        
        buttons = self.driver.find_elements(By.TAG_NAME, "button")
        for btn in buttons:
            if "siguiente" in btn.text.lower() and "paginate" not in (btn.get_attribute("class") or "").lower():
                self.js_click(btn)
                return True
        return False
    
    def check_stock_error(self):
        """Verifica errores de stock y captura los SKUs afectados."""
        has_errors = False
        try:
            # Breve espera para que la tabla cargue
            time.sleep(0.3)
            
            # Obtener todas las filas de la tabla
            rows = self.driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
            
            for row in rows:
                try:
                    if not row.is_displayed():
                        continue
                    
                    # Obtener atributos de la fila
                    row_class = (row.get_attribute("class") or "").lower()
                    row_style = (row.get_attribute("style") or "").lower()
                    bg_color = row.value_of_css_property("background-color")
                    
                    # Múltiples formas de detectar error
                    is_error = False
                    
                    # Por clase CSS
                    if any(x in row_class for x in ["danger", "error", "warning", "red", "alert"]):
                        is_error = True
                    
                    # Por estilo inline
                    if any(x in row_style for x in ["red", "rojo", "#f", "rgb(255", "rgba(255"]):
                        is_error = True
                    
                    # Por color de fondo (RGB)
                    if bg_color:
                        # Detectar tonos de rojo/rosado
                        if "255" in bg_color and ("0," in bg_color or ", 0" in bg_color):
                            is_error = True
                        if "248" in bg_color or "252" in bg_color or "244" in bg_color:
                            is_error = True
                        # rgba(255, 0, 0) o similar
                        if bg_color.startswith("rgba(2") and ", 0," in bg_color:
                            is_error = True
                    
                    # Buscar si alguna celda tiene clase de error
                    cells = row.find_elements(By.TAG_NAME, "td")
                    for cell in cells:
                        cell_class = (cell.get_attribute("class") or "").lower()
                        if any(x in cell_class for x in ["danger", "error", "red"]):
                            is_error = True
                            break
                    
                    if is_error and len(cells) >= 2:
                        has_errors = True
                        codigo = cells[0].text.strip() if cells[0].text else ""
                        descripcion = cells[1].text.strip() if len(cells) > 1 and cells[1].text else ""
                        
                        if codigo or descripcion:
                            sku_info = f"{codigo}"
                            if descripcion:
                                sku_info += f" - {descripcion[:50]}"
                            if sku_info and sku_info not in self.skus_sin_stock:
                                self.skus_sin_stock.append(sku_info)
                                
                except Exception as e:
                    continue
            
            # Método alternativo: buscar directamente elementos con clases de error
            error_selectors = [
                "tr.danger",
                "tr.error", 
                "tr.table-danger",
                "tr[style*='red']",
                "tr[style*='255']",
                ".table-danger",
                "tbody tr.bg-danger"
            ]
            
            for selector in error_selectors:
                try:
                    error_rows = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    for row in error_rows:
                        if row.is_displayed():
                            has_errors = True
                            cells = row.find_elements(By.TAG_NAME, "td")
                            if len(cells) >= 2:
                                codigo = cells[0].text.strip()
                                descripcion = cells[1].text.strip() if len(cells) > 1 else ""
                                sku_info = f"{codigo}"
                                if descripcion:
                                    sku_info += f" - {descripcion[:50]}"
                                if sku_info and sku_info not in self.skus_sin_stock:
                                    self.skus_sin_stock.append(sku_info)
                except:
                    pass
            
            # Método adicional: buscar por texto que indique error de stock
            try:
                body_text = self.driver.find_element(By.TAG_NAME, "body").text.lower()
                if "sin stock" in body_text or "stock insuficiente" in body_text or "no disponible" in body_text:
                    has_errors = True
            except:
                pass
            
            return has_errors
            
        except Exception as e:
            self.log(f"  Error verificando stock: {e}")
            return False
    
    def is_picking_consolidado_checked(self):
        """Verifica si el checkbox 'Picking consolidado' está realmente marcado."""
        try:
            result = self.driver.execute_script("""
                const all = Array.from(document.querySelectorAll('input[type="checkbox"]'));
                for (const cb of all) {
                    const txt = (((cb.closest('label') && cb.closest('label').innerText) || '') + ' ' +
                                ((cb.parentElement && cb.parentElement.innerText) || '')).toLowerCase();
                    if (txt.includes('picking') && txt.includes('consolidado')) {
                        return !!cb.checked;
                    }
                }
                return null;
            """)
            if result is not None:
                return bool(result)
        except:
            pass

        # Fallback por Selenium
        try:
            checkboxes = self.driver.find_elements(By.CSS_SELECTOR, "input[type='checkbox']")
            for cb in checkboxes:
                try:
                    parent_text = (cb.find_element(By.XPATH, "./..").text or "").lower()
                    if "picking" in parent_text and "consolidado" in parent_text:
                        return cb.is_selected()
                except:
                    continue
        except:
            pass

        return False

    def mark_picking_consolidado(self):
        """Intenta marcar 'Picking consolidado' y confirma el estado final."""
        for intento in range(1, 4):
            try:
                if self.is_picking_consolidado_checked():
                    self.log("  Picking consolidado confirmado.")
                    return True

                checkboxes = self.driver.find_elements(By.CSS_SELECTOR, "input[type='checkbox']")
                marcado = False
                for cb in checkboxes:
                    try:
                        parent = cb.find_element(By.XPATH, "./..")
                        if "picking" in parent.text.lower() and "consolidado" in parent.text.lower():
                            self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", cb)
                            time.sleep(0.2)
                            if not cb.is_selected():
                                self.js_click(cb)
                            marcado = True
                            break
                    except:
                        continue

                if not marcado:
                    visible_cbs = [cb for cb in checkboxes if cb.is_displayed()]
                    if len(visible_cbs) >= 3:
                        self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", visible_cbs[2])
                        time.sleep(0.2)
                        if not visible_cbs[2].is_selected():
                            self.js_click(visible_cbs[2])

                time.sleep(0.5)
                if self.is_picking_consolidado_checked():
                    self.log("  Picking consolidado marcado OK.")
                    return True
            except:
                pass

            self.log(f"  (Reintento {intento}/3 marcando 'Picking consolidado')")
            time.sleep(0.4)

        self.log("  ❌ No se pudo confirmar 'Picking consolidado' marcado.")
        return False
    
    def click_crear_ot(self):
        """Hace clic en el botón 'Crear OT' del Paso 5 del wizard.
        Busca con múltiples estrategias y hace scroll para asegurar que esté visible."""
        # Selectores XPATH — el text puede variar por resolución / locale
        selectors = [
            "//button[contains(text(), 'Crear OT')]",
            "//button[contains(text(), 'CREAR OT')]",
            "//button[contains(text(), 'crear ot')]",
            "//button[contains(text(), 'Generar OT')]",
            "//button[contains(text(), 'GENERAR OT')]",
        ]

        for selector in selectors:
            try:
                btn = self.driver.find_element(By.XPATH, selector)
                # Scroll al botón para asegurar que esté en pantalla
                self.driver.execute_script(
                    "arguments[0].scrollIntoView({block: 'center', behavior: 'instant'});", btn)
                time.sleep(0.3)
                try:
                    btn.click()  # Clic nativo primero
                except:
                    self.js_click(btn)  # Fallback JS
                self.log("  Clic en 'Crear OT' OK")
                return True
            except:
                continue

        # Fallback: buscar entre todos los botones visibles
        try:
            for btn in self.driver.find_elements(By.TAG_NAME, "button"):
                txt = btn.text.strip().lower()
                if ("crear" in txt and "ot" in txt) or ("generar" in txt and "ot" in txt):
                    if btn.is_displayed():
                        self.driver.execute_script(
                            "arguments[0].scrollIntoView({block: 'center'});", btn)
                        time.sleep(0.3)
                        try:
                            btn.click()
                        except:
                            self.js_click(btn)
                        self.log(f"  Clic en botón '{btn.text.strip()}' OK (fallback)")
                        return True
        except:
            pass

        self.log("  ⚠️ Botón 'Crear OT' no encontrado")
        return False
    
    def confirm_modal(self):
        time.sleep(0.5)  # Reducido de 1.5
        
        for selector in ["//button[text()='Si']", "//button[text()='Sí']", "//button[contains(text(), 'Si')]"]:
            try:
                btn = self.driver.find_element(By.XPATH, selector)
                if btn.is_displayed():
                    time.sleep(0.2)  # Reducido de 0.5
                    self.js_click(btn)
                    return True
            except:
                continue
        
        try:
            for btn in self.driver.find_elements(By.TAG_NAME, "button"):
                if btn.text.strip().lower() in ["si", "sí"] and btn.is_displayed():
                    self.js_click(btn)
                    return True
        except:
            pass
        return False
    
    def process_batch(self, references):
        ubicacion = self.config["ubicacion"]
        
        # PASO 1
        self.log("\n[1/5] Seleccionando órdenes...")
        
        for i, ref in enumerate(references):
            if not self.running:
                break
            
            if self.search_and_select(ref):
                self.orders_selected.append(ref)
                self.log(f"  [{i+1}/{len(references)}] {ref} OK")
            else:
                self.orders_not_found.append(ref)
                self.log(f"  [{i+1}/{len(references)}] {ref} NO ENCONTRADA")
        
        self.clear_search()
        time.sleep(0.5)  # Reducido de 1
        
        if not self.orders_selected:
            self.log("No hay órdenes para procesar")
            return False
        
        self.log(f"  Seleccionadas: {len(self.orders_selected)}")
        
        self.log("  Siguiente paso...")
        self.click_next()
        time.sleep(DELAY_PAGE)

        # PASO 2 - Seleccionar ubicación
        self.log(f"[2/5] Ubicación: {ubicacion}")
        ubicacion_found = False

        # Esperar que la tabla de ubicaciones cargue (máx 10s)
        deadline_ubi = time.time() + 10
        while time.time() < deadline_ubi:
            try:
                rows_ubi = self.driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
                rows_vis = [r for r in rows_ubi if r.is_displayed() and r.text.strip()
                            and "Cargando" not in r.text]
                if len(rows_vis) > 0:
                    break
            except:
                pass
            time.sleep(0.4)

        # Intentar ampliar paginación de ubicaciones (puede haber muchas)
        try:
            from selenium.webdriver.support.ui import Select as _Sel
            for _val in ["-1", "100", "50"]:
                try:
                    _sel_el = self.driver.find_element(
                        By.CSS_SELECTOR,
                        "select[name*='DataTables'], .dataTables_length select, select[name$='_length']"
                    )
                    _Sel(_sel_el).select_by_value(_val)
                    time.sleep(0.8)
                    break
                except:
                    continue
        except:
            pass

        # MÉTODO 1: JavaScript — busca y clica el radio en todas las filas
        for _intento in range(3):
            try:
                script = f"""
                    var elements = document.querySelectorAll('table tbody tr');
                    for (var i = 0; i < elements.length; i++) {{
                        if (elements[i].textContent.includes('{ubicacion}')) {{
                            elements[i].scrollIntoView({{block: 'center'}});
                            var radio = elements[i].querySelector('input[type="radio"]');
                            if (radio) {{ radio.click(); return true; }}
                        }}
                    }}
                    return false;
                """
                result = self.driver.execute_script(script)
                if result:
                    self.log(f"  {ubicacion} OK")
                    ubicacion_found = True
                    break
            except:
                pass
            time.sleep(0.5)

        # MÉTODO 2: Fallback Python con scroll
        if not ubicacion_found:
            try:
                self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(0.4)
                rows = self.driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
                for row in rows:
                    try:
                        if ubicacion in row.text:
                            self.driver.execute_script(
                                "arguments[0].scrollIntoView({block: 'center'});", row)
                            radios = row.find_elements(By.CSS_SELECTOR, "input[type='radio']")
                            if radios:
                                self.js_click(radios[0])
                                self.log(f"  {ubicacion} OK (fallback)")
                                ubicacion_found = True
                                break
                    except:
                        continue
            except:
                pass

        if not ubicacion_found:
            self.log(f"  ⚠️ ADVERTENCIA: {ubicacion} no encontrada")
            self.log("  Intentando continuar de todos modos...")

        time.sleep(DELAY_STEP)
        self.click_next()
        time.sleep(DELAY_PAGE)
        
        # PASO 3
        self.log("[3/5] Stock...")
        
        # Debug: mostrar info de las primeras filas para diagnosticar
        try:
            rows = self.driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
            self.log(f"  Analizando {len(rows)} filas de stock...")
            
            # Mostrar info de las primeras 5 filas para debug
            for i, row in enumerate(rows[:5]):
                try:
                    if row.is_displayed():
                        row_class = row.get_attribute("class") or "(sin clase)"
                        bg = row.value_of_css_property("background-color")
                        cells = row.find_elements(By.TAG_NAME, "td")
                        first_cell = cells[0].text[:20] if cells else "?"
                        self.log(f"    Fila {i+1}: clase='{row_class}' bg='{bg}' texto='{first_cell}'")
                except:
                    pass
        except Exception as e:
            self.log(f"  Error debug: {e}")
        
        if self.check_stock_error():
            self.log(f"  ⚠️ ADVERTENCIA: {len(self.skus_sin_stock)} SKU(s) SIN STOCK")
            for sku in self.skus_sin_stock:
                self.log(f"    🔴 {sku}")
        else:
            self.log("  ✅ OK (Sin errores detectados)")
        
        time.sleep(DELAY_STEP)
        self.click_next()
        time.sleep(DELAY_PAGE)
        
        # PASO 4
        self.log("[4/5] Operario... OK")
        time.sleep(DELAY_STEP)
        self.click_next()
        time.sleep(DELAY_PAGE)
        
        # PASO 5
        self.log("[5/5] Creando OT...")
        if not self.mark_picking_consolidado():
            self.log("  ❌ Abortado: 'Picking consolidado' es obligatorio para crear OT.")
            return False
        time.sleep(DELAY_STEP)
        
        # Guardar la hora ANTES de crear la OT (para identificarla luego por timestamp)
        tiempo_antes_crear = datetime.now()
        self.log(f"  Hora de creación registrada: {tiempo_antes_crear.strftime('%Y-%m-%dT%H:%M:%S')}")
        
        self.click_crear_ot()
        time.sleep(1.5)
        self.confirm_modal()
        self.dismiss_alerts()  # Limpiar cualquier popup/alert del modal

        # ── Polling inteligente: esperar hasta 15s a que la OT quede guardada ──
        # En vez de un sleep fijo, verificamos cada 0.5s si la página
        # ya no muestra el wizard de creación (señal de que el servidor confirmó)
        self.log("  Esperando confirmación del servidor...")
        _dl_modal = time.time() + 15
        while time.time() < _dl_modal:
            try:
                self.dismiss_alerts()
                # Si el wizard desapareció (ya no hay botón Siguiente visible)
                # es porque el WMS guardó la OT y cambió de vista
                btns = self.driver.find_elements(By.XPATH, "//button[contains(text(),'Siguiente')]")
                btns_vis = [b for b in btns if b.is_displayed()]
                if not btns_vis:
                    self.log("  ✅ Confirmed: wizard cerrado")
                    break
            except:
                pass
            time.sleep(0.5)
        else:
            self.log("  (Timeout wizard — continuando de todas formas)")
        # ── NIVEL 1: Intentar capturar OT desde la página actual (más rápido y confiable) ──
        ot_number = self.capture_ot_from_current_page(tiempo_antes_crear)

        if ot_number:
            self.log(f"\n🎉 ¡OT CREADA EXITOSAMENTE!")
            self.log(f"📋 Número de OT: {ot_number}")
            self.ot_generada = ot_number
            return True

        # ── NIVEL 2: Fallback — buscar en lista de OTs ──
        self.log("  (OT no encontrada en página actual, buscando en lista...)")
        ot_number = self.capture_ot_number(tiempo_antes_crear)

        if ot_number:
            self.log(f"\n🎉 ¡OT CREADA EXITOSAMENTE!")
            self.log(f"📋 Número de OT: {ot_number}")
            self.ot_generada = ot_number
        else:
            self.log("\n¡OT CREADA EXITOSAMENTE!")

        return True

    def capture_ot_from_current_page(self, tiempo_antes_crear=None):
        """
        Intenta leer el número de OT (PCKM...) directamente desde la página actual.
        Si se proporciona tiempo_antes_crear, filtra los códigos cuya fecha
        visible en la página sea >= ese momento (evita capturar OTs antiguas).
        """
        time.sleep(1.5)

        def parse_ts(ts_str):
            return self._parse_wms_timestamp(ts_str)

        # ── Estrategia 1: URL contiene el código ──
        try:
            m = re.search(r'(PCKM\d{6,15})', self.driver.current_url, re.IGNORECASE)
            if m:
                ot_code = m.group(1).upper()
                self.log(f"  ⚡ OT capturada desde URL: {ot_code}")
                return ot_code
        except:
            pass

        # ── Estrategia 2: Filas de tabla — buscar PCKM + timestamp y comparar fecha ──
        try:
            rows = self.driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
            candidatas = []
            for row in rows:
                try:
                    if not row.is_displayed():
                        continue
                    row_text = row.text
                    m_ot = re.search(r'(PCKM\d{6,15})', row_text, re.IGNORECASE)
                    if not m_ot:
                        continue
                    ot_code = m_ot.group(1).upper()
                    row_upper = row_text.upper()
                    creada_ok = "CREADA" in row_upper
                    ubicacion_ok = self.config["ubicacion"] in row_text

                    # Buscar timestamp en el texto de la fila (ej: 2026-02-27T10:14:08.17)
                    m_ts = re.search(
                        r'(\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}:\d{2}(?:\.\d+)?)', row_text)
                    fecha_dt = parse_ts(m_ts.group(1)) if m_ts else None
                    hora_str = m_ts.group(1) if m_ts else "sin fecha"

                    # Si tenemos tiempo_antes_crear, solo aceptar OTs creadas después
                    if tiempo_antes_crear and fecha_dt:
                        if fecha_dt < tiempo_antes_crear:
                            continue  # OT demasiado antigua, ignorar

                    num_match = re.search(r'PCKM0*(\d+)', ot_code)
                    ot_num = int(num_match.group(1)) if num_match else 0
                    candidatas.append({'codigo': ot_code, 'numero': ot_num,
                                       'hora_str': hora_str, 'fecha_dt': fecha_dt,
                                       'creada_ok': creada_ok, 'ubicacion_ok': ubicacion_ok})
                    self.log(f"    ⚡ Candidata (tabla actual): {ot_code} | {hora_str}")
                except:
                    continue

            if candidatas:
                elegida = self._pick_ot_candidate(
                    candidatas, tiempo_antes_crear, self.config["ubicacion"]
                )
                if elegida is None:
                    self.log("  (No hay candidata confiable por tiempo en tabla actual)")
                    return None
                self.log(f"  ⚡ OT capturada desde tabla actual: {elegida['codigo']} | {elegida['hora_str']}")
                return elegida['codigo']
        except:
            pass

        # ── Estrategia 3: Body completo — buscar PCKM + timestamp más reciente ──
        try:
            page_text = self.driver.find_element(By.TAG_NAME, "body").text
            # Encontrar todos los pares (PCKM, timestamp) en el texto
            pckm_positions = [(m.group(1).upper(), m.start())
                              for m in re.finditer(r'(PCKM\d{6,15})', page_text, re.IGNORECASE)]
            ts_positions = [(m.group(1), m.start())
                            for m in re.finditer(
                                r'(\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}:\d{2}(?:\.\d+)?)',
                                page_text)]

            candidatas_body = []
            for ot_code, ot_pos in pckm_positions:
                # Buscar el timestamp más cercano al PCKM (dentro de 300 caracteres)
                fecha_dt = None
                hora_str = "sin fecha"
                for ts_str, ts_pos in ts_positions:
                    if abs(ts_pos - ot_pos) < 300:
                        fecha_dt = parse_ts(ts_str)
                        hora_str = ts_str
                        break

                if tiempo_antes_crear and fecha_dt and fecha_dt < tiempo_antes_crear:
                    continue  # OT antigua, ignorar

                num_match = re.search(r'PCKM0*(\d+)', ot_code)
                ot_num = int(num_match.group(1)) if num_match else 0
                candidatas_body.append({'codigo': ot_code, 'numero': ot_num,
                                        'hora_str': hora_str, 'fecha_dt': fecha_dt,
                                        'creada_ok': True, 'ubicacion_ok': False})

            if candidatas_body:
                elegida = self._pick_ot_candidate(
                    candidatas_body, tiempo_antes_crear, self.config["ubicacion"]
                )
                if elegida is None:
                    self.log("  (No hay candidata confiable por tiempo en body)")
                    return None
                self.log(f"  ⚡ OT capturada desde body: {elegida['codigo']} | {elegida['hora_str']}")
                return elegida['codigo']
        except:
            pass

        # ── Estrategia 4: Notificaciones / toasts / modales ──
        try:
            for sel in [".alert", ".toast", ".swal2-content", ".swal2-html-container",
                        "[class*='success']", ".modal-body", ".modal-content"]:
                try:
                    for el in self.driver.find_elements(By.CSS_SELECTOR, sel):
                        if not el.is_displayed():
                            continue
                        m = re.search(r'(PCKM\d{6,15})', el.text, re.IGNORECASE)
                        if m:
                            ot_code = m.group(1).upper()
                            self.log(f"  ⚡ OT capturada desde notificación ({sel}): {ot_code}")
                            return ot_code
                except:
                    continue
        except:
            pass

        return None  # No encontrado — se usará el fallback capture_ot_number()

    def _parse_wms_timestamp(self, ts_str):
        for fmt in ("%Y-%m-%dT%H:%M:%S.%f", "%Y-%m-%dT%H:%M:%S",
                    "%Y-%m-%d %H:%M:%S.%f", "%Y-%m-%d %H:%M:%S"):
            try:
                return datetime.strptime(ts_str, fmt)
            except ValueError:
                continue
        return None

    def _pick_ot_candidate(self, candidates, tiempo_antes_crear=None, ubicacion=None):
        """Elige candidata más confiable: cercana al momento de creación y priorizando ubicación/estado."""
        if not candidates:
            return None

        if not tiempo_antes_crear:
            return sorted(candidates, key=lambda x: x.get('numero', 0), reverse=True)[0]

        tolerance = timedelta(seconds=OT_CAPTURE_TIME_TOLERANCE_SEC)
        max_delay = timedelta(seconds=OT_CAPTURE_MAX_DELAY_SEC)
        recientes = []

        for c in candidates:
            fecha_dt = c.get('fecha_dt')
            if fecha_dt is None:
                continue
            delta = fecha_dt - tiempo_antes_crear
            if delta < -tolerance or delta > max_delay:
                continue
            c_copy = dict(c)
            c_copy['_delta_seconds'] = delta.total_seconds()
            recientes.append(c_copy)

        if not recientes:
            return None

        if ubicacion:
            por_ubicacion = [c for c in recientes if c.get('ubicacion_ok')]
            if OT_REQUIRE_UBICACION_MATCH:
                if not por_ubicacion:
                    return None
                recientes = por_ubicacion
            elif por_ubicacion:
                recientes = por_ubicacion

        por_estado = [c for c in recientes if c.get('creada_ok', True)]
        if por_estado:
            recientes = por_estado

        recientes.sort(
            key=lambda x: (
                0 if x['_delta_seconds'] >= 0 else 1,
                abs(x['_delta_seconds']),
                -x.get('numero', 0),
            )
        )
        return recientes[0]
    
    def capture_ot_number(self, tiempo_antes_crear=None):
        """
        Captura el número de OT navegando al listado de Órdenes de Trabajo.
        Si se proporciona 'tiempo_antes_crear', busca la OT cuya fecha de creación
        sea >= ese momento (la OT que acabamos de crear).
        """
        ot_number = None
        ubicacion = self.config["ubicacion"]

        try:
            self.log(f"  Buscando OT para ubicación: '{ubicacion}'")

            # Navegar al listado de Órdenes de Trabajo
            ot_url = "https://checkweb-prd-checkwms.azurewebsites.net/OrdenTrabajo/index"
            en_pagina_ot = False
            for intento_nav in range(1, 4):
                try:
                    self.dismiss_alerts()
                    self.driver.get(ot_url)
                except:
                    pass

                deadline_nav = time.time() + 8
                while time.time() < deadline_nav:
                    try:
                        cur = (self.driver.current_url or "").lower()
                        body_txt = (self.driver.find_element(By.TAG_NAME, "body").text or "").lower()
                        if ("/ordentrabajo/index" in cur) or ("listado de órdenes de trabajo" in body_txt) or ("listado de ordenes de trabajo" in body_txt):
                            en_pagina_ot = True
                            break
                    except:
                        pass
                    time.sleep(0.4)

                if en_pagina_ot:
                    break

                self.log(f"  (No se pudo abrir Ordenes de Trabajo en intento {intento_nav}/3, reintentando...)")
                try:
                    self.driver.execute_script("window.location.href = arguments[0];", ot_url)
                except:
                    pass
                time.sleep(0.8)

            if not en_pagina_ot:
                self.log("  ⚠️ No se logró navegar a Ordenes de Trabajo. Revisa sesión/conectividad.")
                return None
            self.log("  Navegación a Ordenes de Trabajo OK.")

            # ── Esperar que la tabla cargue con filas reales (máx 20s) ──
            deadline = time.time() + 20
            while time.time() < deadline:
                try:
                    cur = (self.driver.current_url or "").lower()
                    if "/ordentrabajo/index" not in cur:
                        time.sleep(0.3)
                        continue
                    rows = self.driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
                    visible = [r for r in rows if r.is_displayed() and r.text.strip()
                               and "Cargando" not in r.text
                               and "Sin información" not in r.text]
                    if len(visible) > 0:
                        break
                except:
                    pass
                time.sleep(0.5)
            else:
                self.log("  ⚠️ La tabla de OTs tardó demasiado en cargar.")

            self.log("  Tabla de OTs cargada.")

            # ── PASO 1: Ampliar paginación a 100 registros para ver más OTs ──
            try:
                from selenium.webdriver.support.ui import Select as SeleniumSelect
                for sel_val in ["100", "50", "25"]:
                    try:
                        sel_el = self.driver.find_element(
                            By.CSS_SELECTOR,
                            "select[name*='DataTables'], .dataTables_length select, select[name$='_length']"
                        )
                        SeleniumSelect(sel_el).select_by_value(sel_val)
                        time.sleep(1.5)
                        self.log(f"  Paginación ampliada a {sel_val} registros.")
                        break
                    except:
                        continue
            except:
                pass

            # ── PASO 2: Filtrar la tabla por FECHA+HORA de creación de la OT ──
            # En vez de filtrar por ubicación (que puede no coincdir), usamos la
            # fecha y hora exacta registrada antes de crear la OT.
            # El WMS muestra la fecha debajo de "Estado Actual" en formato ISO.
            search_input = None

            dt_selectors = [
                "input[type='search']",
                ".dataTables_filter input",
                "input[aria-label*='Search']",
                "input[aria-label*='Búsqueda']",
                "input[aria-label*='busqueda']",
                ".dataTables_filter input[type='text']",
            ]
            for selector in dt_selectors:
                try:
                    inputs = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    for inp in inputs:
                        if inp.is_displayed() and inp.is_enabled():
                            search_input = inp
                            break
                    if search_input:
                        break
                except:
                    continue

            def _find_search_input():
                for _selector in dt_selectors:
                    try:
                        _inputs = self.driver.find_elements(By.CSS_SELECTOR, _selector)
                        for _inp in _inputs:
                            if _inp.is_displayed() and _inp.is_enabled():
                                return _inp
                    except:
                        continue
                return None

            def _wait_table_rows(timeout_sec=10):
                _deadline = time.time() + timeout_sec
                while time.time() < _deadline:
                    try:
                        _rows = self.driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
                        _visible = [r for r in _rows if r.is_displayed() and r.text.strip() and "Cargando" not in r.text]
                        if len(_visible) > 0:
                            return True
                    except:
                        pass
                    time.sleep(0.4)
                return False

            def _apply_filter_and_count(_search_input, _filtro_texto, timeout_sec=8):
                try:
                    _search_input.click()
                    time.sleep(0.2)
                    _search_input.send_keys(Keys.CONTROL + "a")
                    _search_input.send_keys(Keys.DELETE)
                    time.sleep(0.2)
                    _search_input.send_keys(_filtro_texto)
                    time.sleep(1.0)
                except Exception:
                    return 0

                _deadline2 = time.time() + timeout_sec
                _last_count = 0
                while time.time() < _deadline2:
                    try:
                        _filas = self.driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
                        _filas_ok = [r for r in _filas if r.is_displayed() and _filtro_texto in r.text]
                        _last_count = len(_filas_ok)
                        if _last_count > 0:
                            return _last_count
                    except:
                        pass
                    time.sleep(0.5)
                return _last_count

            if search_input:
                try:
                    filtro_exitoso = False

                    if tiempo_antes_crear:
                        filtros_fecha = []
                        for dt in [tiempo_antes_crear,
                                   tiempo_antes_crear + timedelta(minutes=1),
                                   tiempo_antes_crear - timedelta(minutes=1)]:
                            ftxt = dt.strftime("%Y-%m-%dT%H:%M")
                            if ftxt not in filtros_fecha:
                                filtros_fecha.append(ftxt)

                        for intento in range(3):
                            if intento > 0:
                                self.log(f"  (Fecha sin resultados aún, reintento {intento+1}/3 con refresh)")
                                self.driver.refresh()
                                _wait_table_rows(timeout_sec=12)
                                try:
                                    from selenium.webdriver.support.ui import Select as SeleniumSelect
                                    sel_el = self.driver.find_element(
                                        By.CSS_SELECTOR,
                                        "select[name*='DataTables'], .dataTables_length select, select[name$='_length']"
                                    )
                                    for sel_val in ["100", "50", "25"]:
                                        try:
                                            SeleniumSelect(sel_el).select_by_value(sel_val)
                                            time.sleep(1.0)
                                            break
                                        except:
                                            continue
                                except:
                                    pass
                                search_input = _find_search_input()
                                if not search_input:
                                    break

                            for filtro_texto in filtros_fecha:
                                self.log(f"  Filtrando DataTable por fecha/hora: '{filtro_texto}'")
                                n = _apply_filter_and_count(search_input, filtro_texto, timeout_sec=7)
                                if n > 0:
                                    self.log(f"  Filtrado OK — {n} fila(s) con '{filtro_texto}'.")
                                    filtro_exitoso = True
                                    break
                            if filtro_exitoso:
                                break

                    if not filtro_exitoso:
                        self.log(f"  Filtrando DataTable por ubicación: '{ubicacion}'")
                        n_ubi = _apply_filter_and_count(search_input, ubicacion, timeout_sec=8)
                        if n_ubi > 0:
                            self.log(f"  Filtrado OK — {n_ubi} fila(s) con '{ubicacion}'.")
                        else:
                            self.log("  (Filtro por ubicación sin resultados inmediatos; se escanearán filas visibles)")
                except Exception as e_search:
                    self.log(f"  (Error al filtrar: {e_search})")
            else:
                self.log("  ⚠️ No se encontró buscador de tabla; se escanea toda la tabla.")

            time.sleep(0.3)


            # ── PASO 3: Leer TODAS las filas visibles (excluir 'Sin información') ──
            rows = self.driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
            visible_rows = [
                r for r in rows
                if r.is_displayed()
                and r.text.strip()
                and "Cargando" not in r.text
                and "Sin información" not in r.text
            ]
            self.log(f"  Filas visibles: {len(visible_rows)}")

            # ── PASO 4: Debug — mostrar primeras 5 filas ──
            for i, row in enumerate(visible_rows[:5]):
                try:
                    texto_debug = row.text[:150].replace('\n', ' | ')
                    self.log(f"    [Fila {i+1}]: {texto_debug}")
                except:
                    pass

            # ── PASO 5: Evaluar candidatas ──
            ots_candidatas = []

            for row in visible_rows:
                try:
                    row_text = row.text

                    # CRITERIO 1: Código OT tipo PCKM (obligatorio)
                    match_ot = re.search(r'(PCKM\d{6,15})', row_text)
                    if not match_ot:
                        continue

                    # CRITERIO 2: Timestamp de la fila (columna Estado Actual)
                    # La fecha aparece debajo del estado, ej: 2026-02-27T10:14:08.17
                    match_hora = re.search(
                        r'(\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}:\d{2}(?:\.\d+)?)', row_text
                    )
                    hora_str = match_hora.group(1) if match_hora else "sin fecha"

                    fecha_dt = self._parse_wms_timestamp(match_hora.group(1)) if match_hora else None

                    # Si tenemos tiempo_antes_crear, SOLO aceptar OTs creadas después
                    # Esta es la comparación principal — la ubicación y estado son secundarios
                    if tiempo_antes_crear and fecha_dt:
                        if fecha_dt < tiempo_antes_crear:
                            continue  # OT demasiado antigua, saltar

                    ot_code = match_ot.group(1)
                    num_match = re.search(r'PCKM0*(\d+)', ot_code)
                    ot_num = int(num_match.group(1)) if num_match else 0
                    row_upper = row_text.upper()

                    self.log(f"    ✅ Candidata: {ot_code} | Fecha: {hora_str}")
                    ots_candidatas.append({
                        'codigo': ot_code,
                        'numero': ot_num,
                        'hora_str': hora_str,
                        'fecha_dt': fecha_dt,
                        'creada_ok': "CREADA" in row_upper,
                        'ubicacion_ok': ubicacion in row_text,
                    })
                except Exception as ex:
                    self.log(f"    (Error procesando fila: {ex})")
                    continue

            # ── PASO 6: Elegir la OT correcta ──
            if ots_candidatas:
                self.log(f"  Total candidatas: {len(ots_candidatas)}")

                elegida = None

                elegida = self._pick_ot_candidate(
                    ots_candidatas, tiempo_antes_crear, ubicacion
                )
                if elegida:
                    self.log(f"  ✅ OT elegida: {elegida['codigo']} | {elegida['hora_str']}")
                else:
                    self.log("  (Sin candidata confiable dentro de la ventana de tiempo)")

                if elegida:
                    ot_number = elegida['codigo']

            else:
                self.log(f"  ⚠️ No se encontró ninguna OT con CREADA + ubicación '{ubicacion}'.")
                self.log("  Esperando 3s y reintentando (la OT puede tardar en aparecer)...")
                time.sleep(3)

                # Refresh y re-filtrar
                self.driver.refresh()
                time.sleep(3)

                # Re-ampliar paginación
                try:
                    from selenium.webdriver.support.ui import Select as _SelR
                    for _v in ["100", "50"]:
                        try:
                            _el = self.driver.find_element(
                                By.CSS_SELECTOR,
                                "select[name*='DataTables'], .dataTables_length select, select[name$='_length']"
                            )
                            _SelR(_el).select_by_value(_v)
                            time.sleep(1)
                            break
                        except:
                            continue
                except:
                    pass

                # Re-filtrar por ubicación
                try:
                    _inp = None
                    for _sel in ["input[type='search']", ".dataTables_filter input"]:
                        try:
                            _inputs = self.driver.find_elements(By.CSS_SELECTOR, _sel)
                            for _i in _inputs:
                                if _i.is_displayed() and _i.is_enabled():
                                    _inp = _i
                                    break
                            if _inp:
                                break
                        except:
                            continue
                    if _inp:
                        _inp.click()
                        _inp.send_keys(Keys.CONTROL + "a")
                        _inp.send_keys(Keys.DELETE)
                        _inp.send_keys(ubicacion)
                        time.sleep(2)
                except:
                    pass

                # Re-evaluar filas
                _rows2 = self.driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
                _vis2 = [r for r in _rows2
                         if r.is_displayed() and r.text.strip()
                         and "Cargando" not in r.text and "Sin información" not in r.text]
                self.log(f"  Filas visibles (reintento): {len(_vis2)}")

                for row in _vis2:
                    try:
                        row_text = row.text
                        if ubicacion not in row_text:
                            continue
                        if "CREADA" not in row_text.upper():
                            continue
                        match_ot = re.search(r'(PCKM\d{6,15})', row_text)
                        if not match_ot:
                            continue
                        ot_code = match_ot.group(1)
                        num_match = re.search(r'PCKM0*(\d+)', ot_code)
                        ot_num = int(num_match.group(1)) if num_match else 0
                        match_hora = re.search(
                            r'(\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}:\d{2}(?:\.\d+)?)', row_text)
                        hora_str = match_hora.group(1) if match_hora else "sin fecha"
                        fecha_dt2 = self._parse_wms_timestamp(match_hora.group(1)) if match_hora else None
                        ots_candidatas.append({
                            'codigo': ot_code,
                            'numero': ot_num,
                            'hora_str': hora_str,
                            'fecha_dt': fecha_dt2,
                            'creada_ok': True,
                            'ubicacion_ok': True,
                        })
                        self.log(f"    ✅ Candidata (reintento): {ot_code} | {hora_str}")
                    except:
                        continue

                if ots_candidatas:
                    elegida_r = self._pick_ot_candidate(
                        ots_candidatas, tiempo_antes_crear, ubicacion
                    )
                    if elegida_r:
                        ot_number = elegida_r['codigo']
                        self.log(f"  ✅ OT encontrada en reintento: {ot_number} | {elegida_r['hora_str']}")

                # ── FALLBACK FINAL: buscar por timestamp sin filtro de ubicación ni estado ──

                # Esto sirve cuando la ubicación no se selección en el paso 2
                if tiempo_antes_crear and ot_number is None:
                    self.log("  Buscando OT por timestamp en TODAS las filas...")
                    # Limpiar el buscador para ver todas las OTs
                    try:
                        if search_input and search_input.is_displayed():
                            search_input.click()
                            search_input.send_keys(Keys.CONTROL + "a")
                            search_input.send_keys(Keys.DELETE)
                            time.sleep(1.5)
                    except:
                        pass

                    # Releer todas las filas
                    all_rows = self.driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
                    all_visible = [
                        r for r in all_rows
                        if r.is_displayed() and r.text.strip()
                        and "Cargando" not in r.text
                        and "Sin información" not in r.text
                    ]
                    self.log(f"  Total filas (sin filtro): {len(all_visible)}")

                    candidatas_fb = []
                    for row in all_visible:
                        try:
                            row_text = row.text
                            m_ot = re.search(r'(PCKM\d{6,15})', row_text)
                            m_hora = re.search(
                                r'(\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}:\d{2}(?:\.\d+)?)', row_text
                            )
                            if not m_ot or not m_hora:
                                continue

                            ot_code = m_ot.group(1)
                            ts = m_hora.group(1)
                            fecha_dt = self._parse_wms_timestamp(ts)

                            if fecha_dt and fecha_dt >= tiempo_antes_crear:
                                num_match = re.search(r'PCKM0*(\d+)', ot_code)
                                ot_num = int(num_match.group(1)) if num_match else 0
                                candidatas_fb.append({
                                    'codigo': ot_code,
                                    'numero': ot_num,
                                    'hora_str': ts,
                                    'fecha_dt': fecha_dt,
                                    'creada_ok': "CREADA" in row_text.upper(),
                                    'ubicacion_ok': ubicacion in row_text,
                                })
                                self.log(f"    📦 Candidata (sin filtro): {ot_code} | {ts}")
                        except:
                            continue

                    if candidatas_fb:
                        elegida_fb = self._pick_ot_candidate(
                            candidatas_fb, tiempo_antes_crear, ubicacion
                        )
                        if elegida_fb:
                            ot_number = elegida_fb['codigo']
                            self.log(f"  ✅ OT encontrada por TIMESTAMP (fallback): {ot_number} | {elegida_fb['hora_str']}")
                        else:
                            self.log("  ⚠️ Candidatas detectadas, pero ninguna fue confiable por tiempo.")
                    else:
                        self.log("  ⚠️ No se encontró OT por timestamp. Revisa manualmente el WMS.")
                elif ot_number is None:
                    self.log("  La OT fue creada pero no pudo identificarse automáticamente.")
                    self.log("  Revisa manualmente el listado en el WMS.")

        except Exception as e:
            self.log(f"  ❌ Error capturando OT: {e}")

        return ot_number
    
    def run(self, references):
        self.log(f"\n{'='*50}")
        self.log(f"WMS {self.canal.upper()} AUTOMATION")
        self.log(f"{'='*50}")
        self.log(f"Operador: {self.operador}")
        self.log(f"Ubicacion destino aplicada: {self.config['ubicacion']}")
        self.log(f"Órdenes: {len(references)} | Destino: {self.config['ubicacion']}")
        
        self.setup_driver()
        
        if not self.login():
            self.log("Error en login")
            self.driver.quit()
            return

        # Verificar que la sesión quedó activa post-login
        if not self.is_session_alive():
            self.log("❌ No se pudo establecer sesión. Verifica credenciales.")
            self.driver.quit()
            return

        if not self.navigate_to_monitor():
            self.log("Error: La tabla no cargó")
            self.driver.quit()
            return
        
        start = datetime.now()
        self.process_batch(references)
        elapsed = datetime.now() - start
        
        # RESUMEN
        self.log(f"\n{'='*50}")
        self.log("RESUMEN")
        self.log(f"{'='*50}")
        self.log(f"Operador: {self.operador}")
        if self.ot_generada:
            self.log(f"📋 OT Generada: {self.ot_generada}")
        self.log(f"Tiempo: {elapsed}")
        self.log(f"Procesadas: {len(self.orders_selected)}")
        self.log(f"No encontradas: {len(self.orders_not_found)}")
        self.log(f"SKUs sin stock: {len(self.skus_sin_stock)}")
        
        if self.orders_not_found:
            self.log(f"\n❌ NO ENCONTRADAS:")
            for ref in self.orders_not_found:
                self.log(f"  -> {ref}")
        
        if self.skus_sin_stock:
            self.log(f"\n🔴 SKUs SIN STOCK (Banderas Rojas):")
            for sku in self.skus_sin_stock:
                self.log(f"  -> {sku}")
        
        self.log("\nCerrando navegador...")
        time.sleep(2)
        self.driver.quit()
        self.log("Finalizado.")
    
    def stop(self):
        self.running = False

# ============== INTERFAZ GRÁFICA MEJORADA ==============

# Colores por canal
CANAL_COLORS = {
    "Falabella": {"primary": "#28a745", "secondary": "#1e7e34", "bg": "#1a1a1a"},      # Verde
    "Mercadolibre": {"primary": "#FFE600", "secondary": "#CCB800", "bg": "#1a1a1a"},   # Amarillo
    "Walmart": {"primary": "#17a2b8", "secondary": "#138496", "bg": "#1a1a1a"},        # Celeste
    "Paris": {"primary": "#001f5b", "secondary": "#001440", "bg": "#1a1a1a"},          # Azul marino
    "Ripley": {"primary": "#dc3545", "secondary": "#c82333", "bg": "#1a1a1a"},         # Rojo
    "Paginas": {"primary": "#9C27B0", "secondary": "#7B1FA2", "bg": "#1a1a1a"}         # Morado
}

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title("DCIC - Sistema de Despachos")
        self.geometry("1400x920")
        self.minsize(1200, 800)
        
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        
        self.pdf_paths = []
        self.references = []
        self.canal_actual = "Falabella"
        self.automation = None
        self.running = False
        self.current_step = 0
        self.authenticated_operator = None
        
        # Configurar drag and drop
        self.drop_target_register = None

        if not self.authenticate_user():
            self.after(100, self.destroy)
            return
        
        self.create_widgets()
        self.apply_canal_theme()
        self.show_git_update_status()

        # Precargar ChromeDriver en background para que esté listo al ejecutar
        threading.Thread(target=preload_driver, daemon=True).start()

    def authenticate_user(self):
        """Muestra login grande con selector de usuario y PIN."""
        max_attempts = 3
        state = {"success": False, "cancelled": False, "attempts": 0}

        login = ctk.CTkToplevel(self)
        login.title("Inicio de sesion")
        login.geometry("560x360")
        login.resizable(False, False)
        login.transient(self)
        login.grab_set()

        self.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() // 2) - 280
        y = self.winfo_y() + (self.winfo_height() // 2) - 180
        login.geometry(f"+{max(0, x)}+{max(0, y)}")

        frame = ctk.CTkFrame(login, corner_radius=14)
        frame.pack(fill="both", expand=True, padx=16, pady=16)

        ctk.CTkLabel(
            frame,
            text="Acceso Operador",
            font=ctk.CTkFont(size=30, weight="bold")
        ).pack(pady=(20, 10))

        ctk.CTkLabel(
            frame,
            text="Selecciona usuario y escribe PIN",
            font=ctk.CTkFont(size=14),
            text_color="#b0b0b0"
        ).pack(pady=(0, 14))

        user_var = ctk.StringVar(value=OPERADORES[0])
        user_menu = ctk.CTkOptionMenu(
            frame,
            values=OPERADORES,
            variable=user_var,
            width=300,
            height=42,
            font=ctk.CTkFont(size=18, weight="bold"),
        )
        user_menu.pack(pady=(0, 10))

        pin_entry = ctk.CTkEntry(
            frame,
            width=300,
            height=44,
            show="*",
            placeholder_text="PIN",
            font=ctk.CTkFont(size=18),
        )
        pin_entry.pack(pady=(0, 8))

        status_lbl = ctk.CTkLabel(
            frame,
            text=f"Intentos: 0/{max_attempts}",
            font=ctk.CTkFont(size=13),
            text_color="#d0d0d0",
        )
        status_lbl.pack(pady=(0, 12))

        def close_as_cancel():
            state["cancelled"] = True
            try:
                login.grab_release()
            except:
                pass
            login.destroy()

        def try_login(event=None):
            operador = user_var.get().strip()
            pin_in = pin_entry.get().strip()

            if pin_in == USER_PINS.get(operador, ""):
                self.authenticated_operator = operador
                state["success"] = True
                try:
                    login.grab_release()
                except:
                    pass
                login.destroy()
                return

            state["attempts"] += 1
            if state["attempts"] >= max_attempts:
                messagebox.showerror("Bloqueado", "Se supero el numero maximo de intentos.", parent=login)
                close_as_cancel()
                return

            status_lbl.configure(
                text=f"PIN incorrecto. Intentos: {state['attempts']}/{max_attempts}",
                text_color="#ff6b6b",
            )
            pin_entry.delete(0, "end")
            pin_entry.focus_set()

        btn_row = ctk.CTkFrame(frame, fg_color="transparent")
        btn_row.pack(pady=(6, 10))

        login_btn = ctk.CTkButton(
            btn_row,
            text="Ingresar",
            width=140,
            height=42,
            command=try_login,
            font=ctk.CTkFont(size=16, weight="bold"),
        )
        login_btn.pack(side="left", padx=8)

        cancel_btn = ctk.CTkButton(
            btn_row,
            text="Cancelar",
            width=120,
            height=42,
            command=close_as_cancel,
            fg_color="#666666",
            hover_color="#555555",
            font=ctk.CTkFont(size=15),
        )
        cancel_btn.pack(side="left", padx=8)

        def apply_user_color(*_):
            operador = user_var.get().strip()
            base = USER_COLORS.get(operador, "#2b2b2b")
            text_color = "#111111" if base.lower() == "#ffffff" else "#ffffff"
            user_menu.configure(
                fg_color=base,
                button_color=base,
                button_hover_color=base,
                text_color=text_color,
            )
            login_btn.configure(fg_color=base, hover_color=base, text_color=text_color)

        user_var.trace_add("write", apply_user_color)
        apply_user_color()

        login.protocol("WM_DELETE_WINDOW", close_as_cancel)
        login.bind("<Return>", try_login)
        login.bind("<Escape>", lambda e: close_as_cancel())
        pin_entry.focus_set()

        self.wait_window(login)
        return state["success"] and not state["cancelled"]

    def show_git_update_status(self):
        """Muestra en log el resultado de auto-actualizacion git (si existe)."""
        try:
            status_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), GIT_UPDATE_STATUS_FILE)
            if not os.path.exists(status_path):
                return

            with open(status_path, "r", encoding="utf-8", errors="ignore") as f:
                txt = f.read().strip()

            low = txt.lower()
            if "already up to date" in low or "already up-to-date" in low:
                self.log("Git: sin actualizaciones (ya estaba al dia).", "info")
            elif "fast-forward" in low or "updating " in low:
                self.log("Git: actualizacion descargada correctamente.", "success")
            elif "no_git_repo" in low:
                self.log("Git: carpeta sin repositorio .git (sin auto-actualizacion).", "warning")
            elif "fatal:" in low or "error" in low:
                self.log("Git: no se pudo actualizar automaticamente.", "warning")
            else:
                self.log("Git: verificacion de actualizaciones completada.", "info")

            try:
                os.remove(status_path)
            except:
                pass
        except Exception as e:
            self.log(f"Git: error leyendo estado de actualizacion ({e})", "warning")
    
    def create_widgets(self):
        # Frame principal con gradiente
        self.main_frame = ctk.CTkFrame(self, fg_color="#1a1a1a")
        self.main_frame.pack(fill="both", expand=True, padx=0, pady=0)
        
        # ===== HEADER =====
        self.header_frame = ctk.CTkFrame(self.main_frame, height=80, fg_color="#2d2d2d")
        self.header_frame.pack(fill="x", padx=0, pady=0)
        self.header_frame.pack_propagate(False)
        
        # Logo/Título
        self.title_label = ctk.CTkLabel(
            self.header_frame, 
            text="⚡ Automatización DCIC",
            font=ctk.CTkFont(size=28, weight="bold"),
            text_color="#ffffff"
        )
        self.title_label.pack(side="left", padx=20, pady=20)
        
        # Subtítulo
        self.subtitle_label = ctk.CTkLabel(
            self.header_frame,
            text="Sistema de Despachos",
            font=ctk.CTkFont(size=14),
            text_color="#888888"
        )
        self.subtitle_label.pack(side="left", padx=5, pady=20)
        
        # Selector de canal (derecha)
        self.canal_frame = ctk.CTkFrame(self.header_frame, fg_color="transparent")
        self.canal_frame.pack(side="right", padx=20, pady=15)
        
        self.canal_label = ctk.CTkLabel(
            self.canal_frame, 
            text="Canal:", 
            font=ctk.CTkFont(size=14),
            text_color="#888888"
        )
        self.canal_label.pack(side="left", padx=5)
        
        self.canal_var = ctk.StringVar(value="Falabella")
        self.canal_menu = ctk.CTkOptionMenu(
            self.canal_frame,
            values=list(CANALES.keys()),
            variable=self.canal_var,
            command=self.on_canal_change,
            width=160,
            height=35,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#FF6B00",
            button_color="#CC5500",
            button_hover_color="#AA4400"
        )
        self.canal_menu.pack(side="left", padx=5)

        self.operador_label = ctk.CTkLabel(
            self.canal_frame,
            text="Operador:",
            font=ctk.CTkFont(size=14),
            text_color="#888888"
        )
        self.operador_label.pack(side="left", padx=(10, 5))

        default_operador = self.authenticated_operator or OPERADORES[0]
        self.operador_var = ctk.StringVar(value=default_operador)
        self.operador_menu = ctk.CTkOptionMenu(
            self.canal_frame,
            values=OPERADORES,
            variable=self.operador_var,
            width=120,
            height=35,
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color="#3a3a3a",
            button_color="#2b2b2b",
            button_hover_color="#1f1f1f"
        )
        self.operador_menu.pack(side="left", padx=5)
        # Sesión bloqueada al operador autenticado
        if self.authenticated_operator:
            self.operador_menu.configure(state="disabled")
        
        # ===== BARRA DE PROGRESO =====
        self.progress_frame = ctk.CTkFrame(self.main_frame, height=60, fg_color="#252525")
        self.progress_frame.pack(fill="x", padx=0, pady=0)
        self.progress_frame.pack_propagate(False)
        
        self.steps = ["📋 PDFs", "🔍 Extraer", "✅ Seleccionar", "📍 Ubicación", "📦 Stock", "🎫 Crear OT"]
        self.step_labels = []
        self.step_indicators = []
        
        steps_container = ctk.CTkFrame(self.progress_frame, fg_color="transparent")
        steps_container.pack(expand=True, pady=10)
        
        for i, step in enumerate(self.steps):
            frame = ctk.CTkFrame(steps_container, fg_color="transparent")
            frame.pack(side="left", padx=15)
            
            indicator = ctk.CTkLabel(
                frame,
                text="○",
                font=ctk.CTkFont(size=20),
                text_color="#555555"
            )
            indicator.pack()
            self.step_indicators.append(indicator)
            
            label = ctk.CTkLabel(
                frame,
                text=step,
                font=ctk.CTkFont(size=11),
                text_color="#666666"
            )
            label.pack()
            self.step_labels.append(label)
        
        # ===== CONTENIDO PRINCIPAL =====
        self.content_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.content_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # Panel izquierdo (PDFs y Referencias)
        self.left_panel = ctk.CTkFrame(self.content_frame, fg_color="#252525", corner_radius=15)
        self.left_panel.pack(side="left", fill="both", expand=True, padx=(0, 10), pady=5)
        
        # Área de Drop (PDFs)
        self.drop_frame = ctk.CTkFrame(
            self.left_panel, 
            height=120, 
            fg_color="#1e1e1e",
            border_width=2,
            border_color="#444444",
            corner_radius=10
        )
        self.drop_frame.pack(fill="x", padx=15, pady=15)
        self.drop_frame.pack_propagate(False)
        
        self.drop_label = ctk.CTkLabel(
            self.drop_frame,
            text="📂 Haz clic para seleccionar PDFs\no arrástralos aquí",
            font=ctk.CTkFont(size=14),
            text_color="#888888"
        )
        self.drop_label.pack(expand=True)
        
        # Hacer el frame clickeable
        self.drop_frame.bind("<Button-1>", lambda e: self.select_pdfs())
        self.drop_label.bind("<Button-1>", lambda e: self.select_pdfs())
        
        # Lista de PDFs seleccionados
        self.pdf_listbox = ctk.CTkTextbox(
            self.left_panel, 
            height=60,
            fg_color="#1e1e1e",
            text_color="#cccccc",
            font=ctk.CTkFont(size=12)
        )
        self.pdf_listbox.pack(fill="x", padx=15, pady=(0, 10))
        self.pdf_listbox.configure(state="disabled")
        
        # Botón extraer
        self.extract_btn = ctk.CTkButton(
            self.left_panel,
            text="🔍 EXTRAER REFERENCIAS",
            command=self.extract_references,
            height=45,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#FF6B00",
            hover_color="#CC5500"
        )
        self.extract_btn.pack(fill="x", padx=15, pady=5)
        
        # Referencias encontradas
        self.ref_header = ctk.CTkFrame(self.left_panel, fg_color="transparent")
        self.ref_header.pack(fill="x", padx=15, pady=(10, 5))
        
        self.ref_label = ctk.CTkLabel(
            self.ref_header,
            text="Referencias encontradas: 0",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color="#ffffff"
        )
        self.ref_label.pack(side="left")
        
        self.ref_textbox = ctk.CTkTextbox(
            self.left_panel,
            fg_color="#1e1e1e",
            text_color="#00ff00",
            font=ctk.CTkFont(family="Consolas", size=12)
        )
        self.ref_textbox.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        self.ref_textbox.configure(state="disabled")
        
        # Panel derecho (Log)
        self.right_panel = ctk.CTkFrame(self.content_frame, fg_color="#252525", corner_radius=15)
        self.right_panel.pack(side="right", fill="both", expand=True, padx=(10, 0), pady=5)
        
        # Header del log
        self.log_header = ctk.CTkFrame(self.right_panel, fg_color="transparent")
        self.log_header.pack(fill="x", padx=15, pady=(15, 5))
        
        self.log_title = ctk.CTkLabel(
            self.log_header,
            text="📋 Log de Ejecución",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color="#ffffff"
        )
        self.log_title.pack(side="left")
        
        # Textbox del log
        self.log_textbox = ctk.CTkTextbox(
            self.right_panel,
            fg_color="#0d0d0d",
            font=ctk.CTkFont(family="Consolas", size=11)
        )
        self.log_textbox.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        self.log_textbox.configure(state="disabled")
        
        # ===== FOOTER (Controles) =====
        self.footer_frame = ctk.CTkFrame(self.main_frame, height=80, fg_color="#2d2d2d")
        self.footer_frame.pack(fill="x", padx=0, pady=0)
        self.footer_frame.pack_propagate(False)
        
        # Botones de control
        self.control_container = ctk.CTkFrame(self.footer_frame, fg_color="transparent")
        self.control_container.pack(expand=True, pady=15)
        
        self.start_btn = ctk.CTkButton(
            self.control_container,
            text="▶️  EJECUTAR AUTOMATIZACIÓN",
            command=self.start_automation,
            width=280,
            height=50,
            font=ctk.CTkFont(size=16, weight="bold"),
            fg_color="#28a745",
            hover_color="#218838"
        )
        self.start_btn.pack(side="left", padx=10)
        
        self.stop_btn = ctk.CTkButton(
            self.control_container,
            text="⏹️  DETENER",
            command=self.stop_automation,
            width=150,
            height=50,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#dc3545",
            hover_color="#c82333",
            state="disabled"
        )
        self.stop_btn.pack(side="left", padx=10)
        
        # Botón NUEVO (limpiar todo)
        self.new_btn = ctk.CTkButton(
            self.control_container,
            text="🔄  NUEVO",
            command=self.reset_all,
            width=120,
            height=50,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#6c757d",
            hover_color="#5a6268"
        )
        self.new_btn.pack(side="left", padx=10)
        
        # Status
        self.status_label = ctk.CTkLabel(
            self.control_container,
            text="● Listo",
            font=ctk.CTkFont(size=12),
            text_color="#28a745"
        )
        self.status_label.pack(side="left", padx=20)
    def reset_all(self):
        """Limpia todo para empezar de nuevo."""
        # Limpiar PDFs
        self.pdf_paths = []
        self.references = []
        
        # Reset área de drop
        self.drop_label.configure(
            text="📂 Haz clic para seleccionar PDFs\no arrástralos aquí",
            text_color="#888888"
        )
        
        # Limpiar lista de PDFs
        self.pdf_listbox.configure(state="normal")
        self.pdf_listbox.delete("1.0", "end")
        self.pdf_listbox.configure(state="disabled")
        
        # Limpiar referencias
        self.ref_label.configure(text="Referencias encontradas: 0")
        self.ref_textbox.configure(state="normal")
        self.ref_textbox.delete("1.0", "end")
        self.ref_textbox.configure(state="disabled")
        
        # Limpiar log
        self.log_textbox.configure(state="normal")
        self.log_textbox.delete("1.0", "end")
        self.log_textbox.configure(state="disabled")
        
        # Reset progreso
        self.update_progress(0)
        
        # Reset status
        self.status_label.configure(text="● Listo", text_color="#28a745")
        
        # Log mensaje
        self.log("🔄 Listo para nuevo proceso", "info")
    
    def apply_canal_theme(self):
        """Aplica los colores según el canal seleccionado."""
        colors = CANAL_COLORS.get(self.canal_actual, CANAL_COLORS["Falabella"])
        primary = colors["primary"]
        secondary = colors["secondary"]
        
        # Actualizar colores
        self.canal_menu.configure(fg_color=primary, button_color=secondary)
        self.extract_btn.configure(fg_color=primary, hover_color=secondary)
        self.title_label.configure(text_color=primary)
        
        # Actualizar borde del drop area
        self.drop_frame.configure(border_color=primary)
    
    def update_progress(self, step):
        """Actualiza la barra de progreso visual."""
        self.current_step = step
        
        colors = CANAL_COLORS.get(self.canal_actual, CANAL_COLORS["Falabella"])
        primary = colors["primary"]
        
        for i, (indicator, label) in enumerate(zip(self.step_indicators, self.step_labels)):
            if i < step:
                # Completado
                indicator.configure(text="●", text_color="#28a745")
                label.configure(text_color="#28a745")
            elif i == step:
                # Actual
                indicator.configure(text="◉", text_color=primary)
                label.configure(text_color=primary)
            else:
                # Pendiente
                indicator.configure(text="○", text_color="#555555")
                label.configure(text_color="#666666")
    
    def on_canal_change(self, value):
        self.canal_actual = value
        self.apply_canal_theme()
        self.log(f"🔄 Canal cambiado a: {value}", "info")
        
        if self.pdf_paths:
            self.extract_references()
    
    def select_pdfs(self):
        files = filedialog.askopenfilenames(
            title="Seleccionar Manifiestos PDF",
            filetypes=[("PDF", "*.pdf")]
        )
        
        if files:
            self.pdf_paths = list(files)
            self.update_progress(1)
            
            # Actualizar área de drop
            self.drop_label.configure(
                text=f"✅ {len(self.pdf_paths)} archivo(s) seleccionado(s)",
                text_color="#28a745"
            )
            
            # Mostrar lista de PDFs
            self.pdf_listbox.configure(state="normal")
            self.pdf_listbox.delete("1.0", "end")
            for path in self.pdf_paths:
                self.pdf_listbox.insert("end", f"📄 {os.path.basename(path)}\n")
            self.pdf_listbox.configure(state="disabled")
            
            self.log(f"📂 Seleccionados {len(self.pdf_paths)} archivos PDF", "info")
            
            # Detectar canal automáticamente
            detected = detect_canal_from_pdf(self.pdf_paths[0])
            if detected:
                self.canal_actual = detected
                self.canal_var.set(detected)
                self.apply_canal_theme()
                self.log(f"🔍 Canal detectado: {detected}", "success")
    
    def extract_references(self):
        if not self.pdf_paths:
            messagebox.showwarning("Aviso", "Primero selecciona archivos PDF")
            return
        
        self.status_label.configure(text="● Extrayendo...", text_color="#FFE600")
        self.update_progress(2)
        self.update()
        
        self.references = extract_references(self.pdf_paths, self.canal_actual)
        
        # Mostrar referencias
        self.ref_label.configure(text=f"Referencias encontradas: {len(self.references)}")
        
        self.ref_textbox.configure(state="normal")
        self.ref_textbox.delete("1.0", "end")
        for i, ref in enumerate(self.references, 1):
            self.ref_textbox.insert("end", f"{i:3}. {ref}\n")
        self.ref_textbox.configure(state="disabled")
        
        self.status_label.configure(text="● Listo", text_color="#28a745")
        self.log(f"✅ Extraídas {len(self.references)} referencias para {self.canal_actual}", "success")
    
    def log(self, message, msg_type="normal"):
        """Agrega mensaje al log con color según tipo."""
        self.log_textbox.configure(state="normal")
        
        # Agregar timestamp
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        # Configurar tags para colores (solo una vez)
        try:
            self.log_textbox._textbox.tag_configure("ot_highlight", foreground="#00BFFF", font=("Consolas", 12, "bold"))
            self.log_textbox._textbox.tag_configure("error_red", foreground="#FF4444")  # Solo color, tamaño normal
            self.log_textbox._textbox.tag_configure("success_green", foreground="#00FF00")
            self.log_textbox._textbox.tag_configure("warning_yellow", foreground="#FFD700")
        except:
            pass
        
        # Formatear mensaje según tipo
        if msg_type == "success":
            prefix = "✅"
        elif msg_type == "error":
            prefix = "❌"
        elif msg_type == "warning":
            prefix = "⚠️"
        elif msg_type == "info":
            prefix = "ℹ️"
        else:
            prefix = "  "
        
        full_message = f"[{timestamp}] {prefix} {message}\n"
        
        # Insertar con color especial según contenido
        try:
            # OT - Azul y grande
            if "Número de OT:" in message or "OT Generada:" in message or "OT encontrada:" in message:
                self.log_textbox._textbox.insert("end", full_message, "ot_highlight")
            # Sin stock - Rojo
            elif "SIN STOCK" in message or "🔴" in message or "Banderas Rojas" in message:
                self.log_textbox._textbox.insert("end", full_message, "error_red")
            # NO ENCONTRADAS - Rojo
            elif "NO ENCONTRADAS" in message or "NO ENCONTRADA" in message:
                self.log_textbox._textbox.insert("end", full_message, "error_red")
            # Éxito
            elif msg_type == "success" or "EXITOSAMENTE" in message:
                self.log_textbox._textbox.insert("end", full_message, "success_green")
            # Warning
            elif msg_type == "warning" or "ADVERTENCIA" in message:
                self.log_textbox._textbox.insert("end", full_message, "warning_yellow")
            else:
                self.log_textbox.insert("end", full_message)
        except:
            self.log_textbox.insert("end", full_message)
        
        self.log_textbox.see("end")
        self.log_textbox.configure(state="disabled")
        self.update()
    
    def start_automation(self):
        if not self.references:
            messagebox.showwarning("Aviso", "No hay referencias para procesar.\nExtrae las referencias primero.")
            return
        
        if not messagebox.askyesno("Confirmar", f"¿Procesar {len(self.references)} referencias de {self.canal_actual}?"):
            return
        
        self.running = True
        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.extract_btn.configure(state="disabled")
        self.canal_menu.configure(state="disabled")
        self.operador_menu.configure(state="disabled")
        self.status_label.configure(text="● Ejecutando...", text_color="#FFE600")
        
        # Ejecutar en thread separado
        thread = threading.Thread(target=self.run_automation)
        thread.daemon = True
        thread.start()
    
    def run_automation(self):
        try:
            # Wrapper del log para actualizar progreso
            def log_wrapper(msg):
                self.log(msg)
                # Detectar paso actual
                if "[1/5]" in msg:
                    self.after(0, lambda: self.update_progress(3))
                elif "[2/5]" in msg:
                    self.after(0, lambda: self.update_progress(4))
                elif "[3/5]" in msg:
                    self.after(0, lambda: self.update_progress(5))
                elif "[4/5]" in msg or "[5/5]" in msg:
                    self.after(0, lambda: self.update_progress(6))
            
            operador = (self.authenticated_operator or self.operador_var.get().strip() or "Sin definir")
            self.automation = WMSAutomation(
                self.canal_actual,
                log_callback=log_wrapper,
                operador=operador
            )
            self.automation.run(self.references.copy())
        except Exception as e:
            self.log(f"Error: {e}", "error")
        finally:
            self.running = False
            self.after(0, self.on_automation_complete)
    
    def on_automation_complete(self):
        self.start_btn.configure(state="normal")
        self.stop_btn.configure(state="disabled")
        self.extract_btn.configure(state="normal")
        self.canal_menu.configure(state="normal")
        if self.authenticated_operator:
            self.operador_menu.configure(state="disabled")
        else:
            self.operador_menu.configure(state="normal")
        self.status_label.configure(text="● Completado", text_color="#28a745")
        self.update_progress(6)
        self.append_ot_audit_row()
        
        # Traer ventana al frente
        self.lift()
        self.focus_force()
        
        # Notificación sonora (3 beeps)
        try:
            for _ in range(3):
                winsound.Beep(800, 200)  # Frecuencia 800Hz, duración 200ms
                time.sleep(0.1)
            winsound.PlaySound("SystemExclamation", winsound.SND_ALIAS)
        except:
            pass
        
        # Mostrar mensaje popup
        messagebox.showinfo("✅ Completado", f"Automatización de {self.canal_actual} finalizada.\\n\\nRevisa el log para ver el resumen.")
    
    def append_ot_audit_row(self):
        """Guarda trazabilidad local de OT por operador."""
        try:
            if not self.automation:
                return

            ot_code = (self.automation.ot_generada or "").strip()
            if not ot_code:
                self.log("No se registro OT en historial (no se capturo OT automatica).", "warning")
                return

            row = {
                "fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "operador": self.authenticated_operator or self.operador_var.get().strip() or "Sin definir",
                "canal": self.canal_actual,
                "ot": ot_code,
                "referencias": len(self.references),
                "procesadas": len(self.automation.orders_selected),
                "no_encontradas": len(self.automation.orders_not_found),
                "skus_sin_stock": len(self.automation.skus_sin_stock),
            }

            file_exists = os.path.exists(OT_AUDIT_CSV)
            with open(OT_AUDIT_CSV, "a", newline="", encoding="utf-8-sig") as f:
                writer = csv.DictWriter(
                    f,
                    fieldnames=[
                        "fecha",
                        "operador",
                        "canal",
                        "ot",
                        "referencias",
                        "procesadas",
                        "no_encontradas",
                        "skus_sin_stock",
                    ],
                )
                if not file_exists:
                    writer.writeheader()
                writer.writerow(row)

            self.log(f"Historial OT actualizado: {ot_code} | {row['operador']}", "success")
        except Exception as e:
            self.log(f"Error guardando historial OT: {e}", "error")

    def stop_automation(self):
        if self.automation:
            self.automation.stop()
            self.log("⏹️ Deteniendo automatización...", "warning")


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()

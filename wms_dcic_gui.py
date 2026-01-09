"""
WMS DCIC - Interfaz Unificada v1.0
==================================
Interfaz gráfica para automatización de:
- Falabella
- Mercadolibre Flex
- Walmart

Requiere: pip install customtkinter pdfplumber selenium webdriver-manager
"""

import sys
import os
import re
import time
import threading
import winsound
from datetime import datetime
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
        "patron_busqueda": r'\b(2000\d{12,14})\b',
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
        "patron": r'^243\d{8}-A$',
        "patron_busqueda": r'\b(243\d{8}-A)\b',
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

WMS_URL = "https://checkweb-prd-checkwms.azurewebsites.net/"
MONITOR_URL = "https://checkweb-prd-checkwms.azurewebsites.net/DocumentoDespacho/monitorsalida"
WMS_USER = "18539597"
WMS_PASS = "185395"

WAIT_TIMEOUT = 60
DELAY_STEP = 0.8      # Reducido de 1.5 para mayor velocidad
DELAY_SEARCH = 1.0    # Reducido de 2.0
DELAY_PAGE = 1.5      # Reducido de 3.0
MAX_RETRIES = 3


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
                if not re.search(r'\b243\d{8}-A\b', text):
                    return "Paginas"
            
            # Ripley: 243XXXXXXXX-A (muy específico por el -A)
            if re.search(r'\b243\d{8}-A\b', text):
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
    def __init__(self, canal, log_callback=None):
        self.canal = canal
        self.config = CANALES[canal]
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
    
    def setup_driver(self):
        options = Options()
        options.add_argument("--start-maximized")
        options.add_argument("--disable-notifications")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_experimental_option("excludeSwitches", ["enable-logging", "enable-automation"])
        options.add_experimental_option("useAutomationExtension", False)
        
        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=options)
        self.wait = WebDriverWait(self.driver, WAIT_TIMEOUT)
        self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    
    def js_click(self, element):
        self.driver.execute_script("arguments[0].click();", element)
    
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
                time.sleep(1)
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
            # Esperar un momento para que la tabla cargue completamente
            time.sleep(1)
            
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
    
    def mark_picking_consolidado(self):
        try:
            checkboxes = self.driver.find_elements(By.CSS_SELECTOR, "input[type='checkbox']")
            for cb in checkboxes:
                try:
                    parent = cb.find_element(By.XPATH, "./..")
                    if "picking" in parent.text.lower() and "consolidado" in parent.text.lower():
                        if not cb.is_selected():
                            self.js_click(cb)
                        return True
                except:
                    continue
            
            visible_cbs = [cb for cb in checkboxes if cb.is_displayed()]
            if len(visible_cbs) >= 3 and not visible_cbs[2].is_selected():
                self.js_click(visible_cbs[2])
                return True
        except:
            pass
        return False
    
    def click_crear_ot(self):
        for selector in ["//button[contains(text(), 'Crear OT')]", "//button[contains(text(), 'CREAR OT')]"]:
            try:
                btn = self.driver.find_element(By.XPATH, selector)
                if btn.is_displayed():
                    self.js_click(btn)
                    return True
            except:
                continue
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
        # PASO 2 - Seleccionar ubicación usando CTRL+F del navegador
        self.log(f"[2/5] Ubicación: {ubicacion}")
        ubicacion_found = False
        
        # Esperar a que la tabla cargue
        time.sleep(1)  # Reducido de 2
        
        # MÉTODO 1: Usar CTRL+F del navegador para buscar y hacer scroll
        try:
            from selenium.webdriver.common.action_chains import ActionChains
            
            # Abrir búsqueda del navegador con CTRL+F
            body = self.driver.find_element(By.TAG_NAME, "body")
            body.send_keys(Keys.CONTROL + "f")
            time.sleep(0.3)  # Reducido de 0.5
            
            # Escribir la ubicación en la búsqueda
            actions = ActionChains(self.driver)
            actions.send_keys(ubicacion)
            actions.perform()
            time.sleep(0.5)  # Reducido de 1
            
            # Presionar Enter para ir al resultado
            actions = ActionChains(self.driver)
            actions.send_keys(Keys.ENTER)
            actions.perform()
            time.sleep(0.3)  # Reducido de 0.5
            
            # Cerrar la búsqueda con Escape
            actions = ActionChains(self.driver)
            actions.send_keys(Keys.ESCAPE)
            actions.perform()
            time.sleep(0.3)  # Reducido de 0.5
            
            self.log(f"  Buscando con CTRL+F...")
        except Exception as e:
            self.log(f"  Error en CTRL+F: {e}")
        
        # Ahora buscar y seleccionar el radio button de la ubicación
        time.sleep(0.5)  # Reducido de 1
        
        # Buscar la fila que contiene la ubicación
        try:
            rows = self.driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
            for row in rows:
                try:
                    if ubicacion in row.text:
                        # Hacer scroll adicional para asegurarse
                        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", row)
                        time.sleep(0.2)  # Reducido de 0.3
                        
                        # Buscar el radio button
                        radios = row.find_elements(By.CSS_SELECTOR, "input[type='radio']")
                        if radios:
                            self.js_click(radios[0])
                            self.log(f"  {ubicacion} OK")
                            ubicacion_found = True
                            break
                except:
                    continue
        except:
            pass
        
        # MÉTODO 2: Si aún no encontró, usar JavaScript para buscar en toda la página
        if not ubicacion_found:
            try:
                # Buscar el elemento que contiene el texto
                script = f"""
                    var elements = document.querySelectorAll('table tbody tr');
                    for (var i = 0; i < elements.length; i++) {{
                        if (elements[i].textContent.includes('{ubicacion}')) {{
                            elements[i].scrollIntoView({{block: 'center'}});
                            var radio = elements[i].querySelector('input[type="radio"]');
                            if (radio) {{
                                radio.click();
                                return true;
                            }}
                        }}
                    }}
                    return false;
                """
                result = self.driver.execute_script(script)
                if result:
                    self.log(f"  {ubicacion} OK (JavaScript)")
                    ubicacion_found = True
            except:
                pass
        
        # MÉTODO 3: Último intento - scroll completo hacia abajo y buscar
        if not ubicacion_found:
            try:
                # Scroll hasta el final de la página
                self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(1)
                
                rows = self.driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
                for row in rows:
                    try:
                        if ubicacion in row.text:
                            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", row)
                            time.sleep(0.3)
                            radios = row.find_elements(By.CSS_SELECTOR, "input[type='radio']")
                            if radios:
                                self.js_click(radios[0])
                                self.log(f"  {ubicacion} OK (scroll final)")
                                ubicacion_found = True
                                break
                    except:
                        continue
            except:
                pass
        
        if not ubicacion_found:
            self.log(f"  ⚠️ ADVERTENCIA: {ubicacion} no encontrada")
            self.log(f"  Intentando continuar de todos modos...")
        
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
        self.mark_picking_consolidado()
        time.sleep(DELAY_STEP)
        self.click_crear_ot()
        time.sleep(1.5)
        self.confirm_modal()
        time.sleep(DELAY_PAGE)
        
        # Capturar número de OT generada
        ot_number = self.capture_ot_number()
        
        if ot_number:
            self.log(f"\n🎉 ¡OT CREADA EXITOSAMENTE!")
            self.log(f"📋 Número de OT: {ot_number}")
            self.ot_generada = ot_number
        else:
            self.log("\n¡OT CREADA EXITOSAMENTE!")
        
        return True
    
    def capture_ot_number(self):
        """Captura el número de OT navegando al listado de Órdenes de Trabajo."""
        ot_number = None
        ubicacion = self.config["ubicacion"]
        
        try:
            self.log(f"  Buscando OT con ubicación: {ubicacion}")
            
            # Navegar al listado de Órdenes de Trabajo (URL correcta con /index)
            ot_url = "https://checkweb-prd-checkwms.azurewebsites.net/OrdenTrabajo/index"
            self.driver.get(ot_url)
            time.sleep(3)
            
            # Hacer doble refresh para asegurar datos actualizados
            self.driver.refresh()
            time.sleep(2)
            self.driver.refresh()
            time.sleep(4)  # Esperar más después del segundo refresh
            
            # Esperar a que cargue la tabla (que no diga "0 to 0 of 0")
            for _ in range(20):  # Intentar hasta 20 veces
                try:
                    # Verificar si hay datos en la tabla
                    info_text = self.driver.find_element(By.CSS_SELECTOR, ".dataTables_info, [class*='info']").text
                    if "0 to 0" not in info_text and "0 of 0" not in info_text:
                        break
                except:
                    pass
                time.sleep(1)
            
            time.sleep(3)  # Espera adicional para asegurar datos completos
            
            # USAR BÚSQUEDA RÁPIDA para filtrar por ubicación
            try:
                search_selectors = [
                    "input[type='search']",
                    ".dataTables_filter input",
                    "[aria-label*='Búsqueda']",
                    "input[placeholder*='Búsqueda']"
                ]
                
                search_input = None
                for selector in search_selectors:
                    try:
                        inputs = self.driver.find_elements(By.CSS_SELECTOR, selector)
                        for inp in inputs:
                            if inp.is_displayed():
                                search_input = inp
                                break
                        if search_input:
                            break
                    except:
                        continue
                
                if search_input:
                    # Limpiar y escribir ubicación en búsqueda rápida
                    search_input.clear()
                    search_input.send_keys(ubicacion)
                    time.sleep(2)  # Esperar que filtre
                    self.log(f"  Filtrando por: {ubicacion}")
            except Exception as e:
                self.log(f"  (Búsqueda rápida no disponible)")
            
            # Esperar un poco más para asegurar que la tabla esté cargada
            time.sleep(2)
            
            # Ahora buscar en las filas
            rows = self.driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
            self.log(f"  Revisando {len(rows)} OTs...")
            
            # Obtener hora actual para comparar
            hora_actual = datetime.now()
            
            # Buscar todas las OTs con estado CREADA y nuestra ubicación
            ots_candidatas = []
            
            for row in rows[:20]:  # Revisar más filas
                try:
                    row_text = row.text
                    
                    # Debug: mostrar primeras filas
                    if len(ots_candidatas) == 0 and rows.index(row) < 3:
                        self.log(f"    Fila {rows.index(row)+1}: {row_text[:80]}...")
                    
                    # Verificar que tenga nuestra ubicación, estado CREADA y código PCKM
                    # Ser más flexible: CREADA puede aparecer como texto
                    tiene_ubicacion = ubicacion in row_text
                    tiene_creada = "CREADA" in row_text.upper()
                    tiene_pckm = "PCKM" in row_text
                    
                    if tiene_ubicacion and tiene_creada and tiene_pckm:
                        # Extraer el código PCKM
                        match_ot = re.search(r'(PCKM\d{9,12})', row_text)
                        # Extraer la hora de creación (formato: 2026-01-08T14:01:33.857)
                        match_hora = re.search(r'(\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2})', row_text)
                        
                        if match_ot:
                            ot_code = match_ot.group(1)
                            hora_creacion = None
                            
                            if match_hora:
                                try:
                                    hora_creacion = datetime.strptime(match_hora.group(1), "%Y-%m-%dT%H:%M:%S")
                                except:
                                    pass
                            
                            ots_candidatas.append({
                                'codigo': ot_code,
                                'hora': hora_creacion,
                                'texto': row_text[:50]
                            })
                except:
                    continue
            
            # Si hay candidatas, tomar la de número más alto (más reciente)
            if ots_candidatas:
                self.log(f"  Encontradas {len(ots_candidatas)} OTs con ubicación correcta")
                
                # Ordenar por número de OT (el más alto es el más reciente)
                # Extraer solo los números del código PCKM para comparar
                for ot in ots_candidatas:
                    try:
                        # Extraer número del código (ej: PCKM000098424 -> 98424)
                        num_match = re.search(r'PCKM0*(\d+)', ot['codigo'])
                        if num_match:
                            ot['numero'] = int(num_match.group(1))
                        else:
                            ot['numero'] = 0
                    except:
                        ot['numero'] = 0
                
                # Ordenar por número (más alto primero)
                ots_candidatas.sort(key=lambda x: x['numero'], reverse=True)
                
                ot_number = ots_candidatas[0]['codigo']
                self.log(f"  ✅ OT más reciente (número más alto): {ot_number}")
            
            # Si no encontró ninguna
            if not ot_number:
                self.log("  ⚠️ No se encontró OT con esa ubicación")
                        
        except Exception as e:
            self.log(f"  ❌ Error capturando OT: {e}")
        
        return ot_number
    
    def run(self, references):
        self.log(f"\n{'='*50}")
        self.log(f"WMS {self.canal.upper()} AUTOMATION")
        self.log(f"{'='*50}")
        self.log(f"Órdenes: {len(references)} | Destino: {self.config['ubicacion']}")
        
        self.setup_driver()
        
        if not self.login():
            self.log("Error en login")
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
        self.geometry("1000x800")
        self.minsize(900, 700)
        
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        
        self.pdf_paths = []
        self.references = []
        self.canal_actual = "Falabella"
        self.automation = None
        self.running = False
        self.current_step = 0
        
        # Configurar drag and drop
        self.drop_target_register = None
        
        self.create_widgets()
        self.apply_canal_theme()
    
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
            
            self.automation = WMSAutomation(self.canal_actual, log_callback=log_wrapper)
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
        self.status_label.configure(text="● Completado", text_color="#28a745")
        self.update_progress(6)
        
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
    
    def stop_automation(self):
        if self.automation:
            self.automation.stop()
            self.log("⏹️ Deteniendo automatización...", "warning")


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()

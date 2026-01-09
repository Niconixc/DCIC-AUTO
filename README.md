#  DCIC - Sistema de Automatización WMS

Sistema de automatización para el procesamiento de órdenes de trabajo en el WMS (Warehouse Management System) de DCIC.

![Python](https://img.shields.io/badge/Python-3.10+-blue.svg)
![Selenium](https://img.shields.io/badge/Selenium-WebDriver-green.svg)
![CustomTkinter](https://img.shields.io/badge/GUI-CustomTkinter-orange.svg)

##  Descripción

Esta aplicación automatiza el proceso completo de despacho en el sistema WMS, desde la extracción de referencias de manifiestos PDF hasta la creación de Órdenes de Trabajo (OT).

### Canales Soportados

| Canal | Patrón de Referencia | Ubicación WMS |
|-------|---------------------|---------------|
| 🟢 Falabella | `32XXXXXXXX` (10 dígitos) | ZDESP-FALA-01 |
| 🔵 Mercadolibre Flex | `32XXXXXXXX` (10 dígitos) | ZDESP-FLEXMELI-01 |
| 🔵 Mercadolibre Colecta | `32XXXXXXXX` (10 dígitos) | ZDESP-COLECTAMELI-01 |
| 🔵 Mercadolibre Bulky | `32XXXXXXXX` (10 dígitos) | ZDESP-BULKYMELI-01 |
| 🟡 Walmart | 13 dígitos | ZDESP-WALMAT-01 |
| 🔴 Ripley | `243XXXXXXXX-A` | ZDESP-RIPLEY-01 |
| 🟣 Páginas | `Nombre.cl-XXXX` | ZDESP-01-01 |

##  Estructura de Archivos

```
DCIC AUTO/
├── wms_dcic_gui.py      # Aplicación principal (GUI + Automatización)
├── requirements.txt     # Dependencias de Python
├── instalar.bat         # Script de instalación automática
├── WMS_DCIC.bat         # Script para ejecutar la aplicación
├── README.md            # Documentación (este archivo)
└── README.txt           # Documentación legacy
```

##  Funcionalidades

###  Extracción de Referencias
- **Detección automática de canal** basada en el contenido del PDF
- **Soporte para PDFs de texto** usando pdfplumber
- **Soporte para PDFs imagen** usando OCR (Tesseract)
- Extracción de múltiples referencias en un solo paso

###  Automatización WMS
- **Login automático** al sistema WMS
- **Navegación al Monitor de Salida**
- **Búsqueda y selección** de órdenes por referencia
- **Selección de ubicación** con CTRL+F del navegador
- **Verificación de stock** y detección de SKUs sin disponibilidad
- **Creación de OT** con confirmación automática
- **Captura del número de OT** generada

###  Reportes
- Log de ejecución en tiempo real con colores
- Resumen final con estadísticas
- Detalle de órdenes no encontradas
- Listado de SKUs sin stock (banderas rojas)

##  Requisitos

### Sistema
- Windows 10/11
- Google Chrome instalado
- Conexión a Internet

### Software
- **Python 3.10 o superior**
  - Descargar de: https://python.org
  - ⚠️ Marcar "Add Python to PATH" durante la instalación

### Para PDFs con imágenes (Mercadolibre)
- **Tesseract OCR**: `winget install UB-Mannheim.TesseractOCR`
- **Poppler**: Descargar y extraer en `C:\poppler\`

##  Instalación

1. **Clonar o descargar** el repositorio
2. **Ejecutar** `instalar.bat` (doble clic)
3. Esperar a que termine la instalación de dependencias

##  Uso

1. Ejecutar `WMS_DCIC.bat` (doble clic)
2. Seleccionar archivos PDF (manifiestos)
3. El canal se detecta automáticamente
4. Click en **"Extraer Referencias"**
5. Verificar las referencias extraídas
6. Click en **"Ejecutar Automatización"**
7. Esperar a que termine el proceso
8. Revisar el **número de OT** generada
9. Click en **"Nuevo"** para procesar otro PDF

##  Configuración

### Credenciales WMS
Las credenciales están en `wms_dcic_gui.py`, líneas 127-128:
```python
WMS_USER = "18539597"
WMS_PASS = "185395"
```

### Tiempos de Espera
Ajustables en las líneas 130-133:
```python
WAIT_TIMEOUT = 60     # Timeout general (segundos)
DELAY_STEP = 0.8      # Delay entre pasos
DELAY_SEARCH = 1.0    # Delay en búsquedas
DELAY_PAGE = 1.5      # Delay para carga de páginas
```

##  Solución de Problemas

| Error | Solución |
|-------|----------|
| "Python no encontrado" | Instalar Python y marcar "Add Python to PATH" |
| "Chrome no encontrado" | Instalar Google Chrome |
| "No se encontraron referencias" | Instalar Tesseract OCR para PDFs imagen |
| "Ubicación no encontrada" | Verificar que la ubicación exista en WMS |
| "OT no capturada" | La OT se creó pero no se pudo leer el número |

##  Dependencias

```
customtkinter>=5.0.0
pdfplumber>=0.7.0
selenium>=4.0.0
webdriver-manager>=3.8.0
pytesseract>=0.3.10
pdf2image>=1.16.0
Pillow>=9.0.0
```

##  Flujo del Proceso

```
┌─────────────────┐
│  Seleccionar    │
│  PDFs           │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  Detectar       │
│  Canal          │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  Extraer        │
│  Referencias    │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  Automatizar    │
│  WMS            │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  Crear OT       │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  Capturar       │
│  Número OT      │
└─────────────────┘
```

##  Versión

- **Versión:** 2.0
- **Fecha:** Enero 2026
- **Desarrollado para:** DCIC

##  Licencia

Uso interno DCIC - Todos los derechos reservados.

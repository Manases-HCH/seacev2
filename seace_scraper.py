import sys
import logging
import os
import glob
from datetime import datetime
from time import sleep
import re

import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

logging.basicConfig(level=logging.INFO, format='%(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

DOWNLOAD_DIR = '/tmp/seace_downloads'


class SeaceScraperCompleto:

    def __init__(self, headless: bool = True):
        self.headless = headless
        self.driver = None
        self.resultados = []

    def iniciar(self):
        """Inicia el navegador"""
        logger.info("🚀 Iniciando navegador...")

        os.makedirs(DOWNLOAD_DIR, exist_ok=True)

        options = Options()

        # CRITICAL: Opciones obligatorias para Cloud Run
        options.add_argument('--headless=new')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--disable-gpu')
        options.add_argument('--disable-software-rasterizer')
        options.add_argument('--disable-extensions')

        # Optimizaciones
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_argument('--window-size=1920,1080')
        options.add_experimental_option("excludeSwitches", ["enable-automation"])

        # ✅ Carpeta de descarga automática + deshabilitar imágenes
        prefs = {
            "download.default_directory": DOWNLOAD_DIR,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
            "profile.managed_default_content_settings.images": 2,
            "profile.default_content_setting_values.notifications": 2,
        }
        options.add_experimental_option("prefs", prefs)

        try:
            self.driver = webdriver.Chrome(options=options)
            logger.info("✅ Chrome iniciado desde PATH")
        except Exception as e:
            logger.info(f"⚠️ Intentando con ruta explícita: {e}")
            service = Service('/usr/local/bin/chromedriver')
            self.driver = webdriver.Chrome(service=service, options=options)
            logger.info("✅ Chrome iniciado con ruta explícita")

        # ✅ Habilitar descargas en headless (necesario en Chrome moderno)
        self.driver.execute_cdp_cmd(
            "Page.setDownloadBehavior",
            {"behavior": "allow", "downloadPath": DOWNLOAD_DIR}
        )
        self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        logger.info("✅ Navegador iniciado\n")

    def cerrar(self):
        """Cierra el navegador"""
        if self.driver:
            self.driver.quit()

    def click(self, xpath: str, wait_after: float = 0.3):
        """Hace clic usando JavaScript con espera configurable"""
        elem = self.driver.find_element(By.XPATH, xpath)
        self.driver.execute_script("arguments[0].scrollIntoView(true);", elem)
        sleep(0.2)
        self.driver.execute_script("arguments[0].click();", elem)
        sleep(wait_after)

    def escribir(self, xpath: str, texto: str):
        """Escribe en un campo"""
        elem = self.driver.find_element(By.XPATH, xpath)
        self.driver.execute_script("arguments[0].value = '';", elem)
        self.driver.execute_script("arguments[0].value = arguments[1];", elem, texto)
        self.driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", elem)
        sleep(0.2)

    def buscar_y_extraer(self, fecha_inicio: datetime, fecha_fin: datetime) -> str:
        """
        Ejecuta la búsqueda y descarga el Excel con el botón 'Exportar a Excel'.
        Retorna la ruta del archivo descargado, o '' si falla.
        """
        logger.info(f"📅 Rango: {fecha_inicio.strftime('%d/%m/%Y')} → {fecha_fin.strftime('%d/%m/%Y')}")

        # Cargar página
        self.driver.get("https://prod2.seace.gob.pe/seacebus-uiwd-pub/buscadorPublico/buscadorPublico.xhtml")
        logger.info("📄 Página cargada")
        logger.info(f"   Título: {self.driver.title} | URL: {self.driver.current_url}")
        sleep(2)

        # Pestaña
        logger.info("🔖 Seleccionando pestaña 'Buscador de Procedimientos'...")
        try:
            tab_link = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//a[@href="#tbBuscador:tab1"]'))
            )
            self.driver.execute_script("arguments[0].scrollIntoView(true);", tab_link)
            sleep(0.5)
            self.driver.execute_script("arguments[0].click();", tab_link)
            sleep(2)
            logger.info("   ✓ Pestaña seleccionada")
        except TimeoutException:
            logger.error("❌ No se pudo seleccionar la pestaña")
            logger.error(self.driver.page_source[:3000])
            return ''

        # Búsqueda avanzada
        logger.info("🔽 Abriendo búsqueda avanzada...")
        self.click('//fieldset/legend')
        sleep(1)

        # Año
        logger.info(f"📅 Seleccionando año: {fecha_inicio.year}")
        self.click('//*[@id="tbBuscador:idFormBuscarProceso:anioConvocatoria_label"]')
        sleep(0.5)
        self.click(f'//*[@id="tbBuscador:idFormBuscarProceso:anioConvocatoria_panel"]/div/ul/li[@data-label="{fecha_inicio.year}"]')
        sleep(0.5)

        # Fechas
        logger.info("📝 Llenando fechas...")
        self.escribir('//*[@id="tbBuscador:idFormBuscarProceso:dfechaInicio_input"]', fecha_inicio.strftime('%d/%m/%Y'))
        self.escribir('//*[@id="tbBuscador:idFormBuscarProceso:dfechaFin_input"]', fecha_fin.strftime('%d/%m/%Y'))

        # Buscar
        logger.info("🔎 Buscando...")
        self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        sleep(0.5)
        self.click('//*[@id="tbBuscador:idFormBuscarProceso:btnBuscarSelToken"]')
        logger.info("⏳ Esperando resultados...")

        try:
            WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="tbBuscador:idFormBuscarProceso:dtProcesos_data"]'))
            )
            sleep(2)
        except TimeoutException:
            logger.error("❌ Tabla de resultados no apareció")
            return ''

        # Verificar si hay datos
        try:
            msg = self.driver.find_element(By.XPATH, '//td[contains(text(), "No se encontraron")]')
            if msg.is_displayed():
                logger.info("ℹ️  No hay datos para estas fechas")
                return ''
        except NoSuchElementException:
            pass
            
        try:
            filas = self.driver.find_elements(By.XPATH, '//*[@id="tbBuscador:idFormBuscarProceso:dtProcesos_data"]/tr')
            logger.info(f"   📊 Filas visibles en tabla: {len(filas)}")
        except:
            pass
        
        # ✅ Snapshot ANTES del clic para detectar solo archivos nuevos
        archivos_previos = set(glob.glob(os.path.join(DOWNLOAD_DIR, '*.xls*')))
        logger.info(f"   📂 Archivos previos en carpeta: {len(archivos_previos)}")

        # Exportar a Excel
        logger.info("📥 Haciendo clic en 'Exportar a Excel'...")
        try:
            btn_exportar = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, 'tbBuscador:idFormBuscarProceso:btnExportar'))
            )
            self.driver.execute_script("arguments[0].scrollIntoView(true);", btn_exportar)
            sleep(0.5)
            self.driver.execute_script("arguments[0].click();", btn_exportar)
            logger.info("   ✓ Clic en exportar ejecutado")
        except TimeoutException:
            logger.error("❌ No se encontró el botón Exportar")
            return ''

        # Esperar descarga
        archivo = self._esperar_descarga(archivos_previos=archivos_previos, timeout=30)
        return archivo

    def _esperar_descarga(self, archivos_previos: set = None, timeout: int = 30) -> str:
        """Espera a que aparezca un archivo .xlsx NUEVO (ignorando los previos)"""
        logger.info(f"⏳ Esperando descarga en {DOWNLOAD_DIR}...")

        if archivos_previos is None:
            archivos_previos = set()

        for i in range(timeout):
            sleep(1)

            archivos_actuales = set(glob.glob(os.path.join(DOWNLOAD_DIR, '*.xlsx')))
            archivos_actuales |= set(glob.glob(os.path.join(DOWNLOAD_DIR, '*.xls')))

            # Solo archivos nuevos, sin temporales de Chrome
            archivos_nuevos = {
                a for a in (archivos_actuales - archivos_previos)
                if not a.endswith('.crdownload')
            }

            if archivos_nuevos:
                archivo = max(archivos_nuevos, key=os.path.getmtime)
                logger.info(f"✅ Descarga completada: {archivo}")
                return archivo

            if i % 5 == 0:
                logger.info(f"   ... esperando ({i}s)")

        logger.error("❌ Timeout: archivo no descargado")
        return ''

    def renombrar_archivo(self, archivo_original: str, fecha_inicio: datetime) -> str:
        if not archivo_original or not os.path.exists(archivo_original):
            return ''

        nombre_nuevo = os.path.join(
            DOWNLOAD_DIR,
            f"LICIT_PROD2_{fecha_inicio.strftime('%y%m%d')}.xlsx"
        )
    
        # Detectar formato real por cabecera de bytes
        with open(archivo_original, 'rb') as f:
            cabecera = f.read(8)
    
        es_xlsx = cabecera[:4] == b'PK\x03\x04'
        es_xls  = cabecera[:8] == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'
        logger.info(f"   Formato detectado — XLSX={es_xlsx} | XLS={es_xls} | HTML={not es_xlsx and not es_xls}")
    
        try:
            if es_xlsx:
                os.rename(archivo_original, nombre_nuevo)
    
            elif es_xls:
                import xlrd
                from openpyxl import Workbook
                logger.info("🔄 Convirtiendo XLS → XLSX...")
                wb_old = xlrd.open_workbook(archivo_original)
                ws_old = wb_old.sheet_by_index(0)
                wb_new = Workbook()
                ws_new = wb_new.active
                for row in range(ws_old.nrows):
                    ws_new.append(ws_old.row_values(row))
                wb_new.save(nombre_nuevo)
                os.remove(archivo_original)
    
            else:
                # HTML disfrazado — caso más común en SEACE/JSF
                logger.info("🔄 Convirtiendo HTML → XLSX...")
                try:
                    dfs = pd.read_html(archivo_original, encoding='utf-8')
                except Exception:
                    dfs = pd.read_html(archivo_original, encoding='latin-1')
                with pd.ExcelWriter(nombre_nuevo, engine='openpyxl') as writer:
                    for idx, df in enumerate(dfs):
                        df.to_excel(writer, sheet_name=f'Hoja{idx+1}', index=False)
                os.remove(archivo_original)
    
        except Exception as e:
            logger.error(f"❌ Error convirtiendo: {e}")
            if os.path.exists(archivo_original):
                os.rename(archivo_original, nombre_nuevo)
    
        logger.info(f"📄 Renombrado a: {nombre_nuevo}")
        return nombre_nuevo


def pedir_fecha(texto: str) -> datetime:
    """Pide una fecha al usuario"""
    while True:
        try:
            entrada = input(texto).strip()
            for sep in ['/', '-', '.']:
                if sep in entrada:
                    partes = entrada.split(sep)
                    if len(partes) == 3:
                        dia, mes, anio = int(partes[0]), int(partes[1]), int(partes[2])
                        if 1 <= dia <= 31 and 1 <= mes <= 12 and 2000 <= anio <= 2030:
                            return datetime(anio, mes, dia)
            print("❌ Formato: DD/MM/YYYY (ej: 25/12/2025)")
        except ValueError as e:
            print(f"❌ Error: {e}")


def main():
    print("\n" + "=" * 70)
    print("🚀 SEACE SCRAPER COMPLETO - EXPORTAR A EXCEL")
    print("=" * 70)
    print("ℹ️  El navegador se ejecutará en segundo plano (sin ventana)")
    print("=" * 70)

    modo_headless = True

    if '--visible' in sys.argv:
        modo_headless = False
        sys.argv.remove('--visible')
        print("\n⚠️  Modo VISIBLE activado (verás el navegador)")

    if len(sys.argv) >= 3:
        try:
            fecha_inicio = datetime.strptime(sys.argv[1], '%Y-%m-%d')
            fecha_fin = datetime.strptime(sys.argv[2], '%Y-%m-%d')
            print(f"\n📅 Fechas desde argumentos:")
        except ValueError:
            print("\n❌ Error: Formato incorrecto")
            print("   Uso: python seace_scraper.py YYYY-MM-DD YYYY-MM-DD")
            return
    else:
        print("\n📅 Formato: DD/MM/YYYY (ejemplo: 25/01/2026)\n")
        fecha_inicio = pedir_fecha("📅 Fecha inicio: ")
        fecha_fin = pedir_fecha("📅 Fecha fin:    ")

    if fecha_fin < fecha_inicio:
        print("\n❌ La fecha fin debe ser posterior")
        return

    print("\n" + "-" * 70)
    print(f"✓ Inicio: {fecha_inicio.strftime('%d/%m/%Y')}")
    print(f"✓ Fin:    {fecha_fin.strftime('%d/%m/%Y')}")
    print(f"✓ Días:   {(fecha_fin - fecha_inicio).days + 1}")
    print("-" * 70)

    if len(sys.argv) < 3:
        conf = input("\n¿Continuar? (s/n): ").strip().lower()
        if conf not in ['s', 'si', 'sí', 'yes', 'y']:
            print("\n❌ Cancelado")
            return

    print("\n" + "=" * 70)
    print("🚀 INICIANDO EXTRACCIÓN...")
    print("=" * 70 + "\n")

    scraper = SeaceScraperCompleto(headless=modo_headless)

    try:
        scraper.iniciar()
        archivo = scraper.buscar_y_extraer(fecha_inicio, fecha_fin)

        if archivo:
            archivo_final = scraper.renombrar_archivo(archivo, fecha_inicio)
            logger.info("⏳ Esperando antes de cerrar...")
            sleep(3)
            print("\n" + "=" * 70)
            print("✅ ¡EXTRACCIÓN COMPLETADA!")
            print("=" * 70)
            print(f"\n💾 Archivo: {archivo_final}\n")
        else:
            print("\n" + "=" * 70)
            print("⚠️  SIN RESULTADOS")
            print("=" * 70 + "\n")

    except Exception as e:
        print(f"\n❌ ERROR: {e}\n")
        import traceback
        traceback.print_exc()
    finally:
        scraper.cerrar()


if __name__ == "__main__":
    main()

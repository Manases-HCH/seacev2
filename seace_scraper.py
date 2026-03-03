import sys
import logging
import os
import glob
from datetime import datetime
from time import sleep

import pandas as pd
import xlrd
from openpyxl import Workbook
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

logging.basicConfig(level=logging.INFO, format='%(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# ✅ Ruta compatible con Cloud Run (Linux) y Windows local
if sys.platform == 'win32':
    DOWNLOAD_DIR = os.path.join(os.path.expanduser('~'), 'Desktop', 'seace_downloads')
else:
    DOWNLOAD_DIR = '/tmp/seace_downloads'


class SeaceScraperCompleto:

    def __init__(self, headless: bool = True):
        self.headless = headless
        self.driver = None

    def iniciar(self):
        """Inicia el navegador con configuración para Cloud Run"""
        logger.info("🚀 Iniciando navegador...")
        os.makedirs(DOWNLOAD_DIR, exist_ok=True)

        download_dir_abs = os.path.abspath(DOWNLOAD_DIR)
        logger.info(f"📂 Carpeta de descarga: {download_dir_abs}")

        options = Options()

        # ✅ Obligatorio en Cloud Run (contenedor sin display)
        options.add_argument('--headless=new')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--disable-gpu')
        options.add_argument('--disable-software-rasterizer')
        options.add_argument('--disable-extensions')
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_argument('--window-size=1920,1080')
        options.add_argument('--single-process')           # ✅ Mejor en Cloud Run
        options.add_argument('--disable-setuid-sandbox')   # ✅ Requerido en algunos entornos GCP
        options.add_experimental_option("excludeSwitches", ["enable-automation"])

        prefs = {
            "download.default_directory": download_dir_abs,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
            "safebrowsing.disable_download_protection": True,  # ✅ Evita bloqueos de descarga
            "profile.managed_default_content_settings.images": 2,
            "profile.default_content_setting_values.notifications": 2,
        }
        options.add_experimental_option("prefs", prefs)

        try:
            self.driver = webdriver.Chrome(options=options)
            logger.info("✅ Chrome iniciado desde PATH")
        except Exception as e:
            logger.warning(f"⚠️ Intentando con ruta explícita: {e}")
            service = Service('/usr/local/bin/chromedriver')
            self.driver = webdriver.Chrome(service=service, options=options)
            logger.info("✅ Chrome iniciado con ruta explícita")

        # ✅ Habilitar descargas en headless con ruta absoluta
        self.driver.execute_cdp_cmd(
            "Page.setDownloadBehavior",
            {"behavior": "allow", "downloadPath": download_dir_abs}
        )
        self.driver.execute_script(
            "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
        )
        logger.info(f"✅ Navegador listo\n")

    def cerrar(self):
        """Cierra el navegador"""
        if self.driver:
            self.driver.quit()

    def click(self, xpath: str, wait_after: float = 0.3):
        """Clic via JavaScript (evita intercepción de elementos JSF/PrimeFaces)"""
        elem = self.driver.find_element(By.XPATH, xpath)
        self.driver.execute_script("arguments[0].scrollIntoView(true);", elem)
        sleep(0.2)
        self.driver.execute_script("arguments[0].click();", elem)
        sleep(wait_after)

    def escribir(self, xpath: str, texto: str):
        """Escribe en un campo y dispara evento change para JSF"""
        elem = self.driver.find_element(By.XPATH, xpath)
        self.driver.execute_script("arguments[0].value = '';", elem)
        self.driver.execute_script("arguments[0].value = arguments[1];", elem, texto)
        self.driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", elem)
        sleep(0.2)

    def buscar_y_extraer(self, fecha_inicio: datetime, fecha_fin: datetime) -> str:
        """
        Ejecuta la búsqueda y descarga el Excel.
        Retorna la ruta del archivo descargado, o '' si falla.
        """
        logger.info(f"📅 Rango: {fecha_inicio.strftime('%d/%m/%Y')} → {fecha_fin.strftime('%d/%m/%Y')}")

        # Cargar página
        self.driver.get("https://prod2.seace.gob.pe/seacebus-uiwd-pub/buscadorPublico/buscadorPublico.xhtml")
        sleep(3)  # JSF necesita tiempo para inicializar
        logger.info(f"📄 Página cargada: {self.driver.title}")

        # Seleccionar pestaña
        try:
            tab_link = WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.XPATH, '//a[@href="#tbBuscador:tab1"]'))
            )
            self.driver.execute_script("arguments[0].click();", tab_link)
            sleep(2)
            logger.info("✓ Pestaña seleccionada")
        except TimeoutException:
            logger.error("❌ No se pudo seleccionar la pestaña")
            return ''

        # Búsqueda avanzada
        self.click('//fieldset/legend')
        sleep(1)

        # Año
        logger.info(f"📅 Seleccionando año: {fecha_inicio.year}")
        self.click('//*[@id="tbBuscador:idFormBuscarProceso:anioConvocatoria_label"]')
        sleep(0.8)
        self.click(
            f'//*[@id="tbBuscador:idFormBuscarProceso:anioConvocatoria_panel"]'
            f'/div/ul/li[@data-label="{fecha_inicio.year}"]'
        )
        sleep(0.8)

        # Fechas
        logger.info("📝 Llenando fechas...")
        self.escribir(
            '//*[@id="tbBuscador:idFormBuscarProceso:dfechaInicio_input"]',
            fecha_inicio.strftime('%d/%m/%Y')
        )
        sleep(0.3)
        self.escribir(
            '//*[@id="tbBuscador:idFormBuscarProceso:dfechaFin_input"]',
            fecha_fin.strftime('%d/%m/%Y')
        )
        sleep(0.3)

        # Buscar
        logger.info("🔎 Ejecutando búsqueda...")
        self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        sleep(0.5)
        self.click('//*[@id="tbBuscador:idFormBuscarProceso:btnBuscarSelToken"]')

        try:
            WebDriverWait(self.driver, 20).until(
                EC.presence_of_element_located(
                    (By.XPATH, '//*[@id="tbBuscador:idFormBuscarProceso:dtProcesos_data"]')
                )
            )
            sleep(2)
            logger.info("✓ Resultados cargados")
        except TimeoutException:
            logger.error("❌ Tabla de resultados no apareció")
            return ''

        # Sin datos
        try:
            msg = self.driver.find_element(
                By.XPATH, '//td[contains(text(), "No se encontraron")]'
            )
            if msg.is_displayed():
                logger.info("ℹ️  Sin datos para este rango de fechas")
                return ''
        except NoSuchElementException:
            pass

        # Snapshot previo
        archivos_previos = set(glob.glob(os.path.join(DOWNLOAD_DIR, '*')))
        logger.info(f"📂 Archivos previos en carpeta: {len(archivos_previos)}")

        # ✅ Exportar — JS click para no fallar con submit JSF
        logger.info("📥 Exportando a Excel...")
        try:
            btn = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located(
                    (By.ID, 'tbBuscador:idFormBuscarProceso:btnExportar')
                )
            )
            self.driver.execute_script("arguments[0].scrollIntoView(true);", btn)
            sleep(0.5)
            self.driver.execute_script("arguments[0].click();", btn)
            logger.info("   ✓ Clic en exportar ejecutado")
        except TimeoutException:
            logger.error("❌ Botón Exportar no encontrado")
            return ''

        archivo = self._esperar_descarga(archivos_previos=archivos_previos, timeout=45)
        return archivo

    def _esperar_descarga(self, archivos_previos: set = None, timeout: int = 45) -> str:
        """Espera la descarga detectando archivos nuevos, incluyendo .crdownload en progreso"""
        logger.info(f"⏳ Esperando descarga...")

        if archivos_previos is None:
            archivos_previos = set()

        crdownload_detectado = False

        for i in range(timeout):
            sleep(1)

            todos = set(glob.glob(os.path.join(DOWNLOAD_DIR, '*')))
            nuevos = todos - archivos_previos

            # Descarga en progreso
            en_progreso = {a for a in nuevos if a.endswith('.crdownload')}
            if en_progreso and not crdownload_detectado:
                logger.info("   ⬇️  Descarga iniciada...")
                crdownload_detectado = True

            # Archivo terminado
            completos = {
                a for a in nuevos
                if (a.endswith('.xlsx') or a.endswith('.xls'))
                and not a.endswith('.crdownload')
            }

            if completos:
                archivo = max(completos, key=os.path.getmtime)
                logger.info(f"✅ Descarga completada: {archivo}")
                return archivo

            if i % 5 == 0:
                logger.info(f"   ... {i}s | nuevos={len(nuevos)} | en progreso={len(en_progreso)}")

        logger.error("❌ Timeout: archivo no descargado")
        logger.error(f"   Contenido carpeta: {os.listdir(DOWNLOAD_DIR)}")
        return ''

    def renombrar_archivo(self, archivo_original: str, fecha_inicio: datetime) -> str:
        """
        Renombra y convierte el archivo al formato LICIT_PROD2_AAMMDD.xlsx.
        SEACE suele enviar HTML o XLS antiguo con extensión .xlsx → se convierte.
        """
        if not archivo_original or not os.path.exists(archivo_original):
            return ''

        nombre_nuevo = os.path.join(
            DOWNLOAD_DIR,
            f"LICIT_PROD2_{fecha_inicio.strftime('%y%m%d')}.xlsx"
        )

        # Detectar formato real por cabecera de bytes
        with open(archivo_original, 'rb') as f:
            cabecera = f.read(8)

        es_xlsx = cabecera[:4] == b'PK\x03\x04'                          # ZIP → xlsx real
        es_xls  = cabecera[:8] == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'   # OLE2 → xls binario
        # Si no es ninguno → probablemente HTML disfrazado (caso más común en SEACE)

        logger.info(f"   Formato detectado — XLSX={es_xlsx} | XLS={es_xls} | HTML={not es_xlsx and not es_xls}")

        try:
            if es_xlsx:
                # Ya es xlsx válido, solo renombrar
                if archivo_original != nombre_nuevo:
                    os.rename(archivo_original, nombre_nuevo)
                logger.info("✅ Archivo xlsx válido, renombrado directamente")

            elif es_xls:
                # XLS binario antiguo → convertir a xlsx
                logger.info("🔄 Convirtiendo XLS binario → XLSX...")
                wb_old = xlrd.open_workbook(archivo_original)
                ws_old = wb_old.sheet_by_index(0)
                wb_new = Workbook()
                ws_new = wb_new.active
                for row in range(ws_old.nrows):
                    ws_new.append(ws_old.row_values(row))
                wb_new.save(nombre_nuevo)
                os.remove(archivo_original)
                logger.info("✅ Conversión XLS→XLSX completada")

            else:
                # HTML disfrazado de Excel (caso típico de SEACE/JSF)
                logger.info("🔄 Convirtiendo HTML→XLSX con pandas...")
                try:
                    dfs = pd.read_html(archivo_original, encoding='utf-8')
                except Exception:
                    dfs = pd.read_html(archivo_original, encoding='latin-1')

                with pd.ExcelWriter(nombre_nuevo, engine='openpyxl') as writer:
                    for idx, df in enumerate(dfs):
                        sheet = f'Hoja{idx + 1}'
                        df.to_excel(writer, sheet_name=sheet, index=False)
                        logger.info(f"   Hoja '{sheet}': {len(df)} filas")

                os.remove(archivo_original)
                logger.info("✅ Conversión HTML→XLSX completada")

        except Exception as e:
            logger.error(f"❌ Error convirtiendo: {e}")
            # Fallback: renombrar de todas formas
            if os.path.exists(archivo_original):
                os.rename(archivo_original, nombre_nuevo)

        logger.info(f"📄 Archivo final: {nombre_nuevo}")
        return nombre_nuevo


# ──────────────────────────────────────────────
# CLI
# ──────────────────────────────────────────────

def pedir_fecha(texto: str) -> datetime:
    while True:
        try:
            entrada = input(texto).strip()
            for sep in ['/', '-', '.']:
                if sep in entrada:
                    partes = entrada.split(sep)
                    if len(partes) == 3:
                        d, m, a = int(partes[0]), int(partes[1]), int(partes[2])
                        if 1 <= d <= 31 and 1 <= m <= 12 and 2000 <= a <= 2030:
                            return datetime(a, m, d)
            print("❌ Formato: DD/MM/YYYY (ej: 25/01/2026)")
        except ValueError as e:
            print(f"❌ Error: {e}")


def main():
    print("\n" + "=" * 70)
    print("🚀 SEACE SCRAPER - EXPORTAR A EXCEL")
    print(f"   Plataforma: {sys.platform} | Descarga en: {os.path.abspath(DOWNLOAD_DIR)}")
    print("=" * 70)

    modo_headless = True
    if '--visible' in sys.argv:
        modo_headless = False
        sys.argv.remove('--visible')
        print("⚠️  Modo VISIBLE activado (solo para debug local)")

    if len(sys.argv) >= 3:
        try:
            fecha_inicio = datetime.strptime(sys.argv[1], '%Y-%m-%d')
            fecha_fin    = datetime.strptime(sys.argv[2], '%Y-%m-%d')
        except ValueError:
            print("❌ Uso: python sea.py YYYY-MM-DD YYYY-MM-DD")
            sys.exit(1)
    else:
        print("\n📅 Ingresa las fechas (formato DD/MM/YYYY):\n")
        fecha_inicio = pedir_fecha("📅 Fecha inicio: ")
        fecha_fin    = pedir_fecha("📅 Fecha fin:    ")

    if fecha_fin < fecha_inicio:
        print("❌ La fecha fin debe ser posterior a la fecha inicio")
        sys.exit(1)

    print(f"\n✓ Inicio : {fecha_inicio.strftime('%d/%m/%Y')}")
    print(f"✓ Fin    : {fecha_fin.strftime('%d/%m/%Y')}")
    print(f"✓ Días   : {(fecha_fin - fecha_inicio).days + 1}")

    scraper = SeaceScraperCompleto(headless=modo_headless)
    try:
        scraper.iniciar()
        archivo = scraper.buscar_y_extraer(fecha_inicio, fecha_fin)

        if archivo:
            final = scraper.renombrar_archivo(archivo, fecha_inicio)
            sleep(2)
            print("\n" + "=" * 70)
            print("✅ EXTRACCIÓN COMPLETADA")
            print(f"💾 Archivo: {final}")
            print("=" * 70)
            sys.exit(0)
        else:
            print("\n⚠️  SIN RESULTADOS para el rango indicado")
            sys.exit(1)

    except Exception as e:
        print(f"\n❌ ERROR FATAL: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(2)
    finally:
        scraper.cerrar()


if __name__ == "__main__":
    main()

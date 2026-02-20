import sys
import logging
import os
import glob
from datetime import datetime
from time import sleep

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

    def iniciar(self):
        logger.info("🚀 Iniciando navegador...")

        os.makedirs(DOWNLOAD_DIR, exist_ok=True)

        options = Options()
        options.add_argument('--headless=new')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--disable-gpu')
        options.add_argument('--disable-software-rasterizer')
        options.add_argument('--disable-extensions')
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_argument('--window-size=1920,1080')
        options.add_experimental_option("excludeSwitches", ["enable-automation"])

        # ✅ CLAVE: configurar carpeta de descarga automática
        prefs = {
            "download.default_directory": DOWNLOAD_DIR,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
            "profile.default_content_setting_values.notifications": 2,
            "profile.managed_default_content_settings.images": 2,
        }
        options.add_experimental_option("prefs", prefs)

        try:
            self.driver = webdriver.Chrome(options=options)
        except Exception as e:
            logger.info(f"⚠️ Intentando ruta explícita: {e}")
            service = Service('/usr/local/bin/chromedriver')
            self.driver = webdriver.Chrome(service=service, options=options)

        # Habilitar descargas en headless (necesario en Chrome moderno)
        self.driver.execute_cdp_cmd(
            "Page.setDownloadBehavior",
            {"behavior": "allow", "downloadPath": DOWNLOAD_DIR}
        )
        self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        logger.info("✅ Navegador iniciado\n")

    def cerrar(self):
        if self.driver:
            self.driver.quit()

    def click(self, xpath: str, wait_after: float = 0.3):
        elem = self.driver.find_element(By.XPATH, xpath)
        self.driver.execute_script("arguments[0].scrollIntoView(true);", elem)
        sleep(0.2)
        self.driver.execute_script("arguments[0].click();", elem)
        sleep(wait_after)

    def escribir(self, xpath: str, texto: str):
        elem = self.driver.find_element(By.XPATH, xpath)
        self.driver.execute_script("arguments[0].value = '';", elem)
        self.driver.execute_script("arguments[0].value = arguments[1];", elem, texto)
        self.driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", elem)
        sleep(0.2)

    def buscar_y_extraer(self, fecha_inicio: datetime, fecha_fin: datetime) -> str:
        """
        Realiza la búsqueda y hace clic en 'Exportar a Excel'.
        Retorna la ruta del archivo descargado, o '' si falla.
        """
        logger.info(f"📅 Rango: {fecha_inicio.strftime('%d/%m/%Y')} → {fecha_fin.strftime('%d/%m/%Y')}")

        # Cargar página
        self.driver.get("https://prod2.seace.gob.pe/seacebus-uiwd-pub/buscadorPublico/buscadorPublico.xhtml")
        logger.info("📄 Página cargada")
        logger.info(f"   Título: {self.driver.title} | URL: {self.driver.current_url}")
        sleep(3)

        # Pestaña
        logger.info("🔖 Seleccionando pestaña...")
        try:
            tab = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//a[@href="#tbBuscador:tab1"]'))
            )
            self.driver.execute_script("arguments[0].scrollIntoView(true);", tab)
            sleep(0.5)
            self.driver.execute_script("arguments[0].click();", tab)
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
        logger.info(f"📅 Año: {fecha_inicio.year}")
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
                logger.info("ℹ️  Sin datos para este rango")
                return ''
        except NoSuchElementException:
            pass

        # ✅ EXPORTAR A EXCEL
        logger.info("📥 Haciendo clic en 'Exportar a Excel'...")
        try:
            btn_exportar = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.ID, 'tbBuscador:idFormBuscarProceso:btnExportar'))
            )
            self.driver.execute_script("arguments[0].scrollIntoView(true);", btn_exportar)
            sleep(0.5)
            self.driver.execute_script("arguments[0].click();", btn_exportar)
            logger.info("   ✓ Clic en exportar ejecutado")
        except TimeoutException:
            logger.error("❌ No se encontró el botón Exportar")
            return ''

        # Esperar a que el archivo se descargue
        archivo = self._esperar_descarga(timeout=30)
        return archivo

    def _esperar_descarga(self, timeout: int = 30) -> str:
        """Espera a que aparezca un .xlsx en la carpeta de descargas"""
        logger.info(f"⏳ Esperando descarga en {DOWNLOAD_DIR}...")

        for i in range(timeout):
            sleep(1)
            # Buscar archivos xlsx que NO sean temporales (.crdownload)
            archivos = glob.glob(os.path.join(DOWNLOAD_DIR, '*.xlsx'))
            archivos += glob.glob(os.path.join(DOWNLOAD_DIR, '*.xls'))

            if archivos:
                # Tomar el más reciente
                archivo = max(archivos, key=os.path.getmtime)
                logger.info(f"✅ Descarga completada: {archivo}")
                return archivo

            if i % 5 == 0:
                logger.info(f"   ... esperando ({i}s)")

        logger.error("❌ Timeout: archivo no descargado")
        return ''

    def renombrar_archivo(self, archivo_original: str, fecha_inicio: datetime) -> str:
        """Renombra el archivo descargado al formato LICIT_PROD2_AAMMDD.xlsx"""
        if not archivo_original or not os.path.exists(archivo_original):
            return ''

        nombre_nuevo = os.path.join(
            DOWNLOAD_DIR,
            f"LICIT_PROD2_{fecha_inicio.strftime('%y%m%d')}.xlsx"
        )
        os.rename(archivo_original, nombre_nuevo)
        logger.info(f"📄 Renombrado a: {nombre_nuevo}")
        return nombre_nuevo


def main():
    print("\n" + "=" * 70)
    print("🚀 SEACE SCRAPER - EXPORTAR A EXCEL")
    print("=" * 70)

    # Argumentos: python seace_scraper_export.py 2026-02-19 2026-02-19
    if len(sys.argv) >= 3:
        try:
            fecha_inicio = datetime.strptime(sys.argv[1], '%Y-%m-%d')
            fecha_fin = datetime.strptime(sys.argv[2], '%Y-%m-%d')
        except ValueError:
            print("❌ Uso: python seace_scraper_export.py YYYY-MM-DD YYYY-MM-DD")
            return
    else:
        print("\n📅 Formato: DD/MM/YYYY\n")
        while True:
            try:
                fecha_inicio = datetime.strptime(input("Fecha inicio: ").strip(), '%d/%m/%Y')
                break
            except ValueError:
                print("❌ Usa DD/MM/YYYY")
        while True:
            try:
                fecha_fin = datetime.strptime(input("Fecha fin:    ").strip(), '%d/%m/%Y')
                break
            except ValueError:
                print("❌ Usa DD/MM/YYYY")

    if fecha_fin < fecha_inicio:
        print("❌ La fecha fin debe ser posterior")
        return

    scraper = SeaceScraperExport(headless=True)

    try:
        scraper.iniciar()
        archivo = scraper.buscar_y_exportar(fecha_inicio, fecha_fin)

        if archivo:
            archivo_final = scraper.renombrar_archivo(archivo, fecha_inicio)
            print(f"\n✅ Archivo listo: {archivo_final}\n")
        else:
            print("\n⚠️  No se pudo obtener el archivo\n")

    except Exception as e:
        print(f"\n❌ ERROR: {e}\n")
        import traceback
        traceback.print_exc()
    finally:
        scraper.cerrar()


if __name__ == "__main__":
    main()

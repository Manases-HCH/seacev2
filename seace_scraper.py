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
        options.add_argument('--headless=old')
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
        
    def _scrapear_tabla_html(self, fecha_inicio: datetime) -> str:
        """Extrae datos directamente del HTML, cambiando a 20 filas por página"""
        logger.info("🔄 Scrapeando tabla directamente...")
        try:
            from io import StringIO

            # ✅ Cambiar a 20 filas por página (máximo disponible)
            try:
                selector = self.driver.find_element(By.XPATH,
                    '//*[@id="tbBuscador:idFormBuscarProceso:dtProcesos_paginator_bottom"]'
                    '//select[contains(@class,"ui-paginator-rpp-options")]'
                )
                self.driver.execute_script("arguments[0].value = '20';", selector)
                self.driver.execute_script(
                    "arguments[0].dispatchEvent(new Event('change', {bubbles:true}));", selector
                )
                sleep(2)
                WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.XPATH,
                        '//*[@id="tbBuscador:idFormBuscarProceso:dtProcesos_data"]/tr[2]'
                    ))
                )
                sleep(1)
                logger.info("   ✅ Cambiado a 20 filas por página")
            except (NoSuchElementException, TimeoutException):
                logger.info("   ℹ️ No se pudo cambiar paginación")

            # Columnas que nos interesan (sin "Acciones")
            columnas = ['N°', 'Entidad', 'Fecha Publicacion', 'Nomenclatura',
                        'Reiniciado Desde', 'Objeto', 'Descripcion',
                        'Cod SNIP', 'Cod CUI', 'VR/VE', 'Moneda', 'Version SEACE']

            todas_las_filas = []
            pagina = 1

            while True:
                logger.info(f"   📄 Página {pagina}...")
                
                datos_pagina = []
                filas_count = len(self.driver.find_elements(By.XPATH,
                    '//*[@id="tbBuscador:idFormBuscarProceso:dtProcesos_data"]/tr'
                ))
        
                for i in range(filas_count):
                    try:
                        celdas = self.driver.find_elements(By.XPATH,
                            f'//*[@id="tbBuscador:idFormBuscarProceso:dtProcesos_data"]/tr[{i+1}]/td'
                        )
                        if len(celdas) < 12:
                            continue
                        valores = []
                        for c in celdas[:12]:
                            try:
                                valores.append(c.text.strip())
                            except:
                                valores.append('')
                        datos_pagina.append(valores)
                    except Exception as e:
                        logger.warning(f"   ⚠️ Fila {i+1} saltada: {e}")
                        continue

                if datos_pagina:
                    todas_las_filas.extend(datos_pagina)
                    logger.info(f"   ✅ {len(datos_pagina)} filas")

                # Siguiente página
                try:
                    btn_next = self.driver.find_element(By.XPATH,
                        '//*[@id="tbBuscador:idFormBuscarProceso:dtProcesos_paginator_bottom"]'
                        '//span[contains(@class,"ui-icon-seek-next")]'
                        '/parent::span[not(contains(@class,"ui-state-disabled"))]'
                    )
                    self.driver.execute_script("arguments[0].click();", btn_next)
                    WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located((By.XPATH,
                            '//*[@id="tbBuscador:idFormBuscarProceso:dtProcesos_data"]/tr[1]'
                        ))
                    )
                    sleep(1.5)
                    pagina += 1
                except NoSuchElementException:
                    logger.info(f"   ✅ Fin. Total páginas: {pagina}")
                    break

            if not todas_las_filas:
                logger.error("❌ Sin datos")
                return ''

            df = pd.DataFrame(todas_las_filas, columns=columnas)
            logger.info(f"   📊 Total filas: {len(df)}")

            archivo = os.path.join(DOWNLOAD_DIR, f"LICIT_PROD2_{fecha_inicio.strftime('%y%m%d')}.xlsx")
            df.to_excel(archivo, index=False, engine='openpyxl')
            logger.info(f"✅ Guardado: {archivo}")
            return archivo

        except Exception as e:
            logger.error(f"❌ Error scraping tabla: {e}")
            import traceback
            traceback.print_exc()
            return ''
            
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

        logger.info("⏳ Esperando filas reales en tabla...")
        try:
            WebDriverWait(self.driver, 20).until(
                EC.presence_of_element_located((By.XPATH,
                    '//*[@id="tbBuscador:idFormBuscarProceso:dtProcesos_data"]/tr[not(contains(@class,"ui-datatable-empty-message"))]'
                ))
            )
            filas = self.driver.find_elements(By.XPATH,
                '//*[@id="tbBuscador:idFormBuscarProceso:dtProcesos_data"]/tr')
            logger.info(f"   📊 Filas visibles: {len(filas)}")
            sleep(3)  # ✅ dejar que JSF estabilice el ViewState
        except TimeoutException:
            logger.warning("⚠️ No se detectaron filas")

        # ✅ Snapshot ANTES del clic
        archivos_previos = set(glob.glob(os.path.join(DOWNLOAD_DIR, '*.xls*')))
        logger.info(f"   📂 Archivos previos: {len(archivos_previos)}")

        # ✅ Exportar — forzar submit del formulario JSF directamente
        logger.info("📥 Exportando a Excel...")
        try:
            btn_exportar = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, 'tbBuscador:idFormBuscarProceso:btnExportar'))
            )
            self.driver.execute_script("arguments[0].scrollIntoView(true);", btn_exportar)
            sleep(1)

            # ✅ Intentar 3 métodos en orden
            exportado = False

            # Método 1: submit directo del formulario JSF
            try:
                self.driver.execute_script("""
                    var btn = arguments[0];
                    var form = btn.closest('form');
                    if (form) {
                        // Simular click nativo que JSF espera
                        var event = new MouseEvent('click', {bubbles: true, cancelable: true});
                        btn.dispatchEvent(event);
                    }
                """, btn_exportar)
                logger.info("   ✓ Método 1 (MouseEvent nativo)")
                exportado = True
            except Exception as e:
                logger.warning(f"   Método 1 falló: {e}")

            sleep(2)

            # Método 2: click directo de Selenium si método 1 no descargó
            archivos_check = set(glob.glob(os.path.join(DOWNLOAD_DIR, '*'))) - archivos_previos
            if not archivos_check:
                try:
                    btn_exportar.click()
                    logger.info("   ✓ Método 2 (Selenium .click())")
                    exportado = True
                except Exception as e:
                    logger.warning(f"   Método 2 falló: {e}")
                sleep(2)

            # Método 3: PrimeFaces API directa
            archivos_check = set(glob.glob(os.path.join(DOWNLOAD_DIR, '*'))) - archivos_previos
            if not archivos_check:
                try:
                    self.driver.execute_script(
                        "PrimeFaces.ab({s:'tbBuscador:idFormBuscarProceso:btnExportar'});"
                    )
                    logger.info("   ✓ Método 3 (PrimeFaces.ab)")
                    exportado = True
                except Exception as e:
                    logger.warning(f"   Método 3 falló: {e}")
                sleep(2)

        except TimeoutException:
            logger.error("❌ Botón Exportar no encontrado")
            return ''

        # Esperar descarga
        archivo = self._esperar_descarga(archivos_previos=archivos_previos, timeout=30)
        # Si llegó vacío → scrapear tabla directamente
        if not archivo or os.path.getsize(archivo) < 10000:
            if archivo:
                logger.warning(f"⚠️ Archivo vacío ({os.path.getsize(archivo)}b) → scraping tabla")
                os.remove(archivo)
            return self._scrapear_tabla_html(fecha_inicio)

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

        # Conservar extensión original
        ext = os.path.splitext(archivo_original)[1]  # .xls o .xlsx
        nombre_nuevo = os.path.join(
            DOWNLOAD_DIR,
            f"LICIT_PROD2_{fecha_inicio.strftime('%y%m%d')}{ext}"
        )

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
            if os.path.basename(archivo).startswith('LICIT_PROD2_'):
                archivo_final = archivo  # ya tiene nombre correcto (viene del scraping)
            else:
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

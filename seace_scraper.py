import sys
import logging
import os
import glob
from datetime import datetime
from time import sleep

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

        # Carpeta de descarga automática + deshabilitar imágenes
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

        # Habilitar descargas en headless (necesario en Chrome moderno)
        self.driver.execute_cdp_cmd(
            "Page.setDownloadBehavior",
            {"behavior": "allow", "downloadPath": DOWNLOAD_DIR}
        )
        self.driver.execute_script(
            "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
        )
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
        self.driver.execute_script(
            "arguments[0].dispatchEvent(new Event('change', {bubbles:true}));", elem
        )
        self.driver.execute_script(
            "arguments[0].dispatchEvent(new Event('input', {bubbles:true}));", elem
        )
        sleep(0.2)

    # ------------------------------------------------------------------
    # SCRAPING DE TABLA (lógica portada del script que funciona en VS Code)
    # ------------------------------------------------------------------
    def _scrapear_tabla_html(self, fecha_inicio: datetime) -> str:
        """
        Extrae datos directamente del HTML página a página.
        Portado del script local (VS Code) que funciona correctamente.
        """
        logger.info("🔄 Scrapeando tabla directamente...")

        # ── 1. Cambiar a 20 filas/página ──────────────────────────────
        try:
            selector = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH,
                    '//*[@id="tbBuscador:idFormBuscarProceso:dtProcesos_paginator_bottom"]'
                    '//select[contains(@class,"ui-paginator-rpp-options")]'
                ))
            )
            self.driver.execute_script("arguments[0].value = '20';", selector)
            self.driver.execute_script(
                "arguments[0].dispatchEvent(new Event('change', {bubbles:true}));", selector
            )
            # Esperar a que la tabla se recargue con la nueva paginación
            WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.XPATH,
                    '//*[@id="tbBuscador:idFormBuscarProceso:dtProcesos_data"]/tr[1]'
                ))
            )
            sleep(1.5)
            logger.info("   ✅ Cambiado a 20 filas/página")
        except (NoSuchElementException, TimeoutException):
            logger.info("   ℹ️ No se pudo cambiar paginación, continuando con el default")

        columnas = [
            'N°', 'Entidad', 'Fecha Publicacion', 'Nomenclatura',
            'Reiniciado Desde', 'Objeto', 'Descripcion',
            'Cod SNIP', 'Cod CUI', 'VR/VE', 'Moneda', 'Version SEACE',
        ]

        todas_las_filas = []
        pagina = 1
        MAX_PAGINAS = 200  # límite de seguridad

        while pagina <= MAX_PAGINAS:
            logger.info(f"   📄 Página {pagina}...")

            # ── 2. Releer filas frescas en cada página ─────────────────
            try:
                WebDriverWait(self.driver, 15).until(
                    EC.presence_of_element_located((By.XPATH,
                        '//*[@id="tbBuscador:idFormBuscarProceso:dtProcesos_data"]/tr[1]'
                    ))
                )
            except TimeoutException:
                logger.warning(f"   ⚠️ Timeout esperando filas en página {pagina}")
                break

            filas_count = len(self.driver.find_elements(By.XPATH,
                '//*[@id="tbBuscador:idFormBuscarProceso:dtProcesos_data"]/tr'
            ))

            if filas_count == 0:
                logger.warning(f"   ⚠️ Página {pagina} sin filas, terminando")
                break

            datos_pagina = []
            for i in range(filas_count):
                try:
                    # Releer cada fila por índice fresco (evita StaleElementException)
                    celdas = self.driver.find_elements(By.XPATH,
                        f'//*[@id="tbBuscador:idFormBuscarProceso:dtProcesos_data"]/tr[{i+1}]/td'
                    )
                    if len(celdas) < 12:
                        continue
                    valores = []
                    for c in celdas[:12]:
                        try:
                            valores.append(c.text.strip())
                        except Exception:
                            valores.append('')
                    # Ignorar filas de mensaje vacío
                    if any(v for v in valores):
                        datos_pagina.append(valores)
                except Exception as e:
                    logger.warning(f"   ⚠️ Fila {i+1} saltada: {e}")
                    continue

            if datos_pagina:
                todas_las_filas.extend(datos_pagina)
                logger.info(f"   ✅ {len(datos_pagina)} filas extraídas")
            else:
                logger.warning(f"   ⚠️ Página {pagina} sin datos válidos, terminando")
                break

            # ── 3. Siguiente página ────────────────────────────────────
            try:
                btn_next = self.driver.find_element(By.XPATH,
                    '//*[@id="tbBuscador:idFormBuscarProceso:dtProcesos_paginator_bottom"]'
                    '//span[contains(@class,"ui-icon-seek-next")]'
                    '/parent::span[not(contains(@class,"ui-state-disabled"))]'
                )
                self.driver.execute_script("arguments[0].click();", btn_next)

                # Esperar a que la primera fila cambie (tabla recargada)
                WebDriverWait(self.driver, 15).until(
                    EC.presence_of_element_located((By.XPATH,
                        '//*[@id="tbBuscador:idFormBuscarProceso:dtProcesos_data"]/tr[1]'
                    ))
                )
                sleep(1.5)
                pagina += 1

            except NoSuchElementException:
                logger.info(f"   ✅ Última página alcanzada ({pagina})")
                break

        # ── 4. Guardar ─────────────────────────────────────────────────
        if not todas_las_filas:
            logger.error("❌ Sin datos para guardar")
            return ''

        df = pd.DataFrame(todas_las_filas, columns=columnas)
        logger.info(f"   📊 TOTAL: {len(df)} filas extraídas")

        # Muestra rápida
        logger.info("📋 Muestra (primeras 3 filas):")
        for _, row in df.head(3).iterrows():
            logger.info(f"   {row['N°']} | {row['Entidad'][:40]} | {row['Nomenclatura']}")

        archivo = os.path.join(
            DOWNLOAD_DIR,
            f"LICIT_PROD2_{fecha_inicio.strftime('%y%m%d')}.xlsx"
        )
        df.to_excel(archivo, index=False, engine='openpyxl')
        logger.info(f"✅ Guardado: {archivo}")
        return archivo

    # ------------------------------------------------------------------
    # BÚSQUEDA + EXPORT (intenta Excel primero, fallback a scraping)
    # ------------------------------------------------------------------
    def buscar_y_extraer(self, fecha_inicio: datetime, fecha_fin: datetime) -> str:
        logger.info(f"📅 Rango: {fecha_inicio.strftime('%d/%m/%Y')} → {fecha_fin.strftime('%d/%m/%Y')}")

        # Cargar página
        self.driver.get(
            "https://prod2.seace.gob.pe/seacebus-uiwd-pub/buscadorPublico/buscadorPublico.xhtml"
        )
        logger.info("📄 Página cargada")
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
            return ''

        # Búsqueda avanzada
        logger.info("🔽 Abriendo búsqueda avanzada...")
        self.click('//fieldset/legend')
        sleep(1)

        # Año
        logger.info(f"📅 Seleccionando año: {fecha_inicio.year}")
        self.click('//*[@id="tbBuscador:idFormBuscarProceso:anioConvocatoria_label"]')
        sleep(0.5)
        self.click(
            f'//*[@id="tbBuscador:idFormBuscarProceso:anioConvocatoria_panel"]'
            f'/div/ul/li[@data-label="{fecha_inicio.year}"]'
        )
        sleep(0.5)

        # Fechas
        logger.info("📝 Llenando fechas...")
        self.escribir(
            '//*[@id="tbBuscador:idFormBuscarProceso:dfechaInicio_input"]',
            fecha_inicio.strftime('%d/%m/%Y')
        )
        self.escribir(
            '//*[@id="tbBuscador:idFormBuscarProceso:dfechaFin_input"]',
            fecha_fin.strftime('%d/%m/%Y')
        )

        # Buscar
        logger.info("🔎 Buscando...")
        self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        sleep(0.5)
        self.click('//*[@id="tbBuscador:idFormBuscarProceso:btnBuscarSelToken"]')

        logger.info("⏳ Esperando filas reales en tabla...")
        try:
            WebDriverWait(self.driver, 20).until(
                EC.presence_of_element_located((By.XPATH,
                    '//*[@id="tbBuscador:idFormBuscarProceso:dtProcesos_data"]'
                    '/tr[not(contains(@class,"ui-datatable-empty-message"))]'
                ))
            )
            filas = self.driver.find_elements(By.XPATH,
                '//*[@id="tbBuscador:idFormBuscarProceso:dtProcesos_data"]/tr'
            )
            logger.info(f"   📊 Filas visibles: {len(filas)}")
            sleep(3)  # dejar que JSF estabilice el ViewState
        except TimeoutException:
            logger.warning("⚠️ No se detectaron filas — posiblemente sin resultados")
            return ''

        # ── Intentar exportar a Excel ──────────────────────────────────
        archivos_previos = set(glob.glob(os.path.join(DOWNLOAD_DIR, '*.xls*')))
        logger.info(f"   📂 Archivos previos: {len(archivos_previos)}")
        logger.info("📥 Exportando a Excel...")

        try:
            btn_exportar = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located(
                    (By.ID, 'tbBuscador:idFormBuscarProceso:btnExportar')
                )
            )
            self.driver.execute_script("arguments[0].scrollIntoView(true);", btn_exportar)
            sleep(1)

            # Método 1: MouseEvent nativo
            try:
                self.driver.execute_script("""
                    var btn = arguments[0];
                    var event = new MouseEvent('click', {bubbles: true, cancelable: true});
                    btn.dispatchEvent(event);
                """, btn_exportar)
                logger.info("   ✓ Método 1 (MouseEvent nativo)")
            except Exception as e:
                logger.warning(f"   Método 1 falló: {e}")

            sleep(2)

            # Método 2: Selenium .click() si no descargó aún
            if not (set(glob.glob(os.path.join(DOWNLOAD_DIR, '*'))) - archivos_previos):
                try:
                    btn_exportar.click()
                    logger.info("   ✓ Método 2 (Selenium .click())")
                except Exception as e:
                    logger.warning(f"   Método 2 falló: {e}")
                sleep(2)

            # Método 3: PrimeFaces API directa
            if not (set(glob.glob(os.path.join(DOWNLOAD_DIR, '*'))) - archivos_previos):
                try:
                    self.driver.execute_script(
                        "PrimeFaces.ab({s:'tbBuscador:idFormBuscarProceso:btnExportar'});"
                    )
                    logger.info("   ✓ Método 3 (PrimeFaces.ab)")
                except Exception as e:
                    logger.warning(f"   Método 3 falló: {e}")
                sleep(2)

        except TimeoutException:
            logger.warning("⚠️ Botón Exportar no encontrado — pasando directo a scraping")

        # Esperar descarga
        archivo = self._esperar_descarga(archivos_previos=archivos_previos, timeout=30)

        # Si el Excel está vacío o no se descargó → scraping directo
        if not archivo or os.path.getsize(archivo) < 10_000:
            if archivo:
                logger.warning(
                    f"⚠️ Archivo vacío ({os.path.getsize(archivo)}b) → scraping tabla"
                )
                os.remove(archivo)
            else:
                logger.warning("⚠️ Sin descarga → scraping tabla")
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

            # Ignorar temporales de Chrome y archivos previos
            archivos_nuevos = {
                a for a in (archivos_actuales - archivos_previos)
                if not a.endswith('.crdownload')
            }

            if archivos_nuevos:
                # Verificar que no haya un .crdownload activo (descarga en progreso)
                crdownloads = glob.glob(os.path.join(DOWNLOAD_DIR, '*.crdownload'))
                if crdownloads:
                    logger.info(f"   ... descarga en progreso ({i}s)")
                    continue
                archivo = max(archivos_nuevos, key=os.path.getmtime)
                logger.info(f"✅ Descarga completada: {archivo}")
                return archivo

            if i % 5 == 0 and i > 0:
                logger.info(f"   ... esperando ({i}s)")

        logger.warning("⚠️ Timeout: archivo no descargado en tiempo límite")
        return ''

    def renombrar_archivo(self, archivo_original: str, fecha_inicio: datetime) -> str:
        if not archivo_original or not os.path.exists(archivo_original):
            return ''

        ext = os.path.splitext(archivo_original)[1]
        nombre_nuevo = os.path.join(
            DOWNLOAD_DIR,
            f"LICIT_PROD2_{fecha_inicio.strftime('%y%m%d')}{ext}"
        )

        # Evitar sobrescribir si ya existe
        if os.path.exists(nombre_nuevo) and nombre_nuevo != archivo_original:
            ts = datetime.now().strftime('%H%M%S')
            nombre_nuevo = nombre_nuevo.replace(ext, f'_{ts}{ext}')

        os.rename(archivo_original, nombre_nuevo)
        logger.info(f"📄 Renombrado a: {nombre_nuevo}")
        return nombre_nuevo


# ──────────────────────────────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────────────────────────────

def pedir_fecha(texto: str) -> datetime:
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


# ──────────────────────────────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────────────────────────────

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
            fecha_fin    = datetime.strptime(sys.argv[2], '%Y-%m-%d')
            print(f"\n📅 Fechas desde argumentos:")
        except ValueError:
            print("\n❌ Error: Formato incorrecto")
            print("   Uso: python seace_scraper.py YYYY-MM-DD YYYY-MM-DD")
            return
    else:
        print("\n📅 Formato: DD/MM/YYYY (ejemplo: 25/01/2026)\n")
        fecha_inicio = pedir_fecha("📅 Fecha inicio: ")
        fecha_fin    = pedir_fecha("📅 Fecha fin:    ")

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
                archivo_final = archivo
            else:
                archivo_final = scraper.renombrar_archivo(archivo, fecha_inicio)

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

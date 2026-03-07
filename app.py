import os
import logging
import pandas as pd
from time import sleep
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


logging.basicConfig(level=logging.INFO, format="%(levelname)s - %(message)s")
logger = logging.getLogger(__name__)


class SeaceScraperCompleto:

    def __init__(self, headless=True):

        self.headless = headless
        self.driver = None
        self.resultados = []

        # Cloud Run solo permite escritura en /tmp
        self.output_dir = "/tmp"
        os.makedirs(self.output_dir, exist_ok=True)

    # ---------------------------------------------------
    # INICIAR DRIVER
    # ---------------------------------------------------

    def iniciar(self):

        logger.info("🚀 iniciando navegador")

        options = Options()

        if self.headless:
            options.add_argument("--headless=new")

        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        options.add_argument("--window-size=1920,1080")

        self.driver = webdriver.Chrome(options=options)

        logger.info("✅ navegador iniciado")

    # ---------------------------------------------------
    # BUSCAR
    # ---------------------------------------------------

    def buscar(self, fecha_inicio, fecha_fin):

        url = "https://prod2.seace.gob.pe/seacebus-uiwd-pub/buscadorPublico/buscadorPublico.xhtml"

        logger.info("🌐 abriendo SEACE")

        self.driver.get(url)

        WebDriverWait(self.driver, 20).until(
            EC.presence_of_element_located(
                (By.ID, "tbBuscador:idFormBuscarProceso:dfechaInicio_input")
            )
        )

        # insertar fechas
        self.driver.execute_script("""
        document.getElementById('tbBuscador:idFormBuscarProceso:dfechaInicio_input').value = arguments[0];
        document.getElementById('tbBuscador:idFormBuscarProceso:dfechaFin_input').value = arguments[1];
        """,
        fecha_inicio.strftime("%d/%m/%Y"),
        fecha_fin.strftime("%d/%m/%Y"))

        sleep(1)

        # click buscar
        self.driver.find_element(
            By.ID,
            "tbBuscador:idFormBuscarProceso:btnBuscarSelToken"
        ).click()

        logger.info("⏳ esperando resultados")

        WebDriverWait(self.driver, 30).until(
            EC.presence_of_element_located(
                (By.XPATH, '//*[@id="tbBuscador:idFormBuscarProceso:dtProcesos_data"]/tr')
            )
        )

        logger.info("✅ tabla cargada")

    # ---------------------------------------------------
    # EXTRAER FILAS DE UNA PAGINA
    # ---------------------------------------------------

    def extraer_pagina(self):

        filas = self.driver.execute_script("""

        const rows = document.querySelectorAll(
        "#tbBuscador\\\\:idFormBuscarProceso\\\\:dtProcesos_data tr"
        );

        const data = [];

        rows.forEach(r => {

            const cols = r.querySelectorAll("td");

            if(cols.length >= 12){

                data.push([
                    cols[0].innerText.trim(),
                    cols[1].innerText.trim(),
                    cols[2].innerText.trim(),
                    cols[3].innerText.trim(),
                    cols[4].innerText.trim(),
                    cols[5].innerText.trim(),
                    cols[6].innerText.trim(),
                    cols[7].innerText.trim(),
                    cols[8].innerText.trim(),
                    cols[9].innerText.trim(),
                    cols[10].innerText.trim(),
                    cols[11].innerText.trim()
                ])

            }

        })

        return data

        """)

        return filas

    # ---------------------------------------------------
    # SCRAPEAR TODAS LAS PAGINAS
    # ---------------------------------------------------

    def scrapear_tabla(self):

        columnas = [
            'N°',
            'Entidad',
            'Fecha Publicacion',
            'Nomenclatura',
            'Reiniciado Desde',
            'Objeto',
            'Descripcion',
            'Cod SNIP',
            'Cod CUI',
            'VR/VE',
            'Moneda',
            'Version SEACE'
        ]

        data_total = []

        pagina = 1

        while True:

            logger.info(f"📄 página {pagina}")

            data = self.extraer_pagina()

            if not data:
                break

            data_total.extend(data)

            logger.info(f"   filas: {len(data)}")

            try:

                next_btn = self.driver.find_element(
                    By.CSS_SELECTOR,
                    "#tbBuscador\\:idFormBuscarProceso\\:dtProcesos_paginator_bottom .ui-icon-seek-next"
                )

                parent = next_btn.find_element(By.XPATH, "..")

                if "ui-state-disabled" in parent.get_attribute("class"):
                    break

                self.driver.execute_script("arguments[0].click()", parent)

                sleep(1.5)

                pagina += 1

            except:
                break

        df = pd.DataFrame(data_total, columns=columnas)

        logger.info(f"📊 TOTAL FILAS: {len(df)}")

        self.resultados = df

        return df

    # ---------------------------------------------------
    # BUSCAR Y EXTRAER
    # ---------------------------------------------------

    def buscar_y_extraer(self, fecha_inicio, fecha_fin):

        try:

            self.buscar(fecha_inicio, fecha_fin)

            df = self.scrapear_tabla()

            if len(df) == 0:
                logger.warning("⚠️ sin resultados")
                return None

            archivo = os.path.join(
                self.output_dir,
                f"seace_{fecha_inicio.strftime('%Y%m%d')}.xlsx"
            )

            df.to_excel(archivo, index=False)

            logger.info(f"💾 archivo guardado {archivo}")

            return archivo

        except Exception as e:

            logger.error(f"❌ error scraping: {e}")

            return None

    # ---------------------------------------------------
    # RENOMBRAR ARCHIVO
    # ---------------------------------------------------

    def renombrar_archivo(self, archivo, fecha_inicio):

        try:

            nuevo_nombre = os.path.join(
                self.output_dir,
                f"LICIT_PROD2_{fecha_inicio.strftime('%y%m%d')}.xlsx"
            )

            os.rename(archivo, nuevo_nombre)

            logger.info(f"📄 renombrado → {nuevo_nombre}")

            return nuevo_nombre

        except Exception as e:

            logger.warning(f"No se pudo renombrar: {e}")

            return archivo

    # ---------------------------------------------------
    # CERRAR DRIVER
    # ---------------------------------------------------

    def cerrar(self):

        if self.driver:
            self.driver.quit()
            logger.info("🔒 navegador cerrado")

import os
import logging
from datetime import datetime
from time import sleep

import pandas as pd

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

logging.basicConfig(level=logging.INFO, format="%(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

OUTPUT_DIR = "/tmp"
os.makedirs(OUTPUT_DIR, exist_ok=True)


class SeaceScraperCompleto:

    def __init__(self, headless=True):
        self.headless = headless
        self.driver = None
        self.resultados = []

    # -------------------------
    # INICIAR DRIVER
    # -------------------------
    def iniciar(self):

        options = Options()

        if self.headless:
            options.add_argument("--headless=new")

        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        options.add_argument("--window-size=1920,1080")

        logger.info("🚀 Iniciando Chrome")

        self.driver = webdriver.Chrome(options=options)

    # -------------------------
    # BUSCAR
    # -------------------------
    def buscar(self, fecha_inicio, fecha_fin):

        url = "https://prod2.seace.gob.pe/seacebus-uiwd-pub/buscadorPublico/buscadorPublico.xhtml"

        logger.info("🌐 Abriendo SEACE")

        self.driver.get(url)

        WebDriverWait(self.driver, 30).until(
            EC.presence_of_element_located(
                (By.ID, "tbBuscador:idFormBuscarProceso:dfechaInicio_input")
            )
        )

        self.driver.execute_script("""
        document.getElementById('tbBuscador:idFormBuscarProceso:dfechaInicio_input').value = arguments[0];
        document.getElementById('tbBuscador:idFormBuscarProceso:dfechaFin_input').value = arguments[1];
        """,
        fecha_inicio.strftime("%d/%m/%Y"),
        fecha_fin.strftime("%d/%m/%Y")
        )

        sleep(1)

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

    # -------------------------
    # EXTRAER PAGINA
    # -------------------------
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

    # -------------------------
    # SCRAPEAR TABLA
    # -------------------------
    def scrapear_tabla(self):

        columnas = [
            'N°','Entidad','Fecha Publicacion','Nomenclatura',
            'Reiniciado Desde','Objeto','Descripcion',
            'Cod SNIP','Cod CUI','VR/VE','Moneda','Version SEACE'
        ]

        data_total = []
        pagina = 1

        while True:

            logger.info(f"📄 Página {pagina}")

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

    # -------------------------
    # BUSCAR Y EXTRAER
    # -------------------------
    def buscar_y_extraer(self, fecha_inicio, fecha_fin):

        self.buscar(fecha_inicio, fecha_fin)

        df = self.scrapear_tabla()

        if df.empty:
            return None

        archivo = os.path.join(
            OUTPUT_DIR,
            f"seace_{fecha_inicio.strftime('%Y%m%d')}.xlsx"
        )

        df.to_excel(archivo, index=False)

        logger.info(f"💾 archivo guardado {archivo}")

        return archivo

    # -------------------------
    # RENOMBRAR ARCHIVO
    # -------------------------
    def renombrar_archivo(self, archivo, fecha_inicio):

        nuevo = os.path.join(
            OUTPUT_DIR,
            f"LICIT_PROD2_{fecha_inicio.strftime('%y%m%d')}.xlsx"
        )

        os.rename(archivo, nuevo)

        return nuevo

    # -------------------------
    # CERRAR
    # -------------------------
    def cerrar(self):

        if self.driver:
            logger.info("🔒 Cerrando navegador")
            self.driver.quit()


# -------------------------
# MAIN (EJECUCIÓN LOCAL)
# -------------------------
if __name__ == "__main__":

    import sys

    logger.info("🚀 Ejecutando scraper en modo local")

    if len(sys.argv) >= 3:

        fecha_inicio = datetime.strptime(sys.argv[1], "%Y-%m-%d")
        fecha_fin = datetime.strptime(sys.argv[2], "%Y-%m-%d")

    else:

        fecha_inicio = datetime(2026,3,5)
        fecha_fin = datetime(2026,3,5)

    scraper = SeaceScraperCompleto(headless=False)

    try:

        scraper.iniciar()

        archivo = scraper.buscar_y_extraer(fecha_inicio, fecha_fin)

        if archivo:
            logger.info(f"✅ archivo generado: {archivo}")
        else:
            logger.warning("⚠️ sin resultados")

    finally:

        scraper.cerrar()

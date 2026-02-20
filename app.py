from flask import Flask, request, jsonify, send_file
from datetime import datetime
import os
import logging
from seace_scraper import SeaceScraperCompleto

app = Flask(__name__)
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


@app.route('/')
def home():
    return jsonify({
        "status": "online",
        "service": "SEACE Scraper API",
        "endpoints": {
            "/health": "GET - Health check",
            "/scrape": "POST - Ejecutar scraping (params: fecha_inicio, fecha_fin)"
        }
    })


@app.route('/health')
def health():
    return jsonify({"status": "healthy"})


@app.route('/scrape', methods=['POST'])
def scrape():
    scraper = None

    try:
        logger.info("📥 Recibida solicitud de scraping")
        data = request.json

        if not data:
            return jsonify({"error": "No se envió JSON en el body"}), 400

        if 'fecha_inicio' not in data or 'fecha_fin' not in data:
            return jsonify({"error": "Faltan parámetros: fecha_inicio y fecha_fin"}), 400

        try:
            fecha_inicio = datetime.strptime(data['fecha_inicio'], '%Y-%m-%d')
            fecha_fin = datetime.strptime(data['fecha_fin'], '%Y-%m-%d')
        except ValueError:
            return jsonify({"error": "Formato de fecha inválido. Use YYYY-MM-DD"}), 400

        logger.info(f"📅 Fechas: {fecha_inicio.strftime('%Y-%m-%d')} → {fecha_fin.strftime('%Y-%m-%d')}")

        # Ejecutar scraper — ahora devuelve ruta del archivo descargado
        scraper = SeaceScraperCompleto(headless=True)
        scraper.iniciar()
        archivo = scraper.buscar_y_extraer(fecha_inicio, fecha_fin)

        if not archivo:
            logger.warning("⚠️ No se encontraron resultados o fallo la descarga")
            return jsonify({
                "error": "No se encontraron resultados",
                "fecha_inicio": data['fecha_inicio'],
                "fecha_fin": data['fecha_fin']
            }), 404

        # Renombrar al formato estándar
        archivo_final = scraper.renombrar_archivo(archivo, fecha_inicio)
        nombre_archivo = f"LICIT_PROD2_{fecha_inicio.strftime('%y%m%d')}.xlsx"

        logger.info(f"📤 Enviando archivo: {nombre_archivo}")

        return send_file(
            archivo_final or archivo,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=nombre_archivo
        )

    except Exception as e:
        logger.error(f"❌ Error: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500

    finally:
        if scraper:
            try:
                scraper.cerrar()
                logger.info("🔒 Navegador cerrado")
            except:
                pass


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    logger.info(f"🚀 Iniciando servidor en puerto {port}")
    app.run(host='0.0.0.0', port=port)

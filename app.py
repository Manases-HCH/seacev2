from flask import Flask, request, jsonify, send_file
from datetime import datetime
import os
import logging
import tempfile
import pandas as pd
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
    archivo_temporal = None
    
    try:
        logger.info("📥 Recibida solicitud de scraping")
        data = request.json
        
        if not data:
            return jsonify({"error": "No se envió JSON en el body"}), 400
        
        if 'fecha_inicio' not in data or 'fecha_fin' not in data:
            return jsonify({"error": "Faltan parámetros: fecha_inicio y fecha_fin"}), 400
        
        # Parsear fechas
        try:
            fecha_inicio = datetime.strptime(data['fecha_inicio'], '%Y-%m-%d')
            fecha_fin = datetime.strptime(data['fecha_fin'], '%Y-%m-%d')
        except ValueError:
            return jsonify({"error": "Formato de fecha inválido. Use YYYY-MM-DD"}), 400
        
        logger.info(f"📅 Fechas: {fecha_inicio.strftime('%Y-%m-%d')} → {fecha_fin.strftime('%Y-%m-%d')}")
        
        # Crear y ejecutar scraper
        scraper = SeaceScraperCompleto(headless=True)
        scraper.iniciar()
        exito = scraper.buscar_y_extraer(fecha_inicio, fecha_fin)
        
        if not exito or not scraper.resultados:
            logger.warning("⚠️ No se encontraron resultados")
            return jsonify({
                "error": "No se encontraron resultados",
                "fecha_inicio": data['fecha_inicio'],
                "fecha_fin": data['fecha_fin']
            }), 404
        
        # Generar nombre de archivo
        fecha_formato = fecha_inicio.strftime('%y%m%d')  # AAMMDD
        nombre_archivo = f"LICIT_PROD2_{fecha_formato}.xlsx"
        
        logger.info(f"✅ Scraping exitoso: {len(scraper.resultados)} registros")
        logger.info(f"💾 Generando archivo: {nombre_archivo}")
        
        # Crear archivo temporal
        with tempfile.NamedTemporaryFile(mode='wb', delete=False, suffix='.xlsx') as tmp_file:
            archivo_temporal = tmp_file.name
        
        # Guardar Excel en archivo temporal
        df = pd.DataFrame(scraper.resultados)
        
        # Ordenar columnas
        columnas_orden = [
            'N°',
            'Fecha',
            'Entidad Solicitante',
            'Descripción del Requerimiento',
            'Nomenclatura',
            'Objeto',
            'Region',
            'Valor Referencial',
            'Moneda',
            'CUBSO',
            'Fecha de Inicio',
            'Fecha de Fin'
        ]
        
        columnas_existentes = [col for col in columnas_orden if col in df.columns]
        df = df[columnas_existentes]
        
        df.to_excel(archivo_temporal, index=False, engine='openpyxl')
        
        logger.info(f"📤 Enviando archivo: {nombre_archivo}")
        
        # Enviar archivo
        return send_file(
            archivo_temporal,
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
        # Cerrar navegador siempre
        if scraper:
            try:
                scraper.cerrar()
                logger.info("🔒 Navegador cerrado")
            except:
                pass
        
        # Limpiar archivo temporal después de enviarlo
        # (Flask se encarga de esto automáticamente con send_file)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    logger.info(f"🚀 Iniciando servidor en puerto {port}")
    app.run(host='0.0.0.0', port=port)

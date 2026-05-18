"""
main.py - Orquestador principal del RPA de Facturacion.

"""

import os
import sys
import time
from pathlib import Path
from dotenv import load_dotenv
from loguru import logger

# Cargar variables de entorno antes de importar modulos del proyecto
load_dotenv()

# Agregar raiz del proyecto al path para importaciones absolutas
sys.path.insert(0, str(Path(__file__).parent))

from src.utils.logger import configurar_logger
from src.utils.reporte import generar_reporte_resumen
from src.excel.lector import leer_filas_pendientes
from src.excel.escritor import escribir_resultado_fila, escribir_celda
from src.sap.controlador import ControladorSAP
from src.pdf.generador import generar_pdf_factura, construir_nombre_pdf


def main() -> None:
    """Funcion principal que ejecuta el proceso completo de facturacion."""

    # ----------------------------------------------------------------
    # INICIALIZACION
    # ----------------------------------------------------------------
    configurar_logger()
    logger.info("=" * 60)
    logger.info("   INICIO - RPA FACTURACION")
    logger.info("=" * 60)

    # Leer configuracion del proceso
    ruta_excel = os.getenv("INPUT_EXCEL_PATH", r"C:\RPA-FACTURACION\input\facturas.xlsx")
    ruta_pdfs = os.getenv("OUTPUT_PDF_PATH", r"C:\RPA-FACTURACION\output\pdfs")
    fila_inicio = int(os.getenv("FILA_INICIO", "2"))
    max_reintentos = int(os.getenv("REINTENTOS_POR_FILA", "3"))

    logger.info(f"Excel de entrada : {ruta_excel}")
    logger.info(f"Carpeta PDFs     : {ruta_pdfs}")
    logger.info(f"Fila de inicio   : {fila_inicio}")
    logger.info(f"Max. reintentos  : {max_reintentos}")

    # Verificar que el Excel existe
    if not Path(ruta_excel).exists():
        logger.error(f"Archivo Excel no encontrado: {ruta_excel}")
        logger.error("Por favor coloca el archivo en la carpeta 'input/' antes de ejecutar.")
        sys.exit(1)

    # ----------------------------------------------------------------
    # CONEXION SAP
    # ----------------------------------------------------------------
    controlador = ControladorSAP()
    logger.info("Iniciando sesion en SAP...")
    sesion_ok = controlador.iniciar_sesion()

    if not sesion_ok:
        logger.error("No se pudo iniciar sesion en SAP. Abortando proceso.")
        sys.exit(1)

    logger.info("Sesion SAP iniciada correctamente.")

    # ----------------------------------------------------------------
    # PROCESAMIENTO FILA POR FILA
    # ----------------------------------------------------------------
    contadores = {"total": 0, "exitosas": 0, "errores": 0}
    detalles_errores = []

    for datos_fila in leer_filas_pendientes(ruta_excel, fila_inicio):
        numero_fila = datos_fila["numero_fila"]
        # Compatible con Excel simple (orden) y Excel completo (numero_pedido)
        numero_pedido = str(
            datos_fila.get("orden") or datos_fila.get("numero_pedido") or ""
        ).strip()
        contadores["total"] += 1

        logger.info("-" * 60)
        logger.info(f"Procesando fila {numero_fila} | Orden: {numero_pedido}")

        # Marcar como EN PROCESO inmediatamente en el Excel
        escribir_celda(ruta_excel, numero_fila, "estado", "EN PROCESO")

        resultado_sap = {}
        exito = False
        mensaje_error = ""

        # Intentar hasta max_reintentos veces
        for intento in range(1, max_reintentos + 1):
            try:
                logger.info(f"Intento {intento}/{max_reintentos} para fila {numero_fila}")
                resultado_sap = controlador.procesar_factura(datos_fila)
                exito = True
                break

            except Exception as error:
                mensaje_error = str(error)
                logger.warning(
                    f"Fila {numero_fila} | Intento {intento} fallido: {mensaje_error}"
                )
                if intento < max_reintentos:
                    logger.info(f"Reintentando en 3 segundos...")
                    time.sleep(3)

        # ----------------------------------------------------------------
        # RESULTADO EXITOSO
        # ----------------------------------------------------------------
        if exito:
            # Construir resultados segun las columnas que existan en el Excel
            resultados_excel = {"estado": "EXITOSO"}
            if resultado_sap.get("numero_factura_sap"):
                resultados_excel["numero_factura_sap"] = resultado_sap["numero_factura_sap"]
            if resultado_sap.get("fecha_factura"):
                resultados_excel["fecha_factura"] = resultado_sap["fecha_factura"]
            if resultado_sap.get("monto_total"):
                resultados_excel["monto_total"] = resultado_sap["monto_total"]

            # Generar PDF si el Excel tiene columna pdf_generado
            if "pdf_generado" in [str(v).lower() for v in datos_fila.keys()]:
                try:
                    nombre_pdf = construir_nombre_pdf(datos_fila, resultado_sap)
                    ruta_pdf_completa = os.path.join(ruta_pdfs, nombre_pdf)
                    generar_pdf_factura(datos_fila, resultado_sap, ruta_pdf_completa)
                    resultados_excel["pdf_generado"] = nombre_pdf
                    logger.info(f"PDF generado: {nombre_pdf}")
                except Exception as error_pdf:
                    logger.error(f"Error generando PDF para fila {numero_fila}: {error_pdf}")

            escribir_resultado_fila(ruta_excel, numero_fila, resultados_excel)
            contadores["exitosas"] += 1
            logger.info(
                f"Fila {numero_fila} EXITOSA | "
                f"Factura SAP: {resultado_sap.get('numero_factura_sap', 'N/A')}"
            )

        # ----------------------------------------------------------------
        # RESULTADO CON ERROR
        # ----------------------------------------------------------------
        else:
            resultados_error = {"estado": "ERROR"}
            if "mensaje_error" in [str(v).lower() for v in datos_fila.keys()]:
                resultados_error["mensaje_error"] = mensaje_error[:200]
            escribir_resultado_fila(ruta_excel, numero_fila, resultados_error)

            contadores["errores"] += 1
            detalles_errores.append({
                "fila":    numero_fila,
                "pedido":  numero_pedido,
                "error":   mensaje_error,
            })
            logger.error(
                f"Fila {numero_fila} finalizo con ERROR tras {max_reintentos} intentos: "
                f"{mensaje_error}"
            )

        logger.info(
            f"Progreso: {contadores['total']} procesadas | "
            f"{contadores['exitosas']} exitosas | "
            f"{contadores['errores']} errores"
        )

    # ----------------------------------------------------------------
    # CIERRE Y REPORTE FINAL
    # ----------------------------------------------------------------
    logger.info("-" * 60)
    logger.info("Cerrando sesion SAP...")
    controlador.cerrar_sesion()

    logger.info("=" * 60)
    logger.info("   FIN DEL PROCESO - RPA FACTURACION")
    logger.info("=" * 60)
    logger.info(f"Total procesadas : {contadores['total']}")
    logger.info(f"Exitosas         : {contadores['exitosas']}")
    logger.info(f"Con error        : {contadores['errores']}")

    # Generar reporte resumen en archivo .txt
    if contadores["total"] > 0:
        ruta_reporte = generar_reporte_resumen(
            total_filas=contadores["total"],
            exitosas=contadores["exitosas"],
            errores=contadores["errores"],
            detalles_errores=detalles_errores,
        )
        logger.info(f"Reporte guardado en: {ruta_reporte}")
    else:
        logger.info("No habia filas pendientes de procesar.")


if __name__ == "__main__":
    main()

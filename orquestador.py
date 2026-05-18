"""
orquestador.py

Pipeline principal del RPA.
Lee las filas pendientes del Cuadro Maestro y las pasa
por cada transaccion SAP en orden. Escribe el resultado
en la columna ESTADO del Excel al finalizar.
"""

import sys
import time
import importlib.util
import ctypes
import xlwings as xw
from datetime import datetime
from pathlib import Path
from dotenv import load_dotenv
from loguru import logger

load_dotenv()
sys.path.insert(0, str(Path(__file__).parent))

from Transacciones.Loguin   import ejecutar_login
from helpers.cuadro_maestro import leer_sin_estado_ok, RUTA_CUADRO_MAESTRO

# ---------------------------------------------------------------------------
# CONFIGURACION DE LOGS
# ---------------------------------------------------------------------------
LOG_DIR = Path(__file__).parent / "output" / "logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)

logger.remove()
logger.add(
    sys.stdout,
    format="<green>{time:HH:mm:ss}</green> | <level>{level:<8}</level> | {message}",
    level="DEBUG",
)
logger.add(
    LOG_DIR / f"orquestador_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log",
    format="{time:YYYY-MM-DD HH:mm:ss} | {level:<8} | {message}",
    level="DEBUG",
    encoding="utf-8",
)


# ---------------------------------------------------------------------------
# PIPELINE: (nombre_legible, archivo, nombre_funcion)
# El orden aqui define el orden de ejecucion para cada fila.
# ---------------------------------------------------------------------------
PIPELINE = [
    ("ZSD007",     "T_ZSD007",     "ejecutar_ZSD007"),
    ("VA02",       "T_va02",       "ejecutar_VA02"),
    ("VL02N",      "T_vl02n",      "ejecutar_VL02N"),
    ("VF01",       "T_vf01",       "ejecutar_VF01"),
    ("ZEDOCPE001", "T_ZEDOCPE001", "ejecutar_ZEDOCPE001"),
]

# Pausa interactiva después de este paso (None = sin pausa)
# Cambia a "VA02" para probar hasta VA02 y pausar antes de continuar
PAUSAR_DESPUES_DE = "VA02"

COL_ESTADO  = "ESTADO"
COL_COD_PED = "COD PED"


# ---------------------------------------------------------------------------
# UTILIDADES
# ---------------------------------------------------------------------------

def _cargar_transaccion(archivo: str, funcion: str):
    """Carga dinamicamente una transaccion desde Transacciones/T_XXXX.py."""
    ruta = Path(__file__).parent / "Transacciones" / f"{archivo}.py"
    spec = importlib.util.spec_from_file_location(archivo, ruta)
    modulo = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(modulo)
    return getattr(modulo, funcion)


def _msgbox_continuar(paso: str, cod_ped: str) -> bool:
    """
    Muestra un MessageBox de Windows.
    Retorna True si el usuario hace click en Aceptar (continuar).
    Retorna False si hace click en Cancelar (detener esta fila).
    """
    MB_OKCANCEL    = 0x01
    MB_ICONQUESTION = 0x20
    IDOK = 1

    mensaje = (
        f"Paso '{paso}' completado para pedido {cod_ped}.\n\n"
        f"Haz clic en Aceptar para continuar con los siguientes pasos,\n"
        f"o Cancelar para detener esta fila."
    )
    resultado = ctypes.windll.user32.MessageBoxW(0, mensaje, "RPA - Pausa de verificación", MB_OKCANCEL | MB_ICONQUESTION)
    return resultado == IDOK


def _escribir_celda(ruta: str, numero_fila_excel: int, nombre_columna: str, valor) -> None:
    """Escribe un valor en la columna indicada usando xlwings (preserva formato y fórmulas)."""
    app = xw.App(visible=False)
    try:
        wb = app.books.open(ruta)
        ws = wb.sheets["maestro"]

        # Buscar índice de columna en fila de encabezado (fila 5)
        col_idx = None
        for col in range(1, 200):
            celda = ws.cells(5, col).value
            if celda is None:
                break
            if str(celda).strip() == nombre_columna:
                col_idx = col
                break

        if col_idx is None:
            logger.warning(f"Columna '{nombre_columna}' no encontrada en el Excel.")
            wb.close()
            return

        ws.cells(numero_fila_excel, col_idx).value = valor
        wb.save()
        logger.debug(f"Excel actualizado → fila {numero_fila_excel} | {nombre_columna} = {valor}")
    finally:
        app.quit()


def _escribir_estado(ruta: str, numero_fila_excel: int, valor: str) -> None:
    _escribir_celda(ruta, numero_fila_excel, COL_ESTADO, valor)


# ---------------------------------------------------------------------------
# PIPELINE POR FILA
# ---------------------------------------------------------------------------

def ejecutar_pipeline(session, fila: dict, numero_fila_excel: int) -> bool:
    """Ejecuta todas las transacciones del PIPELINE para una fila."""
    cod_ped = str(fila.get(COL_COD_PED, "?"))
    cliente = str(fila.get("CLIENTE", ""))

    logger.info("=" * 55)
    logger.info(f"FILA {numero_fila_excel} | Pedido: {cod_ped} | Cliente: {cliente}")
    logger.info("=" * 55)

    pasos_ok   = []
    pasos_fail = []

    for nombre, archivo, nombre_fn in PIPELINE:
        inicio = time.time()
        logger.info(f"[PASO] Iniciando {nombre}...")

        try:
            funcion   = _cargar_transaccion(archivo, nombre_fn)
            resultado = funcion(session, fila)
            duracion  = round(time.time() - inicio, 1)

            if resultado is False:
                logger.error(f"[PASO] {nombre} retornó False ({duracion}s) — deteniendo fila.")
                pasos_fail.append(nombre)
                _escribir_estado(RUTA_CUADRO_MAESTRO, numero_fila_excel, f"ERROR_{nombre}")
                _log_resumen_fila(cod_ped, pasos_ok, pasos_fail)
                return False

            # Si la transaccion retorna un dict, escribir cada campo en el Excel
            if isinstance(resultado, dict):
                for col_nombre, col_valor in resultado.items():
                    _escribir_celda(RUTA_CUADRO_MAESTRO, numero_fila_excel, col_nombre, col_valor)

            logger.success(f"[PASO] {nombre} completado en {duracion}s.")
            pasos_ok.append(nombre)

            # Pausa interactiva si este paso coincide con PAUSAR_DESPUES_DE
            if PAUSAR_DESPUES_DE and nombre == PAUSAR_DESPUES_DE:
                logger.info(f"[PAUSA] Esperando confirmación del usuario tras {nombre}...")
                continuar = _msgbox_continuar(nombre, cod_ped)
                if not continuar:
                    logger.warning(f"[PAUSA] Usuario canceló. Deteniendo fila {numero_fila_excel}.")
                    _escribir_estado(RUTA_CUADRO_MAESTRO, numero_fila_excel, f"PAUSADO_{nombre}")
                    _log_resumen_fila(cod_ped, pasos_ok, pasos_fail)
                    return False
                logger.info("[PAUSA] Usuario confirmó. Continuando pipeline...")

        except Exception as e:
            duracion = round(time.time() - inicio, 1)
            logger.error(f"[PASO] {nombre} excepción ({duracion}s): {e}")
            pasos_fail.append(nombre)
            _escribir_estado(RUTA_CUADRO_MAESTRO, numero_fila_excel, f"ERROR_{nombre}")
            _log_resumen_fila(cod_ped, pasos_ok, pasos_fail)
            return False

        time.sleep(1)

    _escribir_estado(RUTA_CUADRO_MAESTRO, numero_fila_excel, "OK")
    _log_resumen_fila(cod_ped, pasos_ok, pasos_fail)
    return True


def _log_resumen_fila(cod_ped: str, ok: list, fail: list) -> None:
    logger.info(f"--- Resumen pedido {cod_ped} ---")
    for paso in ok:
        logger.info(f"   ✓ {paso}")
    for paso in fail:
        logger.error(f"   ✗ {paso}")


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

def main():
    logger.info("Iniciando RPA Orquestador")
    logger.info(f"Cuadro Maestro: {RUTA_CUADRO_MAESTRO}")

    pendientes = leer_sin_estado_ok()
    logger.info(f"Filas pendientes encontradas: {len(pendientes)}")

    if pendientes.empty:
        logger.info("No hay filas pendientes. Proceso terminado.")
        return

    logger.info("Iniciando sesion SAP...")
    session = ejecutar_login()
    if not session:
        logger.error("No se pudo iniciar sesion SAP. Abortando.")
        return
    logger.success("Sesion SAP iniciada.")

    exitosas = 0
    errores  = 0

    for df_idx, fila in pendientes.iterrows():
        numero_fila_excel = df_idx + 6
        ok = ejecutar_pipeline(session, fila.to_dict(), numero_fila_excel)
        if ok:
            exitosas += 1
        else:
            errores += 1

    logger.info("=" * 55)
    logger.info("RESUMEN FINAL")
    logger.info(f"  Exitosas : {exitosas}")
    logger.info(f"  Errores  : {errores}")
    logger.info(f"  Total    : {exitosas + errores}")
    logger.info("=" * 55)
    logger.info(f"Log guardado en: {LOG_DIR}")


if __name__ == "__main__":
    main()

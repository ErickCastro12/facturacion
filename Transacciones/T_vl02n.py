import time
from loguru import logger


def _contexto(fila: dict) -> str:
    return (
        f"Pedido={fila.get('COD PED', '?')} | "
        f"Cliente={fila.get('CLIENTE', '?')} | "
        f"BK={fila.get('BUSQUEDA', '?')}"
    )


def ejecutar_VL02N(session, fila: dict):
    ctx = _contexto(fila)
    try:
        logger.info(f"[VL02N] Iniciando | {ctx}")
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nVL02N"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(2)

        logger.debug(f"[VL02N] Transaccion abierta | {ctx}")
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        time.sleep(1)

        logger.success(f"[VL02N] Completado OK | {ctx}")
        return True

    except Exception as e:
        logger.error(f"[VL02N] ✗ Excepción inesperada: {e} | {ctx}")
        return False

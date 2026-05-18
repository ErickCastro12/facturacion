import time
from loguru import logger


def _contexto(fila: dict) -> str:
    return (
        f"Pedido={fila.get('COD PED', '?')} | "
        f"Cliente={fila.get('CLIENTE', '?')} | "
        f"BK={fila.get('BUSQUEDA', '?')}"
    )


def ejecutar_ZEDOCPE001(session, fila: dict):
    ctx = _contexto(fila)
    try:
        logger.info(f"[ZEDOCPE001] Iniciando | {ctx}")
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nZEDOCPE001"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(2)

        logger.debug(f"[ZEDOCPE001] Transaccion abierta | {ctx}")
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        time.sleep(1)

        logger.success(f"[ZEDOCPE001] Completado OK | {ctx}")
        return True

    except Exception as e:
        logger.error(f"[ZEDOCPE001] ✗ Excepción inesperada: {e} | {ctx}")
        return False

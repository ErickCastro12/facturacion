import sys
import time
import pandas as pd
from pathlib import Path
from loguru import logger

sys.path.insert(0, str(Path(__file__).parent.parent))


def _contexto(fila: dict) -> str:
    return (
        f"Pedido={fila.get('COD PED', '?')} | "
        f"Cliente={fila.get('CLIENTE', '?')} | "
        f"BK={fila.get('BUSQUEDA', '?')}"
    )


def ejecutar_VF01(session, fila: dict):
    cod_ped = str(int(float(fila["COD PED"])))
    ctx     = _contexto(fila)

    try:
        logger.info(f"[VF01] Iniciando | {ctx}")
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nVF01"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(2)

        logger.debug(f"[VF01] Transaccion abierta | {ctx}")

        # Número de pedido
        session.findById("/app/con[0]/ses[0]/wnd[0]/usr/tblSAPMV60ATCTRL_ERF_FAKT/ctxtKOMFK-VBELN[0,0]").text = cod_ped
        logger.info(f"[VF01] Pedido ingresado: {cod_ped} | {ctx}")

        # Fecha de factura — columna ETD en formato DD.MM.YYYY
        etd = fila.get("ETD")
        if etd is not None and str(etd).strip() not in ("", "NaT", "nan"):
            try:
                fecha_fkdat = pd.to_datetime(etd).strftime("%d.%m.%Y")
                session.findById("/app/con[0]/ses[0]/wnd[0]/usr/ctxtRV60A-FKDAT").text = fecha_fkdat
                logger.info(f"[VF01] Fecha factura (FKDAT)={fecha_fkdat} (ETD={etd}) | {ctx}")
            except Exception as e:
                logger.error(f"[VF01] ✗ Error convirtiendo ETD: {e} | {ctx}")
                return False
        else:
            logger.warning(f"[VF01] ETD vacía, campo FKDAT no completado | {ctx}")

        session.findById("wnd[0]").sendVKey(0)
        time.sleep(2)

        # Validar que el botón TC_HEAD esté visible y activo antes de hacer click
        try:
            btn_head = session.findById("/app/con[0]/ses[0]/wnd[0]/usr/btnTC_HEAD")
            if not btn_head.changeable and not btn_head.visible:
                logger.error(f"[VF01] ✗ Botón TC_HEAD no está disponible | {ctx}")
                return False
            btn_head.press()
            time.sleep(2)
            logger.info(f"[VF01] Botón TC_HEAD presionado | {ctx}")
        except Exception as e:
            logger.error(f"[VF01] ✗ No se encontró o no se pudo presionar TC_HEAD: {e} | {ctx}")
            return False

        # Pestaña Datos Adicionales
        try:
            tab = session.findById("/app/con[0]/ses[0]/wnd[0]/usr/tabsTABSTRIP_OVERVIEW/tabpKFCU")
            tab.select()
            time.sleep(1)
            logger.info(f"[VF01] Pestaña 'Datos Adicionales' (tabpKFCU) abierta | {ctx}")
        except Exception as e:
            logger.error(f"[VF01] ✗ No se pudo abrir pestaña Datos Adicionales: {e} | {ctx}")
            return False

        # Serie — parte antes del "-" de la columna COMMERCIAL INVOICE
        commercial_invoice = str(fila.get("COMMERCIAL INVOICE", "")).strip()
        if commercial_invoice and commercial_invoice not in ("nan", "NaT"):
            serie = commercial_invoice.split("-")[0].strip()
            try:
                sub_kfcu = "/app/con[0]/ses[0]/wnd[0]/usr/tabsTABSTRIP_OVERVIEW/tabpKFCU/ssubSUBSCREEN_BODY:SAPMV60A:6101/ssubCUSTOMER_SCREEN:ZSD_BILL_DATOS_ADIC:0001"
                partes     = commercial_invoice.split("-", 1)
                correlativo = partes[1].strip() if len(partes) > 1 else ""

                session.findById(f"{sub_kfcu}/ctxtVBRK-Z_SERIES").text = serie
                logger.info(f"[VF01] Z_SERIES='{serie}' (COMMERCIAL INVOICE='{commercial_invoice}') | {ctx}")

                if correlativo:
                    correlativo_4 = correlativo[-4:]
                    session.findById(f"{sub_kfcu}/txtVBRK-Z_CORREL_PREFACT").text = correlativo_4
                    logger.info(f"[VF01] Z_CORREL_PREFACT='{correlativo_4}' (de '{correlativo}') | {ctx}")
                else:
                    logger.warning(f"[VF01] No se encontró correlativo en '{commercial_invoice}' | {ctx}")
            except Exception as e:
                logger.error(f"[VF01] ✗ Error escribiendo Z_SERIES: {e} | {ctx}")
                return False
        else:
            logger.warning(f"[VF01] COMMERCIAL INVOICE vacío, campo Z_SERIES no completado | {ctx}")

        logger.success(f"[VF01] Completado OK | {ctx}")
        return True

    except Exception as e:
        logger.error(f"[VF01] ✗ Excepción inesperada: {e} | {ctx}")
        return False

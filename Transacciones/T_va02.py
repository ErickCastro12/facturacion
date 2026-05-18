import time
import sys
from pathlib import Path
from loguru import logger

sys.path.insert(0, str(Path(__file__).parent.parent))
from helpers.cuadro_maestro import obtener_cod_interlocutor

BASE_PATH    = r"/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,{}]"
PARTNER_PATH = r"/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,{}]"


def _contexto(fila: dict) -> str:
    return (
        f"Pedido={fila.get('COD PED', '?')} | "
        f"Cliente={fila.get('CLIENTE', '?')} | "
        f"BK={fila.get('BUSQUEDA', '?')}"
    )


def encontrar_primera_fila_vacia(session, ctx: str) -> int | None:
    i = 0
    while True:
        try:
            texto = session.findById(BASE_PATH.format(i)).text.strip()
            if texto == "":
                logger.debug(f"[VA02] Fila vacia encontrada en posicion {i} | {ctx}")
                return i
            logger.debug(f"[VA02] Fila {i} ocupada con '{texto}' | {ctx}")
            i += 1
        except Exception as e:
            logger.warning(f"[VA02] Fin de tabla en fila {i}: {e} | {ctx}")
            return None


def asignar_fila(session, idx_fila: int, key_opcion: str, valor_partner: str, ctx: str) -> bool:
    try:
        session.findById(BASE_PATH.format(idx_fila)).key = key_opcion
        logger.debug(f"[VA02] Combo '{key_opcion}' asignado en fila {idx_fila} | {ctx}")
        time.sleep(0.3)

        campo = session.findById(PARTNER_PATH.format(idx_fila))
        campo.setFocus()
        campo.text = valor_partner
        logger.debug(f"[VA02] Partner '{valor_partner}' asignado en fila {idx_fila} | {ctx}")
        return True
    except Exception as e:
        logger.error(f"[VA02] ✗ Error asignando fila {idx_fila}: {e} | {ctx}")
        return False


def ejecutar_VA02(session, fila: dict):
    cod_ped = str(int(float(fila["COD PED"])))
    ctx     = _contexto(fila)

    try:
        logger.info(f"[VA02] Iniciando | {ctx}")
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nVA02"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(2)

        session.findById("/app/con[0]/ses[0]/wnd[0]/usr/ctxtVBAK-VBELN").text = cod_ped
        time.sleep(2)
        session.findById("/app/con[0]/ses[0]/wnd[0]/usr/btnBT_SUCH").press()
        time.sleep(1)

        try:
            ventana = session.findById("/app/con[0]/ses[0]/wnd[1]")
            logger.debug(f"[VA02] Ventana emergente detectada: '{ventana.text}' | {ctx}")
            session.findById("/app/con[0]/ses[0]/wnd[1]/tbar[0]/btn[0]").press()
        except Exception:
            logger.debug(f"[VA02] Sin ventana emergente | {ctx}")
        time.sleep(1)

        valorneto = session.findById("/app/con[0]/ses[0]/wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBAK-NETWR").text
        moneda    = session.findById("/app/con[0]/ses[0]/wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/ctxtVBAK-WAERK").text
        logger.info(f"[VA02] Valor neto={valorneto} | Moneda={moneda} | {ctx}")

        # Validacion monto vs Excel
        monto_excel = fila.get("Monto total CFR")
        if monto_excel is not None and str(valorneto).strip():
            try:
                sap_float   = float(str(valorneto).replace(",", "").strip())
                excel_float = float(monto_excel)
                if sap_float == excel_float:
                    logger.success(f"[VA02] ✓ Monto coincide: SAP={sap_float} | Excel={excel_float} | {ctx}")
                else:
                    diferencia = sap_float - excel_float
                    logger.warning(
                        f"[VA02] ⚠ DIFERENCIA EN MONTO — "
                        f"SAP={sap_float} | Excel={excel_float} | "
                        f"Diferencia={diferencia:+.2f} | {ctx}"
                    )
            except Exception:
                logger.warning(f"[VA02] No se pudo comparar monto | {ctx}")

        session.findById("/app/con[0]/ses[0]/wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press()
        time.sleep(2)
        session.findById(r"/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07").select()
        time.sleep(2)

        opr_logist = str(fila.get("OPR LOGIST", "")).strip()
        try:
            cod_interlocutor = obtener_cod_interlocutor(opr_logist)
            logger.info(f"[VA02] OPR LOGIST={opr_logist} → COD interlocutor={cod_interlocutor} | {ctx}")
        except ValueError as e:
            logger.error(f"[VA02] ✗ {e} | {ctx}")
            return False

        fila_sap = encontrar_primera_fila_vacia(session, ctx)
        if fila_sap is None:
            logger.error(f"[VA02] ✗ No hay filas disponibles en tabla de partners | {ctx}")
            return False

        ok = asignar_fila(session, fila_sap, "ZO", cod_interlocutor, ctx)
        if not ok:
            return False

        logger.success(f"[VA02] Completado OK | {ctx}")
        return {"Monto total CFR": valorneto}

    except Exception as e:
        logger.error(f"[VA02] ✗ Excepción inesperada: {e} | {ctx}")
        return False

import time
import pandas as pd
from loguru import logger


def _sap_a_float(valor_sap: str) -> float:
    """Convierte formato SAP '4,080.000' a float 4080.0"""
    return float(str(valor_sap).replace(",", "").strip())


def _contexto(fila: dict) -> str:
    """Retorna string de contexto para logs: pedido | cliente | embarque."""
    return (
        f"Pedido={fila.get('COD PED', '?')} | "
        f"Cliente={fila.get('CLIENTE', '?')} | "
        f"BK={fila.get('BUSQUEDA', '?')}"
    )


def ejecutar_ZSD007(session, fila: dict):
    cod_ped = str(int(float(fila["COD PED"])))
    ctx     = _contexto(fila)

    try:
        logger.info(f"[ZSD007] Iniciando | {ctx}")
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nZSD007"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(2)

        session.findById("wnd[0]/usr/txt%_P_VSTEL_%_APP_%-TEXT")
        logger.debug(f"[ZSD007] Transaccion abierta | {ctx}")

        session.findById("/app/con[0]/ses[0]/wnd[0]/usr/ctxtP_SONUM").text = cod_ped
        time.sleep(2)
        session.findById("/app/con[0]/ses[0]/wnd[0]/tbar[1]/btn[8]").press()
        time.sleep(2)
        session.findById("/app/con[0]/ses[0]/wnd[0]/tbar[1]/btn[37]")
        session.findById("/app/con[0]/ses[0]/wnd[0]/tbar[1]/btn[37]").press()
        time.sleep(5)
        session.findById("/app/con[0]/ses[0]/wnd[0]/tbar[1]/btn[38]").press()
        time.sleep(5)

        vpanelgrabado = session.findById("/app/con[0]/ses[0]/wnd[0]/sbar/pane[0]").text
        logger.debug(f"[ZSD007] Barra de estado SAP: '{vpanelgrabado}' | {ctx}")

        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        time.sleep(5)

        session.findById("/app/con[0]/ses[0]/wnd[0]/usr/ctxtP_SONUM").text = cod_ped
        time.sleep(2)
        session.findById("/app/con[0]/ses[0]/wnd[0]/tbar[1]/btn[8]").press()
        time.sleep(2)

        vpanelgrilla = session.findById("/app/con[0]/ses[0]/wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell")

        col_list = list(vpanelgrilla.columnOrder)
        logger.debug(f"[ZSD007] Columnas grilla: {col_list} | {ctx}")

        datos = []
        for i in range(vpanelgrilla.rowCount):
            registro = {col: vpanelgrilla.getCellValue(i, col) for col in col_list}
            datos.append(registro)

        df = pd.DataFrame(datos)
        logger.info(f"[ZSD007] Filas obtenidas de SAP: {len(df)} | {ctx}")

        # ----------------------------------------------------------------
        # TOTALES (ultima fila de la grilla)
        # ----------------------------------------------------------------
        ultima        = vpanelgrilla.rowCount - 1
        total_kwmeng  = vpanelgrilla.getCellValue(ultima, "KWMENG")
        total_vemng   = vpanelgrilla.getCellValue(ultima, "VEMNG")
        logger.info(f"[ZSD007] Total KWMENG={total_kwmeng} | Total VEMNG={total_vemng} | {ctx}")

        # ----------------------------------------------------------------
        # VALIDACION CAJAS: VEMNG SAP vs "Número de cajas" Excel
        # ----------------------------------------------------------------
        cajas_excel = fila.get("Número de cajas")
        if cajas_excel is not None and str(total_vemng).strip():
            vemng_float = _sap_a_float(total_vemng)
            cajas_float = float(cajas_excel)

            if vemng_float == cajas_float:
                logger.success(
                    f"[ZSD007] ✓ Cajas coinciden: SAP={vemng_float} | Excel={cajas_float} | {ctx}"
                )
            else:
                diferencia = vemng_float - cajas_float
                logger.error(
                    f"[ZSD007] ✗ DIFERENCIA EN CAJAS — "
                    f"SAP={vemng_float} | Excel={cajas_float} | "
                    f"Diferencia={diferencia:+.0f} | {ctx}"
                )
                return False
        else:
            logger.warning(f"[ZSD007] Sin dato 'Número de cajas' en Excel para comparar | {ctx}")

        # ----------------------------------------------------------------
        # VALIDACION MONTOS: KWMENG SAP vs "Monto total CFR" Excel
        # ----------------------------------------------------------------
        monto_excel = fila.get("Monto total CFR")
        if monto_excel is not None and str(total_kwmeng).strip():
            try:
                kwmeng_float = _sap_a_float(total_kwmeng)
                monto_float  = float(monto_excel)

                if kwmeng_float == monto_float:
                    logger.success(
                        f"[ZSD007] ✓ Monto CFR coincide: SAP={kwmeng_float} | Excel={monto_float} | {ctx}"
                    )
                else:
                    diferencia = kwmeng_float - monto_float
                    logger.warning(
                        f"[ZSD007] ⚠ DIFERENCIA EN MONTO CFR — "
                        f"SAP={kwmeng_float} | Excel={monto_float} | "
                        f"Diferencia={diferencia:+.2f} | {ctx}"
                    )
            except Exception:
                logger.warning(f"[ZSD007] No se pudo comparar monto CFR | {ctx}")
        else:
            logger.warning(f"[ZSD007] Sin dato 'Monto total CFR' en Excel para comparar | {ctx}")

        # ----------------------------------------------------------------
        # VALIDAR STATUS VERDE EN TODAS LAS FILAS (excepto totales)
        # ----------------------------------------------------------------
        ICONO_VERDE     = "@08\QSemáf.verde: Go; correcto@"
        filas_no_verdes = []

        for i in range(ultima):
            icono = vpanelgrilla.getCellValue(i, "ICON")
            if icono != ICONO_VERDE:
                filas_no_verdes.append({
                    "fila":  i,
                    "SONUM": vpanelgrilla.getCellValue(i, "SONUM"),
                    "MATNR": vpanelgrilla.getCellValue(i, "MATNR"),
                    "ICON":  icono,
                })

        if not filas_no_verdes:
            logger.success(f"[ZSD007] ✓ Todas las filas tienen status verde | {ctx}")
        else:
            for f in filas_no_verdes:
                logger.error(
                    f"[ZSD007] ✗ Fila sin status verde — "
                    f"Fila={f['fila']} | SONUM={f['SONUM']} | MATNR={f['MATNR']} | "
                    f"ICON={f['ICON']} | {ctx}"
                )
            return False

        logger.success(f"[ZSD007] Completado OK | {ctx}")
        return df

    except Exception as e:
        logger.error(f"[ZSD007] Excepción inesperada: {e} | {ctx}")
        return False

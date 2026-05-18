import time
import sys
import pandas as pd
from pathlib import Path
from loguru import logger

sys.path.insert(0, str(Path(__file__).parent.parent))
from helpers.cuadro_maestro import obtener_cod_interlocutor
from helpers.codigos_sap    import obtener_puerto_embarque

BASE_PATH_TPL    = "/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/{tab}/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,{fila}]"
PARTNER_PATH_TPL = "/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/{tab}/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,{fila}]"

# Nombres del Excel → código ISO SAP de 2 letras
# Cubre: alpha-3, nombres en español, nombres en inglés y alias comunes
_MAPA_PAIS = {
    # ── ISO alpha-3 ────────────────────────────────────────────────────────
    "USA": "US", "PER": "PE", "COL": "CO", "CHL": "CL", "ECU": "EC",
    "MEX": "MX", "CAN": "CA", "BRA": "BR", "ESP": "ES", "CHN": "CN",
    "JPN": "JP", "DEU": "DE", "NLD": "NL", "GBR": "GB", "FRA": "FR",
    "ITA": "IT", "AUS": "AU", "NZL": "NZ", "ARG": "AR", "BOL": "BO",
    "PRY": "PY", "URY": "UY", "VEN": "VE", "PAN": "PA", "CRI": "CR",
    "GTM": "GT", "HND": "HN", "SLV": "SV", "NIC": "NI", "DOM": "DO",
    "CUB": "CU", "BEL": "BE", "CHE": "CH", "AUT": "AT", "PRT": "PT",
    "SWE": "SE", "NOR": "NO", "DNK": "DK", "FIN": "FI", "GRC": "GR",
    "TUR": "TR", "RUS": "RU", "IND": "IN", "KOR": "KR", "THA": "TH",
    "VNM": "VN", "IDN": "ID", "MYS": "MY", "PHL": "PH", "SGP": "SG",
    "ZAF": "ZA", "NGA": "NG", "EGY": "EG", "MAR": "MA", "KEN": "KE",
    "ISR": "IL", "SAU": "SA", "ARE": "AE", "QAT": "QA", "KWT": "KW",
    "HKG": "HK", "TWN": "TW", "CZE": "CZ", "POL": "PL", "HUN": "HU",
    "ROU": "RO", "HRV": "HR", "SVK": "SK", "SVN": "SI", "BGR": "BG",
    # ── Nombres en español ────────────────────────────────────────────────
    "PERU": "PE", "PERÚ": "PE",
    "ESTADOS UNIDOS": "US", "EE.UU.": "US",
    "MEXICO": "MX", "MÉXICO": "MX",
    "BRASIL": "BR",
    "ALEMANIA": "DE",
    "ESPAÑA": "ES",
    "JAPÓN": "JP", "JAPON": "JP",
    "PAISES BAJOS": "NL", "PAÍSES BAJOS": "NL", "HOLANDA": "NL",
    "REINO UNIDO": "GB",
    "FRANCIA": "FR",
    "ITALIA": "IT",
    "CANADA": "CA", "CANADÁ": "CA",
    "NUEVA ZELANDA": "NZ",
    "ARGENTINA": "AR",
    "COLOMBIA": "CO",
    "CHILE": "CL",
    "ECUADOR": "EC",
    "BOLIVIA": "BO",
    "PARAGUAY": "PY",
    "URUGUAY": "UY",
    "VENEZUELA": "VE",
    "PANAMA": "PA", "PANAMÁ": "PA",
    "COSTA RICA": "CR",
    "GUATEMALA": "GT",
    "HONDURAS": "HN",
    "EL SALVADOR": "SV",
    "NICARAGUA": "NI",
    "REPUBLICA DOMINICANA": "DO", "REPÚBLICA DOMINICANA": "DO",
    "CUBA": "CU",
    "BELGICA": "BE", "BÉLGICA": "BE",
    "SUIZA": "CH",
    "AUSTRIA": "AT",
    "PORTUGAL": "PT",
    "SUECIA": "SE",
    "NORUEGA": "NO",
    "DINAMARCA": "DK",
    "FINLANDIA": "FI",
    "GRECIA": "GR",
    "TURQUIA": "TR", "TURQUÍA": "TR",
    "RUSIA": "RU",
    "INDIA": "IN",
    "COREA DEL SUR": "KR",
    "TAILANDIA": "TH",
    "VIETNAM": "VN",
    "INDONESIA": "ID",
    "MALASIA": "MY",
    "FILIPINAS": "PH",
    "SINGAPUR": "SG",
    "SUDAFRICA": "ZA", "SUDÁFRICA": "ZA",
    "MARRUECOS": "MA",
    "ISRAEL": "IL",
    "ARABIA SAUDI": "SA", "ARABIA SAUDÍ": "SA",
    "EMIRATOS": "AE", "EMIRATOS ARABES UNIDOS": "AE",
    "QATAR": "QA",
    "KUWAIT": "KW",
    "HONG KONG": "HK",
    "TAIWAN": "TW",
    "REPUBLICA CHECA": "CZ", "REPÚBLICA CHECA": "CZ",
    "POLONIA": "PL",
    "HUNGRIA": "HU", "HUNGRÍA": "HU",
    "RUMANIA": "RO", "RUMANÍA": "RO",
    "CROACIA": "HR",
    "ESLOVAQUIA": "SK",
    "ESLOVENIA": "SI",
    "BULGARIA": "BG",
    # ── Nombres en inglés ─────────────────────────────────────────────────
    "UNITED STATES": "US", "UNITED STATES OF AMERICA": "US",
    "BRAZIL": "BR",
    "GERMANY": "DE",
    "SPAIN": "ES",
    "JAPAN": "JP",
    "NETHERLANDS": "NL",
    "UK": "GB", "UNITED KINGDOM": "GB",
    "FRANCE": "FR",
    "ITALY": "IT",
    "AUSTRALIA": "AU",
    "NEW ZEALAND": "NZ",
    "BELGIUM": "BE",
    "SWITZERLAND": "CH",
    "SWEDEN": "SE",
    "NORWAY": "NO",
    "DENMARK": "DK",
    "FINLAND": "FI",
    "GREECE": "GR",
    "TURKEY": "TR",
    "RUSSIA": "RU",
    "SOUTH KOREA": "KR",
    "THAILAND": "TH",
    "MALAYSIA": "MY",
    "PHILIPPINES": "PH",
    "SINGAPORE": "SG",
    "SOUTH AFRICA": "ZA",
    "MOROCCO": "MA",
    "SAUDI ARABIA": "SA",
    "UAE": "AE",
    "CHINA": "CN",
    "POLAND": "PL",
    "ROMANIA": "RO",
    "CROATIA": "HR",
    "CZECH REPUBLIC": "CZ",
    # ── Alias adicionales detectados en el Cuadro Maestro ─────────────────
    "CHIPRE": "CY", "CYPRUS": "CY",
    "EAU": "AE",                        # Emiratos Arabes Unidos (abrev. francesa)
    "EMIRATOS ARABES UNIDOS": "AE",
    "EGIPTO": "EG", "EGYPT": "EG",
    "IRAQ": "IQ", "IRAK": "IQ",
    "LIBANO": "LB", "LÍBANO": "LB", "LEBANON": "LB",
    "SINGAPOUR": "SG",                  # grafía francesa usada en el Excel
    "ESPANA": "ES",                     # alias sin ñ por encoding
}


def _contexto(fila: dict) -> str:
    return (
        f"Pedido={fila.get('COD PED', '?')} | "
        f"Cliente={fila.get('CLIENTE', '?')} | "
        f"BK={fila.get('BUSQUEDA', '?')}"
    )


def _obtener_codigo_puerto_f4(session, campo_path: str, discharge_port: str, ctx: str) -> str | None:
    """
    Abre el F4 del campo ZZ_PUEDES, lee la grilla de puertos filtrada
    por el país ya ingresado en SKTO, y retorna el código que coincide
    con el DISCHARGE PORT del Excel.
    """
    try:
        session.findById(campo_path).setFocus()
        session.findById("wnd[0]").sendVKey(4)  # F4
        time.sleep(2)

        # Leer grilla del popup
        grilla = session.findById("wnd[1]/usr/cntlALV_CONTAINER/shellcont/shell")
        total  = grilla.rowCount
        logger.debug(f"[VA02] F4 puertos: {total} registros encontrados | {ctx}")

        busqueda = discharge_port.strip().upper()

        for i in range(total):
            denominacion = str(grilla.getCellValue(i, "BEZTL")).strip().upper()
            codigo       = str(grilla.getCellValue(i, "PTOAD")).strip()
            if busqueda in denominacion or denominacion in busqueda:
                logger.info(f"[VA02] Puerto encontrado: '{denominacion}' → {codigo} | {ctx}")
                # Seleccionar fila y cerrar popup
                grilla.selectedRows = str(i)
                session.findById("wnd[1]").sendVKey(2)  # Enter para confirmar
                time.sleep(1)
                return codigo

        logger.error(f"[VA02] ✗ '{discharge_port}' no encontrado en la grilla F4 | {ctx}")
        session.findById("wnd[1]").sendVKey(12)  # F12 para cerrar sin seleccionar
        time.sleep(1)
        return None

    except Exception as e:
        logger.error(f"[VA02] ✗ Error leyendo F4 de puertos: {e} | {ctx}")
        try:
            session.findById("wnd[1]").sendVKey(12)
        except Exception:
            pass
        return None


TABSTRIP_PATH = "/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD"


def _buscar_tab(session, nombre: str, ctx: str) -> str | None:
    """Busca una pestaña por nombre y retorna su ID completo (tabpXXX) para usar con findById."""
    tabstrip = session.findById(TABSTRIP_PATH)
    for i in range(tabstrip.Children.Count):
        tab = tabstrip.Children.Item(i)
        if nombre.lower() in str(tab.Text).lower():
            tab_id = f"tabp{tab.Name}"   # tab.Name = "T\07" → necesitamos "tabpT\07"
            logger.debug(f"[VA02] Pestaña '{tab.Text}' encontrada en posición {i} (ID={tab_id}) | {ctx}")
            return tab_id
    logger.error(f"[VA02] ✗ Pestaña '{nombre}' no encontrada | {ctx}")
    return None


def interlocutor_ya_existe(session, tab_id: str, ctx: str) -> bool:
    """Verifica si 'Operador Logístico' ya está asignado en la tabla de partners."""
    i = 0
    while True:
        try:
            path = BASE_PATH_TPL.format(tab=tab_id, fila=i)
            tipo = session.findById(path).text.strip()
            if tipo == "":
                return False
            if "Operador" in tipo:
                logger.info(f"[VA02] '{tipo}' ya existe en fila {i}, se omite asignación | {ctx}")
                return True
            i += 1
        except Exception:
            return False


def encontrar_primera_fila_vacia(session, tab_id: str, ctx: str) -> int | None:
    i = 0
    while True:
        try:
            path  = BASE_PATH_TPL.format(tab=tab_id, fila=i)
            texto = session.findById(path).text.strip()
            if texto == "":
                logger.debug(f"[VA02] Fila vacia encontrada en posicion {i} | {ctx}")
                return i
            logger.debug(f"[VA02] Fila {i} ocupada con '{texto}' | {ctx}")
            i += 1
        except Exception as e:
            logger.warning(f"[VA02] Fin de tabla en fila {i}: {e} | {ctx}")
            return None


def asignar_fila(session, tab_id: str, idx_fila: int, key_opcion: str, valor_partner: str, ctx: str) -> bool:
    try:
        session.findById(BASE_PATH_TPL.format(tab=tab_id, fila=idx_fila)).key = key_opcion
        logger.debug(f"[VA02] Combo '{key_opcion}' asignado en fila {idx_fila} | {ctx}")
        time.sleep(0.3)

        campo = session.findById(PARTNER_PATH_TPL.format(tab=tab_id, fila=idx_fila))
        campo.setFocus()
        campo.text = valor_partner
        session.findById("wnd[0]").sendVKey(0)
        logger.debug(f"[VA02] Partner '{valor_partner}' asignado + Enter en fila {idx_fila} | {ctx}")
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

        # Pestaña Interlocutores (busqueda dinamica por nombre)
        tab_id = _buscar_tab(session, "Interlocutor", ctx)
        if tab_id is None:
            return False
        session.findById(f"{TABSTRIP_PATH}/{tab_id}").select()
        time.sleep(2)
        logger.info(f"[VA02] Pestaña Interlocutor abierta (ID={tab_id}) | {ctx}")

        # Codigo interlocutor segun OPR LOGIST
        opr_logist = str(fila.get("OPR LOGIST", "")).strip()
        try:
            cod_interlocutor = obtener_cod_interlocutor(opr_logist)
            logger.info(f"[VA02] OPR LOGIST={opr_logist} → COD interlocutor={cod_interlocutor} | {ctx}")
        except ValueError as e:
            logger.error(f"[VA02] ✗ {e} | {ctx}")
            return False

        if interlocutor_ya_existe(session, tab_id, ctx):
            logger.info(f"[VA02] Interlocutor ya registrado, no se requiere asignación | {ctx}")
        else:
            fila_sap = encontrar_primera_fila_vacia(session, tab_id, ctx)
            if fila_sap is None:
                logger.error(f"[VA02] ✗ No hay filas disponibles en tabla de partners | {ctx}")
                return False
            ok = asignar_fila(session, tab_id, fila_sap, "ZO", cod_interlocutor, ctx)
            if not ok:
                return False

        # Pestaña Datos Exportacion (busqueda dinamica por nombre)
        tab_exp_id = _buscar_tab(session, "Datos Exportación", ctx)
        if tab_exp_id is None:
            return False
        session.findById(f"{TABSTRIP_PATH}/{tab_exp_id}").select()
        time.sleep(2)
        logger.info(f"[VA02] Pestaña Datos Exportación abierta (ID={tab_exp_id}) | {ctx}")

        # Campos de exportacion
        pais_destino    = str(fila.get("PAIS", "")).strip()
        discharge_port  = str(fila.get("DISCHARGE PORT", "")).strip()
        loading_port    = str(fila.get("LOADING PORT", "")).strip()

        sub = f"/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/{tab_exp_id}/ssubSUBSCREEN_BODY:SAPMV45A:4323/subCUSTOMER_SCREEN:ZSD_DMA_DATOS_ADICIONALES:1100/tabsZZ_TABSTRIP/tabpTBS_1/ssubTBS_1_REF1:ZSD_DMA_DATOS_ADICIONALES:1110"

        def _limpiar(path):
            try:
                session.findById(path).text = ""
            except Exception:
                pass

        # SKFROM — país de origen, siempre Perú
        _limpiar(f"{sub}/ctxtVBAK-SKFROM")
        session.findById(f"{sub}/ctxtVBAK-SKFROM").text = "PE"
        logger.debug(f"[VA02] SKFROM=PE (origen fijo Perú) | {ctx}")

        # SKTO — país de destino desde columna PAIS, normalizado a código ISO 2 letras
        pais_upper  = pais_destino.upper()
        if len(pais_upper) == 2:
            codigo_pais = pais_upper  # ya es código ISO
        else:
            codigo_pais = _MAPA_PAIS.get(pais_upper)
            if codigo_pais is None:
                logger.error(f"[VA02] ✗ País '{pais_destino}' no encontrado en _MAPA_PAIS — agregar al diccionario | {ctx}")
                return False
            logger.debug(f"[VA02] PAIS '{pais_destino}' normalizado → '{codigo_pais}' | {ctx}")
        _limpiar(f"{sub}/ctxtVBAK-SKTO")
        session.findById(f"{sub}/ctxtVBAK-SKTO").text = codigo_pais
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)
        logger.info(f"[VA02] SKTO={codigo_pais} | {ctx}")

        # ZZ_PUEDES — puerto de descarga, leido dinamicamente via F4
        _limpiar(f"{sub}/ctxtVBAK-ZZ_PUEDES")
        cod_puerto_destino = _obtener_codigo_puerto_f4(
            session, f"{sub}/ctxtVBAK-ZZ_PUEDES", discharge_port, ctx
        )
        if cod_puerto_destino is None:
            logger.error(f"[VA02] ✗ No se pudo obtener código para DISCHARGE PORT='{discharge_port}' | {ctx}")
            return False
        logger.info(f"[VA02] ZZ_PUEDES={cod_puerto_destino} (DISCHARGE PORT={discharge_port}) | {ctx}")

        # ZZ_PUEEMB — puerto de embarque, match por LOADING PORT
        try:
            cod_puerto_embarque = obtener_puerto_embarque(loading_port)
            _limpiar(f"{sub}/ctxtVBAK-ZZ_PUEEMB")
            session.findById(f"{sub}/ctxtVBAK-ZZ_PUEEMB").text = cod_puerto_embarque
            logger.info(f"[VA02] ZZ_PUEEMB={cod_puerto_embarque} (LOADING PORT={loading_port}) | {ctx}")
        except ValueError as e:
            logger.warning(f"[VA02] ⚠ {e} — campo ZZ_PUEEMB no completado | {ctx}")

        # Fecha de arribo — columna ETA en formato DD.MM.YYYY
        eta = fila.get("ETA")
        if eta is not None and str(eta).strip() not in ("", "NaT", "nan"):
            try:
                fecha_eta = pd.to_datetime(eta).strftime("%d.%m.%Y")
                _limpiar(f"{sub}/ctxtVBAK-ZZ_FEC_ARRIBO")
                session.findById(f"{sub}/ctxtVBAK-ZZ_FEC_ARRIBO").text = fecha_eta
                logger.info(f"[VA02] ZZ_FEC_ARRIBO={fecha_eta} (ETA={eta}) | {ctx}")
            except Exception as e:
                logger.error(f"[VA02] ✗ Error convirtiendo fecha ETA: {e} | {ctx}")
                return False
        else:
            fecha_eta = None
            logger.warning(f"[VA02] ETA vacía, campo ZZ_FEC_ARRIBO no completado | {ctx}")

        # Fecha de zarpe — columna ETD en formato DD.MM.YYYY
        etd = fila.get("ETD")
        if etd is not None and str(etd).strip() not in ("", "NaT", "nan"):
            try:
                fecha_etd = pd.to_datetime(etd).strftime("%d.%m.%Y")
                _limpiar(f"{sub}/ctxtVBAK-ZZ_FEC_ZARPE_R")
                session.findById(f"{sub}/ctxtVBAK-ZZ_FEC_ZARPE_R").text = fecha_etd
                logger.info(f"[VA02] ZZ_FEC_ZARPE_R={fecha_etd} (ETD={etd}) | {ctx}")
            except Exception as e:
                logger.error(f"[VA02] ✗ Error convirtiendo fecha ETD: {e} | {ctx}")
                return False
        else:
            logger.warning(f"[VA02] ETD vacía, campo ZZ_FEC_ZARPE_R no completado | {ctx}")

        # Fecha de cargío y cierre FCL — columna ETD PLANTA en formato DD.MM.YYYY
        etd_planta = fila.get("ETD PLANTA")
        if etd_planta is not None and str(etd_planta).strip() not in ("", "NaT", "nan"):
            try:
                fecha_etd_planta = pd.to_datetime(etd_planta).strftime("%d.%m.%Y")
                _limpiar(f"{sub}/ctxtVBAK-ZZ_FEC_CARGIO")
                session.findById(f"{sub}/ctxtVBAK-ZZ_FEC_CARGIO").text = fecha_etd_planta
                logger.info(f"[VA02] ZZ_FEC_CARGIO={fecha_etd_planta} (ETD PLANTA={etd_planta}) | {ctx}")
                _limpiar(f"{sub}/ctxtVBAK-Z_FE_CLOSE_FCL")
                session.findById(f"{sub}/ctxtVBAK-Z_FE_CLOSE_FCL").text = fecha_etd_planta
                logger.info(f"[VA02] Z_FE_CLOSE_FCL={fecha_etd_planta} (ETD PLANTA={etd_planta}) | {ctx}")
            except Exception as e:
                logger.error(f"[VA02] ✗ Error convirtiendo fecha ETD PLANTA: {e} | {ctx}")
                return False
        else:
            logger.warning(f"[VA02] ETD PLANTA vacía, campos ZZ_FEC_CARGIO y Z_FE_CLOSE_FCL no completados | {ctx}")

        # Número de booking — columna BK/AWB
        bk_awb = str(fila.get("BK/AWB", "")).strip()
        if bk_awb and bk_awb not in ("nan", "NaT"):
            _limpiar(f"{sub}/txtVBAK-ZZ_BOOKING")
            session.findById(f"{sub}/txtVBAK-ZZ_BOOKING").text = bk_awb
            logger.info(f"[VA02] ZZ_BOOKING={bk_awb} | {ctx}")
        else:
            logger.warning(f"[VA02] BK/AWB vacío, campo ZZ_BOOKING no completado | {ctx}")

        # Bill of Lading — columna BL
        bl = str(fila.get("BL", "")).strip()
        if bl and bl not in ("nan", "NaT"):
            _limpiar(f"{sub}/txtVBAK-KZGBE")
            session.findById(f"{sub}/txtVBAK-KZGBE").text = bl
            logger.info(f"[VA02] KZGBE={bl} | {ctx}")
        else:
            logger.warning(f"[VA02] BL vacío, campo KZGBE no completado | {ctx}")

        # Nombre de nave — columna VESSEL
        vessel = str(fila.get("VESSEL", "")).strip()
        if vessel and vessel not in ("nan", "NaT"):
            _limpiar(f"{sub}/txtVBAK-ZZ_NAVE")
            session.findById(f"{sub}/txtVBAK-ZZ_NAVE").text = vessel
            logger.info(f"[VA02] ZZ_NAVE={vessel} | {ctx}")
        else:
            logger.warning(f"[VA02] VESSEL vacío, campo ZZ_NAVE no completado | {ctx}")

        # Número de viaje — columna VIAJE
        viaje = str(fila.get("VIAJE", "")).strip()
        if viaje and viaje not in ("nan", "NaT"):
            _limpiar(f"{sub}/txtVBAK-Z_VIAJE")
            session.findById(f"{sub}/txtVBAK-Z_VIAJE").text = viaje
            logger.info(f"[VA02] Z_VIAJE={viaje} | {ctx}")
        else:
            logger.warning(f"[VA02] VIAJE vacío, campo Z_VIAJE no completado | {ctx}")

        # Forma de envío y modalidad de flete — columna C/P → P=02 SIL PREPAID, C=01 SIL COLLECT
        cp = str(fila.get("C/P", "")).strip().upper()
        _mapa_cp = {"P": "02", "C": "01"}
        if cp in _mapa_cp:
            session.findById(f"{sub}/cmbVBAK-ZZ_FORM_ENV").key = _mapa_cp[cp]
            logger.info(f"[VA02] ZZ_FORM_ENV={_mapa_cp[cp]} (C/P={cp}) | {ctx}")
            session.findById(f"{sub}/cmbVBAK-ZZ_MOD_FLETE").key = _mapa_cp[cp]
            logger.info(f"[VA02] ZZ_MOD_FLETE={_mapa_cp[cp]} (C/P={cp}) | {ctx}")
        else:
            logger.warning(f"[VA02] C/P='{cp}' no reconocido (esperado P o C), campos ZZ_FORM_ENV y ZZ_MOD_FLETE no completados | {ctx}")

        # Precinto aduana — columna PRECINTO ADUANA
        precinto = str(fila.get("PRECINTO ADUANA", "")).strip()
        if precinto and precinto not in ("nan", "NaT"):
            _limpiar(f"{sub}/txtVBAK-Z_PRECINTO")
            session.findById(f"{sub}/txtVBAK-Z_PRECINTO").text = precinto
            logger.info(f"[VA02] Z_PRECINTO={precinto} | {ctx}")
        else:
            logger.warning(f"[VA02] PRECINTO ADUANA vacío, campo Z_PRECINTO no completado | {ctx}")

        # Precinto línea — columna PRECINTO LINEA
        precinto_lin = str(fila.get("PRECINTO LINEA", "")).strip()
        if precinto_lin and precinto_lin not in ("nan", "NaT"):
            _limpiar(f"{sub}/txtVBAK-Z_PRECINTO_LIN")
            session.findById(f"{sub}/txtVBAK-Z_PRECINTO_LIN").text = precinto_lin
            logger.info(f"[VA02] Z_PRECINTO_LIN={precinto_lin} | {ctx}")
        else:
            logger.warning(f"[VA02] PRECINTO LINEA vacío, campo Z_PRECINTO_LIN no completado | {ctx}")

        # Número de contenedor — columna CONTAINER NUM
        container = str(fila.get("CONTAINER NUM", "")).strip()
        if container and container not in ("nan", "NaT"):
            _limpiar(f"{sub}/txtVBAK-Z_FCL")
            session.findById(f"{sub}/txtVBAK-Z_FCL").text = container
            logger.info(f"[VA02] Z_FCL={container} | {ctx}")
        else:
            logger.warning(f"[VA02] CONTAINER NUM vacío, campo Z_FCL no completado | {ctx}")

        # Peso bruto — columna Peso Bruto (kg)
        peso = str(fila.get("Peso Bruto (kg)", "")).strip()
        if peso and peso not in ("nan", "NaT"):
            _limpiar(f"{sub}/txtVBAK-Z_PESO")
            session.findById(f"{sub}/txtVBAK-Z_PESO").text = peso
            logger.info(f"[VA02] Z_PESO={peso} | {ctx}")
        else:
            logger.warning(f"[VA02] Peso Bruto (kg) vacío, campo Z_PESO no completado | {ctx}")

        # Fecha documento — columna ETA (mismo valor que ZZ_FEC_ARRIBO)
        if fecha_eta is not None:
            _limpiar(f"{sub}/ctxtVBAK-Z_FECDOC")
            session.findById(f"{sub}/ctxtVBAK-Z_FECDOC").text = fecha_eta
            logger.info(f"[VA02] Z_FECDOC={fecha_eta} (ETA) | {ctx}")
        else:
            logger.warning(f"[VA02] ETA vacía, campo Z_FECDOC no completado | {ctx}")

        # Combobox tipo de transporte — siempre "1 Transporte marítimo"
        session.findById(f"{sub}/cmbVBAK-EXPVZ").key = "1"
        logger.info(f"[VA02] EXPVZ=1 (Transporte marítimo) | {ctx}")

        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)

        logger.success(f"[VA02] Completado OK | {ctx}")
        return {"Monto total CFR": valorneto}

    except Exception as e:
        logger.error(f"[VA02] ✗ Excepción inesperada: {e} | {ctx}")
        return False

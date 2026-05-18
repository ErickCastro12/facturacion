# Transacciones/T-va02.py
import time

BASE_PATH    = r"/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,{}]"
PARTNER_PATH = r"/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,{}]"
TABLA_PATH   = r"/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW"


def encontrar_primera_fila_vacia(session):
    fila = 0
    while True:
        try:
            texto = session.findById(BASE_PATH.format(fila)).text.strip()
            if texto == "":
                print(f"[VA02] Fila vacia encontrada: {fila}")
                return fila
            else:
                print(f"[VA02] Fila {fila} ocupada con: '{texto}', siguiente...")
                fila += 1
        except Exception as e:
            print(f"[VA02] Fin de tabla en fila {fila}: {e}")
            return None


def asignar_fila(session, fila, key_opcion, valor_partner):
    try:
        # Combo (select list)
        session.findById(BASE_PATH.format(fila)).key = key_opcion
        print(f"[VA02] Fila {fila} -> combo '{key_opcion}' asignado OK")
        time.sleep(0.3)

        # Campo de texto partner - misma fila
        campo = session.findById(PARTNER_PATH.format(fila))
        campo.setFocus()
        campo.text = valor_partner
        # session.findById("wnd[0]").sendVKey(0)
        print(f"[VA02] Fila {fila} -> partner '{valor_partner}' asignado OK")

        return True
    except Exception as e:
        print(f"[VA02] Error asignando fila {fila}: {e}")
        return False


def ejecutar_VA02(session):
    try:
        print("[VA02] Navegando a transaccion...")
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nVA02"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(2)

        session.findById("/app/con[0]/ses[0]/wnd[0]/usr/ctxtVBAK-VBELN").text = "20039216"
        time.sleep(2)
        session.findById("/app/con[0]/ses[0]/wnd[0]/usr/btnBT_SUCH").press()
        time.sleep(1)

        try:
            ventana = session.findById("/app/con[0]/ses[0]/wnd[1]")
            print(f"[VA02] Ventana emergente: {ventana.text}")
            session.findById("/app/con[0]/ses[0]/wnd[1]/tbar[0]/btn[0]").press()
        except:
            print("[VA02] Sin ventana emergente, continuando...")
        time.sleep(1)

        valorneto = session.findById("/app/con[0]/ses[0]/wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBAK-NETWR").text
        moneda    = session.findById("/app/con[0]/ses[0]/wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/ctxtVBAK-WAERK").text
        print(f"[VA02] Valor neto: {valorneto} | Moneda: {moneda}")

        session.findById("/app/con[0]/ses[0]/wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press()
        time.sleep(2)
        session.findById(r"/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07").select()
        time.sleep(2)

        fila = encontrar_primera_fila_vacia(session)
        if fila is None:
            print("[VA02] No hay filas disponibles")
            return False

        # Usa la misma fila para el combo y el campo de texto
        return asignar_fila(session, fila, "ZO", "10445")

    except Exception as e:
        print(f"[VA02] Error: {e}")
        return False
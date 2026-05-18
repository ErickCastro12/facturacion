import time

def ejecutar_VL02N(session, fila: dict):
    try:
        print("[VL02N] Navegando a transaccion...")
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nVL02N"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(2)

    
        print("[VL02N] Transaccion abierta correctamente")

        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        time.sleep(1)
        print("[VL02N] Regresando al menu principal")
        return True

    except Exception as e:
        print(f"[VL02N] Error: {e}")
        return False

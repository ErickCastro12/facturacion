import time

def ejecutar_ZEDOCPE001(session, fila: dict):
    try:
        print("[ZEDOCPE001] Navegando a transaccion...")
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nZEDOCPE001"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(2)

        print("[ZEDOCPE001] Transaccion abierta correctamente")

        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        time.sleep(1)
        print("[ZEDOCPE001] Regresando al menu principal")
        return True

    except Exception as e:
        print(f"[ZEDOCPE001] Error: {e}")
        return False

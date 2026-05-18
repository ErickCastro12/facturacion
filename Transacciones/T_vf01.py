import time

def ejecutar_VF01(session, fila: dict):
    try:
        print("[VF01] Navegando a transaccion...")
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nVF01"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(2)

       
        print("[VF01] Transaccion abierta correctamente")

        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        time.sleep(1)
        print("[VF01] Regresando al menu principal")
        return True

    except Exception as e:
        print(f"[VF01] Error: {e}")
        return False

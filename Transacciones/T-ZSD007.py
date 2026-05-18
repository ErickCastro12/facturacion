import time

def ejecutar_ZSD007(session):
    try:
        print("[ZSD007] Navegando a transaccion...")
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nZSD007"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(2)

        session.findById("wnd[0]/usr/txt%_P_VSTEL_%_APP_%-TEXT")
        print("[ZSD007] Transaccion abierta correctamente")

        session.findById("/app/con[0]/ses[0]/wnd[0]/usr/ctxtP_SONUM").text =  "20039216"       
        time.sleep(2)  
        session.findById("/app/con[0]/ses[0]/wnd[0]/tbar[1]/btn[8]").press()
        time.sleep(2) 
        session.findById("/app/con[0]/ses[0]/wnd[0]/tbar[1]/btn[37]")
        print("Encontro correctamente btn rechzaso")
        session.findById("/app/con[0]/ses[0]/wnd[0]/tbar[1]/btn[37]").press() 
        time.sleep(5)
        session.findById("/app/con[0]/ses[0]/wnd[0]/tbar[1]/btn[38]").press()
        time.sleep(5) 
        vpanelgrabado = session.findById("/app/con[0]/ses[0]/wnd[0]/sbar/pane[0]").text
        print(vpanelgrabado)
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        time.sleep(5)
        print("[ZSD007] Regresando al menu principal")
        session.findById("/app/con[0]/ses[0]/wnd[0]/usr/ctxtP_SONUM").text = "20039216"
        time.sleep(2)
        session.findById("/app/con[0]/ses[0]/wnd[0]/tbar[1]/btn[8]").press()
        time.sleep(2)

        vpanelgrilla = session.findById("/app/con[0]/ses[0]/wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell")
        print(f"[ZSD007] Tipo de componente: {vpanelgrilla.type}")

        col_list = list(vpanelgrilla.columnOrder)
        print(f"[ZSD007] Columnas encontradas: {col_list}")

        datos = []
        for fila in range(vpanelgrilla.rowCount):
            registro = {col: vpanelgrilla.getCellValue(fila, col) for col in col_list}
            datos.append(registro)

        import pandas as pd
        df = pd.DataFrame(datos)
        print(f"[ZSD007] Filas obtenidas: {len(df)}")
        print(df)

        # Total dinamico en la ultima fila
        ultima_fila = vpanelgrilla.rowCount - 1
        total_kwmeng = vpanelgrilla.getCellValue(ultima_fila, "KWMENG")
        total_vemng  = vpanelgrilla.getCellValue(ultima_fila, "VEMNG")
        print(f"[ZSD007] Total KWMENG: {total_kwmeng}")
        print(f"[ZSD007] Total VEMNG:  {total_vemng}")

        # Validar status verde en todas las filas (excepto la de totales)
        ICONO_VERDE = "@08\QSemáf.verde: Go; correcto@"
        filas_no_verdes = []

        for fila in range(ultima_fila):
            icono = vpanelgrilla.getCellValue(fila, "ICON")
            if icono != ICONO_VERDE:
                sonum = vpanelgrilla.getCellValue(fila, "SONUM")
                matnr = vpanelgrilla.getCellValue(fila, "MATNR")
                filas_no_verdes.append({"fila": fila, "SONUM": sonum, "MATNR": matnr, "ICON": icono})

        if not filas_no_verdes:
            print("[ZSD007] Todas las filas tienen status verde")
        else:
            print(f"[ZSD007] ADVERTENCIA: {len(filas_no_verdes)} fila(s) sin status verde:")
            for f in filas_no_verdes:
                print(f"  Fila {f['fila']} | SONUM: {f['SONUM']} | MATNR: {f['MATNR']} | ICON: {f['ICON']}")

        return df

    except Exception as e:
        print(f"[ZSD007] Error: {e}")
        return False
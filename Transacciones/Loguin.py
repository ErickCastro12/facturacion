import os
import win32com.client
import subprocess
import psutil
import time
import pandas as pd
from dotenv import load_dotenv

load_dotenv()

def cerrar_sap():
    for proc in psutil.process_iter():
        try:
            if "saplogon" in proc.name().lower() or "sapgui" in proc.name().lower():
                proc.kill()
                print(f"Proceso cerrado: {proc.name()}")
        except Exception:
            pass

    time.sleep(2)

def iniciar_sap(ruta_sap, ambiente):
    try:
        cerrar_sap()

        subprocess.Popen(ruta_sap)
        time.sleep(5)

        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine

        connection = application.OpenConnection(ambiente, True)

        for i in range(15):

            if connection.Children.Count > 0:
                print(f"Sesion obtenida en intento {i+1}")
                return connection.Children(0)

            print(f"Esperando sesion... intento {i+1}/15")
            time.sleep(1)

        print("Timeout: no se genero sesion")
        return None

    except Exception as e:
        print("Error al iniciar SAP:", e)
        return None

def login_sap(session, user, password):

    try:
        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = user
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = password
        session.findById("wnd[0]/usr/txtRSYST-LANGU").text = "ES"

        session.findById("wnd[0]").sendVKey(0)

        time.sleep(2)

        titulo = session.findById("wnd[0]").text

        if "SAP Easy Access" in titulo:
            print("Login correcto")
        else:
            print("Revisar credenciales - titulo:", titulo)

        return session

    except Exception as e:
        print("Error en login:", e)
        return None

def leer_excel():

    try:

        ruta_excel = os.getenv("INPUT_EXCEL_PATH", r"C:\RPA-FACTURACION\input\facturas.xlsx")

        df = pd.read_excel(ruta_excel)

        print("\n===== CONTENIDO EXCEL =====")
        print(df)

        print("\n===== PRIMERAS FILAS =====")
        print(df.head())

    except Exception as e:
        print("Error leyendo Excel:", e)

def ejecutar_login():

    ruta_sap = os.getenv("SAP_RUTA_EJECUTABLE", r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe")
    ambiente  = os.getenv("SAP_AMBIENTE", "SAP PRD")
    user      = os.getenv("SAP_USUARIO")
    password  = os.getenv("SAP_PASSWORD")

    session = iniciar_sap(ruta_sap, ambiente)

    if session:
        session = login_sap(session, user, password)

        # Leer Excel luego del login
        leer_excel()

    return session

if __name__ == "__main__":

    session = ejecutar_login()

    print("\nSesion:", session)
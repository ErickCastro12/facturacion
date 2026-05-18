import win32com.client
import subprocess
import time

subprocess.Popen(r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe")
time.sleep(5)

SapGuiAuto = win32com.client.GetObject("SAPGUI")
app = SapGuiAuto.GetScriptingEngine

try:
    connections = app.GetConnectionList()
    print(f"Total conexiones: {connections.Length}")
    for i in range(connections.Length):
        c = connections.Item(i)
        print(f"  [{i}] {repr(c.Name)}")
except Exception as e:
    print("Error listando conexiones:", e)
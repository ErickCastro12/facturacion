"""
helpers/cuadro_maestro.py

Arquitectura del Cuadro Maestro SEA.
Define columnas, grupos y funciones de lectura reutilizables.

Hoja principal: 'maestro'
Encabezado en fila 5 (index 4), datos desde fila 6.
Total columnas: 109 | Total registros: ~2361
"""

import os
import pandas as pd
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

# ---------------------------------------------------------------------------
# RUTA POR DEFECTO
# ---------------------------------------------------------------------------
RUTA_CUADRO_MAESTRO = os.getenv(
    "CUADRO_MAESTRO_PATH",
    str(Path(__file__).parent.parent / "Cmaster" / "2025 - SEA - CUADRO MAESTRO.xlsx")
)
HOJA_MAESTRO = "maestro"
FILA_ENCABEZADO = 4  # índice 0-based (fila 5 en Excel)


# ---------------------------------------------------------------------------
# GRUPOS DE COLUMNAS
# ---------------------------------------------------------------------------

COLS_IDENTIFICACION = {
    "BUSQUEDA":   "Código de búsqueda / número de booking",
    "año":   "Año del embarque",
    "MEDIO":      "Medio de transporte (SEA / AIR)",
    "# DE EMB":   "Número de embarques del pedido",
    "PRO":        "Producto procesado (SI / NO)",
}

COLS_PEDIDO = {
    "COD PED":    "Código de pedido SAP",
    "CULTIVO":    "Tipo de cultivo (BB=Blueberry, GRA=Granada, etc.)",
    "PLANTA":     "Planta de origen (CARAZ, SANTIAGO, etc.)",
    "SEM DES":    "Semana de despacho",
}

COLS_CARGA = {
    "ETD PLANTA":       "Fecha ETD desde planta",
    "DÍA":         "Día de la semana de carga",
    "HORA":             "Hora de carga en planta",
    "Número de cajas":  "Cantidad de cajas",
    "Peso Bruto (kg)":  "Peso bruto en kilogramos",
    "PALLETS":          "Cantidad de pallets",
    "PRECINTO ADUANA":  "Número de precinto de aduana",
    "PPRECINTO LINEA":  "Número de precinto de línea naviera",
    "CUT OFF":          "Fecha y hora de corte del buque",
}

COLS_DESTINO = {
    "CLIENTE":          "Nombre del cliente final",
    "CONSIGNEE":        "Consignatario en destino",
    "LOADING PORT":     "Puerto de embarque",
    "DISCHARGE PORT":   "Puerto de descarga",
    "PAIS":             "País de destino",
    "REGION":           "Región geográfica del destino",
}

COLS_TRANSPORTE = {
    "BK/AWB":               "Booking o Air Waybill",
    "BL":                   "Bill of Lading",
    "CONTAINER NUM":        "Número de contenedor",
    "Té":              "Tipo de equipo (AC/CT, G6, etc.)",
    "OPR LOGIST":           "Operador logístico",
    "Aerolínea/Naviera": "Línea naviera o aerolínea",
    "VESSEL":               "Nombre del buque",
    "viaje":                "Número de viaje del buque",
    "FREIGHT":              "Condición de flete (COLLECT / PREPAID)",
    "INLAND SIL":           "Flete inland (PREPAID / COLLECT)",
}

COLS_FECHAS_TRANSITO = {
    "ETD WEEK":     "Semana de ETD",
    "ETD":          "Fecha estimada de salida (puerto origen)",
    "ETA WEEK":     "Semana de ETA",
    "ETA":          "Fecha estimada de llegada (puerto destino)",
    "ETA FINAL":    "ETA final ajustada",
    "TT":           "Transit time en días",
    "LIBERACION BL": "Fecha de liberación del BL",
    "ETA LLEGADA ACT": "ETA de llegada actual/real",
    "ESPERA EN PUERTO": "Días de espera en puerto destino",
    "TT DEMORADO":  "Días de demora sobre el TT normal",
    "STATUS DE ARRIBO": "Estado del arribo (ETA A TIEMPO / DEMORADO)",
}

COLS_DOCUMENTOS = {
    "DAM":                    "Número de DAM (Declaración Aduanera de Mercancías)",
    "F. PL":                  "Fecha del Packing List",
    "F. PH":                  "Fecha del Certificado Fitosanitario",
    "F. BL":                  "Fecha del Bill of Lading",
    "F. CO":                  "Fecha del Certificado de Origen",
    "Nº CO DOCS FISICOS COMENTARIOS": "Número de CO / comentarios documentos físicos",
    "DOC VIR":                "¿Se enviaron documentos virtuales? (SI/NO)",
    "F. ENVIO DOC VIR":       "Fecha de envío de documentos virtuales",
    "F. ENVIO DOC FÍSICOS": "Fecha de envío de documentos físicos",
    "PRIMER DHL":             "Número de guía DHL principal",
    "DHL":                    "Confirmación de envío DHL (X = enviado)",
    "SEGUNDO DHL":            "Número de guía DHL secundario",
    "STATUS DOC FIS":         "Estado documentos físicos (SI = OK)",
    "STATUS FACT":            "Estado factura (SI = OK)",
    "STATUS  PL":             "Estado Packing List (SI = OK)",
    "STATUS PH":              "Estado Fitosanitario (SI = OK)",
    "STATUS  BL":             "Estado BL (SI = OK)",
    "STATUS CO":              "Estado Certificado de Origen (SI = OK)",
    "LAR":                    "Estado LAR",
}

COLS_FACTURACION = {
    "COMMERCIAL INVOICE":           "Número de factura comercial",
    "INVOICE SUNAT":                "Número de factura SUNAT",
    "F. FACTURA":                   "Fecha de facturación",
    "C/P":                          "Condición comercial (C=CFR / P=FOB)",
    "Cod. entrega de exportación": "Código de entrega SAP",
    "Cod. Factura SAP":             "Número de factura en SAP",
    "Monto total CFR":              "Monto total en condición CFR",
    "Moneda":                       "Moneda de la factura",
    "Precio FOB total":             "Precio FOB total",
    "Precio Flete $":               "Precio del flete en dólares",
    "Precio FOB unitario":          "Precio FOB por unidad",
    "Aprobado":                     "¿Aprobado por gerencia? (OK)",
    "Fecha de envío AGENTE ADUANA": "Fecha de envío al agente de aduana",
    "ENVIADO POR CORREO":           "Estado de envío por correo",
    "MES FACTURA SUNAT":            "Mes de facturación SUNAT (ene, feb, ...)",
    "ESTADO FACTURACIÓN":      "Estado del proceso de facturación RPA",
}

COLS_ADUANA = {
    "STATUS DE DUA":          "Estado de la DUA (REGULARIZADO / PENDIENTE)",
    "FECHA DE REGULARIZACION": "Fecha de regularización aduanera",
    "DIAS DE REGULARIZACION": "Días para regularizar la DUA",
    "T DE ENTREGA V DOCS":    "Tiempo de entrega documentos virtuales (días)",
    "T DE ENTREGA F DOCS":    "Tiempo de entrega documentos físicos (días)",
    "TT DEL ENVIO":           "Transit time del envío de documentos",
    "COMENTARIO DEL ENVIO":   "Comentario sobre el envío de documentos",
    "DIF HORAS CARGA":        "Diferencia en horas respecto a la carga programada",
    "STATUS":                 "Estado general del embarque (EN DESTINO, EN TRANSITO, etc.)",
}

COLS_PAGOS = {
    "ANTICIPOS":        "¿Tiene anticipo? (SI/NO)",
    "STATUS ANTICIPO":  "Estado del anticipo (PAGADO / PENDIENTE)",
    "FECHA PAGO":       "Fecha de pago del anticipo",
}

COLS_ORGANICOS = {
    "CERT ORGANICOS":   "¿Requiere certificado orgánico? (SI/NO)",
    "STATUS NOP":       "Estado del certificado NOP",
    "ENVIO DOCS ORG":   "Fecha de envío de documentos orgánicos",
}

COLS_COSTOS_ADICIONALES = {
    "FLETE DET":        "Flete de detención",
    "SIL DET (SIN IGV)": "Costo SIL de detención sin IGV",
    "FLETE KILO":       "Flete por kilogramo",
}

COLS_INCIDENTES = {
    "CAMBIOS EN PED POR PLANNING": "Cambios en el pedido realizados por planning",
    "INCIDENTES (SI/NO)":          "¿Hubo incidentes? (SI/NO)",
    "RESPONSABLE":                 "Responsable del incidente",
    "RECLAMO NAVIERA U OPLOG":     "Reclamo a naviera u operador logístico",
    "TIPO DE INCIDENTE":           "Tipo de incidente ocurrido",
    "COMENTARIOS":                 "Comentarios adicionales sobre el embarque",
}

COLS_CALIDAD = {
    "TIPO DE CULTIVO":          "Clasificación del cultivo",
    "SCORE":                    "Score de calidad",
    "PALLETS RECHAZADOS/POBRES": "Cantidad de pallets rechazados o de baja calidad",
    "INCIDENTE CALIDAD":        "¿Hubo incidente de calidad?",
}

COLS_ESTADO = {
    "ESTADO FACTURACIÓN":  "Estado del proceso de facturación RPA",
    "ESTADO":                   "Estado general del registro (OK / PENDIENTE / ERROR)",
}

# Mapa completo: nombre_columna → descripción
TODAS_LAS_COLUMNAS: dict = {
    **COLS_IDENTIFICACION,
    **COLS_PEDIDO,
    **COLS_CARGA,
    **COLS_DESTINO,
    **COLS_TRANSPORTE,
    **COLS_FECHAS_TRANSITO,
    **COLS_DOCUMENTOS,
    **COLS_FACTURACION,
    **COLS_ADUANA,
    **COLS_PAGOS,
    **COLS_ORGANICOS,
    **COLS_COSTOS_ADICIONALES,
    **COLS_INCIDENTES,
    **COLS_CALIDAD,
    **COLS_ESTADO,
}


# ---------------------------------------------------------------------------
# FUNCIONES DE LECTURA
# ---------------------------------------------------------------------------

def leer_maestro(ruta: str = RUTA_CUADRO_MAESTRO) -> pd.DataFrame:
    """Lee la hoja 'maestro' completa y retorna un DataFrame limpio."""
    df = pd.read_excel(ruta, sheet_name=HOJA_MAESTRO, header=FILA_ENCABEZADO)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all")
    return df


def leer_columnas(columnas: list[str], ruta: str = RUTA_CUADRO_MAESTRO) -> pd.DataFrame:
    """Retorna solo las columnas indicadas del cuadro maestro."""
    df = leer_maestro(ruta)
    cols_existentes = [c for c in columnas if c in df.columns]
    return df[cols_existentes]


def leer_sin_estado_ok(ruta: str = RUTA_CUADRO_MAESTRO) -> pd.DataFrame:
    """Retorna filas donde ESTADO no sea 'OK' (pendientes, vacíos o con error)."""
    df = leer_maestro(ruta)
    col = "ESTADO"
    if col not in df.columns:
        raise KeyError(f"Columna '{col}' no encontrada en el archivo.")
    pendientes = df[
        df[col].isna() | (df[col].astype(str).str.strip().str.upper() != "OK")
    ].copy()
    return pendientes


def leer_pendientes_facturar(ruta: str = RUTA_CUADRO_MAESTRO) -> pd.DataFrame:
    """Retorna filas con ESTADO FACTURACIÓN vacío y COD PED válido."""
    df = leer_maestro(ruta)
    col_estado = "ESTADO FACTURACIÓN"
    col_pedido = "COD PED"

    if col_estado not in df.columns or col_pedido not in df.columns:
        raise KeyError(f"Columnas requeridas no encontradas: '{col_estado}', '{col_pedido}'")

    pendientes = df[
        df[col_pedido].notna() &
        (df[col_estado].isna() | (df[col_estado].astype(str).str.strip() == ""))
    ].copy()

    return pendientes


def leer_por_pedido(cod_pedido: str | int, ruta: str = RUTA_CUADRO_MAESTRO) -> pd.DataFrame:
    """Retorna las filas que coincidan con un código de pedido SAP."""
    df = leer_maestro(ruta)
    return df[df["COD PED"].astype(str).str.strip() == str(cod_pedido).strip()]


def resumen_estado(ruta: str = RUTA_CUADRO_MAESTRO) -> pd.Series:
    """Cuenta los embarques agrupados por STATUS general."""
    df = leer_maestro(ruta)
    return df["STATUS"].value_counts(dropna=False)


# ---------------------------------------------------------------------------
# USO DIRECTO (debug)
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    df = leer_maestro()
    print(f"Filas totales   : {len(df)}")
    print(f"Columnas totales: {len(df.columns)}")
    print("\nPrimeras 5 filas (columnas clave):")
    cols_clave = ["BUSQUEDA", "COD PED", "CLIENTE", "ETD", "ETA", "STATUS", "ESTADO FACTURACIÓN"]
    cols_disp = [c for c in cols_clave if c in df.columns]
    print(df[cols_disp].head(5).to_string(index=False))
    print("\nResumen de status:")
    print(resumen_estado())

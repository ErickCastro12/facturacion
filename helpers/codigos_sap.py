"""
helpers/codigos_sap.py

Tablas de referencia de códigos SAP para campos de exportación.
Cuando lleguen nuevos países, agregar su bloque correspondiente.
"""

# ---------------------------------------------------------------------------
# PUERTOS DE DESTINO POR PAÍS
# Estructura: { "CODIGO_PAIS": { "NOMBRE_CIUDAD": "CODIGO_SAP" } }
# ---------------------------------------------------------------------------

PUERTOS_DESTINO = {
    "US": {
        "ATLANTA":          "Z001",
        "AUSTIN":           "Z002",
        "BALTIMORE":        "Z003",
        "BOSTON":           "Z004",
        "CARLISLE":         "Z005",
        "CHARLESTON":       "Z006",
        "CHARLOTTE":        "Z007",
        "CHATTANOOGA":      "Z008",
        "CHICAGO":          "Z009",
        "CHINO":            "Z010",
        "DAYTON":           "Z011",
        "DENVER":           "Z012",
        "EDISON":           "Z013",
        "FONTANA":          "Z014",
        "FORT WAYNE":       "Z015",
        "GALESBURG":        "Z016",
        "GENEVA":           "Z017",
        "HANNIBAL":         "Z018",
        "HARRISON":         "Z019",
        "HOUSTON":          "Z020",
        "INDIANAPOLIS":     "Z021",
        "JANESVILLE":       "Z022",
        "KANSAS CITY":      "Z023",
        "LANSING":          "Z024",
        "LONG BEACH":       "Z025",
        "LOS ANGELES":      "Z026",
        "MESQUITE":         "Z027",
        "MIAMI":            "Z028",
        "MINNEAPOLIS":      "Z029",
        "MIRA LOMA":        "Z030",
        "MODESTO":          "Z031",
        "MONTGOMERY":       "Z032",
        "NAPERVILLE":       "Z033",
        "NAPLES":           "Z034",
        "NEW JERSEY":       "Z035",
        "NEW ORLEANS":      "Z036",
        "NEW YORK":         "Z037",
        "NORFOLK":          "Z038",
        "OAKLAND":          "Z039",
        "ORLANDO":          "Z040",
        "PALMYRA":          "Z041",
        "PHILADELPHIA":     "Z042",
        "PHOENIX":          "Z043",
        "POMPANO BEACH":    "Z044",
        "PORT EVERGLADES":  "Z045",
        "PORTLAND":         "Z046",
        "ROANOKE":          "Z047",
        "SAN ANTONIO":      "Z048",
        "SAVANNAH":         "Z049",
        "SEATTLE":          "Z050",
        "SOCIAL CIRCLE":    "Z051",
        "STOCKTON":         "Z052",
        "TACOMA":           "Z053",
        "TAMPA":            "Z054",
        "VINELAND":         "Z055",
        "YODER":            "Z056",
        "STA. MARIA, CA":   "Z057",
        "MILTON":           "Z058",
        "NEWPORT":          "Z059",
        "TIPP CITY":        "Z060",
        "BROCKTON, MA":     "Z061",
        "NEWPORT BEACH":    "Z062",
        "LIVERPOOL":        "Z063",
        "BOLINGBROOK":      "Z064",
        "CANTON":           "Z065",
        "EVERETT":          "Z066",
        "SAN FRANCISCO":    "Z067",
        "NEWARK":           "Z068",
        "FLOWER MOUND":     "Z069",
        "PORT ELIZABETH":   "Z070",
        "MINOOKA":          "Z071",
        "NAZARETH":         "Z072",
        "NORTHWEST":        "Z073",
        "LACEY":            "Z074",
        "IRVING":           "Z075",
        "DAYTONA":          "Z076",
        "LEBANON":          "Z078",
        "JOLIET":           "Z079",
        "WILMINGTON":       "Z080",
        "GLOUCESTER":       "Z081",
        "HUENEME":          "Z082",
    },
    "PE": {
        "CALLAO PORT":  "Z001",
        "CALLAO":       "Z001",
        "PAITA PORT":   "Z002",
        "PAITA":        "Z002",
        "LIMA AIRPORT": "Z003",
        "LIMA":         "Z003",
        "PISCO PORT":   "Z004",
        "PISCO":        "Z004",
        "CHANCAY PORT": "Z005",
        "CHANCAY":      "Z005",
    },
    # Agregar más países aquí cuando se tengan los códigos:
    # "GB": { "LONDON": "XXXX", ... },
    # "NL": { "ROTTERDAM": "XXXX", ... },
}

# ---------------------------------------------------------------------------
# PUERTOS DE EMBARQUE (origen Peru)
# Pendiente: agregar códigos SAP cuando sean compartidos
# ---------------------------------------------------------------------------
PUERTOS_EMBARQUE = {
    # CALLAO - Z001
    "CALLAO":           "Z001",
    "CALLAO PORT":      "Z001",
    "PUERTO CALLAO":    "Z001",
    # PAITA - Z002
    "PAITA":            "Z002",
    "PAITA PORT":       "Z002",
    "PUERTO PAITA":     "Z002",
    # LIMA AIRPORT - Z003
    "LIMA AIRPORT":     "Z003",
    "LIMA":             "Z003",
    "AEROPUERTO LIMA":  "Z003",
    # PISCO - Z004
    "PISCO":            "Z004",
    "PISCO PORT":       "Z004",
    "PUERTO PISCO":     "Z004",
    # CHANCAY - Z005
    "CHANCAY":          "Z005",
    "CHANCAY PORT":     "Z005",
    "PUERTO CHANCAY":   "Z005",
    # BOLIVAR - alias de CALLAO PORT
    "BOLIVAR":          "Z001",
}


# ---------------------------------------------------------------------------
# FUNCIONES DE LOOKUP
# ---------------------------------------------------------------------------

def obtener_puerto_destino(pais: str, ciudad: str) -> str:
    """
    Retorna el código SAP del puerto de destino.
    pais:   código ISO del país (ej: 'US')
    ciudad: nombre de la ciudad del Discharge Port (ej: 'PHILADELPHIA')
    """
    pais_upper   = pais.strip().upper()
    ciudad_upper = ciudad.strip().upper()

    pais_puertos = PUERTOS_DESTINO.get(pais_upper)
    if pais_puertos is None:
        raise ValueError(f"País '{pais}' no tiene puertos de destino configurados en codigos_sap.py")

    codigo = pais_puertos.get(ciudad_upper)
    if codigo is None:
        raise ValueError(
            f"Ciudad '{ciudad}' no encontrada para país '{pais}'. "
            f"Ciudades disponibles: {list(pais_puertos.keys())}"
        )
    return codigo


def obtener_puerto_embarque(puerto: str) -> str:
    """Retorna el código SAP del puerto de embarque (origen Perú)."""
    codigo = PUERTOS_EMBARQUE.get(puerto.strip().upper())
    if codigo is None:
        raise ValueError(
            f"Puerto de embarque '{puerto}' no configurado en codigos_sap.py. "
            f"Puertos disponibles: {list(PUERTOS_EMBARQUE.keys())}"
        )
    return codigo

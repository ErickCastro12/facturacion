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


# ---------------------------------------------------------------------------
# MAPA DE PAÍSES — nombre/alpha-3 del Excel → código ISO alpha-2 para SAP
# Cubre los países detectados en el Cuadro Maestro + aliases comunes
# ---------------------------------------------------------------------------
MAPA_PAIS: dict[str, str] = {
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
    "CYP": "CY", "IRQ": "IQ", "LBN": "LB",
    # ── Nombres en español ────────────────────────────────────────────────
    "PERU": "PE", "PERÚ": "PE",
    "ESTADOS UNIDOS": "US", "EE.UU.": "US",
    "MEXICO": "MX", "MÉXICO": "MX",
    "BRASIL": "BR",
    "ALEMANIA": "DE",
    "ESPAÑA": "ES", "ESPANA": "ES",
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
    "EMIRATOS": "AE", "EMIRATOS ARABES UNIDOS": "AE", "EAU": "AE",
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
    "CHIPRE": "CY", "CYPRUS": "CY",
    "EGIPTO": "EG",
    "IRAQ": "IQ", "IRAK": "IQ",
    "LIBANO": "LB", "LÍBANO": "LB",
    "SINGAPOUR": "SG",
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
    "EGYPT": "EG",
    "IRAQ": "IQ",
    "LEBANON": "LB",
    "CYPRUS": "CY",
}


def normalizar_pais(valor: str) -> str:
    """
    Convierte el valor de la columna PAIS del Excel al código ISO alpha-2 que
    espera SAP. Si ya es de 2 letras lo devuelve tal cual. Lanza ValueError
    si no se puede resolver.
    """
    v = valor.strip().upper()
    if len(v) == 2:
        return v
    codigo = MAPA_PAIS.get(v)
    if codigo is None:
        raise ValueError(
            f"País '{valor}' no encontrado en MAPA_PAIS (codigos_sap.py). "
            "Agregar la entrada correspondiente."
        )
    return codigo

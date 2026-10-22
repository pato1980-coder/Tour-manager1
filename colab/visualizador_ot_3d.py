# -*- coding: utf-8 -*-
"""
Visualizador 3D de Órdenes de Trabajo (OT) de mantenimiento — Google Colab
=========================================================================

Maqueta espacial 3D interactiva (Plotly FigureWidget + ipywidgets) que ubica
cada departamento en una grilla cuadrada centrada en (0,0) dentro de su piso,
con tamaño y color de burbuja según la densidad de fallas.

Uso en Colab:
    1. Sube el Excel a '/content/OT_CLASIFICADAS_ARBOL (1).xlsx'.
    2. Pega este script en una celda y ejecútala.
"""

# ---------------------------------------------------------------------------
# 0. DEPENDENCIAS Y ENTORNO
# ---------------------------------------------------------------------------
import re
import math
import time
import subprocess
import sys

try:
    import anywidget  # noqa: F401  (requerido por FigureWidget en plotly >= 6)
except ImportError:
    subprocess.run([sys.executable, "-m", "pip", "install", "-q", "anywidget"], check=False)

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import ipywidgets as widgets
from IPython.display import display

try:
    from google.colab import output as colab_output
    colab_output.enable_custom_widget_manager()  # necesario para FigureWidget en Colab
except ImportError:
    pass  # fuera de Colab (Jupyter local) no hace falta

# ---------------------------------------------------------------------------
# 1. CONFIGURACIÓN
# ---------------------------------------------------------------------------
RUTA_EXCEL = "/content/OT_CLASIFICADAS_ARBOL (1).xlsx"
COL_FECHA = "Creado"
COL_UNIDAD = "unidad"
COL_CLASE = "ESTADO CLASIFICACIÓN"  # opcional: habilita el filtro por clasificación
COL_TIPO = "ITEM"                   # opcional: tipo de falla para las sub-burbujas
SIN_TIPO = "Sin ítem"
MARGEN_EJES = 1.0          # margen estático de seguridad en X/Y/Z
GRID_MIN, GRID_MAX = 3, 7  # grilla dinámica 3x3 … 7x7

# Posiciones reales según planos de arquitectura (metros desde el centro de la planta,
# X hacia la derecha y Y hacia arriba de la lámina). "escala" pasa metros a unidades
# de grilla para que cada planta quepa en ±3. Los edificios que no aparecen aquí
# usan la grilla cuadrada genérica.
PLANTAS = {
    # Somma Asturias, láminas MP-PG-110 (piso 3), 111 (pisos 4–15) y 112 (pisos 16–18):
    # 12 deptos por piso en anillo alrededor del núcleo de ascensores y escalera.
    "Somma Asturias": {
        "escala": 3 / 16.8,
        "deptos": {
            9: (-13.2, 10.0), 11: (-4.8, 10.5), 12: (4.8, 10.5), 10: (13.2, 10.0),
            7: (-12.4, 2.85), 8: (12.4, 2.85),
            5: (-12.4, -2.85), 6: (12.4, -2.85),
            3: (-13.2, -10.0), 1: (-4.8, -10.5), 2: (4.8, -10.5), 4: (13.2, -10.0),
        },
        # Contorno exterior de la planta y del núcleo central (patio + circulaciones)
        "contornos": [
            [(-16.8, 13.4), (16.8, 13.4), (16.8, -13.3), (-16.8, -13.3), (-16.8, 13.4)],
            [(-8.4, 6.3), (8.4, 6.3), (8.4, -6.4), (-8.4, -6.4), (-8.4, 6.3)],
        ],
    },
    # Somma Las Clarisas, láminas LC A006 (piso 2), A007 (piso 3) y A008 (pisos 4–10),
    # escala 1:100. Torres A (bloque recto + inicio de la curva) y B (ala curva) en una
    # misma planta, por eso las claves son (torre, depto). Posiciones tomadas de las
    # etiquetas N° de la lámina del piso 3.
    "Somma Las Clarisas": {
        "escala": 3 / 31.4,
        "unir_torres": True,
        "deptos": {
            ("A", 1): (-2.6, 13.2), ("A", 2): (3.5, 18.9), ("A", 3): (6.2, 21.4),
            ("A", 4): (12.8, 25.7), ("A", 5): (21.0, 31.4), ("A", 6): (20.7, 23.0),
            ("A", 7): (18.3, 19.4), ("A", 8): (18.3, 15.9), ("A", 9): (18.3, 12.5),
            ("A", 10): (20.6, 7.1), ("A", 11): (8.8, 8.6), ("A", 12): (5.9, 6.9),
            ("B", 1): (-21.0, -29.3), ("B", 2): (-20.2, -23.6), ("B", 3): (-18.7, -15.7),
            ("B", 4): (-17.7, -12.3), ("B", 5): (-14.2, -4.6), ("B", 6): (-12.7, -1.7),
            ("B", 7): (-8.4, 5.4), ("B", 8): (-6.5, 8.1), ("B", 9): (2.5, 2.1),
            ("B", 10): (1.0, -4.1), ("B", 11): (-2.1, -7.6), ("B", 12): (-2.4, -11.3),
            ("B", 13): (-2.6, -14.5), ("B", 14): (-3.2, -18.3), ("B", 15): (-3.2, -21.7),
            ("B", 16): (0.5, -31.4), ("B", 17): (-3.4, -31.1), ("B", 18): (-7.2, -30.9),
            ("B", 19): (-10.2, -30.6), ("B", 20): (-14.3, -30.1),
        },
        # Recorrido de cada torre uniendo sus deptos en orden (anillo alrededor del patio)
        "torres_contorno": ["A", "B"],  # cada contorno solo en los pisos de su torre
        "contornos": [
            [(-2.6, 13.2), (3.5, 18.9), (6.2, 21.4), (12.8, 25.7), (21.0, 31.4), (20.7, 23.0),
             (18.3, 19.4), (18.3, 15.9), (18.3, 12.5), (20.6, 7.1), (8.8, 8.6), (5.9, 6.9),
             (-2.6, 13.2)],
            [(-21.0, -29.3), (-20.2, -23.6), (-18.7, -15.7), (-17.7, -12.3), (-14.2, -4.6),
             (-12.7, -1.7), (-8.4, 5.4), (-6.5, 8.1), (2.5, 2.1), (1.0, -4.1), (-2.1, -7.6),
             (-2.4, -11.3), (-2.6, -14.5), (-3.2, -18.3), (-3.2, -21.7), (0.5, -31.4),
             (-3.4, -31.1), (-7.2, -30.9), (-10.2, -30.6), (-14.3, -30.1), (-21.0, -29.3)],
        ],
    },
    # Ronda Lo Ovalle, lámina A05 "Planta general piso tipo" (escala 1:125), que muestra
    # las dos torres en su posición real: Boreal (deptos 1–9) hacia Av. Lo Ovalle y
    # Austral (deptos 10–17) hacia calle Jorge Cáceres. Las etiquetas de la lámina están
    # junto al pasillo central; cada punto se llevó al centro de su lado del pasillo.
    "Ronda Lo Ovalle": {
        "escala": 3 / 42.6,
        "unir_torres": True,
        "deptos": {
            ("Boreal", 1): (-7.1, 33.2), ("Boreal", 2): (-7.1, 35.0), ("Boreal", 3): (0.4, 35.0),
            ("Boreal", 4): (0.4, 30.2), ("Boreal", 5): (0.4, 25.7), ("Boreal", 6): (0.4, 16.2),
            ("Boreal", 7): (0.4, 10.3), ("Boreal", 8): (-7.1, 10.3), ("Boreal", 9): (-7.1, 14.8),
            ("Austral", 10): (-2.3, -12.6), ("Austral", 11): (-2.3, -8.1),
            ("Austral", 12): (5.9, -12.8), ("Austral", 13): (5.9, -16.0),
            ("Austral", 14): (5.9, -25.8), ("Austral", 15): (5.9, -30.9),
            ("Austral", 16): (-2.3, -35.5), ("Austral", 17): (-2.3, -31.1),
        },
        # Contorno de cada torre (muros exteriores), solo en los pisos de esa torre
        "torres_contorno": ["Boreal", "Austral"],
        "contornos": [
            [(-10.0, 42.6), (5.0, 42.6), (5.0, 5.0), (-10.0, 5.0), (-10.0, 42.6)],
            [(-6.4, -0.3), (10.0, -0.3), (10.0, -42.6), (-6.4, -42.6), (-6.4, -0.3)],
        ],
    },
    # Somma Ines de Suarez, lámina A-07 (piso 3), con torres MAYOR (deptos 1–43, en C
    # alrededor del patio) y MENOR (deptos 1–7, bloque al sur-oriente) en su posición real.
    # La lámina está impresa reducida (dice 1:100, pero los deslindes A-B = 106,51 m y
    # F-E = 164,16 m dan 0,0956 m por punto); las coordenadas usan esa calibración.
    # Las plantas de los pisos 2, 4, 5–8 y 9 tienen la misma distribución (±1 m).
    "Somma Ines de Suarez": {
        "escala": 3 / 45.6,
        "unir_torres": True,
        "deptos": {
            ("MAYOR", 1): (37.6, 29.0), ("MAYOR", 2): (37.6, 21.4), ("MAYOR", 3): (22.2, 20.1), ("MAYOR", 4): (12.3, 20.1),
            ("MAYOR", 5): (2.8, 20.1), ("MAYOR", 6): (-6.1, 22.0), ("MAYOR", 7): (-11.2, 20.8), ("MAYOR", 8): (-19.2, 20.8),
            ("MAYOR", 9): (-23.1, 20.8), ("MAYOR", 10): (-31.8, 14.5), ("MAYOR", 11): (-31.8, 7.9), ("MAYOR", 12): (-30.9, 0.6),
            ("MAYOR", 13): (-30.9, -9.2), ("MAYOR", 14): (-31.9, -17.0), ("MAYOR", 15): (-22.2, -22.5), ("MAYOR", 16): (-12.8, -22.5),
            ("MAYOR", 17): (-4.8, -22.5), ("MAYOR", 18): (4.4, -22.5), ("MAYOR", 19): (12.6, -22.5), ("MAYOR", 20): (11.9, -31.8),
            ("MAYOR", 21): (3.3, -31.8), ("MAYOR", 22): (-6.8, -31.8), ("MAYOR", 23): (-16.1, -31.9), ("MAYOR", 24): (-20.2, -31.9),
            ("MAYOR", 25): (-27.9, -31.9), ("MAYOR", 26): (-33.1, -31.9), ("MAYOR", 27): (-40.9, -31.4), ("MAYOR", 28): (-40.5, -17.1),
            ("MAYOR", 29): (-40.5, -10.1), ("MAYOR", 30): (-40.9, -5.1), ("MAYOR", 31): (-40.9, 3.7), ("MAYOR", 32): (-40.9, 7.9),
            ("MAYOR", 33): (-40.9, 15.2), ("MAYOR", 34): (-42.1, 30.3), ("MAYOR", 35): (-32.8, 29.8), ("MAYOR", 36): (-23.1, 29.8),
            ("MAYOR", 37): (-15.7, 29.8), ("MAYOR", 38): (-6.9, 30.4), ("MAYOR", 39): (-2.6, 30.4), ("MAYOR", 40): (5.8, 30.4),
            ("MAYOR", 41): (10.4, 30.4), ("MAYOR", 42): (18.5, 30.4), ("MAYOR", 43): (24.7, 30.4),
            ("MENOR", 1): (30.2, -14.3), ("MENOR", 2): (30.2, -4.6), ("MENOR", 3): (39.5, -4.8), ("MENOR", 4): (39.5, -14.5),
            ("MENOR", 5): (39.5, -24.1), ("MENOR", 6): (39.5, -31.3), ("MENOR", 7): (32.5, -30.7),
        },
        # Contorno de cada torre (muros exteriores), solo en los pisos de esa torre
        "torres_contorno": ["MAYOR", "MENOR"],
        "contornos": [
            [(-45.6, 35.3), (43.9, 35.3), (43.9, 15.8), (-25.5, 15.8), (-25.5, -15.8),
             (17.5, -15.8), (17.5, -35.3), (-45.6, -35.3), (-45.6, 35.3)],
            [(23.8, 0.5), (45.6, 0.5), (45.6, -35.3), (23.8, -35.3), (23.8, 0.5)],
        ],
    },
    # Somma Parque Bustamante, lámina M10 (piso 3, escala 1:125): torre NORTE (deptos 1–8,
    # bloque girado) y ORIENTE (deptos 9–17, bloque horizontal) en su posición real.
    # Las etiquetas están junto al pasillo central de cada torre; cada punto se llevó al
    # centro de su lado del pasillo. La torre NORTE de los pisos 20 a 25 (M27–M29) tiene
    # la misma distribución.
    "Somma Parque Bustamante": {
        "escala": 3 / 28.5,
        "unir_torres": True,
        "deptos": {
            ("NORTE", 1): (10.5, -2.4), ("NORTE", 2): (11.2, 1.9), ("NORTE", 3): (12.5, 9.0),
            ("NORTE", 4): (13.5, 14.6), ("NORTE", 5): (22.6, 13.0), ("NORTE", 6): (20.7, 2.2),
            ("NORTE", 7): (19.6, -3.6), ("NORTE", 8): (18.2, -11.7), ("ORIENTE", 9): (0.8, -18.6),
            ("ORIENTE", 10): (-5.4, -18.6), ("ORIENTE", 11): (-13.8, -18.6), ("ORIENTE", 12): (-22.8, -18.6),
            ("ORIENTE", 13): (-22.8, -11.0), ("ORIENTE", 14): (-15.8, -11.0), ("ORIENTE", 15): (-11.2, -11.0),
            ("ORIENTE", 16): (-4.2, -11.0), ("ORIENTE", 17): (1.6, -11.0),
        },
        # Contorno de cada torre, solo en los pisos de esa torre
        "torres_contorno": ["NORTE", "ORIENTE"],
        "contornos": [
            [(10.2, 22.4), (28.5, 19.1), (21.7, -18.2), (3.5, -14.9), (10.2, 22.4)],
            [(-28.5, -7.1), (5.9, -7.1), (5.9, -22.4), (-28.5, -22.4), (-28.5, -7.1)],
        ],
    },
}
TODOS = "Todos"

# Escala continua Azul -> Morado -> Rojo
ESCALA_COLORES = [
    [0.00, "#0b1a7a"],
    [0.25, "#2a1f9e"],
    [0.50, "#8a2be2"],
    [0.75, "#d0207a"],
    [1.00, "#ff1a1a"],
]

# ---------------------------------------------------------------------------
# 2. CARGA Y LIMPIEZA DE DATOS
# ---------------------------------------------------------------------------
df = pd.read_excel(RUTA_EXCEL)
df.columns = [str(c).strip() for c in df.columns]


def buscar_columna(df, nombre, alternativas=()):
    """Devuelve el nombre real de una columna sin distinguir mayúsculas/acentos."""
    normalizar = lambda s: (
        str(s).strip().lower()
        .replace("á", "a").replace("é", "e").replace("í", "i")
        .replace("ó", "o").replace("ú", "u")
    )
    objetivos = [normalizar(nombre)] + [normalizar(a) for a in alternativas]
    for col in df.columns:
        if normalizar(col) in objetivos:
            return col
    return None


col_fecha = buscar_columna(df, COL_FECHA, ("fecha", "fecha creacion", "created"))
col_unidad = buscar_columna(df, COL_UNIDAD, ("unidad", "depto", "departamento"))
col_edificio = buscar_columna(df, "Edificio", ("edificio", "torre", "proyecto", "propiedad"))
if col_fecha is None or col_unidad is None:
    raise KeyError(f"No se encontraron las columnas '{COL_FECHA}' y/o '{COL_UNIDAD}'. "
                   f"Columnas disponibles: {list(df.columns)}")
if col_edificio is None:
    df["Edificio"] = "Edificio único"
    col_edificio = "Edificio"


def parsear_fecha(valor):
    """Parseo flexible de fechas con formatos mixtos 12h ('a. m.'/'p. m.') y 24h."""
    if pd.isna(valor):
        return pd.NaT
    if isinstance(valor, (pd.Timestamp, np.datetime64)) or hasattr(valor, "year"):
        return pd.Timestamp(valor)
    if isinstance(valor, (int, float)):  # número de serie de Excel
        return pd.Timestamp("1899-12-30") + pd.to_timedelta(float(valor), unit="D")

    s = str(valor).strip().lower()
    s = s.replace("\xa0", " ").replace(" ", " ")
    # Quita la zona horaria final ('CLST', 'CLT', 'UTC'…) sin tocar am/pm
    s = re.sub(r"\s+(?!(?:am|pm)$)[a-z]{3,5}$", "", s)
    # Normaliza variantes: 'a. m.', 'a.m.', 'a m', 'am' -> 'AM' (ídem p. m.)
    s = re.sub(r"\b(a)\s*\.?\s*(m)\s*\.?(?=\s|$)", "AM", s)
    s = re.sub(r"\b(p)\s*\.?\s*(m)\s*\.?(?=\s|$)", "PM", s)
    s = re.sub(r"\s+", " ", s).strip()

    es_12h = s.endswith("AM") or s.endswith("PM")
    formatos_12h = [
        "%d/%m/%Y %I:%M:%S %p", "%d/%m/%Y %I:%M %p",
        "%d-%m-%Y %I:%M:%S %p", "%d-%m-%Y %I:%M %p",
        "%Y-%m-%d %I:%M:%S %p", "%Y-%m-%d %I:%M %p",
        "%d/%m/%y %I:%M:%S %p", "%d/%m/%y %I:%M %p",
    ]
    formatos_24h = [
        "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y",
        "%d-%m-%Y %H:%M:%S", "%d-%m-%Y %H:%M", "%d-%m-%Y",
        "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d",
        "%d/%m/%y %H:%M:%S", "%d/%m/%y %H:%M", "%d/%m/%y",
    ]
    for fmt in (formatos_12h if es_12h else formatos_24h):
        try:
            return pd.to_datetime(s, format=fmt)
        except (ValueError, TypeError):
            continue
    # Último recurso: inferencia de pandas (día primero, estilo Chile)
    return pd.to_datetime(s, dayfirst=True, errors="coerce")


df["Fecha"] = df[col_fecha].apply(parsear_fecha)


def extraer_piso_depto(unidad):
    """
    Extrae (Torre, Piso_Num, Depto_Num) de textos como '1204', 'A-1305',
    'NORTE-1403', 'B-B903', 'Santo Domingo-903' o '12-04'. Convención: los
    2 últimos dígitos son el departamento y los anteriores el piso; el texto
    antes del primer guion (si tiene letras) es la torre.
    """
    if pd.isna(unidad):
        return ("", np.nan, np.nan)
    s = str(unidad).strip()
    if re.fullmatch(r"\d+\.0", s):  # números leídos como float por Excel
        s = s[:-2]
    torre = s.split("-")[0].strip() if "-" in s else ""
    if not re.search(r"[^\W\d_]", torre):  # sin letras: no es nombre de torre
        torre = ""
    # Formato explícito piso-depto (p.ej. '12-04')
    m = re.search(r"(\d{1,3})\s*[-/]\s*(\d{1,2})\s*$", s)
    if m:
        return (torre, int(m.group(1)), int(m.group(2)))
    grupos = re.findall(r"\d+", s)
    if not grupos:
        return (torre, np.nan, np.nan)
    num = grupos[-1]  # el último bloque numérico es el nº de unidad
    if len(num) >= 3:
        return (torre, int(num[:-2]), int(num[-2:]))
    return (torre, 1, int(num))  # unidades de 1-2 dígitos: se asumen en piso 1


df[["Torre", "Piso_Num", "Depto_Num"]] = df[col_unidad].apply(
    lambda u: pd.Series(extraer_piso_depto(u)))

filas_antes = len(df)
df = df.dropna(subset=["Fecha", "Piso_Num", "Depto_Num"]).copy()
df["Piso_Num"] = df["Piso_Num"].astype(int)
df["Depto_Num"] = df["Depto_Num"].astype(int)
df["Edificio"] = df[col_edificio].astype(str).str.strip()
# Edificios con varias torres: cada torre se trata como un edificio aparte
# (salvo los que tienen plano con las torres juntas, que se muestran en una sola planta)
juntas = df["Edificio"].isin([e for e, p in PLANTAS.items() if p.get("unir_torres")])
varias = (df.groupby("Edificio")["Torre"].transform("nunique") > 1) & ~juntas
df.loc[varias, "Edificio"] = (df.loc[varias, "Edificio"] + " · "
                              + df.loc[varias, "Torre"].replace("", "sin torre"))
col_clase = buscar_columna(df, COL_CLASE, ("clasificacion", "estado clasificacion"))
df["Clase"] = (df[col_clase].fillna("Sin clasificación").astype(str).str.strip()
               if col_clase else "Sin clasificación")
col_tipo = buscar_columna(df, COL_TIPO, ("item", "tipo falla"))
df["Tipo"] = (df[col_tipo].fillna(SIN_TIPO).astype(str).str.strip().replace("", SIN_TIPO)
              if col_tipo else SIN_TIPO)
df["Depto_Txt"] = np.where(juntas, "Torre " + df["Torre"] + " · Depto ", "Depto ") \
    + df["Depto_Num"].astype(str)
df["Depto_ID"] = (df["Edificio"] + " | " + df["Torre"] + " | P" + df["Piso_Num"].astype(str)
                  + "-D" + df["Depto_Num"].astype(str).str.zfill(2))
print(f"Registros válidos: {len(df)} de {filas_antes} "
      f"({filas_antes - len(df)} descartados por fecha/unidad inválida)")

# Densidad acumulada de fallas por departamento único (todo el conjunto)
deptos = (df.groupby(["Depto_ID", "Edificio", "Torre", "Piso_Num", "Depto_Num", "Depto_Txt"])
            .size().reset_index(name="Densidad_Total"))

# ---------------------------------------------------------------------------
# 3. MODELAMIENTO ESPACIAL: GRILLA CUADRADA CENTRADA EN (0,0)
# ---------------------------------------------------------------------------
def tamano_grilla(n):
    """Lado de grilla dinámico (3x3 … 7x7) según el nº de deptos del piso."""
    lado = math.ceil(math.sqrt(max(n, 1)))
    if lado > GRID_MAX:
        print(f"Aviso: piso con {n} deptos excede {GRID_MAX}x{GRID_MAX}; se usa {lado}x{lado}.")
        return lado
    return max(GRID_MIN, lado)


def asignar_coordenadas(grupo):
    grupo = grupo.sort_values("Depto_Num").copy()
    plano = PLANTAS.get(grupo["Edificio"].iloc[0])
    claves = (list(zip(grupo["Torre"], grupo["Depto_Num"])) if plano and plano.get("unir_torres")
              else list(grupo["Depto_Num"]))
    if plano and all(k in plano["deptos"] for k in claves):
        grupo["X"] = [plano["deptos"][k][0] * plano["escala"] for k in claves]
        grupo["Y"] = [plano["deptos"][k][1] * plano["escala"] for k in claves]
        grupo["Grid_Size"] = 0  # posición según plano, no grilla
        return grupo
    grid_size = tamano_grilla(len(grupo))
    offset = (grid_size - 1) / 2.0           # centrado matemático en (0,0)
    idx = np.arange(len(grupo))
    grupo["X"] = (idx % grid_size) - offset
    grupo["Y"] = (idx // grid_size) - offset
    grupo["Grid_Size"] = grid_size
    return grupo


deptos = pd.concat(
    [asignar_coordenadas(g) for _, g in deptos.groupby(["Edificio", "Piso_Num"])],
    ignore_index=True,
)
deptos["Z"] = deptos["Piso_Num"]
df = df.merge(deptos[["Depto_ID", "X", "Y", "Z", "Grid_Size", "Densidad_Total"]],
              on="Depto_ID", how="left")

# ---------------------------------------------------------------------------
# 4. ESCALA FIJA ABSOLUTA (calculada ANTES de cualquier filtro)
# ---------------------------------------------------------------------------
LIM_XY = float(max(deptos["X"].abs().max(), deptos["Y"].abs().max())) + MARGEN_EJES
RANGO_X = [-LIM_XY, LIM_XY]
RANGO_Y = [-LIM_XY, LIM_XY]
RANGO_Z = [float(deptos["Z"].min()) - MARGEN_EJES, float(deptos["Z"].max()) + MARGEN_EJES]
DENSIDAD_MAX = int(deptos["Densidad_Total"].max())

# Proporción espacial fija (Z algo más alto para que se lean los pisos)
span_z = RANGO_Z[1] - RANGO_Z[0]
ASPECTO = dict(x=1, y=1, z=max(0.6, min(2.5, span_z / (2 * LIM_XY) * 0.6)))


def tamano_marcador(d):
    if d <= 1:
        return 4
    if d <= 5:
        return 8
    if d <= 10:
        return 14
    if d <= 15:
        return 22
    return 34


# Colores categóricos fijos por tipo (legibles sobre fondo negro); se reciclan si faltan
COLORES_TIPO = [
    "#FD3216", "#00FE35", "#6A76FC", "#FED4C4", "#FE00CE", "#0DF9FF", "#F6F926",
    "#FF9616", "#479B55", "#EEA6FB", "#DC587D", "#D626FF", "#6E899C", "#00B5F7",
    "#B68E00", "#C9FBE5", "#FF0092", "#22FFA7", "#E3EE9E", "#86CE00", "#BC7196",
    "#7E7DCD", "#FC6955", "#E48F72", "#FFFFFF", "#A0A0FF", "#FFD700",
]
COLOR_TIPO = {}
for t in df["Tipo"].value_counts().index:
    COLOR_TIPO[t] = ("#8A8F98" if t == SIN_TIPO
                     else COLORES_TIPO[len([k for k in COLOR_TIPO if k != SIN_TIPO]) % len(COLORES_TIPO)])


def posiciones_esfera(n, r):
    """Reparte n puntos sobre una esfera (espiral de Fibonacci), corrigiendo la
    proporción de los ejes para que la molécula se vea redonda."""
    rz = r * (span_z / ASPECTO["z"]) / ((2 * LIM_XY) / ASPECTO["x"])
    oro = math.pi * (3 - math.sqrt(5))
    out = []
    for k in range(n):
        yk = 0.0 if n == 1 else 1 - (k / (n - 1)) * 2
        rad, th = math.sqrt(1 - yk * yk), k * oro
        out.append((math.cos(th) * rad * r, math.sin(th) * rad * r, yk * rz))
    return out


# ---------------------------------------------------------------------------
# 5. FIGURA (tema oscuro, ejes congelados)
# ---------------------------------------------------------------------------
estilo_eje = dict(
    showbackground=True, backgroundcolor="#000000",
    gridcolor="#2a2a2a", zerolinecolor="#444444",
    color="#bbbbbb", autorange=False,
)

fig = go.FigureWidget(
    data=[go.Scatter3d(
        x=[], y=[], z=[], mode="markers",
        marker=dict(
            size=[], color=[], colorscale=ESCALA_COLORES,
            cmin=0, cmax=DENSIDAD_MAX,  # escala de color también fija
            opacity=0.9, line=dict(width=0),
            colorbar=dict(title=dict(text="Fallas", font=dict(color="#ddd")),
                          tickfont=dict(color="#ddd"), thickness=14),
        ),
        text=[], hovertemplate="%{text}<extra></extra>",
    ),
        # Trazas de la "molécula" (vacías hasta hacer doble clic en un depto)
        go.Scatter3d(x=[], y=[], z=[], mode="lines", hoverinfo="skip",
                     line=dict(color="#9a9aae", width=4)),
        go.Scatter3d(x=[], y=[], z=[], mode="markers", text=[],
                     hovertemplate="%{text}<extra></extra>",
                     marker=dict(size=18, color="#ffffff", opacity=1)),
        go.Scatter3d(x=[], y=[], z=[], mode="markers+text", text=[], hovertext=[],
                     hovertemplate="%{hovertext}<extra></extra>",
                     textposition="middle center", textfont=dict(color="#000", size=11),
                     marker=dict(size=[], color=[], opacity=1)),
        # Contorno de planta real (solo edificios con plano en PLANTAS)
        go.Scatter3d(x=[], y=[], z=[], mode="lines", hoverinfo="skip",
                     line=dict(color="#3b4058", width=2)),
    ],
    layout=go.Layout(
        template="plotly_dark",
        paper_bgcolor="#000000", plot_bgcolor="#000000",
        height=720, margin=dict(l=0, r=0, t=50, b=0),
        title=dict(text="Maqueta 3D de Fallas por Departamento", font=dict(color="#eee")),
        uirevision="fijo",  # conserva la cámara al actualizar datos
        showlegend=False,
        scene=dict(
            xaxis=dict(range=RANGO_X, title="X (grilla)", **estilo_eje),
            yaxis=dict(range=RANGO_Y, title="Y (grilla)", **estilo_eje),
            zaxis=dict(range=RANGO_Z, title="Piso", dtick=1, **estilo_eje),
            aspectmode="manual", aspectratio=ASPECTO,
            camera=dict(eye=dict(x=1.6, y=1.6, z=0.9)),
            bgcolor="#000000",
        ),
    ),
)

# ---------------------------------------------------------------------------
# 6. CONTROLES INTERACTIVOS
# ---------------------------------------------------------------------------
edificios = sorted(df["Edificio"].unique())
dd_edificio = widgets.Dropdown(options=[TODOS] + edificios, value=TODOS,
                               description="Edificio:", layout=widgets.Layout(width="280px"))
dd_piso = widgets.Dropdown(options=[TODOS], value=TODOS, description="Piso:",
                           layout=widgets.Layout(width="180px"))

fechas = sorted(df["Fecha"].dt.normalize().unique())
fechas = [pd.Timestamp(f) for f in fechas]
clases = df["Clase"].value_counts().index.tolist()
dd_clase = widgets.Dropdown(options=[TODOS] + clases, value=TODOS, description="Clasif.:",
                            layout=widgets.Layout(width="300px",
                                                  display="flex" if len(clases) > 1 else "none"))
sl_fechas = widgets.SelectionRangeSlider(
    options=[(f.strftime("%d/%m/%Y"), f) for f in fechas],
    index=(0, len(fechas) - 1), description="Fechas:",
    continuous_update=False, layout=widgets.Layout(width="620px"),
)
lbl_resumen = widgets.HTML()

# Simbología del depto desplegado (oculta mientras no haya ninguno)
html_leyenda = widgets.HTML()
btn_cerrar = widgets.Button(description="Cerrar", layout=widgets.Layout(width="90px"))
caja_leyenda = widgets.VBox(
    [html_leyenda, btn_cerrar],
    layout=widgets.Layout(display="none", width="290px", padding="10px 12px",
                          border="1px solid #2c2c3a", margin="50px 8px 0 0"),
)
caja_leyenda.add_class("ot-leyenda")
estado = {"abierto": None, "puntos": [], "ultimo": (None, 0.0)}


def actualizar_pisos(*_):
    base = df if dd_edificio.value == TODOS else df[df["Edificio"] == dd_edificio.value]
    pisos = [TODOS] + [str(p) for p in sorted(base["Piso_Num"].unique())]
    actual = dd_piso.value
    dd_piso.options = pisos
    dd_piso.value = actual if actual in pisos else TODOS


def actualizar_grafico(*_):
    f_ini, f_fin = sl_fechas.value
    m = (df["Fecha"] >= f_ini) & (df["Fecha"] < f_fin + pd.Timedelta(days=1))
    if dd_edificio.value != TODOS:
        m &= df["Edificio"] == dd_edificio.value
    if dd_piso.value != TODOS:
        m &= df["Piso_Num"] == int(dd_piso.value)
    if dd_clase.value != TODOS:
        m &= df["Clase"] == dd_clase.value

    # Densidad dentro del rango filtrado; posiciones y escalas siguen fijas
    sub = (df[m].groupby(["Depto_ID", "Edificio", "Piso_Num", "Depto_Num", "Depto_Txt", "X", "Y", "Z"])
               .size().reset_index(name="Densidad"))
    estado["puntos"] = sub["Depto_ID"].tolist()
    if estado["abierto"] not in estado["puntos"]:
        estado["abierto"] = None  # el filtro dejó fuera al depto abierto
    abierto = estado["abierto"]

    textos = [f"<b>{r.Edificio}</b><br>Piso {r.Piso_Num} · {r.Depto_Txt}"
              f"<br>Fallas en rango: <b>{r.Densidad}</b>"
              for r in sub.itertuples()]

    with fig.batch_update():
        t = fig.data[0]
        t.x, t.y, t.z = sub["X"].tolist(), sub["Y"].tolist(), sub["Z"].tolist()
        t.marker.size = [tamano_marcador(d) for d in sub["Densidad"]]
        t.marker.color = sub["Densidad"].tolist()
        t.text = textos
        t.marker.opacity = 0.12 if abierto else 0.9  # oscurece el resto al desplegar
        dibujar_molecula(df[m & (df["Depto_ID"] == abierto)] if abierto else None)
        dibujar_contorno()
        # Re-afirma los límites absolutos para que nada se auto-escale
        fig.layout.scene.xaxis.range = RANGO_X
        fig.layout.scene.yaxis.range = RANGO_Y
        fig.layout.scene.zaxis.range = RANGO_Z

    lbl_resumen.value = (f"<span style='color:#9aa0a6'>{int(m.sum())} OT · "
                         f"{len(sub)} departamentos visibles · máx. "
                         f"{int(sub['Densidad'].max()) if len(sub) else 0} fallas/depto · "
                         f"Clic en una burbuja para ver sus fallas por tipo</span>")


dd_edificio.observe(lambda ch: (actualizar_pisos(), actualizar_grafico()), names="value")
dd_piso.observe(actualizar_grafico, names="value")
dd_clase.observe(actualizar_grafico, names="value")


def dibujar_contorno():
    """Dibuja el contorno de la planta en cada piso si el edificio tiene plano real."""
    plano = PLANTAS.get(dd_edificio.value)
    lx, ly, lz = [], [], []
    if plano:
        for k, c in enumerate(plano["contornos"]):
            # Cada contorno solo en los pisos de su torre, si el plano la indica
            sel = deptos["Edificio"] == dd_edificio.value
            if plano.get("torres_contorno"):
                sel &= deptos["Torre"] == plano["torres_contorno"][k]
            pisos = sorted(deptos.loc[sel, "Z"].unique())
            if dd_piso.value != TODOS:
                pisos = [z for z in pisos if z == int(dd_piso.value)]
            for z in pisos:
                lx += [x * plano["escala"] for x, _ in c] + [None]
                ly += [y * plano["escala"] for _, y in c] + [None]
                lz += [int(z)] * len(c) + [None]
    fig.data[4].x, fig.data[4].y, fig.data[4].z = lx, ly, lz


def dibujar_molecula(ots_depto):
    """Sub-burbujas por tipo alrededor del depto: tamaño = nº de OT, color = tipo."""
    enlaces, nucleo, subs = fig.data[1], fig.data[2], fig.data[3]
    if ots_depto is None or ots_depto.empty:
        for tr in (enlaces, nucleo, subs):
            tr.x, tr.y, tr.z = [], [], []
        caja_leyenda.layout.display = "none"
        return
    r0 = ots_depto.iloc[0]
    cx, cy, cz = float(r0["X"]), float(r0["Y"]), float(r0["Z"])
    conteo = ots_depto["Tipo"].value_counts()
    max_n = int(conteo.iloc[0])
    lx, ly, lz, sx, sy, sz = [], [], [], [], [], []
    for (dx, dy, dz) in posiciones_esfera(len(conteo), 1.4):
        lx += [cx, cx + dx, None]; ly += [cy, cy + dy, None]; lz += [cz, cz + dz, None]
        sx.append(cx + dx); sy.append(cy + dy); sz.append(cz + dz)
    enlaces.x, enlaces.y, enlaces.z = lx, ly, lz
    nucleo.x, nucleo.y, nucleo.z = [cx], [cy], [cz]
    nucleo.text = [f"<b>{r0['Edificio']}</b><br>Piso {r0['Piso_Num']} · {r0['Depto_Txt']}"]
    subs.x, subs.y, subs.z = sx, sy, sz
    subs.marker.size = [12 + 26 * math.sqrt(n / max_n) for n in conteo]
    subs.marker.color = [COLOR_TIPO[t] for t in conteo.index]
    subs.text = [str(n) for n in conteo]
    subs.hovertext = [f"<b>{t}</b><br>{n} OT" for t, n in conteo.items()]

    filas = "".join(
        f"<div style='display:flex;align-items:center;gap:8px;margin:3px 0'>"
        f"<span style='width:12px;height:12px;border-radius:50%;background:{COLOR_TIPO[t]};"
        f"flex:none'></span><span style='flex:1'>{t}</span><b>{n}</b></div>"
        for t, n in conteo.items())
    html_leyenda.value = (
        f"<div style='color:#ddd;font-size:13px'><div style='color:#fff;font-weight:600'>"
        f"{r0['Edificio']}</div><div style='color:#9aa0a6;font-size:12px;margin-bottom:8px'>"
        f"Piso {r0['Piso_Num']} · {r0['Depto_Txt']} · {len(ots_depto)} OT por tipo</div>"
        f"{filas}<div style='color:#777;font-size:11px;margin-top:8px'>"
        f"Clic en otra burbuja para cambiar</div></div>")
    caja_leyenda.layout.display = "flex"


def al_hacer_clic(trace, points, selector):
    """Un clic abre o cierra la molécula del depto.

    En 3D, Plotly solo registra el clic si el botón sigue apretado durante un cuadro de
    dibujo (un doble clic rápido casi nunca llega), y un clic largo puede llegar repetido:
    se ignoran los clics del mismo depto durante 0,8 s tras abrir o cerrar.
    """
    if not points.point_inds:
        return
    if trace is fig.data[0]:
        depto = estado["puntos"][points.point_inds[0]]
    else:
        depto = estado["abierto"]  # clic en el núcleo de la molécula
    ahora = time.time()
    previo, t_previo = estado["ultimo"]
    if depto == previo and ahora - t_previo < 0.8:
        return
    estado["ultimo"] = (depto, ahora)
    estado["abierto"] = None if estado["abierto"] == depto else depto
    actualizar_grafico()


def cerrar_molecula(_=None):
    estado["abierto"] = None
    actualizar_grafico()


fig.data[0].on_click(al_hacer_clic)
fig.data[2].on_click(al_hacer_clic)
btn_cerrar.on_click(cerrar_molecula)
sl_fechas.observe(actualizar_grafico, names="value")

# ---------------------------------------------------------------------------
# 7. BOTÓN PANTALLA COMPLETA (API nativa + fallback CSS para el iFrame de Colab)
# ---------------------------------------------------------------------------
# Nota: ipywidgets.HTML inserta el contenido vía innerHTML, por lo que las
# etiquetas <script> NO se ejecutan; los atributos inline (onclick) y <style>
# sí funcionan. Toda la lógica va en el onclick.
JS_FULLSCREEN = r"""
(function(btn){
  var box = btn.closest('.ot-viz-container') || document.querySelector('.ot-viz-container');
  if (!box) return;
  var FB = 'ot-fs-fallback';
  function relayout(){
    setTimeout(function(){
      window.dispatchEvent(new Event('resize'));
      box.querySelectorAll('.js-plotly-plot').forEach(function(p){
        if (window.Plotly && Plotly.Plots) { try { Plotly.Plots.resize(p); } catch(e){} }
      });
    }, 120);
  }
  function setIframe(h){
    try { if (window.google && google.colab && google.colab.output)
            google.colab.output.setIframeHeight(h, true); } catch(e){}
  }
  function label(on){ btn.innerHTML = on ? '&#x2715;&nbsp; Salir de pantalla completa'
                                         : '&#x26F6;&nbsp; Pantalla completa'; }
  function exitFallback(){
    box.classList.remove(FB); document.documentElement.classList.remove('ot-fs-lock');
    setIframe(0); label(false); relayout();
  }
  function enterFallback(){
    box.classList.add(FB); document.documentElement.classList.add('ot-fs-lock');
    setIframe(Math.max(window.screen.availHeight || 0, window.innerHeight, 900));
    label(true); relayout();
    if (!window.__otEscBound){
      window.__otEscBound = true;
      document.addEventListener('keydown', function(e){
        var b = document.querySelector('.ot-viz-container.' + FB);
        if (e.key === 'Escape' && b){ b.querySelector('.ot-fs-btn').click(); }
      });
    }
  }
  var fsEl = document.fullscreenElement || document.webkitFullscreenElement;
  if (fsEl){
    (document.exitFullscreen || document.webkitExitFullscreen).call(document);
    box.classList.remove(FB); label(false); relayout(); return;
  }
  if (box.classList.contains(FB)){ exitFallback(); return; }

  var req = box.requestFullscreen || box.webkitRequestFullscreen || box.msRequestFullscreen;
  var enabled = document.fullscreenEnabled || document.webkitFullscreenEnabled;
  if (req && enabled){
    try {
      var p = req.call(box);
      // La clase también aplica el layout 100vw/100vh dentro del modo nativo
      box.classList.add(FB); label(true); relayout();
      if (p && p.catch){ p.catch(function(){ box.classList.remove(FB); enterFallback(); }); }
      document.addEventListener('fullscreenchange', function h(){
        if (!document.fullscreenElement){
          document.removeEventListener('fullscreenchange', h);
          box.classList.remove(FB); label(false); relayout();
        }
      });
      return;
    } catch(e){ /* sandbox del iFrame: usar fallback */ }
  }
  enterFallback();
})(this);
""".replace("\n", " ")

CSS_FULLSCREEN = """
<style>
  .ot-fs-btn{
    background:linear-gradient(135deg,#1f4fff,#8a2be2 55%,#ff1a1a);
    color:#fff;border:none;border-radius:10px;padding:8px 18px;
    font:600 13px/1.2 system-ui,-apple-system,'Segoe UI',Roboto,sans-serif;
    letter-spacing:.3px;cursor:pointer;box-shadow:0 4px 14px rgba(138,43,226,.45);
    transition:transform .15s ease, box-shadow .15s ease, filter .15s ease;
  }
  .ot-fs-btn:hover{transform:translateY(-1px);filter:brightness(1.12);
    box-shadow:0 6px 20px rgba(255,26,26,.45);}
  .ot-fs-btn:active{transform:translateY(0);}
  .ot-viz-container{background:#000 !important;border-radius:8px;padding:6px;}
  .ot-leyenda{background:rgba(8,8,14,.92) !important;border-radius:10px;align-self:flex-start;}
  .ot-leyenda button{background:#222 !important;color:#eee !important;border:1px solid #444 !important;
    border-radius:6px !important;}
  /* Fallback de alta prioridad: maximiza el contenedor a toda la ventana */
  .ot-viz-container.ot-fs-fallback{
    position:fixed !important;top:0 !important;left:0 !important;
    width:100vw !important;height:100vh !important;
    max-width:100vw !important;max-height:100vh !important;
    margin:0 !important;padding:8px !important;border-radius:0 !important;
    z-index:2147483647 !important;background:#000 !important;
    overflow:auto !important;box-sizing:border-box !important;
  }
  .ot-viz-container.ot-fs-fallback .js-plotly-plot,
  .ot-viz-container.ot-fs-fallback .plot-container,
  .ot-viz-container.ot-fs-fallback .svg-container{
    width:100% !important;height:calc(100vh - 110px) !important;
  }
  .ot-viz-container:fullscreen{width:100vw;height:100vh;background:#000;}
  html.ot-fs-lock, html.ot-fs-lock body{overflow:hidden !important;}
</style>
"""

btn_fullscreen = widgets.HTML(
    CSS_FULLSCREEN
    + f'<button class="ot-fs-btn" onclick="{JS_FULLSCREEN}">'
      '&#x26F6;&nbsp; Pantalla completa</button>'
)

# ---------------------------------------------------------------------------
# 8. ENSAMBLE Y DESPLIEGUE
# ---------------------------------------------------------------------------
fila_filtros = widgets.HBox([dd_edificio, dd_piso, dd_clase, btn_fullscreen],
                            layout=widgets.Layout(align_items="center", gap="12px"))
panel = widgets.VBox([fila_filtros, sl_fechas, lbl_resumen,
                      widgets.HBox([caja_leyenda, fig])])
panel.add_class("ot-viz-container")

actualizar_pisos()
actualizar_grafico()
display(panel)

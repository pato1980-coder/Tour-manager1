# -*- coding: utf-8 -*-
"""
Visualizador 3D de Órdenes de Trabajo (OT) — versión HTML autónoma
=================================================================

Misma maqueta 3D que el script de Colab, pero sin Colab ni Jupyter:
lee el Excel con Python y genera un único archivo .html (Plotly.js + filtros
en JavaScript) que se abre en cualquier navegador, sin servidor.

Uso:
    pip install pandas openpyxl plotly
    python visualizador_ot_3d_html.py "OT_CLASIFICADAS_ARBOL (1).xlsx"
    python visualizador_ot_3d_html.py datos.xlsx -o maqueta.html --cdn

Por defecto Plotly.js se incrusta en el HTML (funciona sin internet, ~4.5 MB);
con --cdn se carga desde internet y el archivo pesa unos pocos KB + datos.
"""

import argparse
import json
import math
import re
import sys
import webbrowser
from pathlib import Path

import numpy as np
import pandas as pd

# ---------------------------------------------------------------------------
# 1. CONFIGURACIÓN
# ---------------------------------------------------------------------------
COL_FECHA = "Creado"
COL_UNIDAD = "unidad"
COL_CLASE = "ESTADO CLASIFICACIÓN"  # opcional: habilita el filtro por clasificación
COL_TIPO = "ITEM"                   # opcional: tipo de falla para las sub-burbujas
SIN_TIPO = "Sin ítem"

# Colores categóricos para los tipos (legibles sobre fondo negro); se reciclan si faltan
COLORES_TIPO = [
    "#FD3216", "#00FE35", "#6A76FC", "#FED4C4", "#FE00CE", "#0DF9FF", "#F6F926",
    "#FF9616", "#479B55", "#EEA6FB", "#DC587D", "#D626FF", "#6E899C", "#00B5F7",
    "#B68E00", "#C9FBE5", "#FF0092", "#22FFA7", "#E3EE9E", "#86CE00", "#BC7196",
    "#7E7DCD", "#FC6955", "#E48F72", "#FFFFFF", "#A0A0FF", "#FFD700",
]
COLOR_SIN_TIPO = "#8A8F98"

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
}
MARGEN_EJES = 1.0          # margen estático de seguridad en X/Y/Z
GRID_MIN, GRID_MAX = 3, 7  # grilla dinámica 3x3 … 7x7

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


def cargar_datos(ruta_excel):
    df = pd.read_excel(ruta_excel)
    df.columns = [str(c).strip() for c in df.columns]

    col_fecha = buscar_columna(df, COL_FECHA, ("fecha", "fecha creacion", "created"))
    col_unidad = buscar_columna(df, COL_UNIDAD, ("unidad", "depto", "departamento"))
    col_edificio = buscar_columna(df, "Edificio", ("edificio", "torre", "proyecto", "propiedad"))
    if col_fecha is None or col_unidad is None:
        raise KeyError(f"No se encontraron las columnas '{COL_FECHA}' y/o '{COL_UNIDAD}'. "
                       f"Columnas disponibles: {list(df.columns)}")
    if col_edificio is None:
        df["Edificio"] = "Edificio único"
        col_edificio = "Edificio"

    df["Fecha"] = df[col_fecha].apply(parsear_fecha)
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
    return df


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


def construir_modelo(df):
    # Densidad acumulada de fallas por departamento único (todo el conjunto)
    deptos = (df.groupby(["Depto_ID", "Edificio", "Torre", "Piso_Num", "Depto_Num", "Depto_Txt"])
                .size().reset_index(name="Densidad_Total"))
    deptos = pd.concat(
        [asignar_coordenadas(g) for _, g in deptos.groupby(["Edificio", "Piso_Num"])],
        ignore_index=True,
    )
    deptos["Z"] = deptos["Piso_Num"]

    # Escala fija absoluta (calculada ANTES de cualquier filtro)
    lim_xy = float(max(deptos["X"].abs().max(), deptos["Y"].abs().max())) + MARGEN_EJES
    rango_z = [float(deptos["Z"].min()) - MARGEN_EJES, float(deptos["Z"].max()) + MARGEN_EJES]
    span_z = rango_z[1] - rango_z[0]

    # Cada OT se reduce a (índice de depto, día, clase, tipo) para filtrar en el navegador
    idx_depto = {d: i for i, d in enumerate(deptos["Depto_ID"])}
    clases = df["Clase"].value_counts().index.tolist()
    idx_clase = {c: i for i, c in enumerate(clases)}
    tipos = df["Tipo"].value_counts().index.tolist()
    idx_tipo = {t: i for i, t in enumerate(tipos)}
    colores_tipo, k = [], 0
    for t in tipos:
        if t == SIN_TIPO:
            colores_tipo.append(COLOR_SIN_TIPO)
        else:
            colores_tipo.append(COLORES_TIPO[k % len(COLORES_TIPO)])
            k += 1
    dias = df["Fecha"].dt.normalize()
    fecha_0 = dias.min()
    ots = np.column_stack([
        df["Depto_ID"].map(idx_depto).to_numpy(),
        (dias - fecha_0).dt.days.to_numpy(),
        df["Clase"].map(idx_clase).to_numpy(),
        df["Tipo"].map(idx_tipo).to_numpy(),
    ]).astype(int)

    return {
        "deptos": {
            "edificio": deptos["Edificio"].tolist(),
            "piso": deptos["Piso_Num"].astype(int).tolist(),
            "depto": deptos["Depto_Txt"].tolist(),
            "x": deptos["X"].astype(float).tolist(),
            "y": deptos["Y"].astype(float).tolist(),
            "z": deptos["Z"].astype(int).tolist(),
        },
        "ots": ots.ravel().tolist(),  # cuartetos aplanados [depto, día, clase, tipo, …]
        "clases": clases,
        "tipos": tipos,
        "coloresTipo": colores_tipo,
        "fecha0": fecha_0.strftime("%Y-%m-%d"),
        "nDias": int((dias.max() - fecha_0).days) + 1,
        "rangoX": [-lim_xy, lim_xy],
        "rangoY": [-lim_xy, lim_xy],
        "rangoZ": rango_z,
        "aspecto": {"x": 1, "y": 1, "z": max(0.6, min(2.5, span_z / (2 * lim_xy) * 0.6))},
        "densidadMax": int(deptos["Densidad_Total"].max()),
        "escala": ESCALA_COLORES,
        # Contornos de planta (en unidades de grilla) y pisos donde dibujarlos
        "plantas": {
            ed: {
                "contornos": [[[x * p["escala"], y * p["escala"]] for x, y in c]
                              for c in p["contornos"]],
                "pisos": sorted(deptos.loc[deptos["Edificio"] == ed, "Z"].astype(int).unique().tolist()),
            }
            for ed, p in PLANTAS.items() if (deptos["Edificio"] == ed).any()
        },
    }


# ---------------------------------------------------------------------------
# 4. PLANTILLA HTML (Plotly.js + filtros + pantalla completa)
# ---------------------------------------------------------------------------
PLANTILLA = r"""<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Maqueta 3D de Fallas</title>
__PLOTLY__
<style>
  :root { color-scheme: dark; }
  * { box-sizing: border-box; }
  html, body { margin: 0; background: #000; color: #ddd;
    font: 14px/1.4 system-ui, -apple-system, 'Segoe UI', Roboto, sans-serif; }
  #contenedor { background: #000; padding: 12px 16px; min-height: 100vh;
    display: flex; flex-direction: column; gap: 10px; }
  .fila { display: flex; flex-wrap: wrap; align-items: center; gap: 12px 18px; }
  label { color: #9aa0a6; font-size: 13px; }
  select { background: #111; color: #eee; border: 1px solid #333; border-radius: 8px;
    padding: 6px 10px; font: inherit; min-width: 120px; }
  .rango { display: flex; flex-direction: column; gap: 4px; flex: 1 1 360px; max-width: 640px; }
  .rango-pista { position: relative; height: 26px; }
  .rango-pista input[type=range] { position: absolute; inset: 0; width: 100%; margin: 0;
    background: none; pointer-events: none; -webkit-appearance: none; appearance: none; }
  .rango-pista input[type=range]::-webkit-slider-thumb { pointer-events: auto;
    -webkit-appearance: none; width: 16px; height: 16px; border-radius: 50%;
    background: #8a2be2; border: 2px solid #fff; cursor: pointer; }
  .rango-pista input[type=range]::-moz-range-thumb { pointer-events: auto; width: 14px;
    height: 14px; border-radius: 50%; background: #8a2be2; border: 2px solid #fff; cursor: pointer; }
  .rango-fondo { position: absolute; left: 0; right: 0; top: 11px; height: 4px;
    border-radius: 2px; background: #333; }
  .rango-activo { position: absolute; top: 11px; height: 4px; border-radius: 2px;
    background: linear-gradient(90deg, #1f4fff, #8a2be2, #ff1a1a); }
  #textoFechas { color: #eee; font-variant-numeric: tabular-nums; }
  #resumen { color: #9aa0a6; font-size: 13px; }
  .aviso-error { color: #ffb86b; border: 1px solid #5a3a1a; border-radius: 8px;
    padding: 12px 14px; max-width: 70ch; }
  #zona { position: relative; flex: 1 1 auto; display: flex; min-height: 560px; }
  #grafico { flex: 1 1 auto; min-width: 0; width: 100%; }
  #leyenda { position: absolute; top: 54px; left: 8px; width: min(300px, calc(100% - 16px));
    max-height: calc(100% - 70px); overflow: auto; background: rgba(8, 8, 14, .9);
    border: 1px solid #2c2c3a; border-radius: 10px; padding: 12px 14px; z-index: 5;
    box-shadow: 0 8px 28px rgba(0, 0, 0, .6); }
  #leyenda h3 { margin: 0 0 2px; font-size: 14px; color: #fff; text-wrap: balance; }
  #leyenda .sub { color: #9aa0a6; font-size: 12px; margin-bottom: 10px; }
  #leyenda ul { list-style: none; margin: 0; padding: 0; display: grid; gap: 5px; }
  #leyenda li { display: grid; grid-template-columns: 14px 1fr auto; gap: 8px; align-items: center;
    font-size: 13px; color: #ddd; }
  #leyenda .punto { width: 12px; height: 12px; border-radius: 50%; }
  #leyenda .n { font-variant-numeric: tabular-nums; color: #fff; font-weight: 600; }
  #leyenda .pie { display: flex; justify-content: space-between; align-items: center; gap: 8px;
    margin-top: 12px; color: #777; font-size: 11px; }
  #leyenda button { background: #222; color: #eee; border: 1px solid #444; border-radius: 6px;
    padding: 4px 10px; font: inherit; font-size: 12px; cursor: pointer; }
  #leyenda button:hover, #leyenda button:focus-visible { background: #333; outline: none; }
  .btn-fs { margin-left: auto; background: linear-gradient(135deg, #1f4fff, #8a2be2 55%, #ff1a1a);
    color: #fff; border: none; border-radius: 10px; padding: 8px 18px;
    font: 600 13px/1.2 inherit; letter-spacing: .3px; cursor: pointer;
    box-shadow: 0 4px 14px rgba(138, 43, 226, .45);
    transition: transform .15s ease, box-shadow .15s ease, filter .15s ease; }
  .btn-fs:hover { transform: translateY(-1px); filter: brightness(1.12);
    box-shadow: 0 6px 20px rgba(255, 26, 26, .45); }
  /* Fallback de alta prioridad si el navegador bloquea la API nativa */
  #contenedor.fs-fallback { position: fixed !important; inset: 0 !important;
    width: 100vw !important; height: 100vh !important; z-index: 2147483647 !important;
    overflow: auto !important; }
  #contenedor:fullscreen, #contenedor.fs-fallback { height: 100vh; }
  #contenedor:fullscreen #zona, #contenedor.fs-fallback #zona { min-height: 0; }
  @media (max-width: 640px) { #contenedor { padding: 10px 16px; } .btn-fs { margin-left: 0; } }
</style>
</head>
<body>
<div id="contenedor">
  <div class="fila">
    <label>Edificio <select id="selEdificio"></select></label>
    <label>Piso <select id="selPiso"></select></label>
    <label>Clasificación <select id="selClase"></select></label>
    <div class="rango">
      <label>Fechas: <span id="textoFechas"></span></label>
      <div class="rango-pista">
        <div class="rango-fondo"></div><div class="rango-activo" id="rangoActivo"></div>
        <input type="range" id="desde" min="0" step="1">
        <input type="range" id="hasta" min="0" step="1">
      </div>
    </div>
    <button class="btn-fs" id="btnFs">&#x26F6;&nbsp; Pantalla completa</button>
  </div>
  <div id="resumen"></div>
  <div id="zona">
    <div id="grafico"></div>
    <aside id="leyenda" hidden></aside>
  </div>
</div>
<script>
const M = __DATOS__;
const TODOS = 'Todos';
const D = M.deptos, nDeptos = D.x.length;
const f0 = new Date(M.fecha0 + 'T00:00:00');
const fmt = i => new Date(f0.getTime() + i * 864e5).toLocaleDateString('es-CL');
const $ = id => document.getElementById(id);

function tamano(d) { return d <= 1 ? 4 : d <= 5 ? 8 : d <= 10 ? 14 : d <= 15 ? 22 : 34; }
function opciones(sel, valores, actual) {
  sel.innerHTML = '';
  [TODOS, ...valores].forEach(v => sel.add(new Option(v, v)));
  sel.value = valores.map(String).includes(String(actual)) ? actual : TODOS;
}

// Controles
const edificios = [...new Set(D.edificio)].sort();
opciones($('selEdificio'), edificios, TODOS);
opciones($('selClase'), M.clases, TODOS);
$('selClase').hidden = M.clases.length < 2;
$('selClase').parentElement.hidden = M.clases.length < 2;
function actualizarPisos() {
  const e = $('selEdificio').value;
  const pisos = [...new Set(D.piso.filter((_, i) => e === TODOS || D.edificio[i] === e))]
    .sort((a, b) => a - b);
  opciones($('selPiso'), pisos, $('selPiso').value);
}
['desde', 'hasta'].forEach(id => { $(id).max = M.nDias - 1; });
$('desde').value = 0; $('hasta').value = M.nDias - 1;

// Figura con ejes congelados (sin auto-escalado)
const eje = (titulo, rango, extra) => Object.assign({
  title: { text: titulo }, range: rango, autorange: false,
  showbackground: true, backgroundcolor: '#000', gridcolor: '#2a2a2a',
  zerolinecolor: '#444', color: '#bbb'
}, extra || {});
const layout = {
  template: 'plotly_dark', paper_bgcolor: '#000', plot_bgcolor: '#000',
  font: { color: '#ddd' }, margin: { l: 0, r: 0, t: 50, b: 0 }, autosize: true,
  title: { text: 'Maqueta 3D de Fallas por Departamento', font: { color: '#eee' } },
  uirevision: 'fijo',  // conserva la cámara al filtrar
  scene: {
    xaxis: eje('X (grilla)', M.rangoX), yaxis: eje('Y (grilla)', M.rangoY),
    zaxis: eje('Piso', M.rangoZ, { dtick: 1 }),
    aspectmode: 'manual', aspectratio: M.aspecto,
    camera: { eye: { x: 1.6, y: 1.6, z: 0.9 } }, bgcolor: '#000'
  }
};

// Detalle tipo "molécula": doble clic en un depto despliega sub-burbujas por tipo
let abierto = -1;            // índice del depto desplegado (-1 = ninguno)
let puntoADepto = [];        // índice de punto de la traza principal -> índice de depto
const esc = s => String(s).replace(/[&<>"]/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[ch]);

// Posiciones repartidas sobre una esfera (espiral de Fibonacci), corregidas por la
// proporción de los ejes para que la molécula se vea redonda y no aplastada
function posicionesEsfera(n, r) {
  const kx = (M.rangoX[1] - M.rangoX[0]) / M.aspecto.x;
  const kz = (M.rangoZ[1] - M.rangoZ[0]) / M.aspecto.z;
  const rz = r * kz / kx, oro = Math.PI * (3 - Math.sqrt(5)), out = [];
  for (let k = 0; k < n; k++) {
    const yk = n === 1 ? 0 : 1 - (k / (n - 1)) * 2, rad = Math.sqrt(1 - yk * yk), th = k * oro;
    out.push([Math.cos(th) * rad * r, Math.sin(th) * rad * r, yk * rz]);
  }
  return out;
}

function trazasMolecula(i, porTipo) {
  const tipos = porTipo.map((n, t) => [t, n]).filter(([, n]) => n > 0).sort((p, q) => q[1] - p[1]);
  const maxN = tipos.length ? tipos[0][1] : 1;
  const pos = posicionesEsfera(tipos.length, 1.4);
  const cx = D.x[i], cy = D.y[i], cz = D.z[i];
  const lx = [], ly = [], lz = [], sx = [], sy = [], sz = [], sc = [], ss = [], st = [], sh = [];
  tipos.forEach(([t, n], k) => {
    const [dx, dy, dz] = pos[k];
    lx.push(cx, cx + dx, null); ly.push(cy, cy + dy, null); lz.push(cz, cz + dz, null);
    sx.push(cx + dx); sy.push(cy + dy); sz.push(cz + dz);
    sc.push(M.coloresTipo[t]); ss.push(12 + 26 * Math.sqrt(n / maxN));
    st.push(String(n)); sh.push(`<b>${esc(M.tipos[t])}</b><br>${n} OT`);
  });
  return [
    { type: 'scatter3d', mode: 'lines', x: lx, y: ly, z: lz, hoverinfo: 'skip', showlegend: false,
      line: { color: '#9a9aae', width: 4 } },
    { type: 'scatter3d', mode: 'markers', x: [cx], y: [cy], z: [cz], showlegend: false,
      text: [`<b>${esc(D.edificio[i])}</b><br>Piso ${D.piso[i]} · ${esc(D.depto[i])}`],
      hovertemplate: '%{text}<extra></extra>',
      marker: { size: 18, color: '#ffffff', opacity: 1 } },
    { type: 'scatter3d', mode: 'markers+text', x: sx, y: sy, z: sz, text: st, hovertext: sh,
      hovertemplate: '%{hovertext}<extra></extra>', showlegend: false,
      textposition: 'middle center', textfont: { color: '#000', size: 11 },
      marker: { size: ss, color: sc, opacity: 1 } },
  ];
}

function mostrarLeyenda(i, porTipo) {
  const lg = $('leyenda');
  if (i < 0) { lg.hidden = true; lg.innerHTML = ''; return; }
  const total = porTipo.reduce((p, q) => p + q, 0);
  const filas = porTipo.map((n, t) => [t, n]).filter(([, n]) => n > 0).sort((p, q) => q[1] - p[1])
    .map(([t, n]) => `<li><span class="punto" style="background:${M.coloresTipo[t]}"></span>`
      + `<span>${esc(M.tipos[t])}</span><span class="n">${n}</span></li>`).join('');
  lg.innerHTML = `<h3>${esc(D.edificio[i])}</h3>`
    + `<div class="sub">Piso ${D.piso[i]} · ${esc(D.depto[i])} · ${total} OT por tipo</div>`
    + `<ul>${filas}</ul>`
    + `<div class="pie"><span>Doble clic en otra burbuja para cambiar · Esc para cerrar</span>`
    + `<button type="button" id="cerrarLeyenda">Cerrar</button></div>`;
  lg.hidden = false;
  $('cerrarLeyenda').addEventListener('click', () => { abierto = -1; actualizar(); });
}

function actualizar() {
  let a = +$('desde').value, b = +$('hasta').value;
  if (a > b) [a, b] = [b, a];
  const pct = n => (n / Math.max(1, M.nDias - 1)) * 100;
  $('rangoActivo').style.left = pct(a) + '%';
  $('rangoActivo').style.width = (pct(b) - pct(a)) + '%';
  $('textoFechas').textContent = fmt(a) + ' – ' + fmt(b);

  const e = $('selEdificio').value, p = $('selPiso').value;
  const cl = $('selClase').value === TODOS ? -1 : M.clases.indexOf($('selClase').value);
  const visible = i => (e === TODOS || D.edificio[i] === e) && (p === TODOS || String(D.piso[i]) === p);
  const cuenta = new Int32Array(nDeptos);
  const porTipo = new Array(M.tipos.length).fill(0);
  let nOt = 0;
  for (let k = 0; k < M.ots.length; k += 4) {
    const i = M.ots[k], dia = M.ots[k + 1];
    if (dia >= a && dia <= b && (cl < 0 || M.ots[k + 2] === cl) && visible(i)) {
      cuenta[i]++; nOt++;
      if (i === abierto) porTipo[M.ots[k + 3]]++;
    }
  }
  if (abierto >= 0 && !cuenta[abierto]) abierto = -1;  // el filtro dejó fuera al depto abierto
  const x = [], y = [], z = [], c = [], s = [], t = [];
  let max = 0;
  puntoADepto = [];
  for (let i = 0; i < nDeptos; i++) {
    const d = cuenta[i];
    if (!d) continue;
    puntoADepto.push(i);
    x.push(D.x[i]); y.push(D.y[i]); z.push(D.z[i]); c.push(d); s.push(tamano(d));
    t.push(`<b>${D.edificio[i]}</b><br>Piso ${D.piso[i]} · ${esc(D.depto[i])}<br>Fallas en rango: <b>${d}</b>`);
    max = Math.max(max, d);
  }
  const trazas = [{
    type: 'scatter3d', mode: 'markers', x, y, z, text: t, showlegend: false,
    hovertemplate: '%{text}<extra></extra>',
    marker: { size: s, color: c, colorscale: M.escala, cmin: 0, cmax: M.densidadMax,
      opacity: abierto >= 0 ? 0.12 : 0.9, line: { width: 0 },  // oscurece el resto al desplegar
      colorbar: { title: { text: 'Fallas' }, thickness: 14 } }
  }];
  if (abierto >= 0) trazas.push(...trazasMolecula(abierto, porTipo));
  const planta = M.plantas[e];
  if (planta) {  // edificio con plano real: dibuja el contorno en cada piso
    const lx = [], ly = [], lz = [];
    planta.pisos.filter(z => p === TODOS || String(z) === p).forEach(z => {
      planta.contornos.forEach(c => {
        c.forEach(([px, py]) => { lx.push(px); ly.push(py); lz.push(z); });
        lx.push(null); ly.push(null); lz.push(null);
      });
    });
    trazas.push({ type: 'scatter3d', mode: 'lines', x: lx, y: ly, z: lz, hoverinfo: 'skip',
      showlegend: false, line: { color: '#3b4058', width: 2 } });
  }
  bajoCursor = -1;  // los índices de punto cambian al redibujar
  Plotly.react('grafico', trazas, layout, { responsive: true, displaylogo: false });
  enlazarClics();
  mostrarLeyenda(abierto, porTipo);
  $('resumen').textContent = `${nOt} OT · ${x.length} departamentos visibles · máx. ${max} fallas/depto`
    + ' · Doble clic en una burbuja para ver sus fallas por tipo';
}

// En 3D, plotly_click solo llega si el botón sigue apretado durante un cuadro de dibujo,
// así que un doble clic rápido casi nunca se registra. Se usa el hover (que sí es fiable)
// para saber qué depto está bajo el cursor y el evento dblclick nativo del navegador.
let bajoCursor = -1, clicsEnlazados = false;
function enlazarClics() {
  if (clicsEnlazados) return;
  clicsEnlazados = true;
  const g = $('grafico');
  g.on('plotly_hover', ev => {
    const pt = ev.points && ev.points[0];
    if (!pt) return;
    if (pt.curveNumber === 0) bajoCursor = puntoADepto[pt.pointNumber];
    else if (pt.curveNumber === 2) bajoCursor = abierto;  // el núcleo de la molécula
    else bajoCursor = -1;
  });
  g.on('plotly_unhover', () => { bajoCursor = -1; });
  g.addEventListener('dblclick', () => {
    if (bajoCursor < 0) return;
    abierto = abierto === bajoCursor ? -1 : bajoCursor;
    actualizar();
  });
}

// Redibujar el 3D es caro: se agrupan los cambios para dibujar como máximo una vez por cuadro
let redibujoPendiente = false;
function programarActualizacion() {
  if (redibujoPendiente) return;
  redibujoPendiente = true;
  requestAnimationFrame(() => { redibujoPendiente = false; actualizar(); });
}

$('selEdificio').addEventListener('change', () => { actualizarPisos(); programarActualizacion(); });
$('selPiso').addEventListener('change', programarActualizacion);
$('selClase').addEventListener('change', programarActualizacion);
$('desde').addEventListener('input', programarActualizacion);
$('hasta').addEventListener('input', programarActualizacion);

// Pantalla completa: API nativa con fallback CSS
const caja = $('contenedor'), btn = $('btnFs');
const etiqueta = on => { btn.innerHTML = on ? '&#x2715;&nbsp; Salir de pantalla completa'
                                            : '&#x26F6;&nbsp; Pantalla completa'; };
const redimensionar = () => setTimeout(() => Plotly.Plots.resize('grafico'), 120);
const enFallback = () => caja.classList.contains('fs-fallback');
function fallback(on) {
  caja.classList.toggle('fs-fallback', on);
  document.documentElement.style.overflow = on ? 'hidden' : '';
  etiqueta(on); redimensionar();
}
btn.addEventListener('click', () => {
  if (document.fullscreenElement) { document.exitFullscreen(); return; }
  if (enFallback()) { fallback(false); return; }
  if (caja.requestFullscreen && document.fullscreenEnabled) {
    caja.requestFullscreen().catch(() => fallback(true));
  } else {
    fallback(true);
  }
});
document.addEventListener('fullscreenchange', () => {
  etiqueta(!!document.fullscreenElement); redimensionar();
});
document.addEventListener('keydown', ev => {
  if (ev.key !== 'Escape') return;
  if (abierto >= 0) { abierto = -1; actualizar(); return; }
  if (enFallback()) fallback(false);
});

// Sin Plotly o sin WebGL el gráfico 3D no puede dibujarse: avisar en vez de quedar en blanco
function webglDisponible() {
  try {
    const c = document.createElement('canvas');
    return !!(window.WebGLRenderingContext && (c.getContext('webgl') || c.getContext('experimental-webgl')));
  } catch (e) { return false; }
}
function mostrarError(msg) {
  $('grafico').innerHTML = '<p class="aviso-error">' + msg + '</p>';
}
actualizarPisos();
if (typeof Plotly === 'undefined') {
  mostrarError('No se pudo cargar la librería de gráficos (Plotly). Revisa tu conexión '
    + 'o genera el archivo sin la opción --cdn.');
} else if (!webglDisponible()) {
  mostrarError('Tu navegador no tiene WebGL activo, y los gráficos 3D lo necesitan. Activa la '
    + 'aceleración por hardware en la configuración del navegador o prueba con otro navegador.');
} else {
  try { actualizar(); } catch (e) { mostrarError('Error al dibujar el gráfico: ' + e.message); }
}
</script>
</body>
</html>
"""


def generar_html(modelo, salida, usar_cdn=False):
    if usar_cdn:
        plotly_tag = '<script src="https://cdnjs.cloudflare.com/ajax/libs/plotly.js/2.35.2/plotly.min.js" charset="utf-8"></script>'
    else:
        from plotly.offline import get_plotlyjs
        plotly_tag = f"<script>{get_plotlyjs()}</script>"
    datos = json.dumps(modelo, ensure_ascii=False).replace("</", "<\\/")
    html = PLANTILLA.replace("__PLOTLY__", plotly_tag).replace("__DATOS__", datos)
    Path(salida).write_text(html, encoding="utf-8")
    return Path(salida).resolve()


# ---------------------------------------------------------------------------
# 5. EJECUCIÓN
# ---------------------------------------------------------------------------
def main(argv=None):
    ap = argparse.ArgumentParser(description="Genera la maqueta 3D de OT como HTML autónomo.")
    ap.add_argument("excel", nargs="?", default="OT_CLASIFICADAS_ARBOL (1).xlsx",
                    help="Ruta del Excel de órdenes de trabajo")
    ap.add_argument("-o", "--salida", default="maqueta_ot_3d.html", help="Archivo HTML de salida")
    ap.add_argument("--cdn", action="store_true",
                    help="Cargar Plotly.js desde internet en vez de incrustarlo")
    ap.add_argument("--no-abrir", action="store_true", help="No abrir el navegador al terminar")
    args = ap.parse_args(argv)

    df = cargar_datos(args.excel)
    if df.empty:
        sys.exit("No quedaron registros válidos para graficar.")
    ruta = generar_html(construir_modelo(df), args.salida, usar_cdn=args.cdn)
    print(f"HTML generado: {ruta}")
    if not args.no_abrir:
        webbrowser.open(ruta.as_uri())


if __name__ == "__main__":
    main()

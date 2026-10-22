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
    # Somma Plaza Ñuñoa, láminas IRZ-AR-PL-06 a 19 (pisos 2 a 25, escala 1:100), todas en
    # las mismas coordenadas. Barra con pasillo central: torre B (deptos 1–14) al poniente
    # y torre A al oriente; en los pisos 3 a 7 la torre A suma un ala perpendicular
    # (deptos 7–17). Las etiquetas de la lámina ("301A", "301B") están junto al pasillo;
    # cada punto se llevó al centro de su lado. La distribución cambia según el piso.
    # En PL-08 (pisos 4–7) el rótulo 416A está en el margen; se usa la posición del piso 3.
    "Somma Plaza Ñuñoa": {
        "escala": 3 / 40.1,
        "unir_torres": True,
        # Respaldo para pisos fuera de los rangos de abajo: la planta del piso 3
        "deptos": {
            ("A", 1): (8.0, -5.4), ("A", 2): (9.0, 5.6), ("A", 3): (12.8, -4.6),
            ("A", 4): (13.7, 4.3), ("A", 5): (21.2, -4.5), ("A", 6): (18.9, 4.3),
            ("A", 7): (25.4, -7.3), ("A", 8): (25.3, -14.7), ("A", 9): (36.5, -14.5),
            ("A", 10): (36.6, -9.1), ("A", 11): (36.6, -4.1), ("A", 12): (36.6, 1.0),
            ("A", 13): (36.6, 5.9), ("A", 14): (36.6, 10.9), ("A", 15): (36.6, 13.7),
            ("A", 16): (25.2, 13.9), ("A", 17): (25.2, 6.7), ("B", 1): (-0.7, -5.4),
            ("B", 2): (-5.2, 4.3), ("B", 3): (-10.9, -4.5), ("B", 4): (-11.3, 4.3),
            ("B", 5): (-16.5, -4.5), ("B", 6): (-16.5, 4.3), ("B", 7): (-21.7, -4.5),
            ("B", 8): (-21.7, 4.3), ("B", 9): (-27.5, -4.5), ("B", 10): (-26.9, 4.3),
            ("B", 11): (-32.1, -4.5), ("B", 12): (-32.1, 4.3), ("B", 13): (-37.1, -5.4),
            ("B", 14): (-37.1, 5.4),
        },
        # Planta propia por rango de pisos (láminas PL-06 a PL-19; el piso 1 usa la del piso 2)
        "deptos_por_piso": [
            ((1, 2), {
                ("A", 1): (8.1, -5.4), ("A", 2): (9.0, 5.6), ("A", 3): (13.7, 4.3),
                ("A", 4): (18.9, 4.3), ("B", 1): (-0.7, -5.8), ("B", 2): (-5.2, 4.3),
                ("B", 3): (-10.9, -4.5), ("B", 4): (-11.3, 4.3), ("B", 5): (-16.5, -4.5),
                ("B", 6): (-16.5, 4.3), ("B", 7): (-21.7, -4.5), ("B", 8): (-21.7, 4.3),
                ("B", 9): (-27.5, -4.5), ("B", 10): (-26.9, 4.3), ("B", 11): (-32.1, -4.5),
                ("B", 12): (-32.1, 4.3), ("B", 13): (-37.0, -5.4), ("B", 14): (-37.1, 5.3),
            }),
            ((3, 3), {
                ("A", 1): (8.0, -5.4), ("A", 2): (9.0, 5.6), ("A", 3): (12.8, -4.6),
                ("A", 4): (13.7, 4.3), ("A", 5): (21.2, -4.5), ("A", 6): (18.9, 4.3),
                ("A", 7): (25.4, -7.3), ("A", 8): (25.3, -14.7), ("A", 9): (36.5, -14.5),
                ("A", 10): (36.6, -9.1), ("A", 11): (36.6, -4.1), ("A", 12): (36.6, 1.0),
                ("A", 13): (36.6, 5.9), ("A", 14): (36.6, 10.9), ("A", 15): (36.6, 13.7),
                ("A", 16): (25.2, 13.9), ("A", 17): (25.2, 6.7), ("B", 1): (-0.7, -5.4),
                ("B", 2): (-5.2, 4.3), ("B", 3): (-10.9, -4.5), ("B", 4): (-11.3, 4.3),
                ("B", 5): (-16.5, -4.5), ("B", 6): (-16.5, 4.3), ("B", 7): (-21.7, -4.5),
                ("B", 8): (-21.7, 4.3), ("B", 9): (-27.5, -4.5), ("B", 10): (-26.9, 4.3),
                ("B", 11): (-32.1, -4.5), ("B", 12): (-32.1, 4.3), ("B", 13): (-37.1, -5.4),
                ("B", 14): (-37.1, 5.4),
            }),
            ((4, 7), {
                ("A", 1): (7.6, -5.8), ("A", 2): (9.0, 4.3), ("A", 3): (12.5, -4.5),
                ("A", 4): (13.4, 4.3), ("A", 5): (20.9, -4.5), ("A", 6): (18.6, 4.3),
                ("A", 7): (24.8, -7.3), ("A", 8): (25.0, -14.7), ("A", 9): (36.3, -14.5),
                ("A", 10): (36.4, -9.0), ("A", 11): (36.4, -4.0), ("A", 12): (36.5, 1.0),
                ("A", 13): (36.5, 6.0), ("A", 14): (36.5, 11.0), ("A", 15): (36.3, 13.7),
                ("A", 16): (25.2, 13.9), ("A", 17): (24.8, 6.7), ("B", 1): (-0.9, -5.9),
                ("B", 2): (-5.4, 4.3), ("B", 3): (-11.1, -4.5), ("B", 4): (-11.5, 4.3),
                ("B", 5): (-16.7, -4.5), ("B", 6): (-16.7, 4.3), ("B", 7): (-21.9, -4.5),
                ("B", 8): (-21.9, 4.3), ("B", 9): (-27.7, -4.5), ("B", 10): (-27.1, 4.3),
                ("B", 11): (-32.3, -4.5), ("B", 12): (-32.3, 4.3), ("B", 13): (-37.3, -5.4),
                ("B", 14): (-37.4, 5.4),
            }),
            ((8, 8), {
                ("A", 1): (7.9, -5.8), ("A", 2): (9.3, 4.3), ("A", 3): (12.8, -4.5),
                ("A", 4): (13.7, 4.3), ("A", 5): (21.2, -4.5), ("A", 7): (26.8, -4.5),
                ("A", 8): (28.8, 5.6), ("B", 1): (-0.7, -5.9), ("B", 2): (-5.2, 4.3),
                ("B", 3): (-10.9, -4.5), ("B", 4): (-11.3, 4.3), ("B", 5): (-16.5, -4.5),
                ("B", 6): (-16.5, 4.3), ("B", 7): (-21.7, -4.5), ("B", 8): (-21.7, 4.3),
                ("B", 9): (-27.5, -4.5), ("B", 10): (-26.9, 4.3), ("B", 11): (-32.1, -4.5),
                ("B", 12): (-32.1, 4.3), ("B", 13): (-36.9, -5.2), ("B", 14): (-36.9, 5.0),
            }),
            ((9, 9), {
                ("A", 1): (7.9, -5.8), ("A", 2): (9.3, 4.3), ("A", 3): (12.8, -4.5),
                ("A", 4): (13.7, 4.3), ("A", 5): (21.2, -4.5), ("A", 7): (26.8, -4.5),
                ("A", 8): (28.8, 5.1), ("B", 1): (-0.7, -5.9), ("B", 2): (-5.2, 4.3),
                ("B", 3): (-10.9, -4.5), ("B", 4): (-11.3, 4.3), ("B", 5): (-16.5, -4.5),
                ("B", 6): (-16.5, 4.3), ("B", 7): (-21.7, -4.5), ("B", 8): (-21.7, 4.3),
                ("B", 9): (-27.5, -4.5), ("B", 10): (-26.9, 4.3), ("B", 11): (-32.2, -5.9),
                ("B", 12): (-32.1, 5.6),
            }),
            ((10, 12), {
                ("A", 1): (7.5, -5.8), ("A", 2): (9.0, 4.3), ("A", 3): (12.5, -4.5),
                ("A", 4): (13.4, 4.3), ("A", 5): (20.9, -4.5), ("A", 6): (18.6, 4.3),
                ("A", 7): (26.5, -4.5), ("A", 8): (28.6, 5.2), ("B", 1): (-0.9, -5.9),
                ("B", 2): (-5.5, 4.3), ("B", 3): (-11.2, -4.5), ("B", 4): (-11.6, 4.3),
                ("B", 5): (-16.8, -4.5), ("B", 6): (-16.8, 4.3), ("B", 7): (-22.0, -4.5),
                ("B", 8): (-22.0, 4.3), ("B", 9): (-27.8, -4.5), ("B", 10): (-27.2, 4.3),
                ("B", 11): (-32.5, -5.9), ("B", 12): (-32.5, 5.6),
            }),
            ((13, 13), {
                ("A", 1): (7.8, -5.8), ("A", 2): (9.3, 4.3), ("A", 3): (12.8, -4.5),
                ("A", 4): (13.7, 4.3), ("A", 5): (21.2, -4.5), ("A", 6): (18.9, 4.3),
                ("A", 7): (26.8, -4.5), ("A", 8): (28.9, 5.1), ("B", 1): (-0.6, -5.9),
                ("B", 2): (-5.2, 4.3), ("B", 3): (-10.9, -4.5), ("B", 4): (-11.3, 4.3),
                ("B", 5): (-16.5, -4.5), ("B", 6): (-16.5, 4.3), ("B", 7): (-21.7, -4.5),
                ("B", 8): (-21.7, 4.3), ("B", 9): (-27.1, -5.9), ("B", 10): (-27.1, 5.6),
            }),
            ((14, 15), {
                ("A", 1): (7.5, -5.8), ("A", 2): (9.0, 4.3), ("A", 3): (12.5, -4.5),
                ("A", 4): (13.4, 4.3), ("A", 5): (20.9, -4.5), ("A", 6): (18.6, 4.3),
                ("A", 7): (26.5, -4.5), ("A", 8): (28.6, 5.2), ("B", 1): (-0.9, -5.9),
                ("B", 2): (-5.5, 4.3), ("B", 3): (-11.2, -4.5), ("B", 4): (-11.6, 4.3),
                ("B", 5): (-16.8, -4.5), ("B", 6): (-16.8, 4.3), ("B", 7): (-22.0, -4.5),
                ("B", 8): (-22.0, 4.3), ("B", 9): (-27.4, -5.9), ("B", 10): (-27.4, 5.6),
            }),
            ((16, 16), {
                ("A", 1): (7.8, -5.8), ("A", 2): (9.3, 4.3), ("A", 3): (12.8, -4.5),
                ("A", 4): (13.7, 4.3), ("A", 5): (21.2, -4.5), ("A", 6): (18.9, 4.3),
                ("A", 7): (26.8, -4.5), ("A", 8): (28.9, 5.1), ("B", 1): (-0.6, -5.9),
                ("B", 2): (-5.2, 4.3), ("B", 3): (-10.9, -4.5), ("B", 4): (-11.3, 4.3),
                ("B", 5): (-16.5, -4.5), ("B", 6): (-16.5, 4.3), ("B", 7): (-21.9, -5.9),
                ("B", 8): (-22.0, 5.6),
            }),
            ((17, 18), {
                ("A", 1): (7.5, -5.8), ("A", 2): (9.0, 4.3), ("A", 3): (12.5, -4.5),
                ("A", 4): (13.4, 4.3), ("A", 5): (20.9, -4.5), ("A", 6): (18.6, 4.3),
                ("A", 7): (26.5, -4.5), ("A", 8): (28.6, 5.2), ("B", 1): (-0.9, -5.9),
                ("B", 2): (-5.5, 4.3), ("B", 3): (-11.2, -4.5), ("B", 4): (-11.6, 4.3),
                ("B", 5): (-16.8, -4.5), ("B", 6): (-16.8, 4.3), ("B", 7): (-22.2, -5.8),
                ("B", 8): (-22.4, 5.6),
            }),
            ((19, 19), {
                ("A", 1): (7.8, -5.8), ("A", 2): (9.3, 4.3), ("A", 3): (12.8, -4.5),
                ("A", 4): (13.7, 4.3), ("A", 5): (21.2, -4.5), ("A", 6): (18.9, 4.3),
                ("A", 7): (26.8, -4.5), ("A", 8): (28.9, 5.1), ("B", 1): (-0.6, -5.9),
                ("B", 2): (-5.2, 4.3), ("B", 3): (-10.9, -4.5), ("B", 4): (-11.3, 4.3),
                ("B", 5): (-16.7, -5.9), ("B", 6): (-16.9, 5.6),
            }),
            ((20, 21), {
                ("A", 1): (7.5, -5.8), ("A", 2): (9.0, 4.3), ("A", 3): (12.5, -4.5),
                ("A", 4): (13.4, 4.3), ("A", 5): (20.9, -4.5), ("A", 6): (18.6, 4.3),
                ("A", 7): (26.5, -4.5), ("A", 8): (28.6, 5.2), ("B", 1): (-0.9, -5.9),
                ("B", 2): (-5.5, 4.3), ("B", 3): (-11.2, -4.5), ("B", 4): (-11.6, 4.3),
                ("B", 5): (-17.0, -5.9), ("B", 6): (-17.1, 5.6),
            }),
            ((22, 22), {
                ("A", 1): (7.8, -5.8), ("A", 2): (9.3, 4.3), ("A", 3): (12.8, -4.5),
                ("A", 4): (13.7, 4.3), ("A", 5): (21.2, -4.5), ("A", 6): (18.9, 4.3),
                ("A", 7): (26.8, -4.5), ("A", 8): (28.9, 5.1), ("B", 1): (-0.6, -5.9),
                ("B", 2): (-5.2, 4.3), ("B", 3): (-11.3, -5.9), ("B", 4): (-11.6, 5.6),
            }),
            ((23, 25), {
                ("A", 1): (7.5, -5.8), ("A", 2): (9.0, 4.3), ("A", 3): (12.5, -4.5),
                ("A", 4): (13.4, 4.3), ("A", 5): (20.9, -4.5), ("A", 6): (18.6, 4.3),
                ("A", 7): (26.5, -4.5), ("A", 8): (28.6, 5.2), ("B", 1): (-0.9, -5.9),
                ("B", 2): (-5.5, 4.3), ("B", 3): (-11.5, -5.9), ("B", 4): (-11.8, 5.6),
            }),
        ],
        # Contornos: la barra se escalona en altura (su borde poniente, medido en cada
        # lámina, retrocede ~27 m entre el piso 2 y el 25), cada una en su rango de pisos;
        # el último es el ala de la torre A, solo en los pisos 3 a 7.
        "pisos_contorno": [(1, 2), (3, 7), (8, 8), (9, 9), (10, 12), (13, 13), (14, 15), (16, 16), (17, 18), (19, 19), (20, 21), (22, 22), (23, 25), (3, 7)],
        "contornos": [
            [(-40.7, 7.9), (30.3, 7.9), (30.3, -8.1), (-40.7, -8.1), (-40.7, 7.9)],  # pisos 1–2
            [(-40.7, 7.9), (22.8, 7.9), (22.8, -8.1), (-40.7, -8.1), (-40.7, 7.9)],  # pisos 3–7
            [(-40.7, 7.9), (30.3, 7.9), (30.3, -8.1), (-40.7, -8.1), (-40.7, 7.9)],  # piso 8
            [(-40.3, 7.9), (30.3, 7.9), (30.3, -8.1), (-40.3, -8.1), (-40.3, 7.9)],  # piso 9
            [(-35.3, 7.9), (30.3, 7.9), (30.3, -8.1), (-35.3, -8.1), (-35.3, 7.9)],  # pisos 10–12
            [(-34.2, 7.9), (30.3, 7.9), (30.3, -8.1), (-34.2, -8.1), (-34.2, 7.9)],  # piso 13
            [(-30.1, 7.9), (30.3, 7.9), (30.3, -8.1), (-30.1, -8.1), (-30.1, 7.9)],  # pisos 14–15
            [(-29.1, 7.9), (30.3, 7.9), (30.3, -8.1), (-29.1, -8.1), (-29.1, 7.9)],  # piso 16
            [(-24.9, 7.9), (30.3, 7.9), (30.3, -8.1), (-24.9, -8.1), (-24.9, 7.9)],  # pisos 17–18
            [(-23.8, 7.9), (30.3, 7.9), (30.3, -8.1), (-23.8, -8.1), (-23.8, 7.9)],  # piso 19
            [(-19.6, 7.9), (30.3, 7.9), (30.3, -8.1), (-19.6, -8.1), (-19.6, 7.9)],  # pisos 20–21
            [(-18.6, 7.9), (30.3, 7.9), (30.3, -8.1), (-18.6, -8.1), (-18.6, 7.9)],  # piso 22
            [(-13.4, 7.9), (30.3, 7.9), (30.3, -8.1), (-13.4, -8.1), (-13.4, 7.9)],  # pisos 23–25
            [(22.8, 17.8), (40.1, 17.8), (40.1, -17.8), (22.8, -17.8), (22.8, 17.8)],  # ala torre A
        ],
    },
    # Somma Vista Calan, láminas 2101-A-105 a A-109 (escala 1:75): una barra con pasillo
    # central dividida al medio en "Edificio S" (poniente) y "Edificio N" (oriente), en
    # espejo. Se asume SVCA = Edificio S y SVCB = Edificio N (la lámina no lo indica).
    # Los pisos 2 a 9 tienen la misma planta (A-106, pisos 3 al 7); el piso 10 es el
    # "piso retirado" (A-109), más angosto y con otra numeración, alineado con las demás
    # por el rótulo de la línea divisoria. Las etiquetas están junto al pasillo; cada
    # punto se llevó al centro de su fila de departamentos.
    "Somma Vista Calan": {
        "escala": 3 / 34.3,
        "unir_torres": True,
        "deptos": {
            ("SVCA", 1): (-6.5, 5.3), ("SVCA", 2): (-10.2, 5.3), ("SVCA", 3): (-13.9, 5.3),
            ("SVCA", 4): (-16.9, 5.3), ("SVCA", 5): (-22.9, 5.3), ("SVCA", 6): (-28.1, 0.1),
            ("SVCA", 7): (-26.9, -4.2), ("SVCA", 8): (-22.8, -4.2), ("SVCA", 9): (-11.7, -4.2),
            ("SVCA", 10): (-3.0, -4.2), ("SVCB", 1): (6.2, 5.3), ("SVCB", 2): (9.9, 5.3),
            ("SVCB", 3): (13.6, 5.3), ("SVCB", 4): (16.7, 5.3), ("SVCB", 5): (22.6, 5.3),
            ("SVCB", 6): (27.8, 0.1), ("SVCB", 7): (26.6, -4.2), ("SVCB", 8): (22.5, -4.2),
            ("SVCB", 9): (11.3, -4.2), ("SVCB", 10): (2.7, -4.2),
        },
        "deptos_por_piso": [
            ((10, 10), {
                ("SVCA", 1): (-7.4, 4.0), ("SVCA", 2): (-21.6, -0.1), ("SVCA", 3): (-20.4, -3.3),
                ("SVCA", 4): (-11.5, -3.3), ("SVCA", 5): (-2.9, -3.3), ("SVCB", 1): (7.4, 4.0),
                ("SVCB", 2): (20.3, 4.0), ("SVCB", 3): (23.0, -0.0), ("SVCB", 4): (21.7, -3.3),
                ("SVCB", 5): (11.5, -3.3), ("SVCB", 6): (2.9, -3.3),
            }),
        ],
        # Contorno: la barra en los pisos 2 a 9 y el piso retirado (10), más angosto;
        # la planta es en espejo, así que el contorno es simétrico respecto de la división
        "pisos_contorno": [(2, 9), (10, 10)],
        "contornos": [
            [(-34.3, 9.9), (34.3, 9.9), (34.3, -9.9), (-34.3, -9.9), (-34.3, 9.9)],
            [(-31.4, 7.2), (31.4, 7.2), (31.4, -6.2), (-31.4, -6.2), (-31.4, 7.2)],
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
    posiciones = plano and plano["deptos"]
    piso = grupo["Piso_Num"].iloc[0]
    for (desde, hasta), por_piso in (plano or {}).get("deptos_por_piso", []):
        if desde <= piso <= hasta:  # planta propia de este rango de pisos
            posiciones = por_piso
    if plano and all(k in posiciones for k in claves):
        grupo["X"] = [posiciones[k][0] * plano["escala"] for k in claves]
        grupo["Y"] = [posiciones[k][1] * plano["escala"] for k in claves]
        grupo["Grid_Size"] = 0  # posición según plano, no grilla
        return grupo
    grid_size = tamano_grilla(len(grupo))
    offset = (grid_size - 1) / 2.0           # centrado matemático en (0,0)
    idx = np.arange(len(grupo))
    grupo["X"] = (idx % grid_size) - offset
    grupo["Y"] = (idx // grid_size) - offset
    grupo["Grid_Size"] = grid_size
    return grupo


def pisos_contorno(deptos, edificio, plano, k):
    """Pisos donde dibujar el contorno k: los de su torre y rango de pisos, si el plano los indica."""
    sel = deptos["Edificio"] == edificio
    torres = plano.get("torres_contorno")
    if torres and torres[k]:
        sel &= deptos["Torre"] == torres[k]
    rango = (plano.get("pisos_contorno") or [None] * len(plano["contornos"]))[k]
    if rango:
        sel &= deptos["Z"].between(*rango)
    return sorted(int(z) for z in deptos.loc[sel, "Z"].unique())


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
            ed: [
                {"puntos": [[x * p["escala"], y * p["escala"]] for x, y in c],
                 "pisos": pisos_contorno(deptos, ed, p, k)}
                for k, c in enumerate(p["contornos"])
            ]
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
  [hidden] { display: none !important; }
  .tl { gap: 10px 12px; }
  .tl button { background: #1a1a24; color: #eee; border: 1px solid #3a3a50; border-radius: 8px;
    padding: 6px 12px; font: inherit; font-size: 13px; cursor: pointer; }
  .tl button:hover, .tl button:focus-visible { background: #262636; outline: none; }
  #btnTl { border-color: #8a2be2; }
  #ctlTl { display: flex; flex-wrap: wrap; align-items: center; gap: 10px 12px; flex: 1 1 auto; }
  #tlDia { flex: 1 1 220px; min-width: 0; max-width: 520px; accent-color: #8a2be2; }
  #tlTexto { color: #eee; font-variant-numeric: tabular-nums; font-size: 13px; }
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
  <div class="fila tl">
    <button type="button" id="btnTl">&#x25B6;&nbsp; Time lapse (30 días por falla)</button>
    <div id="ctlTl" hidden>
      <button type="button" id="tlPlay">&#x23F8;&nbsp; Pausa</button>
      <button type="button" id="tlInicio" title="Volver al inicio">&#x23EE;</button>
      <label>Velocidad <select id="tlVel">
        <option value="5">5 días/s</option><option value="10" selected>10 días/s</option>
        <option value="20">20 días/s</option><option value="40">40 días/s</option>
      </select></label>
      <input type="range" id="tlDia" min="0" step="1" aria-label="Fecha del time lapse">
      <span id="tlTexto"></span>
      <button type="button" id="tlSalir">Salir del time lapse</button>
    </div>
  </div>
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

// Time lapse: cada OT queda activa VIDA_OT días desde su creación y luego se apaga
const VIDA_OT = 30;
const tl = { activo: false, corriendo: false, dia: 0, acum: 0, timer: null };

// Máximo de OT activas a la vez en un mismo depto durante [a, b]: fija la escala de
// tamaño y color del time lapse para que no salte mientras se reproduce
function maxActivas(a, b, pasa) {
  const llegan = new Map();
  for (let k = 0; k < M.ots.length; k += 4) {
    const dia = M.ots[k + 1];
    if (dia < a || dia > b || !pasa(k)) continue;
    if (!llegan.has(dia)) llegan.set(dia, []);
    llegan.get(dia).push(M.ots[k]);
  }
  const n = new Int32Array(nDeptos);
  let max = 1;
  for (let d = a; d <= b; d++) {
    (llegan.get(d) || []).forEach(i => { if (++n[i] > max) max = n[i]; });
    (llegan.get(d - VIDA_OT) || []).forEach(i => { n[i]--; });
  }
  return max;
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
  const pasa = k => (cl < 0 || M.ots[k + 2] === cl) && visible(M.ots[k]);
  // Ventana de días a contar: todo el rango, o en el time lapse las OT aún activas
  let lo = a, hi = b, maxTl = 1;
  if (tl.activo) {
    tl.dia = Math.min(Math.max(tl.dia, a), b);
    lo = Math.max(a, tl.dia - VIDA_OT + 1); hi = tl.dia;
    maxTl = maxActivas(a, b, pasa);
    $('tlDia').min = a; $('tlDia').max = b; $('tlDia').value = tl.dia;
  }
  const cuenta = new Int32Array(nDeptos);
  const porTipo = new Array(M.tipos.length).fill(0);
  let nOt = 0;
  for (let k = 0; k < M.ots.length; k += 4) {
    const i = M.ots[k], dia = M.ots[k + 1];
    if (dia >= lo && dia <= hi && pasa(k)) {
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
    x.push(D.x[i]); y.push(D.y[i]); z.push(D.z[i]); c.push(d);
    s.push(tl.activo ? 6 + 26 * Math.sqrt(d / maxTl) : tamano(d));
    t.push(`<b>${D.edificio[i]}</b><br>Piso ${D.piso[i]} · ${esc(D.depto[i])}<br>`
      + (tl.activo ? `OT activas (últimos ${VIDA_OT} días): <b>${d}</b>` : `Fallas en rango: <b>${d}</b>`));
    max = Math.max(max, d);
  }
  const trazas = [{
    type: 'scatter3d', mode: 'markers', x, y, z, text: t, showlegend: false,
    hovertemplate: '%{text}<extra></extra>',
    marker: { size: s, color: c, colorscale: M.escala, cmin: 0, cmax: tl.activo ? maxTl : M.densidadMax,
      opacity: abierto >= 0 ? 0.12 : 0.9, line: { width: 0 },  // oscurece el resto al desplegar
      colorbar: { title: { text: tl.activo ? 'OT activas' : 'Fallas' }, thickness: 14 } }
  }];
  if (abierto >= 0) trazas.push(...trazasMolecula(abierto, porTipo));
  const planta = M.plantas[e];
  if (planta) {  // edificio con plano real: dibuja el contorno en cada piso
    const lx = [], ly = [], lz = [];
    planta.forEach(c => {
      c.pisos.filter(z => p === TODOS || String(z) === p).forEach(z => {
        c.puntos.forEach(([px, py]) => { lx.push(px); ly.push(py); lz.push(z); });
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
  if (tl.activo) {
    $('tlTexto').textContent = `${fmt(tl.dia)} · ${nOt} OT activas en ${x.length} deptos`;
    $('resumen').textContent = `Time lapse: cada OT se muestra ${VIDA_OT} días desde su creación y luego se apaga`
      + ` · escala fija hasta ${maxTl} OT activas por depto · Espacio para pausar`;
  } else {
    $('resumen').textContent = `${nOt} OT · ${x.length} departamentos visibles · máx. ${max} fallas/depto`
      + ' · Doble clic en una burbuja para ver sus fallas por tipo';
  }
}

function rangoFechas() {
  const a = +$('desde').value, b = +$('hasta').value;
  return a <= b ? [a, b] : [b, a];
}
function tlReproducir(on) {
  const [a, b] = rangoFechas();
  clearInterval(tl.timer);
  tl.corriendo = on;
  if (on) {
    if (tl.dia >= b) tl.dia = a;  // al terminar, reproducir de nuevo parte desde el inicio
    tl.acum = 0;
    tl.timer = setInterval(tlPaso, 100);
  }
  $('tlPlay').innerHTML = on ? '&#x23F8;&nbsp; Pausa'
    : (tl.dia >= b ? '&#x21BB;&nbsp; Repetir' : '&#x25B6;&nbsp; Reproducir');
}
function tlPaso() {
  const b = rangoFechas()[1];
  tl.acum += +$('tlVel').value / 10;  // 10 pasos por segundo
  const n = Math.floor(tl.acum);
  if (!n) return;
  tl.acum -= n;
  tl.dia = Math.min(b, tl.dia + n);
  programarActualizacion();
  if (tl.dia >= b) tlReproducir(false);
}
function tlIniciar() {
  tl.activo = true;
  tl.dia = rangoFechas()[0];
  $('btnTl').hidden = true; $('ctlTl').hidden = false;
  programarActualizacion();
  tlReproducir(true);
}
function tlSalir() {
  tlReproducir(false);
  tl.activo = false;
  $('btnTl').hidden = false; $('ctlTl').hidden = true;
  programarActualizacion();
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
$('btnTl').addEventListener('click', tlIniciar);
$('tlSalir').addEventListener('click', tlSalir);
$('tlPlay').addEventListener('click', () => tlReproducir(!tl.corriendo));
// Saltar de fecha no reinicia la reproducción; en pausa solo actualiza el texto del botón
function tlIrA(dia) {
  tl.dia = dia; tl.acum = 0;
  if (!tl.corriendo) tlReproducir(false);
  programarActualizacion();
}
$('tlInicio').addEventListener('click', () => tlIrA(rangoFechas()[0]));
$('tlDia').addEventListener('input', () => tlIrA(+$('tlDia').value));

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
  // Espacio pausa o reanuda el time lapse (salvo escribiendo en un control)
  if (ev.key === ' ' && tl.activo && !/^(SELECT|INPUT|BUTTON)$/.test(document.activeElement.tagName)) {
    ev.preventDefault(); tlReproducir(!tl.corriendo); return;
  }
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

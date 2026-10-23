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
MARGEN_EJES = 5.0          # margen estático de seguridad en planta (m)
ALTURA_PISO = 2.5          # altura de piso a piso (m): la maqueta va a escala real
MODULO_GRILLA = 6.0        # separación (m) entre deptos en la grilla de edificios sin plano
GRID_MIN, GRID_MAX = 3, 7  # grilla dinámica 3x3 … 7x7

# Posiciones reales según planos de arquitectura (metros desde el centro de la planta,
# X hacia la derecha y Y hacia arriba de la lámina). Los edificios que no aparecen aquí
# usan la grilla cuadrada genérica.
PLANTAS = {
    # Somma Asturias, láminas MP-PG-110 (piso 3), 111 (pisos 4–15), 112 (pisos 16–18) y
    # 113 (piso 19): 12 deptos por piso en anillo alrededor del patio central, con la escalera
    # al poniente y los ascensores al oriente. El piso 19 solo tiene el bloque central (deptos
    # 01, 02, 11 y 12); en las alas quedan la piscina y la sala de calderas sobre la cubierta.
    # Los pisos 1 y 2 son el zócalo comercial (sin deptos). Medidas tomadas de la planta tipo
    # escaneada, calibrada con las cotas totales (33,57 m de ancho y 31,00 m de fondo).
    "Somma Asturias": {
        "deptos": {
            9: (-13.2, 10.0), 11: (-4.8, 10.5), 12: (4.8, 10.5), 10: (13.2, 10.0),
            7: (-12.4, 2.85), 8: (12.4, 2.85),
            5: (-12.4, -2.85), 6: (12.4, -2.85),
            3: (-13.2, -10.0), 1: (-4.8, -10.5), 2: (4.8, -10.5), 4: (13.2, -10.0),
        },
        "deptos_por_piso": [
            ((19, 19), {11: (-4.8, 10.5), 12: (4.8, 10.5), 1: (-4.8, -10.5), 2: (4.8, -10.5)}),
        ],
        # Muros exteriores: la planta completa (pisos 1–18) y el bloque central del piso 19
        "contornos": [
            [(-15.12, 13.75), (15.12, 13.75), (15.12, -13.7), (-15.12, -13.7), (-15.12, 13.75)],
            [(-9.63, 13.75), (9.63, 13.75), (9.63, -13.7), (-9.63, -13.7), (-9.63, 13.75)],
        ],
        "pisos_contorno": [(1, 18), (19, 19)],
        # Volúmenes: alas hasta la cubierta del piso 18 (+45,0) y bloque central hasta el 19 (+47,5)
        "volumenes": [(0, 0.0, 45.0), (1, 45.0, 47.5)],
        "areas_comunes": [
            ("Corredor", (-9.63, 5.8, 9.63, 7.53), 0.0, 47.5),
            ("Corredor", (-9.63, -7.55, 9.63, -5.91), 0.0, 47.5),
            ("Corredor", (6.12, -5.91, 7.92, 5.8), 0.0, 47.5),
            ("Escaleras", (-7.83, -5.91, -3.6, 5.8), 0.0, 47.5),
            ("Ascensores", (3.6, -3.23, 6.12, 5.8), 0.0, 47.5),
            ("Sala de máquinas", (3.6, -3.23, 6.12, 5.8), 47.5, 50.5),
        ],
        # Terrazas: la franja norte y sur del bloque central (1,8 m) y la columna de terrazas de
        # cada ala (poniente y oriente, solo pisos 3–18), con sus retornos hasta el muro
        "terrazas": {
            "Asturias": [
                ((-9.63, 15.5), (9.63, 15.5)), ((-9.63, 15.5), (-9.63, 13.75)),
                ((9.63, 15.5), (9.63, 13.75)),
                ((-9.63, -15.5), (9.63, -15.5)), ((-9.63, -15.5), (-9.63, -13.7)),
                ((9.63, -15.5), (9.63, -13.7)),
                ((-16.85, 12.95), (-16.85, -12.95), (3, 18)), ((-16.85, 12.95), (-15.12, 12.95), (3, 18)),
                ((-16.85, -12.95), (-15.12, -12.95), (3, 18)),
                ((16.85, 12.95), (16.85, -12.95), (3, 18)), ((15.12, 12.95), (16.85, 12.95), (3, 18)),
                ((15.12, -12.95), (16.85, -12.95), (3, 18)),
            ],
        },
        # Muros entre deptos (las alas solo hasta el piso 18)
        "particiones": {
            "Asturias": [
                ((-16.85, 6.25), (-9.63, 6.25), (3, 18)), ((-16.85, 0.0), (-7.92, 0.0), (3, 18)),
                ((-16.85, -6.2), (-9.63, -6.2), (3, 18)),
                ((9.63, 6.25), (16.85, 6.25), (3, 18)), ((7.92, 0.0), (16.85, 0.0), (3, 18)),
                ((9.63, -6.2), (16.85, -6.2), (3, 18)),
                ((-9.63, 13.75), (-9.63, 7.53), (3, 18)), ((9.63, 13.75), (9.63, 7.53), (3, 18)),
                ((-9.63, -7.55), (-9.63, -13.7), (3, 18)), ((9.63, -7.55), (9.63, -13.7), (3, 18)),
                ((0.0, 15.5), (0.0, 7.53)), ((0.0, -7.55), (0.0, -15.5)),
            ],
        },
        "fachada": {
            "Alas": {"contorno": 0, "ciegos": []},
            "Piso 19": {"contorno": 1, "ciegos": []},
        },
    },
    # Somma Las Clarisas, láminas LC A006 (piso 2), A007 (piso 3) y A008 (pisos 4–10),
    # escala 1:100. Torres A (bloque recto + inicio de la curva) y B (ala curva) en una
    # misma planta, por eso las claves son (torre, depto). Posiciones tomadas de las
    # etiquetas N° de la lámina del piso 3.
    "Somma Las Clarisas": {
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
        # Contorno de cada torre (muros exteriores, medidos en la planta A05 ampliada), solo en
        # los pisos de esa torre
        "torres_contorno": ["Boreal", "Austral"],
        # Silueta de cada torre según elevaciones y cortes (láminas A16 a A20): niveles N.P.T.
        # de la cubierta. Boreal = Ed. Lo Ovalle (pisos 1–16, mirador +40,00, cubierta +43,15);
        # Austral = Ed. Jorge Cáceres (pisos 1–22, azotea +55,28). (índice de contorno, desde, hasta)
        "volumenes": [(0, 0.0, 43.15), (1, 0.2, 55.28)],
        # Áreas comunes según la planta general A05: corredor central y núcleo (escaleras y
        # ascensores) de cada torre, de suelo a cubierta. (tipo, (x0, y0, x1, y1), desde, hasta)
        "areas_comunes": [
            ("Corredor", (-4.12, 8.03, -2.49, 34.92), 0.0, 40.0),
            ("Escaleras", (-10.21, 20.15, -6.68, 27.65), 0.0, 43.15),
            ("Ascensores", (-6.68, 20.37, -4.26, 26.55), 0.0, 43.15),
            ("Corredor", (0.82, -38.28, 2.36, -6.09), 0.2, 55.28),
            ("Escaleras", (-5.14, -26.15, -2.05, -16.89), 0.2, 55.28),
            ("Ascensores", (-2.05, -25.93, 0.82, -19.54), 0.2, 55.28),
            # Salas de máquinas sobre el núcleo (cajas oscuras sobre la cubierta en A16 y A17)
            ("Sala de máquinas", (-10.21, 20.15, -4.26, 27.65), 43.15, 46.15),
            ("Sala de máquinas", (-5.14, -26.15, 0.82, -16.89), 55.28, 58.78),
        ],
        # Fachada (elevaciones A16 a A20): por torre, el contorno que la forma y los muros ciegos
        # del núcleo de escaleras (la franja oscura de suelo a cubierta en la fachada poniente).
        # El resto del perímetro se dibuja como vidrio entre bordes de losa, con parapeto en la cubierta.
        "fachada": {
            "Boreal": {"contorno": 0, "ciegos": [((-10.56, 20.15), (-10.56, 27.65))]},
            "Austral": {"contorno": 1, "ciegos": [((-5.14, -26.15), (-5.14, -16.89))]},
        },
        # Antepechos de terrazas (planta A05): paños de 1 m de alto en la línea de fachada,
        # centrados en cada rótulo "TERRAZA" y de ~3 m de largo, en cada piso de su torre
        "terrazas": {
            "Boreal": [
                ((5.23, 34.73), (5.23, 37.73)), ((-10.56, 34.73), (-10.56, 37.73)),
                ((5.23, 28.82), (5.23, 31.82)), ((5.23, 23.4), (5.23, 26.4)),
                ((5.23, 14.82), (5.23, 17.82)), ((-10.56, 14.8), (-10.56, 17.8)),
                ((-10.56, 6.72), (-10.56, 9.72)), ((5.23, 6.72), (5.23, 9.72)),
            ],
            "Austral": [
                ((-5.14, -8.12), (-5.14, -5.12)), ((9.19, -11.63), (9.19, -8.63)),
                ((9.19, -15.84), (9.19, -12.84)), ((-5.14, -16.26), (-5.14, -13.26)),
                ((9.19, -26.8), (9.19, -23.8)), ((-5.14, -31.16), (-5.14, -28.16)),
                ((9.19, -33.05), (9.19, -30.05)), ((-5.14, -39.38), (-5.14, -36.38)),
            ],
        },
        # Muros entre departamentos (planta A05), como muros de piso a cielo en cada piso de su torre
        "particiones": {
            "Boreal": [((-10.56, 33.82), (-4.12, 33.82)), ((-10.56, 27.96), (-4.12, 27.96)),
                       ((-10.56, 20.15), (-4.12, 20.15)), ((-10.56, 13.98), (-4.12, 13.98)),
                       ((-2.49, 33.82), (5.23, 33.82)), ((-2.49, 27.96), (5.23, 27.96)),
                       ((-2.49, 20.15), (5.23, 20.15)), ((-2.49, 13.98), (5.23, 13.98))],
            "Austral": [((-5.14, -9.61), (0.82, -9.61)), ((-5.14, -16.89), (0.82, -16.89)),
                        ((-5.14, -26.15), (0.82, -26.15)), ((-5.14, -32.98), (0.82, -32.98)),
                        ((2.36, -13.8), (9.19, -13.8)), ((2.36, -17.99), (9.19, -17.99)),
                        ((2.36, -27.69), (9.19, -27.69))],
        },
        "contornos": [
            [(-10.56, 42.65), (5.23, 42.65), (5.23, 4.06), (-10.56, 4.06), (-10.56, 42.65)],
            [(-5.14, -2.12), (9.19, -2.12), (9.19, -41.81), (-5.14, -41.81), (-5.14, -2.12)],
        ],
    },
    # Somma Ines de Suarez, lámina A-07 (piso 3), con torres MAYOR (deptos 1–43, en C
    # alrededor del patio) y MENOR (deptos 1–7, bloque al sur-oriente) en su posición real.
    # La lámina está impresa reducida (dice 1:100, pero los deslindes A-B = 106,51 m y
    # F-E = 164,16 m dan 0,0956 m por punto); las coordenadas usan esa calibración.
    # Las plantas de los pisos 2, 4, 5–8 y 9 tienen la misma distribución (±1 m).
    "Somma Ines de Suarez": {
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
    # Somma Valle La Piramide (proyecto Vitapark, Rinconada El Salto), lámina ARQ-002
    # "Ubicación / Piso tipo" (1:500, A1). Bloque en E: barra al sur y tres alas al norte.
    # Cada depto rotula sus números de piso apilados en el mismo lugar (207, 307, 407…);
    # se usan las posiciones del piso 2. El segundo bloque de la lámina no está numerado.
    "Somma Valle La Piramide": {
        "deptos": {
            1: (48.5, 12.9), 2: (52.1, 10.6), 3: (46.4, 9.5), 4: (52.0, 3.1), 5: (46.4, 1.9),
            6: (52.0, -1.1), 7: (46.4, -2.5), 8: (52.0, -8.7), 9: (52.0, -13.0), 10: (47.4, -15.5),
            11: (40.4, -15.5), 12: (38.9, -10.6), 13: (35.5, -15.5), 14: (33.9, -10.6), 15: (28.4, -15.5),
            16: (26.9, -10.6), 17: (23.5, -15.5), 18: (21.9, -10.6), 19: (14.5, -15.5), 20: (12.9, -10.6),
            21: (7.4, -15.5), 22: (-5.2, -9.6), 23: (-3.8, -15.5), 24: (-3.1, 0.2), 25: (-8.3, 0.2),
            26: (-3.1, 6.2), 27: (-8.3, 6.2), 28: (-1.9, 14.2), 29: (-9.5, 14.2), 30: (-9.9, -15.5),
            31: (-17.2, -15.5), 32: (-17.3, -10.8), 33: (-21.8, -15.5), 34: (-24.2, -10.8), 35: (-30.8, -15.5),
            36: (-33.5, -10.8), 37: (-36.9, -15.5), 38: (-38.2, -10.8), 39: (-45.8, -15.5), 40: (-51.3, -13.0),
            41: (-51.3, -8.6), 42: (-45.9, -2.5), 43: (-51.0, -1.0), 44: (-45.7, 1.9), 45: (-51.0, 3.2),
            46: (-45.8, 9.5), 47: (-51.5, 10.4), 48: (-48.5, 13.2),
        },
        "contornos": [
            [(-57.8, -19.8), (57.8, -19.8), (57.8, 19.8), (40.1, 19.8), (40.1, -4.0), (3.1, -4.0), (3.1, 19.8), (-14.6, 19.8), (-14.6, -4.0), (-39.2, -4.0), (-39.2, 19.8), (-57.8, 19.8), (-57.8, -19.8)],
        ],
    },
    # Ronda Terravista, láminas A02 (piso 2), A04 (pisos 4 a 12), A07 (pisos 13 a 21) y A08
    # (piso 23), escala 1:75, todas en las mismas coordenadas. Barra con pasillo central y
    # bloques en los extremos. Los rótulos ("405-1205" = depto 05 de los pisos 4 a 12) están
    # junto al pasillo: en el tramo central cada punto se llevó al centro de su fila y en los
    # extremos al centro del bloque. Pisos sin lámina: el 3 usa la A04 y el 22 la A07.
    "Ronda Terravista": {
        "deptos": {
            1: (13.6, -5.4), 2: (18.6, -5.4), 3: (24.4, -4.6), 4: (24.4, 0.0), 5: (24.4, 4.0),
            6: (17.2, 4.3), 7: (11.3, 4.3), 8: (-1.6, 4.3), 9: (-6.1, 4.3), 10: (-14.9, 4.3),
            11: (-24.4, 3.6), 12: (-24.4, 0.0), 13: (-24.4, -4.8), 14: (-18.6, -5.4), 15: (-13.4, -5.4),
            16: (-8.8, -5.4), 17: (-3.8, -5.4),
        },
        "deptos_por_piso": [
            ((2, 2), {
                1: (13.4, -5.4), 2: (18.7, -5.4), 3: (24.4, -4.4), 4: (24.4, -0.2), 5: (24.4, 4.4),
                6: (17.1, 4.3), 7: (11.3, 4.3), 9: (-6.2, 4.3), 10: (-14.9, 4.3), 11: (-24.4, 3.7),
                12: (-24.4, -0.2), 13: (-24.4, -4.8), 14: (-18.8, -5.4), 15: (-13.3, -5.4), 16: (-9.0, -5.4),
                17: (-3.7, -5.4),
            }),
            ((13, 22), {
                1: (13.5, -5.4), 2: (18.7, -5.4), 3: (24.4, -5.2), 4: (24.4, 0.0), 5: (24.4, 4.2),
                6: (16.9, 4.3), 7: (11.2, 4.3), 8: (-1.8, 4.3), 9: (-5.9, 4.3), 10: (-15.2, 4.3),
                11: (-24.4, 3.9), 12: (-24.4, -0.4), 13: (-24.4, -5.1), 14: (-14.1, -5.4), 16: (-8.9, -5.4),
                17: (-3.8, -5.4),
            }),
            ((23, 23), {
                1: (13.5, -5.4), 2: (18.9, -5.4), 6: (17.2, 4.3), 7: (11.3, 4.3), 8: (-1.7, 4.3),
                9: (-6.0, 4.3), 10: (-14.9, 4.3), 14: (-14.3, -5.4), 16: (-8.8, -5.4), 17: (-3.8, -5.4),
            }),
        ],
        "contornos": [
            [(-28.6, 10.0), (28.6, 10.0), (28.6, -10.0), (-28.6, -10.0), (-28.6, 10.0)],
        ],
    },
    # Ronda Santo Domingo, láminas A-104 (piso 2), A-105 (3, 4, 6 y 9), A-106 (5, 7, 8, 10
    # y 11) y A-107 (12), escala 1:100 en A0, todas en las mismas coordenadas. Planta en U:
    # dos alas con dos columnas de deptos y una barra al sur. Los pisos 3 a 12 comparten
    # planta ("301 al 1201"); el piso 2 difiere en la posición de los deptos 18, 24 y 29.
    "Ronda Santo Domingo": {
        "deptos": {
            1: (-7.5, -16.7), 2: (-4.2, -16.6), 3: (1.8, -16.6), 4: (7.1, -16.7), 5: (15.2, -8.3),
            6: (15.1, -4.9), 7: (15.1, 1.3), 8: (15.0, 7.1), 9: (14.9, 15.8), 10: (15.0, 19.0),
            11: (23.5, 22.8), 12: (23.7, 13.3), 13: (23.8, 9.9), 14: (23.6, 1.4), 15: (23.7, -4.8),
            16: (23.6, -8.4), 17: (23.7, -17.1), 18: (23.0, -24.7), 19: (13.4, -24.7), 20: (4.7, -24.8),
            21: (1.5, -24.8), 22: (-7.7, -24.6), 23: (-16.3, -24.8), 24: (-23.4, -24.8), 25: (-23.4, -17.4),
            26: (-23.5, -8.2), 27: (-23.6, -4.8), 28: (-23.5, 3.9), 29: (-22.7, 13.3), 30: (-15.9, 13.3),
            31: (-15.1, 3.8), 32: (-15.0, -2.4), 33: (-15.0, -8.4),
        },
        "deptos_por_piso": [
            ((2, 2), {
                1: (-7.6, -16.4), 2: (-4.3, -16.4), 3: (1.7, -16.4), 4: (7.7, -16.5), 5: (15.2, -8.2),
                6: (15.2, -5.1), 7: (15.2, 1.0), 8: (15.0, 7.0), 9: (15.0, 16.0), 10: (15.1, 19.0),
                11: (23.5, 22.6), 12: (23.7, 13.1), 13: (23.7, 9.8), 14: (23.8, 1.0), 15: (23.7, -4.9),
                16: (23.7, -8.4), 17: (23.6, -17.2), 18: (19.7, -25.3), 19: (13.5, -25.0), 20: (4.7, -25.0),
                21: (1.5, -24.8), 22: (-7.2, -24.8), 23: (-16.3, -24.8), 24: (-19.4, -25.8), 25: (-23.4, -17.1),
                26: (-23.6, -8.1), 27: (-23.8, -5.0), 28: (-23.5, 4.0), 29: (-21.4, 12.7), 30: (-16.0, 13.4),
                31: (-15.1, 4.0), 32: (-15.1, -2.0), 33: (-15.1, -8.1),
            }),
        ],
        "contornos": [
            [(-27.7, 18.3), (-11.4, 18.3), (-11.4, -12.7), (11.9, -12.7), (11.9, 27.9), (27.7, 27.9), (27.7, -27.9), (-27.7, -27.9), (-27.7, 18.3)],
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
    posiciones = plano and plano["deptos"]
    piso = grupo["Piso_Num"].iloc[0]
    for (desde, hasta), por_piso in (plano or {}).get("deptos_por_piso", []):
        if desde <= piso <= hasta:  # planta propia de este rango de pisos
            posiciones = por_piso
    if plano and all(k in posiciones for k in claves):
        grupo["X"] = [posiciones[k][0] for k in claves]
        grupo["Y"] = [posiciones[k][1] for k in claves]
        grupo["Grid_Size"] = 0  # posición según plano, no grilla
        return grupo
    grid_size = tamano_grilla(len(grupo))
    offset = (grid_size - 1) / 2.0           # centrado matemático en (0,0)
    idx = np.arange(len(grupo))
    grupo["X"] = ((idx % grid_size) - offset) * MODULO_GRILLA
    grupo["Y"] = ((idx // grid_size) - offset) * MODULO_GRILLA
    grupo["Grid_Size"] = grid_size
    return grupo


def paños_fachada(plano):
    """Paños de fachada de cada torre con elevaciones, como (x0, y0, x1, y1, z0, z1, piso):
    bordes de losa (25 cm) y vidrio entre ellos en todo el perímetro, salvo los muros ciegos
    (de piso a piso), más el borde de la cubierta y su parapeto (1 m, piso None)."""
    paños = {"losas": [], "vidrios": [], "ciegos": [], "parapetos": []}
    for f in plano.get("fachada", {}).values():
        k = f["contorno"]
        z0, z1 = next(v[1:] for v in plano["volumenes"] if v[0] == k)
        c = plano["contornos"][k]
        aristas = [(c[i], c[i + 1]) for i in range(len(c) - 1)]
        # tramos de vidrio: cada arista sin los muros ciegos que caen sobre ella
        vidrios = []
        for (ax, ay), (bx, by) in aristas:
            largo = math.hypot(bx - ax, by - ay)
            ux, uy = (bx - ax) / largo, (by - ay) / largo
            cortes = []
            for extremos in f.get("ciegos", []):
                if all(abs((px - ax) * uy - (py - ay) * ux) < 0.05 for px, py in extremos):
                    t0, t1 = sorted((px - ax) * ux + (py - ay) * uy for px, py in extremos)
                    cortes.append((max(t0, 0), min(t1, largo)))
            t = 0
            for t0, t1 in sorted(cortes) + [(largo, largo)]:
                if t0 - t > 0.05:
                    vidrios.append(((ax + ux * t, ay + uy * t), (ax + ux * t0, ay + uy * t0)))
                t = max(t, t1)
        niveles = [n * ALTURA_PISO for n in range(round(z0 / ALTURA_PISO), int(z1 / ALTURA_PISO) + 1)
                   if n * ALTURA_PISO < z1 - 0.5] + [z1]
        for za, zb in zip(niveles, niveles[1:]):
            n = round(za / ALTURA_PISO) + 1
            for (x0, y0), (x1, y1) in aristas:
                paños["losas"].append((x0, y0, x1, y1, za, za + 0.25, n))
            for (x0, y0), (x1, y1) in vidrios:
                paños["vidrios"].append((x0, y0, x1, y1, za + 0.25, zb, n))
            for (x0, y0), (x1, y1) in f.get("ciegos", []):
                paños["ciegos"].append((x0, y0, x1, y1, za, zb, n))
        for (x0, y0), (x1, y1) in aristas:
            paños["losas"].append((x0, y0, x1, y1, z1 - 0.25, z1, None))
            paños["parapetos"].append((x0, y0, x1, y1, z1, z1 + 1, None))
    return paños


def agregar_deptos_sin_ot(deptos):
    """Agrega con 0 OT los deptos que el plano tiene en cada piso pero que no registran OT,
    para que la maqueta muestre el edificio completo (la burbuja más chica = 0 OT). Cada
    torre se completa solo entre su piso más bajo y más alto con OT."""
    filas = []
    for ed, plano in PLANTAS.items():
        sub = deptos[deptos["Edificio"] == ed]
        if sub.empty:
            continue
        juntas = plano.get("unir_torres")
        for torre, st in (sub.groupby("Torre") if juntas else [(sub["Torre"].iloc[0], sub)]):
            for piso in range(int(st["Piso_Num"].min()), int(st["Piso_Num"].max()) + 1):
                mapa = plano["deptos"]
                for (desde, hasta), por_piso in plano.get("deptos_por_piso", []):
                    if desde <= piso <= hasta:
                        mapa = por_piso
                existentes = set(st.loc[st["Piso_Num"] == piso, "Depto_Num"])
                for clave in mapa:
                    if juntas and clave[0] != torre:
                        continue
                    num = clave[1] if juntas else clave
                    if num in existentes:
                        continue
                    filas.append({
                        "Depto_ID": f"{ed} | {torre} | P{piso}-D{num:02d}", "Edificio": ed,
                        "Torre": torre, "Piso_Num": piso, "Depto_Num": num,
                        "Depto_Txt": f"Torre {torre} · Depto {num}" if juntas else f"Depto {num}",
                        "Densidad_Total": 0,
                    })
    return pd.concat([deptos, pd.DataFrame(filas)], ignore_index=True) if filas else deptos


deptos = agregar_deptos_sin_ot(deptos)
deptos = pd.concat(
    [asignar_coordenadas(g) for _, g in deptos.groupby(["Edificio", "Piso_Num"])],
    ignore_index=True,
)
deptos["Z"] = (deptos["Piso_Num"] - 1) * ALTURA_PISO  # piso 1 a nivel de suelo (m), como en los cortes
df = df.merge(deptos[["Depto_ID", "X", "Y", "Z", "Grid_Size", "Densidad_Total"]],
              on="Depto_ID", how="left")

# ---------------------------------------------------------------------------
# 4. ESCALA FIJA ABSOLUTA (calculada ANTES de cualquier filtro)
# ---------------------------------------------------------------------------
# En metros e incluyendo los contornos de planta de los edificios presentes
_contornos = [pt for ed, p in PLANTAS.items() if (deptos["Edificio"] == ed).any()
              for c in p["contornos"] for pt in c]
LIM_XY = max([float(deptos["X"].abs().max()), float(deptos["Y"].abs().max())]
             + [max(abs(x), abs(y)) for x, y in _contornos]) + MARGEN_EJES
RANGO_X = [-LIM_XY, LIM_XY]
RANGO_Y = [-LIM_XY, LIM_XY]
_techo = max([float(deptos["Z"].max())] + [v[2] for ed, p in PLANTAS.items()
                                           if (deptos["Edificio"] == ed).any() for v in p.get("volumenes", [])]
             + [a[3] for ed, p in PLANTAS.items() if (deptos["Edificio"] == ed).any() for a in p.get("areas_comunes", [])])
RANGO_Z = [-ALTURA_PISO / 2, _techo + ALTURA_PISO]
DENSIDAD_MAX = int(deptos["Densidad_Total"].max())
# rótulos del eje vertical: el primer piso y luego cada 5
PISOS = [n for n in range(int(deptos["Piso_Num"].min()), int(deptos["Piso_Num"].max()) + 1)
         if n == int(deptos["Piso_Num"].min()) or n % 5 == 0]

# Proporción espacial real: la misma escala en planta y en altura
span_z = RANGO_Z[1] - RANGO_Z[0]
ASPECTO = dict(x=1, y=1, z=span_z / (2 * LIM_XY))


def tamano_marcador(d):
    """La burbuja más chica es un depto sin OT; las grandes (y rojas) son las más críticas."""
    if d == 0:
        return 2
    if d <= 1:
        return 3
    if d <= 5:
        return 5
    if d <= 10:
        return 9
    if d <= 15:
        return 16
    return 28


# Colores de las áreas comunes (cajas translúcidas): índices de traza 6 en adelante
COLOR_COMUN = {"Corredor": "#262b38", "Escaleras": "#2e3342", "Ascensores": "#222838",
               "Sala de máquinas": "#1c2030"}
# Fachada translúcida (torres con elevaciones): color y opacidad de cada tipo de paño
ESTILO_FACHADA = {"losas": ("#5d6684", 0.55), "vidrios": ("#4f6fa8", 0.10),
                  "ciegos": ("#262b38", 0.60), "parapetos": ("#3d4456", 0.55)}
I_COMUN = 6                              # primer índice de traza de las áreas comunes
I_MUROS = I_COMUN + len(COLOR_COMUN)     # muros entre deptos; el siguiente, antepechos
I_FACHADA = I_MUROS + 2                  # paños de fachada, en el orden de ESTILO_FACHADA

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
        # Líneas finas: silueta de las torres y particiones entre deptos
        go.Scatter3d(x=[], y=[], z=[], mode="lines", hoverinfo="skip",
                     line=dict(color="#4b5170", width=1)),
    ] + [
        # Corredores, escaleras y ascensores: cajas blanquecinas translúcidas
        go.Mesh3d(x=[], y=[], z=[], i=[], j=[], k=[], name=tipo, color=color, opacity=0.45,
                  flatshading=True, hoverinfo="skip", lighting=dict(ambient=0.6, diffuse=0.2, specular=0, roughness=1, fresnel=0))
        for tipo, color in COLOR_COMUN.items()
    ] + [
        # Muros entre deptos (piso a cielo) y antepechos de terrazas (1 m): índices I_MUROS y siguiente,
        # y luego los paños de fachada (I_FACHADA en adelante)
        go.Mesh3d(x=[], y=[], z=[], i=[], j=[], k=[], name=nombre, color=color, opacity=opacidad,
                  flatshading=True, hoverinfo="skip", lighting=dict(ambient=0.6, diffuse=0.2, specular=0, roughness=1, fresnel=0))
        for nombre, color, opacidad in [("Muros", "#8f98b8", 0.30), ("Antepechos", "#3d4456", 0.55)]
        + [(nombre, color, opacidad) for nombre, (color, opacidad) in ESTILO_FACHADA.items()]
    ],
    layout=go.Layout(
        template="plotly_dark",
        paper_bgcolor="#000000", plot_bgcolor="#000000",
        height=720, margin=dict(l=0, r=0, t=50, b=0),
        title=dict(text="Maqueta 3D de Fallas por Departamento", font=dict(color="#eee")),
        uirevision="fijo",  # conserva la cámara al actualizar datos
        showlegend=False,
        scene=dict(
            xaxis=dict(range=RANGO_X, title="X (m)", **estilo_eje),
            yaxis=dict(range=RANGO_Y, title="Y (m)", **estilo_eje),
            # altura real (2,5 m por piso) rotulada con el número de piso
            zaxis=dict(range=RANGO_Z, title=f"Piso ({str(ALTURA_PISO).replace('.', ',')} m c/u)",
                       tickvals=[(n - 1) * ALTURA_PISO for n in PISOS], ticktext=[str(n) for n in PISOS],
                       **estilo_eje),
            aspectmode="manual", aspectratio=ASPECTO,
            camera=dict(eye=dict(x=0.95, y=0.95, z=0.6)),
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

# Time lapse: cada OT queda activa VIDA_OT días desde su creación y luego se apaga.
# Play trae sus propios botones de reproducir, pausa y detener.
VIDA_OT = 30
DIA0 = fechas[0]
N_DIAS = (fechas[-1] - DIA0).days + 1
df["Dia"] = (df["Fecha"].dt.normalize() - DIA0).dt.days
tl_btn = widgets.ToggleButton(description="Time lapse (30 días por falla)", icon="play",
                              layout=widgets.Layout(width="260px"))
tl_play = widgets.Play(value=0, min=0, max=N_DIAS - 1, step=2, interval=200)
tl_dia = widgets.IntSlider(value=0, min=0, max=N_DIAS - 1, readout=False,
                           layout=widgets.Layout(width="360px"))
widgets.jslink((tl_play, "value"), (tl_dia, "value"))
tl_vel = widgets.Dropdown(options=[("5 días/s", 1), ("10 días/s", 2), ("20 días/s", 4), ("40 días/s", 8)],
                          value=2, description="Velocidad:", layout=widgets.Layout(width="190px"))
tl_lbl = widgets.HTML()
tl_ctl = widgets.HBox([tl_play, tl_dia, tl_vel, tl_lbl],
                      layout=widgets.Layout(display="none", align_items="center", gap="10px"))
_cache_tl = {}


def max_activas(m_rango, clave):
    """Máximo de OT activas a la vez en un mismo depto: fija la escala del time lapse
    para que no salte mientras se reproduce (se guarda por combinación de filtros)."""
    if clave not in _cache_tl:
        sub = df.loc[m_rango, ["Dia", "Depto_ID"]]
        if sub.empty:
            _cache_tl[clave] = 1
        else:
            tabla = sub.groupby(["Dia", "Depto_ID"]).size().unstack(fill_value=0)
            tabla = tabla.reindex(range(tabla.index.min(), tabla.index.max() + 1), fill_value=0)
            _cache_tl[clave] = int(max(1, tabla.rolling(VIDA_OT, min_periods=1).sum().to_numpy().max()))
    return _cache_tl[clave]

# Simbología del depto desplegado (oculta mientras no haya ninguno)
html_leyenda = widgets.HTML()
btn_cerrar = widgets.Button(description="Cerrar", layout=widgets.Layout(width="90px"))
caja_leyenda = widgets.VBox(
    [html_leyenda, btn_cerrar],
    layout=widgets.Layout(display="none", width="290px", padding="10px 12px",
                          border="1px solid #2c2c3a", margin="50px 8px 0 0"),
)
caja_leyenda.add_class("ot-leyenda")
estado = {"abierto": None, "puntos": [], "densidad": [], "ultimo": (None, 0.0)}


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

    # En el time lapse solo cuentan las OT creadas en los últimos VIDA_OT días
    escala_tl = None
    if tl_btn.value:
        escala_tl = max_activas(m, (f_ini, f_fin, dd_edificio.value, dd_piso.value, dd_clase.value))
        dia = DIA0 + pd.Timedelta(days=tl_dia.value)
        m &= (df["Fecha"] >= dia - pd.Timedelta(days=VIDA_OT - 1)) & (df["Fecha"] < dia + pd.Timedelta(days=1))

    # Densidad dentro del rango filtrado; posiciones y escalas siguen fijas. Los deptos del
    # edificio/piso sin OT en el rango también se dibujan, como el punto más chico.
    base = deptos
    if dd_edificio.value != TODOS:
        base = base[base["Edificio"] == dd_edificio.value]
    if dd_piso.value != TODOS:
        base = base[base["Piso_Num"] == int(dd_piso.value)]
    cuenta = df[m].groupby("Depto_ID").size().rename("Densidad")
    sub = base[["Depto_ID", "Edificio", "Piso_Num", "Depto_Num", "Depto_Txt", "X", "Y", "Z"]].merge(
        cuenta, left_on="Depto_ID", right_index=True, how="left").fillna({"Densidad": 0})
    sub["Densidad"] = sub["Densidad"].astype(int)
    estado["densidad"] = sub["Densidad"].tolist()
    con_ot = int((sub["Densidad"] > 0).sum())
    estado["puntos"] = sub["Depto_ID"].tolist()
    if estado["abierto"] not in set(sub.loc[sub["Densidad"] > 0, "Depto_ID"]):
        estado["abierto"] = None  # el filtro dejó fuera al depto abierto (o quedó sin OT)
    abierto = estado["abierto"]

    etiqueta = f"OT activas (últimos {VIDA_OT} días)" if escala_tl else "Fallas en rango"
    textos = [f"<b>{r.Edificio}</b><br>Piso {r.Piso_Num} · {r.Depto_Txt}<br>"
              + (f"{etiqueta}: <b>{r.Densidad}</b>" if r.Densidad else "Sin OT")
              for r in sub.itertuples()]

    with fig.batch_update():
        t = fig.data[0]
        t.x, t.y, t.z = sub["X"].tolist(), sub["Y"].tolist(), sub["Z"].tolist()
        t.marker.size = [(3 + 25 * math.sqrt(d / escala_tl) if d else 2) if escala_tl
                         else tamano_marcador(d) for d in sub["Densidad"]]
        t.marker.color = sub["Densidad"].tolist()
        t.marker.cmax = escala_tl or DENSIDAD_MAX
        t.marker.colorbar.title.text = "OT activas" if escala_tl else "Fallas"
        t.text = textos
        t.marker.opacity = 0.12 if abierto else 1  # opacas; se oscurecen al desplegar
        dibujar_molecula(df[m & (df["Depto_ID"] == abierto)] if abierto else None)
        dibujar_contorno()
        # Re-afirma los límites absolutos para que nada se auto-escale (no en cada cuadro del
        # time lapse: ya están fijos y tocar el layout redibuja la escena y corta el giro)
        if not escala_tl:
            fig.layout.scene.xaxis.range = RANGO_X
            fig.layout.scene.yaxis.range = RANGO_Y
            fig.layout.scene.zaxis.range = RANGO_Z

    if escala_tl:
        tl_lbl.value = (f"<span style='color:#eee'>{dia:%d-%m-%Y} · {int(m.sum())} OT activas "
                        f"en {con_ot} deptos</span>")
        lbl_resumen.value = (f"<span style='color:#9aa0a6'>Time lapse: cada OT se muestra {VIDA_OT} días "
                             f"desde su creación y luego se apaga · escala fija hasta {escala_tl} OT "
                             f"activas por depto</span>")
        return
    lbl_resumen.value = (f"<span style='color:#9aa0a6'>{int(m.sum())} OT en "
                         f"{con_ot} departamentos · máx. "
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
            # Cada contorno solo en los pisos de su torre y rango, si el plano los indica
            sel = deptos["Edificio"] == dd_edificio.value
            if plano.get("torres_contorno") and plano["torres_contorno"][k]:
                sel &= deptos["Torre"] == plano["torres_contorno"][k]
            if plano.get("pisos_contorno") and plano["pisos_contorno"][k]:
                sel &= deptos["Piso_Num"].between(*plano["pisos_contorno"][k])
            pisos = sorted(deptos.loc[sel, "Piso_Num"].unique())
            if dd_piso.value != TODOS:
                pisos = [n for n in pisos if n == int(dd_piso.value)]
            for n in pisos:
                lx += [x for x, _ in c] + [None]
                ly += [y for _, y in c] + [None]
                lz += [(int(n) - 1) * ALTURA_PISO] * len(c) + [None]
    fx, fy, fz = [], [], []  # líneas finas
    muros, antepechos = (dict(x=[], y=[], z=[], i=[], j=[], k=[]) for _ in range(2))
    cajas = {tipo: dict(x=[], y=[], z=[], i=[], j=[], k=[]) for tipo in COLOR_COMUN}
    fachada = {clave: dict(x=[], y=[], z=[], i=[], j=[], k=[]) for clave in ESTILO_FACHADA}
    if plano:
        # Fachada: bordes de losa, vidrio, muros ciegos del núcleo y parapeto de cubierta
        for clave, paños in paños_fachada(plano).items():
            for x0, y0, x1, y1, z0, z1, n in paños:
                if dd_piso.value == TODOS or (n is not None and n == int(dd_piso.value)):
                    agregar_paño(fachada[clave], x0, y0, x1, y1, z0, z1)
        # Silueta de cada torre (si hay elevaciones): aristas de suelo a cubierta
        if dd_piso.value == TODOS:
            for k, z0, z1 in plano.get("volumenes", []):
                c = plano["contornos"][k]
                for x, y in c[:-1]:
                    fx += [x, x, None]; fy += [y, y, None]; fz += [z0, z1, None]
                for zz in (z0, z1):
                    fx += [x for x, _ in c] + [None]
                    fy += [y for _, y in c] + [None]
                    fz += [zz] * len(c) + [None]
        # Particiones entre deptos (línea en el piso y muro de piso a cielo) y antepechos de
        # terrazas (1 m de alto), en cada piso de su torre
        for torre, segs in plano.get("particiones", {}).items():
            pisos = sorted(deptos.loc[(deptos["Edificio"] == dd_edificio.value) & (deptos["Torre"] == torre),
                                      "Piso_Num"].unique())
            if dd_piso.value != TODOS:
                pisos = [n for n in pisos if n == int(dd_piso.value)]
            for n in pisos:
                zz = (int(n) - 1) * ALTURA_PISO
                for (x0, y0), (x1, y1), *rango in segs:
                    if rango and not rango[0][0] <= n <= rango[0][1]:
                        continue  # tramo con rango de pisos propio
                    # línea en el piso y en el cielo, y el muro entre ambas
                    fx += [x0, x1, None, x0, x1, None]; fy += [y0, y1, None, y0, y1, None]
                    fz += [zz, zz, None, zz + ALTURA_PISO, zz + ALTURA_PISO, None]
                    agregar_paño(muros, x0, y0, x1, y1, zz, zz + ALTURA_PISO)
                for (x0, y0), (x1, y1), *rango in plano.get("terrazas", {}).get(torre, []):
                    if rango and not rango[0][0] <= n <= rango[0][1]:
                        continue
                    agregar_paño(antepechos, x0, y0, x1, y1, zz, zz + 1)
        # Corredores y núcleos: de suelo a cubierta, o solo el tramo del piso elegido
        for tipo, (x0, y0, x1, y1), z0, z1 in plano.get("areas_comunes", []):
            if dd_piso.value != TODOS:
                n = int(dd_piso.value)
                z0, z1 = max(z0, (n - 1) * ALTURA_PISO), min(z1, n * ALTURA_PISO)
                if z1 <= z0:
                    continue
            m = cajas.setdefault(tipo, dict(x=[], y=[], z=[], i=[], j=[], k=[]))
            b = len(m["x"])
            for vx, vy, vz in [(x0, y0, z0), (x1, y0, z0), (x1, y1, z0), (x0, y1, z0),
                               (x0, y0, z1), (x1, y0, z1), (x1, y1, z1), (x0, y1, z1)]:
                m["x"].append(vx); m["y"].append(vy); m["z"].append(vz)
            for a1, a2, a3 in [(0, 1, 2), (0, 2, 3), (4, 5, 6), (4, 6, 7), (0, 1, 5), (0, 5, 4),
                               (1, 2, 6), (1, 6, 5), (2, 3, 7), (2, 7, 6), (3, 0, 4), (3, 4, 7)]:
                m["i"].append(b + a1); m["j"].append(b + a2); m["k"].append(b + a3)
    fig.data[4].x, fig.data[4].y, fig.data[4].z = lx, ly, lz
    fig.data[5].x, fig.data[5].y, fig.data[5].z = fx, fy, fz
    for traza, tipo in zip(fig.data[I_COMUN:I_MUROS], COLOR_COMUN):
        traza.update(**cajas[tipo])
    fig.data[I_MUROS].update(**muros)
    fig.data[I_MUROS + 1].update(**antepechos)
    for traza, clave in zip(fig.data[I_FACHADA:], ESTILO_FACHADA):
        traza.update(**fachada[clave])


def agregar_paño(m, x0, y0, x1, y1, z0, z1):
    """Paño vertical: el segmento en planta (x0, y0)-(x1, y1) levantado de z0 a z1."""
    b = len(m["x"])
    m["x"] += [x0, x1, x1, x0]; m["y"] += [y0, y1, y1, y0]; m["z"] += [z0, z0, z1, z1]
    m["i"] += [b, b]; m["j"] += [b + 1, b + 2]; m["k"] += [b + 2, b + 3]


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
    for (dx, dy, dz) in posiciones_esfera(len(conteo), 0.07 * (RANGO_X[1] - RANGO_X[0])):
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
        if not estado["densidad"][points.point_inds[0]]:
            return  # depto sin OT: no hay tipos que desplegar
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


def alternar_time_lapse(ch):
    tl_ctl.layout.display = "flex" if ch.new else "none"
    tl_btn.icon = "stop" if ch.new else "play"
    tl_btn.description = "Salir del time lapse" if ch.new else "Time lapse (30 días por falla)"
    if ch.new:  # parte en el inicio del rango de fechas elegido
        tl_play.value = (sl_fechas.value[0] - DIA0).days
    actualizar_grafico()


tl_btn.observe(alternar_time_lapse, names="value")

tl_dia.observe(actualizar_grafico, names="value")
tl_vel.observe(lambda ch: setattr(tl_play, "step", ch.new), names="value")

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
                      widgets.HBox([tl_btn, tl_ctl], layout=widgets.Layout(gap="12px")),
                      widgets.HBox([caja_leyenda, fig])])
panel.add_class("ot-viz-container")

actualizar_pisos()
actualizar_grafico()
display(panel)

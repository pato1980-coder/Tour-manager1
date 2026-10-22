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
MARGEN_EJES = 1.0          # margen estático de seguridad en X/Y/Z
GRID_MIN, GRID_MAX = 3, 7  # grilla dinámica 3x3 … 7x7
TODOS = "Todos"

# Escala continua Azul -> Morado -> Rojo
ESCALA_COLORES = [
    [0.00, "#1f4fff"],
    [0.25, "#4b3cff"],
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
varias = df.groupby("Edificio")["Torre"].transform("nunique") > 1
df.loc[varias, "Edificio"] = (df.loc[varias, "Edificio"] + " · "
                              + df.loc[varias, "Torre"].replace("", "sin torre"))
col_clase = buscar_columna(df, COL_CLASE, ("clasificacion", "estado clasificacion"))
df["Clase"] = (df[col_clase].fillna("Sin clasificación").astype(str).str.strip()
               if col_clase else "Sin clasificación")
df["Depto_ID"] = (df["Edificio"] + " | P" + df["Piso_Num"].astype(str)
                  + "-D" + df["Depto_Num"].astype(str).str.zfill(2))
print(f"Registros válidos: {len(df)} de {filas_antes} "
      f"({filas_antes - len(df)} descartados por fecha/unidad inválida)")

# Densidad acumulada de fallas por departamento único (todo el conjunto)
deptos = (df.groupby(["Depto_ID", "Edificio", "Piso_Num", "Depto_Num"])
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
    return 22


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
    )],
    layout=go.Layout(
        template="plotly_dark",
        paper_bgcolor="#000000", plot_bgcolor="#000000",
        height=720, margin=dict(l=0, r=0, t=50, b=0),
        title=dict(text="Maqueta 3D de Fallas por Departamento", font=dict(color="#eee")),
        uirevision="fijo",  # conserva la cámara al actualizar datos
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
    sub = (df[m].groupby(["Depto_ID", "Edificio", "Piso_Num", "Depto_Num", "X", "Y", "Z"])
               .size().reset_index(name="Densidad"))

    textos = [f"<b>{r.Edificio}</b><br>Piso {r.Piso_Num} · Depto {r.Depto_Num}"
              f"<br>Fallas en rango: <b>{r.Densidad}</b>"
              for r in sub.itertuples()]

    with fig.batch_update():
        t = fig.data[0]
        t.x, t.y, t.z = sub["X"].tolist(), sub["Y"].tolist(), sub["Z"].tolist()
        t.marker.size = [tamano_marcador(d) for d in sub["Densidad"]]
        t.marker.color = sub["Densidad"].tolist()
        t.text = textos
        # Re-afirma los límites absolutos para que nada se auto-escale
        fig.layout.scene.xaxis.range = RANGO_X
        fig.layout.scene.yaxis.range = RANGO_Y
        fig.layout.scene.zaxis.range = RANGO_Z

    lbl_resumen.value = (f"<span style='color:#9aa0a6'>{int(m.sum())} OT · "
                         f"{len(sub)} departamentos visibles · máx. "
                         f"{int(sub['Densidad'].max()) if len(sub) else 0} fallas/depto</span>")


dd_edificio.observe(lambda ch: (actualizar_pisos(), actualizar_grafico()), names="value")
dd_piso.observe(actualizar_grafico, names="value")
dd_clase.observe(actualizar_grafico, names="value")
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
panel = widgets.VBox([fila_filtros, sl_fechas, lbl_resumen, fig])
panel.add_class("ot-viz-container")

actualizar_pisos()
actualizar_grafico()
display(panel)

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
MARGEN_EJES = 1.0          # margen estático de seguridad en X/Y/Z
GRID_MIN, GRID_MAX = 3, 7  # grilla dinámica 3x3 … 7x7

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
    Extrae (Piso_Num, Depto_Num) de textos como '1204', 'Depto 1204',
    'Torre A - 304', 'D-1204', '12-04'. Convención: los 2 últimos dígitos
    son el departamento y los anteriores el piso.
    """
    if pd.isna(unidad):
        return (np.nan, np.nan)
    s = str(unidad).strip()
    if re.fullmatch(r"\d+\.0", s):  # números leídos como float por Excel
        s = s[:-2]
    # Formato explícito piso-depto (p.ej. '12-04')
    m = re.search(r"(\d{1,3})\s*[-/]\s*(\d{1,2})\s*$", s)
    if m:
        return (int(m.group(1)), int(m.group(2)))
    grupos = re.findall(r"\d+", s)
    if not grupos:
        return (np.nan, np.nan)
    num = max(grupos, key=len)  # el bloque numérico más largo es el nº de unidad
    if len(num) >= 3:
        return (int(num[:-2]), int(num[-2:]))
    return (1, int(num))  # unidades de 1-2 dígitos: se asumen en piso 1


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
    df[["Piso_Num", "Depto_Num"]] = df[col_unidad].apply(
        lambda u: pd.Series(extraer_piso_depto(u)))

    filas_antes = len(df)
    df = df.dropna(subset=["Fecha", "Piso_Num", "Depto_Num"]).copy()
    df["Piso_Num"] = df["Piso_Num"].astype(int)
    df["Depto_Num"] = df["Depto_Num"].astype(int)
    df["Edificio"] = df[col_edificio].astype(str).str.strip()
    df["Depto_ID"] = (df["Edificio"] + " | P" + df["Piso_Num"].astype(str)
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
    grid_size = tamano_grilla(len(grupo))
    offset = (grid_size - 1) / 2.0           # centrado matemático en (0,0)
    idx = np.arange(len(grupo))
    grupo["X"] = (idx % grid_size) - offset
    grupo["Y"] = (idx // grid_size) - offset
    grupo["Grid_Size"] = grid_size
    return grupo


def construir_modelo(df):
    # Densidad acumulada de fallas por departamento único (todo el conjunto)
    deptos = (df.groupby(["Depto_ID", "Edificio", "Piso_Num", "Depto_Num"])
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

    # Cada OT se reduce a (índice de depto, día) para filtrar en el navegador
    idx_depto = {d: i for i, d in enumerate(deptos["Depto_ID"])}
    dias = df["Fecha"].dt.normalize()
    fecha_0 = dias.min()
    ots = np.column_stack([
        df["Depto_ID"].map(idx_depto).to_numpy(),
        (dias - fecha_0).dt.days.to_numpy(),
    ]).astype(int)

    return {
        "deptos": {
            "edificio": deptos["Edificio"].tolist(),
            "piso": deptos["Piso_Num"].astype(int).tolist(),
            "depto": deptos["Depto_Num"].astype(int).tolist(),
            "x": deptos["X"].astype(float).tolist(),
            "y": deptos["Y"].astype(float).tolist(),
            "z": deptos["Z"].astype(int).tolist(),
        },
        "ots": ots.ravel().tolist(),  # pares aplanados [depto, día, depto, día, …]
        "fecha0": fecha_0.strftime("%Y-%m-%d"),
        "nDias": int((dias.max() - fecha_0).days) + 1,
        "rangoX": [-lim_xy, lim_xy],
        "rangoY": [-lim_xy, lim_xy],
        "rangoZ": rango_z,
        "aspecto": {"x": 1, "y": 1, "z": max(0.6, min(2.5, span_z / (2 * lim_xy) * 0.6))},
        "densidadMax": int(deptos["Densidad_Total"].max()),
        "escala": ESCALA_COLORES,
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
  #grafico { flex: 1 1 auto; min-height: 560px; width: 100%; }
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
  #contenedor:fullscreen #grafico, #contenedor.fs-fallback #grafico { min-height: 0; }
  @media (max-width: 640px) { #contenedor { padding: 10px 16px; } .btn-fs { margin-left: 0; } }
</style>
</head>
<body>
<div id="contenedor">
  <div class="fila">
    <label>Edificio <select id="selEdificio"></select></label>
    <label>Piso <select id="selPiso"></select></label>
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
  <div id="grafico"></div>
</div>
<script>
const M = __DATOS__;
const TODOS = 'Todos';
const D = M.deptos, nDeptos = D.x.length;
const f0 = new Date(M.fecha0 + 'T00:00:00');
const fmt = i => new Date(f0.getTime() + i * 864e5).toLocaleDateString('es-CL');
const $ = id => document.getElementById(id);

function tamano(d) { return d <= 1 ? 4 : d <= 5 ? 8 : d <= 10 ? 14 : 22; }
function opciones(sel, valores, actual) {
  sel.innerHTML = '';
  [TODOS, ...valores].forEach(v => sel.add(new Option(v, v)));
  sel.value = valores.map(String).includes(String(actual)) ? actual : TODOS;
}

// Controles
const edificios = [...new Set(D.edificio)].sort();
opciones($('selEdificio'), edificios, TODOS);
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

function actualizar() {
  let a = +$('desde').value, b = +$('hasta').value;
  if (a > b) [a, b] = [b, a];
  const pct = n => (n / Math.max(1, M.nDias - 1)) * 100;
  $('rangoActivo').style.left = pct(a) + '%';
  $('rangoActivo').style.width = (pct(b) - pct(a)) + '%';
  $('textoFechas').textContent = fmt(a) + ' – ' + fmt(b);

  const e = $('selEdificio').value, p = $('selPiso').value;
  const visible = i => (e === TODOS || D.edificio[i] === e) && (p === TODOS || String(D.piso[i]) === p);
  const cuenta = new Int32Array(nDeptos);
  let nOt = 0;
  for (let k = 0; k < M.ots.length; k += 2) {
    const i = M.ots[k], dia = M.ots[k + 1];
    if (dia >= a && dia <= b && visible(i)) { cuenta[i]++; nOt++; }
  }
  const x = [], y = [], z = [], c = [], s = [], t = [];
  let max = 0;
  for (let i = 0; i < nDeptos; i++) {
    const d = cuenta[i];
    if (!d) continue;
    x.push(D.x[i]); y.push(D.y[i]); z.push(D.z[i]); c.push(d); s.push(tamano(d));
    t.push(`<b>${D.edificio[i]}</b><br>Piso ${D.piso[i]} · Depto ${D.depto[i]}<br>Fallas en rango: <b>${d}</b>`);
    max = Math.max(max, d);
  }
  Plotly.react('grafico', [{
    type: 'scatter3d', mode: 'markers', x, y, z, text: t,
    hovertemplate: '%{text}<extra></extra>',
    marker: { size: s, color: c, colorscale: M.escala, cmin: 0, cmax: M.densidadMax,
      opacity: 0.9, line: { width: 0 },
      colorbar: { title: { text: 'Fallas' }, thickness: 14 } }
  }], layout, { responsive: true, displaylogo: false });
  $('resumen').textContent = `${nOt} OT · ${x.length} departamentos visibles · máx. ${max} fallas/depto`;
}

$('selEdificio').addEventListener('change', () => { actualizarPisos(); actualizar(); });
$('selPiso').addEventListener('change', actualizar);
$('desde').addEventListener('input', actualizar);
$('hasta').addEventListener('input', actualizar);

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
document.addEventListener('keydown', ev => { if (ev.key === 'Escape' && enFallback()) fallback(false); });

actualizarPisos();
actualizar();
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

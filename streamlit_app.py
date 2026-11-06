import altair as alt
import openpyxl
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from pathlib import Path

st.set_page_config(
    page_title="Town Hall 2026",
    page_icon=None,
    layout="wide",
    initial_sidebar_state="expanded",
)

BASE = Path(__file__).parent

EXCEL_PATH = BASE / "EXCEL" / "Town Hall Julio 2026.xlsx"
# Logo en negativo (marcas off-white, fondo transparente) para el sidebar navy.
LOGO_PATH = BASE / "imagenes" / "ancora_blanco.png"

MESES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun",
         "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]

# Proyección de ingreso para el cierre del año 2026 (dato provisto, editable).
PROY_ANUAL_2026 = 160_000_000

# ── Paleta de marca Ancora ──────────────────────────────────────────────────────
# Primarios: navy #13375c · off-white #F6F6F6 · teal #1d7b8a · dorado #e9ba40.
MARCA_NAVY = "#13375c"
MARCA_CLARO = "#F6F6F6"
MARCA_TEAL = "#1d7b8a"
MARCA_ORO = "#e9ba40"

# Paleta categórica por área, anclada en los primarios de marca y extendida con
# tonos armónicos para que las 9 áreas se distingan sobre fondo oscuro.
COLORES = {
    "BENEFICIOS":    MARCA_TEAL,   # teal (marca)
    "LF":            MARCA_ORO,    # dorado (marca)
    "DAÑOS":         "#3a6ea5",    # azul acero (navy aclarado)
    "FIANZAS":       "#d97b3c",    # terracota (complementa el dorado)
    "LP":            "#4aa3b0",    # teal claro
    "AUTOS":         "#8a6d4b",    # bronce
    "BONO":          "#6f9bc4",    # azul medio
    "ANI":           "#b58fb0",    # malva
    "AFFINITY":      "#9aa5b1",    # gris azulado (familia off-white)
    "HONORARIOS":    "#c99a2e",    # dorado oscuro
    "INTERNACIONAL": "#aec7e8",
}
# Series de comparación por año, en la familia de marca: gris (2024), navy acero
# (2025), dorado = meta/PTO y teal = real 2026 (la línea protagonista).
COLOR_2024 = "#9aa5b1"   # gris azulado · dotted
COLOR_2025 = "#3a6ea5"   # navy acero · dashed
COLOR_PTO  = MARCA_ORO   # dorado · dashdot (línea de meta)
COLOR_2026 = MARCA_TEAL  # teal · sólido (protagonista)
C_NEG = "#e05c5c"        # rojo (semántico: por debajo)
C_POS = MARCA_TEAL       # teal (semántico: a favor)

# ── Helpers ────────────────────────────────────────────────────────────────────
def fmt(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "—"
    return f"${v:,.0f}"

def fmt_m(v):
    # Importe exacto, sin redondear a millones ni a miles (mismo formato que fmt).
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "—"
    return f"${v:,.0f}"

def pct(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "—"
    return f"{'+' if v >= 0 else ''}{v*100:.1f}%"

def sf(v):
    return float(v) if isinstance(v, (int, float)) else 0.0

def norm(s):
    if not isinstance(s, str):
        return s
    return (s.replace("DA�OS", "DAÑOS").replace("DANOS", "DAÑOS")
             .replace("�REA", "AREA").replace("�", "Ñ").strip())

# ── Altair: tema oscuro de marca ────────────────────────────────────────────────
# Vega-Lite por defecto pinta fondo blanco y texto oscuro; el tablero vive sobre
# fondo navy (#0f2137, via config.toml), asi que forzamos fondo transparente y
# texto claro en ejes/leyenda/titulo, con la fuente de marca (Montserrat).
FUENTE = "Montserrat, sans-serif"
EJE_COLOR = "#c7d2e0"
GRID_COLOR = "#2e4d70"
DOMINIO_COLOR = "#3a5578"

def estilizar(chart):
    """Aplica el tema oscuro (fondo transparente, texto claro, Montserrat) a un chart Altair ya armado."""
    return (
        chart
        .properties(background="transparent")
        .configure_view(strokeOpacity=0)
        .configure_axis(labelColor=EJE_COLOR, titleColor=EJE_COLOR,
                        gridColor=GRID_COLOR, domainColor=DOMINIO_COLOR,
                        labelFont=FUENTE, titleFont=FUENTE,
                        labelFontSize=11, titleFontSize=12)
        .configure_legend(labelColor=EJE_COLOR, titleColor=EJE_COLOR,
                          labelFont=FUENTE, titleFont=FUENTE)
        .configure_title(color=MARCA_CLARO, font=FUENTE)
    )

EJE_DOLAR = alt.Axis(labelExpr="'$' + format(datum.value, ',.0f')")

# ── Carga ──────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="Cargando datos...")
def load_data():
    """Lee los bloques por año buscando los encabezados 'AREA', no filas fijas."""
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)

    def etiqueta(txt):
        s = str(txt).upper()
        if "PTO" in s or "PRESUP" in s:
            return "PTO"
        for y in ("2026", "2025", "2024"):
            if y in s:
                return y
        return None

    def leer_hoja(nombre):
        ws = wb[nombre]
        rows = list(ws.iter_rows(min_row=1, values_only=True))
        bloques = {}
        for i, r in enumerate(rows):
            if not r or r[0] is None:
                continue
            if norm(str(r[0])).upper() not in ("AREA", "ÁREA"):
                continue
            # la etiqueta del bloque vive en la fila anterior
            lab = None
            for back in (1, 2):
                if i - back >= 0 and rows[i - back] and rows[i - back][0] is not None:
                    lab = etiqueta(rows[i - back][0])
                    if lab:
                        break
            if not lab:
                continue
            datos = {}
            for r2 in rows[i + 1:]:
                if not r2 or r2[0] is None or not str(r2[0]).strip():
                    continue
                a = norm(str(r2[0])).upper()
                if a == "TOTAL":
                    break
                datos[a] = [sf(r2[c]) for c in range(1, 13)]
            if datos:
                bloques[lab] = datos
        return bloques

    ing = leer_hoja("Ingresos")
    vn = leer_hoja("VN")
    return ing, vn

try:
    ING, VN = load_data()
except Exception as e:
    st.error(f"Error al abrir el archivo: {e}")
    st.stop()

if not ING:
    st.error(f"No se encontraron bloques de datos en el archivo:\n\n`{EXCEL_PATH}`")
    st.stop()

# Áreas excluidas de todo el tablero.
EXCLUIR_AREAS = {"INTERNACIONAL", "ANI"}
for _grupo in (ING, VN):
    for _bloque in _grupo.values():
        for _a in list(_bloque):
            if _a in EXCLUIR_AREAS:
                del _bloque[_a]

def meses_con_datos(bloque):
    """Ultimo mes con dato real (para comparar YTD contra el mismo periodo)."""
    if not bloque:
        return 0
    ult = 0
    for m in range(12):
        if sum(v[m] for v in bloque.values()) > 0:
            ult = m + 1
    return ult

N_MES = meses_con_datos(ING.get("2026", {}))
PERIODO = f"{MESES[0]}–{MESES[N_MES-1]}" if N_MES else "sin datos"

def total(bloque, hasta=12):
    return sum(sum(v[:hasta]) for v in bloque.values())

def por_area(bloque, hasta=12):
    return {a: sum(v[:hasta]) for a, v in bloque.items()}

def serie_mensual(bloque, areas=None):
    ms = [0.0] * 12
    for a, v in bloque.items():
        if areas is not None and a not in areas:
            continue
        for m in range(12):
            ms[m] += v[m]
    return ms

# ── Componentes de gráficas (Altair) ────────────────────────────────────────────
def grafico_linea_comparativa(series_base, bloque_real, n_mes, height):
    """Línea mensual con 3 series de referencia (punteada/discontinua/dashdot) más
    la serie 'Real 2026', gruesa y sólida, cortada en el último mes con dato real.
    series_base: lista de (nombre, bloque, color, patrón_de_raya)."""
    orden = [s[0] for s in series_base] + (["Real 2026"] if bloque_real else [])
    colores = [s[2] for s in series_base] + ([COLOR_2026] if bloque_real else [])
    rayas = [s[3] for s in series_base] + ([[]] if bloque_real else [])

    filas = []
    for nombre, bloque, color, dash in series_base:
        if not bloque:
            continue
        for m, v in zip(MESES, serie_mensual(bloque)):
            filas.append({"Serie": nombre, "Mes": m, "Valor": v, "ValorTxt": fmt(v)})
    if bloque_real:
        for m, v in zip(MESES[:n_mes], serie_mensual(bloque_real)[:n_mes]):
            filas.append({"Serie": "Real 2026", "Mes": m, "Valor": v, "ValorTxt": fmt(v)})
    df = pd.DataFrame(filas)

    base = alt.Chart(df).encode(
        x=alt.X("Mes:N", sort=MESES, title=None, axis=alt.Axis(labelAngle=0)),
        y=alt.Y("Valor:Q", title=None, axis=EJE_DOLAR),
        color=alt.Color("Serie:N", sort=orden,
                         scale=alt.Scale(domain=orden, range=colores),
                         legend=alt.Legend(title=None, orient="top")),
        strokeDash=alt.StrokeDash("Serie:N", sort=orden,
                                   scale=alt.Scale(domain=orden, range=rayas),
                                   legend=alt.Legend(title=None, orient="top")),
        tooltip=[alt.Tooltip("Serie:N", title="Serie"), alt.Tooltip("Mes:N", title="Mes"),
                 alt.Tooltip("ValorTxt:N", title="Monto")],
    )
    finas = base.transform_filter(alt.datum.Serie != "Real 2026").mark_line(
        strokeWidth=2, interpolate="monotone")
    gruesa = base.transform_filter(alt.datum.Serie == "Real 2026").mark_line(
        strokeWidth=4, interpolate="monotone")
    puntos = base.mark_point(filled=True, size=38, opacity=0.9)
    return (finas + gruesa + puntos).properties(height=height)

def grafico_donut(vals, height=460):
    """Donut con etiqueta 'Área · NN%' alrededor, coloreado por área (paleta COLORES)."""
    areas = list(vals.keys())
    total_v = sum(vals.values()) or 1
    df = pd.DataFrame({"Área": areas, "Valor": [vals[a] for a in areas]})
    df["Porcentaje"] = df["Valor"] / total_v
    df["ValorTxt"] = df["Valor"].map(fmt)
    df["Etiqueta"] = [f"{a} · {p*100:.1f}%" for a, p in zip(df["Área"], df["Porcentaje"])]
    # Rebanadas muy chicas (< 3%) se saltan la etiqueta: a ese tamaño el texto de
    # vecinos se encima (sí quedan disponibles al pasar el mouse via tooltip).
    df["MostrarEtiqueta"] = df["Porcentaje"] >= 0.03
    escala_color = alt.Scale(domain=areas, range=[COLORES.get(a, "#999") for a in areas])

    base = alt.Chart(df).encode(
        theta=alt.Theta("Valor:Q", stack=True, sort="descending"),
        order=alt.Order("Valor:Q", sort="descending"),
        tooltip=[alt.Tooltip("Área:N"), alt.Tooltip("ValorTxt:N", title="Monto"),
                 alt.Tooltip("Porcentaje:Q", title="% del total", format=".1%")],
    )
    arcos = base.mark_arc(innerRadius=height * 0.26, outerRadius=height * 0.42,
                           cornerRadius=2, stroke="#0f2137", strokeWidth=1).encode(
        color=alt.Color("Área:N", scale=escala_color, legend=None),
    )
    etiquetas = base.transform_filter(alt.datum.MostrarEtiqueta).mark_text(
        radius=height * 0.49, fontSize=11, color=EJE_COLOR
    ).encode(
        text="Etiqueta:N",
    )
    return (arcos + etiquetas).properties(height=height)

# ── Estilos (mismos que los otros tableros) ────────────────────────────────────
st.markdown("""
<style>
/* Fuente de marca Ancora: Montserrat. El @import debe ir primero en el <style>. */
@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600;700&display=swap');
/* Montserrat en todo EXCEPTO los iconos Material, que llevan su propia fuente. */
html, body, [data-testid="stAppViewContainer"], [data-testid="stSidebar"], .stApp,
.stApp *:not([data-testid="stIconMaterial"]) {
    font-family: 'Montserrat', sans-serif !important;
}
[data-testid="stIconMaterial"] { font-family: 'Material Symbols Rounded' !important; }
/* Cuerpo sobre fondo oscuro: Medium (500). */
[data-testid="stAppViewContainer"], [data-testid="stMarkdownContainer"],
p, span:not([data-testid="stIconMaterial"]), label, td, th, li,
div[data-testid="stMetricValue"], div[data-testid="stMetricDelta"], .stDataFrame { font-weight: 500; }
/* Títulos: Semibold (600). */
h1, h2, h3, h4, h5, h6, div[data-testid="stMetricLabel"] p { font-weight: 600 !important; }

section[data-testid="stSidebar"] img { pointer-events: none; }
section[data-testid="stSidebar"] [data-testid="StyledFullScreenButton"] { display: none; }
section[data-testid="stSidebar"] > div:first-child { padding-top: 0 !important; }
section[data-testid="stSidebar"] .block-container { padding-top: 0 !important; }
div[data-testid="stRadio"] > div { gap: 2px !important; }
div[data-testid="stRadio"] label {
    display: flex !important; align-items: center !important;
    padding: 10px 14px !important; border-radius: 8px !important;
    cursor: pointer !important; font-size: 0.9rem !important;
    margin-bottom: 2px !important; transition: background 0.15s !important; border: none !important;
}
div[data-testid="stRadio"] label:hover { background: rgba(255,255,255,0.08) !important; }
div[data-testid="stRadio"] label:has(input:checked) {
    background: rgba(255,255,255,0.13) !important; font-weight: 600 !important;
}
div[data-testid="stRadio"] label > div:first-child { display: none !important; }
div[data-testid="stRadio"] { border: none !important; }
/* Montos exactos son largos: fuente menor y sin corte para que quepan completos. */
div[data-testid="stMetricValue"] {
    font-size: 1.35rem !important; line-height: 1.2 !important;
    white-space: nowrap !important; overflow: visible !important;
}
div[data-testid="stMetricValue"] > div { overflow: visible !important; }
/* Tarjetas de la misma altura: la columna ya se estira a la más alta; hago que la
   tarjeta (wrapper + bloque con borde) llene esa altura en vez de quedarse corta. */
div[data-testid="stColumn"] > div[data-testid="stVerticalBlock"] > div[data-testid="stLayoutWrapper"],
div[data-testid="stColumn"] > div[data-testid="stVerticalBlock"] > div[data-testid="stLayoutWrapper"] > div[data-testid="stVerticalBlock"] {
    height: 100% !important;
}
</style>
""", unsafe_allow_html=True)

# Evita que el traductor del navegador reformatee los montos ("$18.6M" -> "18,6 millones").
components.html(
    "<script>try{var d=window.parent.document.documentElement;"
    "d.setAttribute('translate','no');d.classList.add('notranslate');}catch(e){}</script>",
    height=0,
)

with st.sidebar:
    if LOGO_PATH.exists():
        st.image(str(LOGO_PATH), use_container_width=True)
    st.markdown("---")
    st.caption("TABLEROS")
    pagina = st.radio(
        "",
        ["Panorama", "Ingreso Semestral", "Venta Nueva"],
        label_visibility="collapsed",
        key="nav",
    )
    st.markdown("---")
    st.caption(f"Town Hall · Datos a {MESES[N_MES-1] if N_MES else '—'} 2026")

# ── Series base ────────────────────────────────────────────────────────────────
B26, B25, B24, BPTO = ING.get("2026", {}), ING.get("2025", {}), ING.get("2024", {}), ING.get("PTO", {})
V26, V25, V24, VPTO = VN.get("2026", {}), VN.get("2025", {}), VN.get("2024", {}), VN.get("PTO", {})

# Comparaciones contra años completos: lo acumulado de 2026 se mide contra el
# total anual de 2025, 2024 y el presupuesto. Es la vista de "avance del año".
ING_2026  = total(B26)
ING_2025  = total(B25)
ING_2024  = total(B24)
PTO_ANUAL = total(BPTO)

AREAS = sorted(
    {a for b in (B26, B25, B24, BPTO) for a in b},
    key=lambda a: -sum(B26.get(a, [0]*12)[:N_MES] or [0]),
)

# ══════════════════════════════════════════════════════════════════════════════
# PANORAMA
# ══════════════════════════════════════════════════════════════════════════════
if pagina == "Panorama":
    st.title("Town Hall · 2026")
    st.caption("Ingresos y venta nueva por área · Comparativo contra años completos · Importes en MXN")

    if N_MES < 12:
        st.info(f"2026 lleva **{N_MES} de 12 meses** ({PERIODO}). El crecimiento de 2026 se mide "
                f"contra el **mismo periodo** de 2025 ({PERIODO}); el de 2025 es año completo vs 2024.")

    # Crecimiento: 2025 completo vs 2024 completo; 2026 (parcial) vs el mismo periodo de 2025.
    ing_25_ytd = total(B25, N_MES)
    g_ing_25 = ING_2025 / ING_2024 - 1 if ING_2024 else None
    g_ing_26 = ING_2026 / ing_25_ytd - 1 if ing_25_ytd else None
    vn_24, vn_25, vn_26 = total(V24), total(V25), total(V26)
    vn_25_ytd = total(V25, N_MES)
    g_vn_25 = vn_25 / vn_24 - 1 if vn_24 else None
    g_vn_26 = vn_26 / vn_25_ytd - 1 if vn_25_ytd else None

    g2024, g2025, g2026 = st.columns(3)

    with g2024:
        with st.container(border=True):
            st.markdown("##### 2024")
            a, b = st.columns(2)
            a.metric("Ingreso", fmt_m(ING_2024),
                     delta="año completo", delta_color="off")
            b.metric("Venta Nueva", fmt_m(vn_24),
                     delta="año completo", delta_color="off")

    with g2025:
        with st.container(border=True):
            st.markdown("##### 2025")
            a, b = st.columns(2)
            a.metric("Ingreso", fmt_m(ING_2025),
                     delta=f"{pct(g_ing_25)} vs 2024" if g_ing_25 is not None else None)
            b.metric("Venta Nueva", fmt_m(vn_25),
                     delta=f"{pct(g_vn_25)} vs 2024" if g_vn_25 is not None else None)

    with g2026:
        with st.container(border=True):
            st.markdown(f"##### 2026 · {N_MES} de 12 meses")
            a, b = st.columns(2)
            a.metric("Ingreso", fmt_m(ING_2026),
                     delta=f"{pct(g_ing_26)} vs 2025 (mismo periodo)" if g_ing_26 is not None else None)
            b.metric("Venta Nueva", fmt_m(vn_26),
                     delta=f"{pct(g_vn_26)} vs 2025 (mismo periodo)" if g_vn_26 is not None else None)
            st.caption(f"Proyección de ingreso al cierre del año: **{fmt(PROY_ANUAL_2026)}**")

    st.divider()
    st.subheader("Ingreso mensual: 2026 vs 2025 vs 2024")
    chart_ing_mensual = grafico_linea_comparativa(
        [("2024", B24, COLOR_2024, [2, 2]),
         ("2025", B25, COLOR_2025, [6, 3]),
         ("PTO 2026", BPTO, COLOR_PTO, [6, 2, 2, 2])],
        B26, N_MES, 430,
    )
    st.altair_chart(estilizar(chart_ing_mensual), use_container_width=True)

    st.divider()
    col_mix1, col_mix2 = st.columns(2)
    with col_mix1:
        st.subheader(f"Mix de ingreso 2026 ({N_MES}m)")
        vals = {a: v for a, v in por_area(B26, N_MES).items() if v > 0}
        st.altair_chart(estilizar(grafico_donut(vals, 460)), use_container_width=True)
    with col_mix2:
        st.subheader("Mix de presupuesto 2026")
        vals_pto = {a: v for a, v in por_area(BPTO).items() if v > 0}
        st.altair_chart(estilizar(grafico_donut(vals_pto, 460)), use_container_width=True)

    st.divider()
    st.subheader("Por área: 2024 vs 2025 vs PTO 2026 vs Real 2026")
    st.caption(f"2024 y 2025 son año completo y PTO 2026 es el presupuesto anual; "
               f"Real 2026 es lo acumulado ({PERIODO}, {N_MES} de 12 meses).")
    a24, a25, apto, areal = por_area(B24), por_area(B25), por_area(BPTO), por_area(B26)
    areas_y = sorted(
        [a for a in AREAS if any(d.get(a, 0) > 0 for d in (a24, a25, apto, areal))],
        key=lambda a: -apto.get(a, 0),
    )
    a25_ytd = por_area(B25, N_MES)   # 2025 al mismo periodo que el Real (para su crecimiento)

    def _crece(nuevo, viejo, suf):
        """Sufijo con el % de crecimiento vs el año/periodo de referencia."""
        if not viejo:
            return ""
        g = nuevo / viejo - 1
        return f" · {'+' if g >= 0 else ''}{g*100:.0f}% {suf}"

    def _texto_real(a):
        """Monto real + crecimiento vs 2025 (mismo periodo) + % que falta al PTO."""
        real, pto = areal.get(a, 0), apto.get(a, 0)
        txt = fmt_m(real) + _crece(real, a25_ytd.get(a, 0), "vs 25")
        if pto > 0:
            falta = (pto - real) / pto
            txt += " · meta cumplida" if falta <= 0 else f" · falta {falta*100:.0f}% PTO"
        return txt

    # (etiqueta, datos, color, referencia_para_crecimiento, sufijo)
    series = [("2024", a24, COLOR_2024, None, ""),
              ("2025", a25, COLOR_2025, a24, "vs 24"),
              ("PTO 2026", apto, COLOR_PTO, a25, "vs 25"),
              (f"Real 2026 ({N_MES}m)", areal, COLOR_2026, None, "")]
    orden_cmp = [s[0] for s in series]
    colores_cmp = [s[2] for s in series]

    filas_cmp = []
    for etq, d, col, ref, suf in series:
        es_real = d is areal
        for a in areas_y:
            v = d.get(a, 0)
            if es_real:
                txt = _texto_real(a)
            elif ref is not None:
                txt = fmt_m(v) + _crece(v, ref.get(a, 0), suf)
            else:  # 2024 es la base, sin crecimiento
                txt = fmt_m(v)
            filas_cmp.append({"Área": a, "Serie": etq, "Valor": v,
                              "ValorMedio": v / 2, "Texto": txt, "EsReal": es_real})
    df_cmp = pd.DataFrame(filas_cmp)
    # Dominio con margen extra a la derecha para que la etiqueta "por fuera" del
    # Real (puede incluir "· falta NN% PTO") no se corte contra el borde del chart.
    max_cmp = df_cmp["Valor"].max() or 1

    base_cmp = alt.Chart(df_cmp).encode(
        y=alt.Y("Área:N", sort=areas_y, title=None),
        yOffset=alt.YOffset("Serie:N", sort=orden_cmp),
        x=alt.X("Valor:Q", title=None, scale=alt.Scale(domain=[0, max_cmp * 1.6]),
                axis=EJE_DOLAR),
        color=alt.Color("Serie:N", sort=orden_cmp,
                         scale=alt.Scale(domain=orden_cmp, range=colores_cmp),
                         legend=alt.Legend(title=None, orient="top")),
        tooltip=[alt.Tooltip("Área:N"), alt.Tooltip("Serie:N"),
                 alt.Tooltip("Texto:N", title="Detalle")],
    )
    barras_cmp = base_cmp.mark_bar(cornerRadiusEnd=2)
    texto_dentro = base_cmp.transform_filter(~alt.datum.EsReal).mark_text(
        align="center", baseline="middle", fontSize=10,
    ).encode(x=alt.X("ValorMedio:Q"), text="Texto:N", color=alt.value("#111111"))
    texto_fuera = base_cmp.transform_filter(alt.datum.EsReal).mark_text(
        align="left", baseline="middle", fontSize=10, dx=6,
    ).encode(x=alt.X("Valor:Q"), text="Texto:N", color=alt.value("#e8e8e8"))
    chart_cmp = (barras_cmp + texto_dentro + texto_fuera).properties(height=760)
    st.altair_chart(estilizar(chart_cmp), use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# INGRESOS MENSUAL
# ══════════════════════════════════════════════════════════════════════════════
if pagina == "Ingreso Semestral":
    SEM = 6
    def h1(bl, a):
        return sum(bl.get(a, [0] * 12)[:SEM])

    # (etiqueta, bloque, color, bloque del año anterior para el % de crecimiento)
    series = [("1er sem 2024", B24, COLOR_2024, None),
              ("1er sem 2025", B25, COLOR_2025, B24),
              ("1er sem 2026", B26, COLOR_2026, B25)]
    areas_sem = [a for a in AREAS
                 if any(h1(bl, a) > 0 for _, bl, _, _ in series)]

    # ── Carrusel de crecimiento por línea de negocio (de 2 en 2) ────────────────
    st.subheader("Crecimiento por línea de negocio · 1er semestre")
    PASO = 2
    n = len(areas_sem)
    total_pag = max(1, (n + PASO - 1) // PASO)
    if "ing_pag" not in st.session_state:
        st.session_state.ing_pag = 0

    c_prev, c_lbl, c_next = st.columns([1, 6, 1])
    if c_prev.button("◀", key="ing_prev", use_container_width=True):
        st.session_state.ing_pag = (st.session_state.ing_pag - 1) % total_pag
    if c_next.button("▶", key="ing_next", use_container_width=True):
        st.session_state.ing_pag = (st.session_state.ing_pag + 1) % total_pag
    pag = st.session_state.ing_pag % total_pag
    c_lbl.markdown(
        f"<div style='text-align:center;padding-top:8px;color:#9aa0a6'>"
        f"Líneas {pag*PASO+1}–{min((pag+1)*PASO, n)} de {n}</div>",
        unsafe_allow_html=True,
    )

    fila = st.columns(PASO)
    for col, a in zip(fila, areas_sem[pag*PASO:(pag+1)*PASO]):
        s24, s25, s26 = h1(B24, a), h1(B25, a), h1(B26, a)
        with col, st.container(border=True):
            st.markdown(f"##### {a}")
            m1, m2, m3 = st.columns(3)
            m1.metric("2024", fmt_m(s24))
            m2.metric("2025 vs 24", fmt_m(s25),
                      delta=pct(s25 / s24 - 1) if s24 > 0 else None)
            m3.metric("2026 vs 25", fmt_m(s26),
                      delta=pct(s26 / s25 - 1) if s25 > 0 else None)

    st.divider()
    st.subheader("Ingreso por área · primer semestre por año")
    orden_sem = [s[0] for s in series]
    colores_sem = [s[2] for s in series]
    filas_sem = []
    for etq, bl, col, prev in series:
        for a in areas_sem:
            v = h1(bl, a)
            txt = ""
            if prev is not None:
                pv = h1(prev, a)
                txt = pct(v / pv - 1) if pv > 0 else ""
            filas_sem.append({"Área": a, "Serie": etq, "Valor": v, "Texto": txt})
    df_sem = pd.DataFrame(filas_sem)

    base_sem = alt.Chart(df_sem).encode(
        x=alt.X("Área:N", sort=areas_sem, title=None,
                axis=alt.Axis(labelAngle=-30, labelOverlap=False)),
        xOffset=alt.XOffset("Serie:N", sort=orden_sem),
        y=alt.Y("Valor:Q", title=None, axis=EJE_DOLAR),
        color=alt.Color("Serie:N", sort=orden_sem,
                         scale=alt.Scale(domain=orden_sem, range=colores_sem),
                         legend=alt.Legend(title=None, orient="top")),
        tooltip=[alt.Tooltip("Área:N"), alt.Tooltip("Serie:N"),
                 alt.Tooltip("Valor:Q", title="Monto", format=",.0f")],
    )
    barras_sem = base_sem.mark_bar(cornerRadiusEnd=2)
    texto_sem = base_sem.transform_filter(alt.datum.Texto != "").mark_text(
        align="center", baseline="bottom", dy=-3, fontSize=10,
    ).encode(text="Texto:N", color=alt.value(EJE_COLOR))
    chart_sem = (barras_sem + texto_sem).properties(height=460)
    st.altair_chart(estilizar(chart_sem), use_container_width=True)

    st.divider()
    st.subheader("Real 2026 vs Presupuesto anual por área")
    st.caption(f"Real acumulado ({PERIODO}, {N_MES} de 12 meses) contra el presupuesto anual.")
    real_a, pto_a = por_area(B26), por_area(BPTO)
    areas_cmp = [a for a in AREAS if real_a.get(a, 0) > 0 or pto_a.get(a, 0) > 0]
    orden_rp = ["PTO", "Real"]
    colores_rp = [COLOR_PTO, COLOR_2026]
    filas_rp = []
    for etq, d in [("PTO", pto_a), ("Real", real_a)]:
        for a in areas_cmp:
            filas_rp.append({"Área": a, "Serie": etq, "Valor": d.get(a, 0),
                             "ValorTxt": fmt_m(d.get(a, 0))})
    df_rp = pd.DataFrame(filas_rp)
    base_rp = alt.Chart(df_rp).encode(
        x=alt.X("Área:N", sort=areas_cmp, title=None,
                axis=alt.Axis(labelAngle=-30, labelOverlap=False)),
        xOffset=alt.XOffset("Serie:N", sort=orden_rp),
        y=alt.Y("Valor:Q", title=None, axis=EJE_DOLAR),
        color=alt.Color("Serie:N", sort=orden_rp,
                         scale=alt.Scale(domain=orden_rp, range=colores_rp),
                         legend=alt.Legend(title=None, orient="top")),
        tooltip=[alt.Tooltip("Área:N"), alt.Tooltip("Serie:N"),
                 alt.Tooltip("ValorTxt:N", title="Monto")],
    )
    chart_rp = base_rp.mark_bar(cornerRadiusEnd=2).properties(height=420)
    st.altair_chart(estilizar(chart_rp), use_container_width=True)

    # La tabla de detalle muestra el bloque 2026 Real.
    año = "2026"
    bloque = B26
    hasta = N_MES
    areas_b = [a for a in AREAS if a in bloque and sum(bloque[a]) > 0]

    st.divider()
    st.subheader(f"Detalle mensual 2026 · {PERIODO}")
    filas = []
    for a in areas_b:
        fila = {"Área": a}
        for m in range(hasta):
            fila[MESES[m]] = fmt(bloque[a][m])
        fila["Total"] = fmt(sum(bloque[a][:hasta]))
        filas.append(fila)
    tot_fila = {"Área": "TOTAL"}
    for m in range(hasta):
        tot_fila[MESES[m]] = fmt(sum(bloque[a][m] for a in areas_b))
    tot_fila["Total"] = fmt(sum(sum(bloque[a][:hasta]) for a in areas_b))
    filas.append(tot_fila)
    st.dataframe(pd.DataFrame(filas), use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# VENTA NUEVA
# ══════════════════════════════════════════════════════════════════════════════
if pagina == "Venta Nueva":
    # Vars anuales que siguen usando la gráfica y la tabla de abajo.
    vn_ytd = total(V26)
    vnpto_anual = total(VPTO)
    vn25, vn24 = total(V25), total(V24)

    # Primer semestre (Ene–Jun) para las tarjetas.
    SEM = 6
    vs24, vs25 = total(V24, SEM), total(V25, SEM)
    vspto, vsreal = total(VPTO, SEM), total(V26, SEM)
    with st.container(border=True):
        st.markdown("##### VENTA NUEVA · 1er semestre")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("1er sem 2024", fmt_m(vs24))
        c2.metric("2025 vs 24", fmt_m(vs25),
                  delta=pct(vs25 / vs24 - 1) if vs24 else None)
        c3.metric("Real 2026 vs 25", fmt_m(vsreal),
                  delta=pct(vsreal / vs25 - 1) if vs25 else None)
        c4.metric("PTO 2026", fmt_m(vspto),
                  delta=f"avance {vsreal/vspto*100:.0f}% del real" if vspto else None,
                  delta_color="off")

    st.divider()
    st.subheader("Venta Nueva mensual")
    chart_vn_mensual = grafico_linea_comparativa(
        [("2024", V24, COLOR_2024, [2, 2]),
         ("2025", V25, COLOR_2025, [6, 3]),
         ("PTO 2026", VPTO, COLOR_PTO, [6, 2, 2, 2])],
        V26, N_MES, 420,
    )
    st.altair_chart(estilizar(chart_vn_mensual), use_container_width=True)

    st.divider()
    col_a, col_b = st.columns(2)
    vn_area = {a: v for a, v in por_area(V26).items() if v > 0}
    vnpto_area = por_area(VPTO)
    with col_a:
        st.subheader("VN por área · acumulado 2026")
        areas_vn_acc = list(vn_area.keys())
        df_vn_acc = pd.DataFrame({"Área": areas_vn_acc,
                                  "Valor": [vn_area[a] for a in areas_vn_acc]})
        df_vn_acc["ValorTxt"] = df_vn_acc["Valor"].map(fmt_m)
        escala_vn_acc = alt.Scale(domain=areas_vn_acc,
                                   range=[COLORES.get(a, "#999") for a in areas_vn_acc])
        max_vn_acc = df_vn_acc["Valor"].max() or 1
        base_vn_acc = alt.Chart(df_vn_acc).encode(
            y=alt.Y("Área:N", sort=areas_vn_acc, title=None),
            x=alt.X("Valor:Q", title=None, scale=alt.Scale(domain=[0, max_vn_acc * 1.3]),
                    axis=EJE_DOLAR),
            color=alt.Color("Área:N", scale=escala_vn_acc, legend=None),
            tooltip=[alt.Tooltip("Área:N"), alt.Tooltip("ValorTxt:N", title="Monto")],
        )
        barras_vn_acc = base_vn_acc.mark_bar(cornerRadiusEnd=2)
        texto_vn_acc = base_vn_acc.mark_text(
            align="left", baseline="middle", dx=6, fontSize=10,
        ).encode(x="Valor:Q", text="ValorTxt:N", color=alt.value(EJE_COLOR))
        chart_a = (barras_vn_acc + texto_vn_acc).properties(height=380)
        st.altair_chart(estilizar(chart_a), use_container_width=True)
    with col_b:
        st.subheader("VN acumulado vs PTO anual por área")
        areas_vn = [a for a in vnpto_area if vnpto_area[a] > 0 or vn_area.get(a, 0) > 0]
        orden_vn2 = ["PTO", "Real"]
        colores_vn2 = [COLOR_PTO, COLOR_2026]
        # % de avance (Real ÷ PTO) fijo sobre cada barra teal (Real).
        filas_vn2 = []
        for etq, d in [("PTO", vnpto_area), ("Real", vn_area)]:
            for a in areas_vn:
                v = d.get(a, 0)
                txt = (f"{vn_area.get(a, 0)/vnpto_area[a]*100:.0f}%"
                       if etq == "Real" and vnpto_area.get(a) else "")
                filas_vn2.append({"Área": a, "Serie": etq, "Valor": v,
                                  "ValorTxt": fmt_m(v), "Texto": txt})
        df_vn2 = pd.DataFrame(filas_vn2)
        base_vn2 = alt.Chart(df_vn2).encode(
            x=alt.X("Área:N", sort=areas_vn, title=None,
                    axis=alt.Axis(labelAngle=-45, labelOverlap=False, labelFontSize=10)),
            xOffset=alt.XOffset("Serie:N", sort=orden_vn2),
            y=alt.Y("Valor:Q", title=None, axis=EJE_DOLAR),
            color=alt.Color("Serie:N", sort=orden_vn2,
                             scale=alt.Scale(domain=orden_vn2, range=colores_vn2),
                             legend=alt.Legend(title=None, orient="top")),
            tooltip=[alt.Tooltip("Área:N"), alt.Tooltip("Serie:N"),
                     alt.Tooltip("ValorTxt:N", title="Monto")],
        )
        barras_vn2 = base_vn2.mark_bar(cornerRadiusEnd=2)
        texto_vn2 = base_vn2.transform_filter(alt.datum.Texto != "").mark_text(
            align="center", baseline="bottom", dy=-3, fontSize=11,
        ).encode(text="Texto:N", color=alt.value("#e0e0e0"))
        chart_c = (barras_vn2 + texto_vn2).properties(height=380)
        st.altair_chart(estilizar(chart_c), use_container_width=True)

    filas = []
    for a in sorted(set(vn_area) | set(k for k, v in vnpto_area.items() if v > 0),
                    key=lambda x: -vn_area.get(x, 0)):
        r, p = vn_area.get(a, 0), vnpto_area.get(a, 0)
        filas.append({"Área": a, "Real acumulado": fmt(r), "PTO anual": fmt(p),
                      "Falta": fmt(r - p),
                      "Avance": pct(r/p) if p else "—"})
    filas.append({"Área": "TOTAL", "Real acumulado": fmt(vn_ytd),
                  "PTO anual": fmt(vnpto_anual),
                  "Falta": fmt(vn_ytd - vnpto_anual),
                  "Avance": pct(vn_ytd/vnpto_anual) if vnpto_anual else "—"})
    st.dataframe(pd.DataFrame(filas), use_container_width=True, hide_index=True)

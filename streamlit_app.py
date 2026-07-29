import openpyxl
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
import streamlit.components.v1 as components
from pathlib import Path

st.set_page_config(
    page_title="Town Hall 2026",
    page_icon=None,
    layout="wide",
    initial_sidebar_state="expanded",
)

# Rutas relativas al propio archivo: funcionan igual en Windows y en Linux (la nube).
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
    fig = go.Figure()
    for nombre, bloque, color, dash in [
        ("2024", B24, COLOR_2024, "dot"),
        ("2025", B25, COLOR_2025, "dash"),
        ("PTO 2026", BPTO, COLOR_PTO, "dashdot"),
    ]:
        if bloque:
            fig.add_scatter(name=nombre, x=MESES, y=serie_mensual(bloque),
                            mode="lines+markers",
                            line=dict(color=color, width=2, dash=dash))
    if B26:
        s26 = serie_mensual(B26)[:N_MES]
        fig.add_scatter(name="Real 2026", x=MESES[:N_MES], y=s26,
                        mode="lines+markers",
                        line=dict(color=COLOR_2026, width=4))
    fig.update_layout(yaxis_tickformat="$,.0f", height=430,
                      margin=dict(t=30, b=20),
                      legend=dict(orientation="h", y=1.1))
    st.plotly_chart(fig, use_container_width=True)

    st.divider()
    col_mix1, col_mix2 = st.columns(2)
    with col_mix1:
        st.subheader(f"Mix de ingreso 2026 ({N_MES}m)")
        vals = {a: v for a, v in por_area(B26, N_MES).items() if v > 0}
        fig_p = go.Figure(go.Pie(
            labels=list(vals.keys()), values=list(vals.values()),
            marker_colors=[COLORES.get(a, "#999") for a in vals],
            hole=0.42, textinfo="label+percent", sort=True,
        ))
        fig_p.update_layout(height=460, margin=dict(t=20, b=20), showlegend=False)
        st.plotly_chart(fig_p, use_container_width=True)
    with col_mix2:
        st.subheader("Mix de presupuesto 2026")
        vals_pto = {a: v for a, v in por_area(BPTO).items() if v > 0}
        fig_ppto = go.Figure(go.Pie(
            labels=list(vals_pto.keys()), values=list(vals_pto.values()),
            marker_colors=[COLORES.get(a, "#999") for a in vals_pto],
            hole=0.42, textinfo="label+percent", sort=True,
        ))
        fig_ppto.update_layout(height=460, margin=dict(t=20, b=20), showlegend=False)
        st.plotly_chart(fig_ppto, use_container_width=True)

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

    fig_cmp = go.Figure()
    for etq, d, col, ref, suf in series:
        es_real = d is areal
        if es_real:
            textos = [_texto_real(a) for a in areas_y]
        elif ref is not None:
            textos = [fmt_m(d.get(a, 0)) + _crece(d.get(a, 0), ref.get(a, 0), suf)
                      for a in areas_y]
        else:  # 2024 es la base, sin crecimiento
            textos = [fmt_m(d.get(a, 0)) for a in areas_y]
        fig_cmp.add_bar(
            name=etq, orientation="h", y=areas_y,
            x=[d.get(a, 0) for a in areas_y], marker_color=col,
            text=textos,
            # El Real lleva su etiqueta fija por fuera; las demas por dentro.
            textposition="outside" if es_real else "inside",
            insidetextanchor="middle",
            textfont=dict(size=10, color=("#e0e0e0" if es_real else "#111")),
            cliponaxis=False,
            constraintext="none" if es_real else "both",
        )
    fig_cmp.update_layout(barmode="group", height=760, xaxis_tickformat="$,.0f",
                          yaxis=dict(autorange="reversed"),
                          margin=dict(t=30, b=20, r=210),
                          uniformtext=dict(minsize=8, mode="show"),
                          legend=dict(orientation="h", y=1.06))
    st.plotly_chart(fig_cmp, use_container_width=True)

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
    fig = go.Figure()
    for etq, bl, col, prev in series:
        y = [h1(bl, a) for a in areas_sem]
        txt = None
        if prev is not None:
            txt = [pct(h1(bl, a) / h1(prev, a) - 1) if h1(prev, a) > 0 else ""
                   for a in areas_sem]
        fig.add_bar(name=etq, x=areas_sem, y=y, marker_color=col,
                    text=txt, textposition="outside",
                    textfont=dict(size=10), cliponaxis=False)
    fig.update_layout(barmode="group", bargap=0.25, bargroupgap=0.05,
                      yaxis_tickformat="$,.0f", height=460,
                      margin=dict(t=50, b=20),
                      legend=dict(orientation="h", y=1.12))
    st.plotly_chart(fig, use_container_width=True)

    st.divider()
    st.subheader("Real 2026 vs Presupuesto anual por área")
    st.caption(f"Real acumulado ({PERIODO}, {N_MES} de 12 meses) contra el presupuesto anual.")
    real_a, pto_a = por_area(B26), por_area(BPTO)
    areas_cmp = [a for a in AREAS if real_a.get(a, 0) > 0 or pto_a.get(a, 0) > 0]
    fig_rp = go.Figure()
    fig_rp.add_bar(name="PTO", x=areas_cmp, y=[pto_a.get(a, 0) for a in areas_cmp],
                   marker_color=COLOR_PTO)
    fig_rp.add_bar(name="Real", x=areas_cmp, y=[real_a.get(a, 0) for a in areas_cmp],
                   marker_color=COLOR_2026)
    fig_rp.update_layout(barmode="group", yaxis_tickformat="$,.0f", height=420,
                         margin=dict(t=30, b=20),
                         legend=dict(orientation="h", y=1.1))
    st.plotly_chart(fig_rp, use_container_width=True)

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
    fig = go.Figure()
    for nombre, bloque, color, dash in [
        ("2024", V24, COLOR_2024, "dot"),
        ("2025", V25, COLOR_2025, "dash"),
        ("PTO 2026", VPTO, COLOR_PTO, "dashdot"),
    ]:
        if bloque:
            fig.add_scatter(name=nombre, x=MESES, y=serie_mensual(bloque),
                            mode="lines+markers",
                            line=dict(color=color, width=2, dash=dash))
    if V26:
        fig.add_scatter(name="Real 2026", x=MESES[:N_MES],
                        y=serie_mensual(V26)[:N_MES],
                        mode="lines+markers", line=dict(color=COLOR_2026, width=4))
    fig.update_layout(yaxis_tickformat="$,.0f", height=420,
                      margin=dict(t=30, b=20),
                      legend=dict(orientation="h", y=1.1))
    st.plotly_chart(fig, use_container_width=True)

    st.divider()
    col_a, col_b = st.columns(2)
    vn_area = {a: v for a, v in por_area(V26).items() if v > 0}
    vnpto_area = por_area(VPTO)
    with col_a:
        st.subheader("VN por área · acumulado 2026")
        fig_a = go.Figure(go.Bar(
            x=list(vn_area.values()), y=list(vn_area.keys()), orientation="h",
            marker_color=[COLORES.get(a, "#999") for a in vn_area],
            text=[fmt_m(v) for v in vn_area.values()], textposition="auto",
        ))
        fig_a.update_layout(height=380, xaxis_tickformat="$,.0f",
                            yaxis=dict(autorange="reversed"), margin=dict(t=20, b=20))
        st.plotly_chart(fig_a, use_container_width=True)
    with col_b:
        st.subheader("VN acumulado vs PTO anual por área")
        areas_vn = [a for a in vnpto_area if vnpto_area[a] > 0 or vn_area.get(a, 0) > 0]
        fig_c = go.Figure()
        fig_c.add_bar(name="PTO", x=areas_vn, y=[vnpto_area.get(a, 0) for a in areas_vn],
                      marker_color=COLOR_PTO)
        fig_c.add_bar(name="Real", x=areas_vn, y=[vn_area.get(a, 0) for a in areas_vn],
                      marker_color=COLOR_2026)
        fig_c.update_layout(barmode="group", height=380, yaxis_tickformat="$,.0f",
                            margin=dict(t=20, b=20),
                            legend=dict(orientation="h", y=1.12))
        st.plotly_chart(fig_c, use_container_width=True)

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

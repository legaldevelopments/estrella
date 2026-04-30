"""
Tablero Predial La Estrella 2026 — Streamlit
Análisis de límites municipales (Acuerdo 021/2025) + Críticas del sistema catastral.

Fuentes:
  Predial_Estrella_2026_Analisis.xlsx  → generado por procesar_predial_estrella.py
  Criticas 2026 Municipio de La Estrella V_1.xlsx → hojas de diferencias gestor/V6
  Destinaciones La Estrella.xlsx → catálogo de códigos

Ejecutar:  streamlit run tablero_estrella.py
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import numpy as np
import os, io

# ── CONFIG ────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Predial La Estrella 2026",
    page_icon="⭐",
    layout="wide",
    initial_sidebar_state="expanded",
)

DIR      = os.path.dirname(__file__)
EXCEL    = os.path.join(DIR, "Predial_Estrella_2026_Analisis.xlsx")
CRITICAS = os.path.join(DIR, "Criticas 2026 Municipio de La Estrella V_1.xlsx")
DEST_XLS = os.path.join(DIR, "Destinaciones La Estrella.xlsx")

AZUL_OSC  = "#1F4E79"
AZUL_MED  = "#2E75B6"
AZUL_CLAR = "#BDD7EE"
VERDE     = "#70AD47"
AMBAR     = "#FFD966"
ROJO      = "#FF7C80"
NARANJA   = "#F4B183"
GRIS      = "#D9D9D9"
CAFE      = "#833C11"

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
  .stApp { background-color: #F0F4F8; }
  [data-testid="stSidebar"] {
      background: linear-gradient(180deg, #1F4E79 0%, #2E75B6 100%);
  }
  [data-testid="stSidebar"] * { color: white !important; }
  [data-testid="stSidebar"] .stSelectbox label,
  [data-testid="stSidebar"] .stMultiSelect label,
  [data-testid="stSidebar"] .stCheckbox label { color: #BDD7EE !important; font-weight:600; }

  .kpi-card {
      background: white; border-radius: 12px; padding: 18px 14px;
      text-align: center; box-shadow: 0 2px 8px rgba(0,0,0,0.10);
      border-top: 4px solid #2E75B6; min-height: 118px;
      display: flex; flex-direction: column; justify-content: center;
  }
  .kpi-card.verde   { border-top-color: #70AD47; }
  .kpi-card.ambar   { border-top-color: #FFD966; }
  .kpi-card.rojo    { border-top-color: #FF7C80; }
  .kpi-card.oscuro  { border-top-color: #1F4E79; }
  .kpi-card.naranja { border-top-color: #F4B183; }
  .kpi-card.cafe    { border-top-color: #833C11; }
  .kpi-valor { font-size:1.6rem; font-weight:800; color:#1F4E79; line-height:1.1; }
  .kpi-label { font-size:0.70rem; color:#666; margin-top:5px; font-weight:500;
               text-transform:uppercase; letter-spacing:0.5px; }
  .kpi-sub   { font-size:0.76rem; color:#2E75B6; font-weight:600; margin-top:3px; }

  .sec-tit {
      background: linear-gradient(90deg,#1F4E79,#2E75B6);
      color:white; padding:7px 16px; border-radius:8px;
      font-size:0.92rem; font-weight:700; letter-spacing:0.4px;
      margin:20px 0 10px 0;
  }
  .sec-crit {
      background: linear-gradient(90deg,#833C11,#C55A11);
      color:white; padding:7px 16px; border-radius:8px;
      font-size:0.92rem; font-weight:700;
      margin:20px 0 10px 0;
  }
  .header-main {
      background: linear-gradient(135deg,#1F4E79 0%,#2E75B6 55%,#3A9AD9 100%);
      border-radius:14px; padding:26px 30px; color:white;
      margin-bottom:22px; box-shadow:0 4px 15px rgba(31,78,121,0.3);
  }
  .header-main h1 { margin:0; font-size:1.75rem; font-weight:800; }
  .header-main p  { margin:4px 0 0; opacity:0.85; font-size:0.9rem; }
  /* tabs */
  .stTabs [data-baseweb="tab-list"] { gap: 6px; }
  .stTabs [data-baseweb="tab"] {
      color: #1F4E79 !important; font-weight: 600; font-size: 0.88rem;
  }
  .stTabs [aria-selected="true"] {
      color: #1F4E79 !important;
      border-bottom: 3px solid #2E75B6 !important;
  }
  /* ejes y textos de plotly */
  .js-plotly-plot .gtitle, .js-plotly-plot .xtick text, .js-plotly-plot .ytick text {
      fill: #1F4E79 !important;
  }

  .nota-legal {
      background:#FFF9E6; border-left:4px solid #FFD966;
      border-radius:6px; padding:10px 14px; font-size:0.80rem; color:#555; margin-top:8px;
  }
  .badge-crit {
      display:inline-block; background:#FF7C80; color:white;
      border-radius:10px; padding:2px 8px; font-size:0.75rem; font-weight:700;
  }
  .badge-ok {
      display:inline-block; background:#70AD47; color:white;
      border-radius:10px; padding:2px 8px; font-size:0.75rem; font-weight:700;
  }
</style>
""", unsafe_allow_html=True)


# ── HELPERS ───────────────────────────────────────────────────────────────────
def fmt_cop(v):
    if pd.isna(v) or v == 0:
        return "$0"
    if abs(v) >= 1e9:
        return f"${v/1e9:,.2f} MM"
    if abs(v) >= 1e6:
        return f"${v/1e6:,.1f} M"
    return f"${v:,.0f}"


def kpi(col, val, lab, sub="", color=""):
    with col:
        st.markdown(f"""
        <div class="kpi-card {color}">
          <div class="kpi-valor">{val}</div>
          <div class="kpi-label">{lab}</div>
          <div class="kpi-sub">{sub}</div>
        </div>""", unsafe_allow_html=True)


ORDEN_RANGOS = ["0–10 M", "10–50 M", "50–100 M", "100–300 M", "300 M–1 MM", ">1 MM"]


# ── CARGA DE DATOS ────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="Cargando base Estrella 2026…")
def cargar_base():
    df = pd.read_excel(EXCEL, sheet_name="Predios_Analisis", header=1)

    num_cols = [
        "AREA_TERRENO_ANT", "AREA_TERRENO_ACT", "DIF_AREA_TERRENO",
        "AREA_CONST_ANT", "AREA_CONST_ACT", "DIF_AREA_CONST",
        "AVALUO_2025", "AVALUO_2026", "DIF_AVALUO", "PCT_AUMENTO_AVALUO",
        "TARIFA_2025", "TARIFA_2026",
        "PREDIAL_2025", "SOBRETASA_2025",
        "PREDIAL_SIN_TOPE", "SOBRETASA_SIN_TOPE",
        "PCT_TOPE", "IPU_2026", "SOBRETASA_2026",
        "EXCESO_TOPE", "VAR_IMPTO_%",
    ]
    for c in num_cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    yn = lambda col: df[col].astype(str).str.strip().str.upper() == "SÍ" if col in df.columns else pd.Series(False, index=df.index)
    df["_nuevo"]       = yn("ES_PREDIO_NUEVO")
    df["_zona_diff"]   = yn("ZONA_DIFERENTE")
    df["_sin_dest"]    = yn("SIN_DESTINACION")
    df["_dest_cambia"] = yn("DEST_CAMBIA")
    df["_dif_der"]     = yn("DIF_DERECHO")
    df["_con_exceso"]  = df["EXCESO_TOPE"].fillna(0) > 0

    for col in ["DEST_2026", "DEST_2025", "SECTOR_2026", "CATEGORIA_LIMITE", "ESTADO_FINAL_LIQ", "RANGO_AVALUO"]:
        if col in df.columns:
            df[col] = df[col].fillna("Sin dato")

    df["_rng_ord"] = pd.Categorical(df["RANGO_AVALUO"], categories=ORDEN_RANGOS + ["Sin dato"], ordered=True)

    # Deduplicar por FICHA: las columnas fiscales corresponden al predio, no al propietario
    if "FICHA" in df.columns:
        n_antes = len(df)
        df = df.drop_duplicates(subset=["FICHA"], keep="first")
        n_dup = n_antes - len(df)
        if n_dup > 0:
            df.attrs["_filas_duplicadas"] = n_dup

    return df


@st.cache_data(show_spinner="Cargando críticas del sistema…")
def cargar_criticas():
    xl = pd.ExcelFile(CRITICAS)
    result = {}
    for sn in xl.sheet_names:
        try:
            result[sn] = xl.parse(sn)
        except Exception:
            result[sn] = pd.DataFrame()
    return result


@st.cache_data(show_spinner="Cargando catálogo de destinaciones…")
def cargar_dest_catalog():
    df = pd.read_excel(DEST_XLS, sheet_name=0)
    if df.shape[1] >= 2:
        return dict(zip(df.iloc[:, 0].astype(str), df.iloc[:, 1].astype(str)))
    return {}


if not os.path.exists(EXCEL):
    st.error(
        f"No se encontró **{EXCEL}**. "
        "Ejecute primero `procesar_predial_estrella.py` para generar el archivo."
    )
    st.stop()

df_all    = cargar_base()
criticas  = cargar_criticas()
dest_cat  = cargar_dest_catalog()


# ── SIDEBAR ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⭐ Predial La Estrella 2026")
    st.markdown("**Acuerdo 021/2025 · Parágrafo Transitorio**")
    st.divider()
    st.markdown("#### Filtros globales")

    sectores_disp = ["Todos"] + sorted(df_all["SECTOR_2026"].dropna().unique().tolist())
    sel_sector = st.selectbox("Sector:", sectores_disp)

    cats_disp = sorted(df_all["CATEGORIA_LIMITE"].dropna().unique())
    sel_cat = st.multiselect("Categoría de límite:", cats_disp, placeholder="Todas")

    dests_disp = sorted(df_all["DEST_2026"].dropna().unique())
    sel_dest = st.multiselect("Destino económico 2026:", dests_disp, placeholder="Todos")

    sel_exceso = st.checkbox("Solo predios con exceso sobre tope", value=False)

    st.divider()
    umbral_av = st.slider("Umbral 'cambio extremo' avalúo (%):", 20, 500, 100, 10)

    st.divider()
    st.markdown("""
    <div style='font-size:0.72rem; opacity:0.85;'>
    📂 Fuentes:<br>
    <i>INFORME AUMENTO 2026 v3-FINAL<br>Archivo Estrella).xlsx</i><br>
    <i>Criticas 2026 Municipio de<br>La Estrella V_1.xlsx</i><br><br>
    ⚖️ Marco legal:<br>
    · Acuerdo 021/2025<br>
    · Párr. transitorio: límites a–e<br>
    · a) Lotes: máx 5× ant.<br>
    · b) E1-2 ≤135 SMMLV: 100% IPC<br>
    · c) Nuevos: impuesto pleno<br>
    · e) Demás actualizados: máx 50%
    </div>""", unsafe_allow_html=True)


# ── FILTRAR ───────────────────────────────────────────────────────────────────
df = df_all.copy()
if sel_sector != "Todos":
    df = df[df["SECTOR_2026"] == sel_sector]
if sel_cat:
    df = df[df["CATEGORIA_LIMITE"].isin(sel_cat)]
if sel_dest:
    df = df[df["DEST_2026"].isin(sel_dest)]
if sel_exceso:
    df = df[df["_con_exceso"]]


# ── HEADER ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="header-main">
  <h1>⭐ Análisis Predial — Municipio de La Estrella 2026</h1>
  <p>Actualización catastral · Acuerdo 021/2025 · Parágrafo Transitorio · Límites municipales a–e · Críticas sistema gestor vs V6</p>
</div>""", unsafe_allow_html=True)

n_f = len(df); n_t = len(df_all)
dup_removed = df_all.attrs.get("_filas_duplicadas", 0)
if dup_removed:
    st.warning(f"Se eliminaron **{dup_removed:,}** filas duplicadas por FICHA (múltiples propietarios por predio). "
               f"Los totales fiscales se calculan sobre predios únicos.")
if n_f < n_t:
    st.info(f"Mostrando **{n_f:,}** de **{n_t:,}** predios según filtros aplicados.")
else:
    st.info(f"**{n_t:,}** predios totales únicos · Municipio de La Estrella · Actualización catastral 2026")


# ── TABS PRINCIPALES ──────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 Resumen General",
    "⚖️ Análisis de Límites",
    "🔍 Críticas del Sistema",
    "💰 Avalúos",
    "📋 Detalle de Predios",
])


# ════════════════════════════════════════════════════════════════════════════════
# TAB 1 — RESUMEN GENERAL
# ════════════════════════════════════════════════════════════════════════════════
with tab1:
    predial_2025  = df["PREDIAL_2025"].fillna(0).sum()
    sin_tope_2026 = df["PREDIAL_SIN_TOPE"].fillna(0).sum()
    ipu_2026      = df["IPU_2026"].fillna(0).sum()
    exceso_total  = df["EXCESO_TOPE"].fillna(0).sum()
    n_exceso      = int(df["_con_exceso"].sum())
    var_c = (ipu_2026 - predial_2025) / predial_2025 * 100 if predial_2025 else 0
    var_s = (sin_tope_2026 - predial_2025) / predial_2025 * 100 if predial_2025 else 0

    st.markdown('<div class="sec-tit">📊 Indicadores Clave</div>', unsafe_allow_html=True)
    ks = st.columns(8)
    kpi(ks[0], f"{n_f:,}",          "Total Predios",              "Municipio La Estrella",    "oscuro")
    kpi(ks[1], fmt_cop(predial_2025), "Recaudo Predial 2025",      "Base anterior",            "")
    kpi(ks[2], fmt_cop(sin_tope_2026),"Liq. Sin Tope 2026",        "Sin aplicar límites",      "ambar")
    kpi(ks[3], fmt_cop(ipu_2026),    "IPU 2026 Correcto",          "Con límites municipales",  "verde")
    kpi(ks[4], f"{var_c:+.1f}%",    "Var. Correcto vs 2025",      "Recaudo con topes",        "verde" if var_c >= 0 else "rojo")
    kpi(ks[5], f"{var_s:+.1f}%",    "Var. Sin Tope vs 2025",      "Sin aplicar límites",      "rojo" if var_s > 50 else "ambar")
    kpi(ks[6], f"{n_exceso:,}",     "Predios con Exceso",         "Liq. > tope municipal",    "rojo")
    kpi(ks[7], fmt_cop(exceso_total),"Exceso Total Sobre Tope",   "Ahorro contribuyentes",    "naranja")

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="sec-tit">📈 Comparativo Fiscal</div>', unsafe_allow_html=True)

    c1, c2, c3 = st.columns([2.2, 1.8, 1.8])

    with c1:
        fig_bar = go.Figure(go.Bar(
            x=["Recaudo\nAnterior 2025", "Liq. Sin Tope\n2026", "IPU Correcto\n2026"],
            y=[predial_2025, sin_tope_2026, ipu_2026],
            marker_color=[AZUL_MED, ROJO if sin_tope_2026 > ipu_2026 else AMBAR, VERDE],
            text=[fmt_cop(v) for v in [predial_2025, sin_tope_2026, ipu_2026]],
            textposition="outside", textfont=dict(size=10, color="#1F4E79"),
        ))
        fig_bar.update_layout(
            title=dict(text="Comparativo de Recaudo ($)", font=dict(size=13, color=AZUL_OSC)),
            xaxis=dict(tickfont=dict(color="#1F4E79")),
            yaxis=dict(tickformat="$,.0f", showgrid=True, gridcolor="#eee", tickfont=dict(color="#1F4E79")),
            plot_bgcolor="white", paper_bgcolor="white",
            margin=dict(t=50, b=20, l=10, r=20), height=330, showlegend=False,
        )
        st.plotly_chart(fig_bar, use_container_width=True)

    with c2:
        cat_counts = df["CATEGORIA_LIMITE"].value_counts()
        fig_pie = go.Figure(go.Pie(
            labels=cat_counts.index.tolist(),
            values=cat_counts.values.tolist(),
            hole=0.50,
            marker_colors=[VERDE, AZUL_MED, AMBAR, NARANJA, ROJO, GRIS][:len(cat_counts)],
            textinfo="percent", textfont=dict(size=8, color="#1F4E79"),
            hovertemplate="%{label}: %{value:,}<extra></extra>",
        ))
        fig_pie.update_layout(
            title=dict(text="Distribución Categorías de Límite", font=dict(size=12, color=AZUL_OSC)),
            showlegend=True,
            legend=dict(font=dict(size=8, color="#1F4E79"), orientation="v"),
            margin=dict(t=50, b=0, l=0, r=0), height=330,
            paper_bgcolor="white",
            annotations=[dict(text=f"<b>{n_f:,}</b><br>predios",
                              x=0.5, y=0.5, font_size=11, showarrow=False, font_color=AZUL_OSC)],
        )
        st.plotly_chart(fig_pie, use_container_width=True)

    with c3:
        fig_var = go.Figure(go.Bar(
            x=["Var. Sin\nTope %", "Var. Con\nLímites %"],
            y=[var_s, var_c],
            marker_color=[ROJO if var_s > 50 else AMBAR, VERDE if var_c >= 0 else ROJO],
            text=[f"{var_s:+.1f}%", f"{var_c:+.1f}%"],
            textposition="outside", textfont=dict(size=12, color="#1F4E79"),
        ))
        fig_var.add_hline(y=50, line_dash="dot", line_color=NARANJA,
                          annotation_text="Límite e = 50%", annotation_position="bottom right")
        fig_var.update_layout(
            title=dict(text="Variación Recaudo vs 2025", font=dict(size=13, color=AZUL_OSC)),
            xaxis=dict(tickfont=dict(color="#1F4E79")),
            yaxis=dict(ticksuffix="%", showgrid=True, gridcolor="#eee", tickfont=dict(color="#1F4E79")),
            plot_bgcolor="white", paper_bgcolor="white",
            margin=dict(t=50, b=20, l=10, r=10), height=330, showlegend=False,
        )
        st.plotly_chart(fig_var, use_container_width=True)

    # Marcadores de criticidad
    st.markdown('<div class="sec-tit">🚩 Marcadores de Criticidad en Base</div>', unsafe_allow_html=True)
    mk1, mk2, mk3, mk4, mk5 = st.columns(5)
    kpi(mk1, f"{int(df['_nuevo'].sum()):,}",       "Predios Nuevos\n(gestor, no en V6)",    "",  "ambar")
    kpi(mk2, f"{int(df['_zona_diff'].sum()):,}",   "Zonas ≠ 100%",                          "",  "naranja")
    kpi(mk3, f"{int(df['_sin_dest'].sum()):,}",    "Sin Destinación\nEconómica",            "",  "rojo")
    kpi(mk4, f"{int(df['_dest_cambia'].sum()):,}", "Destinación\nCambia Gestor→V6",         "",  "cafe")
    kpi(mk5, f"{int(df['_dif_der'].sum()):,}",     "Diferencia\nDerecho Gestor/V6",         "",  "rojo")


# ════════════════════════════════════════════════════════════════════════════════
# TAB 2 — ANÁLISIS DE LÍMITES
# ════════════════════════════════════════════════════════════════════════════════
with tab2:
    st.markdown('<div class="sec-tit">⚖️ Análisis por Categoría de Límite Municipal</div>',
                unsafe_allow_html=True)

    df_cat = (
        df.groupby("CATEGORIA_LIMITE")
        .agg(
            N_Predios=("FICHA", "count"),
            Con_Exceso=("_con_exceso", "sum"),
            Predial_2025=("PREDIAL_2025", "sum"),
            Liq_Sin_Tope=("PREDIAL_SIN_TOPE", "sum"),
            IPU_2026=("IPU_2026", "sum"),
            Exceso_Tope=("EXCESO_TOPE", "sum"),
        )
        .reset_index()
    )
    df_cat["Var_%"] = (
        (df_cat["IPU_2026"] - df_cat["Predial_2025"])
        / df_cat["Predial_2025"].replace(0, np.nan) * 100
    ).round(2)
    df_cat["Ahorro_Contrib"] = df_cat["Exceso_Tope"]
    df_cat = df_cat.sort_values("IPU_2026", ascending=False)

    ca1, ca2 = st.columns(2)
    with ca1:
        fig_cat = go.Figure(go.Bar(
            x=df_cat["N_Predios"], y=df_cat["CATEGORIA_LIMITE"],
            orientation="h", marker_color=AZUL_MED,
            text=df_cat["N_Predios"].apply(lambda v: f"{int(v):,}"),
            textposition="outside",
        ))
        fig_cat.update_layout(
            title="N° Predios por Categoría",
            plot_bgcolor="white", paper_bgcolor="white",
            xaxis=dict(tickfont=dict(color="#1F4E79")),
            yaxis=dict(tickfont=dict(color="#1F4E79"), autorange="reversed"),
            margin=dict(t=50, b=20, l=10, r=60), height=340,
        )
        st.plotly_chart(fig_cat, use_container_width=True)

    with ca2:
        fig_ipu = go.Figure()
        fig_ipu.add_trace(go.Bar(
            name="Predial 2025",
            x=df_cat["CATEGORIA_LIMITE"], y=df_cat["Predial_2025"],
            marker_color=AZUL_CLAR,
        ))
        fig_ipu.add_trace(go.Bar(
            name="IPU 2026 Correcto",
            x=df_cat["CATEGORIA_LIMITE"], y=df_cat["IPU_2026"],
            marker_color=VERDE,
        ))
        fig_ipu.add_trace(go.Bar(
            name="Exceso sobre Tope",
            x=df_cat["CATEGORIA_LIMITE"], y=df_cat["Exceso_Tope"],
            marker_color=ROJO,
        ))
        fig_ipu.update_layout(
            title="Recaudo por Categoría ($)",
            barmode="group", plot_bgcolor="white", paper_bgcolor="white",
            xaxis=dict(tickfont=dict(color="#1F4E79"), tickangle=-15),
            yaxis=dict(tickformat="$,.0f", showgrid=True, gridcolor="#eee", tickfont=dict(color="#1F4E79")),
            legend=dict(orientation="h", yanchor="bottom", y=1.01),
            margin=dict(t=60, b=30, l=10, r=10), height=340,
        )
        st.plotly_chart(fig_ipu, use_container_width=True)

    # Tabla resumen por categoría
    col_cfg_cat = {
        "Predial_2025":  st.column_config.NumberColumn("Predial 2025 ($)", format="$ %,.0f"),
        "Liq_Sin_Tope":  st.column_config.NumberColumn("Liq. Sin Tope ($)", format="$ %,.0f"),
        "IPU_2026":      st.column_config.NumberColumn("IPU 2026 ($)", format="$ %,.0f"),
        "Exceso_Tope":   st.column_config.NumberColumn("Exceso Tope ($)", format="$ %,.0f"),
        "Var_%":         st.column_config.NumberColumn("Var. %", format="%.2f %%"),
        "Con_Exceso":    st.column_config.NumberColumn("Con Exceso", format="%d"),
    }
    st.dataframe(df_cat.reset_index(drop=True), use_container_width=True,
                 height=300, column_config=col_cfg_cat)

    # Por destino y rango
    st.markdown('<div class="sec-tit">📊 Distribución por Destino y Rango de Avalúo</div>',
                unsafe_allow_html=True)

    dt1, dt2, dt3 = st.tabs(["N° Predios", "IPU 2026 ($)", "Exceso Tope ($)"])
    df_mat = df[df["RANGO_AVALUO"].isin(ORDEN_RANGOS)].copy()

    def pivote(val_col, aggfn="sum"):
        if df_mat.empty or val_col not in df_mat.columns:
            return pd.DataFrame()
        piv = pd.pivot_table(df_mat, values=val_col, index="DEST_2026",
                             columns="RANGO_AVALUO", aggfunc=aggfn, fill_value=0, observed=True)
        cols_o = [c for c in ORDEN_RANGOS if c in piv.columns]
        piv = piv[cols_o]
        piv["TOTAL"] = piv.sum(axis=1)
        piv = piv.sort_values("TOTAL", ascending=False)
        fila = piv.sum(); fila.name = "▶ TOTAL"
        return pd.concat([piv, fila.to_frame().T])

    with dt1:
        p1 = pivote("FICHA", "count")
        if not p1.empty:
            st.dataframe(p1.astype(int).style.background_gradient(cmap="Blues", axis=None, subset=ORDEN_RANGOS),
                         use_container_width=True, height=380)
    with dt2:
        p2 = pivote("IPU_2026")
        if not p2.empty:
            st.dataframe(p2.style.format("${:,.0f}").background_gradient(cmap="Greens", axis=None, subset=ORDEN_RANGOS),
                         use_container_width=True, height=380)
    with dt3:
        p3 = pivote("EXCESO_TOPE")
        if not p3.empty:
            st.dataframe(p3.style.format("${:,.0f}").background_gradient(cmap="Reds", axis=None, subset=ORDEN_RANGOS),
                         use_container_width=True, height=380)

    # Cambios extremos avalúo
    st.markdown(f'<div class="sec-tit">🚨 Predios con Variación Avalúo ≥ {umbral_av}%</div>',
                unsafe_allow_html=True)

    col_var = "PCT_AUMENTO_AVALUO"
    if col_var in df.columns:
        df_ext = df[df[col_var].fillna(0).abs() >= umbral_av].copy().sort_values(col_var, ascending=False)
    else:
        df_ext = pd.DataFrame()

    st.markdown(f"**{len(df_ext):,}** predios con variación de avalúo ≥ {umbral_av}% "
                f"· de estos **{int(df_ext['_con_exceso'].sum() if not df_ext.empty else 0):,}** con exceso sobre tope.")

    if not df_ext.empty:
        cols_ext = [c for c in [
            "FICHA", "DEST_2025", "DEST_2026", "SECTOR_2026",
            "AVALUO_2025", "AVALUO_2026", "DIF_AVALUO", "PCT_AUMENTO_AVALUO",
            "PREDIAL_2025", "PREDIAL_SIN_TOPE", "IPU_2026",
            "EXCESO_TOPE", "CATEGORIA_LIMITE",
        ] if c in df_ext.columns]

        cc_ext = {
            "AVALUO_2025": st.column_config.NumberColumn(format="$ %,.0f"),
            "AVALUO_2026": st.column_config.NumberColumn(format="$ %,.0f"),
            "DIF_AVALUO":  st.column_config.NumberColumn(format="$ %,.0f"),
            "PCT_AUMENTO_AVALUO": st.column_config.NumberColumn(format="%.2f %%"),
            "PREDIAL_2025": st.column_config.NumberColumn(format="$ %,.0f"),
            "PREDIAL_SIN_TOPE": st.column_config.NumberColumn(format="$ %,.0f"),
            "IPU_2026": st.column_config.NumberColumn(format="$ %,.0f"),
            "EXCESO_TOPE": st.column_config.NumberColumn(format="$ %,.0f"),
        }
        st.dataframe(df_ext[cols_ext].reset_index(drop=True),
                     use_container_width=True, height=360, column_config=cc_ext)

        buf_ext = io.BytesIO()
        with pd.ExcelWriter(buf_ext, engine="openpyxl") as wr:
            df_ext[cols_ext].to_excel(wr, index=False, sheet_name="Cambios_Extremos")
        st.download_button("⬇️ Descargar predios cambio extremo (.xlsx)", data=buf_ext.getvalue(),
                           file_name="estrella_cambios_extremos.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ════════════════════════════════════════════════════════════════════════════════
# TAB 3 — CRÍTICAS DEL SISTEMA
# ════════════════════════════════════════════════════════════════════════════════
with tab3:
    st.markdown('<div class="sec-crit">🔍 Críticas del Sistema — Diferencias Gestor vs V6</div>',
                unsafe_allow_html=True)

    DESCRIPCIONES = {
        "DEST EN GESTOR Y NO EN V6":       ("⚠️ Códigos de destinación en gestor no existentes en V6", NARANJA),
        "PREDIOS EN GESTOR Y NO EN V6":    ("🆕 Predios nuevos en el gestor (no estaban en V6)", AMBAR),
        "PREDIOS EN V6 Y NO EN GESTOR":    ("🗑️ Predios eliminados: estaban en V6 y ya no están en gestor", GRIS),
        "PRED-PROP EN V6 Y NO EN GESTOR":  ("👤 Predios-propietarios en V6 sin correspondencia en gestor", GRIS),
        "PRED-PROP V6 Y NO GESTOR DEUDA":  ("💸 Predios-propietarios en V6 sin gestor CON DEUDA", ROJO),
        "PRED-PROP GESTOR Y V6 <> DER":    ("⚖️ Predios con diferencia en % de derecho entre sistemas", ROJO),
        "PREDIOS EN GESTOR Y EN V6":       ("✅ Predios que cruzan correctamente en ambas bases", VERDE),
        "TERCEROS EN GESTOR Y NO EN V6":   ("👥 Terceros nuevos en gestor no registrados en V6", AMBAR),
        "PROP EN GESTOR Y NO EN V6":       ("👤 Propietarios en gestor sin registro en V6", AMBAR),
        "PROP EN V6 Y NO EN GESTOR":       ("👤 Propietarios en V6 sin correspondencia en gestor", GRIS),
        "PRED-PROP IGUAL EN GESTOR Y V6":  ("✅ Predios-propietarios iguales en ambas bases", VERDE),
        "PROP EN V6 Y NO GESTOR (DEUDA)":  ("💸 Propietarios en V6 sin gestor CON DEUDA", ROJO),
        "DEST QUE CAMBIAN":                ("🔄 Predios con cambio de destinación entre sistemas", NARANJA),
        "AVALUOS ENTRE GESTOR Y V6":       ("💰 Comparativo de avalúos entre gestor y V6", AZUL_MED),
    }

    TOTALES_OK = {"PREDIOS EN GESTOR Y EN V6", "PRED-PROP IGUAL EN GESTOR Y V6"}

    # Resumen tarjetas
    st.markdown("#### Resumen de diferencias")
    cols_r = st.columns(4)
    items = [(k, v) for k, v in criticas.items() if k not in TOTALES_OK]
    for i, (sname, df_c) in enumerate(items):
        desc, color = DESCRIPCIONES.get(sname, (sname, GRIS))
        n = len(df_c) if isinstance(df_c, pd.DataFrame) else 0
        badge = f'<span class="badge-crit">{n:,}</span>' if n > 0 else f'<span class="badge-ok">0</span>'
        with cols_r[i % 4]:
            st.markdown(f"**{desc}**<br>{badge}", unsafe_allow_html=True)
            st.markdown("---")

    # Tabs por categoría
    crit_keys_order = [
        "PREDIOS EN GESTOR Y NO EN V6",
        "PREDIOS EN V6 Y NO EN GESTOR",
        "PRED-PROP GESTOR Y V6 <> DER",
        "PRED-PROP V6 Y NO GESTOR DEUDA",
        "PROP EN V6 Y NO GESTOR (DEUDA)",
        "DEST QUE CAMBIAN",
        "TERCEROS EN GESTOR Y NO EN V6",
        "PROP EN GESTOR Y NO EN V6",
        "PROP EN V6 Y NO EN GESTOR",
        "PRED-PROP EN V6 Y NO EN GESTOR",
        "PREDIOS EN GESTOR Y EN V6",
        "PRED-PROP IGUAL EN GESTOR Y V6",
        "DEST EN GESTOR Y NO EN V6",
    ]

    tab_labels = []
    for k in crit_keys_order:
        if k in criticas:
            n = len(criticas[k]) if isinstance(criticas[k], pd.DataFrame) else 0
            short = k[:22] + "…" if len(k) > 22 else k
            tab_labels.append(f"{short} ({n:,})")

    sub_tabs = st.tabs(tab_labels)
    for i, (k, sub_t) in enumerate(zip(crit_keys_order, sub_tabs)):
        if k not in criticas:
            continue
        df_c = criticas[k] if isinstance(criticas[k], pd.DataFrame) else pd.DataFrame()
        desc, color = DESCRIPCIONES.get(k, (k, GRIS))
        with sub_t:
            st.markdown(f"**{desc}** — {len(df_c):,} registros")

            # Enriquecer destinos si tiene columna de código
            if "Destinación Gestor" in df_c.columns:
                df_c = df_c.copy()
                df_c["Nombre Dest. Gestor"] = df_c["Destinación Gestor"].astype(str).map(
                    lambda x: dest_cat.get(x, x))
            if "Destinación V6" in df_c.columns:
                df_c = df_c.copy()
                df_c["Nombre Dest. V6"] = df_c["Destinación V6"].astype(str).map(
                    lambda x: dest_cat.get(x, x))

            if not df_c.empty:
                # Formato dinero para columnas de deuda/valor
                cc = {}
                for col in df_c.columns:
                    if "deuda" in col.lower() or "valor" in col.lower() or "avaluo" in col.lower():
                        cc[col] = st.column_config.NumberColumn(col, format="$ %,.0f")
                    elif "derecho" in col.lower() or "%" in col.lower():
                        cc[col] = st.column_config.NumberColumn(col, format="%.4f")

                busq_c = st.text_input(f"🔎 Filtrar en {k[:30]}…", key=f"busq_{i}",
                                       placeholder="Ficha, documento, nombre…")
                df_show = df_c.copy()
                if busq_c:
                    mask = df_show.apply(
                        lambda col: col.astype(str).str.contains(busq_c, case=False, na=False)
                    ).any(axis=1)
                    df_show = df_show[mask]

                st.markdown(f"**{len(df_show):,}** registros")
                st.dataframe(df_show.reset_index(drop=True), use_container_width=True,
                             height=380, column_config=cc)

                buf_c = io.BytesIO()
                with pd.ExcelWriter(buf_c, engine="openpyxl") as wr:
                    df_show.to_excel(wr, index=False, sheet_name="Critica")
                st.download_button(
                    f"⬇️ Descargar {k[:30]}… (.xlsx)",
                    data=buf_c.getvalue(),
                    file_name=f"critica_{i+1}_{k[:30].replace(' ','_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_crit_{i}",
                )
            else:
                st.success(f"Sin registros en esta crítica.")


# ════════════════════════════════════════════════════════════════════════════════
# TAB 4 — AVALÚOS
# ════════════════════════════════════════════════════════════════════════════════
with tab4:
    st.markdown('<div class="sec-tit">💰 Comparativo de Avalúos — Gestor vs V6</div>',
                unsafe_allow_html=True)

    df_av = criticas.get("AVALUOS ENTRE GESTOR Y V6", pd.DataFrame())
    if isinstance(df_av, pd.DataFrame) and not df_av.empty:
        df_av = df_av.copy()
        # Normalizar nombres de columnas
        col_map = {c: c.strip() for c in df_av.columns}
        df_av = df_av.rename(columns=col_map)

        col_gest = next((c for c in df_av.columns if "gestor" in c.lower()), None)
        col_v6   = next((c for c in df_av.columns if "v6" in c.lower() or "v 6" in c.lower()), None)
        col_pct  = next((c for c in df_av.columns if "%" in c or "incremento" in c.lower()), None)
        col_dif  = next((c for c in df_av.columns if "diferencia" in c.lower()), None)

        for col in [col_gest, col_v6, col_pct, col_dif]:
            if col and col in df_av.columns:
                df_av[col] = pd.to_numeric(df_av[col], errors="coerce")

        n_av = len(df_av)
        av1, av2, av3, av4 = st.columns(4)
        kpi(av1, f"{n_av:,}",               "Predios Comparados",      "Gestor vs V6",          "oscuro")
        if col_gest:
            kpi(av2, fmt_cop(df_av[col_gest].sum()), "Total Avalúo Gestor",   "Nuevo (actualizado)",  "verde")
        if col_v6:
            kpi(av3, fmt_cop(df_av[col_v6].sum()),   "Total Avalúo V6",       "Anterior (sistema)",   "ambar")
        if col_pct and col_pct in df_av.columns:
            med_pct = df_av[col_pct].median()
            kpi(av4, f"{med_pct:,.1f}%",             "% Incremento Mediano",  "Avalúo gestor vs V6",  "naranja")

        st.markdown("<br>", unsafe_allow_html=True)
        a1, a2 = st.columns(2)

        with a1:
            if col_pct and col_pct in df_av.columns:
                df_av_filt = df_av[df_av[col_pct].abs() < 2000]
                fig_hist_av = px.histogram(
                    df_av_filt, x=col_pct, nbins=60,
                    color_discrete_sequence=[AZUL_MED],
                    labels={col_pct: "% Incremento Avalúo (Gestor vs V6)"},
                    title="Distribución % Incremento Avalúo",
                )
                fig_hist_av.add_vline(x=0,   line_dash="dash", line_color=AZUL_OSC)
                fig_hist_av.add_vline(x=100, line_dash="dot",  line_color=ROJO,
                                      annotation_text="100% (Ley 44 referencia)")
                fig_hist_av.add_vline(x=50,  line_dash="dot",  line_color=NARANJA,
                                      annotation_text="50% (Límite e)")
                fig_hist_av.update_layout(
                    plot_bgcolor="white", paper_bgcolor="white",
                    height=360, margin=dict(t=50, b=20, l=10, r=10),
                    xaxis=dict(tickfont=dict(color="#1F4E79")),
                    yaxis=dict(title="N° predios", tickfont=dict(color="#1F4E79")),
                )
                st.plotly_chart(fig_hist_av, use_container_width=True)

        with a2:
            if col_gest and col_v6 and col_gest in df_av.columns and col_v6 in df_av.columns:
                df_sc = df_av[df_av[col_gest].notna() & df_av[col_v6].notna()].copy()
                muestra = df_sc.sample(min(len(df_sc), 2000), random_state=42)
                if col_pct and col_pct in muestra.columns:
                    muestra["_excede_ley44"] = muestra[col_pct] > 100
                    muestra["_cat"] = muestra["_excede_ley44"].map(
                        {True: "Incremento > 100%", False: "Incremento ≤ 100%"})
                else:
                    muestra["_cat"] = "Sin dato"

                fig_sc = px.scatter(
                    muestra, x=col_v6, y=col_gest,
                    color="_cat",
                    color_discrete_map={"Incremento > 100%": ROJO, "Incremento ≤ 100%": VERDE, "Sin dato": GRIS},
                    title="Avalúo V6 vs Avalúo Gestor (muestra 2.000)",
                    opacity=0.5,
                )
                # Línea de igualdad
                max_val = max(df_sc[col_gest].quantile(0.99), df_sc[col_v6].quantile(0.99))
                fig_sc.add_shape(type="line", x0=0, y0=0, x1=max_val, y1=max_val,
                                 line=dict(color=AZUL_OSC, dash="dash"))
                fig_sc.add_shape(type="line", x0=0, y0=0, x1=max_val, y1=max_val * 2,
                                 line=dict(color=ROJO, dash="dot"))
                fig_sc.update_layout(
                    plot_bgcolor="white", paper_bgcolor="white",
                    height=360, margin=dict(t=50, b=20, l=10, r=10),
                    legend=dict(orientation="h", yanchor="bottom", y=1.01),
                    xaxis=dict(tickformat="$,.0f", tickfont=dict(color="#1F4E79")),
                    yaxis=dict(tickformat="$,.0f", tickfont=dict(color="#1F4E79")),
                )
                fig_sc.update_traces(marker=dict(size=4))
                st.plotly_chart(fig_sc, use_container_width=True)

        # Análisis Ley 44
        if col_pct and col_pct in df_av.columns:
            st.markdown('<div class="sec-tit">⚖️ Referencia Ley 44 — Incrementos en Avalúo</div>',
                        unsafe_allow_html=True)
            st.markdown("""
            La Ley 44/1990 Art. 6 establece que el impuesto no puede superar el **doble del año anterior**
            (200% del impuesto anterior). Aquí se analiza el incremento en avalúo como referencia —
            un incremento > 100% en avalúo sugiere riesgo de sobrepasar el doble del impuesto anterior
            si la tarifa se mantiene constante. Los límites locales del Acuerdo 021/2025 son más estrictos.
            """)

            df_av["_cat_pct"] = pd.cut(
                df_av[col_pct].fillna(0),
                bins=[-float("inf"), 0, 50, 100, 200, float("inf")],
                labels=["Baja/negativa (≤0%)", "Baja (0–50%)", "Moderada (50–100%)",
                        "Alta (100–200%)", "Muy alta (>200%)"],
            )
            cat_av = df_av["_cat_pct"].value_counts().reset_index()
            cat_av.columns = ["Rango de Incremento", "N° Predios"]
            cat_av = cat_av.sort_values("Rango de Incremento")

            c_av1, c_av2 = st.columns([1.5, 2])
            with c_av1:
                st.dataframe(cat_av.reset_index(drop=True), use_container_width=True, height=220)
            with c_av2:
                fig_cat_av = go.Figure(go.Bar(
                    x=cat_av["Rango de Incremento"], y=cat_av["N° Predios"],
                    marker_color=[VERDE, AZUL_CLAR, AMBAR, NARANJA, ROJO][:len(cat_av)],
                    text=cat_av["N° Predios"].apply(lambda v: f"{int(v):,}"),
                    textposition="outside",
                ))
                fig_cat_av.update_layout(
                    title="N° Predios por Rango de Incremento en Avalúo",
                    plot_bgcolor="white", paper_bgcolor="white",
                    height=260, margin=dict(t=50, b=20, l=10, r=10),
                    xaxis=dict(tickfont=dict(color="#1F4E79"), tickangle=-20),
                    yaxis=dict(tickfont=dict(color="#1F4E79")),
                )
                st.plotly_chart(fig_cat_av, use_container_width=True)

        # Tabla detalle avalúos
        st.markdown("**Detalle comparativo avalúos (filtrable)**")
        busq_av = st.text_input("🔎 Filtrar por ficha o dirección:", key="busq_av", placeholder="Escriba para filtrar…")
        df_av_show = df_av.copy()
        if busq_av:
            mask = df_av_show.apply(lambda col: col.astype(str).str.contains(busq_av, case=False, na=False)).any(axis=1)
            df_av_show = df_av_show[mask]

        cc_av = {c: st.column_config.NumberColumn(c, format="$ %,.0f") for c in [col_gest, col_v6, col_dif] if c}
        if col_pct:
            cc_av[col_pct] = st.column_config.NumberColumn(col_pct, format="%.2f %%")

        cols_av_vis = [c for c in df_av.columns if c != "_cat_pct" and not c.startswith("_")]
        st.markdown(f"**{len(df_av_show):,}** registros")
        st.dataframe(df_av_show[cols_av_vis].reset_index(drop=True),
                     use_container_width=True, height=360, column_config=cc_av)

        buf_av = io.BytesIO()
        with pd.ExcelWriter(buf_av, engine="openpyxl") as wr:
            df_av_show[cols_av_vis].to_excel(wr, index=False, sheet_name="Avaluos_Comparativo")
        st.download_button("⬇️ Descargar comparativo avalúos (.xlsx)", data=buf_av.getvalue(),
                           file_name="estrella_avaluos_comparativo.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    else:
        st.info("No se encontró la hoja AVALUOS ENTRE GESTOR Y V6 en el archivo de críticas.")


# ════════════════════════════════════════════════════════════════════════════════
# TAB 5 — DETALLE DE PREDIOS
# ════════════════════════════════════════════════════════════════════════════════
with tab5:
    st.markdown('<div class="sec-tit">📋 Resumen por Destino Económico</div>', unsafe_allow_html=True)

    df_res_dest = (
        df.groupby("DEST_2026")
        .agg(
            N_Predios=("FICHA", "count"),
            Predial_2025=("PREDIAL_2025", "sum"),
            Liq_Sin_Tope=("PREDIAL_SIN_TOPE", "sum"),
            IPU_2026=("IPU_2026", "sum"),
            Exceso_Tope=("EXCESO_TOPE", "sum"),
            Con_Exceso=("_con_exceso", "sum"),
            Pred_Nuevos=("_nuevo", "sum"),
            Sin_Dest=("_sin_dest", "sum"),
        )
        .reset_index()
    )
    df_res_dest["Var_IPU_%"] = (
        (df_res_dest["IPU_2026"] - df_res_dest["Predial_2025"])
        / df_res_dest["Predial_2025"].replace(0, np.nan) * 100
    ).round(2)
    df_res_dest = df_res_dest.sort_values("IPU_2026", ascending=False)

    cc_res = {
        "Predial_2025":  st.column_config.NumberColumn("Predial 2025 ($)", format="$ %,.0f"),
        "Liq_Sin_Tope":  st.column_config.NumberColumn("Liq. Sin Tope ($)", format="$ %,.0f"),
        "IPU_2026":      st.column_config.NumberColumn("IPU 2026 ($)", format="$ %,.0f"),
        "Exceso_Tope":   st.column_config.NumberColumn("Exceso Tope ($)", format="$ %,.0f"),
        "Var_IPU_%":     st.column_config.NumberColumn("Var. %", format="%.2f %%"),
    }
    st.dataframe(df_res_dest.reset_index(drop=True), use_container_width=True,
                 height=340, column_config=cc_res)

    csv_dest = df_res_dest.to_csv(index=False).encode("utf-8-sig")
    st.download_button("⬇️ Descargar resumen por destino (.csv)", data=csv_dest,
                       file_name="estrella_resumen_destino.csv", mime="text/csv")

    st.markdown('<div class="sec-tit">📋 Detalle Individual de Predios</div>', unsafe_allow_html=True)

    cols_vis = [c for c in [
        "FICHA", "SECTOR_2026", "DEST_2025", "DEST_2026",
        "ESTRATO_2025", "ESTRATO_2026",
        "AVALUO_2025", "AVALUO_2026", "PCT_AUMENTO_AVALUO",
        "TARIFA_2025", "TARIFA_2026",
        "PREDIAL_2025", "PREDIAL_SIN_TOPE",
        "ESTADO_FINAL_LIQ", "CATEGORIA_LIMITE",
        "IPU_2026", "SOBRETASA_2026",
        "EXCESO_TOPE", "VAR_IMPTO_%",
        "RANGO_AVALUO",
        "ES_PREDIO_NUEVO", "ZONA_DIFERENTE", "SIN_DESTINACION",
        "DEST_CAMBIA", "DIF_DERECHO", "NOVEDADES",
    ] if c in df.columns]

    busq = st.text_input("🔎 Buscar por ficha, destino o novedad:",
                         placeholder="Escriba para filtrar…", key="busq_pred")
    df_vis = df[cols_vis].copy()
    if busq:
        mask = df_vis.apply(lambda col: col.astype(str).str.contains(busq, case=False, na=False)).any(axis=1)
        df_vis = df_vis[mask]

    st.markdown(f"**{len(df_vis):,}** registros")

    cc_vis = {}
    for c in ["AVALUO_2025", "AVALUO_2026", "PREDIAL_2025", "PREDIAL_SIN_TOPE", "IPU_2026",
              "SOBRETASA_2026", "EXCESO_TOPE"]:
        if c in cols_vis:
            cc_vis[c] = st.column_config.NumberColumn(c, format="$ %,.0f")
    for c in ["PCT_AUMENTO_AVALUO", "VAR_IMPTO_%"]:
        if c in cols_vis:
            cc_vis[c] = st.column_config.NumberColumn(c, format="%.2f %%")
    for c in ["TARIFA_2025", "TARIFA_2026"]:
        if c in cols_vis:
            cc_vis[c] = st.column_config.NumberColumn(c, format="%.2f ‰")

    st.dataframe(df_vis.reset_index(drop=True), use_container_width=True,
                 height=460, column_config=cc_vis)

    csv_all = df_vis.to_csv(index=False).encode("utf-8-sig")
    st.download_button("⬇️ Descargar tabla filtrada (.csv)", data=csv_all,
                       file_name="estrella_predios_2026.csv", mime="text/csv")

    buf_all = io.BytesIO()
    with pd.ExcelWriter(buf_all, engine="openpyxl") as wr:
        df_vis.to_excel(wr, index=False, sheet_name="Predios")
    st.download_button("⬇️ Descargar tabla filtrada (.xlsx)", data=buf_all.getvalue(),
                       file_name="estrella_predios_2026.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ── NOTA LEGAL ────────────────────────────────────────────────────────────────
st.markdown("""
<div class="nota-legal">
⚖️ <strong>Nota legal:</strong>
Límites Parágrafo Transitorio 2026 — Acuerdo 021/2025:
<b>a)</b> Lotes urbanizables no urbanizados/urbanizados no edificados: máx 5× impuesto anterior |
<b>b)</b> Viviendas E1-2 ≤135 SMMLV sin modif. área significativa: máx 100% IPC |
<b>c)</b> Predios nuevos: impuesto pleno (tarifa × avalúo) |
<b>d)</b> Autoestimación: Ley 44 — <em>NO APLICA (sin avalúos por autoestimación)</em> |
<b>e)</b> Demás inmuebles actualizados: máx 50% del impuesto liquidado año anterior.
EXCESO_TOPE = PREDIAL_SIN_TOPE − IPU_2026. Verificar con Estatuto Tributario Municipal vigente.
</div>""", unsafe_allow_html=True)

st.markdown(
    "<br><center style='color:#aaa; font-size:0.76rem;'>"
    "Municipio de La Estrella · Predial 2026 · Acuerdo 021/2025 · Parágrafo Transitorio"
    "</center>", unsafe_allow_html=True,
)




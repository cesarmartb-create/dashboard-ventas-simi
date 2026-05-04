import streamlit as st
import pandas as pd
import plotly.express as px
from pagina_compras import render_compras

st.set_page_config(
    page_title="Dashboard de Ventas - Farmacias",
    page_icon="💊",
    layout="wide"
)

st.markdown("""
<style>
    .stApp { background-color: #f4f6fb; }
    section[data-testid="stSidebar"] { background-color: #1a2340; }
    section[data-testid="stSidebar"] * { color: white !important; }

    /* Texto negro en cajas de archivos subidos */
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] span,
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] p,
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] small,
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] div,
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] li,
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] *,
    section[data-testid="stSidebar"] input[type="number"] {
        color: #1a2340 !important;
        font-weight: 600 !important;
    }

    /* Fondo blanco en inputs de metas */
    section[data-testid="stSidebar"] input[type="number"] {
        background-color: white !important;
        border-radius: 6px !important;
        padding: 6px 10px !important;
    }

    /* Fondo de las cajitas de archivos */
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] [class*="uploadedFileName"],
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] [class*="fileSize"] {
        color: #1a2340 !important;
    }
    section[data-testid="stSidebar"] .stSelectbox label,
    section[data-testid="stSidebar"] .stMultiSelect label,
    section[data-testid="stSidebar"] .stFileUploader label {
        color: #a0aec0 !important;
        font-size: 11px !important;
        text-transform: uppercase;
        letter-spacing: 1px;
    }
    [data-testid="stMetricValue"] {
        font-size: 1.8rem !important;
        font-weight: 700 !important;
        color: #1a2340 !important;
    }
    [data-testid="metric-container"] {
        background: white;
        border-radius: 12px;
        padding: 20px 24px;
        box-shadow: 0 1px 4px rgba(0,0,0,0.08);
        border-top: 3px solid #3b82f6;
    }
    h1 { color: #1a2340 !important; font-size: 1.6rem !important; }
    h2, h3 { color: #1a2340 !important; }
    .block-container { padding-top: 1.5rem; }
    div[data-testid="stDataFrame"] {
        background: white;
        border-radius: 12px;
        padding: 8px;
        box-shadow: 0 1px 4px rgba(0,0,0,0.08);
    }
</style>
""", unsafe_allow_html=True)

AZUL       = "#3b82f6"
AZUL_LISTA = ["#3b82f6","#60a5fa","#93c5fd","#1d4ed8","#2563eb","#6366f1","#8b5cf6","#06b6d4"]
VERDE      = "#10b981"
NARANJA    = "#f59e0b"
DORADO     = "#f59e0b"

# ── FUNCIONES ────────────────────────────────────────────

@st.cache_data
def cargar_personal(contenido):
    df = pd.read_excel(contenido, engine='openpyxl')
    df.columns = ['Empresa','Farmacia','Lugar','Genero','Nombres','Apellidos','Apellido2','Codigo','Cargo']
    df = df.dropna(subset=['Codigo'])
    mapping = {}
    for _, row in df.iterrows():
        codigo = str(row['Codigo']).strip()
        nombres = str(row['Nombres']).strip() if pd.notna(row['Nombres']) else ''
        ap1     = str(row['Apellidos']).strip() if pd.notna(row['Apellidos']) else ''
        nombre_completo = f"{nombres} {ap1}".strip()
        if codigo and nombre_completo:
            mapping[codigo] = nombre_completo
    return mapping

@st.cache_data
def leer_archivo(nombre, contenido):
    df = pd.read_excel(contenido, engine='openpyxl')
    df = df.dropna(subset=['Farmacia'])
    df = df[df['Importe Acumulado'].apply(
        lambda x: str(x).replace('.','').replace('-','').strip().isdigit()
        if pd.notna(x) else False
    )]
    df['Importe Acumulado'] = pd.to_numeric(df['Importe Acumulado'], errors='coerce')
    df['Piezas Acumuladas'] = pd.to_numeric(df['Piezas Acumuladas'], errors='coerce')
    df['Tickets Acum.']     = pd.to_numeric(df['Tickets Acum.'], errors='coerce')
    df['Periodo']           = df['Periodo'].astype(str).str[:6]
    df['Hora']              = pd.to_numeric(df['Hora'], errors='coerce')
    df['Dia']               = pd.to_numeric(df['Dia'], errors='coerce')
    df['Nombre Farmacia']   = df['Farmacia'].str.split('-').str[1:].str.join('-').str.strip()
    df['Es G8']         = df['Premio'].astype(str).str.upper().str.strip() == 'SI'
    return df

def combinar_sin_duplicados(archivos):
    datos_por_periodo = {}
    for archivo in archivos:
        df = leer_archivo(archivo.name, archivo)
        for periodo in df['Periodo'].unique():
            df_p = df[df['Periodo'] == periodo]
            if periodo not in datos_por_periodo or len(df_p) > len(datos_por_periodo[periodo]):
                datos_por_periodo[periodo] = df_p
    if not datos_por_periodo:
        return pd.DataFrame()
    return pd.concat(datos_por_periodo.values(), ignore_index=True)

def card_chart(fig):
    fig.update_layout(
        plot_bgcolor='white', paper_bgcolor='white',
        font_color='#1a2340', font_family='sans-serif',
        margin=dict(l=10, r=10, t=40, b=10),
        title_font_size=13, title_font_color='#1a2340',
        xaxis=dict(gridcolor='#e5e7eb', linecolor='#e5e7eb'),
        yaxis=dict(gridcolor='#e5e7eb', linecolor='#e5e7eb'),
    )
    return fig

def resolver_nombre(codigo, mapping):
    return mapping.get(str(codigo).strip(), str(codigo).strip())

def calcular_metricas(df, grupo):
    base = df.groupby(grupo).agg(
        Importe = ('Importe Acumulado', 'sum'),
        Tickets = ('Tickets Acum.',     'sum'),
        Piezas  = ('Piezas Acumuladas', 'sum')
    ).reset_index()
    premio = df[df['Es G8']].groupby(grupo)['Importe Acumulado'].sum().reset_index()
    premio.columns = [grupo, 'Importe G8']
    base = base.merge(premio, on=grupo, how='left')
    base['Importe G8'] = base['Importe G8'].fillna(0)
    base = base[base['Tickets'] > 0]
    base['Ticket Promedio'] = (base['Importe']  / base['Tickets']).round(0)
    base['Pct G8']      = (base['Importe G8'] / base['Importe'] * 100).round(1)
    base['Arts por Ticket'] = (base['Piezas']   / base['Tickets']).round(2)
    return base

def ranking_html(df_rank, col_nombre, col_valor, formato, color, titulo):
    max_val = df_rank[col_valor].max()
    html = f"""<div style='background:white;border-radius:12px;padding:16px;
                box-shadow:0 1px 4px rgba(0,0,0,0.08);margin-bottom:8px'>
               <p style='font-weight:700;color:#1a2340;margin-bottom:12px;font-size:14px'>{titulo}</p>"""
    for pos, (i, row) in enumerate(df_rank.iterrows(), start=1):
        pct = (row[col_valor] / max_val * 100) if max_val > 0 else 0
        valor_fmt = formato(row[col_valor])
        html += f"""
        <div style='margin-bottom:10px'>
            <div style='display:flex;justify-content:space-between;align-items:center'>
                <span style='color:#374151;font-size:12px;font-weight:600'>#{pos} {row[col_nombre]}</span>
                <span style='color:#1a2340;font-size:13px;font-weight:700'>{valor_fmt}</span>
            </div>
            <div style='background:#e5e7eb;border-radius:4px;height:5px;width:100%;margin-top:4px'>
                <div style='background:{color};height:5px;border-radius:4px;width:{pct:.0f}%'></div>
            </div>
        </div>"""
    html += "</div>"
    return html

def generar_html_benchmarking(bench, meta_ticket, meta_g8_min, meta_g8_max, meta_arts,
                               min_tickets, meta_codigos, periodo_sel, farmacia_sel, rango_fechas=""):
    periodos_str = ", ".join(sorted(periodo_sel)) if periodo_sel else "Todos"
    farmacia_str = ", ".join(farmacia_sel) if isinstance(farmacia_sel, list) else str(farmacia_sel)

    def dot_color(val, meta_min, meta_max=None):
        if meta_max is not None:
            if meta_min <= val <= meta_max:
                return "#10b981"
            diff = min(abs(val - meta_min), abs(val - meta_max))
            if diff <= 3:
                return "#f59e0b"
            return "#ef4444"
        if meta_min == 0:
            return "#9ca3af"
        p = val / meta_min * 100
        if p >= 100:
            return "#10b981"
        if p >= 80:
            return "#f59e0b"
        return "#ef4444"

    def dot(color):
        return (
            "<span style=\"display:inline-block;width:10px;height:10px;"
            "border-radius:50%;background:" + color + ";"
            "margin-right:5px;vertical-align:middle\"></span>"
        )

    def score_bar_html(score):
        c = "#10b981" if score >= 95 else ("#f59e0b" if score >= 80 else "#ef4444")
        return (
            "<div style=\"display:flex;align-items:center;gap:6px\">"
            "<div style=\"flex:1;background:#e5e7eb;border-radius:4px;height:8px\">"
            "<div style=\"background:" + c + ";height:8px;border-radius:4px;"
            "width:" + str(int(score)) + "%\"></div></div>"
            "<span style=\"font-weight:700;font-size:12px;color:#1a2340;"
            "min-width:34px\">" + str(score) + "</span></div>"
        )

    filas = ""
    for i, row in bench.iterrows():
        bg  = "#f9fafb" if i % 2 == 0 else "white"
        c_t = dot_color(row["Ticket Prom"],  meta_ticket)
        c_g = dot_color(row["Pct G8"],       meta_g8_min, meta_g8_max)
        c_a = dot_color(row["Arts Ticket"],  meta_arts)
        c_n = dot_color(row["Num Tickets"],  min_tickets)
        c_c = dot_color(row["Codigos"],      meta_codigos)
        filas += (
            "<tr style=\"background:" + bg + ";border-bottom:1px solid #f3f4f6\">"
            "<td style=\"padding:8px 10px;font-size:12px;color:#9ca3af\">" + str(i) + "</td>"
            "<td style=\"padding:8px 10px;font-size:13px;font-weight:600;color:#1a2340\">"
            + str(row["Nombre Vendedor"]) + "</td>"
            "<td style=\"padding:8px 10px;text-align:center;font-size:12px\">"
            + dot(c_t) + "$ " + "{:,.0f}".format(row["Ticket Prom"]) + "</td>"
            "<td style=\"padding:8px 10px;text-align:center;font-size:12px\">"
            + dot(c_g) + "{:.1f}".format(row["Pct G8"]) + "%</td>"
            "<td style=\"padding:8px 10px;text-align:center;font-size:12px\">"
            + dot(c_a) + "{:.2f}".format(row["Arts Ticket"]) + "</td>"
            "<td style=\"padding:8px 10px;text-align:center;font-size:12px\">"
            + dot(c_n) + "{:,}".format(int(row["Num Tickets"])) + "</td>"
            "<td style=\"padding:8px 10px;text-align:center;font-size:12px\">"
            + dot(c_c) + str(int(row["Codigos"])) + "</td>"
            "<td style=\"padding:8px 10px;min-width:150px\">"
            + score_bar_html(row["Score"]) + "</td>"
            "</tr>"
        )

    resumen_verde    = len(bench[bench["Score"] >= 95])
    resumen_amarillo = len(bench[(bench["Score"] >= 80) & (bench["Score"] < 95)])
    resumen_rojo     = len(bench[bench["Score"] < 80])

    html = (
        "<div style=\"font-family:sans-serif;background:white;padding:24px\">"
        "<h2 style=\"margin:0 0 4px;color:#1a2340\">🎯 Benchmarking de vendedores</h2>"
        "<p style=\"margin:0 0 16px;color:#6b7280;font-size:12px\">"
        "📅 Período: " + periodos_str
        + (" (" + rango_fechas + ")" if rango_fechas else "")
        + " &nbsp;|&nbsp; 🏪 " + farmacia_str + "</p>"
        ""
        ""
        "<div style=\"display:flex;gap:24px;margin-bottom:16px;flex-wrap:wrap\">"
        "<div><div style=\"font-size:11px;color:#6b7280;text-transform:uppercase\">🎫 Ticket Promedio meta</div>"
        "<div style=\"font-size:20px;font-weight:700;color:#1a2340\">$ "
        + "{:,.0f}".format(meta_ticket) + "</div></div>"
        "<div><div style=\"font-size:11px;color:#6b7280;text-transform:uppercase\">🏅 G8 rango</div>"
        "<div style=\"font-size:20px;font-weight:700;color:#1a2340\">"
        + "{:.1f}".format(meta_g8_min) + "% - " + "{:.1f}".format(meta_g8_max) + "%</div></div>"
        "<div><div style=\"font-size:11px;color:#6b7280;text-transform:uppercase\">🛒 Arts/tick meta</div>"
        "<div style=\"font-size:20px;font-weight:700;color:#1a2340\">"
        + "{:.2f}".format(meta_arts) + "</div></div>"
        "<div><div style=\"font-size:11px;color:#6b7280;text-transform:uppercase\">🧾 Tickets mínimos</div>"
        "<div style=\"font-size:20px;font-weight:700;color:#1a2340\">"
        + "{:,}".format(int(min_tickets)) + "</div></div>"
        "<div><div style=\"font-size:11px;color:#6b7280;text-transform:uppercase\">🔢 Códigos meta</div>"
        "<div style=\"font-size:20px;font-weight:700;color:#1a2340\">"
        + str(int(meta_codigos)) + "</div></div>"
        "</div>"
        "<div style=\"display:flex;gap:12px;margin-bottom:20px\">"
        "<div style=\"flex:1;background:#dcfce7;border-radius:8px;padding:12px;text-align:center\">"
        "<div style=\"font-size:26px;font-weight:800;color:#10b981\">" + str(resumen_verde) + "</div>"
        "<div style=\"font-size:12px;color:#374151\">Sobre meta (Score mayor 95)</div></div>"
        "<div style=\"flex:1;background:#fef9c3;border-radius:8px;padding:12px;text-align:center\">"
        "<div style=\"font-size:26px;font-weight:800;color:#ca8a04\">" + str(resumen_amarillo) + "</div>"
        "<div style=\"font-size:12px;color:#374151\">Cerca de meta (Score 80-94)</div></div>"
        "<div style=\"flex:1;background:#fee2e2;border-radius:8px;padding:12px;text-align:center\">"
        "<div style=\"font-size:26px;font-weight:800;color:#dc2626\">" + str(resumen_rojo) + "</div>"
        "<div style=\"font-size:12px;color:#374151\">Bajo meta (Score menor 80)</div></div>"
        "</div>"
        "<table style=\"width:100%;border-collapse:collapse\">"
        "<thead><tr style=\"background:#f8faff;border-bottom:2px solid #e5e7eb\">"
        "<th style=\"padding:10px;text-align:left;font-size:11px;color:#6b7280;width:30px\">#</th>"
        "<th style=\"padding:10px;text-align:left;font-size:11px;color:#6b7280\">VENDEDOR</th>"
        "<th style=\"padding:10px;text-align:center;font-size:11px;color:#6b7280\">TICKET PROMEDIO</th>"
        "<th style=\"padding:10px;text-align:center;font-size:11px;color:#6b7280\">G8%</th>"
        "<th style=\"padding:10px;text-align:center;font-size:11px;color:#6b7280\">ARTS/TICK</th>"
        "<th style=\"padding:10px;text-align:center;font-size:11px;color:#6b7280\">TICKETS</th>"
        "<th style=\"padding:10px;text-align:center;font-size:11px;color:#6b7280\">CODIGOS</th>"
        "<th style=\"padding:10px;text-align:left;font-size:11px;color:#6b7280;min-width:150px\">SCORE</th>"
        "</tr></thead>"
        "<tbody>" + filas + "</tbody></table>"
        "<p style=\"margin-top:20px;font-size:11px;color:#9ca3af;text-align:center\">"
        "Grupo Baco - Benchmarking de Vendedores - Farmacias Dr. Simi</p>"
        "</div>"
    )
    return html


def generar_html_impresion(ticket_prom, pct_premio, arts_ticket,
                            rank_local, rank_vend, periodo_sel, farmacia_sel):
    def filas_ranking(df, col_nombre, col_valor, formato, color):
        html = ""
        max_v = df[col_valor].max()
        for i, row in df.iterrows():
            pct = (row[col_valor] / max_v * 100) if max_v > 0 else 0
            html += f"""
            <tr>
              <td style='padding:4px 8px;font-size:12px;color:#374151'>#{i+1} {row[col_nombre]}</td>
              <td style='padding:4px 8px;font-size:12px;font-weight:700;text-align:right'>{formato(row[col_valor])}</td>
              <td style='padding:4px 8px;width:120px'>
                <div style='background:#e5e7eb;border-radius:3px;height:6px'>
                  <div style='background:{color};height:6px;border-radius:3px;width:{pct:.0f}%'></div>
                </div>
              </td>
            </tr>"""
        return html

    html = (
        "<div style='font-family:sans-serif;background:white;padding:24px'>"
        "<div style='display:flex;justify-content:space-between;align-items:center;"
        "margin-bottom:20px;border-bottom:2px solid #3b82f6;padding-bottom:12px'>"
        "<div><h2 style='margin:0;color:#1a2340'>💊 Informe de Gestión — Farmacias Dr. Simi</h2>"
        "<p style='margin:4px 0 0;color:#6b7280;font-size:13px'>Grupo Baco · Período: "
        + periodos_str + " · " + farmacia_str +
        "</p></div></div>"
        "<h3 style='color:#1a2340;margin-bottom:12px'>Indicadores de gestión</h3>"
        "<div style='display:flex;gap:16px;margin-bottom:24px'>"
        "<div style='flex:1;border-radius:10px;padding:16px;border-top:4px solid #3b82f6;background:#f8faff'>"
        "<p style='margin:0;font-size:12px;color:#6b7280;text-transform:uppercase'>🎫 Ticket promedio</p>"
        "<p style='margin:8px 0 0;font-size:24px;font-weight:700;color:#1a2340'>" + ticket_str + "</p></div>"
        "<div style='flex:1;border-radius:10px;padding:16px;border-top:4px solid #10b981;background:#f0fdf4'>"
        "<p style='margin:0;font-size:12px;color:#6b7280;text-transform:uppercase'>🏅 % G8 s/ venta</p>"
        "<p style='margin:8px 0 0;font-size:24px;font-weight:700;color:#1a2340'>" + premio_str + "</p></div>"
        "<div style='flex:1;border-radius:10px;padding:16px;border-top:4px solid #f59e0b;background:#fffbeb'>"
        "<p style='margin:0;font-size:12px;color:#6b7280;text-transform:uppercase'>🛒 Artículos por ticket</p>"
        "<p style='margin:8px 0 0;font-size:24px;font-weight:700;color:#1a2340'>" + arts_str + "</p></div>"
        "</div>"
        "<div style='display:flex;gap:20px;margin-bottom:20px'>"
        "<div style='flex:1'><h4 style='color:#1a2340;margin-bottom:8px'>🎫 Ticket promedio — Por local</h4>"
        "<table style='width:100%;border-collapse:collapse'>" + t_tick_loc + "</table></div>"
        "<div style='flex:1'><h4 style='color:#1a2340;margin-bottom:8px'>🎫 Ticket promedio — Por vendedor</h4>"
        "<table style='width:100%;border-collapse:collapse'>" + t_tick_ven + "</table></div>"
        "</div>"
        "<div style='display:flex;gap:20px;margin-bottom:20px'>"
        "<div style='flex:1'><h4 style='color:#1a2340;margin-bottom:8px'>🏅 % G8 — Por local</h4>"
        "<table style='width:100%;border-collapse:collapse'>" + t_prem_loc + "</table></div>"
        "<div style='flex:1'><h4 style='color:#1a2340;margin-bottom:8px'>🏅 % G8 — Por vendedor</h4>"
        "<table style='width:100%;border-collapse:collapse'>" + t_prem_ven + "</table></div>"
        "</div>"
        "<div style='display:flex;gap:20px'>"
        "<div style='flex:1'><h4 style='color:#1a2340;margin-bottom:8px'>🛒 Arts/ticket — Por local</h4>"
        "<table style='width:100%;border-collapse:collapse'>" + t_arts_loc + "</table></div>"
        "<div style='flex:1'><h4 style='color:#1a2340;margin-bottom:8px'>🛒 Arts/ticket — Por vendedor</h4>"
        "<table style='width:100%;border-collapse:collapse'>" + t_arts_ven + "</table></div>"
        "</div>"
        "<p style='margin-top:24px;font-size:11px;color:#9ca3af;text-align:center'>"
        "💊 Grupo Baco · Dashboard de Ventas · Farmacias Dr. Simi</p>"
        "</div>"
    )
    return html

# ── SIDEBAR ──────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 💊 Farmacias")
    st.markdown("*Panel de gestión*")
    st.markdown("---")

    pagina = st.radio(
        "NAVEGACIÓN",
        ["🏠 Dashboard", "📈 Ventas", "📦 Compras"],
        index=1,
        label_visibility="collapsed"
    )
    st.markdown("---")

    archivo_personal = st.file_uploader(
        "PERSONAL (opcional)", type=["xlsx"],
        help="Sube el archivo de personal para ver nombres."
    )
    mapping_nombres = {}
    if archivo_personal:
        mapping_nombres = cargar_personal(archivo_personal)
        st.caption(f"✅ {len(mapping_nombres)} vendedores cargados")

    st.markdown("---")
    archivos = st.file_uploader(
        "CARGAR DATOS DE VENTAS", type=["xlsx"],
        accept_multiple_files=True
    )

    farmacia_sel  = ["Todas"]
    periodo_sel   = []
    familia_sel   = "Todas las familias"
    excluir_vend  = []

    if archivos:
        # Guardar TODOS los archivos en session_state
        st.session_state["vtas_files"] = archivos
        with st.spinner("Procesando..."):
            df_raw = combinar_sin_duplicados(archivos)

        df_raw['Nombre Vendedor'] = df_raw['Vendedor'].apply(
            lambda x: resolver_nombre(x, mapping_nombres)
        )

        st.markdown("**ARCHIVOS CARGADOS**")
        for p in sorted(df_raw['Periodo'].unique()):
            filas = len(df_raw[df_raw['Periodo'] == p])
            st.caption(f"✅ {p} · {filas:,} registros")

        st.markdown("---")

        # ✅ MULTISELECT de farmacias
        farmacias_lista = sorted(df_raw['Nombre Farmacia'].dropna().unique().tolist())
        farmacia_sel = st.multiselect(
            "FARMACIA",
            farmacias_lista,
            default=farmacias_lista,
            help="Puedes elegir una o varias farmacias."
        )

        periodos_op = sorted(df_raw['Periodo'].dropna().unique().tolist())
        periodo_sel = st.multiselect("PERÍODO", periodos_op, default=periodos_op)

        familias_op = ["Todas las familias"] + sorted(df_raw['Familia'].dropna().unique().tolist())
        familia_sel = st.selectbox("FAMILIA", familias_op)

        todos_vendedores = sorted(df_raw['Nombre Vendedor'].dropna().unique().tolist())
        excluir_vend = st.multiselect(
            "EXCLUIR VENDEDORES",
            todos_vendedores,
            help="Selecciona los vendedores QF u otros que quieras excluir."
        )

        st.markdown("---")
        st.markdown("**📅 FILTRO DE DÍAS**")

        orden_dias = ['LUNES','MARTES','MIERCOLES','JUEVES','VIERNES','SABADO','DOMINGO']
        dias_disponibles = [d for d in orden_dias if d in df_raw['Dia Semana'].unique()]
        dias_sel = st.multiselect(
            "DÍA DE LA SEMANA",
            dias_disponibles,
            default=dias_disponibles,
            help="Filtra por día de la semana"
        )

        dias_mes_disponibles = sorted(df_raw['Dia'].dropna().unique().astype(int).tolist())
        dias_mes_sel = st.multiselect(
            "DÍA DEL MES",
            dias_mes_disponibles,
            default=[],
            help="Deja vacío para ver todos los días. Selecciona uno o más días específicos."
        )

        st.markdown("---")
        st.markdown("**🎯 METAS MANUALES**")
        st.caption("Define tú las metas para el benchmarking")

        meta_ticket_input  = st.number_input("🎫 Ticket promedio ($)",    min_value=0,   value=9000,  step=100)
        st.caption("🏅 G8 — rango aceptable")
        g8col1, g8col2 = st.columns(2)
        with g8col1:
            meta_g8_min_input = st.number_input("Mín %", min_value=0.0, value=30.0, step=0.5, format="%.1f")
        with g8col2:
            meta_g8_max_input = st.number_input("Máx %", min_value=0.0, value=45.0, step=0.5, format="%.1f")
        meta_arts_input    = st.number_input("🛒 Artículos por ticket",   min_value=0.0, value=3.0,   step=0.1, format="%.1f")
        meta_codigos_input = st.number_input("🔢 Códigos distintos",      min_value=0,   value=700,   step=10)
        st.caption("🧾 Mínimo de tickets para ser evaluado")
        min_tickets_input  = st.number_input("Tickets mínimos",           min_value=0,   value=500,   step=50)

        usar_metas_manuales = st.toggle("Usar metas manuales", value=False,
            help="Activa para usar tus metas. Desactiva para usar TOP 25% automático.")

        st.markdown("---")
        st.caption(f"🟢 En vivo · {len(df_raw):,} registros")

# ── MAIN ─────────────────────────────────────────────────
# ── NAVEGACIÓN ───────────────────────────────────────────
if pagina == "📦 Compras":
    render_compras()
    st.stop()

if pagina == "🏠 Dashboard":
    st.markdown("## 🏠 Dashboard General")
    st.caption("Resumen ejecutivo · Grupo Baco · Dr. Simi")
    st.info("👈 Selecciona **📈 Ventas** o **📦 Compras** para ver los dashboards detallados.")
    st.stop()

st.markdown("## 📊 Dashboard de Ventas")
st.caption("General · Comercial · Datos reales · Grupo Baco · Dr. Simi")

if not archivos:
    st.info("👈 Carga uno o más archivos Excel de ventas desde el panel izquierdo para comenzar.")
    st.stop()

# FILTROS
df = df_raw.copy()
if farmacia_sel:
    df = df[df['Nombre Farmacia'].isin(farmacia_sel)]
if periodo_sel:
    df = df[df['Periodo'].isin(periodo_sel)]
if familia_sel != "Todas las familias":
    df = df[df['Familia'] == familia_sel]
if excluir_vend:
    df = df[~df['Nombre Vendedor'].isin(excluir_vend)]

# Filtro día de la semana
if dias_sel and len(dias_sel) < len(dias_disponibles):
    df = df[df['Dia Semana'].isin(dias_sel)]

# Filtro día del mes
if dias_mes_sel:
    df = df[df['Dia'].isin(dias_mes_sel)]

if df.empty:
    st.warning("⚠️ No hay datos con los filtros seleccionados.")
    st.stop()

# ── KPIs FILA 1 ───────────────────────────────────────────
ingresos       = df['Importe Acumulado'].sum()
tickets        = df['Tickets Acum.'].sum()
piezas         = df['Piezas Acumuladas'].sum()
ticket_prom    = ingresos / tickets if tickets > 0 else 0
importe_premio = df[df['Es G8']]['Importe Acumulado'].sum()
pct_premio     = (importe_premio / ingresos * 100) if ingresos > 0 else 0
arts_ticket    = (piezas / tickets) if tickets > 0 else 0

k1, k2, k3, k4 = st.columns(4)
with k1:
    st.metric("💰 Ingresos",          f"$ {ingresos:,.0f}")
with k2:
    st.metric("🧾 Tickets emitidos",  f"{tickets:,.0f}")
with k3:
    st.metric("📦 Piezas vendidas",   f"{piezas:,.0f}")
with k4:
    st.metric("🏪 Farmacias activas", f"{df['Nombre Farmacia'].nunique()}")

st.markdown("<br>", unsafe_allow_html=True)

# ── KPIs FILA 2: Gestión + Botón imprimir ────────────────
st.markdown("#### Indicadores de gestión")
g1, g2, g3 = st.columns(3)
with g1:
    st.metric("🎫 Ticket promedio",      f"$ {ticket_prom:,.0f}")
with g2:
    st.metric("🏅 % G8 s/ venta",    f"{pct_premio:.1f}%")
with g3:
    st.metric("🛒 Artículos por ticket", f"{arts_ticket:.2f}")
imprimir = False

st.markdown("---")

# ── IMPRESIÓN ─────────────────────────────────────────────
rank_local = calcular_metricas(df, 'Nombre Farmacia')
rank_vend  = calcular_metricas(df, 'Nombre Vendedor')

if imprimir:
    html_informe = generar_html_benchmarking(
        bench, meta_ticket, meta_g8_min, meta_g8_max, meta_arts,
        min_tickets, meta_codigos,
        periodo_sel, farmacia_sel
    )
    st.markdown(html_informe, unsafe_allow_html=True)
    st.components.v1.html(
        "<script>window.onload=function(){window.print();}</script>",
        height=0
    )
    st.stop()

# ── EVOLUCIÓN MENSUAL ─────────────────────────────────────
st.markdown("### 📅 Evolución mensual")
evol = df.groupby('Periodo').agg(
    Ingresos = ('Importe Acumulado', 'sum'),
    Tickets  = ('Tickets Acum.',     'sum'),
    Piezas   = ('Piezas Acumuladas', 'sum')
).reset_index().sort_values('Periodo')

evol['Ticket Prom']    = (evol['Ingresos'] / evol['Tickets']).round(0)
evol['Periodo Label']  = evol['Periodo'].apply(
    lambda x: ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'][int(str(x)[4:6])-1]
              + ' ' + str(x)[:4]
)

ev1, ev2, ev3 = st.columns(3)

with ev1:
    fig_ev1 = px.bar(evol, x='Periodo Label', y='Ingresos',
                     title='💰 Ingresos por mes',
                     color_discrete_sequence=[AZUL], text_auto='.2s')
    fig_ev1 = card_chart(fig_ev1)
    fig_ev1.update_layout(xaxis_title='', yaxis_title='$ Ingresos', showlegend=False)
    st.plotly_chart(fig_ev1, use_container_width=True)

with ev2:
    fig_ev2 = px.bar(evol, x='Periodo Label', y='Tickets',
                     title='🧾 Tickets por mes',
                     color_discrete_sequence=[VERDE], text_auto='.2s')
    fig_ev2 = card_chart(fig_ev2)
    fig_ev2.update_layout(xaxis_title='', yaxis_title='Tickets', showlegend=False)
    st.plotly_chart(fig_ev2, use_container_width=True)

with ev3:
    fig_ev3 = px.bar(evol, x='Periodo Label', y='Ticket Prom',
                     title='🎫 Ticket promedio por mes',
                     color_discrete_sequence=[NARANJA], text_auto='.2s')
    fig_ev3 = card_chart(fig_ev3)
    fig_ev3.update_layout(xaxis_title='', yaxis_title='$ Ticket prom.', showlegend=False)
    st.plotly_chart(fig_ev3, use_container_width=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── FILA: Área diaria + Donut ─────────────────────────────
col1, col2 = st.columns([2, 1])

with col1:
    _df_sem = df.copy()
    _df_sem['_fecha'] = pd.to_datetime(
        _df_sem['Periodo'].astype(str).str[:4] + '-' +
        _df_sem['Periodo'].astype(str).str[4:6] + '-' +
        _df_sem['Dia'].fillna(1).astype(int).astype(str).str.zfill(2), errors='coerce')
    _df_sem['Lunes'] = _df_sem['_fecha'] - pd.to_timedelta(_df_sem['_fecha'].dt.dayofweek, unit='d')
    _df_sem = _df_sem.dropna(subset=['Lunes'])
    ventas_sem = _df_sem.groupby('Lunes').agg(
        Ingresos = ('Importe Acumulado', 'sum'),
        Tickets  = ('Tickets Acum.',     'sum')
    ).reset_index().sort_values('Lunes')
    ventas_sem['Ticket Prom'] = (ventas_sem['Ingresos'] / ventas_sem['Tickets']).round(0)
    ventas_sem['Label'] = ventas_sem['Lunes'].dt.strftime('%d/%m')
    import plotly.graph_objects as go
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=ventas_sem['Label'], y=ventas_sem['Ingresos'],
        name='Ingresos', marker_color=AZUL, yaxis='y1'
    ))
    fig.add_trace(go.Scatter(
        x=ventas_sem['Label'], y=ventas_sem['Ticket Prom'],
        name='Ticket promedio', mode='lines+markers',
        line=dict(color=NARANJA, width=2),
        marker=dict(size=5), yaxis='y2'
    ))
    fig.update_layout(
        title='Ingresos semanales + ticket promedio',
        plot_bgcolor='white', paper_bgcolor='white',
        font_color='#1a2340', font_family='sans-serif',
        margin=dict(l=10, r=10, t=40, b=10),
        title_font_size=13,
        xaxis=dict(title='Semana (lunes)', gridcolor='#e5e7eb'),
        yaxis=dict(title='$ Ingresos', gridcolor='#e5e7eb'),
        yaxis2=dict(title='$ Ticket prom.', overlaying='y', side='right', showgrid=False),
        legend=dict(orientation='h', y=-0.2)
    )
    st.plotly_chart(fig, use_container_width=True)

with col2:
    top_fam = df.groupby('Familia')['Importe Acumulado'].sum().nlargest(8).reset_index()
    fig2 = px.pie(top_fam, names='Familia', values='Importe Acumulado',
                  title='Ventas por familia terapéutica',
                  hole=0.5, color_discrete_sequence=AZUL_LISTA)
    fig2 = card_chart(fig2)
    fig2.update_layout(legend=dict(font=dict(size=9)))
    st.plotly_chart(fig2, use_container_width=True)

# ── FILA: Farmacia + Hora ─────────────────────────────────
col3, col4 = st.columns([1, 1])

with col3:
    por_farm = df.groupby('Nombre Farmacia')['Importe Acumulado'].sum().sort_values(ascending=True).reset_index()
    fig3 = px.bar(por_farm, x='Importe Acumulado', y='Nombre Farmacia',
                  orientation='h', title='Ventas por farmacia',
                  color_discrete_sequence=[AZUL])
    fig3 = card_chart(fig3)
    fig3.update_layout(xaxis_title='$ Ingresos', yaxis_title='', showlegend=False)
    st.plotly_chart(fig3, use_container_width=True)

with col4:
    _hora = df.groupby('Hora').agg(
        Ingresos  = ('Importe Acumulado', 'sum'),
        Tickets   = ('Tickets Acum.',     'sum'),
        Piezas    = ('Piezas Acumuladas', 'sum')
    ).reset_index().sort_values('Hora')
    _hora = _hora[_hora['Tickets'] > 0]
    _hora['Ticket Prom'] = (_hora['Ingresos'] / _hora['Tickets']).round(0)
    _hora['Arts Ticket'] = (_hora['Piezas']   / _hora['Tickets']).round(2)
    import plotly.graph_objects as go
    fig4 = go.Figure()
    fig4.add_trace(go.Bar(
        x=_hora['Hora'], y=_hora['Ingresos'],
        name='$ Ingresos', marker_color=AZUL, yaxis='y1'
    ))
    fig4.add_trace(go.Scatter(
        x=_hora['Hora'], y=_hora['Ticket Prom'],
        name='Ticket promedio', mode='lines+markers',
        line=dict(color=NARANJA, width=2),
        marker=dict(size=5), yaxis='y2'
    ))
    fig4.add_trace(go.Scatter(
        x=_hora['Hora'], y=_hora['Arts Ticket'],
        name='Arts/ticket', mode='lines+markers',
        line=dict(color=VERDE, width=2, dash='dot'),
        marker=dict(size=5), yaxis='y3'
    ))
    fig4.update_layout(
        title='Ingresos · Ticket promedio · Arts/ticket por hora',
        plot_bgcolor='white', paper_bgcolor='white',
        font_color='#1a2340', font_family='sans-serif',
        margin=dict(l=10, r=60, t=40, b=10),
        title_font_size=13,
        xaxis=dict(title='Hora', gridcolor='#e5e7eb'),
        yaxis =dict(title='$ Ingresos',    gridcolor='#e5e7eb'),
        yaxis2=dict(title='$ Ticket prom.', overlaying='y', side='right', showgrid=False, position=1.0),
        yaxis3=dict(title='Arts/ticket',    overlaying='y', side='right', showgrid=False, anchor='free', position=0.92),
        legend=dict(orientation='h', y=-0.25)
    )
    st.plotly_chart(fig4, use_container_width=True)

# ── Top vendedores + Top G8 ───────────────────────────────
col5, col6 = st.columns(2)

with col5:
    top_vend = df.groupby('Nombre Vendedor')['Importe Acumulado'].sum().nlargest(10).reset_index()
    top_vend.columns = ['Vendedor','Ventas']
    top_vend = top_vend.reset_index(drop=True)
    st.markdown("### 🏆 Top vendedores · Venta total")
    for i, row in top_vend.iterrows():
        pct = row['Ventas'] / top_vend['Ventas'].max()
        st.markdown(f"""
        <div style='margin-bottom:10px;background:white;border-radius:8px;padding:8px 12px;box-shadow:0 1px 3px rgba(0,0,0,0.07)'>
            <span style='color:#6b7280;font-size:11px'>#{i+1} {row['Vendedor']}</span><br>
            <div style='background:#e5e7eb;border-radius:4px;height:5px;width:100%;margin:4px 0'>
                <div style='background:{AZUL};height:5px;border-radius:4px;width:{pct*100:.0f}%'></div>
            </div>
            <span style='color:#1a2340;font-size:13px;font-weight:700'>$ {row['Ventas']/1e6:.1f}M</span>
        </div>
        """, unsafe_allow_html=True)

with col6:
    df_g8   = df[df['Es G8']]
    top_g8  = df_g8.groupby('Nombre Vendedor')['Importe Acumulado'].sum().nlargest(10).reset_index()
    top_g8.columns = ['Vendedor','Ventas G8']
    top_g8  = top_g8.reset_index(drop=True)
    st.markdown("### 🥇 Top vendedores · Venta G8")
    for i, row in top_g8.iterrows():
        pct = row['Ventas G8'] / top_g8['Ventas G8'].max()
        st.markdown(f"""
        <div style='margin-bottom:10px;background:white;border-radius:8px;padding:8px 12px;box-shadow:0 1px 3px rgba(0,0,0,0.07)'>
            <span style='color:#6b7280;font-size:11px'>#{i+1} {row['Vendedor']}</span><br>
            <div style='background:#e5e7eb;border-radius:4px;height:5px;width:100%;margin:4px 0'>
                <div style='background:{DORADO};height:5px;border-radius:4px;width:{pct*100:.0f}%'></div>
            </div>
            <span style='color:#1a2340;font-size:13px;font-weight:700'>$ {row['Ventas G8']/1e6:.1f}M</span>
        </div>
        """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── Tickets por local y por vendedor ─────────────────────
st.markdown("### 🧾 Número de tickets")
t1, t2 = st.columns(2)

with t1:
    tick_local = df.groupby('Nombre Farmacia')['Tickets Acum.'].sum().sort_values(ascending=True).reset_index()
    tick_local.columns = ['Farmacia','Tickets']
    fig_tl = px.bar(tick_local, x='Tickets', y='Farmacia',
                    orientation='h', title='🏪 Tickets por local',
                    color_discrete_sequence=[AZUL], text='Tickets')
    fig_tl = card_chart(fig_tl)
    fig_tl.update_traces(texttemplate='%{text:,.0f}', textposition='outside')
    fig_tl.update_layout(xaxis_title='N° tickets', yaxis_title='', showlegend=False)
    st.plotly_chart(fig_tl, use_container_width=True)

with t2:
    tick_vend = df.groupby('Nombre Vendedor')['Tickets Acum.'].sum().nlargest(12).sort_values(ascending=True).reset_index()
    tick_vend.columns = ['Vendedor','Tickets']
    fig_tv = px.bar(tick_vend, x='Tickets', y='Vendedor',
                    orientation='h', title='👤 Tickets por vendedor · Top 12',
                    color_discrete_sequence=[VERDE], text='Tickets')
    fig_tv = card_chart(fig_tv)
    fig_tv.update_traces(texttemplate='%{text:,.0f}', textposition='outside')
    fig_tv.update_layout(xaxis_title='N° tickets', yaxis_title='', showlegend=False)
    st.plotly_chart(fig_tv, use_container_width=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── Ranking códigos por vendedor ──────────────────────────
st.markdown("### 🔢 Códigos distintos vendidos por trabajador · Top 12")
codigos_x_vend = df.groupby('Nombre Vendedor')['Producto'].nunique().nlargest(12).reset_index()
codigos_x_vend.columns = ['Vendedor','Códigos']
codigos_x_vend = codigos_x_vend.sort_values('Códigos', ascending=True)
fig_cod = px.bar(codigos_x_vend, x='Códigos', y='Vendedor',
                 orientation='h',
                 title='Cantidad de códigos distintos vendidos · Top 12',
                 color_discrete_sequence=[VERDE], text='Códigos')
fig_cod = card_chart(fig_cod)
fig_cod.update_traces(textposition='outside')
fig_cod.update_layout(xaxis_title='Nº de códigos distintos', yaxis_title='', showlegend=False)
st.plotly_chart(fig_cod, use_container_width=True)

st.markdown("<br>", unsafe_allow_html=True)

st.markdown("## 📊 Rankings por indicador")
st.markdown("---")

st.markdown("### 🎫 Ticket promedio")
r1a, r1b = st.columns(2)
with r1a:
    datos = rank_local.nlargest(11,'Ticket Promedio')[['Nombre Farmacia','Ticket Promedio']].reset_index(drop=True)
    datos.index += 1
    st.markdown(ranking_html(datos,'Nombre Farmacia','Ticket Promedio',
        lambda x: f"$ {x:,.0f}", AZUL, "🏪 Por local"), unsafe_allow_html=True)
with r1b:
    datos = rank_vend.nlargest(10,'Ticket Promedio')[['Nombre Vendedor','Ticket Promedio']].reset_index(drop=True)
    datos.index += 1
    st.markdown(ranking_html(datos,'Nombre Vendedor','Ticket Promedio',
        lambda x: f"$ {x:,.0f}", AZUL, "👤 Por vendedor"), unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)
st.markdown("### 🏅 % G8 sobre venta total")
r2a, r2b = st.columns(2)
with r2a:
    datos = rank_local.nlargest(11,'Pct G8')[['Nombre Farmacia','Pct G8']].reset_index(drop=True)
    datos.index += 1
    st.markdown(ranking_html(datos,'Nombre Farmacia','Pct G8',
        lambda x: f"{x:.1f}%", VERDE, "🏪 Por local"), unsafe_allow_html=True)
with r2b:
    datos = rank_vend.nlargest(10,'Pct G8')[['Nombre Vendedor','Pct G8']].reset_index(drop=True)
    datos.index += 1
    st.markdown(ranking_html(datos,'Nombre Vendedor','Pct G8',
        lambda x: f"{x:.1f}%", VERDE, "👤 Por vendedor"), unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)
st.markdown("### 🛒 Artículos por ticket")
r3a, r3b = st.columns(2)
with r3a:
    datos = rank_local.nlargest(11,'Arts por Ticket')[['Nombre Farmacia','Arts por Ticket']].reset_index(drop=True)
    datos.index += 1
    st.markdown(ranking_html(datos,'Nombre Farmacia','Arts por Ticket',
        lambda x: f"{x:.2f}", NARANJA, "🏪 Por local"), unsafe_allow_html=True)
with r3b:
    datos = rank_vend.nlargest(10,'Arts por Ticket')[['Nombre Vendedor','Arts por Ticket']].reset_index(drop=True)
    datos.index += 1
    st.markdown(ranking_html(datos,'Nombre Vendedor','Arts por Ticket',
        lambda x: f"{x:.2f}", NARANJA, "👤 Por vendedor"), unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── BENCHMARKING: SEMÁFORO + SCORE ──────────────────────
st.markdown("---")
_col_bench_titulo, _col_bench_btn = st.columns([3, 1])
with _col_bench_titulo:
    st.markdown("## 🎯 Benchmarking de vendedores")
    st.caption("Meta = promedio del TOP 25% · Score 0–100 ponderado por influencia en venta")
with _col_bench_btn:
    st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
    imprimir_bench = st.button("🖨️ Imprimir", use_container_width=True, type="primary", key="btn_bench")

# ── Calcular métricas por vendedor ───────────────────────
bench = df.groupby('Nombre Vendedor').agg(
    Importe = ('Importe Acumulado', 'sum'),
    Tickets = ('Tickets Acum.',     'sum'),
    Piezas  = ('Piezas Acumuladas', 'sum'),
    Codigos = ('Producto',          'nunique')
).reset_index()

premio_vend = df[df['Es G8']].groupby('Nombre Vendedor')['Importe Acumulado'].sum().reset_index()
premio_vend.columns = ['Nombre Vendedor','Premio']
bench = bench.merge(premio_vend, on='Nombre Vendedor', how='left')
bench['Premio'] = bench['Premio'].fillna(0)

bench = bench[bench['Tickets'] > 0]
bench['Ticket Prom']  = (bench['Importe'] / bench['Tickets']).round(0)
bench['Pct G8']   = (bench['Premio']  / bench['Importe'] * 100).round(1)
bench['Arts Ticket']  = (bench['Piezas']  / bench['Tickets']).round(2)
bench['Num Tickets']  = bench['Tickets'].round(0)

# ── Calcular metas (TOP 25% o manual) ────────────────────
def meta_top25(serie):
    corte = serie.quantile(0.75)
    return serie[serie >= corte].mean()

if usar_metas_manuales:
    meta_ticket   = float(meta_ticket_input)
    meta_g8_min   = float(meta_g8_min_input)
    meta_g8_max   = float(meta_g8_max_input)
    meta_arts     = float(meta_arts_input)
    meta_codigos  = float(meta_codigos_input)
    min_tickets   = float(min_tickets_input)
else:
    meta_ticket   = meta_top25(bench['Ticket Prom'])
    meta_g8_min   = bench['Pct G8'].quantile(0.25)
    meta_g8_max   = bench['Pct G8'].quantile(0.75)
    meta_arts     = meta_top25(bench['Arts Ticket'])
    meta_codigos  = meta_top25(bench['Codigos'])
    min_tickets   = bench['Num Tickets'].quantile(0.25)

# Filtrar vendedores evaluables por minimo de tickets
# (sin filtro de tickets mínimos — se muestran todos los vendedores)

# ── Score 0-100 por vendedor ──────────────────────────────
W = {'ticket': 0.35, 'g8': 0.10, 'arts': 0.25, 'codigos': 0.20, 'ntickets': 0.10}

def score_ind(val, meta):
    if meta == 0:
        return 0
    return min(100, round((val / meta) * 100, 1))

def score_g8_rango(val, g8_min, g8_max):
    if g8_max == 0:
        return 0
    if g8_min <= val <= g8_max:
        return 100.0
    elif val < g8_min:
        return round(min(100, (val / g8_min) * 100), 1) if g8_min > 0 else 0
    else:
        exceso = val - g8_max
        return round(max(0, 100 - (exceso / g8_max) * 50), 1)

bench['S_ticket']   = bench['Ticket Prom'].apply(lambda x: score_ind(x, meta_ticket))
bench['S_g8']       = bench['Pct G8'].apply(lambda x: score_g8_rango(x, meta_g8_min, meta_g8_max))
bench['S_arts']     = bench['Arts Ticket'].apply(lambda x: score_ind(x, meta_arts))
bench['S_codigos']  = bench['Codigos'].apply(lambda x: score_ind(x, meta_codigos))
bench['S_ntickets'] = bench['Num Tickets'].apply(lambda x: score_ind(x, min_tickets))

bench['Score'] = (
    bench['S_ticket']   * W['ticket']   +
    bench['S_g8']       * W['g8']       +
    bench['S_arts']     * W['arts']     +
    bench['S_codigos']  * W['codigos']  +
    bench['S_ntickets'] * W['ntickets']
).round(1)

bench = bench.sort_values('Score', ascending=False).reset_index(drop=True)
bench.index += 1

# ── Función semáforo ──────────────────────────────────────
def semaforo(val, meta):
    if meta == 0:
        return "⚪"
    pct = val / meta * 100
    if pct >= 100:
        return "🟢"
    elif pct >= 80:
        return "🟡"
    else:
        return "🔴"

def semaforo_g8(val, g8_min, g8_max):
    if g8_min <= val <= g8_max:
        return "🟢"
    diff = min(abs(val - g8_min), abs(val - g8_max))
    if diff <= 3:
        return "🟡"
    return "🔴"

def barra_score(score):
    if score >= 95:
        color = "#10b981"
    elif score >= 80:
        color = "#f59e0b"
    else:
        color = "#ef4444"
    return (
        "<div style='display:flex;align-items:center;gap:8px'>"
        "<div style='flex:1;background:#e5e7eb;border-radius:6px;height:10px'>"
        "<div style='background:" + color + ";height:10px;border-radius:6px;width:" + str(int(score)) + "%'></div>"
        "</div>"
        "<span style='font-weight:700;color:#1a2340;min-width:40px'>" + str(score) + "</span>"
        "</div>"
    )

# ── Mostrar metas vigentes ────────────────────────────────
titulo_bench = "🎯 Metas manuales" if usar_metas_manuales else "🎯 Metas vigentes (TOP 25%)"
st.markdown(f"#### {titulo_bench}")
m1, m2, m3, m4, m5 = st.columns(5)
with m1:
    st.metric("🎫 Ticket Promedio meta", f"$ {meta_ticket:,.0f}")
with m2:
    st.metric("🏅 G8 rango",            f"{meta_g8_min:.1f}% – {meta_g8_max:.1f}%")
with m3:
    st.metric("🛒 Arts/tick meta",       f"{meta_arts:.2f}")
with m4:
    st.metric("🧾 Tickets mínimos",      f"{int(min_tickets):,}")
with m5:
    st.metric("🔢 Códigos meta",         f"{meta_codigos:.0f}")

st.markdown("<br>", unsafe_allow_html=True)

# ── Lógica imprimir benchmarking ────────────────────────
if imprimir_bench:
    _df_f = df.copy()
    _df_f["_fecha"] = pd.to_datetime(
        _df_f["Periodo"].str[:4] + "-" + _df_f["Periodo"].str[4:6] + "-" +
        _df_f["Dia"].fillna(1).astype(int).astype(str).str.zfill(2), errors="coerce"
    )
    _desde = _df_f["_fecha"].min().strftime("%d/%m/%Y")
    _hasta = _df_f["_fecha"].max().strftime("%d/%m/%Y")
    rango_fechas = _desde + " al " + _hasta
    html_informe = generar_html_benchmarking(
        bench, meta_ticket, meta_g8_min, meta_g8_max, meta_arts,
        min_tickets, meta_codigos,
        periodo_sel, farmacia_sel, rango_fechas
    )
    html_completo = (
        "<!DOCTYPE html><html><head><meta charset='utf-8'>"
        "<style>body{font-family:sans-serif;margin:0;padding:0}"
        "@media print{@page{margin:15mm}"
        "*{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important}}"
        "</style></head><body>"
        + html_informe +
        "<script>window.onload=function(){setTimeout(function(){window.print();},500);}</script>"
        "</body></html>"
    )
    st.components.v1.html(html_completo, height=900, scrolling=True)
    st.stop()

# ── Tabla semáforo ────────────────────────────────────────
# ── Resumen semáforo ─────────────────────────────────────────────
resumen_verde    = len(bench[bench['Score'] >= 95])
resumen_amarillo = len(bench[(bench['Score'] >= 80) & (bench['Score'] < 95)])
resumen_rojo     = len(bench[bench['Score'] < 80])

r1, r2, r3 = st.columns(3)
with r1:
    st.markdown(f"""<div style='background:#dcfce7;border-radius:12px;padding:20px;text-align:center'>
        <div style='font-size:36px;font-weight:800;color:#10b981'>{resumen_verde}</div>
        <div style='font-size:13px;color:#374151'>🟢 Sobre meta<br>(Score ≥ 95)</div>
    </div>""", unsafe_allow_html=True)
with r2:
    st.markdown(f"""<div style='background:#fef9c3;border-radius:12px;padding:20px;text-align:center'>
        <div style='font-size:36px;font-weight:800;color:#ca8a04'>{resumen_amarillo}</div>
        <div style='font-size:13px;color:#374151'>🟡 Cerca de meta<br>(Score 80–94)</div>
    </div>""", unsafe_allow_html=True)
with r3:
    st.markdown(f"""<div style='background:#fee2e2;border-radius:12px;padding:20px;text-align:center'>
        <div style='font-size:36px;font-weight:800;color:#dc2626'>{resumen_rojo}</div>
        <div style='font-size:13px;color:#374151'>🔴 Bajo meta<br>(Score &lt; 80)</div>
    </div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

st.markdown("#### 🚦 Semáforo por vendedor")
st.caption("🟢 Llega o supera la meta · 🟡 Entre 75–99% de la meta · 🔴 Bajo 75% de la meta")
col_sel, _ = st.columns([1,3])
with col_sel:
    n_opcion     = st.selectbox("Mostrar vendedores", [10, 20, 50, 100, "Todos"], index=4)
    n_vendedores = len(bench) if n_opcion == "Todos" else int(n_opcion)
bench_vista = bench.head(n_vendedores)

html_tabla = """
<div style='background:white;border-radius:12px;padding:16px;
     box-shadow:0 1px 4px rgba(0,0,0,0.08);overflow-x:auto'>
<table style='width:100%;border-collapse:collapse;font-family:sans-serif'>
<thead>
  <tr style='border-bottom:2px solid #e5e7eb'>
    <th style='padding:10px 8px;text-align:left;font-size:12px;color:#6b7280'>#</th>
    <th style='padding:10px 8px;text-align:left;font-size:12px;color:#6b7280'>VENDEDOR</th>
    <th style='padding:10px 8px;text-align:center;font-size:12px;color:#6b7280'>🎫 TICKET PROMEDIO</th>
    <th style='padding:10px 8px;text-align:center;font-size:12px;color:#6b7280'>🏅 G8%</th>
    <th style='padding:10px 8px;text-align:center;font-size:12px;color:#6b7280'>🛒 ARTS/TICK</th>
    <th style='padding:10px 8px;text-align:center;font-size:12px;color:#6b7280'>🧾 TICKETS</th>
    <th style='padding:10px 8px;text-align:center;font-size:12px;color:#6b7280'>🔢 CÓDIGOS</th>
    <th style='padding:10px 8px;text-align:left;font-size:12px;color:#6b7280;min-width:160px'>SCORE</th>
  </tr>
</thead>
<tbody>
"""

for i, row in bench_vista.iterrows():
    bg = "#f9fafb" if i % 2 == 0 else "white"
    s_t  = semaforo(row['Ticket Prom'],  meta_ticket)
    s_p  = semaforo_g8(row['Pct G8'],    meta_g8_min, meta_g8_max)
    s_a  = semaforo(row['Arts Ticket'],  meta_arts)
    s_n  = semaforo(row['Num Tickets'],  min_tickets)
    s_c  = semaforo(row['Codigos'],      meta_codigos)
    barra = barra_score(row['Score'])

    html_tabla += (
        "<tr style='background:" + bg + ";border-bottom:1px solid #f3f4f6'>"
        "<td style='padding:10px 8px;font-size:12px;color:#9ca3af'>" + str(i) + "</td>"
        "<td style='padding:10px 8px;font-size:13px;font-weight:600;color:#1a2340'>" + str(row['Nombre Vendedor']) + "</td>"
        "<td style='padding:10px 8px;text-align:center;font-size:13px'>"
        + s_t + " <span style='font-size:11px;color:#374151'>$ " + f"{row['Ticket Prom']:,.0f}" + "</span></td>"
        "<td style='padding:10px 8px;text-align:center;font-size:13px'>"
        + s_p + " <span style='font-size:11px;color:#374151'>" + f"{row['Pct G8']:.1f}" + "%</span></td>"
        "<td style='padding:10px 8px;text-align:center;font-size:13px'>"
        + s_a + " <span style='font-size:11px;color:#374151'>" + f"{row['Arts Ticket']:.2f}" + "</span></td>"
        "<td style='padding:10px 8px;text-align:center;font-size:13px'>"
        + s_n + " <span style='font-size:11px;color:#374151'>" + f"{int(row['Num Tickets']):,}" + "</span></td>"
        "<td style='padding:10px 8px;text-align:center;font-size:13px'>"
        + s_c + " <span style='font-size:11px;color:#374151'>" + str(int(row['Codigos'])) + "</span></td>"
        "<td style='padding:10px 8px'>" + barra + "</td>"
        "</tr>"
    )

html_tabla += "</tbody></table></div>"
st.markdown(html_tabla, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── Top 10 Score ──────────────────────────────────────────
st.markdown("#### 🏅 Ranking Score global · Top 10")
top_score = bench.head(10).copy()
fig_score = px.bar(
    top_score.sort_values('Score', ascending=True),
    x='Score', y='Nombre Vendedor',
    orientation='h',
    title='Score global (0-100) · Top 10 vendedores',
    color='Score',
    color_continuous_scale=['#ef4444','#f59e0b','#10b981'],
    range_color=[0,100],
    text='Score'
)
fig_score = card_chart(fig_score)
fig_score.update_traces(textposition='outside')
fig_score.update_layout(
    xaxis_title='Score', yaxis_title='',
    showlegend=False, coloraxis_showscale=False,
    xaxis_range=[0,110]
)
st.plotly_chart(fig_score, use_container_width=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── PRODUCTOS ─────────────────────────────────────────────
st.markdown("---")
st.markdown("### 📦 Productos más vendidos · Top 15")
top_prod = df.groupby('Producto').agg(
    Unidades=('Piezas Acumuladas','sum'),
    Ventas  =('Importe Acumulado','sum'),
    Familia =('Familia','first')
).nlargest(15,'Ventas').reset_index()
top_prod['Participación'] = (top_prod['Ventas']/top_prod['Ventas'].sum()*100).round(2)
top_prod['Ventas Fmt']    = top_prod['Ventas'].apply(lambda x: f"$ {x:,.0f}")
top_prod['Part. Fmt']     = top_prod['Participación'].apply(lambda x: f"{x:.2f}%")
st.dataframe(
    top_prod[['Producto','Familia','Unidades','Ventas Fmt','Part. Fmt']].rename(columns={
        'Ventas Fmt':'Ventas','Part. Fmt':'Participación'
    }),
    use_container_width=True, hide_index=True
)

st.markdown("---")
st.caption("💊 Grupo Baco · Dashboard de Ventas · Farmacias Dr. Simi")

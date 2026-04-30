import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta

AZUL       = "#3b82f6"
VERDE      = "#10b981"
NARANJA    = "#f59e0b"
ROJO       = "#ef4444"
AZUL_LISTA = ["#3b82f6","#60a5fa","#93c5fd","#1d4ed8","#2563eb","#6366f1","#8b5cf6","#06b6d4"]

@st.cache_data
def leer_facturacion(contenido):
    df = pd.read_excel(contenido, sheet_name='REGISTRO', engine='openpyxl')
    df['Fecha Documento']   = pd.to_datetime(df['Fecha Documento'],   errors='coerce')
    df['Fecha Vencimiento'] = pd.to_datetime(df['Fecha Vencimiento'], errors='coerce')
    df['Monto']  = pd.to_numeric(df['Monto'],  errors='coerce').fillna(0)
    df['Cargo']  = pd.to_numeric(df['Cargo'],  errors='coerce').fillna(0)
    df['Abono']  = pd.to_numeric(df['Abono'],  errors='coerce').fillna(0)
    df['Semana'] = pd.to_numeric(df['Semana'], errors='coerce')
    df['Año']    = pd.to_numeric(df['Año'],    errors='coerce')
    return df

@st.cache_data
def leer_ventas(nombre, contenido):
    try:
        df = pd.read_excel(contenido, engine='openpyxl')
    except Exception:
        return pd.DataFrame()
    if 'Farmacia' not in df.columns:
        return pd.DataFrame()
    df = df.dropna(subset=['Farmacia'])
    if 'Importe Acumulado' not in df.columns:
        return pd.DataFrame()
    df = df[df['Importe Acumulado'].apply(
        lambda x: str(x).replace('.','').replace('-','').strip().isdigit()
        if pd.notna(x) else False
    )]
    df['Importe Acumulado'] = pd.to_numeric(df['Importe Acumulado'], errors='coerce')
    df['Periodo'] = df['Periodo'].astype(str).str[:6]
    df['Dia']     = pd.to_numeric(df['Dia'], errors='coerce')
    df['Semana_calc'] = ((df['Dia'] - 1) // 7 + 1).clip(1, 5)
    df['Nombre Farmacia'] = df['Farmacia'].str.split('-').str[1:].str.join('-').str.strip()
    df['Local'] = df['Farmacia'].str.split('-').str[0].str.strip()
    return df

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

def render_compras():
    st.markdown("## 📦 Dashboard de Compras")
    st.caption("Facturación · Pagos · Relación Compra/Venta · Grupo Baco · Dr. Simi")

    # ── SIDEBAR COMPRAS ───────────────────────────────────
    with st.sidebar:
        st.markdown("**📂 DATOS COMPRAS**")

        archivo_fact = st.file_uploader(
            "Facturación (.xlsx)",
            type=["xlsx"],
            key="uploader_fact",
            help="Sube el archivo FACTURACION_SIMI.xlsx"
        )
        # Guardar en session_state para no perder al navegar
        if archivo_fact:
            st.session_state["fact_file"] = archivo_fact
        elif "fact_file" in st.session_state:
            archivo_fact = st.session_state["fact_file"]

        # Ventas se toman automáticamente desde la página de Ventas
        archivo_vtas = st.session_state.get("vtas_file", None)
        if archivo_vtas:
            st.caption("✅ Ventas cargadas desde página Ventas")
        else:
            st.caption("⚠️ Carga primero los archivos en 📈 Ventas para ver ratio Compra/Venta")

    if not archivo_fact:
        st.info("👈 Carga el archivo de facturación desde el panel izquierdo para comenzar.")
        return

    df = leer_facturacion(archivo_fact)

    # ── FILTROS SIDEBAR ───────────────────────────────────
    with st.sidebar:
        años_op = sorted(df['Año'].dropna().unique().astype(int).tolist())
        año_sel = st.multiselect("AÑO", años_op, default=años_op, key="año_fact")

        meses_orden = ['enero','febrero','marzo','abril','mayo','junio',
                       'julio','agosto','septiembre','octubre','noviembre','diciembre']
        meses_op = [m for m in meses_orden if m in df['Mes'].unique()]
        mes_sel  = st.multiselect("MES", meses_op, default=meses_op, key="mes_fact")

        locales_op = sorted(df['Local'].dropna().unique().tolist())
        local_sel  = st.multiselect("LOCAL", locales_op, default=locales_op, key="local_fact")

        cats_op = sorted(df['Categoría'].dropna().unique().tolist())
        cat_sel = st.multiselect("CATEGORÍA", cats_op, default=cats_op, key="cat_fact")

    # Aplicar filtros
    df = df[df['Año'].isin(año_sel)]
    df = df[df['Mes'].isin(mes_sel)]
    df = df[df['Local'].isin(local_sel)]
    df = df[df['Categoría'].isin(cat_sel)]

    if df.empty:
        st.warning("⚠️ No hay datos con los filtros seleccionados.")
        return

    hoy = pd.Timestamp(datetime.today().date())

    # ── KPIs ─────────────────────────────────────────────
    df_fact  = df[df['Tipo De Movimiento'] == 'Factura']
    df_nc    = df[df['Tipo De Movimiento'] == 'Nota de Crédito']

    total_fact   = df_fact['Monto'].sum()
    total_nc     = df_nc['Monto'].sum()
    neto         = total_fact - total_nc

    vencidas     = df_fact[df_fact['Estatus'] == 'Vencida']['Monto'].sum()
    por_vencer   = df_fact[df_fact['Estatus'] == 'Por vencer']['Monto'].sum()
    prox_49      = df_fact[
        (df_fact['Fecha Vencimiento'] >= hoy) &
        (df_fact['Fecha Vencimiento'] <= hoy + timedelta(days=49))
    ]['Monto'].sum()

    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.metric("💰 Total facturado",      f"$ {total_fact:,.0f}")
    with k2:
        st.metric("📋 Notas de crédito",     f"$ {total_nc:,.0f}")
    with k3:
        st.metric("🔢 Neto (fact - NC)",     f"$ {neto:,.0f}")
    with k4:
        st.metric("🏪 Locales activos",      f"{df['Local'].nunique()}")

    st.markdown("<br>", unsafe_allow_html=True)

    p1, p2, p3 = st.columns(3)
    with p1:
        st.metric("⚠️ Vencido",             f"$ {vencidas:,.0f}",
                  delta="Requiere atención" if vencidas > 0 else "Sin vencidos",
                  delta_color="inverse" if vencidas > 0 else "normal")
    with p2:
        st.metric("🕐 Por vencer",           f"$ {por_vencer:,.0f}")
    with p3:
        st.metric("📅 Próx. 49 días",        f"$ {prox_49:,.0f}")

    st.markdown("---")

    # ── TABLA PRÓXIMOS VENCIMIENTOS ───────────────────────
    st.markdown("### 📅 Próximos vencimientos — 49 días")
    df_prox = df_fact[
        (df_fact['Fecha Vencimiento'] >= hoy) &
        (df_fact['Fecha Vencimiento'] <= hoy + timedelta(days=49))
    ].copy()

    if not df_prox.empty:
        df_prox['Días restantes'] = (df_prox['Fecha Vencimiento'] - hoy).dt.days

        # Calcular viernes siguiente al vencimiento (día de pago real)
        def viernes_siguiente(fecha):
            if pd.isna(fecha):
                return None
            dias_hasta_viernes = (4 - fecha.weekday()) % 7
            if dias_hasta_viernes == 0:
                dias_hasta_viernes = 7  # Si vence viernes, paga el viernes siguiente
            return fecha + timedelta(days=dias_hasta_viernes)

        df_prox['Fecha Pago'] = df_prox['Fecha Vencimiento'].apply(viernes_siguiente)
        df_prox['Semana Pago'] = df_prox['Fecha Pago'].apply(
            lambda x: x.strftime("Viernes %d/%m/%Y") if pd.notna(x) else "—"
        )

        # Tabla pivoteada: filas=viernes de pago, columnas=empresa, total
        pivot = df_prox.groupby(['Semana Pago','Fecha Pago','Empresa'])['Monto'].sum().reset_index()
        pivot_tabla = pivot.pivot_table(
            index=['Fecha Pago','Semana Pago'],
            columns='Empresa',
            values='Monto',
            aggfunc='sum'
        ).fillna(0)
        pivot_tabla['💰 TOTAL'] = pivot_tabla.sum(axis=1)
        pivot_tabla = pivot_tabla.sort_index(level='Fecha Pago')

        # Fila TOTAL GENERAL
        total_row = pivot_tabla.sum()
        total_row.name = (pd.Timestamp('2099-01-01'), '📊 TOTAL GENERAL')
        pivot_tabla = pd.concat([pivot_tabla, total_row.to_frame().T])

        # Mostrar solo la etiqueta de semana como índice
        pivot_tabla.index = [idx[1] for idx in pivot_tabla.index]

        # Formatear como $
        pivot_fmt = pivot_tabla.map(lambda x: f"$ {x:,.0f}" if isinstance(x, (int,float)) and x > 0 else "—")
        st.dataframe(pivot_fmt, use_container_width=True)

        # Gráfico usando pivot sin formatear
        pivot_grafico = pivot.copy()
        orden_sem = ['Esta semana','Semana 2','Semana 3','Semana 4',
                     'Semana 5','Semana 6','Semana 7']
        # Gráfico por viernes de pago — ordenado por fecha real
        pivot_grafico = df_prox.groupby(['Fecha Pago','Semana Pago','Empresa'])['Monto'].sum().reset_index()
        pivot_grafico = pivot_grafico.sort_values('Fecha Pago')

        fig_vto = px.bar(
            pivot_grafico, x='Semana Pago', y='Monto', color='Empresa',
            title='Pagos por viernes — próximos 49 días',
            color_discrete_sequence=AZUL_LISTA,
            category_orders={'Semana Pago': pivot_grafico['Semana Pago'].tolist()}
        )
        fig_vto = card_chart(fig_vto)
        fig_vto.update_layout(xaxis_title='Viernes de pago', yaxis_title='$ Monto',
                              xaxis_tickangle=-20)
        st.plotly_chart(fig_vto, use_container_width=True)
    else:
        st.success("✅ No hay vencimientos en los próximos 49 días.")

    st.markdown("---")

    # ── RATIO COMPRA / VENTA ──────────────────────────────
    st.markdown("### 📊 Relación Compra / Venta semanal")
    st.caption("Solo Mercadería · Solo Facturas · Sin Notas de Crédito")

    if archivo_vtas:
        # Puede ser un archivo o una lista
        if isinstance(archivo_vtas, list):
            archivo_vtas = archivo_vtas[0]
        df_v = leer_ventas(archivo_vtas.name, archivo_vtas)
        if df_v.empty:
            archivo_vtas = None

        # Compras: solo Mercadería + solo Facturas
        df_merc = df[
            (df['Categoría'] == 'Mercadería') &
            (df['Tipo De Movimiento'] == 'Factura')
        ].copy()

        # Agrupar compras por semana del año
        compras_sem = df_merc.groupby('Semana')['Monto'].sum().reset_index()
        compras_sem.columns = ['Semana','Compras']

        # Ventas: calcular semana si no existe
        if 'Semana_calc' not in df_v.columns:
            if 'Dia' in df_v.columns:
                df_v['Semana_calc'] = ((pd.to_numeric(df_v['Dia'], errors='coerce') - 1) // 7 + 1).clip(1,5)
            else:
                st.warning("⚠️ No se puede calcular semana desde el archivo de ventas.")
                archivo_vtas = None

        if archivo_vtas is not None:
            ventas_sem = df_v.groupby('Semana_calc')['Importe Acumulado'].sum().reset_index()
            ventas_sem.columns = ['Semana','Ventas']

            # Merge por semana
                ratio_df = compras_sem.merge(ventas_sem, on='Semana', how='inner')
        ratio_df = ratio_df[ratio_df['Ventas'] > 0]
        ratio_df['Ratio %'] = (ratio_df['Compras'] / ratio_df['Ventas'] * 100).round(1)
        ratio_df['Semana Label'] = 'S' + ratio_df['Semana'].astype(str)

        def color_ratio(r):
            if r > 100: return ROJO
            if r > 75:  return NARANJA
            if r >= 60: return VERDE
            return AZUL

        ratio_df['Color'] = ratio_df['Ratio %'].apply(color_ratio)

        # KPIs ratio
        r1, r2, r3, r4 = st.columns(4)
        ratio_prom = ratio_df['Ratio %'].mean()
        with r1:
            st.metric("📊 Ratio promedio",  f"{ratio_prom:.1f}%")
        with r2:
            st.metric("🎯 Meta ideal",      "60% – 75%")
        with r3:
            semanas_ok = len(ratio_df[ratio_df['Ratio %'].between(60,75)])
            st.metric("✅ Semanas en meta", f"{semanas_ok} de {len(ratio_df)}")
        with r4:
            semanas_riesgo = len(ratio_df[ratio_df['Ratio %'] > 100])
            st.metric("🔴 Semanas > 100%", f"{semanas_riesgo}")

        st.markdown("<br>", unsafe_allow_html=True)

        # Gráfico ratio
        fig_ratio = px.bar(
            ratio_df, x='Semana Label', y='Ratio %',
            title='Ratio Compra/Venta % por semana (Mercadería)',
            color='Color',
            color_discrete_map={VERDE:VERDE, NARANJA:NARANJA, ROJO:ROJO, AZUL:AZUL}
        )
        fig_ratio = card_chart(fig_ratio)
        fig_ratio.add_hline(y=60,  line_dash="dot", line_color=VERDE,  annotation_text="Meta mín 60%")
        fig_ratio.add_hline(y=75,  line_dash="dot", line_color=NARANJA, annotation_text="Meta máx 75%")
        fig_ratio.add_hline(y=100, line_dash="dot", line_color=ROJO,   annotation_text="Límite 100%")
        fig_ratio.update_layout(xaxis_title='Semana', yaxis_title='Ratio %', showlegend=False)
        st.plotly_chart(fig_ratio, use_container_width=True)

        # Gráfico barras lado a lado
        ratio_melt = ratio_df.melt(
            id_vars='Semana Label',
            value_vars=['Compras','Ventas'],
            var_name='Tipo', value_name='Monto'
        )
        fig_cv = px.bar(
            ratio_melt, x='Semana Label', y='Monto', color='Tipo',
            barmode='group',
            title='Compras vs Ventas por semana (Mercadería)',
            color_discrete_map={'Compras': ROJO, 'Ventas': AZUL}
        )
        fig_cv = card_chart(fig_cv)
        fig_cv.update_layout(xaxis_title='Semana', yaxis_title='$ Monto')
        st.plotly_chart(fig_cv, use_container_width=True)

    else:
        st.info("👈 Carga el archivo de ventas para ver la relación Compra/Venta.")

    st.markdown("---")

    # ── EVOLUCIÓN MENSUAL ─────────────────────────────────
    st.markdown("### 📈 Evolución mensual de facturación")

    meses_num = {'enero':1,'febrero':2,'marzo':3,'abril':4,'mayo':5,'junio':6,
                 'julio':7,'agosto':8,'septiembre':9,'octubre':10,'noviembre':11,'diciembre':12}
    df['Mes Num'] = df['Mes'].map(meses_num)

    evol = df[df['Tipo De Movimiento']=='Factura'].groupby(
        ['Año','Mes','Mes Num']
    )['Monto'].sum().reset_index().sort_values(['Año','Mes Num'])
    evol['Label'] = evol['Mes'].str.capitalize() + ' ' + evol['Año'].astype(str)

    fig_evol = px.bar(evol, x='Label', y='Monto',
                      title='Facturación mensual',
                      color_discrete_sequence=[AZUL], text_auto='.2s')
    fig_evol = card_chart(fig_evol)
    fig_evol.update_layout(xaxis_title='', yaxis_title='$ Monto', showlegend=False)
    st.plotly_chart(fig_evol, use_container_width=True)

    # ── FACTURACIÓN POR LOCAL ─────────────────────────────
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("### 🏪 Facturación por local")
        por_local = df[df['Tipo De Movimiento']=='Factura'].groupby(
            'Local')['Monto'].sum().sort_values(ascending=True).reset_index()
        fig_local = px.bar(por_local, x='Monto', y='Local',
                           orientation='h', title='Facturación por local',
                           color_discrete_sequence=[AZUL])
        fig_local = card_chart(fig_local)
        fig_local.update_layout(xaxis_title='$ Monto', yaxis_title='', showlegend=False)
        st.plotly_chart(fig_local, use_container_width=True)

    with col2:
        st.markdown("### 🗂️ Por categoría")
        por_cat = df[df['Tipo De Movimiento']=='Factura'].groupby(
            'Categoría')['Monto'].sum().sort_values(ascending=False).reset_index()
        fig_cat = px.pie(por_cat, names='Categoría', values='Monto',
                         title='Distribución por categoría',
                         hole=0.5, color_discrete_sequence=AZUL_LISTA)
        fig_cat = card_chart(fig_cat)
        fig_cat.update_layout(legend=dict(font=dict(size=10)))
        st.plotly_chart(fig_cat, use_container_width=True)

    # ── SUBCATEGORÍAS ─────────────────────────────────────
    st.markdown("### 📋 Top subcategorías")
    por_sub = df[df['Tipo De Movimiento']=='Factura'].groupby(
        'Subcategoria')['Monto'].sum().nlargest(15).sort_values(ascending=True).reset_index()
    fig_sub = px.bar(por_sub, x='Monto', y='Subcategoria',
                     orientation='h', title='Top 15 subcategorías',
                     color_discrete_sequence=[VERDE])
    fig_sub = card_chart(fig_sub)
    fig_sub.update_layout(xaxis_title='$ Monto', yaxis_title='', showlegend=False)
    st.plotly_chart(fig_sub, use_container_width=True)

    # ── ESTADO DE PAGOS ───────────────────────────────────
    st.markdown("### 💳 Estado de pagos")
    col3, col4 = st.columns(2)
    with col3:
        estado = df[df['Tipo De Movimiento']=='Factura'].groupby(
            'Estatus')['Monto'].sum().reset_index()
        colores_estado = {
            'Pagada':'#10b981','Por vencer':'#f59e0b',
            'Vencida':'#ef4444','No vencido':'#3b82f6'
        }
        fig_est = px.pie(estado, names='Estatus', values='Monto',
                         title='Estado de pagos',
                         color='Estatus',
                         color_discrete_map=colores_estado,
                         hole=0.5)
        fig_est = card_chart(fig_est)
        st.plotly_chart(fig_est, use_container_width=True)

    with col4:
        estado_local = df[df['Tipo De Movimiento']=='Factura'].groupby(
            ['Local','Estatus'])['Monto'].sum().reset_index()
        fig_el = px.bar(estado_local, x='Local', y='Monto', color='Estatus',
                        title='Estado de pagos por local',
                        color_discrete_map=colores_estado)
        fig_el = card_chart(fig_el)
        fig_el.update_layout(xaxis_title='', yaxis_title='$ Monto')
        st.plotly_chart(fig_el, use_container_width=True)

    st.markdown("---")
    st.caption("💊 Grupo Baco · Dashboard de Compras · Farmacias Dr. Simi")

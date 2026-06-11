import streamlit as st

def render_kpis(resumen_df, df_disponibles_full, df_vencimientos):
    """
    Muestra los KPIs ejecutivos:
    KPI 1: Clientes con riesgo (Crítico o Alto)
    KPI 2: Tickets disponibles totales
    KPI 3: Tickets próximos a vencer
    KPI 4: Monto potencial de venta (Tickets a reponer)
    KPI 5: Clientes críticos
    """
    st.markdown("### Dashboard Ejecutivo")
    
    if resumen_df.empty:
        st.warning("No hay datos para mostrar con los filtros actuales.")
        return

    # KPI 1: Clientes en riesgo (Crítico + Alto) - Conteo de clientes únicos en riesgo
    clientes_riesgo = resumen_df[resumen_df["Riesgo"].isin(["Crítico", "Alto"])]["RazonSocial"].nunique()
    
    # KPI 2: Tickets disponibles totales (usamos df_disponibles_full filtrado, o la suma de Saldo_Disponible de resumen)
    tickets_disponibles = int(resumen_df["Saldo_Disponible"].sum())
    
    # KPI 3: Tickets próximos a vencer
    tickets_vencer = len(df_vencimientos) if not df_vencimientos.empty else 0
    
    # KPI 4: Monto potencial de venta (Estimación de tickets a reponer)
    # Ejemplo de estimación: Para clientes críticos/altos, se les debería vender al menos el consumo de 4 semanas (promedio semanal * 4)
    resumen_critico = resumen_df[resumen_df["Riesgo"].isin(["Crítico", "Alto"])]
    tickets_reponer = int(resumen_critico["Promedio_Semanal"].sum() * 4)
    
    # KPI 5: Clientes Críticos
    clientes_criticos = resumen_df[resumen_df["Riesgo"] == "Crítico"]["RazonSocial"].nunique()

    # Estilos CSS de las tarjetas usando métricas nativas de Streamlit
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        st.metric(label="Clientes en Riesgo", value=clientes_riesgo)
    with col2:
        st.metric(label="Tickets Disponibles", value=f"{tickets_disponibles:,}")
    with col3:
        st.metric(label="Próximos a Vencer (30d)", value=f"{tickets_vencer:,}")
    with col4:
        st.metric(label="Potencial Reposición", value=f"{tickets_reponer:,} tkts")
    with col5:
        st.metric(label="Clientes Críticos (< 1 sem)", value=clientes_criticos, delta="- Atención", delta_color="inverse")
    
    st.markdown("---")

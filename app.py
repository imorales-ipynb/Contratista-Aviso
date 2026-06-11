import streamlit as st
import pandas as pd
from database import cargar_datos, recargar_base_datos
from data_processing import procesar_datos
from components.kpis import render_kpis
from components.charts import render_charts
from components.tables import render_tables
import datetime

st.set_page_config(
    page_title="Gestión Contratistas",
    page_icon="🍔",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    [data-testid="stMetric"] {
        background-color: #f8f9fa;
        padding: 15px;
        border-radius: 10px;
        box-shadow: 2px 2px 10px rgba(0,0,0,0.05);
        text-align: center;
    }
    [data-testid="stMetric"]:hover {
        transform: translateY(-2px);
        transition: 0.3s;
        box-shadow: 2px 5px 15px rgba(0,0,0,0.1);
    }
    [data-testid="stMetricLabel"] p,
    [data-testid="stMetricLabel"] div {
        color: #555555 !important;
        font-size: 0.85rem !important;
    }
    [data-testid="stMetricValue"] div {
        color: #0e1117 !important;
        font-size: 1.6rem !important;
        font-weight: 700 !important;
    }
    [data-testid="stMetricDelta"] div {
        color: #555555 !important;
    }
</style>
""", unsafe_allow_html=True)

st.title("Gestión Contratistas")

col1, col2 = st.columns([4, 1])
with col2:
    if st.button("🔄 Actualizar Base de Datos"):
        with st.spinner("Descargando datos desde SQL Server..."):
            recargar_base_datos(use_mock=False)
            st.rerun()

# 1. Cargar datos
df_consumo, df_jerarquia, df_servicios = cargar_datos()

if df_consumo.empty:
    st.error("No se encontraron datos. Intenta actualizar la base de datos.")
    st.stop()

# 2. Procesar
resumen_df, df_vencimientos, df_full = procesar_datos(df_consumo, df_jerarquia, df_servicios)

if resumen_df.empty:
    st.warning("El procesamiento de datos no generó resultados.")
    st.stop()

# 3. Renderizar componentes (sin filtros globales)
render_kpis(resumen_df, df_full, df_vencimientos)
render_charts(resumen_df, df_full, df_vencimientos)
render_tables(resumen_df)

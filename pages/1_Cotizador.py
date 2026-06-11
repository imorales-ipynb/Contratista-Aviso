import streamlit as st
from database import cargar_precios, cargar_clientes
from components.quoter import render_cotizador

st.set_page_config(
    page_title="Cotizador de Servicios",
    page_icon="📋",
    layout="wide"
)

df_precios  = cargar_precios()
df_clientes = cargar_clientes()

render_cotizador(df_precios, df_clientes)

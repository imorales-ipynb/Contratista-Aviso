import streamlit as st
import pandas as pd
from datetime import date

def render_sidebar(df_full):
    """Renderiza la barra lateral y devuelve el diccionario de filtros seleccionados."""
    st.sidebar.header("Filtros Globales")
    
    # Extraer valores únicos de las columnas para los filtros
    if df_full.empty:
        return {}
        
    areas = ["Todas"] + sorted(df_full["Area"].dropna().unique().tolist())
    gops = ["Todos"] + sorted(df_full["GOP"].dropna().unique().tolist())
    jops = ["Todos"] + sorted(df_full["JOP"].dropna().unique().tolist())
    casinos = ["Todos"] + sorted(df_full["Casino"].dropna().unique().tolist())
    ccs = ["Todos"] + sorted(df_full["CC"].dropna().unique().tolist())
    clientes = ["Todos"] + sorted(df_full["RazonSocial"].dropna().unique().tolist())
    servicios = ["Todos"] + sorted(df_full["NombreServicio"].dropna().unique().tolist())
    grupos_servicio = ["Todos"] + sorted(df_full["Grupo"].dropna().unique().tolist())
    tipos_servicio = ["Todos"] + sorted(df_full["Tipo"].dropna().unique().tolist())

    filtros = {}
    
    filtros["Area"] = st.sidebar.selectbox("Área", areas)
    filtros["GOP"] = st.sidebar.selectbox("GOP", gops)
    filtros["JOP"] = st.sidebar.selectbox("JOP", jops)
    filtros["Casino"] = st.sidebar.selectbox("Casino", casinos)
    filtros["CC"] = st.sidebar.selectbox("Centro de Costo", ccs)
    filtros["RazonSocial"] = st.sidebar.selectbox("Cliente", clientes)
    filtros["Grupo"] = st.sidebar.selectbox("Grupo Servicio", grupos_servicio)
    filtros["Tipo"] = st.sidebar.selectbox("Tipo Servicio", tipos_servicio)
    filtros["NombreServicio"] = st.sidebar.selectbox("Servicio", servicios)
    
    return filtros

def aplicar_filtros(df, filtros):
    """Aplica el diccionario de filtros a un DataFrame dado."""
    if df.empty:
        return df
        
    df_filtrado = df.copy()
    for col, val in filtros.items():
        if val not in ["Todos", "Todas"]:
            if col in df_filtrado.columns:
                df_filtrado = df_filtrado[df_filtrado[col] == val]
    return df_filtrado

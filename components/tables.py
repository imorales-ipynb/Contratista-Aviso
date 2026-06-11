import streamlit as st
import pandas as pd
from io import BytesIO

_fragment = getattr(st, "fragment", getattr(st, "experimental_fragment", lambda f: f))


def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Gestión Comercial')
    return output.getvalue()


@_fragment
def render_tables(resumen_df):
    if resumen_df.empty:
        return

    st.markdown("### Tablas de Detalle")

    # ── Filtros ────────────────────────────────────────────────────────────────
    col_f1, col_f2, col_f3, col_f4 = st.columns(4)

    with col_f1:
        gops = ["Todos"]
        if "GOP" in resumen_df.columns:
            gops += sorted(resumen_df["GOP"].dropna().unique().tolist())
        gop_sel = st.selectbox("Gerencia:", options=gops, key="tbl_gop")

    df_f = resumen_df.copy()
    if gop_sel != "Todos" and "GOP" in df_f.columns:
        df_f = df_f[df_f["GOP"] == gop_sel]

    with col_f2:
        casinos = ["Todos"] + sorted(df_f["Casino"].dropna().unique().tolist())
        if st.session_state.get("tbl_casino", "Todos") not in casinos:
            st.session_state["tbl_casino"] = "Todos"
        casino_sel = st.selectbox("Casino:", options=casinos, key="tbl_casino")

    if casino_sel != "Todos":
        df_f = df_f[df_f["Casino"] == casino_sel]

    with col_f3:
        servicios = ["Todos"]
        if "NombreServicio" in df_f.columns:
            servicios += sorted(df_f["NombreServicio"].dropna().unique().tolist())
        if st.session_state.get("tbl_servicio", "Todos") not in servicios:
            st.session_state["tbl_servicio"] = "Todos"
        servicio_sel = st.selectbox("Servicio:", options=servicios, key="tbl_servicio")

    if servicio_sel != "Todos" and "NombreServicio" in df_f.columns:
        df_f = df_f[df_f["NombreServicio"] == servicio_sel]

    with col_f4:
        orden_riesgo = ["Crítico", "Alto", "Medio", "Bajo"]
        riesgos = ["Todos"] + [r for r in orden_riesgo if "Riesgo" in df_f.columns and r in df_f["Riesgo"].values]
        riesgo_sel = st.selectbox("Riesgo:", options=riesgos, key="tbl_riesgo")

    if riesgo_sel != "Todos":
        df_f = df_f[df_f["Riesgo"] == riesgo_sel]

    # ── Preparar columnas ──────────────────────────────────────────────────────
    cols_mostrar = [
        "RazonSocial", "RutContratista", "CC", "Casino", "NombreServicio",
        "Saldo_Disponible", "Promedio_Semanal", "Semanas_Cubiertas",
        "Riesgo", "Tendencia", "GOP", "JOP", "Recomendacion_Comercial"
    ]
    cols_existentes = [c for c in cols_mostrar if c in df_f.columns]
    df_mostrar = df_f[cols_existentes].copy()

    if "Promedio_Semanal" in df_mostrar.columns:
        df_mostrar["Promedio_Semanal"] = df_mostrar["Promedio_Semanal"].round(0).astype(int)
    if "Semanas_Cubiertas" in df_mostrar.columns:
        df_mostrar["Semanas_Cubiertas"] = df_mostrar["Semanas_Cubiertas"].clip(upper=998).round(0).astype(int)

    # ── Tabs ──────────────────────────────────────────────────────────────────
    tab1, tab2 = st.tabs(["Ranking Clientes Críticos", "Gestión Comercial (Completo)"])

    with tab1:
        st.markdown("#### Top 20 Clientes Críticos")
        df_criticos = (
            df_mostrar[df_mostrar["Riesgo"] == "Crítico"]
            .sort_values(by="Semanas_Cubiertas", ascending=True)
            .head(20)
        )
        if not df_criticos.empty:
            st.dataframe(df_criticos, use_container_width=True)
        else:
            st.success("No se encontraron clientes críticos en esta selección.")

    with tab2:
        st.markdown("#### Detalle Completo")
        st.dataframe(df_mostrar, use_container_width=True)

        excel_data = to_excel(df_mostrar)
        st.download_button(
            label="📥 Descargar Excel",
            data=excel_data,
            file_name='gestion_comercial_contratistas.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

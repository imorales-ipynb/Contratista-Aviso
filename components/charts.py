import streamlit as st
import plotly.express as px
import pandas as pd

_fragment = getattr(st, "fragment", getattr(st, "experimental_fragment", lambda f: f))


@_fragment
def render_charts(resumen_df, df_full, df_vencimientos):
    if resumen_df.empty:
        return

    st.markdown("### Visualizaciones")
    tab1, tab2 = st.tabs(["Consumo por Casino", "Vencimientos"])

    # 1. Consumo por Casino
    with tab1:
        st.markdown("#### Consumo Semanal por Casino")

        df_usados = df_full[df_full["Usado"] == "SI"].copy()
        if not df_usados.empty:
            df_usados['FechaUso'] = pd.to_datetime(df_usados['FechaUso'], errors='coerce')
            df_usados = df_usados.dropna(subset=['FechaUso'])
            df_usados['SemanaFecha'] = (
                df_usados['FechaUso'] - pd.to_timedelta(df_usados['FechaUso'].dt.dayofweek, unit='D')
            ).dt.normalize()

            fecha_min = df_usados['FechaUso'].min().date()
            fecha_max = df_usados['FechaUso'].max().date()

            col_f1, col_f2, col_f3, col_f4 = st.columns(4)

            with col_f1:
                fecha_desde = st.date_input("Desde:", value=fecha_min, min_value=fecha_min, max_value=fecha_max, key="chart_desde")
            with col_f2:
                fecha_hasta = st.date_input("Hasta:", value=fecha_max, min_value=fecha_min, max_value=fecha_max, key="chart_hasta")

            df_filtrado = df_usados[
                (df_usados['FechaUso'].dt.date >= fecha_desde) &
                (df_usados['FechaUso'].dt.date <= fecha_hasta)
            ]

            with col_f3:
                gops = ["Todos"]
                if "GOP" in df_filtrado.columns:
                    gops += sorted(df_filtrado["GOP"].dropna().unique().tolist())
                gop_sel = st.selectbox("Gerencia:", options=gops, key="chart_gop")

            df_filtrado_gop = df_filtrado if gop_sel == "Todos" else df_filtrado[df_filtrado["GOP"] == gop_sel]

            with col_f4:
                casinos = ["Todos"] + sorted(df_filtrado_gop["Casino"].dropna().unique().tolist())
                # Resetear casino si ya no está en la lista tras cambio de GOP
                if st.session_state.get("chart_casino", "Todos") not in casinos:
                    st.session_state["chart_casino"] = "Todos"
                casino_sel = st.selectbox("Casino:", options=casinos, key="chart_casino")

            df_casino = df_filtrado_gop if casino_sel == "Todos" else df_filtrado_gop[df_filtrado_gop["Casino"] == casino_sel]

            if not df_casino.empty:
                df_grouped = (
                    df_casino
                    .groupby(["SemanaFecha", "Casino"])
                    .size()
                    .reset_index(name="Tickets")
                    .sort_values("SemanaFecha")
                )

                fig = px.bar(
                    df_grouped, x="SemanaFecha", y="Tickets", color="Casino",
                    barmode="stack" if casino_sel == "Todos" else "relative"
                )
                fig.update_traces(hovertemplate="%{x|%d %b %Y}<br>%{y} tickets<extra>%{fullData.name}</extra>")
                fig.update_layout(
                    xaxis_title="Semana",
                    yaxis_title="Tickets Consumidos",
                    hovermode="x unified",
                    xaxis=dict(tickformat="%d %b", tickangle=-30),
                    legend_title="Casino",
                    bargap=0.2
                )
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("No hay datos para los filtros seleccionados.")
        else:
            st.info("No hay datos históricos de consumo para mostrar.")

    # 2. Vencimientos
    with tab2:
        st.markdown("#### Tickets próximos a vencer")
        if not df_vencimientos.empty:
            orden_rangos = ["7 días", "15 días", "30 días"]

            st.markdown("##### Por Servicio")
            venc_servicio = df_vencimientos.groupby(["Rango_Vencimiento", "NombreServicio"]).size().reset_index(name="Tickets")
            fig_serv = px.bar(
                venc_servicio, x="Rango_Vencimiento", y="Tickets", color="NombreServicio",
                barmode="group", category_orders={"Rango_Vencimiento": orden_rangos}
            )
            fig_serv.update_traces(hovertemplate="%{x}: %{y} tickets<extra>%{fullData.name}</extra>")
            fig_serv.update_layout(xaxis_title="Próximos a Vencer", yaxis_title="Tickets", legend_title="Servicio")
            st.plotly_chart(fig_serv, use_container_width=True)

            st.markdown("##### Por Casino y Cliente")
            cols_disp = [c for c in ["RazonSocial", "Casino"] if c in df_vencimientos.columns]
            if len(cols_disp) == 2:
                venc_cliente = (
                    df_vencimientos
                    .groupby(["RazonSocial", "Casino"])
                    .size()
                    .reset_index(name="Tickets")
                    .sort_values("Tickets", ascending=True)
                )
                alto = max(400, venc_cliente["RazonSocial"].nunique() * 28)
                fig_cli = px.bar(
                    venc_cliente, x="Tickets", y="RazonSocial", color="Casino",
                    orientation="h", barmode="stack"
                )
                fig_cli.update_traces(hovertemplate="%{y}<br>%{x} tickets<extra>%{fullData.name}</extra>")
                fig_cli.update_layout(
                    xaxis_title="Tickets por Vencer",
                    yaxis_title="",
                    legend_title="Casino",
                    height=alto,
                    yaxis=dict(tickfont=dict(size=11))
                )
                st.plotly_chart(fig_cli, use_container_width=True)
        else:
            st.success("No hay tickets próximos a vencer (30 días) en la selección actual.")

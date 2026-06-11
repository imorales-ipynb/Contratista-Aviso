import streamlit as st
import pandas as pd
from utils.historial import (cargar_historial, cargar_excel_cotizacion,
                              cargar_pdf_cotizacion, eliminar_cotizacion)

st.set_page_config(
    page_title="Historial de Cotizaciones",
    page_icon="📂",
    layout="wide"
)

st.title("Historial de Cotizaciones")

historial = cargar_historial()

if not historial:
    st.info("Aún no hay cotizaciones guardadas. Genera y exporta una cotización para verla aquí.")
    st.stop()

# ── Tabla resumen ─────────────────────────────────────────────────────────────
df_hist = pd.DataFrame([{
    "Fecha Emisión": r["fecha_emision"],
    "Vigencia":      r["vigencia"],
    "Casino":        r["casino"],
    "Cliente":       r["cliente"],
    "RUT":           r["rut"],
    "Cond. Pago":    r["condicion_pago"],
    "N° Servicios":  r["n_servicios"],
    "Total Neto":    f"${r['total_neto']:,.0f}",
    "IVA":           f"${r['iva']:,.0f}",
    "Total":         f"${r['total']:,.0f}",
    "_id":           r["id"],
} for r in historial])

st.markdown(f"**{len(df_hist)} cotización(es) registrada(s)**")
st.dataframe(df_hist.drop(columns=["_id"]), use_container_width=True, hide_index=True)

# ── Detalle y descarga ────────────────────────────────────────────────────────
st.markdown("---")
st.markdown("#### Ver detalle o descargar")

opciones = {
    f"{r['fecha_emision']}  —  {r['casino']}  —  {r['cliente'] or 'Sin cliente'}": r["id"]
    for r in historial
}

seleccion = st.selectbox("Seleccionar cotización:", options=list(opciones.keys()))

if seleccion:
    cid      = opciones[seleccion]
    registro = next(r for r in historial if r["id"] == cid)

    col_info, col_acciones = st.columns([3, 1])

    with col_info:
        st.markdown(f"**Casino:** {registro['casino']}  &nbsp;|&nbsp;  "
                    f"**Cliente:** {registro['cliente'] or '—'}  &nbsp;|&nbsp;  "
                    f"**RUT:** {registro['rut'] or '—'}")
        st.markdown(f"**Emisión:** {registro['fecha_emision']}  &nbsp;|&nbsp;  "
                    f"**Vigencia:** {registro['vigencia']}  &nbsp;|&nbsp;  "
                    f"**Cond. Pago:** {registro['condicion_pago']}")
        st.markdown(f"**Total Neto:** ${registro['total_neto']:,.0f}  &nbsp;|&nbsp;  "
                    f"**IVA:** ${registro['iva']:,.0f}  &nbsp;|&nbsp;  "
                    f"**Total:** ${registro['total']:,.0f}")

        if registro.get("items"):
            df_items = pd.DataFrame(registro["items"])
            df_items["Precio"]   = df_items["Precio"].apply(lambda x: f"${x:,.0f}")
            df_items["Subtotal"] = df_items["Subtotal"].apply(lambda x: f"${x:,.0f}")
            st.dataframe(df_items, use_container_width=True, hide_index=True)

    with col_acciones:
        excel_bytes = cargar_excel_cotizacion(cid)
        if excel_bytes:
            st.download_button(
                label="📥 Descargar Excel",
                data=excel_bytes,
                file_name=f"Cotizacion_{cid}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        else:
            st.warning("Excel no disponible.")

        pdf_bytes = cargar_pdf_cotizacion(cid)
        if pdf_bytes:
            st.download_button(
                label="📄 Descargar PDF",
                data=pdf_bytes,
                file_name=f"Cotizacion_{cid}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )

        st.markdown("---")
        if st.button("🗑️ Eliminar", use_container_width=True, key="btn_eliminar"):
            st.session_state["_confirmar_eliminar"] = cid

        if st.session_state.get("_confirmar_eliminar") == cid:
            st.warning("¿Confirmar eliminación?")
            col_si, col_no = st.columns(2)
            with col_si:
                if st.button("Sí", key="confirmar_si"):
                    eliminar_cotizacion(cid)
                    st.session_state.pop("_confirmar_eliminar", None)
                    st.rerun()
            with col_no:
                if st.button("No", key="confirmar_no"):
                    st.session_state.pop("_confirmar_eliminar", None)
                    st.rerun()

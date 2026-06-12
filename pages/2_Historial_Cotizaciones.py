import streamlit as st
import pandas as pd
from utils.historial import (cargar_historial, cargar_excel_cotizacion,
                              cargar_pdf_cotizacion, eliminar_cotizacion,
                              exportar_backup, importar_backup)

st.set_page_config(
    page_title="Historial de Cotizaciones",
    page_icon="📂",
    layout="wide"
)

st.title("Historial de Cotizaciones")

historial = cargar_historial()

# ── Backup / Restore (sidebar) ────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### Respaldo")
    backup_bytes = exportar_backup()
    st.download_button(
        label="💾 Descargar Backup JSON",
        data=backup_bytes,
        file_name="backup_historial.json",
        mime="application/json",
        use_container_width=True,
    )
    st.markdown("---")
    st.markdown("**Restaurar desde backup:**")
    uploaded = st.file_uploader("Subir backup JSON", type=["json"], key="backup_upload")
    if uploaded is not None:
        if st.button("⬆️ Restaurar", use_container_width=True, key="btn_restaurar"):
            try:
                importar_backup(uploaded.read())
                st.success("Backup restaurado correctamente.")
                st.rerun()
            except Exception as e:
                st.error(f"Error al restaurar: {e}")

if not historial:
    st.info("Aún no hay cotizaciones guardadas. Genera y exporta una cotización para verla aquí.")
    st.stop()

# ── Tabla resumen ─────────────────────────────────────────────────────────────
df_hist = pd.DataFrame([{
    "N° Cotización": r.get("numero", ""),
    "Fecha Emisión": r["fecha_emision"],
    "Vigencia":      r["vigencia"],
    "Casino":        r["casino"],
    "Cliente":       r.get("cliente", ""),
    "RUT":           r.get("rut", ""),
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

def _label(r):
    num = r.get("numero", "")
    lbl = f"{r['fecha_emision']}  —  {r['casino']}  —  {r.get('cliente') or 'Sin cliente'}"
    if num:
        lbl = f"{num}  |  {lbl}"
    return lbl

opciones = {_label(r): r["id"] for r in historial}

seleccion = st.selectbox("Seleccionar cotización:", options=list(opciones.keys()))

if seleccion:
    cid      = opciones[seleccion]
    registro = next(r for r in historial if r["id"] == cid)
    numero   = registro.get("numero", "")

    col_info, col_acciones = st.columns([3, 1])

    with col_info:
        if numero:
            st.markdown(f"**N° Cotización:** `{numero}`")
        st.markdown(f"**Casino:** {registro['casino']}  &nbsp;|&nbsp;  "
                    f"**Cliente:** {registro.get('cliente') or '—'}  &nbsp;|&nbsp;  "
                    f"**RUT:** {registro.get('rut') or '—'}")
        st.markdown(f"**Emisión:** {registro['fecha_emision']}  &nbsp;|&nbsp;  "
                    f"**Vigencia:** {registro['vigencia']}  &nbsp;|&nbsp;  "
                    f"**Cond. Pago:** {registro['condicion_pago']}")
        st.markdown(f"**Total Neto:** ${registro['total_neto']:,.0f}  &nbsp;|&nbsp;  "
                    f"**IVA:** ${registro['iva']:,.0f}  &nbsp;|&nbsp;  "
                    f"**Total:** ${registro['total']:,.0f}")

        if registro.get("items"):
            df_items = pd.DataFrame(registro["items"])
            if "Precio" in df_items.columns:
                df_items["Precio"]   = df_items["Precio"].apply(lambda x: f"${x:,.0f}")
            if "Subtotal" in df_items.columns:
                df_items["Subtotal"] = df_items["Subtotal"].apply(lambda x: f"${x:,.0f}")
            st.dataframe(df_items, use_container_width=True, hide_index=True)

    with col_acciones:
        nombre_base = f"Cotizacion_{numero}_{registro['casino'].replace(' ', '_')}" if numero else f"Cotizacion_{cid}"

        excel_bytes = cargar_excel_cotizacion(cid)
        if excel_bytes:
            st.download_button(
                label="📥 Descargar Excel",
                data=excel_bytes,
                file_name=f"{nombre_base}.xlsx",
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
                file_name=f"{nombre_base}.pdf",
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

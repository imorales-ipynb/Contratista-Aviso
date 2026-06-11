import streamlit as st
import pandas as pd
import datetime
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from fpdf import FPDF

# ── Constantes empresa ────────────────────────────────────────────────────────
EMAIL_EMPRESA    = "venta.ticket@casinoexpress.cl"
CONDICIONES_PAGO = ["Anticipado", "Tarjeta de Crédito", "Tarjeta de Débito",
                    "Crédito 30 días", "Crédito 60 días"]

# ── Helpers Excel ─────────────────────────────────────────────────────────────
def _celda(ws, fila, col, valor="", bold=False, size=10, color_font="000000",
           bg=None, alineacion="left", wrap=False, numero_fmt=None, borde=None):
    c = ws.cell(row=fila, column=col, value=valor)
    c.font = Font(name="Calibri", bold=bold, size=size, color=color_font)
    c.alignment = Alignment(horizontal=alineacion, vertical="center", wrap_text=wrap)
    if bg:
        c.fill = PatternFill(fill_type="solid", fgColor=bg)
    if numero_fmt:
        c.number_format = numero_fmt
    if borde:
        c.border = borde
    return c

def _borde_fino():
    s = Side(style="thin", color="BFBFBF")
    return Border(left=s, right=s, top=s, bottom=s)


def _exportar_excel(df_cot, casino, fecha, vigencia, cliente, rut, condicion_pago):
    wb = Workbook()
    ws = wb.active
    ws.title = "Cotización"

    AZUL    = "1F4E79"
    AZUL_CL = "D6E4F0"
    GRIS_H  = "595959"
    GRIS_F  = "F2F2F2"
    NARANJA = "C55A11"
    B       = _borde_fino()

    anchos = [5, 30, 18, 20, 14, 16, 10, 14]
    for i, a in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(i)].width = a

    ws.merge_cells("A1:H1")
    _celda(ws, 1, 1, "COTIZACIÓN DE SERVICIOS", bold=True, size=16,
           color_font=AZUL, alineacion="center")
    ws.row_dimensions[1].height = 34

    info_doc = [
        ("Cliente:",        cliente,                "Emisión:",     fecha.strftime("%d/%m/%Y")),
        ("RUT:",            rut,                    "Vigencia:",    vigencia.strftime("%d/%m/%Y")),
        ("Casino:",         casino,                 "Cond. Pago:",  condicion_pago),
        ("Email contacto:", EMAIL_EMPRESA,          "",             ""),
    ]
    for i, (l1, v1, l2, v2) in enumerate(info_doc, start=2):
        ws.merge_cells(f"A{i}:B{i}"); ws.merge_cells(f"C{i}:D{i}")
        ws.merge_cells(f"E{i}:F{i}"); ws.merge_cells(f"G{i}:H{i}")
        _celda(ws, i, 1, f"{l1}  {v1}", bold=(i == 2), size=10,
               color_font=AZUL if i == 2 else "000000")
        _celda(ws, i, 3, "")
        _celda(ws, i, 5, f"{l2}  {v2}", bold=(l2 != ""), size=10)
        _celda(ws, i, 7, "")
        ws.row_dimensions[i].height = 15

    ws.row_dimensions[6].height = 6

    HEADERS   = ["N°", "Servicio", "Tipo Servicio", "Alias",
                 "Cód. Servicio", "Precio Unitario", "Cantidad", "Subtotal"]
    HEADER_ROW = 7
    for ci, h in enumerate(HEADERS, 1):
        _celda(ws, HEADER_ROW, ci, h, bold=True, size=10, color_font="FFFFFF",
               bg=AZUL, alineacion="center", borde=B)
    ws.row_dimensions[HEADER_ROW].height = 22

    fmt_pesos  = '$#,##0'
    data_start = HEADER_ROW + 1
    for idx, (_, row) in enumerate(df_cot.iterrows(), 1):
        r      = data_start + idx - 1
        bg_row = GRIS_F if idx % 2 == 0 else None
        vals   = [idx, row.get("NombreServicio",""), row.get("TipoServicio",""),
                  row.get("Alias",""), row.get("Codigo Servicio",""),
                  row.get("Precio",0), row.get("Cantidad",0), row.get("Subtotal",0)]
        for ci, val in enumerate(vals, 1):
            alin = "right" if ci in (6,8) else ("center" if ci in (1,7) else "left")
            fmt  = fmt_pesos if ci in (6,8) else None
            _celda(ws, r, ci, val, size=10, bg=bg_row, alineacion=alin,
                   numero_fmt=fmt, borde=B)
        ws.row_dimensions[r].height = 16

    neto          = df_cot["Subtotal"].sum()
    iva           = neto * 0.19
    total_con_iva = neto + iva
    total_row     = data_start + len(df_cot)
    iva_row       = total_row + 1
    total_fin_row = total_row + 2

    def _fila_total(fila, etiqueta, valor, destacado=False):
        ws.merge_cells(f"A{fila}:G{fila}")
        _celda(ws, fila, 1, etiqueta, bold=destacado, size=11,
               color_font=AZUL if destacado else "000000",
               bg=AZUL_CL if destacado else "EBF3FB",
               alineacion="right", borde=B)
        _celda(ws, fila, 8, valor, bold=destacado, size=11,
               color_font=AZUL if destacado else "000000",
               bg=AZUL_CL if destacado else "EBF3FB",
               alineacion="right", numero_fmt=fmt_pesos, borde=B)
        ws.row_dimensions[fila].height = 18

    _fila_total(total_row,     "Total Neto", neto)
    _fila_total(iva_row,       "IVA (19%)",  iva)
    _fila_total(total_fin_row, "TOTAL",      total_con_iva, destacado=True)

    INFO_START = total_fin_row + 3
    h = INFO_START

    ws.merge_cells(f"A{h}:D{h}")
    _celda(ws, h, 1, "Datos de Transferencia", bold=True, size=10,
           color_font="FFFFFF", bg=GRIS_H, alineacion="left", borde=B)
    ws.row_dimensions[h].height = 18

    transferencia = ["Banco: Chile", "Cuenta Corriente: 167-01052-02",
                     "Rut: 78.793.360-2", "Casino Express S.A", f"Mail: {EMAIL_EMPRESA}"]
    for i, linea in enumerate(transferencia, 1):
        ws.merge_cells(f"A{h+i}:D{h+i}")
        _celda(ws, h+i, 1, linea, size=9, borde=B)
        ws.row_dimensions[h+i].height = 14

    texto_vigencia = (
        "Los tickets comprados tienen una vigencia de 100 días a contar de la fecha de "
        "emisión. Art. 41 Ley 19496. Por la naturaleza del servicio contratado, el prestador "
        "no efectuará devolución alguna de dinero en caso de no uso por parte del cliente, "
        "salvo que el servicio no esté disponible, dentro de los días y horas en que se presta."
    )
    ws.merge_cells(f"E{h}:H{h+len(transferencia)}")
    c_vig = ws.cell(row=h, column=5, value=texto_vigencia)
    c_vig.font = Font(name="Calibri", size=9, color=NARANJA)
    c_vig.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c_vig.border = B

    h = h + len(transferencia) + 1
    texto_horario = (
        "Los pagos realizados posterior a las 14:00 horas serán considerados como pagos del día "
        "siguiente, esto incluye los pagos informados al correo posterior el horario indicado."
    )
    alto_inf = 5
    ws.merge_cells(f"A{h}:D{h+alto_inf-1}")
    c_hor = ws.cell(row=h, column=1, value=texto_horario)
    c_hor.font = Font(name="Calibri", size=9, italic=True)
    c_hor.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    c_hor.border = B

    ws.merge_cells(f"E{h}:H{h}")
    _celda(ws, h, 5,
           "Estimado cliente para recibir nuestros servicios en el casino debe realizar el pago con anticipación:",
           bold=True, size=9, color_font="FFFFFF", bg=GRIS_H,
           alineacion="left", wrap=True, borde=B)
    ws.row_dimensions[h].height = 28

    instrucciones = [
        "Pago mediante transferencia electrónica 24 horas hábiles de anticipación.",
        "Pago mediante Webpay 48 horas hábiles de anticipación.",
        "Se solicita por favor programar sus pagos para evitar suspensión de los servicios.",
    ]
    for j, inst in enumerate(instrucciones, 1):
        ws.merge_cells(f"E{h+j}:H{h+j}")
        _celda(ws, h+j, 5, inst, size=9, color_font=NARANJA, borde=B)
        ws.row_dimensions[h+j].height = 14
    for r in range(h+1, h+alto_inf):
        ws.row_dimensions[r].height = 14

    output = BytesIO()
    wb.save(output)
    return output.getvalue()


def _exportar_pdf(df_cot, casino, fecha, vigencia, cliente, rut, condicion_pago):
    """Genera un PDF de la cotización con el mismo contenido que el Excel."""

    class PDF(FPDF):
        pass

    pdf = PDF(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_margins(15, 15, 15)
    W = 180  # ancho útil

    # Colores
    AZUL    = (31, 78, 121)
    AZUL_CL = (214, 228, 240)
    NARANJA = (197, 90, 17)
    GRIS_H  = (89, 89, 89)
    GRIS_F  = (242, 242, 242)

    # ── Título ────────────────────────────────────────────────────────────────
    pdf.set_font("Helvetica", "B", 16)
    pdf.set_text_color(*AZUL)
    pdf.cell(W, 12, "COTIZACION DE SERVICIOS", border=0, ln=1, align="C")
    pdf.ln(2)

    # ── Info cliente/documento ─────────────────────────────────────────────────
    info_rows = [
        ("Cliente:", cliente or "-", "Emision:", fecha.strftime("%d/%m/%Y")),
        ("RUT:",     rut     or "-", "Vigencia:", vigencia.strftime("%d/%m/%Y")),
        ("Casino:",  casino,         "Cond. Pago:", condicion_pago),
        ("Email:",   EMAIL_EMPRESA,  "",            ""),
    ]
    L, V = 22, 68
    for l1, v1, l2, v2 in info_rows:
        pdf.set_font("Helvetica", "B", 9); pdf.set_text_color(0, 0, 0)
        pdf.cell(L, 6, l1, ln=0)
        pdf.set_font("Helvetica", "", 9)
        pdf.cell(V, 6, v1, ln=0)
        pdf.set_font("Helvetica", "B", 9)
        pdf.cell(L, 6, l2, ln=0)
        pdf.set_font("Helvetica", "", 9)
        pdf.cell(V, 6, v2, ln=1)
    pdf.ln(4)

    # ── Encabezado tabla ──────────────────────────────────────────────────────
    HEADERS = ["N", "Servicio", "Tipo", "Alias", "Codigo", "Precio Unit.", "Cant.", "Subtotal"]
    COL_W   = [8, 50, 22, 22, 18, 22, 12, 26]   # suma = 180
    ALIGNS  = ["C", "L", "L", "L", "C", "R", "C", "R"]

    pdf.set_fill_color(*AZUL)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 9)
    for h, w, a in zip(HEADERS, COL_W, ALIGNS):
        pdf.cell(w, 8, h, border=1, ln=0, align=a, fill=True)
    pdf.ln()

    # ── Filas de servicios ────────────────────────────────────────────────────
    neto = 0.0
    for idx, (_, row) in enumerate(df_cot.iterrows(), 1):
        subtotal = float(row.get("Cantidad", 0)) * float(row.get("Precio", 0))
        neto += subtotal
        fill = (idx % 2 == 0)
        pdf.set_fill_color(*(GRIS_F if fill else (255, 255, 255)))
        pdf.set_text_color(0, 0, 0)
        pdf.set_font("Helvetica", "", 9)

        nombre = str(row.get("NombreServicio", ""))
        if len(nombre) > 30:
            nombre = nombre[:27] + "..."

        vals = [str(idx), nombre,
                str(row.get("TipoServicio", ""))[:18],
                str(row.get("Alias", ""))[:18],
                str(row.get("Codigo Servicio", "")),
                f"${float(row.get('Precio', 0)):,.0f}",
                str(int(row.get("Cantidad", 0))),
                f"${subtotal:,.0f}"]
        for val, w, a in zip(vals, COL_W, ALIGNS):
            pdf.cell(w, 7, val, border=1, ln=0, align=a, fill=True)
        pdf.ln()

    # ── Totales ───────────────────────────────────────────────────────────────
    iva   = neto * 0.19
    total = neto + iva
    TW    = sum(COL_W[:-1])  # 154 mm

    def _total(label, valor, destacado=False):
        if destacado:
            pdf.set_fill_color(*AZUL_CL)
            pdf.set_text_color(*AZUL)
            pdf.set_font("Helvetica", "B", 10)
        else:
            pdf.set_fill_color(235, 243, 251)
            pdf.set_text_color(0, 0, 0)
            pdf.set_font("Helvetica", "", 9)
        pdf.cell(TW, 7, label, border=1, ln=0, align="R", fill=True)
        pdf.cell(COL_W[-1], 7, f"${valor:,.0f}", border=1, ln=1, align="R", fill=True)

    _total("Total Neto", neto)
    _total("IVA (19%)", iva)
    _total("TOTAL", total, destacado=True)
    pdf.ln(5)

    # ── Pie: transferencia e instrucciones ────────────────────────────────────
    y0   = pdf.get_y()
    xL   = pdf.l_margin
    xR   = xL + 93
    cw   = 87

    # Encabezados pie
    pdf.set_fill_color(*GRIS_H)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_xy(xL, y0)
    pdf.cell(cw, 7, "Datos de Transferencia", border=1, ln=0, fill=True)
    pdf.set_xy(xR, y0)
    pdf.cell(cw, 7, "Instrucciones de Pago", border=1, ln=1, fill=True)

    y1 = pdf.get_y()
    transferencia = ["Banco: Chile", "Cuenta Corriente: 167-01052-02",
                     "Rut: 78.793.360-2", "Casino Express S.A",
                     f"Mail: {EMAIL_EMPRESA}"]
    instrucciones = [
        "Para recibir servicios, pague con anticipacion:",
        "- Transferencia: 24 hrs habiles de anticipacion.",
        "- Webpay: 48 hrs habiles de anticipacion.",
        "- Programe pagos para evitar suspension.",
        "",
    ]
    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Helvetica", "", 8)
    for i, linea in enumerate(transferencia):
        pdf.set_xy(xL, y1 + i * 5.5)
        pdf.cell(cw, 5.5, linea, border=1)
    pdf.set_text_color(*NARANJA)
    for i, inst in enumerate(instrucciones):
        pdf.set_xy(xR, y1 + i * 5.5)
        pdf.cell(cw, 5.5, inst, border=1)

    # Texto legal
    pdf.set_y(y1 + len(transferencia) * 5.5 + 3)
    pdf.set_text_color(*NARANJA)
    pdf.set_font("Helvetica", "I", 8)
    texto_legal = (
        "Los tickets comprados tienen una vigencia de 100 dias a contar de la fecha de emision. "
        "Art. 41 Ley 19496. Por la naturaleza del servicio contratado, el prestador no efectuara "
        "devolucion alguna de dinero en caso de no uso, salvo que el servicio no este disponible."
    )
    pdf.multi_cell(W, 4.5, texto_legal, border=1, align="L")

    return bytes(pdf.output())


# ── UI Cotizador ──────────────────────────────────────────────────────────────

def render_cotizador(df_precios, df_clientes=None):
    st.title("Cotizador de Servicios")

    if "cot_items" not in st.session_state:
        st.session_state.cot_items = []
    if "cot_casino_prev" not in st.session_state:
        st.session_state.cot_casino_prev = None

    if df_precios.empty:
        st.warning("No hay datos de precios cargados. Conéctate a FasesUAT para continuar.")
        if st.button("🔄 Cargar Precios desde FasesUAT"):
            from database import recargar_precios
            with st.spinner("Descargando precios..."):
                recargar_precios()
            st.rerun()
        return

    hoy           = datetime.date.today()
    vigencia_fecha = hoy + datetime.timedelta(days=7)

    # ── Casino ────────────────────────────────────────────────────────────────
    col_cas, col_reload = st.columns([6, 1])
    with col_cas:
        casinos_disp = sorted(df_precios["Nombre Casino"].dropna().unique().tolist())
        casino_sel   = st.selectbox("Casino:", options=casinos_disp, key="cot_casino")
    with col_reload:
        st.markdown("<div style='margin-top:28px'>", unsafe_allow_html=True)
        if st.button("🔄 Actualizar Precios", key="cot_reload"):
            from database import recargar_precios, recargar_clientes
            with st.spinner("Actualizando..."):
                recargar_precios()
                recargar_clientes()
            st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

    if st.session_state.cot_casino_prev != casino_sel:
        st.session_state.cot_items = []
        st.session_state.cot_casino_prev = casino_sel
        for k in ("cot_cliente_sel", "cot_rut_sel", "cot_nombre_manual", "cot_rut_manual"):
            st.session_state.pop(k, None)

    # ── Datos del Cliente ─────────────────────────────────────────────────────
    st.markdown("#### Datos del Cliente")

    manual = st.checkbox("Ingresar cliente manualmente (no está en el listado)",
                         key="cot_manual", value=False)

    if manual:
        # Entrada libre
        col_c1, col_c2, col_c3 = st.columns(3)
        with col_c1:
            nombre_manual = st.text_input("Nombre del Cliente:",
                                          key="cot_nombre_manual",
                                          placeholder="Razón social...")
        with col_c2:
            rut_manual = st.text_input("RUT:",
                                       key="cot_rut_manual",
                                       placeholder="XX.XXX.XXX-X")
        with col_c3:
            condicion = st.selectbox("Condiciones de Pago:",
                                     options=CONDICIONES_PAGO, key="cot_condicion")
        cliente_val = nombre_manual.strip()
        rut_auto    = rut_manual.strip()
    else:
        # Búsqueda en listado
        df_uniq = pd.DataFrame()
        if df_clientes is not None and not df_clientes.empty:
            df_uniq = (df_clientes[["RazonSocial", "RutContratista"]]
                       .drop_duplicates()
                       .sort_values("RazonSocial")
                       .reset_index(drop=True))

        PLACEHOLDER_NOMBRE = "— Buscar por nombre —"
        PLACEHOLDER_RUT    = "— Buscar por RUT —"

        opciones_nombre = [PLACEHOLDER_NOMBRE] + (df_uniq["RazonSocial"].tolist() if not df_uniq.empty else [])
        opciones_rut    = [PLACEHOLDER_RUT]    + (sorted(df_uniq["RutContratista"].astype(str).unique().tolist()) if not df_uniq.empty else [])

        if st.session_state.get("cot_cliente_sel", PLACEHOLDER_NOMBRE) not in opciones_nombre:
            st.session_state["cot_cliente_sel"] = PLACEHOLDER_NOMBRE
        if st.session_state.get("cot_rut_sel", PLACEHOLDER_RUT) not in opciones_rut:
            st.session_state["cot_rut_sel"] = PLACEHOLDER_RUT

        def _on_nombre():
            nombre = st.session_state.get("cot_cliente_sel", PLACEHOLDER_NOMBRE)
            if nombre != PLACEHOLDER_NOMBRE and not df_uniq.empty:
                match = df_uniq[df_uniq["RazonSocial"] == nombre]
                if not match.empty:
                    st.session_state["cot_rut_sel"] = str(match["RutContratista"].iloc[0])
            else:
                st.session_state["cot_rut_sel"] = PLACEHOLDER_RUT

        def _on_rut():
            rut = st.session_state.get("cot_rut_sel", PLACEHOLDER_RUT)
            if rut != PLACEHOLDER_RUT and not df_uniq.empty:
                match = df_uniq[df_uniq["RutContratista"].astype(str) == rut]
                if not match.empty:
                    st.session_state["cot_cliente_sel"] = match["RazonSocial"].iloc[0]
            else:
                st.session_state["cot_cliente_sel"] = PLACEHOLDER_NOMBRE

        col_c1, col_c2, col_c3 = st.columns(3)
        with col_c1:
            nombre_sel = st.selectbox("Nombre del Cliente:", options=opciones_nombre,
                                      key="cot_cliente_sel", on_change=_on_nombre)
        with col_c2:
            rut_sel = st.selectbox("RUT:", options=opciones_rut,
                                   key="cot_rut_sel", on_change=_on_rut)
        with col_c3:
            condicion = st.selectbox("Condiciones de Pago:",
                                     options=CONDICIONES_PAGO, key="cot_condicion")

        cliente_val = nombre_sel if nombre_sel != PLACEHOLDER_NOMBRE else ""
        rut_auto    = rut_sel    if rut_sel    != PLACEHOLDER_RUT    else ""

    col_d1, col_d2, col_d3 = st.columns(3)
    with col_d1:
        st.text_input("Email de contacto:", value=EMAIL_EMPRESA, disabled=True, key="cot_email")
    with col_d2:
        st.text_input("Fecha de Emisión:", value=hoy.strftime("%d/%m/%Y"), disabled=True, key="cot_emision")
    with col_d3:
        st.text_input("Vigencia de Cotización:", value=vigencia_fecha.strftime("%d/%m/%Y"),
                      disabled=True, key="cot_vigencia")

    # ── Buscador de servicios ─────────────────────────────────────────────────
    st.markdown("---")
    COLS_ITEM       = ["NombreServicio", "TipoServicio", "Alias", "Codigo Servicio", "Precio"]
    servicios_casino = (
        df_precios[df_precios["Nombre Casino"] == casino_sel]
        [[c for c in COLS_ITEM if c in df_precios.columns]]
        .drop_duplicates()
    )
    ya_agregados = {item["NombreServicio"] for item in st.session_state.cot_items}
    servicios_disp = servicios_casino[~servicios_casino["NombreServicio"].isin(ya_agregados)]

    col_search, col_add = st.columns([5, 1])
    with col_search:
        if servicios_disp.empty:
            st.info("Todos los servicios de este casino ya están en la cotización.")
            servicio_sel = None
        else:
            servicio_sel = st.selectbox(
                "Buscar servicio:",
                options=sorted(servicios_disp["NombreServicio"].unique().tolist()),
                key="cot_search"
            )
    with col_add:
        st.markdown("<div style='margin-top:28px'>", unsafe_allow_html=True)
        agregar = st.button("➕ Agregar", use_container_width=True, key="cot_agregar")
        st.markdown("</div>", unsafe_allow_html=True)

    if agregar and servicio_sel:
        fila = servicios_disp[servicios_disp["NombreServicio"] == servicio_sel].iloc[0]
        st.session_state.cot_items.append({
            "NombreServicio":  fila.get("NombreServicio", ""),
            "TipoServicio":    fila.get("TipoServicio", ""),
            "Alias":           fila.get("Alias", ""),
            "Codigo Servicio": fila.get("Codigo Servicio", ""),
            "Precio":          fila.get("Precio", 0),
            "Cantidad":        1,
        })
        st.rerun()

    # ── Vista previa ──────────────────────────────────────────────────────────
    st.markdown("#### Vista previa de cotización")

    if not st.session_state.cot_items:
        st.info("Selecciona un servicio y presiona **➕ Agregar** para comenzar.")
        return

    df_preview = pd.DataFrame(st.session_state.cot_items)
    df_preview["Eliminar"] = False

    df_editado = st.data_editor(
        df_preview,
        column_config={
            "NombreServicio":  st.column_config.TextColumn("Servicio",      disabled=True, width="large"),
            "TipoServicio":    st.column_config.TextColumn("Tipo",          disabled=True),
            "Alias":           st.column_config.TextColumn("Alias",         disabled=True),
            "Codigo Servicio": st.column_config.TextColumn("Código",        disabled=True),
            "Precio":          st.column_config.NumberColumn("Precio Unit.", disabled=True, format="$%d"),
            "Cantidad":        st.column_config.NumberColumn("Cantidad",    min_value=1, step=1, format="%d"),
            "Eliminar":        st.column_config.CheckboxColumn("🗑️"),
        },
        hide_index=True,
        use_container_width=True,
        key=f"cot_tabla_{casino_sel}_{len(st.session_state.cot_items)}",
    )

    for i, row in df_editado.iterrows():
        if i < len(st.session_state.cot_items):
            qty = row.get("Cantidad", 1)
            st.session_state.cot_items[i]["Cantidad"] = max(1, int(qty) if pd.notna(qty) else 1)

    col_del, col_clear, _ = st.columns([1, 1, 5])
    with col_del:
        if st.button("Eliminar marcados", key="cot_del"):
            st.session_state.cot_items = [
                item for i, item in enumerate(st.session_state.cot_items)
                if i >= len(df_editado) or not df_editado.iloc[i].get("Eliminar", False)
            ]
            st.rerun()
    with col_clear:
        if st.button("Limpiar todo", key="cot_clear"):
            st.session_state.cot_items = []
            st.rerun()

    # ── Totales ───────────────────────────────────────────────────────────────
    df_cot = df_editado.drop(columns=["Eliminar"]).copy()
    df_cot["Subtotal"] = df_cot["Cantidad"] * df_cot["Precio"]

    col_k1, col_k2, col_k3 = st.columns(3)
    with col_k1:
        st.metric("Servicios en cotización", len(df_cot))
    with col_k2:
        st.metric("Total tickets cotizados", f"{int(df_cot['Cantidad'].sum()):,}")
    with col_k3:
        st.metric("Total Cotización", f"${df_cot['Subtotal'].sum():,.0f}")

    # ── Exportar ──────────────────────────────────────────────────────────────
    nombre_cliente  = cliente_val or ""
    nombre_base     = f"Cotizacion_{casino_sel.replace(' ', '_')}_{hoy.strftime('%Y%m%d')}"

    export_args = dict(
        df_cot        = df_cot,
        casino        = casino_sel,
        fecha         = hoy,
        vigencia      = vigencia_fecha,
        cliente       = nombre_cliente,
        rut           = rut_auto,
        condicion_pago= condicion,
    )

    excel_data = _exportar_excel(**export_args)
    pdf_data   = _exportar_pdf(**export_args)

    col_dl1, col_dl2 = st.columns(2)

    with col_dl1:
        if st.download_button(
            label="📥 Descargar Excel",
            data=excel_data,
            file_name=f"{nombre_base}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        ):
            try:
                from utils.historial import guardar_cotizacion
                guardar_cotizacion(
                    **export_args,
                    excel_bytes = excel_data,
                    pdf_bytes   = pdf_data,
                )
                st.success("Cotización guardada en el historial.")
            except Exception as e:
                st.warning(f"No se pudo guardar en historial: {e}")

    with col_dl2:
        st.download_button(
            label="📄 Descargar PDF",
            data=pdf_data,
            file_name=f"{nombre_base}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

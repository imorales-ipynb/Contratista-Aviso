import json
import os
import datetime
import pandas as pd

HISTORIAL_PATH   = os.path.join("data", "historial_cotizaciones.json")
COTIZACIONES_DIR = os.path.join("data", "cotizaciones")


def _asegurar_dirs():
    os.makedirs(COTIZACIONES_DIR, exist_ok=True)


def guardar_cotizacion(df_cot, casino, fecha, vigencia, cliente, rut,
                       condicion_pago, excel_bytes, pdf_bytes=None):
    _asegurar_dirs()

    ts  = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    cid = f"{ts}_{casino[:20].replace(' ', '_')}"

    # Guardar Excel
    excel_path = os.path.join(COTIZACIONES_DIR, f"{cid}.xlsx")
    with open(excel_path, "wb") as f:
        f.write(excel_bytes)

    # Guardar PDF (si se proporcionó)
    pdf_path = None
    if pdf_bytes:
        pdf_path = os.path.join(COTIZACIONES_DIR, f"{cid}.pdf")
        with open(pdf_path, "wb") as f:
            f.write(pdf_bytes)

    items = df_cot[["NombreServicio", "TipoServicio", "Alias",
                    "Precio", "Cantidad", "Subtotal"]].to_dict("records")

    registro = {
        "id":             cid,
        "fecha_emision":  fecha.strftime("%Y-%m-%d"),
        "vigencia":       vigencia.strftime("%Y-%m-%d"),
        "casino":         casino,
        "cliente":        cliente,
        "rut":            rut,
        "condicion_pago": condicion_pago,
        "n_servicios":    len(df_cot),
        "total_neto":     round(float(df_cot["Subtotal"].sum()), 0),
        "iva":            round(float(df_cot["Subtotal"].sum() * 0.19), 0),
        "total":          round(float(df_cot["Subtotal"].sum() * 1.19), 0),
        "archivo":        excel_path,
        "archivo_pdf":    pdf_path,
        "items":          items,
    }

    historial = _leer_json()
    historial.append(registro)

    with open(HISTORIAL_PATH, "w", encoding="utf-8") as f:
        json.dump(historial, f, ensure_ascii=False, indent=2)

    return cid


def cargar_historial() -> list:
    return list(reversed(_leer_json()))


def cargar_excel_cotizacion(cid: str):
    path = os.path.join(COTIZACIONES_DIR, f"{cid}.xlsx")
    if not os.path.exists(path):
        return None
    with open(path, "rb") as f:
        return f.read()


def cargar_pdf_cotizacion(cid: str):
    path = os.path.join(COTIZACIONES_DIR, f"{cid}.pdf")
    if not os.path.exists(path):
        return None
    with open(path, "rb") as f:
        return f.read()


def eliminar_cotizacion(cid: str):
    historial = _leer_json()
    historial = [r for r in historial if r["id"] != cid]
    with open(HISTORIAL_PATH, "w", encoding="utf-8") as f:
        json.dump(historial, f, ensure_ascii=False, indent=2)

    for ext in (".xlsx", ".pdf"):
        path = os.path.join(COTIZACIONES_DIR, f"{cid}{ext}")
        if os.path.exists(path):
            os.remove(path)


def _leer_json() -> list:
    if not os.path.exists(HISTORIAL_PATH):
        return []
    try:
        with open(HISTORIAL_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []

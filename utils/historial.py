import json
import os
import datetime
import pandas as pd

from config import HISTORIAL_JSON as HISTORIAL_PATH, HISTORIAL_DIR as COTIZACIONES_DIR


def _asegurar_dirs():
    os.makedirs(COTIZACIONES_DIR, exist_ok=True)
    parent = os.path.dirname(HISTORIAL_PATH)
    if parent:
        os.makedirs(parent, exist_ok=True)


# ── Correlativos ──────────────────────────────────────────────────────────────

def _siguiente_numero() -> int:
    nums = [r.get("numero_int", 0) for r in _leer_json() if r.get("numero_int")]
    return max(nums, default=0) + 1


def formato_numero(n: int, year: int) -> str:
    return f"COT-{year}-{n:03d}"


# ── CRUD ──────────────────────────────────────────────────────────────────────

def guardar_cotizacion(df_cot, casino, fecha, vigencia, cliente, rut,
                       condicion_pago, excel_bytes, pdf_bytes=None, numero=""):
    _asegurar_dirs()

    ts  = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    cid = f"{ts}_{casino[:20].replace(' ', '_')}"

    excel_path = os.path.join(COTIZACIONES_DIR, f"{cid}.xlsx")
    with open(excel_path, "wb") as f:
        f.write(excel_bytes)

    pdf_path = None
    if pdf_bytes:
        pdf_path = os.path.join(COTIZACIONES_DIR, f"{cid}.pdf")
        with open(pdf_path, "wb") as f:
            f.write(pdf_bytes)

    numero_int = 0
    if numero and "-" in numero:
        try:
            numero_int = int(numero.split("-")[-1])
        except Exception:
            pass

    items = df_cot[["NombreServicio", "TipoServicio", "Alias",
                    "Precio", "Cantidad", "Subtotal"]].to_dict("records")

    registro = {
        "id":             cid,
        "numero":         numero,
        "numero_int":     numero_int,
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
    return open(path, "rb").read() if os.path.exists(path) else None


def cargar_pdf_cotizacion(cid: str):
    path = os.path.join(COTIZACIONES_DIR, f"{cid}.pdf")
    return open(path, "rb").read() if os.path.exists(path) else None


def eliminar_cotizacion(cid: str):
    historial = [r for r in _leer_json() if r["id"] != cid]
    with open(HISTORIAL_PATH, "w", encoding="utf-8") as f:
        json.dump(historial, f, ensure_ascii=False, indent=2)
    for ext in (".xlsx", ".pdf"):
        p = os.path.join(COTIZACIONES_DIR, f"{cid}{ext}")
        if os.path.exists(p):
            os.remove(p)


# ── Backup / Restore ──────────────────────────────────────────────────────────

def exportar_backup() -> bytes:
    return json.dumps(_leer_json(), ensure_ascii=False, indent=2).encode("utf-8")


def importar_backup(contenido: bytes):
    datos = json.loads(contenido.decode("utf-8"))
    _asegurar_dirs()
    with open(HISTORIAL_PATH, "w", encoding="utf-8") as f:
        json.dump(datos, f, ensure_ascii=False, indent=2)


def _leer_json() -> list:
    if not os.path.exists(HISTORIAL_PATH):
        return []
    try:
        with open(HISTORIAL_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []

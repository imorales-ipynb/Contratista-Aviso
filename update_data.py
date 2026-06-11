"""
update_data.py — Actualizador local de datos
Consulta SQL Server, genera archivos parquet y sube a GitHub.
Ejecutar manualmente o programar con el Programador de Tareas de Windows.

Uso:  python update_data.py
Logs: logs/update.log
"""
import os
import sys
import subprocess
import datetime
import pandas as pd
import numpy as np

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from config import DB_CONN_STR_SPC, DB_CONN_STR_JERARQUIA, DB_CONN_STR_FASESUAT

DATA_DIR       = "data"
CONSUMO_PATH   = os.path.join(DATA_DIR, "consumo.parquet")
JERARQUIA_PATH = os.path.join(DATA_DIR, "jerarquia.parquet")
SERVICIOS_PATH = os.path.join(DATA_DIR, "servicios.parquet")
PRECIOS_PATH   = os.path.join(DATA_DIR, "precios.parquet")
CLIENTES_PATH  = os.path.join(DATA_DIR, "clientes.parquet")

PARQUET_FILES = [CONSUMO_PATH, JERARQUIA_PATH, SERVICIOS_PATH, PRECIOS_PATH, CLIENTES_PATH]

_ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")


def log(msg):
    line = f"[{_ts}] {msg}"
    print(line)


def _engine(conn_str):
    from sqlalchemy import create_engine
    return create_engine(conn_str)


def descargar_consumos():
    log("Conectando a BDSPC (consumos + servicios)...")
    engine = _engine(DB_CONN_STR_SPC)

    query = """
    SELECT
        cs.nombreCasino             AS Casino,
        ct.razonSocial              AS RazonSocial,
        cm.idCasino                 AS CC,
        cm.idServicio               AS IdServicio,
        cm.nroCompra                AS NroCompra,
        CASE WHEN cm.usado = 1 THEN 'SI' ELSE 'NO' END AS Usado,
        cm.fechaVencimiento         AS FechaVencimiento,
        cm.fechaUso                 AS FechaUso,
        cm.fechaSincronizacion      AS FechaSincronizacion,
        ct.rut                      AS RutContratista,
        serv.nombreServicio         AS NombreServicio
    FROM BDSPC.dbo.CompraMovimiento cm WITH(NOLOCK)
    JOIN contratista ct  ON cm.idContratista = ct.idContratista
    JOIN Casino      cs  ON cs.idCasino      = cm.idCasino
    JOIN Servicio    serv ON serv.idServicio  = cm.idServicio
    WHERE cm.fechaSincronizacion >= DATEADD(day, -100, GETDATE())
    """

    df = pd.read_sql(query, engine)
    df["CC"] = df["CC"].astype(str)
    for col in ["FechaVencimiento", "FechaUso", "FechaSincronizacion"]:
        df[col] = pd.to_datetime(df[col], errors="coerce")

    hoy = pd.Timestamp(datetime.date.today())
    df["Saldo Pendiente"] = np.where(
        (df["Usado"] == "NO") & (df["FechaVencimiento"] > hoy),
        "Disponible", "No Disponible"
    )

    df_servicios = (
        df[["IdServicio", "NombreServicio"]]
        .drop_duplicates()
        .rename(columns={"IdServicio": "ServicioId", "NombreServicio": "Servicio"})
        .assign(Grupo="Sin Grupo", Tipo="Principal")
    )

    log(f"  Consumos:  {len(df):,} registros")
    log(f"  Servicios: {len(df_servicios)} tipos")
    return df, df_servicios


def descargar_jerarquia():
    log("Conectando a cerberus_minuta (jerarquía)...")
    engine = _engine(DB_CONN_STR_JERARQUIA)

    query = """
    SELECT area AS Area, gop AS GOP, jop AS JOP, cc AS CC, name AS Casino
    FROM cerberus_minuta.dbo.vt_jopgop_ax
    WHERE NOT jop  LIKE '%cerrado%'
      AND NOT name LIKE 'ope%'
      AND NOT gop  LIKE 'da%'
      AND cc = ccpadre
    """

    df = pd.read_sql(query, engine)
    df["CC"] = df["CC"].astype(str)
    log(f"  Jerarquía: {len(df):,} registros")
    return df


def descargar_clientes():
    log("Descargando listado de contratistas (BDSPC, sin filtro fecha)...")
    engine = _engine(DB_CONN_STR_SPC)

    query = """
    SELECT DISTINCT
        cs.nombreCasino AS Casino,
        cm.idCasino     AS CC,
        ct.razonSocial  AS RazonSocial,
        ct.rut          AS RutContratista
    FROM BDSPC.dbo.CompraMovimiento cm WITH(NOLOCK)
    JOIN contratista ct ON cm.idContratista = ct.idContratista
    JOIN Casino      cs ON cs.idCasino      = cm.idCasino
    ORDER BY ct.razonSocial
    """

    df = pd.read_sql(query, engine)
    df["CC"] = df["CC"].astype(str)
    log(f"  Contratistas: {len(df):,} únicos")
    return df


def descargar_precios():
    log("Conectando a FasesUAT (catálogo de precios)...")
    engine = _engine(DB_CONN_STR_FASESUAT)

    query = """
    SELECT
        TJ.CC, TJ.casino AS [Nombre Casino], TJ.area, TJ.gop, TJ.jop,
        S.Cod_ser AS [Codigo Servicio], S.NombreServicio,
        CS.Alias, CS.Precio, CS.FechaDesde, CS.FechaHasta,
        TS.NombreTipoServicio AS TipoServicio
    FROM Entidad.CasinoServicio CS
        JOIN Jerarquia.Casino     C  ON CS.idCasino     = C.idCasino
        JOIN Maestro.Servicio     S  ON CS.idServicio   = S.idServicio
        JOIN Maestro.TipoServicio TS ON S.idTipoServicio = TS.idTipoServicio
        JOIN ppto_ce.dbo.tbl_jerarquia TJ ON C.CentroCosto = TJ.CC
    WHERE CS.FechaHasta = (
        SELECT MAX(FechaHasta) FROM Entidad.CasinoServicio
        WHERE idCasino = CS.idCasino AND idServicio = CS.idServicio
    )
    ORDER BY CS.idCasino, CS.idServicio
    """

    df = pd.read_sql(query, engine)
    log(f"  Precios: {len(df):,} registros")
    return df


def guardar_parquet(df_consumo, df_jerarquia, df_servicios, df_clientes, df_precios):
    os.makedirs(DATA_DIR, exist_ok=True)
    df_consumo.to_parquet(CONSUMO_PATH,     index=False)
    df_jerarquia.to_parquet(JERARQUIA_PATH, index=False)
    df_servicios.to_parquet(SERVICIOS_PATH, index=False)
    df_clientes.to_parquet(CLIENTES_PATH,   index=False)
    df_precios.to_parquet(PRECIOS_PATH,     index=False)
    log("Archivos parquet guardados en data/")


def git_push():
    log("Subiendo cambios a GitHub...")

    # Stage solo los parquet
    subprocess.run(["git", "add"] + PARQUET_FILES, check=True)

    # Commit (puede ser "nothing to commit" si no cambió nada)
    result = subprocess.run(
        ["git", "commit", "-m", f"data: actualización {datetime.datetime.now():%Y-%m-%d %H:%M}"],
        capture_output=True, text=True
    )
    if result.returncode != 0:
        if "nothing to commit" in result.stdout or "nothing to commit" in result.stderr:
            log("Sin cambios en los datos. Push omitido.")
            return
        raise RuntimeError(f"git commit falló:\n{result.stderr}")

    subprocess.run(["git", "push"], check=True)
    log("[OK] Push a GitHub completado. Streamlit Cloud actualizará en ~30 seg.")


def main():
    log("=" * 50)
    log("Inicio de actualización de datos")
    log("=" * 50)

    df_consumo,  df_servicios = descargar_consumos()
    df_jerarquia              = descargar_jerarquia()
    df_clientes               = descargar_clientes()
    df_precios                = descargar_precios()

    guardar_parquet(df_consumo, df_jerarquia, df_servicios, df_clientes, df_precios)

    git_push()

    log("=== Actualización completada ===")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        log(f"[ERROR FATAL] {exc}")
        sys.exit(1)

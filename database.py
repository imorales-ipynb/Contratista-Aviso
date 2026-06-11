import pandas as pd
import numpy as np
import os
import datetime
import streamlit as st
from config import DB_CONN_STR_SPC, DB_CONN_STR_JERARQUIA, DB_CONN_STR_FASESUAT, CLOUD_MODE

# SQLAlchemy y pyodbc solo se necesitan localmente para conectar a SQL Server
if not CLOUD_MODE:
    try:
        from sqlalchemy import create_engine, text
    except ImportError:
        pass

DATA_DIR = "data"
CONSUMO_PATH = os.path.join(DATA_DIR, "consumo.parquet")
JERARQUIA_PATH = os.path.join(DATA_DIR, "jerarquia.parquet")
SERVICIOS_PATH = os.path.join(DATA_DIR, "servicios.parquet")
PRECIOS_PATH   = os.path.join(DATA_DIR, "precios.parquet")
CLIENTES_PATH  = os.path.join(DATA_DIR, "clientes.parquet")

def _test_connection(engine, nombre_bd):
    """Prueba la conexión realmente ejecutando una query simple."""
    try:
        with engine.connect() as conn:
            conn.execute(text("SELECT 1"))
        print(f"[OK] Conexión exitosa a {nombre_bd}")
        return True
    except Exception as e:
        msg = f"[ERROR] Error al conectar a {nombre_bd}: {e}"
        print(msg)
        st.error(msg)
        return False

def _generar_datos_simulados():
    """Genera datos de prueba si no hay conexión a base de datos real."""
    print("Generando datos simulados...")
    np.random.seed(42)
    clientes = [f"Cliente {i}" for i in range(1, 101)]
    ruts = [f"{np.random.randint(10000000, 99999999)}-{np.random.randint(0, 9)}" for _ in clientes]
    
    df_jerarquia = pd.DataFrame({
        "Area": np.random.choice(["Norte", "Sur", "Centro"], 100),
        "Casino": [f"Casino {i}" for i in range(1, 11)] * 10,
        "CC": [f"{i}" for i in range(1001, 1101)],
        "GOP": [f"GOP-{i}" for i in np.random.randint(1, 6, 100)],
        "JOP": [f"JOP-{i}" for i in np.random.randint(1, 11, 100)],
    })
    
    df_servicios = pd.DataFrame({
        "Grupo": ["Alimentación", "Alimentación", "Hospedaje", "Transporte"],
        "Servicio": ["Almuerzo", "Cena", "Cama", "Pasaje"],
        "ServicioId": [1, 2, 3, 4],
        "Tipo": ["Principal", "Principal", "Adicional", "Adicional"]
    })
    
    hoy = datetime.date.today()
    registros = []
    
    for i, cliente in enumerate(clientes):
        servicios_comprados = np.random.choice([1, 2, 3, 4], np.random.randint(1, 4), replace=False)
        for serv_id in servicios_comprados:
            nro_compra = np.random.randint(10000, 99999)
            fecha_compra = hoy - datetime.timedelta(days=np.random.randint(0, 100))
            cantidad_tickets = np.random.randint(50, 500)
            tickets_usados = int(cantidad_tickets * np.random.uniform(0.1, 0.9))
            
            for t in range(cantidad_tickets):
                usado = "SI" if t < tickets_usados else "NO"
                fecha_uso = None
                if usado == "SI":
                    dias_desde_compra = (hoy - fecha_compra).days
                    if dias_desde_compra > 0:
                        fecha_uso = fecha_compra + datetime.timedelta(days=np.random.randint(0, dias_desde_compra + 1))
                    else:
                        fecha_uso = fecha_compra
                
                fecha_vencimiento = fecha_compra + datetime.timedelta(days=100)
                if usado == "NO" and fecha_vencimiento > hoy:
                    saldo_pendiente = "Disponible"
                else:
                    saldo_pendiente = "No Disponible"
                    
                registros.append({
                    "Casino": df_jerarquia.iloc[i]["Casino"],
                    "CC": df_jerarquia.iloc[i]["CC"],
                    "FechaSincronizacion": pd.Timestamp(fecha_compra),
                    "FechaUso": pd.Timestamp(fecha_uso) if fecha_uso else None,
                    "IdServicio": serv_id,
                    "NombreServicio": df_servicios[df_servicios["ServicioId"] == serv_id]["Servicio"].values[0],
                    "MesAno": fecha_compra.strftime("%Y-%m"),
                    "NroCompra": nro_compra,
                    "RazonSocial": cliente,
                    "RutContratista": ruts[i],
                    "RutCC": f"{ruts[i][:8]}",
                    "SemanaAno": fecha_compra.isocalendar()[1],
                    "Usado": usado,
                    "Saldo Pendiente": saldo_pendiente
                })
                
    df_consumo = pd.DataFrame(registros)
    return df_consumo, df_jerarquia, df_servicios

def recargar_base_datos(use_mock=False):
    """
    Descarga la información desde SQL Server y la guarda en Parquet local.
    """
    if not os.path.exists(DATA_DIR):
        os.makedirs(DATA_DIR)
    
    if use_mock:
        st.info("Modo simulado activado. Generando datos de prueba...")
        df_consumo, df_jerarquia, df_servicios = _generar_datos_simulados()
    else:
        # Intentar conexión real
        try:
            st.info(f"Intentando conectar a SQL Server...")
            print(f"Cadena SPC: {DB_CONN_STR_SPC}")
            print(f"Cadena Jerarquia: {DB_CONN_STR_JERARQUIA}")
            
            engine_spc = create_engine(DB_CONN_STR_SPC)
            engine_jerarquia = create_engine(DB_CONN_STR_JERARQUIA)
            
            # Probar la conexión realmente
            ok_spc = _test_connection(engine_spc, "BDSPC")
            ok_jer = _test_connection(engine_jerarquia, "cerberus_minuta")
            
            if not ok_spc or not ok_jer:
                st.warning("No se pudo conectar a una o más bases de datos. Usando datos simulados como respaldo.")
                df_consumo, df_jerarquia, df_servicios = _generar_datos_simulados()
            else:
                st.info("Conexiones OK. Descargando datos...")
                
                query_consumo = """
                select 
                    cs.nombreCasino Casino, 
                    razonSocial as RazonSocial, 
                    cm.idCasino as CC, 
                    cm.idServicio as IdServicio, 
                    nroCompra as NroCompra, 
                    case when usado = 1 then 'SI' else 'NO' end as Usado, 
                    fechaVencimiento as FechaVencimiento, 
                    fechaUso as FechaUso, 
                    fechaSincronizacion as FechaSincronizacion,
                    ct.rut as RutContratista,
                    serv.nombreServicio as NombreServicio
                from BDSPC.dbo.CompraMovimiento cm with(nolock)
                join contratista ct on cm.idContratista = ct.idContratista
                join Casino cs on cs.idCasino = cm.idCasino
                join Servicio serv on serv.idServicio = cm.idServicio
                where cm.fechaSincronizacion >= DATEADD(day, -100, GETDATE())
                """
                
                query_jerarquia = """
                select 
                    area as Area, 
                    gop as GOP, 
                    jop as JOP, 
                    cc as CC, 
                    name as Casino
                from cerberus_minuta.dbo.vt_jopgop_ax
                where not jop like '%cerrado%'     
                     and not name like 'ope%'
                     and not gop like 'da%'
                     and cc = ccpadre
                """
                
                df_consumo = pd.read_sql(query_consumo, engine_spc)
                st.info(f"Consumo: {len(df_consumo):,} registros descargados.")
                
                df_jerarquia = pd.read_sql(query_jerarquia, engine_jerarquia)
                st.info(f"Jerarquía: {len(df_jerarquia):,} registros descargados.")
                
                # Asegurar compatibilidad de tipos para el merge (CC como string)
                df_consumo["CC"] = df_consumo["CC"].astype(str)
                df_jerarquia["CC"] = df_jerarquia["CC"].astype(str)
                
                # Convertir fechas
                df_consumo["FechaVencimiento"] = pd.to_datetime(df_consumo["FechaVencimiento"], errors='coerce')
                df_consumo["FechaUso"] = pd.to_datetime(df_consumo["FechaUso"], errors='coerce')
                df_consumo["FechaSincronizacion"] = pd.to_datetime(df_consumo["FechaSincronizacion"], errors='coerce')
                
                hoy = pd.Timestamp(datetime.date.today())
                
                df_consumo["Saldo Pendiente"] = np.where(
                    (df_consumo["Usado"] == "NO") & (df_consumo["FechaVencimiento"] > hoy),
                    "Disponible",
                    "No Disponible"
                )
                
                # Crear tabla de Servicios a partir de los datos ya descargados
                df_servicios = df_consumo[["IdServicio", "NombreServicio"]].drop_duplicates()
                df_servicios = df_servicios.rename(columns={"IdServicio": "ServicioId", "NombreServicio": "Servicio"})
                df_servicios["Grupo"] = "Sin Grupo"
                df_servicios["Tipo"] = "Principal"
                
                st.success(f"✅ Datos descargados correctamente desde SQL Server.")
                
        except Exception as e:
            msg = f"Error inesperado al leer base de datos: {e}"
            print(msg)
            st.error(msg)
            st.warning("Usando datos simulados como respaldo.")
            df_consumo, df_jerarquia, df_servicios = _generar_datos_simulados()
        
    # Guardar localmente
    df_consumo.to_parquet(CONSUMO_PATH, index=False)
    df_jerarquia.to_parquet(JERARQUIA_PATH, index=False)
    df_servicios.to_parquet(SERVICIOS_PATH, index=False)
    print("Datos guardados en repositorio local.")

def recargar_clientes():
    """Descarga el listado completo de contratistas únicos desde BDSPC (sin filtro de fecha)."""
    if not os.path.exists(DATA_DIR):
        os.makedirs(DATA_DIR)
    try:
        engine = create_engine(DB_CONN_STR_SPC)
        query = """
        SELECT DISTINCT
            cs.nombreCasino     AS Casino,
            cm.idCasino         AS CC,
            ct.razonSocial      AS RazonSocial,
            ct.rut              AS RutContratista
        FROM BDSPC.dbo.CompraMovimiento cm WITH(NOLOCK)
        JOIN contratista ct ON cm.idContratista = ct.idContratista
        JOIN Casino      cs ON cs.idCasino      = cm.idCasino
        ORDER BY ct.razonSocial
        """
        df = pd.read_sql(query, engine)
        df["CC"] = df["CC"].astype(str)
        df.to_parquet(CLIENTES_PATH, index=False)
        print(f"[OK] Clientes descargados: {len(df):,} registros únicos.")
        return df
    except Exception as e:
        print(f"[ERROR] Error al cargar clientes: {e}")
        st.error(f"Error al cargar clientes desde BDSPC: {e}")
        return pd.DataFrame()


def cargar_clientes():
    """Carga el listado de contratistas desde caché local. Si no existe, lo descarga."""
    if not os.path.exists(CLIENTES_PATH):
        if CLOUD_MODE:
            return pd.DataFrame()
        return recargar_clientes()
    try:
        return pd.read_parquet(CLIENTES_PATH)
    except Exception as e:
        print(f"[ERROR] Error al leer clientes locales: {e}")
        return pd.DataFrame()


def recargar_precios():
    """Descarga el catálogo de precios vigentes desde FasesUAT y lo guarda en caché local."""
    if not os.path.exists(DATA_DIR):
        os.makedirs(DATA_DIR)
    try:
        engine = create_engine(DB_CONN_STR_FASESUAT)
        query = """
        SELECT
            TJ.CC                           AS cc,
            TJ.casino                       AS [Nombre Casino],
            TJ.area                         AS area,
            TJ.gop                          AS gop,
            TJ.jop                          AS jop,
            S.Cod_ser                       AS [Codigo Servicio],
            S.NombreServicio,
            CS.Alias,
            CS.Precio,
            CS.FechaDesde,
            CS.FechaHasta,
            TS.NombreTipoServicio           AS TipoServicio
        FROM Entidad.CasinoServicio CS
            JOIN Jerarquia.Casino       C  ON CS.idCasino    = C.idCasino
            JOIN Maestro.Servicio       S  ON CS.idServicio  = S.idServicio
            JOIN Maestro.TipoServicio   TS ON S.idTipoServicio = TS.idTipoServicio
            JOIN ppto_ce.dbo.tbl_jerarquia TJ ON C.CentroCosto = TJ.CC
        WHERE CS.FechaHasta = (
            SELECT MAX(FechaHasta)
            FROM Entidad.CasinoServicio
            WHERE idCasino = CS.idCasino AND idServicio = CS.idServicio
        )
        ORDER BY CS.idCasino, CS.idServicio
        """
        df_precios = pd.read_sql(query, engine)
        df_precios.to_parquet(PRECIOS_PATH, index=False)
        print("[OK] Precios descargados correctamente.")
        return df_precios
    except Exception as e:
        print(f"[ERROR] Error al cargar precios desde FasesUAT: {e}")
        st.error(f"Error al cargar precios desde FasesUAT: {e}")
        return pd.DataFrame()


def cargar_precios():
    """Carga el catálogo de precios desde caché local. Si no existe, intenta descargarlo."""
    if not os.path.exists(PRECIOS_PATH):
        if CLOUD_MODE:
            return pd.DataFrame()
        return recargar_precios()
    try:
        return pd.read_parquet(PRECIOS_PATH)
    except Exception as e:
        print(f"[ERROR] Error al leer precios locales: {e}")
        return pd.DataFrame()


def cargar_datos():
    """
    Carga los datos desde el repositorio local (parquet).
    En CLOUD_MODE nunca intenta conectar a SQL Server.
    """
    parquet_ok = (
        os.path.exists(CONSUMO_PATH) and
        os.path.exists(JERARQUIA_PATH) and
        os.path.exists(SERVICIOS_PATH)
    )
    if not parquet_ok:
        if CLOUD_MODE:
            st.error("Los archivos de datos no están disponibles. Ejecuta update_data.py localmente y vuelve a hacer push.")
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
        print("No se encontró caché local. Iniciando carga inicial...")
        recargar_base_datos(use_mock=False)
        
    try:
        df_consumo = pd.read_parquet(CONSUMO_PATH)
        df_jerarquia = pd.read_parquet(JERARQUIA_PATH)
        df_servicios = pd.read_parquet(SERVICIOS_PATH)
        return df_consumo, df_jerarquia, df_servicios
    except Exception as e:
        print(f"Error al cargar Parquet local: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

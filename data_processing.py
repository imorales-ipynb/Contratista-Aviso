import pandas as pd
import datetime
import numpy as np
from config import RIESGO_CRITICO_SEMANAS, RIESGO_ALTO_SEMANAS, RIESGO_MEDIO_SEMANAS

def get_start_of_week(dt):
    """Devuelve el lunes de la semana a la que pertenece la fecha dt"""
    if pd.isna(dt): return None
    if isinstance(dt, pd.Timestamp):
        dt = dt.date()
    return dt - datetime.timedelta(days=dt.weekday())

def procesar_datos(df_consumo, df_jerarquia, df_servicios):
    """
    Limpia y une los datos.
    Calcula consumo promedio y riesgo.
    """
    if df_consumo.empty or df_jerarquia.empty or df_servicios.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    # Asegurar que las columnas de fecha sean datetime antes de cualquier operación
    for col in ["FechaUso", "FechaSincronizacion", "FechaVencimiento"]:
        if col in df_consumo.columns:
            df_consumo[col] = pd.to_datetime(df_consumo[col], errors='coerce')

    # 1. Unir con Jerarquía y Servicios
    df_full = df_consumo.merge(df_jerarquia, on=["Casino", "CC"], how="left")
    df_full = df_full.merge(df_servicios, left_on="IdServicio", right_on="ServicioId", how="left", suffixes=("", "_servicio"))
    
    # 2. Calcular consumo histórico por semanas calendario completas
    df_usados = df_full[df_full["Usado"] == "SI"].copy()

    # Solo días hábiles (Lunes=0 … Viernes=4); excluir fines de semana
    df_usados = df_usados[df_usados["FechaUso"].dt.dayofweek < 5]

    hoy = datetime.date.today()
    inicio_esta_semana = get_start_of_week(hoy)
    
    # Asignar semana de uso
    df_usados['SemanaUso'] = df_usados['FechaUso'].apply(get_start_of_week)
    
    # Filtrar últimas 12 semanas (excluyendo la semana actual en curso que está incompleta)
    hace_12_semanas = inicio_esta_semana - datetime.timedelta(weeks=12)
    hace_4_semanas = inicio_esta_semana - datetime.timedelta(weeks=4)
    
    # Consumos últimas 12 semanas cerradas
    df_12w = df_usados[(df_usados['SemanaUso'] >= hace_12_semanas) & (df_usados['SemanaUso'] < inicio_esta_semana)]
    
    # Agrupar consumo 12 semanas
    agg_12w = df_12w.groupby(["RazonSocial", "RutContratista", "CC", "Casino", "NombreServicio"]).size().reset_index(name="Consumo_12w")
    agg_12w["Promedio_12w"] = agg_12w["Consumo_12w"] / 12.0
    
    # Consumos últimas 4 semanas cerradas
    df_4w = df_usados[(df_usados['SemanaUso'] >= hace_4_semanas) & (df_usados['SemanaUso'] < inicio_esta_semana)]
    agg_4w = df_4w.groupby(["RazonSocial", "RutContratista", "CC", "Casino", "NombreServicio"]).size().reset_index(name="Consumo_4w")
    agg_4w["Promedio_Semanal"] = agg_4w["Consumo_4w"] / 4.0
    
    # 3. Calcular Saldo Disponible
    df_disponibles = df_full[df_full["Saldo Pendiente"] == "Disponible"].copy()
    
    saldo_actual = df_disponibles.groupby(["RazonSocial", "RutContratista", "CC", "Casino", "NombreServicio"]).size().reset_index(name="Saldo_Disponible")
    
    # 4. Consolidar el resumen por cliente/servicio
    resumen = pd.DataFrame()
    # Empezamos con los clientes que tienen saldo disponible o han consumido en las ultimas 12 semanas
    keys = ["RazonSocial", "RutContratista", "CC", "Casino", "NombreServicio"]
    
    all_keys = pd.concat([agg_12w[keys], saldo_actual[keys]]).drop_duplicates()
    
    resumen = all_keys.merge(agg_12w, on=keys, how="left")
    resumen = resumen.merge(agg_4w, on=keys, how="left")
    resumen = resumen.merge(saldo_actual, on=keys, how="left")
    
    # Rellenar NaNs
    resumen["Consumo_12w"] = resumen["Consumo_12w"].fillna(0)
    resumen["Promedio_12w"] = resumen["Promedio_12w"].fillna(0)
    resumen["Consumo_4w"] = resumen["Consumo_4w"].fillna(0)
    resumen["Promedio_Semanal"] = resumen["Promedio_Semanal"].fillna(0) # Este es el de 4 semanas
    resumen["Saldo_Disponible"] = resumen["Saldo_Disponible"].fillna(0)
    
    # 5. Calcular Semanas Cubiertas
    # Para evitar división por 0, si Promedio_Semanal es 0 pero hay saldo, cobertura es "infinita"
    # Eliminar contratistas sin actividad (saldo cero y sin consumo reciente)
    resumen = resumen[(resumen["Saldo_Disponible"] > 0) | (resumen["Promedio_Semanal"] > 0)].copy()

    resumen["Semanas_Cubiertas"] = np.where(
        resumen["Promedio_Semanal"] > 0,
        resumen["Saldo_Disponible"] / resumen["Promedio_Semanal"],
        999.0  # Sin consumo reciente no hay riesgo, independiente del saldo
    )
    
    # 6. Clasificación de Riesgo
    conditions = [
        resumen["Semanas_Cubiertas"] < RIESGO_CRITICO_SEMANAS,
        (resumen["Semanas_Cubiertas"] >= RIESGO_CRITICO_SEMANAS) & (resumen["Semanas_Cubiertas"] < RIESGO_ALTO_SEMANAS),
        (resumen["Semanas_Cubiertas"] >= RIESGO_ALTO_SEMANAS) & (resumen["Semanas_Cubiertas"] < RIESGO_MEDIO_SEMANAS),
        resumen["Semanas_Cubiertas"] >= RIESGO_MEDIO_SEMANAS
    ]
    choices = ["Crítico", "Alto", "Medio", "Bajo"]
    resumen["Riesgo"] = np.select(conditions, choices, default="Bajo")
    
    # 7. Análisis de Tendencia
    # Creciente si Promedio_4w > Promedio_12w + 10%
    # Decreciente si Promedio_4w < Promedio_12w - 10%
    # Estable si está en el rango
    resumen["Tendencia"] = np.where(
        resumen["Promedio_Semanal"] > resumen["Promedio_12w"] * 1.1, "Creciente",
        np.where(
            resumen["Promedio_Semanal"] < resumen["Promedio_12w"] * 0.9, "Decreciente", "Estable"
        )
    )
    
    # 8. Unir con Jerarquía para filtros adicionales (GOP, JOP, Area)
    # Como agrupamos por Casino y CC, la relación debería ser 1:1, sacamos la primera coincidencia
    jerarquia_unique = df_jerarquia.drop_duplicates(subset=["Casino", "CC"])
    resumen = resumen.merge(jerarquia_unique[["Casino", "CC", "Area", "GOP", "JOP"]], on=["Casino", "CC"], how="left")
    
    # Recomendación Comercial
    resumen["Recomendacion_Comercial"] = np.where(
        resumen["Riesgo"] == "Crítico", "Contactar cliente para reposición inmediata.",
        np.where(resumen["Riesgo"] == "Alto", "Monitorear e iniciar gestión comercial preventiva.", "Sin acción inmediata.")
    )

    # 9. Análisis de Vencimientos
    # Si ya viene FechaVencimiento desde SQL, la usamos; si no, la calculamos
    if "FechaVencimiento" not in df_disponibles.columns or df_disponibles["FechaVencimiento"].isna().all():
        df_disponibles["FechaVencimiento"] = df_disponibles["FechaSincronizacion"] + pd.Timedelta(days=100)
    else:
        df_disponibles["FechaVencimiento"] = pd.to_datetime(df_disponibles["FechaVencimiento"], errors='coerce')

    hoy_ts = pd.Timestamp(hoy)
    df_disponibles["DiasParaVencer"] = (df_disponibles["FechaVencimiento"] - hoy_ts).dt.days
    
    vencimientos = df_disponibles[df_disponibles["DiasParaVencer"] <= 30].copy()
    
    vencimientos["Rango_Vencimiento"] = np.where(
        vencimientos["DiasParaVencer"] <= 7, "7 días",
        np.where(vencimientos["DiasParaVencer"] <= 15, "15 días", "30 días")
    )

    return resumen, vencimientos, df_full

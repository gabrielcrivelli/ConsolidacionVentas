# core_consolidacion.py
import os
import re
from datetime import datetime

import pandas as pd

COLUMNA_CANTIDAD = "Cantidad"
COLUMNAS_DESCRIPTIVAS = ["Marca", "Descripcion", "Departamento", "SubFamilia", "Familia"]

# Prioridades base (se pueden sobreescribir desde la GUI)
PRIORIDAD_DEPARTAMENTOS_DEFAULT = {
    "BEBIDAS SIN ALCOHOL": 100,
    "ADITIVOS PARA LAVADOS": 100,
    "ALMACEN": 100,
    "ACEITES": 90,
    "BEBIDAS": 50,
    "LIMPIEZA Y CUIDADO PERSONAL": 30,
    "LIMPIEZA Y CUIDADO": 30,
    "DESAYUNO": 40,
    "ARROZ": 30,
    "ENLATADOS": 30,
    "ALIM VARIOS": 30,
}

CATEGORIAS_ESPECIALES = ["ELECTRO", "ELECTRODOMESTICOS", "FERRETERIA", "RODADOS"]

MESES_ES = [
    "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
    "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"
]
MAPA_MES = {m: i + 1 for i, m in enumerate(MESES_ES)}


def normalizar_mes(mes_str: str) -> str:
    mes_up = mes_str.strip().upper()
    for m in MESES_ES:
        if m in mes_up:
            return m
    raise ValueError(f"No se pudo normalizar mes desde: {mes_str}")


def parsear_nombre_archivo(nombre: str):
    """
    Devuelve (mes_str, anio_int, sucursal_str_o_None) o None si no se puede parsear.
    Espera algo tipo: '3. MARZO 2025 CORRIENTES.xlsx'
    """
    base = os.path.splitext(os.path.basename(nombre))[0]
    base_up = base.upper().replace("  ", " ")

    # Buscar sucursal conocida
    sucursal = None
    for s in ["HIPER", "CORRIENTES"]:
        if s in base_up:
            sucursal = s
            base_up = base_up.replace(s, "").strip()
            break

    # Mes y año
    m = re.search(r"(\d{1,2})\s*\.?\s+([A-ZÁÉÍÓÚÑ ]+)\s+(\d{4})", base_up)
    if not m:
        return None

    _, mes_txt, anio_txt = m.groups()
    try:
        mes_norm = normalizar_mes(mes_txt)
    except ValueError:
        return None

    anio = int(anio_txt)
    return mes_norm, anio, sucursal


def consolidar_datos(archivos_info, prioridades_depto=None):
    """
    archivos_info: lista de diccionarios:
        {
          "ruta": str,
          "mes": "MARZO",
          "anio": 2025,
          "sucursal": "HIPER"
        }
    prioridades_depto: dict opcional para sobreescribir PRIORIDAD_DEPARTAMENTOS_DEFAULT
    """
    if prioridades_depto is None:
        prioridades = PRIORIDAD_DEPARTAMENTOS_DEFAULT.copy()
    else:
        prioridades = prioridades_depto

    todos = []

    for info in archivos_info:
        ruta = info["ruta"]
        mes = info["mes"]
        anio = info["anio"]
        sucursal = info["sucursal"]

        if not os.path.exists(ruta):
            continue

        df = pd.read_excel(ruta, sheet_name=0)

        if COLUMNA_CANTIDAD not in df.columns or "IdArticulo" not in df.columns:
            continue

        columnas_existentes = [c for c in COLUMNAS_DESCRIPTIVAS if c in df.columns]
        columnas_a_usar = ["IdArticulo"] + columnas_existentes + [COLUMNA_CANTIDAD]

        df_f = df[columnas_a_usar].copy()
        df_f["MES"] = f"{mes} {anio}"
        df_f["SUCURSAL"] = sucursal

        # Limpieza básica
        for col in COLUMNAS_DESCRIPTIVAS:
            if col in df_f.columns:
                df_f[col] = df_f[col].astype(str).str.strip()

        # Normalización de Departamento (se puede extender)
        if "Departamento" in df_f.columns:
            df_f["Departamento"] = df_f["Departamento"].str.upper()
            df_f.loc[df_f["Departamento"] == "ACEITES", "Departamento"] = "ALMACEN"
            df_f.loc[df_f["Departamento"] == "HIGIENE PERSONAL", "Departamento"] = "LIMPIEZA Y CUIDADO"

        todos.append(df_f)

    if not todos:
        raise ValueError("No se pudo leer ningún archivo válido.")

    df = pd.concat(todos, ignore_index=True)

    # Consolidar por prioridad de departamento
    df["PRIORIDAD"] = df["Departamento"].map(prioridades).fillna(0)
    df["CLAVE_PRODUCTO"] = (
        df["IdArticulo"].astype(str)
        + "|" + df.get("Marca", "").astype(str)
        + "|" + df.get("Descripcion", "").astype(str)
        + "|" + df.get("SubFamilia", "").astype(str)
        + "|" + df.get("Familia", "").astype(str)
    )

    idx_max = df.groupby("CLAVE_PRODUCTO")["PRIORIDAD"].idxmax()
    dept_final = df.loc[idx_max, ["CLAVE_PRODUCTO", "Departamento"]].drop_duplicates("CLAVE_PRODUCTO")

    df = df.drop(columns=["Departamento"])
    df = df.merge(
        dept_final.rename(columns={"Departamento": "Departamento"}),
        on="CLAVE_PRODUCTO",
        how="left"
    )
    df = df.drop(columns=["PRIORIDAD", "CLAVE_PRODUCTO"])

    # Agrupar final
    df = df.groupby(
        ["IdArticulo", "Marca", "Descripcion", "Departamento", "SubFamilia", "Familia", "MES", "SUCURSAL"],
        as_index=False
    )[COLUMNA_CANTIDAD].sum()

    # Redondeo
    df[COLUMNA_CANTIDAD] = df[COLUMNA_CANTIDAD].apply(
        lambda x: int(x + 0.5) if x >= 0 else int(x - 0.5)
    )

    return df


def _orden_mes_clave(mes_ano: str):
    # mes_ano: "MARZO 2025"
    partes = mes_ano.split()
    if len(partes) != 2:
        return (9999, 99)
    mes_txt, anio_txt = partes
    anio = int(anio_txt)
    mes_num = MAPA_MES.get(mes_txt, 99)
    return (anio, mes_num)


def generar_reportes(df, ruta_salida):
    """
    Genera:
    - Consolidado
    - Ranking de Ventas
    - Por Sucursal
    - Matriz
    - Evolución Mensual
    - Categorías Especiales
    """
    meses_ordenados = sorted(df["MES"].unique(), key=_orden_mes_clave)
    sucursales = sorted(df["SUCURSAL"].dropna().unique())

    with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
        # 1) Consolidado
        df_temp = df.copy()
        df_temp["MES_SUC"] = df_temp["MES"] + "_" + df_temp["SUCURSAL"]
        idx_cols = ["IdArticulo", "Marca", "Descripcion", "Departamento", "SubFamilia", "Familia"]

        df_pivot = df_temp.pivot_table(
            index=idx_cols,
            columns="MES_SUC",
            values=COLUMNA_CANTIDAD,
            aggfunc="sum",
            fill_value=0
        ).reset_index()

        # Asegurar ints
        for col in df_pivot.columns:
            if col not in idx_cols:
                df_pivot[col] = df_pivot[col].astype(int)

        cols_finales = idx_cols.copy()

        # Columnas por mes (sumando sucursales)
        for mes in meses_ordenados:
            cols_mes = [c for c in df_pivot.columns if c.startswith(mes + "_")]
            if not cols_mes:
                continue
            df_pivot[mes] = df_pivot[cols_mes].sum(axis=1).astype(int)
            cols_finales.append(mes)

        # Totales por sucursal
        total_cols = []
        for suc in sucursales:
            cols_suc = [c for c in df_pivot.columns if c.endswith("_" + suc)]
            if not cols_suc:
                continue
            col_total = f"TOTAL {suc.upper()}"
            df_pivot[col_total] = df_pivot[cols_suc].sum(axis=1).astype(int)
            total_cols.append(col_total)

        # Total Consolidado
        if total_cols:
            df_pivot["TOTAL CONSOLIDADO"] = df_pivot[total_cols].sum(axis=1).astype(int)
            cols_finales.extend(total_cols + ["TOTAL CONSOLIDADO"])

        df_final = df_pivot[cols_finales].sort_values("IdArticulo").reset_index(drop=True)
        df_final.to_excel(writer, sheet_name="Consolidado", index=False)

        # 2) Ranking de Ventas
        ranking = df.groupby(["IdArticulo", "Marca", "Descripcion"]).agg(
            {COLUMNA_CANTIDAD: "sum"}
        ).reset_index()
        ranking = ranking.sort_values(COLUMNA_CANTIDAD, ascending=False).reset_index(drop=True)
        ranking.rename(columns={COLUMNA_CANTIDAD: "Total Vendido"}, inplace=True)
        ranking.to_excel(writer, sheet_name="Ranking de Ventas", index=False)

        # 3) Por Sucursal
        por_suc = df.pivot_table(
            index="SUCURSAL",
            columns="MES",
            values=COLUMNA_CANTIDAD,
            aggfunc="sum",
            fill_value=0
        ).reset_index()
        cols_tot = [m for m in meses_ordenados if m in por_suc.columns]
        if cols_tot:
            por_suc["TOTAL"] = por_suc[cols_tot].sum(axis=1)
        por_suc.to_excel(writer, sheet_name="Por Sucursal", index=False)

        # 4) Matriz
        matriz = df.pivot_table(
            index="Departamento",
            columns="SUCURSAL",
            values=COLUMNA_CANTIDAD,
            aggfunc="sum",
            fill_value=0
        )
        matriz["TOTAL"] = matriz.sum(axis=1)
        matriz.sort_values("TOTAL", ascending=False).to_excel(writer, sheet_name="Matriz")

        # 5) Evolución Mensual
        evol = df.pivot_table(
            index="Departamento",
            columns="MES",
            values=COLUMNA_CANTIDAD,
            aggfunc="sum",
            fill_value=0
        )
        cols_evol = [m for m in meses_ordenados if m in evol.columns]
        evol = evol[cols_evol]
        evol.to_excel(writer, sheet_name="Evolución Mensual")

        # 6) Categorías Especiales
        df_espec = df[df["Departamento"].str.upper().isin(CATEGORIAS_ESPECIALES)].copy()
        if not df_espec.empty:
            tmp = df_espec.copy()
            tmp["MES_SUC"] = tmp["MES"] + "_" + tmp["SUCURSAL"]
            idx_cols_espec = ["IdArticulo", "Marca", "Descripcion", "Departamento", "SubFamilia", "Familia"]
            piv = tmp.pivot_table(
                index=idx_cols_espec,
                columns="MES_SUC",
                values=COLUMNA_CANTIDAD,
                aggfunc="sum",
                fill_value=0
            ).reset_index()

            for col in piv.columns:
                if col not in idx_cols_espec:
                    piv[col] = piv[col].astype(int)

            cols_espec_final = idx_cols_espec.copy()
            for mes in meses_ordenados:
                cols_mes = [c for c in piv.columns if c.startswith(mes + "_")]
                if not cols_mes:
                    continue
                piv[mes] = piv[cols_mes].sum(axis=1).astype(int)
                cols_espec_final.append(mes)

            # Totales por sucursal y consolidado
            total_cols_espec = []
            for suc in sucursales:
                cols_suc = [c for c in piv.columns if c.endswith("_" + suc)]
                if not cols_suc:
                    continue
                col_total = f"TOTAL {suc.upper()}"
                piv[col_total] = piv[cols_suc].sum(axis=1).astype(int)
                total_cols_espec.append(col_total)

            if total_cols_espec:
                piv["TOTAL CONSOLIDADO"] = piv[total_cols_espec].sum(axis=1).astype(int)
                cols_espec_final.extend(total_cols_espec + ["TOTAL CONSOLIDADO"])

            piv_final = piv[cols_espec_final].sort_values("IdArticulo").reset_index(drop=True)
            piv_final.to_excel(writer, sheet_name="Categorías Especiales", index=False)

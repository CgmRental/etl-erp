import sys
import os
from datetime import datetime, timedelta
import polars as pl
import hashlib
import numpy as np

def generar_hash_df(df: pl.DataFrame, columnas_excluir=None):
    if columnas_excluir is None:
        columnas_excluir = []

    columnas = [c for c in df.columns if c not in columnas_excluir]

    exprs = []
    for col in columnas:
        dtype = df.schema[col]

        if dtype in (pl.Float32, pl.Float64, pl.Int32, pl.Int64):
            expr = (
                pl.when(pl.col(col).is_null())
                .then(pl.lit(""))
                .otherwise(pl.col(col).cast(pl.Float64).round(2).cast(pl.Utf8))
            )
        elif dtype == pl.Utf8:
            expr = (
                pl.when(pl.col(col).is_null())
                .then(pl.lit(""))
                .otherwise(pl.col(col).str.strip_chars())
            )
        else:
            expr = (
                pl.when(pl.col(col).is_null())
                .then(pl.lit(""))
                .otherwise(pl.col(col).cast(pl.Utf8))
            )

        exprs.append(expr)

    return df.with_columns(
        pl.concat_str(exprs, separator="|").hash().alias("_hash")
    )

# %%
def detectar_cambios_df(
    df_nuevo: pl.DataFrame,
    df_actual: pl.DataFrame,
    key_col,
    columnas_excluir=None
):
    # 🔧 Normalizar key a lista
    if isinstance(key_col, str):
        key_cols = [key_col]
    else:
        key_cols = key_col

    if columnas_excluir is None:
        columnas_excluir = []

    if df_actual.is_empty():
        print(f"Tabla vacía. Todos los {df_nuevo.height} registros son nuevos")
        return {
            "df_insertar": df_nuevo.clone(),
            "df_actualizar": pl.DataFrame(schema=df_nuevo.schema),
            "df_sin_cambios": pl.DataFrame(schema=df_nuevo.schema),
        }

    print("Generando hashes para comparación...")

    df_nuevo_h = generar_hash_df(df_nuevo, columnas_excluir)
    df_actual_h = generar_hash_df(df_actual, columnas_excluir)

    # 🔗 Join por key (soporta múltiples columnas)
    df_join = df_nuevo_h.join(
        df_actual_h.select([*key_cols, "_hash"]),
        on=key_cols,
        how="left",
        suffix="_actual"
    )

    # 🧠 Clasificación
    df_insertar = df_join.filter(pl.col("_hash_actual").is_null())
    df_comunes = df_join.filter(pl.col("_hash_actual").is_not_null())

    df_actualizar = df_comunes.filter(pl.col("_hash") != pl.col("_hash_actual"))
    df_sin_cambios = df_comunes.filter(pl.col("_hash") == pl.col("_hash_actual"))

    # 🧹 limpiar columnas auxiliares
    def limpiar(df):
        return df.select([c for c in df.columns if c not in ["_hash", "_hash_actual"]])

    df_insertar = limpiar(df_insertar)
    df_actualizar = limpiar(df_actualizar)
    df_sin_cambios = limpiar(df_sin_cambios)

    print("Análisis completado:")
    print(f"• Nuevos: {df_insertar.height}")
    print(f"• Con cambios: {df_actualizar.height}")
    print(f"• Sin cambios: {df_sin_cambios.height}")

    return {
        "df_insertar": df_insertar,
        "df_actualizar": df_actualizar,
        "df_sin_cambios": df_sin_cambios,
    }
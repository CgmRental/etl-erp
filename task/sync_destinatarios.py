import os
import polars as pl
from db.conexion_db import crear_conexion, cerrar_conexion
from db.querys_db import consult_data
from cache import set_destinatarios, get_destinatarios

RUTA_BASE = r'C:\Users\CGM\Projects\Proyecto CGMExPress\data'
FORMATO = ".xlsx"

CENTROS = ['C154', 'C200', 'C040', 'C080']

RENAME_COLUMNS = {
    "Documento": "id_entrega",
    "Destinatario mcía.": "codigo_destinatario",
    "Descripción destinatario de mercancías": "descripcion",
}

def cargar_excel_cabezera():
    dfs = []
    for centro in CENTROS:
        path = os.path.join(RUTA_BASE, f"MONITOR_CABEZERA{centro}{FORMATO}")
        dfs.append(pl.read_excel(path, engine="openpyxl"))
    return pl.concat(dfs)

def sync_destinatarios():
    print("▶ Ejecutando sync_destinatarios")

    # 1️⃣ Leer Excel
    df_excel = (
        cargar_excel_cabezera()
        .rename(RENAME_COLUMNS)
        .select("codigo_destinatario", "descripcion")
        .unique()
    )

    # 2️⃣ Obtener destinatarios DB (cache primero)
    df_destinatario_db = get_destinatarios()

    if df_destinatario_db is None:
        conn = crear_conexion("db_almacen_distribucion")
        data = consult_data(
            conn,
            "SELECT id_destinatario, codigo_destinatario FROM tbl_destinatarios"
        )
        cerrar_conexion(conn)

        df_destinatario_db = pl.DataFrame(data)
        set_destinatarios(df_destinatario_db)

    # 3️⃣ Anti-join (nuevos destinatarios)
    df_nuevos = df_excel.join(
        df_destinatario_db.select("codigo_destinatario"),
        on="codigo_destinatario",
        how="anti"
    )

    if df_nuevos.is_empty():
        print("✔ No hay nuevos destinatarios")
        return

    # 4️⃣ Generar IDs
    last_id = df_destinatario_db.select(
        pl.col("id_destinatario").max()
    ).item()

    df_nuevos = (
        df_nuevos
        .with_row_index("id_destinatario", offset=last_id + 1)
        .with_columns(pl.col("id_destinatario").cast(pl.Int64))
    )

    # 5️⃣ Actualizar cache
    df_destinatario_db = pl.concat([
        df_destinatario_db,
        df_nuevos.select("id_destinatario", "codigo_destinatario")
    ])

    set_destinatarios(df_destinatario_db)

    print(f"✔ Nuevos destinatarios agregados: {df_nuevos.height}")

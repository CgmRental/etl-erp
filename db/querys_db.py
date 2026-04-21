from db.conexion_db import cerrar_conexion
from datetime import datetime
import numpy as np
import polars as pl


def insert_data_polars(connection, tabla, datos, cabezeras, batch_size):
    try:
        cursor = connection.cursor()

        datos_final = datos.select(cabezeras)
        
        num_registros = datos_final.height # Polars usa .height, no .shape[0] (aunque shape funciona, height es más idiomatico)

        columnas = ", ".join([f"`{col}`" for col in cabezeras])
        placeholders = ", ".join(["%s"] * len(cabezeras))
        
        query = f"INSERT INTO {tabla} ({columnas}) VALUES ({placeholders})"
        
        # Iterar usando slice (offset, longitud)
        for i in range(0, num_registros, batch_size):
            # Polars: slice(inicio, cantidad)
            bloque = datos_final.slice(i, batch_size)

            # ESTA ES LA CLAVE: .iter_rows() 
            # Convierte nulos de Polars a None de Python automáticamente
            # Es más rápido y seguro que convertir a numpy
            valores = list(bloque.iter_rows())
            
            cursor.executemany(query, valores)

        connection.commit()
        print(f"✅ Se insertaron: {num_registros} registros en la tabla {tabla} (Polars)")
        
    except Exception as e:
        connection.rollback()
        print(f"❌ Error al insertar datos: {e}")
        raise
    # Recuerda: No cerrar conexión aquí si la usas fuera


def consult_data(connection, query):

    try:
        cursor = connection.cursor(dictionary=True)
        cursor.execute(query)
        return cursor.fetchall()
    except Exception as e:
        print(f"Error al obtener datos: {e}")
        return []

def insert_data_polars_ids(connection, tabla, datos, cabezeras):
    try:
        cursor = connection.cursor()

        datos_final = datos.select(cabezeras)

        num_registros = datos_final.height

        columnas = ", ".join([f"`{col}`" for col in cabezeras])
        placeholders = ", ".join(["%s"] * len(cabezeras))
        query = f"INSERT INTO {tabla} ({columnas}) VALUES ({placeholders})"

        ids = []

        # Insertar fila por fila y capturar ID
        for fila in datos_final.iter_rows():
            cursor.execute(query, fila)
            ids.append(cursor.lastrowid)

        connection.commit()

        # 🔥 Agregar columna id al DataFrame original
        df_resultado = datos_final.with_columns(
            pl.Series("id", ids)
        )

        print(f"✅ Se insertaron: {num_registros} registros en la tabla {tabla}")

        return df_resultado

    except Exception as e:
        connection.rollback()
        print(f"❌ Error al insertar datos: {e}")
        raise

def update_data_polars(
    connection,
    tabla: str,
    df: pl.DataFrame,
    cabezeras: list,
    batch_size: int,
    columna_id: str = "equipo"
):
    try:
        cursor = connection.cursor()

        # 🔹 Validación de columnas
        faltantes = [c for c in cabezeras if c not in df.columns]
        if faltantes:
            print(f"⚠️ Columnas faltantes: {faltantes}")

        # 🔹 columnas a actualizar (excepto ID)
        columnas_update = [
            col for col in cabezeras
            if col in df.columns and col != columna_id
        ]

        # 🔹 SQL dinámico
        set_clause = ", ".join([f"`{col}` = %s" for col in columnas_update])
        query = f"UPDATE {tabla} SET {set_clause} WHERE `{columna_id}` = %s"

        registros_actualizados = 0
        num_registros = df.height

        # 🔹 batch processing con Polars (sin pandas)
        for i in range(0, num_registros, batch_size):
            bloque = df.slice(i, batch_size)

            valores = [
                tuple(
                    fila.get(col, None) for col in columnas_update
                ) + (fila[columna_id],)
                for fila in bloque.to_dicts()
            ]

            cursor.executemany(query, valores)
            registros_actualizados += len(valores)

            # commit parcial cada 5 lotes
            if i > 0 and i % (batch_size * 5) == 0:
                connection.commit()
                print(f"🔄 Actualizados {registros_actualizados}/{num_registros}")

        connection.commit()
        print(f"✅ Se actualizaron {num_registros} registros en {tabla}")

    except Exception as e:
        connection.rollback()
        print(f"❌ Error al actualizar: {e}")
        raise

#############################################################PANDAS####################################################################w

def update_data(connection, tabla, datos, cabezeras, batch_size, columna_id='equipo'):
    try:
        cursor = connection.cursor()
        
        datos_copy = datos.copy()
        
        cols = [col for col in datos_copy.columns]
        datos_copy = datos_copy[cols]
        
        num_registros = datos_copy.shape[0]
        
        # Reemplazar NaN por None (para compatibilidad con MySQL)
        datos_copy = datos_copy.replace({np.nan: None})
        
        # --- 🔍 VERIFICAR COLUMNAS DISPONIBLES ---
        faltantes = [col for col in cabezeras if col not in datos_copy.columns]
        if faltantes:
            print(f" Columnas en cabezeras no encontradas en los datos: {faltantes}")
        
        columnas_update = [
            col for col in cabezeras
            if col in datos_copy.columns and col != columna_id and col != 'fechahora'
        ]
        
        columnas_update = ['fecha_modificacion'] + columnas_update
        
        set_clause = ", ".join([f"`{col}` = %s" for col in columnas_update])
        query = f"UPDATE {tabla} SET {set_clause} WHERE `{columna_id}` = %s"
        
        registros_actualizados = 0
        for i in range(0, len(datos_copy), batch_size):
            bloque = datos_copy.iloc[i:i + batch_size]
            
            valores = []
            for _, fila in bloque.iterrows():
                valores_update = [fila.get(col, None) for col in columnas_update]
                # Añadir el valor del ID al final (para el WHERE)
                valores_update.append(fila[columna_id])
                valores.append(tuple(valores_update))
            
            # Ejecutar el bloque
            cursor.executemany(query, valores)
            registros_actualizados += len(valores)
            
            # Commit parcial cada 5 lotes
            if i > 0 and i % (batch_size * 5) == 0:
                connection.commit()
                print(f" Actualizados {registros_actualizados}/{num_registros} registros...")
        
        connection.commit()
        print(f" Se actualizaron {num_registros} registros en la tabla {tabla}")
        
    except Exception as e:
        connection.rollback()
        print(f" Error al actualizar datos: {e}")
        raise
    finally:
        cerrar_conexion(connection)


def insert_data_pandas(connection, tabla, datos, cabezeras, batch_size):

    try:
        cursor = connection.cursor()

        datos_copy = datos.copy()

        cols = [col for col in datos_copy.columns] 

        num_registros = datos_copy.shape[0]

        datos_copy = datos_copy[cols]
        
        datos_copy = datos_copy.replace({np.nan: None})
        
        columnas = ", ".join([f"`{columna}`" for columna in cabezeras])
        placeholders = ", ".join(["%s"] * len(datos_copy.columns))
        
        # Crear la consulta de inserción
        query = f"INSERT INTO {tabla} ({columnas}) VALUES ({placeholders})"
        
        # Iterar sobre el DataFrame en bloques
        for i in range(0, len(datos_copy), batch_size):
            # Obtener el bloque de datos
            bloque = datos_copy.iloc[i:i+batch_size]

            # Convertir el DataFrame del bloque en una lista de tuplas
            valores = [tuple(fila) for fila in bloque.values]
            
            # Ejecutar la inserción del bloque
            cursor.executemany(query, valores)

        connection.commit()
        print(f" Se insertaron: {num_registros} registros en la tabla {tabla}")
    except Exception as e:
        connection.rollback()
        print(f" Error al insertar datos: {e}")
        raise
    finally:
        cerrar_conexion(connection)
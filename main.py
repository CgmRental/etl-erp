from sap.conexion_sap import connect_to_sap
from db.conexion_db import crear_conexion
from sap.scripts_sap import execute_entregas_ewm,execute_entregas_ewm_v2,execute_guias_remision,execute_materialesIH09
from db.conexion_db import crear_conexion, cerrar_conexion
from db.querys_db import insert_data_polars, consult_data, insert_data_polars_ids, update_data_polars
from funciones import detectar_cambios_df

from datetime import datetime, timedelta
import subprocess
import polars as pl
import pandas as pd
import numpy as np
import time
import os
import json

def main():
    
    # Conexión a SAP
       # Nombre de la conexión en SAP GUI
    def get_db_config(path="config/credenciales_sap.json"):
        ###Carga la configuración de la base de datos desde un archivo JSON.####
        with open(path, "r") as config_file:
            return json.load(config_file)
        
    config = get_db_config()    
    connection_name = config["sap_connection_name"]
    user_sap = config["sap_user"]
    pwa_sap = config["sap_password"]
    formato = ".xlsx"
    ruta_archivos_base = r'C:\Users\CGM\Projects\Proyecto CGMExPress\data'

    parametros_materiales = {
        "material": "",
        "txt_breve_mat": "",
        "t_material": "",
        "laboratorio": "",
        "denom_estand": "",
        "documento": "",
        "n_ant_mat": "",
        "grp_articulo": "",
        "codi_ean": "",
        "pb_mand": "",
        "creado_por": "",
        "modificado": "",
        "ult_modif": "",
        "layout": "/0_DB_MAT2",
    }
    nombre_archivo_IH09 = "materiales_IH09"
    parametros_entregas_ewm = {
        "almacen": "C154",
        "monitor": "SAP",
        "nodo": ""
    }
    
    nombre_archivo_entregas_ewm_cabezera = "MONITOR_CABEZERA"
    nombre_archivo_entregas_ewm_detalle = "MONITOR_DETALLE"
    
    nombre_archivo_entregas_ewm_cabezera_pendiente = "MONITOR_CABEZERA_PENDIENTE"
    nombre_archivo_entregas_ewm_detalle_pendiente = "MONITOR_DETALLE_PENDIENTE"
    nombre_archivo_guias_remision = "guia_remision"
    
    centros = pd.DataFrame({'centro': ['C200', 'C040', 'C080']})
    centros_pendiente = pl.DataFrame({'centro': ['C154','C200', 'C040', 'C080']})
    
    fecha_sistema = (datetime.now() - timedelta(days=0)).strftime("%d.%m.%Y")
    fecha_db = (datetime.now() - timedelta(days=0)).strftime("%Y-%m-%d")
    batch_size = 1000

    print("Fecha del sistema:", fecha_sistema)
    # Conectar a SAP
    
    session = connect_to_sap(connection_name=connection_name,user_sap=user_sap,pwa_sap=pwa_sap)
    if session:
        print(f"Conexión a {connection_name,user_sap,pwa_sap} establecida exitosamente.")

        subprocess.run("echo off | clip", shell=True)

        execute_guias_remision(session,nombre_archivo_guias_remision,formato,ruta_archivos_base,fecha_sistema)
        
        execute_entregas_ewm(session, parametros_entregas_ewm, nombre_archivo_entregas_ewm_cabezera,nombre_archivo_entregas_ewm_detalle, formato, ruta_archivos_base,fecha_sistema)
        for _,centro in centros.iterrows():
            cen = centro['centro']
            execute_entregas_ewm_v2(session,nombre_archivo_entregas_ewm_cabezera,nombre_archivo_entregas_ewm_detalle,formato,ruta_archivos_base,fecha_sistema,cen)
        
        connection = crear_conexion("db_almacen_distribucion")
        entrega_pendientedb = consult_data(connection, "SELECT id_entrega,oficina_expedicion FROM tbl_entrega_pendientes_ewm")
        cerrar_conexion(connection)

        dfentrega_pendiente = pd.DataFrame(entrega_pendientedb)

        dfentrega_pendiente["oficina_expedicion"] = (
        dfentrega_pendiente["oficina_expedicion"]
        .str.replace("UCL-", "", regex=False))

        fecha = ""

        for oficina, df_grupo in dfentrega_pendiente.groupby("oficina_expedicion"):
            print(f"Oficina: {oficina}")
            print(df_grupo[["id_entrega"]])
            subprocess.run("echo off | clip", shell=True)
            df_grupo[["id_entrega"]].to_clipboard(index=False,header=False)

            execute_entregas_ewm_v2(session,nombre_archivo_entregas_ewm_cabezera_pendiente,nombre_archivo_entregas_ewm_detalle_pendiente,formato,ruta_archivos_base,fecha,oficina)
    else:
        print(f"No se pudo establecer conexión con {connection_name,user_sap,pwa_sap}.")
        session.findById("wnd[0]").close()

    time.sleep(5)
      
    session.findById("wnd[0]").close()
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
    
    types_columns_guias_ewm = {
                                "Status": pl.String,
                                "Sociedad": pl.String, 
                                "DocumSAP": pl.Int64,
                                "Ejercicio": pl.String,
                                "TipDocumento": pl.String,
                                "Número SUNAT": pl.String,
                                "Estado del Proceso": pl.String,
                                "Fecha de Creación": pl.String,
                                "Hora Creación": pl.String,
                                "Usuario Creación": pl.String,
                                "Fecha emisión": pl.String,
                                "Nombres y/o Razón Social del Adquiriente": pl.String,
                                "N° CDR": pl.String,
                                "Fecha Recepción": pl.String,
                                "Hora Recepción": pl.String,
                                "Fecha Respuesta": pl.String, 
                                "Hora Respuesta": pl.String, 
                                "Mensaje": pl.String,
                                "Observación": pl.String}


    excel_guias = pl.read_excel(os.path.join(ruta_archivos_base,nombre_archivo_guias_remision+formato),engine="calamine", schema_overrides=types_columns_guias_ewm)


    df_guias = excel_guias.select([ "DocumSAP", "Número SUNAT", "Fecha de Creación", "Hora Creación", "Usuario Creación", "Fecha emisión"])


    rename_columns_guias = {"DocumSAP": "entrega", 
                        "Número SUNAT": "guia_sap", 
                        "Fecha de Creación": "created_at_sap",
                        "Hora Creación": "created_to_the_sap",
                        "Usuario Creación": "usuario", 
                        "Fecha emisión": "fecha_emision", }
    df_guias = df_guias.rename(rename_columns_guias)


    df_guias = df_guias.with_columns(
        pl.col("created_at_sap")
        .str.strptime(pl.Datetime, format="%Y-%m-%d %H:%M:%S", strict=False),

        pl.col("created_to_the_sap")
        .str.strptime(pl.Datetime, format="%Y-%m-%d %H:%M:%S", strict=False),

        pl.col("fecha_emision")
        .str.strptime(pl.Datetime, format="%Y-%m-%d %H:%M:%S", strict=False)
    )

    df_guias = df_guias.with_columns(
        pl.col("created_at_sap").dt.date().alias("created_at_sap"),
        pl.col("created_to_the_sap").dt.time().alias("created_to_the_sap"),
        pl.col("fecha_emision").dt.date().alias("fecha_emision")
    )


    df_guias = df_guias.with_columns(
        (
            pl.col("created_at_sap").cast(pl.Datetime) +
            pl.col("created_to_the_sap").cast(pl.Duration)
        ).alias("created_at_sap")
    )


    df_guias_unicas = df_guias.sort(
        ["entrega", "created_at_sap"]
    ).unique(
        subset=["entrega"], 
        keep="last"
    )


    df_guias_unicas = df_guias_unicas.select(["entrega", "guia_sap","usuario", "fecha_emision", "created_at_sap"])


    consulta_guias_sap = f"SELECT entrega,guia_sap,usuario,fecha_emision, created_at_sap FROM tbl_guia_remision_zosfe004 where created = '{fecha_db}'"


    connection = crear_conexion("db_almacen_distribucion")
    guiasap_db = consult_data(connection, consulta_guias_sap)
    cerrar_conexion(connection)
    df_guiasap_db = pl.DataFrame(guiasap_db)


    resultados_guias = detectar_cambios_df(df_guias_unicas, df_guiasap_db, key_col="entrega")
    df_insertar_guias = resultados_guias['df_insertar']
    df_actualizar_guias = resultados_guias['df_actualizar']
    df_sin_cambios_guias = resultados_guias['df_sin_cambios'] 


    if not df_insertar_guias.is_empty():
        cabecera_guia = ["entrega", "guia_sap", "usuario", "fecha_emision", "created_at_sap", "created"]
        df_insertar_guias = df_insertar_guias.with_columns(pl.lit(datetime.now()).alias("created"))
        connection = crear_conexion("db_almacen_distribucion")
        df_insertar_ids = insert_data_polars_ids(connection, "tbl_guia_remision_zosfe004", df_insertar_guias, cabecera_guia)
        df_insertar_ids = df_insertar_ids.rename({"id": "id_guia_remision", "entrega": "id_entrega"})
        cerrar_conexion(connection)
    else:
        df_insertar_ids = pl.DataFrame({
        "id_entrega": pl.Series([], dtype=pl.Int64),  # o el tipo que corresponda
        "id_guia_remision": pl.Series([], dtype=pl.String)
    })
        print("No hay registros nuevos para insertar.")

    if not df_actualizar_guias.is_empty():
        cabecera_guia = ["entrega", "guia_sap", "usuario", "fecha_emision", "created_at_sap","updated"]
        df_actualizar_guias = df_actualizar_guias.with_columns(pl.lit(datetime.now()).alias("updated"))
        connection = crear_conexion("db_almacen_distribucion")
        update_data_polars(connection, "tbl_guia_remision_zosfe004", df_actualizar_guias, cabecera_guia,batch_size, columna_id="entrega")
        cerrar_conexion(connection)
    else:
        print("No hay registros nuevos para actualizar.")


    types_columns_entregas_ewm = {
                                    "Documento": pl.Int64,
        "Destinatario mcía.": pl.String,
        "Descripción destinatario de mercancías": pl.String,
        "Clase de documento": pl.String,
        "Oficina expedición": pl.String,
        "Status salida de mercancías": pl.String,
        "Concluido": pl.String,
        "Status de picking": pl.String,
        "Autor": pl.String,
        "Cantidad pos.": pl.Int64,
        "Cantidad de unidades de manipulación": pl.Int64,
        "Cantidad de productos": pl.Int64,
        "Fecha de entrega (planificada)": pl.String,
        "Creados el": pl.String,
        "Creados a la(s)": pl.String,}

    rename_columns_entregas_ewm_cabezera = {
        "Documento": "id_entrega",
        "Destinatario mcía.": "codigo_cliente",
        "Descripción destinatario de mercancías": "descripcion",
        "Clase de documento": "clase_documento",
        "Oficina expedición": "oficina_expedicion",
        "Status salida de mercancías": "status_salida_mercancia",
        "Concluido": "concluido",
        "Status de picking": "status_picking",
        "Autor": "autor",
        "Cantidad pos.": "cantidad_posiciones",
        "Cantidad de unidades de manipulación": "cantidad_unidades_manipuladas",
        "Cantidad de productos": "cantidad_productos",
        "Fecha de entrega (planificada)": "Fecha_entrega_planif",
        "Creados el": "created_at_sap",
        "Creados a la(s)": "created_to_the_sap",
    }

    excel_entregas_ewm_cabezeraC154 = pl.read_excel(os.path.join(ruta_archivos_base,nombre_archivo_entregas_ewm_cabezera+formato),engine="calamine", schema_overrides=types_columns_entregas_ewm)
    excel_entregas_ewm_cabezeraC200 = pl.read_excel(os.path.join(ruta_archivos_base,nombre_archivo_entregas_ewm_cabezera+centros["centro"][0]+formato),engine="calamine", schema_overrides=types_columns_entregas_ewm)
    excel_entregas_ewm_cabezeraC040 = pl.read_excel(os.path.join(ruta_archivos_base,nombre_archivo_entregas_ewm_cabezera+centros["centro"][1]+formato),engine="calamine", schema_overrides=types_columns_entregas_ewm)
    excel_entregas_ewm_cabezeraC080 = pl.read_excel(os.path.join(ruta_archivos_base,nombre_archivo_entregas_ewm_cabezera+centros["centro"][2]+formato),engine="calamine", schema_overrides=types_columns_entregas_ewm)

    excel_entregas_ewm_cabezera_pendienteC154 = pl.read_excel(os.path.join(ruta_archivos_base,nombre_archivo_entregas_ewm_cabezera_pendiente+centros_pendiente["centro"][0]+formato),engine="calamine", schema_overrides=types_columns_entregas_ewm)
    excel_entregas_ewm_cabezera_pendienteC200 = pl.read_excel(os.path.join(ruta_archivos_base,nombre_archivo_entregas_ewm_cabezera_pendiente+centros_pendiente["centro"][1]+formato),engine="calamine", schema_overrides=types_columns_entregas_ewm)
    excel_entregas_ewm_cabezera_pendienteC040 = pl.read_excel(os.path.join(ruta_archivos_base,nombre_archivo_entregas_ewm_cabezera_pendiente+centros_pendiente["centro"][2]+formato),engine="calamine", schema_overrides=types_columns_entregas_ewm)
    excel_entregas_ewm_cabezera_pendienteC080 = pl.read_excel(os.path.join(ruta_archivos_base,nombre_archivo_entregas_ewm_cabezera_pendiente+centros_pendiente["centro"][3]+formato),engine="calamine", schema_overrides=types_columns_entregas_ewm)

    df_excel_entregas_ewm_cabezera = pl.concat([excel_entregas_ewm_cabezeraC154, excel_entregas_ewm_cabezeraC200, excel_entregas_ewm_cabezeraC040, excel_entregas_ewm_cabezeraC080, excel_entregas_ewm_cabezera_pendienteC154, excel_entregas_ewm_cabezera_pendienteC200, excel_entregas_ewm_cabezera_pendienteC040, excel_entregas_ewm_cabezera_pendienteC080])
    df_excel_entregas_ewm_cabezera = df_excel_entregas_ewm_cabezera.rename(rename_columns_entregas_ewm_cabezera)

    df_excel_entregas_ewm_cabezera = df_excel_entregas_ewm_cabezera.unique(subset=["id_entrega"], keep="last")

    df_excel_entregas_ewm_cabezera = df_excel_entregas_ewm_cabezera.with_columns(
        pl.col("created_at_sap")
        .str.strptime(pl.Datetime, format="%Y-%m-%d %H:%M:%S", strict=False),

        pl.col("created_to_the_sap")
        .str.strptime(pl.Datetime, format="%Y-%m-%d %H:%M:%S", strict=False),

        pl.col("Fecha_entrega_planif")
        .str.strptime(pl.Datetime, format="%Y-%m-%d %H:%M:%S", strict=False)
    )

    df_excel_entregas_ewm_cabezera = df_excel_entregas_ewm_cabezera.with_columns(
        pl.col("created_at_sap").dt.date().alias("created_at_sap"),
        pl.col("created_to_the_sap").dt.time().alias("created_to_the_sap"),
        pl.col("Fecha_entrega_planif").dt.date().alias("Fecha_entrega_planif")
    )


    df_excel_entregas_ewm_cabezera = df_excel_entregas_ewm_cabezera.with_columns(
        (
            pl.col("created_at_sap").cast(pl.Datetime) +
            pl.col("created_to_the_sap").cast(pl.Duration)
        ).alias("created_at_sap")
    )


    connection = crear_conexion("db_almacen_distribucion")
    cliente_db = consult_data(connection, "SELECT id_cliente,codigo_cliente,descripcion FROM tbl_cliente")
    cerrar_conexion(connection)
    df_cliente_db = pl.DataFrame(cliente_db)

    df_cliente_excel = df_excel_entregas_ewm_cabezera.select(["codigo_cliente","descripcion"]).unique()

    resultados_clientes = detectar_cambios_df(df_cliente_excel, df_cliente_db, key_col="codigo_cliente", columnas_excluir=["id_cliente"])
    df_insertar_clientes = resultados_clientes['df_insertar']
    df_actualizar_clientes = resultados_clientes['df_actualizar']    
    df_sin_cambios_clientes = resultados_clientes['df_sin_cambios'] 

    if not df_insertar_clientes.is_empty():
        cabezera_clientes = ["codigo_cliente","descripcion","created"]
        con = crear_conexion("db_almacen_distribucion")
        df_insertar_clientes = df_insertar_clientes.with_columns(pl.lit(datetime.now()).alias("created"))
        df_clientes_creados = insert_data_polars_ids(con, "tbl_cliente", df_insertar_clientes, cabezera_clientes)
        df_clientes_creados = df_clientes_creados.rename({"id": "id_cliente"})
        df_clientes = pl.concat([
            df_cliente_db.select(["id_cliente","codigo_cliente","descripcion"]),
            df_clientes_creados.select(["id_cliente","codigo_cliente","descripcion"])
        ])
        cerrar_conexion(con)
    else:
        df_clientes = df_cliente_db.select(["id_cliente","codigo_cliente","descripcion"])   
        print('No hay clientes nuevos para insertar.')

    if not df_actualizar_clientes.is_empty():
        cabezera_clientes = ["codigo_cliente","descripcion","updated"]
        df_actualizar_clientes = df_actualizar_clientes.with_columns(pl.lit(datetime.now()).alias("updated"))
        connection = crear_conexion("db_almacen_distribucion")
        update_data_polars(connection, "tbl_cliente", df_actualizar_clientes, cabezera_clientes,batch_size, columna_id="codigo_cliente")
        cerrar_conexion(connection)
    else:    print("No hay clientes nuevos para actualizar.")


    df_entregas_join_cliente = df_excel_entregas_ewm_cabezera.join(df_cliente_db.select(["id_cliente", "codigo_cliente"]), on="codigo_cliente", how="left")
    df_entregas_join_guias = df_entregas_join_cliente.join(df_insertar_ids.select(["id_entrega", "id_guia_remision"]), on=["id_entrega"], how="left")


    df_actualizar_entrega_ewm = df_insertar_ids.select(["id_entrega", "id_guia_remision"]).join(
        df_entregas_join_cliente,
        on = ["id_entrega"],
        how="anti"
    )

    df_entregas_to_mysql_finalizada = df_entregas_join_guias.select(["id_entrega","id_cliente","id_guia_remision","clase_documento","oficina_expedicion","status_salida_mercancia","concluido","status_picking","cantidad_posiciones","cantidad_unidades_manipuladas","cantidad_productos","autor","Fecha_entrega_planif","created_at_sap"]).filter(pl.col("status_salida_mercancia") == "Finalizada")
    df_entregas_to_mysql_pendientes = df_entregas_join_guias.select(["id_entrega","id_cliente","id_guia_remision","clase_documento","oficina_expedicion","status_salida_mercancia","concluido","status_picking","cantidad_posiciones","cantidad_unidades_manipuladas","cantidad_productos","autor","Fecha_entrega_planif","created_at_sap"]).filter(pl.col("status_salida_mercancia") == "No iniciada")

    connection = crear_conexion("db_almacen_distribucion")
    db_entregas=consult_data(connection, f"SELECT id_entrega,status_salida_mercancia,concluido,status_picking,cantidad_posiciones,cantidad_unidades_manipuladas,cantidad_productos FROM tbl_entrega_ewm WHERE created = '{fecha_db}'")
    df_entregas_db = pl.DataFrame(db_entregas)
    cerrar_conexion(connection)

    resultados_entregas = detectar_cambios_df(df_entregas_to_mysql_finalizada,df_entregas_db, key_col="id_entrega", columnas_excluir=["id_cliente","id_guia_remision","clase_documento","oficina_expedicion","autor","Fecha_entrega_planif","created_at_sap"])
    df_insertar_entregas = resultados_entregas['df_insertar']
    df_actualizar_entregas = resultados_entregas['df_actualizar']
    df_sin_cambios_entregas = resultados_entregas['df_sin_cambios']


    if not df_insertar_entregas.is_empty():
        cabezera_entregas = ["id_entrega","id_cliente","id_guia_remision","clase_documento","oficina_expedicion","status_salida_mercancia","concluido",
                            "status_picking","cantidad_posiciones","cantidad_unidades_manipuladas","cantidad_productos","autor",
                            "Fecha_entrega_planif","created_at_sap", "created"]
        df_insertar_entregas = df_insertar_entregas.with_columns(pl.lit(datetime.now()).alias("created"))
        con = crear_conexion("db_almacen_distribucion")
        insert_data_polars(con, "tbl_entrega_ewm", df_insertar_entregas, cabezera_entregas, batch_size)
        cerrar_conexion(con)
    else:
        print("No hay entregas finalizadas para insertar.")

    if not df_actualizar_entregas.is_empty():
        cabezera_entregas = ["id_entrega","id_cliente","id_guia_remision","clase_documento","oficina_expedicion","status_salida_mercancia","concluido",
                            "status_picking","cantidad_posiciones","cantidad_unidades_manipuladas","cantidad_productos","autor",
                            "Fecha_entrega_planif","created_at_sap", "updated"]
        df_actualizar_entregas = df_actualizar_entregas.with_columns(pl.lit(datetime.now()).alias("updated"))
        con = crear_conexion("db_almacen_distribucion")
        update_data_polars(con, "tbl_entrega_ewm", df_actualizar_entregas, cabezera_entregas,batch_size, columna_id="id_entrega")
        cerrar_conexion(con)


    if not df_entregas_to_mysql_pendientes.is_empty():
        cabezera_entregas = ["id_entrega","id_cliente","id_guia_remision","clase_documento","oficina_expedicion","status_salida_mercancia","concluido",
                            "status_picking","cantidad_posiciones","cantidad_unidades_manipuladas","cantidad_productos","autor",
                            "Fecha_entrega_planif","created_at_sap", "created"]
        df_entregas_to_mysql_pendientes = df_entregas_to_mysql_pendientes.with_columns(pl.lit(datetime.now()).alias("created"))
        con = crear_conexion("db_almacen_distribucion")
        consult_data(con, "TRUNCATE TABLE tbl_entrega_pendientes_ewm")
        insert_data_polars(con, "tbl_entrega_pendientes_ewm", df_entregas_to_mysql_pendientes, cabezera_entregas, batch_size)
        cerrar_conexion(con)


    if not df_actualizar_entrega_ewm.is_empty():
        cabecera_actualizar_entrega = ["id_guia_remision", "id_entrega"]
        connection = crear_conexion("db_almacen_distribucion")
        update_data_polars(connection, "tbl_entrega_ewm", df_actualizar_entrega_ewm, cabecera_actualizar_entrega, batch_size, columna_id="id_entrega")
        cerrar_conexion(connection)


    configuracion_tipos_detalle = {"Documento": pl.Int64, 
                        "Número de posición": pl.Int64, 
                        "Pedido": pl.String, 
                        "Pedido clte.": pl.String, 
                        "Número de reserva": pl.String, 
                        "Orden de mantenimiento": pl.String,
                        "Prod.": pl.String, 
                        "Descripción de producto": pl.String,
                        "Ctd.": pl.Decimal(38, 3), 
                        "Unidad de medida": pl.String, 
                        "Lote": pl.String, 
                        "Fecha de picking real": pl.String, 
                        "Hora de picking real": pl.String}   


    rename_columns_entregas_ewm_detalle = {"Documento": "id_entrega", 
                        "Número de posición": "posicion", 
                        "Pedido": "pedido", 
                        "Pedido clte.": "pedido_venta", 
                        "Número de reserva": "reserva", 
                        "Orden de mantenimiento": "orden",
                        "Prod.": "codigo_material",
                        "Descripción de producto": "descripcion",
                        "Ctd.": "cantidad", 
                        "Unidad de medida": "unidad_medida", 
                        "Lote": "lote", 
                        "Fecha de picking real": "fecha_picking", 
                        "Hora de picking real": "hora_picking"}


    excel_entregas_ewm_detalleC154 = pl.read_excel(os.path.join(ruta_archivos_base,nombre_archivo_entregas_ewm_detalle+formato),engine="calamine",schema_overrides=configuracion_tipos_detalle)
    excel_entregas_ewm_detalleC200 = pl.read_excel(os.path.join(ruta_archivos_base,nombre_archivo_entregas_ewm_detalle+centros["centro"][0]+formato),engine="calamine",schema_overrides=configuracion_tipos_detalle)
    excel_entregas_ewm_detalleC040 = pl.read_excel(os.path.join(ruta_archivos_base,nombre_archivo_entregas_ewm_detalle+centros["centro"][1]+formato),engine="calamine",schema_overrides=configuracion_tipos_detalle)
    excel_entregas_ewm_detalleC080 = pl.read_excel(os.path.join(ruta_archivos_base,nombre_archivo_entregas_ewm_detalle+centros["centro"][2]+formato),engine="calamine",schema_overrides=configuracion_tipos_detalle)

    excel_entregas_ewm_detalle_pendienteC154 = pl.read_excel(os.path.join(ruta_archivos_base,nombre_archivo_entregas_ewm_detalle_pendiente+centros_pendiente["centro"][0]+formato),engine="calamine",schema_overrides=configuracion_tipos_detalle)
    excel_entregas_ewm_detalle_pendienteC200 = pl.read_excel(os.path.join(ruta_archivos_base,nombre_archivo_entregas_ewm_detalle_pendiente+centros_pendiente["centro"][1]+formato),engine="calamine",schema_overrides=configuracion_tipos_detalle)
    excel_entregas_ewm_detalle_pendienteC040 = pl.read_excel(os.path.join(ruta_archivos_base,nombre_archivo_entregas_ewm_detalle_pendiente+centros_pendiente["centro"][2]+formato),engine="calamine",schema_overrides=configuracion_tipos_detalle)
    excel_entregas_ewm_detalle_pendienteC080 = pl.read_excel(os.path.join(ruta_archivos_base,nombre_archivo_entregas_ewm_detalle_pendiente+centros_pendiente["centro"][3]+formato),engine="calamine",schema_overrides=configuracion_tipos_detalle)


    df_excel_entregas_ewm_detalle = pl.concat([excel_entregas_ewm_detalleC154, excel_entregas_ewm_detalleC200, excel_entregas_ewm_detalleC040, excel_entregas_ewm_detalleC080, excel_entregas_ewm_detalle_pendienteC154, excel_entregas_ewm_detalle_pendienteC200, excel_entregas_ewm_detalle_pendienteC040, excel_entregas_ewm_detalle_pendienteC080])


    df_excel_entregas_ewm_detalle = df_excel_entregas_ewm_detalle.with_columns(
        
        # Convertir Fecha
        pl.col("Fecha de picking real").str.to_datetime(format="%Y-%m-%d %H:%M:%S", strict=False).dt.date().alias("Fecha de picking real"),
        
        # Convertir Hora
        # Formato común SAP/Excel: "%H:%M:%S" (ej: 14:30:59)
        pl.col("Hora de picking real").str.to_datetime(format="%Y-%m-%d %H:%M:%S", strict=False).dt.time().alias("Hora de picking real")
    ).rename(rename_columns_entregas_ewm_detalle)

    df_excel_entregas_ewm_detalle = df_excel_entregas_ewm_detalle.unique(subset=["id_entrega", "posicion"], keep="last")
    df_excel_entregas_ewm_detalle = df_excel_entregas_ewm_detalle.with_columns(
        (
            pl.col("fecha_picking").cast(pl.Datetime) +
            pl.col("hora_picking").cast(pl.Duration)
        ).alias("fecha_picking")
    )


    df_materiales_unicos = df_excel_entregas_ewm_detalle.select(["codigo_material", "descripcion"]).unique()


    connection = crear_conexion("db_matricial")
    materiales_db = consult_data(connection, "SELECT id_material,codigo_material,descripcion FROM tbl_materiales_ih09")
    cerrar_conexion(connection)
    df_material_db = pl.DataFrame(materiales_db)    


    resultados_materiales = detectar_cambios_df(df_materiales_unicos, df_material_db, key_col="codigo_material", columnas_excluir=["id_material"])
    df_insertar_materiales = resultados_materiales['df_insertar']
    df_actualizar_materiales = resultados_materiales['df_actualizar']
    df_sin_cambios_materiales = resultados_materiales['df_sin_cambios']


    if not df_insertar_materiales.is_empty():
        cabezera_materiales = ['codigo_material','descripcion', 'created']
        connection = crear_conexion("db_matricial")
        df_insertar_materiales = df_insertar_materiales.with_columns(pl.lit(datetime.now()).alias("created"))
        df_materiales_ids = insert_data_polars_ids(connection, "tbl_materiales_ih09", df_insertar_materiales,cabezera_materiales)
        df_materiales_ids = df_materiales_ids.rename({"id": "id_material"})
        df_materiales = pl.concat([df_material_db.select(["id_material","codigo_material","descripcion"])
                                ,df_materiales_ids.select(["id_material","codigo_material","descripcion"])])
        cerrar_conexion(connection)
    else:
        df_materiales = df_material_db.select(["id_material","codigo_material","descripcion"])   
        print("No hay materiales nuevos para insertar.")


    df_excel_entregas_ewm_detalle_join_material = df_excel_entregas_ewm_detalle.join(df_materiales.select(["codigo_material", "id_material"]), on="codigo_material", how="left")


    df_excel_entregas_ewm_detalle_join_material = df_excel_entregas_ewm_detalle_join_material.select(["id_entrega", "posicion", "pedido", "pedido_venta", "reserva", "orden", "id_material", "cantidad", "unidad_medida", "lote", "fecha_picking", "hora_picking"])


    df_excel_entregas_ewm_detalle_finalizada = df_excel_entregas_ewm_detalle_join_material.join(df_entregas_to_mysql_finalizada.select(["id_entrega"]), on="id_entrega", how="inner")


    connection = crear_conexion("db_almacen_distribucion")
    db_detalle_entregas=consult_data(connection, f"SELECT id_detalle_entrega ,id_entrega, posicion, pedido, pedido_venta, reserva, orden, id_material, cantidad, unidad_medida, lote, fecha_picking FROM tbl_detalle_entrega_ewm WHERE created = '{fecha_db}'")
    cerrar_conexion(connection)
    df_detalle_entregas_db = pl.DataFrame(db_detalle_entregas)


    resultados_detalle_entregas = detectar_cambios_df(df_excel_entregas_ewm_detalle_finalizada, df_detalle_entregas_db, key_col=["id_entrega","posicion"], columnas_excluir=["id_detalle_entrega", "pedido", "pedido_venta", "reserva", "orden", "hora_picking"])
    df_insertar_detalle_entregas = resultados_detalle_entregas['df_insertar']
    df_actualizar_detalle_entregas = resultados_detalle_entregas['df_actualizar']
    df_sin_cambios_detalle_entregas = resultados_detalle_entregas['df_sin_cambios']


    if not df_insertar_detalle_entregas.is_empty():
        cabezera_detalle_entregas = ["id_entrega", "posicion", "pedido", "pedido_venta", "reserva", "orden", "id_material", "cantidad", "unidad_medida", "lote", "fecha_picking", "created"]
        df_insertar_detalle_entregas = df_insertar_detalle_entregas.with_columns(pl.lit(datetime.now()).alias("created"))
        connection = crear_conexion("db_almacen_distribucion")
        insert_data_polars(connection, "tbl_detalle_entrega_ewm", df_insertar_detalle_entregas, cabezera_detalle_entregas, batch_size)
        cerrar_conexion(connection)
    else:
        print("No hay detalles de entregas nuevos para insertar.")

    if not df_actualizar_detalle_entregas.is_empty():
        cabezera_detalle_entregas = ["pedido", "pedido_venta", "reserva", "orden", "id_material", "cantidad", "unidad_medida", "lote", "fecha_picking", "updated", "id_entrega", "posicion"]
        df_actualizar_detalle_entregas = df_actualizar_detalle_entregas.join(
            df_detalle_entregas_db.select(["id_detalle_entrega", "id_entrega", "posicion"]),
            on=["id_entrega", "posicion"])
        df_actualizar_detalle_entregas = df_actualizar_detalle_entregas.with_columns(pl.lit(datetime.now()).alias("updated"))
        connection = crear_conexion("db_almacen_distribucion")
        update_data_polars(connection, "tbl_detalle_entrega_ewm", df_actualizar_detalle_entregas, cabezera_detalle_entregas, batch_size, columna_id="id_detalle_entrega")
        cerrar_conexion(connection)

    def cerrar_todos_excel():
        try:
            os.system("taskkill /f /im excel.exe")
            print("Todos los procesos de Excel han sido cerrados.")
        except Exception as e:
            print("Error cerrando Excel:", e)

    cerrar_todos_excel()
    

if __name__ == "__main__":
    main()


    
from sap.conexion_sap import connect_to_sap
from db.conexion_db import crear_conexion
from sap.scripts_sap import execute_entregas_ewm,execute_entregas_ewm_v2,execute_guias_remision,execute_materialesIH09
from datetime import datetime, timedelta

import polars as pl
import pandas as pd
import numpy as np
import time
import os
import json
import hashlib


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
    nombre_archivo_guias_remision = "guia_remision"
    centros = pd.DataFrame({'centro': ['C200', 'C040', 'C080']})
    fecha_sistema = (datetime.now() - timedelta(days=1)).strftime("%d.%m.%Y")
    print("Fecha del sistema:", fecha_sistema)
    # Conectar a SAP
    session = connect_to_sap(connection_name=connection_name,user_sap=user_sap,pwa_sap=pwa_sap)
    if session:
        print(f"Conexión a {connection_name,user_sap,pwa_sap} establecida exitosamente.")

        execute_guias_remision(session,nombre_archivo_guias_remision,formato,ruta_archivos_base,fecha_sistema)
        #execute_materialesIH09(session,parametros_materiales,nombre_archivo_IH09,formato,ruta_archivos_base,fecha_sistema)
        
        if execute_entregas_ewm(session, parametros_entregas_ewm, nombre_archivo_entregas_ewm_cabezera,nombre_archivo_entregas_ewm_detalle, formato, ruta_archivos_base,fecha_sistema):
            for _,centro in centros.iterrows():
                cen = centro['centro']
                execute_entregas_ewm_v2(session,nombre_archivo_entregas_ewm_cabezera,nombre_archivo_entregas_ewm_detalle,formato,ruta_archivos_base,fecha_sistema,cen)

    else:
        print(f"No se pudo establecer conexión con {connection_name,user_sap,pwa_sap}.")
        session.findById("wnd[0]").close()

    time.sleep(5)
      
    session.findById("wnd[0]").close()
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()

    def cerrar_todos_excel():
        try:
            os.system("taskkill /f /im excel.exe")
            print("Todos los procesos de Excel han sido cerrados.")
        except Exception as e:
            print("Error cerrando Excel:", e)

    cerrar_todos_excel()
    

if __name__ == "__main__":
    main()
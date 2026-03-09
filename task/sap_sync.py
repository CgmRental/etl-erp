from sap.conexion_sap import connect_to_sap
from sap.scripts_sap import (
    execute_entregas_ewm,
    execute_entregas_ewm_v2,
    execute_guias_remision
)

import pandas as pd
import time
import os
import json


def get_sap_config(path="config/credenciales_sap.json"):
    with open(path, "r") as config_file:
        return json.load(config_file)


def cerrar_todos_excel():
    try:
        os.system("taskkill /f /im excel.exe")
        print("✔ Excel cerrado")
    except Exception as e:
        print("❌ Error cerrando Excel:", e)


def ejecutar_proceso_sap():
    print("▶ Iniciando proceso SAP")

    config = get_sap_config()

    connection_name = config["sap_connection_name"]
    user_sap = config["sap_user"]
    pwa_sap = config["sap_password"]

    formato = ".xlsx"
    ruta_archivos_base = r'C:\Users\CGM\Projects\Proyecto CGMExPress\data'

    parametros_entregas_ewm = {
        "almacen": "C154",
        "monitor": "SAP",
        "nodo": ""
    }

    nombre_archivo_entregas_ewm_cabezera = "MONITOR_CABEZERA"
    nombre_archivo_entregas_ewm_detalle = "MONITOR_DETALLE"
    nombre_archivo_guias_remision = "guia_remision"

    centros = pd.DataFrame({'centro': ['C200', 'C040', 'C080']})

    fecha_sistema = time.strftime("%d.%m.%Y")

    # 🔌 Conectar a SAP
    session = connect_to_sap(
        connection_name=connection_name,
        user_sap=user_sap,
        pwa_sap=pwa_sap
    )

    if not session:
        print("❌ No se pudo conectar a SAP")
        return

    try:
        print("✔ Conectado a SAP")

        execute_guias_remision(
            session,
            nombre_archivo_guias_remision,
            formato,
            ruta_archivos_base,
            fecha_sistema
        )

        if execute_entregas_ewm(
            session,
            parametros_entregas_ewm,
            nombre_archivo_entregas_ewm_cabezera,
            nombre_archivo_entregas_ewm_detalle,
            formato,
            ruta_archivos_base,
            fecha_sistema
        ):
            for _, centro in centros.iterrows():
                execute_entregas_ewm_v2(
                    session,
                    nombre_archivo_entregas_ewm_cabezera,
                    nombre_archivo_entregas_ewm_detalle,
                    formato,
                    ruta_archivos_base,
                    fecha_sistema,
                    centro["centro"]
                )

    except Exception as e:
        print("❌ Error en proceso SAP:", e)

    finally:
        time.sleep(5)
        session.findById("wnd[0]").close()
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
        cerrar_todos_excel()
        print("✔ Proceso SAP finalizado")

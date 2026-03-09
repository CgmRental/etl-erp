import pandas as pd
import numpy as np
import os
import subprocess

def execute_inventario_mb52(session, parametros,nombre_archivo,formato,ruta):
    """
    Ejecuta la transacción MB52 y llena los campos requeridos.
    :param session: Sesión activa de SAP.
    :param parametros: Diccionario con los valores para los campos necesarios.
    """
    try:
        subprocess.run("echo off | clip", shell=True)
        # Maximizar la ventana
        session.findById("wnd[0]").maximize()

        # Ingresar la transacción MB52
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nMB52"
        session.findById("wnd[0]").sendVKey(0)

        # Configuración de opciones de selección
        session.findById("wnd[0]/usr/chkPA_SOND").selected = True
        session.findById("wnd[0]/usr/chkNEGATIV").selected = False
        session.findById("wnd[0]/usr/chkXMCHB").selected = False
        session.findById("wnd[0]/usr/chkNOZERO").selected = True
        session.findById("wnd[0]/usr/chkNOVALUES").selected = False

        # Llenar los campos de selección
        session.findById("wnd[0]/usr/ctxtMATNR-LOW").text = parametros.get("material", "")
        session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = parametros.get("centro", "")
        session.findById("wnd[0]/usr/ctxtLGORT-LOW").text = parametros.get("almacen", "")
        session.findById("wnd[0]/usr/ctxtCHARG-LOW").text = parametros.get("lote", "")
        session.findById("wnd[0]/usr/ctxtMATART-LOW").text = parametros.get("tipo_material", "")
        session.findById("wnd[0]/usr/ctxtMATKLA-LOW").text = parametros.get("clase_material", "")
        session.findById("wnd[0]/usr/ctxtEKGRUP-LOW").text = parametros.get("grupo_compras", "")
        session.findById("wnd[0]/usr/ctxtSO_SOBKZ-LOW").text = parametros.get("clave_especial", "")

        # Seleccionar opciones avanzadas
        session.findById("wnd[0]/usr/radPA_FLT").setFocus()
        session.findById("wnd[0]/usr/radPA_FLT").select()

        # Asignar variante de parámetros
        session.findById("wnd[0]/usr/ctxtP_VARI").text = parametros.get("variante", "")
        session.findById("wnd[0]/usr/ctxtP_VARI").setFocus()
        session.findById("wnd[0]/usr/ctxtP_VARI").caretPosition = len(parametros.get("variante", ""))

        # Guardar el archivo
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[1]").select()
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/txtGS_EXPORT-FILE_NAME").text = nombre_archivo
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/txtGS_EXPORT-FILE_NAME").caretPosition = len(nombre_archivo)
        session.findById("wnd[1]/tbar[0]/btn[20]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = ruta
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nombre_archivo+formato
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = len(nombre_archivo+formato)
        session.findById("wnd[1]/tbar[0]/btn[11]").press()

        #Finalizar Transaccion
        session.EndTransaction()

        print("Transacción MB52 configurada correctamente.")
    except Exception as e:
        print(f"Error al ejecutar MB52: {e}")

def execute_flota(session, parametros,nombre_archivo,formato,ruta):
    try:
        subprocess.run("echo off | clip", shell=True)
        # Maximizar la ventana
        session.findById("wnd[0]").maximize()

        # Ingresar la transacción ZPMRI1
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nZPMRI0001"
        session.findById("wnd[0]").sendVKey(0)

        session.findById("wnd[0]/usr/chkP_CHECK").selected = False

        # Llenar los campos de selección
        session.findById("wnd[0]/usr/ctxtP_BUKRS").text = parametros.get("sociedad", "")
        session.findById("wnd[0]/usr/ctxtS_IWERK-LOW").text = parametros.get("centro_trabajo", "")
        session.findById("wnd[0]/usr/ctxtS_INGRP-LOW").text = parametros.get("grupo_planificacio", "")
        session.findById("wnd[0]/usr/ctxtS_STORT-LOW").text = parametros.get("almacen", "")
        session.findById("wnd[0]/usr/ctxtS_EQUNR-LOW").text = parametros.get("numero_equipo", "")
        session.findById("wnd[0]/usr/ctxtS_EQTYP-LOW").text = parametros.get("tipo_equipo", "")
        session.findById("wnd[0]/usr/ctxtS_BEBER-LOW").text = parametros.get("responsable", "")
        session.findById("wnd[0]/usr/ctxtS_PROYE-LOW").text = parametros.get("proyecto", "")
        session.findById("wnd[0]/usr/ctxtS_CUSTO-LOW").text = parametros.get("cliente", "")
        session.findById("wnd[0]/usr/ctxtS_KLART-LOW").text = parametros.get("tipo_clase", "")
        session.findById("wnd[0]/usr/ctxtS_CLASS-LOW").text = parametros.get("clase", "")

        # Seleccionar opciones avanzadas
        session.findById("wnd[0]/usr/chkP_CHECK").setFocus()
        session.findById("wnd[0]/usr/txt%_S_IWERK_%_APP_%-TEXT").setFocus()
        session.findById("wnd[0]/usr/txt%_S_IWERK_%_APP_%-TEXT").caretPosition = 21

        #Ejecutar y Seleccionar layaout
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[33]").press()
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").currentCellRow = 5
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectedRows = "5"
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").clickCurrentCell()
        
        # Guardar el archivo
        session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/txtGS_EXPORT-FILE_NAME").text = nombre_archivo
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/txtGS_EXPORT-FILE_NAME").caretPosition = len(nombre_archivo)
        session.findById("wnd[1]/tbar[0]/btn[20]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = ruta
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nombre_archivo+formato
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = len(nombre_archivo+formato)
        session.findById("wnd[1]/tbar[0]/btn[11]").press()

        #Finalizar Transaccion
        session.EndTransaction()

        print("Transacción ZPMRI1 configurada correctamente")
    except Exception as e:
        print(f'Error al ejecutar ZPMRI1: {e}')

def execute_reservas(session, parametros,nombre_archivo,formato,ruta):
    try:
        subprocess.run("echo off | clip", shell=True)
        # Maximizar la ventana
        session.findById("wnd[0]").maximize()

        # Ingresar la transacción MB25
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/nMB25"

        session.findById("wnd[0]").sendVKey(0)

        # Llenar los campos de selección
        session.findById("wnd[0]/usr/ctxtMATNR-LOW").text = parametros.get("material", "")
        session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = parametros.get("centro", "")
        session.findById("wnd[0]/usr/ctxtBDTER-LOW").text = parametros.get("fecha_necesidad", "")
        session.findById("wnd[0]/usr/txtUSNAM-LOW").text = parametros.get("usuario", "")
        session.findById("wnd[0]/usr/txtWEMPF-LOW").text = parametros.get("destinatario", "")

        # Seleccionar opciones avanzadas
        session.findById("wnd[0]/usr/chkP_NO_ACC").setFocus()
        session.findById("wnd[0]/usr/chkP_NO_ACC").selected = False

        # Llenar los campos de selección
        session.findById("wnd[0]/usr/ctxtKOSTL-LOW").text = parametros.get("centro_coste", "")
        session.findById("wnd[0]/usr/ctxtAUFNR-LOW").text = parametros.get("orden", "")
        session.findById("wnd[0]/usr/ctxtPOSID-LOW").text = parametros.get("elemento_pep", "")
        session.findById("wnd[0]/usr/ctxtNPLNR-LOW").text = parametros.get("grafo", "")
        session.findById("wnd[0]/usr/txtP_VORNR").text = parametros.get("operacion", "")
        session.findById("wnd[0]/usr/ctxtANLN1-LOW").text = parametros.get("activo_fijo", "")
        session.findById("wnd[0]/usr/ctxtANLN2-LOW").text = parametros.get("subnumero", "")
        session.findById("wnd[0]/usr/ctxtUMWRK-LOW").text = parametros.get("centro_receptor", "")
        session.findById("wnd[0]/usr/ctxtUMLGO-LOW").text = parametros.get("almacen_receptor", "")
        session.findById("wnd[0]/usr/ctxtKDAUF-LOW").text = parametros.get("pedido_cliente", "")
        session.findById("wnd[0]/usr/txtKDPOS-LOW").text = parametros.get("posicion", "")
        session.findById("wnd[0]/usr/txtKDEIN-LOW").text = parametros.get("reparto", "")

        # Seleccionar opciones avanzadas
        session.findById("wnd[0]/usr/chkP_OPEN").selected = True
        session.findById("wnd[0]/usr/chkP_CANCEL").selected = False
        session.findById("wnd[0]/usr/chkP_CLOSED").selected = True
        session.findById("wnd[0]/usr/chkP_ISSUES").selected = True
        session.findById("wnd[0]/usr/chkP_RECEIP").selected = True

        # Llenar los campos de selección
        session.findById("wnd[0]/usr/ctxtALV_DEF").text = parametros.get("layout", "")
        session.findById("wnd[0]/usr/ctxtALV_DEF").setFocus()
        session.findById("wnd[0]/usr/ctxtALV_DEF").caretPosition = len(parametros.get("layout", ""))

        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[1]").select()

        # Guardar el archivo
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/txtGS_EXPORT-FILE_NAME").text = nombre_archivo
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/txtGS_EXPORT-FILE_NAME").caretPosition = len(nombre_archivo)
        session.findById("wnd[1]/tbar[0]/btn[20]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = ruta
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nombre_archivo+formato
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = len(nombre_archivo+formato)
        session.findById("wnd[1]/tbar[0]/btn[11]").press()

        #Finalizar Transaccion
        session.EndTransaction()

        print("Transacción MB25 configurada correctamente")
    except Exception as e:
        print(f'Error al ejecutar MB25: {e}')

def execute_avisorden(session, parametros,nombre_archivo,formato,ruta):
    try:

        # Maximizar la ventana
        session.findById("wnd[0]").maximize()

         # Ingresar la transacción ZPMRI0002
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nZPMRI0002"
        session.findById("wnd[0]").sendVKey(0)

        # Seleccionar opciones avanzadas
        session.findById("wnd[0]/usr/chkP_PEND").selected = True
        session.findById("wnd[0]/usr/chkP_PROC").selected = True
        session.findById("wnd[0]/usr/chkP_CONC").selected = True
        session.findById("wnd[0]/usr/chkP_OPEND").selected = True
        session.findById("wnd[0]/usr/chkP_OPROC").selected = True
        session.findById("wnd[0]/usr/chkP_OCONC").selected = False

        # Llenar los campos de selección
        session.findById("wnd[0]/usr/ctxtS_SWERK-LOW").text = parametros.get("centro", "")
        session.findById("wnd[0]/usr/ctxtS_INGRP-LOW").text = parametros.get("grupo", "")

        session.findById("wnd[0]/usr/btn%_S_QMNUM_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()

        session.findById("wnd[0]/usr/ctxtS_QMART-LOW").text = parametros.get("clase_aviso", "")
        session.findById("wnd[0]/usr/ctxtS_AUSWK-LOW").text = parametros.get("repercusion", "")
        session.findById("wnd[0]/usr/ctxtS_STORT-LOW").text = parametros.get("emplazamiento", "")
        session.findById("wnd[0]/usr/ctxtS_ZONA-LOW").text = parametros.get("zona", "")
        session.findById("wnd[0]/usr/ctxtS_EQUNR-LOW").text = parametros.get("equipo", "")
        session.findById("wnd[0]/usr/ctxtS_BEBER-LOW").text = parametros.get("area_empresa", "")
        session.findById("wnd[0]/usr/ctxtS_PARNP-LOW").text = parametros.get("proyecto", "")
        session.findById("wnd[0]/usr/ctxtS_PARNC-LOW").text = parametros.get("cliente", "")
        session.findById("wnd[0]/usr/ctxtS_EQART-LOW").text = parametros.get("clase_objeto", "")
        session.findById("wnd[0]/usr/ctxtS_AUART-LOW").text = parametros.get("clase_orden", "")
        session.findById("wnd[0]/usr/ctxtS_MPRES-LOW").text = parametros.get("supervisor", "")

        #Ejecutar y Seleccionar layaout
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[33]").press()
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").currentCellRow = 0
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").firstVisibleRow = 0
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectedRows = "0"
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").clickCurrentCell()
                
        session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
        
        # Guardar el archivo
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/txtGS_EXPORT-FILE_NAME").text = nombre_archivo
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/txtGS_EXPORT-FILE_NAME").caretPosition = len(nombre_archivo)
        session.findById("wnd[1]/tbar[0]/btn[20]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = ruta
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nombre_archivo+formato
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = len(nombre_archivo+formato)
        session.findById("wnd[1]/tbar[0]/btn[11]").press()

        #Finalizar Transaccion
        session.EndTransaction()

        print("Transacción ZPMRI0002 configurada correctamente")
    except Exception as e:
        print(f'Error al ejecutar ZPMRI0002: {e}')

def execute_ordenes(session, parametros,nombre_archivo,formato,ruta):
    try:

        # Maximizar la ventana
        session.findById("wnd[0]").maximize()

        # Ingresar la transacción IW39
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nIW39"
        session.findById("wnd[0]").sendVKey(0)

        #Status Orden
        session.findById("wnd[0]/usr/chkDY_OFN").selected = True
        session.findById("wnd[0]/usr/chkDY_IAR").selected = True
        session.findById("wnd[0]/usr/chkDY_MAB").selected = True
        session.findById("wnd[0]/usr/chkDY_HIS").selected = True

        session.findById("wnd[0]/usr/ctxtSELSCHEM").text = parametros.get("esq_selec", "")
        
        #Seleccion de Ordenes
        session.findById("wnd[0]/usr/btn%_AUFNR_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()

        session.findById("wnd[0]/usr/ctxtAUART-LOW").text = parametros.get("clase_orden", "")
        session.findById("wnd[0]/usr/ctxtSTRNO-LOW").text = parametros.get("ubicacion", "")
        session.findById("wnd[0]/usr/ctxtEQUNR-LOW").text = parametros.get("equipo", "")
        session.findById("wnd[0]/usr/ctxtSERMAT-LOW").text = parametros.get("material", "")
        session.findById("wnd[0]/usr/txtSERIALNR-LOW").text = parametros.get("num_serie", "")
        session.findById("wnd[0]/usr/txtDEVICEID-LOW").text = parametros.get("dat_adic_disp", "")
        session.findById("wnd[0]/usr/ctxtQMNUM-LOW").text = parametros.get("aviso", "")
        session.findById("wnd[0]/usr/ctxtGEWRK-LOW").text = parametros.get("responsable", "")
        session.findById("wnd[0]/usr/txtVAWRK-LOW").text = parametros.get("trabajo", "")
        session.findById("wnd[0]/usr/ctxtDATUV").text = parametros.get("periodo_ini", "")
        session.findById("wnd[0]/usr/ctxtDATUB").text = parametros.get("periodo_fin", "")
        #session.findById("wnd[0]/usr/cmbDY_PARVW").key = parametros.get("interl", "")
        session.findById("wnd[0]/usr/ctxtDY_PARNR").text = parametros.get("interl", "")
        session.findById("wnd[0]/usr/ctxtWAERS").text = parametros.get("moneda", "")

        #Seleccion de Orden de Servicio Mantenimiento
        session.findById("wnd[0]/usr/ctxtSRVDOCID-LOW").text = parametros.get("documento_servicio", "")
        session.findById("wnd[0]/usr/txtS_ITEMID-LOW").text = parametros.get("posicion_servicio", "")
        session.findById("wnd[0]/usr/ctxtSRVPROD-LOW").text = parametros.get("producto_servicio", "")
        session.findById("wnd[0]/usr/ctxtS_SPART-LOW").text = parametros.get("solicitante", "")
        session.findById("wnd[0]/usr/ctxtS_SLSORG-LOW").text = parametros.get("org_ventas", "")
        session.findById("wnd[0]/usr/ctxtS_DSCHNL-LOW").text = parametros.get("canal_distr", "")
        session.findById("wnd[0]/usr/ctxtSRV_DIV-LOW").text = parametros.get("particion", "")

        #Datos Generales/datos de gestion
        session.findById("wnd[0]/usr/chkDY_OBL").selected = False

        session.findById("wnd[0]/usr/ctxtLEAD_AUF-LOW").text = parametros.get("orden_prin", "")
        session.findById("wnd[0]/usr/ctxtMAUFNR-LOW").text = parametros.get("orden_superior", "")
        session.findById("wnd[0]/usr/ctxtIWERK-LOW").text = parametros.get("centro_planif", "")
        session.findById("wnd[0]/usr/ctxtPRIOK-LOW").text = parametros.get("prioridad", "")
        session.findById("wnd[0]/usr/ctxtERNAM-LOW").text = parametros.get("autor", "")
        session.findById("wnd[0]/usr/ctxtERDAT-LOW").text = parametros.get("fecha_creac", "")
        session.findById("wnd[0]/usr/ctxtSTAI1-LOW").text = parametros.get("status_incl", "")
        session.findById("wnd[0]/usr/ctxtSTAE1-LOW").text = parametros.get("status_excl", "")
        session.findById("wnd[0]/usr/txtKTEXT-LOW").text = parametros.get("txt_breve", "")
        session.findById("wnd[0]/usr/ctxtAENAM-LOW").text = parametros.get("modif", "")
        session.findById("wnd[0]/usr/ctxtAEDAT-LOW").text = parametros.get("fech_modif", "")
        session.findById("wnd[0]/usr/ctxtANLBD-LOW").text = parametros.get("disponible", "")
        session.findById("wnd[0]/usr/ctxtGSTRP-LOW").text = parametros.get("fecha_ini_extr", "")
        session.findById("wnd[0]/usr/ctxtGLTRP-LOW").text = parametros.get("fecha_fin_extr", "")
        session.findById("wnd[0]/usr/ctxtWARPL-LOW").text = parametros.get("plan_mant_preve", "")
        session.findById("wnd[0]/usr/ctxtWAPOS-LOW").text = parametros.get("posic_mant", "")
        session.findById("wnd[0]/usr/ctxtREVNR-LOW").text = parametros.get("revision", "")
        session.findById("wnd[0]/usr/ctxtPSPEL-LOW").text = parametros.get("elem_pep_cab", "")
        session.findById("wnd[0]/usr/ctxtKOSTV-LOW").text = parametros.get("ceco", "")
        session.findById("wnd[0]/usr/ctxtGLTRS-LOW").text = parametros.get("fin_progr", "")
        session.findById("wnd[0]/usr/ctxtFTRMI-LOW").text = parametros.get("fecha_lib_real", "")
        session.findById("wnd[0]/usr/ctxtGETRI-LOW").text = parametros.get("fecha_fin_real", "")
        session.findById("wnd[0]/usr/ctxtGSTRI-LOW").text = parametros.get("fecha_ini_real", "")
        session.findById("wnd[0]/usr/ctxtGSTRS-LOW").text = parametros.get("ini_progr", "")
        session.findById("wnd[0]/usr/ctxtANLVD-LOW").text = parametros.get("disponible_dsd_fecha", "")
        session.findById("wnd[0]/usr/ctxtLACD_DAT-LOW").text = parametros.get("fecha_venc_fin", "")
        session.findById("wnd[0]/usr/ctxtPLNNR-LOW").text = parametros.get("grp_hoja_ruta", "")
        session.findById("wnd[0]/usr/txtPLNAL-LOW").text = parametros.get("cont_hoja_ruta", "")
        session.findById("wnd[0]/usr/ctxtPLKNZ-LOW").text = parametros.get("ind_planif_orden", "")
        session.findById("wnd[0]/usr/ctxtBAUTL-LOW").text = parametros.get("conjunto", "")

        #Datos emplazamiento/imputacion
        session.findById("wnd[0]/usr/ctxtSWERK-LOW").text = parametros.get("centro_empl", "")
        session.findById("wnd[0]/usr/ctxtILART-LOW").text = parametros.get("clase_act", "")
        session.findById("wnd[0]/usr/ctxtARBPL-LOW").text = parametros.get("puesto_trab", "")
        session.findById("wnd[0]/usr/ctxtKOSTL-LOW").text = parametros.get("centro_coste", "")
        session.findById("wnd[0]/usr/ctxtPROID-LOW").text = parametros.get("elemento_pep", "")
        session.findById("wnd[0]/usr/ctxtAUFNT-LOW").text = parametros.get("subgrafo", "")
        session.findById("wnd[0]/usr/txtVORUE-LOW").text = parametros.get("operacion_sup", "")
        session.findById("wnd[0]/usr/ctxtADPSP-LOW").text = parametros.get("elemento_refe", "")
        session.findById("wnd[0]/usr/ctxtKDAUF-LOW").text = parametros.get("pedido_clt", "")
        session.findById("wnd[0]/usr/txtKDPOS-LOW").text = parametros.get("pos_ped_clt", "")
        session.findById("wnd[0]/usr/ctxtVKORG-LOW").text = parametros.get("org_ventas", "")
        session.findById("wnd[0]/usr/ctxtVTWEG-LOW").text = parametros.get("canal_distr", "")
        session.findById("wnd[0]/usr/ctxtSPART-LOW").text = parametros.get("sector", "")
        session.findById("wnd[0]/usr/ctxtGSBER-LOW").text = parametros.get("division", "")
        session.findById("wnd[0]/usr/ctxtBUKRS-LOW").text = parametros.get("sociedad", "")
        session.findById("wnd[0]/usr/ctxtANLNR-LOW").text = parametros.get("activo_fijo", "")
        session.findById("wnd[0]/usr/ctxtBEBER-LOW").text = parametros.get("area_emp", "")
        session.findById("wnd[0]/usr/ctxtSTORT-LOW").text = parametros.get("emplazamiento", "")
        session.findById("wnd[0]/usr/txtEQFNR-LOW").text = parametros.get("campo_clasf", "")
        session.findById("wnd[0]/usr/ctxtABCKZ-LOW").text = parametros.get("ind_abc", "")
        session.findById("wnd[0]/usr/ctxtINGPR-LOW").text = parametros.get("grp_mant_ord", "")
        session.findById("wnd[0]/usr/txtMSGRP-LOW").text = parametros.get("local", "")
        session.findById("wnd[0]/usr/txtAUFPL-LOW").text = parametros.get("n_hoja_ruta", "")
        session.findById("wnd[0]/usr/ctxtPLGRP-LOW").text = parametros.get("grp_mant_hruta", "")
        session.findById("wnd[0]/usr/ctxtKUNUM-LOW").text = parametros.get("cliente", "")
        session.findById("wnd[0]/usr/txtGESIST-LOW").text = parametros.get("total_real", "")
        session.findById("wnd[0]/usr/txtGESPLN-LOW").text = parametros.get("total_plan", "")
        session.findById("wnd[0]/usr/ctxtSOGEN-LOW").text = parametros.get("permiso", "")
        session.findById("wnd[0]/usr/ctxtPRCTR-LOW").text = parametros.get("centro_bn", "")
        session.findById("wnd[0]/usr/ctxtPAGESTAT-LOW").text = parametros.get("status_paging", "")
        session.findById("wnd[0]/usr/ctxtVARIANT").text = parametros.get("layout", "")
        session.findById("wnd[0]/usr/ctxtVARIANT").caretPosition =len(parametros.get("layout", ""))
        session.findById("wnd[0]/usr/ctxtMONITOR").text = parametros.get("campo_ref_monit", "")

        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        
        # Guardar el archivo
        session.findById("wnd[0]/tbar[1]/btn[16]").press()
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/txtGS_EXPORT-FILE_NAME").text = nombre_archivo
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/txtGS_EXPORT-FILE_NAME").caretPosition = len(nombre_archivo)
        session.findById("wnd[1]/tbar[0]/btn[20]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = ruta
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nombre_archivo+formato
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = len(nombre_archivo+formato)
        session.findById("wnd[1]/tbar[0]/btn[11]").press()

        #Finalizar Transaccion
        session.EndTransaction()

        print("Transacción IW39 configurada correctamente")
    except Exception as e:
        print(f'Error al ejecutar IW39: {e}')

def execute_pedidospendiente(session, parametros,nombre_archivo,formato,ruta,fecha):
    try:
        
        # Maximizar la ventana
        session.findById("wnd[0]").maximize()

        # Ingresar la transacción ME2N
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nME2N"
        session.findById("wnd[0]").sendVKey(0)

        session.findById("wnd[0]/usr/btn%_EN_EBELN_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()

        subprocess.run("echo off | clip", shell=True)

        session.findById("wnd[0]/usr/ctxtEN_EKORG-LOW").text = parametros.get("org_compra", "")
        session.findById("wnd[0]/usr/ctxtLISTU").text = parametros.get("alca_lista", "")

        cond_seleccion = pd.DataFrame({'WE101','WE107'},columns=['cond_selec'])
        cond_seleccion.to_clipboard(index=False,header=False)

        session.findById("wnd[0]/usr/btn%_SELPA_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()

        subprocess.run("echo off | clip", shell=True)

        session.findById("wnd[0]/usr/ctxtS_BSART-LOW").text = parametros.get("cls_doc", "")
        session.findById("wnd[0]/usr/ctxtS_EKGRP-LOW").text = parametros.get("grp_compra", "")
        session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = parametros.get("centro", "")
        session.findById("wnd[0]/usr/ctxtS_PSTYP-LOW").text = parametros.get("t_posicion", "")
        session.findById("wnd[0]/usr/ctxtS_KNTTP-LOW").text = parametros.get("t_imp", "")
        session.findById("wnd[0]/usr/ctxtS_EINDT-LOW").text = parametros.get("fech_entr", "")
        session.findById("wnd[0]/usr/ctxtP_GULDT").text = parametros.get("validez", "")
        session.findById("wnd[0]/usr/ctxtP_RWEIT").text = parametros.get("cobertura", "")
        session.findById("wnd[0]/usr/ctxtS_LIFNR-LOW").text = parametros.get("proveedor", "")
        session.findById("wnd[0]/usr/ctxtS_RESWK-LOW").text = parametros.get("centro_sumi", "")
        session.findById("wnd[0]/usr/ctxtS_MATNR-LOW").text = parametros.get("material", "")
        session.findById("wnd[0]/usr/ctxtS_MATKL-LOW").text = parametros.get("grp_artic", "")
        session.findById("wnd[0]/usr/ctxtS_BEDAT-LOW").text = fecha
        session.findById("wnd[0]/usr/txtS_EAN11-LOW").text = parametros.get("num_art", "")
        session.findById("wnd[0]/usr/txtS_IDNLF-LOW").text = parametros.get("n_mat_prov", "")
        session.findById("wnd[0]/usr/ctxtS_LTSNR-LOW").text = parametros.get("surt_par_prov", "")
        session.findById("wnd[0]/usr/ctxtS_AKTNR-LOW").text = parametros.get("accion", "")
        session.findById("wnd[0]/usr/ctxtS_SAISO-LOW").text = parametros.get("temporada", "")
        session.findById("wnd[0]/usr/txtS_SAISJ-LOW").text = parametros.get("anio_est", "")
        session.findById("wnd[0]/usr/txtP_TXZ01").text = parametros.get("txt_breve", "")
        session.findById("wnd[0]/usr/txtP_NAME1").text = parametros.get("nomb_prov", "")

        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[23]").press()
        session.findById("wnd[0]/tbar[1]/btn[33]").press()
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").currentCellRow = 0
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectedRows = "0"
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").clickCurrentCell()
        session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/txtGS_EXPORT-FILE_NAME").text = nombre_archivo
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/txtGS_EXPORT-FILE_NAME").caretPosition = len(nombre_archivo)
        session.findById("wnd[1]/tbar[0]/btn[20]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = ruta
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nombre_archivo+formato
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = len(nombre_archivo+formato)
        session.findById("wnd[1]/tbar[0]/btn[11]").press()

        #Finalizar Transaccion
        session.EndTransaction()

        print("Transacción ME2N configurada correctamente")
    except Exception as e:
        print(f'Error al ejecutar ME2N: {e}')

def execute_materialesIH09(session, parametros,nombre_archivo,formato,ruta,fecha):
    
    try:
        #Limpiar Portapeles
        subprocess.run("echo off | clip", shell=True)

        session.findById("wnd[0]").maximize()

        session.findById("wnd[0]/tbar[0]/okcd").Text = "/nIH09"
        session.findById("wnd[0]").sendVKey(0)

        session.findById("wnd[0]/usr/ctxtMS_MATNR-LOW").Text = parametros.get("material", "")
        session.findById("wnd[0]/usr/txtMAKTX-LOW").Text = parametros.get("txt_breve_mat", "")
        session.findById("wnd[0]/usr/ctxtMTART-LOW").Text = parametros.get("t_material", "")
        session.findById("wnd[0]/usr/ctxtLABOR-LOW").Text = parametros.get("laboratorio", "")
        session.findById("wnd[0]/usr/txtNORMT-LOW").Text = parametros.get("denom_estand", "")
        session.findById("wnd[0]/usr/txtZEINR-LOW").Text = parametros.get("documento", "")
        session.findById("wnd[0]/usr/txtBISMT-LOW").Text = parametros.get("n_ant_mat", "")
        session.findById("wnd[0]/usr/ctxtMATKL-LOW").Text = parametros.get("grp_articulo", "")
        session.findById("wnd[0]/usr/txtEAN11-LOW").Text = parametros.get("codi_ean", "")
        session.findById("wnd[0]/usr/ctxtLVORM-LOW").Text = parametros.get("pb_mand", "")
        session.findById("wnd[0]/usr/ctxtERNAM-LOW").Text = parametros.get("creado_por", "")
        session.findById("wnd[0]/usr/ctxtERSDA-LOW").Text = fecha
        session.findById("wnd[0]/usr/ctxtAENAM-LOW").Text = parametros.get("modificado", "")
        session.findById("wnd[0]/usr/ctxtLAEDA-LOW").Text = parametros.get("ult_modif", "")
        session.findById("wnd[0]/usr/ctxtVARIANT").Text = parametros.get("layout", "")

        session.findById("wnd[0]/tbar[1]/btn[8]").Press()

        session.findById("wnd[0]/mbar/menu[4]/menu[2]/menu[1]").Select()
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").ClickCurrentCell()
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").CurrentCellRow = 3
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").SelectAll()
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").ContextMenu()
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").SelectContextMenuItem("&XXL")
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/txtGS_EXPORT-FILE_NAME").Text = nombre_archivo
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/txtGS_EXPORT-FILE_NAME").caretPosition = len(nombre_archivo)
        session.findById("wnd[1]/tbar[0]/btn[20]").Press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").Text = ruta
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = nombre_archivo+formato
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = len(nombre_archivo+formato)
        session.findById("wnd[1]/tbar[0]/btn[11]").Press()

        #Finalizar Transaccion
        session.EndTransaction()

        print("Transacción IH09 configurada correctamente")
    except Exception as e:
        print(f'Error al ejecutar IH09: {e}')

def execute_trazabilidad(session, parametros,nombre_archivo,formato,ruta):
    try:
        subprocess.run("echo off | clip", shell=True)

        # Maximizar la ventana
        session.findById("wnd[0]").maximize()

        # Ingresar la transacción ZMMRI0013
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/nZMMRI0013"
        session.findById("wnd[0]/tbar[0]/btn[0]").Press()

        # Llenar los campos de selección
        session.findById("wnd[0]/usr/ctxtS_EKORG-LOW").Text = parametros.get("org_compra", "")
        session.findById("wnd[0]/usr/ctxtS_BSART-LOW").Text = parametros.get("cls_doc", "")
        session.findById("wnd[0]/usr/ctxtS_BEDAT-LOW").Text = parametros.get("fecha_doc1", "")
        session.findById("wnd[0]/usr/ctxtS_BEDAT-HIGH").text = parametros.get("fecha_doc2", "")
        session.findById("wnd[0]/usr/ctxtS_EBELN-LOW").Text = parametros.get("doc_compra", "")
        session.findById("wnd[0]/usr/ctxtS_KONNR-LOW").Text = parametros.get("contrato", "")
        session.findById("wnd[0]/usr/ctxtS_MATKL-LOW").Text = parametros.get("grp_artic", "")
        session.findById("wnd[0]/usr/ctxtS_EKGRP-LOW").Text = parametros.get("grp_compra", "")
        session.findById("wnd[0]/usr/ctxtS_RESWK-LOW").Text = parametros.get("centro_sumin", "")
        session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").Text = parametros.get("centro", "")
        session.findById("wnd[0]/usr/ctxtS_BUDAT-LOW").Text = parametros.get("fe_cont_entr_sal", "")
        session.findById("wnd[0]/usr/ctxtS_KNTTP-LOW").Text = parametros.get("t_imp", "")
        session.findById("wnd[0]/usr/ctxtS_LIFNR-LOW").Text = parametros.get("proveedor", "")
        session.findById("wnd[0]/usr/ctxtS_MATNR-LOW").Text = parametros.get("material", "")
        session.findById("wnd[0]/usr/ctxtS_MTART-LOW").Text = parametros.get("t_material", "")
        session.findById("wnd[0]/usr/ctxtS_KOSTL-LOW").Text = parametros.get("centro_coste", "")
        session.findById("wnd[0]/usr/ctxtS_AUFNR-LOW").Text = parametros.get("orden", "")
        session.findById("wnd[0]/usr/ctxtS_FEC_L-LOW").Text = parametros.get("fecha_lib", "")
        session.findById("wnd[0]/usr/ctxtS_LFDAT-LOW").Text = parametros.get("fecha_entr_alm", "")
        session.findById("wnd[0]/usr/ctxtS_CPUDT-LOW").Text = parametros.get("fecha_doc_em", "")
        session.findById("wnd[0]/usr/txtS_ERNAM-LOW").Text = parametros.get("usuario_creador", "")
        session.findById("wnd[0]/usr/ctxtS_BANFN-LOW").Text = parametros.get("solic_pedi", "")
        session.findById("wnd[0]/usr/ctxtS_BANPR-LOW").Text = parametros.get("status_trat", "")
        session.findById("wnd[0]/usr/ctxtS_AFNAM-LOW").Text = parametros.get("solicitante", "")
        session.findById("wnd[0]/usr/ctxtS_BADAT-LOW").Text = parametros.get("fech_soli", "")
        session.findById("wnd[0]/usr/ctxtS_CREAD-LOW").Text = parametros.get("usuario_crea_hess", "")
        session.findById("wnd[0]/usr/txtS_DATAT-LOW").Text = parametros.get("fecha_creacion", "")
        session.findById("wnd[0]/usr/ctxtS_HOJAE-LOW").Text = parametros.get("hoja_entr_serv", "")
        session.findById("wnd[0]/usr/ctxtS_ESTAD-LOW").Text = parametros.get("est_doc", "")
        session.findById("wnd[0]/usr/txtS_USUAR-LOW").Text = parametros.get("usuario_cread_hess", "")

        # Ejecutar
        session.findById("wnd[0]/tbar[1]/btn[8]").Press()

        # Escoger layout
        session.findById("wnd[0]/usr/cntlD0100_CONTAINER/shellcont/shell").PressToolbarButton("&MB_VARIANT")
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").ClickCurrentCell()

        # Guardar el archivo
        session.findById("wnd[0]/usr/cntlD0100_CONTAINER/shellcont/shell").PressToolbarContextButton("&MB_EXPORT")
        session.findById("wnd[0]/usr/cntlD0100_CONTAINER/shellcont/shell").SelectContextMenuItem("&XXL")
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/txtGS_EXPORT-FILE_NAME").Text = nombre_archivo
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/txtGS_EXPORT-FILE_NAME").caretPosition = len(nombre_archivo)
        session.findById("wnd[1]/tbar[0]/btn[20]").Press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").Text = ruta
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = nombre_archivo+formato
        session.findById("wnd[1]/usr/ctxtDY_PATH").SetFocus()
        session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = len(nombre_archivo+formato)
        session.findById("wnd[1]/tbar[0]/btn[11]").Press()

        #Finalizar Transaccion
        session.EndTransaction()
        
        print("Transacción ZMMRI0013 configurada correctamente")
    except Exception as e:
        print(f'Error al ejecutar ZMMRI0013: {e}')

def execute_inventario_scwm(session, parametros,nombre_archivo,formato,ruta):

    try:
        subprocess.run("echo off | clip", shell=True)

        session.findById("wnd[0]").maximize()

        session.findById("wnd[0]/tbar[0]/okcd").Text = "/N/SCWM/MON"
        session.findById("wnd[0]").sendVKey(0)

        session.findById("wnd[1]/usr/ctxtP_LGNUM").Text = parametros.get("almacen", "")
        session.findById("wnd[1]/usr/ctxtP_MONIT").Text = parametros.get("monitor", "")
        session.findById("wnd[1]/usr/txtP_REFR").text = parametros.get("nodo", "")
        session.findById("wnd[1]/tbar[0]/btn[8]").Press()
        session.findById("wnd[0]/usr/shell/shellcont[0]/shell").SelectedNode = "N000000137"
        session.findById("wnd[0]/usr/shell/shellcont[0]/shell").DoubleClickNode("N000000137")

        session.findById("wnd[1]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH").Press()
        session.findById("wnd[2]/tbar[0]/btn[16]").Press()
        session.findById("wnd[2]/tbar[0]/btn[24]").Press()
        session.findById("wnd[2]/tbar[0]/btn[8]").Press()

        session.findById("wnd[1]/usr/chkP_LGPLA").selected = False
        session.findById("wnd[1]/usr/chkP_RSRC").selected = True
        session.findById("wnd[1]/usr/chkP_TU").selected = True
        session.findById("wnd[1]/usr/chkP_E_LOGP").selected = False
        session.findById("wnd[1]/usr/ctxtS_STCAT-LOW").text = parametros.get("tipo_stock", "")
        session.findById("wnd[1]/usr/ctxtS_OWNER-LOW").text = parametros.get("propietario", "")
        session.findById("wnd[1]/usr/ctxtS_ENTIT-LOW").text = parametros.get("persona_autorz", "")
        session.findById("wnd[1]/usr/ctxtS_BATCH-LOW").text = parametros.get("lote", "")
        session.findById("wnd[1]/usr/ctxtS_STKSEG-LOW").text = parametros.get("seg_stock", "")
        session.findById("wnd[1]/usr/txtS_STOID-LOW").text = parametros.get("iden_stock", "")
        session.findById("wnd[1]/usr/txtS_SERID-LOW").text = parametros.get("num_ser", "")
        session.findById("wnd[1]/usr/txtS_UII-LOW").text = parametros.get("ident_pieza", "")
        session.findById("wnd[1]/usr/txtS_WIP_NO-LOW").text = parametros.get("num_wip", "")
        session.findById("wnd[1]/usr/ctxtS_HUIDEN-LOW").text = parametros.get("und_manip", "")

        session.findById("wnd[1]/tbar[0]/btn[8]").Press()

        session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell").pressToolbarButton("&MB_VARIANT")
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").clickCurrentCell()
        session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell").pressToolbarContextButton("&MB_EXPORT")
        session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell").selectContextMenuItem("&XXL")
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/txtGS_EXPORT-FILE_NAME").text = nombre_archivo+parametros.get("almacen", "")
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/txtGS_EXPORT-FILE_NAME").caretPosition = len(nombre_archivo+parametros.get("almacen", ""))
        session.findById("wnd[1]/tbar[0]/btn[20]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = ruta
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nombre_archivo+parametros.get("almacen", "")+formato
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = len(nombre_archivo+parametros.get("almacen", "")+formato)
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        session.findById("wnd[0]/tbar[0]/btn[15]").press()
        session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode("F00011")
        session.findById("wnd[0]/tbar[1]/btn[5]").press()
        session.findById("wnd[1]/tbar[0]/btn[12]").press()

        #Finalizar Transaccion
        session.EndTransaction()
        
        print("Transacción /SCWM/MON configurada correctamente")
        return True
    except Exception as e:
        print(f'Error al ejecutar /SCWM/MON: {e}')
        return False

def execute_inventario_scwm_v2(session, parametros,nombre_archivo,formato,ruta,centro):
    try:
        subprocess.run("echo off | clip", shell=True)

        session.findById("wnd[0]").maximize()
        
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/N/SCWM/MON"
        session.findById("wnd[0]/tbar[0]/btn[0]").Press()

        session.findById("wnd[0]/tbar[1]/btn[5]").Press()
        session.findById("wnd[1]/usr/ctxtP_LGNUM").Text = centro
        session.findById("wnd[1]/tbar[0]/btn[8]").Press()
        session.findById("wnd[0]/usr/shell/shellcont[0]/shell").ExpandNode("C000000011")
        session.findById("wnd[0]/usr/shell/shellcont[0]/shell").SelectedNode = "N000000137"
        session.findById("wnd[0]/usr/shell/shellcont[0]/shell").TopNode = "C000000001"
        session.findById("wnd[0]/usr/shell/shellcont[0]/shell").DoubleClickNode("N000000137")
        
        session.findById("wnd[1]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH").Press()
        session.findById("wnd[2]/tbar[0]/btn[16]").Press()
        session.findById("wnd[2]/tbar[0]/btn[24]").Press()
        session.findById("wnd[2]/tbar[0]/btn[8]").Press()

        session.findById("wnd[1]/usr/chkP_LGPLA").selected = False
        session.findById("wnd[1]/usr/chkP_RSRC").selected = True
        session.findById("wnd[1]/usr/chkP_TU").selected = True
        session.findById("wnd[1]/usr/chkP_E_LOGP").selected = False
        session.findById("wnd[1]/usr/ctxtS_STCAT-LOW").text = parametros.get("tipo_stock", "")
        session.findById("wnd[1]/usr/ctxtS_OWNER-LOW").text = parametros.get("propietario", "")
        session.findById("wnd[1]/usr/ctxtS_ENTIT-LOW").text = parametros.get("persona_autorz", "")
        session.findById("wnd[1]/usr/ctxtS_BATCH-LOW").text = parametros.get("lote", "")
        session.findById("wnd[1]/usr/ctxtS_STKSEG-LOW").text = parametros.get("seg_stock", "")
        session.findById("wnd[1]/usr/txtS_STOID-LOW").text = parametros.get("iden_stock", "")
        session.findById("wnd[1]/usr/txtS_SERID-LOW").text = parametros.get("num_ser", "")
        session.findById("wnd[1]/usr/txtS_UII-LOW").text = parametros.get("ident_pieza", "")
        session.findById("wnd[1]/usr/txtS_WIP_NO-LOW").text = parametros.get("num_wip", "")
        session.findById("wnd[1]/usr/ctxtS_HUIDEN-LOW").text = parametros.get("und_manip", "")

        session.findById("wnd[1]/tbar[0]/btn[8]").Press()

        session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell").pressToolbarButton("&MB_VARIANT")
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").clickCurrentCell()
        session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell").pressToolbarContextButton("&MB_EXPORT")
        session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell").selectContextMenuItem("&XXL")
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/txtGS_EXPORT-FILE_NAME").text = nombre_archivo+centro
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/txtGS_EXPORT-FILE_NAME").caretPosition = len(nombre_archivo+centro)
        session.findById("wnd[1]/tbar[0]/btn[20]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = ruta
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nombre_archivo+centro+formato
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = len(nombre_archivo+centro+formato)
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        session.findById("wnd[0]/tbar[0]/btn[15]").press()
        session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode("F00011")
        session.findById("wnd[0]/tbar[1]/btn[5]").press()
        session.findById("wnd[1]/tbar[0]/btn[12]").press()

        #Finalizar Transaccion
        session.EndTransaction()
        
        print("Transacción /SCWM/MON_v2 configurada correctamente")
        return True
    except Exception as e:
        print(f'Error al ejecutar /SCWM/MON_v2: {e}')
        return False

def execute_movimientos_MB51(session, parametros,nombre_archivo,formato,ruta,fecha):
    try:
        session.findById("wnd[0]").maximize()

        session.findById("wnd[0]/tbar[0]/okcd").text = "/NMB51"
        session.findById("wnd[0]").sendVKey(0)

        session.findById("wnd[0]/usr/ctxtMATNR-LOW").text = parametros.get("material", "")
        session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = parametros.get("centro", "")
        session.findById("wnd[0]/usr/ctxtLGORT-LOW").text = parametros.get("almacen", "")
        session.findById("wnd[0]/usr/ctxtCHARG-LOW").text = parametros.get("lote", "")
        session.findById("wnd[0]/usr/ctxtLIFNR-LOW").text = parametros.get("proveedor", "")
        session.findById("wnd[0]/usr/ctxtKUNNR-LOW").text = parametros.get("cliente", "")
        session.findById("wnd[0]/usr/ctxtBWART-LOW").text = parametros.get("clase_movimiento", "")
        session.findById("wnd[0]/usr/ctxtSOBKZ-LOW").text = parametros.get("stock_especial", "")
        session.findById("wnd[0]/usr/ctxtANLN1-LOW").text = parametros.get("actv_fijo", "")
        session.findById("wnd[0]/usr/ctxtAUFNR-LOW").text = parametros.get("orden", "")
        session.findById("wnd[0]/usr/ctxtEBELN-LOW").text = parametros.get("pedido", "")
        session.findById("wnd[0]/usr/ctxtKDAUF-LOW").text = parametros.get("pedido_cliente", "")
        session.findById("wnd[0]/usr/ctxtKOSTL-LOW").text = parametros.get("centro_coste", "")
        session.findById("wnd[0]/usr/txtRSNUM-LOW").text = parametros.get("reserva", "")
        session.findById("wnd[0]/usr/ctxtBUDAT-LOW").text = fecha
        session.findById("wnd[0]/usr/txtUSNAM-LOW").text = parametros.get("usuario", "")
        session.findById("wnd[0]/usr/ctxtVGART-LOW").text = parametros.get("clase_tran", "")
        session.findById("wnd[0]/usr/txtLE_VBELN-LOW").text = parametros.get("entrega", "")
        session.findById("wnd[0]/usr/txtMBLNR-LOW").text = parametros.get("doc_mat", "")
        session.findById("wnd[0]/usr/txtXBLNR-LOW").text = parametros.get("refe", "")
        session.findById("wnd[0]/usr/radRFLAT_L").setFocus()
        session.findById("wnd[0]/usr/radRFLAT_L").select()
        session.findById("wnd[0]/usr/chkDATABASE").selected = True
        session.findById("wnd[0]/usr/chkSHORTDOC").selected = True
        session.findById("wnd[0]/usr/chkARCHIVE").selected = False
        session.findById("wnd[0]/usr/ctxtALV_DEF").text = parametros.get("layout", "")
        session.findById("wnd[0]/usr/ctxtPA_AISTR").text = parametros.get("inf_arch", "")

        session.findById("wnd[0]/tbar[1]/btn[8]").press()

        session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[1]").select()
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/txtGS_EXPORT-FILE_NAME").text = nombre_archivo
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/txtGS_EXPORT-FILE_NAME").caretPosition = len(nombre_archivo)
        session.findById("wnd[1]/tbar[0]/btn[20]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = ruta
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nombre_archivo+formato
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = len(nombre_archivo+formato)
        session.findById("wnd[1]/tbar[0]/btn[11]").press()

        #Finalizar Transaccion
        session.EndTransaction()
        
        print("Transacción MB51 configurada correctamente")

    except Exception as e:
        print(f"Error al ejecutar MB52: {e}")

def execute_pedidospendiente_hist(session, parametros,nombre_archivo,formato,ruta,fecha):
    try:
        
        # Maximizar la ventana
        session.findById("wnd[0]").maximize()

        # Ingresar la transacción ME2N
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nME2N"
        session.findById("wnd[0]").sendVKey(0)

        session.findById("wnd[0]/usr/btn%_EN_EBELN_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()

        subprocess.run("echo off | clip", shell=True)

        session.findById("wnd[0]/usr/ctxtEN_EKORG-LOW").text = parametros.get("org_compra", "")
        session.findById("wnd[0]/usr/ctxtLISTU").text = parametros.get("alca_lista", "")

        cond_seleccion = pd.DataFrame({},columns=['cond_selec'])
        cond_seleccion.to_clipboard(index=False,header=False)

        session.findById("wnd[0]/usr/btn%_SELPA_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()

        subprocess.run("echo off | clip", shell=True)

        session.findById("wnd[0]/usr/ctxtS_BSART-LOW").text = parametros.get("cls_doc", "")
        session.findById("wnd[0]/usr/ctxtS_EKGRP-LOW").text = parametros.get("grp_compra", "")
        session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = parametros.get("centro", "")
        session.findById("wnd[0]/usr/ctxtS_PSTYP-LOW").text = parametros.get("t_posicion", "")
        session.findById("wnd[0]/usr/ctxtS_KNTTP-LOW").text = parametros.get("t_imp", "")
        session.findById("wnd[0]/usr/ctxtS_EINDT-LOW").text = parametros.get("fech_entr", "")
        session.findById("wnd[0]/usr/ctxtP_GULDT").text = parametros.get("validez", "")
        session.findById("wnd[0]/usr/ctxtP_RWEIT").text = parametros.get("cobertura", "")
        session.findById("wnd[0]/usr/ctxtS_LIFNR-LOW").text = parametros.get("proveedor", "")
        session.findById("wnd[0]/usr/ctxtS_RESWK-LOW").text = parametros.get("centro_sumi", "")
        session.findById("wnd[0]/usr/ctxtS_MATNR-LOW").text = parametros.get("material", "")
        session.findById("wnd[0]/usr/ctxtS_MATKL-LOW").text = parametros.get("grp_artic", "")
        session.findById("wnd[0]/usr/ctxtS_BEDAT-LOW").text = fecha
        session.findById("wnd[0]/usr/txtS_EAN11-LOW").text = parametros.get("num_art", "")
        session.findById("wnd[0]/usr/txtS_IDNLF-LOW").text = parametros.get("n_mat_prov", "")
        session.findById("wnd[0]/usr/ctxtS_LTSNR-LOW").text = parametros.get("surt_par_prov", "")
        session.findById("wnd[0]/usr/ctxtS_AKTNR-LOW").text = parametros.get("accion", "")
        session.findById("wnd[0]/usr/ctxtS_SAISO-LOW").text = parametros.get("temporada", "")
        session.findById("wnd[0]/usr/txtS_SAISJ-LOW").text = parametros.get("anio_est", "")
        session.findById("wnd[0]/usr/txtP_TXZ01").text = parametros.get("txt_breve", "")
        session.findById("wnd[0]/usr/txtP_NAME1").text = parametros.get("nomb_prov", "")

        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[23]").press()
        session.findById("wnd[0]/tbar[1]/btn[33]").press()
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").currentCellRow = 0
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectedRows = "0"
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").clickCurrentCell()
        session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/txtGS_EXPORT-FILE_NAME").text = nombre_archivo
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/txtGS_EXPORT-FILE_NAME").caretPosition = len(nombre_archivo)
        session.findById("wnd[1]/tbar[0]/btn[20]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = ruta
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nombre_archivo+formato
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = len(nombre_archivo+formato)
        session.findById("wnd[1]/tbar[0]/btn[11]").press()

        #Finalizar Transaccion
        session.EndTransaction()

        print("Transacción ME2N configurada correctamente")
    except Exception as e:
        print(f'Error al ejecutar ME2N: {e}')

def execute_inventario_IPESA(session,nombre_archivo,formato,ruta):
    try:

        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nZMMRP0002"
        session.findById("wnd[0]").sendVKey(0)

        session.findById("wnd[0]/usr/radRAD1").setFocus()
        session.findById("wnd[0]/usr/radRAD1").select()
        session.findById("wnd[0]/usr/radRSDIS").select()
        session.findById("wnd[0]/usr/radRREP").select()
        session.findById("wnd[0]/usr/radRREP").setFocus()

        session.findById("wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()

        session.findById("wnd[0]/tbar[1]/btn[8]").press()

        session.findById("wnd[0]/tbar[1]/btn[33]").press()
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").clickCurrentCell()
        session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/txtGS_EXPORT-FILE_NAME").text = nombre_archivo
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/txtGS_EXPORT-FILE_NAME").caretPosition = len(nombre_archivo)
        session.findById("wnd[1]/tbar[0]/btn[20]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = ruta
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nombre_archivo+formato
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = len(nombre_archivo+formato)
        session.findById("wnd[1]/tbar[0]/btn[11]").press()

        #Finalizar Transaccion
        session.EndTransaction()

    except Exception as e:
        print(f'Error al ejecutar ME2N: {e}')

def execute_entregas_ewm(session, parametros,nombre_archivo_cabezera,nombre_archivo_detalle,formato,ruta,fecha):
    try:

        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/N/SCWM/MON"
        session.findById("wnd[0]").sendVKey(0)

        session.findById("wnd[1]/usr/ctxtP_LGNUM").Text = parametros.get("almacen", "")
        session.findById("wnd[1]/usr/ctxtP_MONIT").Text = parametros.get("monitor", "")
        session.findById("wnd[1]/usr/txtP_REFR").text = parametros.get("nodo", "")
        session.findById("wnd[1]/tbar[0]/btn[8]").press()

        session.findById("wnd[0]/usr/shell/shellcont[0]/shell").selectedNode = "N000000010"
        session.findById("wnd[0]/usr/shell/shellcont[0]/shell").doubleClickNode("N000000010")
        session.findById("wnd[1]/usr/chkP_REFDOC").selected = True
        session.findById("wnd[1]/usr/ctxtP_CDATFR").text = fecha
        session.findById("wnd[1]/usr/chkP_REFDOC").setFocus()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell").pressToolbarButton("&MB_VARIANT")
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").currentCellRow = 0
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectedRows = "0"
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").clickCurrentCell()
        session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell").pressToolbarContextButton("&MB_EXPORT")
        session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell").selectContextMenuItem("&XXL")
        session.findById("wnd[1]/tbar[0]/btn[20]").press()
        session.findById("wnd[1]").sendVKey(4)
        session.findById("wnd[2]/usr/ctxtDY_PATH").text = ruta
        session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = nombre_archivo_cabezera+formato
        session.findById("wnd[2]/tbar[0]/btn[11]").press()
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell").setCurrentCell(-1,"")
        session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell").selectAll()
        session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell").pressToolbarButton("N000000011")
        session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[1]/shell").pressToolbarButton("&MB_VARIANT")
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").clickCurrentCell()
        session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[1]/shell").pressToolbarContextButton("&MB_EXPORT")
        session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[1]/shell").selectContextMenuItem("&XXL")
        session.findById("wnd[1]/tbar[0]/btn[20]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = ruta
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nombre_archivo_detalle+formato
        session.findById("wnd[1]/tbar[0]/btn[11]").press()

        print("Transacción /SCWM/MON-C154 configurada correctamente")
        return True

    except Exception as e:
        print(f'Error al ejecutar /N/SCWM/MON(C154): {e}')

def execute_entregas_ewm_v2(session,nombre_archivo_cabezera,nombre_archivo_detalle,formato,ruta,fecha,centro):
    try:

        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[1]/btn[5]").press()
        session.findById("wnd[1]/usr/ctxtP_LGNUM").text = centro
        session.findById("wnd[1]/tbar[0]/btn[8]").press()

        session.findById("wnd[0]/usr/shell/shellcont[0]/shell").selectedNode = "N000000010"
        session.findById("wnd[0]/usr/shell/shellcont[0]/shell").doubleClickNode("N000000010")
        session.findById("wnd[1]/usr/chkP_REFDOC").selected = True
        session.findById("wnd[1]/usr/ctxtP_CDATFR").text = fecha
        session.findById("wnd[1]/usr/chkP_REFDOC").setFocus()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell").pressToolbarButton("&MB_VARIANT")
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").currentCellRow = 0
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectedRows = "0"
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").clickCurrentCell()
        session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell").pressToolbarContextButton("&MB_EXPORT")
        session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell").selectContextMenuItem("&XXL")
        session.findById("wnd[1]/tbar[0]/btn[20]").press()
        session.findById("wnd[1]").sendVKey(4)
        session.findById("wnd[2]/usr/ctxtDY_PATH").text = ruta
        session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = nombre_archivo_cabezera+centro+formato
        session.findById("wnd[2]/tbar[0]/btn[11]").press()
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell").setCurrentCell(-1,"")
        session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell").selectAll()
        session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell").pressToolbarButton("N000000011")
        session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[1]/shell").pressToolbarButton("&MB_VARIANT")
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").clickCurrentCell()
        session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[1]/shell").pressToolbarContextButton("&MB_EXPORT")
        session.findById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[1]/shell").selectContextMenuItem("&XXL")
        session.findById("wnd[1]/tbar[0]/btn[20]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = ruta
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nombre_archivo_detalle+centro+formato
        session.findById("wnd[1]/tbar[0]/btn[11]").press()

        print(f"Transacción /SCWM/MON-C{centro} configurada correctamente")

    except Exception as e:
        print(f'Error al ejecutar /N/SCWM/MON: {e}')

def execute_guias_remision(session,nombre_archivo,formato,ruta,fecha):
    try:

        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "ZOSFE004"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/btnGW_BTMN_DOC").press()
        session.findById("wnd[0]/usr/ctxtP_BUKRS").text = "PE02"
        session.findById("wnd[0]/usr/radP_GR").setFocus()
        session.findById("wnd[0]/usr/radP_GR").select()
        session.findById("wnd[0]/usr/ctxtS_FECHA-LOW").text = fecha
        session.findById("wnd[0]/usr/ctxtS_FECHA-LOW").setFocus()
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
        session.findById("wnd[1]/tbar[0]/btn[20]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = ruta
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nombre_archivo+formato
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        
        #Finalizar Transaccion
        session.EndTransaction()

    except Exception as e:
        print(f'Error al ejecutar ME2N: {e}')

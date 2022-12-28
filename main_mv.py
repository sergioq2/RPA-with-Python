import os
import shutil
import pandas as pd
from docxtpl import DocxTemplate
from openpyxl import load_workbook
import smtplib
from email.message import EmailMessage
import ssl
import win32com.client as win32
import time
from docxtpl import InlineImage
from docx.shared import Mm
import io
import folium
from PIL import Image
import imgkit
from html2image import Html2Image
import pyodbc 
import cx_Oracle
import pyautogui
import cv2
import numpy as np
import codecs
import webbrowser
import time
import os.path
from os import path
from datetime import datetime

##Lectura de todas las fuentes de excel
#información es la variable que guarda la información que se lee desde el archivo de Excel original con la información proveniente de Power apps
informacion = pd.read_excel('BD_ExperienciasOperativas_PowerApps.xlsx')
#antecedentes_ow es la variable que guarda la información de los antecedentes descargados de OW, los cuales están en el excel Oneworld.xlsx
antecedentes_ow = pd.read_excel('Oneworld.xlsx',sheet_name='historico_oneworld')
#en directorio correos está el listado de todos los correos de los responsables en realizar los informes de experiencias operativas
directorio_correos = pd.read_excel('directorio_activo.xlsx')
#Base_BD es el archivo con el que se va a comparar las experiencias registradas para identificar si hay nuevos registros agregados
base = pd.read_excel('Base_BD.xlsx')
#template.docx es el template en word en el que finalmente se generará el informe de experiencias operativas
docx_tpl = DocxTemplate("template.docx")


#La función principal es la responsable de ejecutar cada una de las funciones programadas, estas las ejecutará por cada uno de los registros nuevos que identifique.
#Esa identificación de registros nuevos se hace en las primeras lineas de la función, comparando las dimensiones de archivo de registros con el archivo base, y va 
#a aplicar cada una de las funciones a cada uno de los registros nuevos
def principal():
    dim_base = base.shape[0]
    dim_experiencias = informacion.shape[0]
    diferencia = dim_experiencias - dim_base
    if diferencia > 0:
        for reg in range(dim_experiencias):
            if (reg+1) <= dim_base:
                continue
            else:
                if pd.isna(informacion.loc[reg,'Ipid']) == True:
                    ipid = 0
                else:
                    ipid = int(informacion.loc[reg,'Ipid'])
                tipo_elemento, df_modelo = conexion_modelo(ipid)
                completar_informacion(reg, tipo_elemento, df_modelo)
                mapa_dano(reg)
                nombre_informe = crear_informe(reg, tipo_elemento)
                link_documento = guardar_documento(nombre_informe)
                informacion.loc[reg,'LINK_INFORME'] = link_documento
                enviar_correo(reg,link_documento)
    informacion.to_excel('BD_ExperienciasOperativas_PowerApps_db.xlsx',index=False,sheet_name='ExperienciasOperativas')
    informacion.to_excel('Base_BD.xlsx',index=False)
    print("se ejecutó correctamente")

#Esta función realiza la conexión al modelo de red, la primer parte de la función son los parámetros de la conexión a GAGUPROD, y en el execute es todo el query
#con la información a traer, la trabla de la conexión, y los filtros. Al final todo se organiza en un dataframe llamado bd_modelo
def conexion_modelo(num_ipid):
    dsn_tns = cx_Oracle.makedsn('EPM-PO13', 1521, service_name='GAGUPROD')
    conn = cx_Oracle.connect(user='CONGAGUAS', password='congaguas1', dsn=dsn_tns)
    c = conn.cursor()
    c.execute("""
        SELECT
            IPID, DIAMETRO_NOMINAL, MATERIAL, TIPO_RED, FABRICANTE, GRUPO, NOMBRE_OPERACION, NOMBRE_MTTO, FECHA_INSTALACION, COOR_LAT, COOR_LON,
             LONGITUD
        FROM
            GAGUAS.GM_VATUB_PRM_LN
        WHERE
            IPID = {num_ipid}
        """.format(num_ipid=num_ipid))
    bd_modelo = pd.DataFrame(c, columns = ['IPID', 'DIAMETRO_NOMINAL', 'MATERIAL', 'TIPO_RED', 'FABRICANTE', 'GRUPO',
                                     'NOMBRE_OPERACION', 'NOMBRE_MTTO', 'FECHA_INSTALACION', 'COOR_LAT', 'COOR_LON', 'LONGITUD'])
    if bd_modelo.shape[0]>0:
        tipo_elemento = "Redes Primarias"
        return(tipo_elemento, bd_modelo)
    else:
        c.execute("""
        SELECT
            IPID, NUMERO_VALVULA, TIPO_VALVULA, FUNCION_VALVULA, DIAMETRO, FABRICANTE, GRUPO, FECHA_INSTALACION, TIPO_AGUA, COOR_LON, COOR_LAT
        FROM
            GAGUAS.GM_VAVAL_PRM_PT
        WHERE
            IPID = {num_ipid}
        """.format(num_ipid=num_ipid))
    bd_modelo = pd.DataFrame(c, columns = ['IPID', 'NUMERO_VALVULA', 'TIPO_VALVULA', 'FUNCION_VALVULA', 'DIAMETRO', 'FABRICANTE', 
                        'GRUPO', 'FECHA_INSTALACION', 'TIPO_AGUA', 'COOR_LON', 'COOR_LAT'])
    if bd_modelo.shape[0]>0:
        tipo_elemento = "Valvula"
        return(tipo_elemento, bd_modelo)
    else:
        c.execute("""
        SELECT
            IPID, DIAMETRO_NOMINAL, MATERIAL, FABRICANTE, PROFUNDIDAD, FECHA_INSTALACION, COOR_LAT, COOR_LON, LONGITUD, NOMBRE_CIRCUITO
        FROM
            GAGUAS.GM_VATUB_SCN_LN
        WHERE
            IPID = {num_ipid}
        """.format(num_ipid=num_ipid))
    bd_modelo = pd.DataFrame(c, columns = ['IPID', 'DIAMETRO_NOMINAL', 'MATERIAL', 'FABRICANTE', 'PROFUNDIDAD', 'FECHA_INSTALACION',
                        'COOR_LAT', 'COOR_LON', 'LONGITUD', 'NOMBRE_CIRCUITO'])
    if bd_modelo.shape[0]>0:
        tipo_elemento = "Redes Secundarias"
        return(tipo_elemento, bd_modelo)
    else:
        tipo_elemento = "El IPID no existe"
        return(tipo_elemento, bd_modelo)


#La función completar_informacion se encarga de terminar de diligenciar los campos provenientes del modelo de red en el archivo de Excel donde se está registrando la 
#información proveniente del PowerApp
def completar_informacion(registro, tipo_elemento, df_modelo):
    if tipo_elemento == 'Redes Primarias':
        informacion.loc[registro,'TIPO_ELEMENTO'] = tipo_elemento
        informacion.loc[registro,'DIAMETRO'] = df_modelo['DIAMETRO_NOMINAL'].values[0]
        informacion.loc[registro, 'MATERIAL'] = df_modelo['MATERIAL'].values[0]
        informacion.loc[registro, 'TIPO_RED'] = df_modelo['TIPO_RED'].values[0]
        informacion.loc[registro, 'FABRICANTE'] = df_modelo['FABRICANTE'].values[0]
        informacion.loc[registro, 'GRUPO'] = df_modelo['GRUPO'].values[0]
        informacion.loc[registro, 'NOMBRE_OPERACION'] = df_modelo['NOMBRE_OPERACION'].values[0]
        informacion.loc[registro, 'NOMBRE_MTTO'] = df_modelo['NOMBRE_MTTO'].values[0]
        informacion.loc[registro, 'FECHA_INSTALACION'] = df_modelo['FECHA_INSTALACION'].values[0]
        informacion.loc[registro, 'COOR_LAT'] = df_modelo['COOR_LAT'].values[0]
        informacion.loc[registro, 'COOR_LON'] = df_modelo['COOR_LON'].values[0]
        informacion.loc[registro, 'LONGITUD'] = df_modelo['LONGITUD'].values[0]
    elif tipo_elemento == 'Valvula':
        informacion.loc[registro,'TIPO_ELEMENTO'] = tipo_elemento
        informacion.loc[registro, 'NUMERO_VALVULA'] = df_modelo['NUMERO_VALVULA'].values[0]
        informacion.loc[registro, 'TIPO_VALVULA'] = df_modelo['TIPO_VALVULA'].values[0]
        informacion.loc[registro, 'FUNCION_VALVULA'] = df_modelo['FUNCION_VALVULA'].values[0]
        informacion.loc[registro, 'TIPO_AGUA'] = df_modelo['TIPO_AGUA'].values[0]
        informacion.loc[registro, 'FABRICANTE'] = df_modelo['FABRICANTE'].values[0]
        informacion.loc[registro, 'GRUPO'] = df_modelo['GRUPO'].values[0]
        informacion.loc[registro, 'TIPO_AGUA'] = df_modelo['TIPO_AGUA'].values[0]
        informacion.loc[registro, 'DIAMETRO'] = df_modelo['DIAMETRO'].values[0]
        informacion.loc[registro, 'COOR_LAT'] = df_modelo['COOR_LAT'].values[0]
        informacion.loc[registro, 'COOR_LON'] = df_modelo['COOR_LON'].values[0]
        informacion.loc[registro, 'FECHA_INSTALACION'] = df_modelo['FECHA_INSTALACION'].values[0]
    elif tipo_elemento == "Redes Secundarias":
        informacion.loc[registro,'TIPO_ELEMENTO'] = tipo_elemento
        informacion.loc[registro,'DIAMETRO'] = df_modelo['DIAMETRO_NOMINAL'].values[0]
        informacion.loc[registro, 'MATERIAL'] = df_modelo['MATERIAL'].values[0]
        informacion.loc[registro, 'FABRICANTE'] = df_modelo['FABRICANTE'].values[0]
        informacion.loc[registro, 'PROFUNDIDAD'] = df_modelo['PROFUNDIDAD'].values[0]
        informacion.loc[registro, 'FECHA_INSTALACION'] = df_modelo['FECHA_INSTALACION'].values[0]
        informacion.loc[registro, 'COOR_LAT'] = df_modelo['COOR_LAT'].values[0]
        informacion.loc[registro, 'COOR_LON'] = df_modelo['COOR_LON'].values[0]
        informacion.loc[registro, 'LONGITUD'] = df_modelo['LONGITUD'].values[0]
        informacion.loc[registro, 'NOMBRE_CIRCUITO'] = df_modelo['NOMBRE_CIRCUITO'].values[0]
    return(informacion)


#Esta función antecedentes_sistema se encarga primer de extraer los antecedentes que encuentre en el cubo ALFAII, y los junta con los antecedentes que encuentra en OW 
#hallados con la función antecedentes_ow, y finalmente los concatena en un dataframe llamado df_antecedentes trayendo los últimos 5 antecedentes
def antecedentes_sistema(num_ipid, tipo_elemento):
    if tipo_elemento != "El IPID no existe":
        conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=epm-ps04;'
                      'Database=ALFAII;'
                      'Trusted_Connection=yes;')

        sql_query = pd.read_sql_query ("""
            SELECT
                IDSolicitud, IPID, FecCreacionOrden, CausaEvento, ObservacionSolicitud, DesEfectividad
            FROM
                VOPERAGeneralHistoricoVPAA
            WHERE
                DesEfectividad = 'SI'
            AND
                IPID = {num_ipid}
            """.format(num_ipid=num_ipid),conn)
        df = pd.DataFrame(sql_query, columns = ['OT','IPID','FECHA','MOTIVO','DESCRIPCION'])
        df['ORIGEN'] = 'GESTA'
        df = df[df['OT'].notna()]

        df_ow = antecedentes_ow[antecedentes_ow['IPID']==num_ipid]
        df_ow['ORIGEN'] = "ONEWORLD"
        df_antecedentes = pd.concat([df, df_ow], axis=0, ignore_index=True)
        while df_antecedentes.shape[0]<5:
            df_antecedentes = df_antecedentes.append({'OT' : "", 'IPID' : "" , 'FECHA' : "", 'MOTIVO': "", 'DESCRIPCION':"", 'ORIGEN': ""},ignore_index=True)
    else:
        df_antecedentes = pd.DataFrame(index=np.arange(5), columns=['OT','IPID','FECHA','MOTIVO','DESCRIPCION', 'ORIGEN'])
    return(df_antecedentes)


##Esta función crear_informe genera el informe de antecedentes llenando todos los campos el template de word cargado
def crear_informe(registro, tipo_elemento):
    ipid_elemento = informacion.loc[registro,'Ipid']
    numero_ot = int(informacion.loc[registro,'Numero de OT'])
    nombre_func = informacion.loc[registro, 'Funcionario que Registra']
    nombre_informe = str(numero_ot)
    unidad = directorio_correos['UNIDAD'][directorio_correos['NOMBRE']==nombre_func].values[0]
    ruta = 'Informes/'+str(numero_ot)
    contenido = os.listdir(ruta+'//'+'fotos')
    if len(contenido)>0:
        dif = 4 - len(contenido)
        if dif > 0:
            for i in range(dif):
                contenido.append(contenido[0])
        contenido = contenido[0:4]
        imagen1 = InlineImage(docx_tpl, ruta+'//'+'fotos'+'//'+str(contenido[0]), width=Mm(60))
        imagen2 = InlineImage(docx_tpl, ruta+'//'+'fotos'+'//'+str(contenido[1]), width=Mm(60))
        imagen3 = InlineImage(docx_tpl, ruta+'//'+'fotos'+'//'+str(contenido[2]), width=Mm(60))
        imagen4 = InlineImage(docx_tpl, ruta+'//'+'fotos'+'//'+str(contenido[3]), width=Mm(60))
    else:
        imagen1 = "No hay foto"
        imagen2 = "No hay foto"
        imagen3 = "No hay foto"
        imagen4 = "No hay foto"

    if path.exists(ruta+'\\'+'dano.png'):
            ubicacion_dano = InlineImage(docx_tpl, ruta+'\\'+'dano.png' , width=Mm(110))
    else:
            ubicacion_dano = ""

    condicion_entorno = informacion.loc[registro,'Condiciones del Entorno']
    condiciones_arboles = informacion.loc[registro,'Árboles (Raíces)']
    condiciones_trafico = informacion.loc[registro,'Tráfico Pesado']
    condiciones_deformaciones = informacion.loc[registro,'Deformaciones en el Terreno']
    condiciones_incendios = informacion.loc[registro,'Incendios']
    condiciones_clima_ext = informacion.loc[registro,'Condiciones climáticas extremas']
    condiciones_clima = informacion.loc[registro,'Condiciones climáticas']
    condiciones_terceros = informacion.loc[registro,'Manipulación por tercero']
    condiciones_freatico = informacion.loc[registro,'Nivel freático']
    condiciones_suelo = informacion.loc[registro,'Suelo orgánico']
    condiciones_material = informacion.loc[registro,'Material de lleno no apropiado']


    antecedentes = antecedentes_sistema(ipid_elemento, tipo_elemento)

    date_f = informacion.loc[registro,'FECHA_INSTALACION']
    if pd.isnull(date_f) == True:
        date_f = '1900-01-01'
        date_f = datetime.strptime(date_f, '%Y-%m-%d')

    fecha_r = informacion.loc[registro,'Fecha de Registro']

    context = {
            'UNIDAD' : unidad,
            'NOMBRE_PERSONA' : informacion.loc[registro,'Funcionario que Registra'],
            'FECHA' : fecha_r,
            'OT': informacion.loc[registro,'Numero de OT'],
            'OPERACION': informacion.loc[registro,'NOMBRE_OPERACION'],
            'CIRCUITO': informacion.loc[registro,'NOMBRE_CIRCUITO'],
            'COORDENADA_NOR': informacion.loc[registro,'COOR_LAT'],
            'COORDENADA_OCC': informacion.loc[registro,'COOR_LON'],
            'ELEMENTO': informacion.loc[registro,'TIPO_ELEMENTO'],
            'MATERIAL': informacion.loc[registro,'MATERIAL'],
            'DIAMETRO': informacion.loc[registro,'DIAMETRO'],
            'PROFUNDIDAD': informacion.loc[registro,'PROFUNDIDAD'],
            'IPID': int(informacion.loc[registro,'Ipid']),
            'FABRICANTE': informacion.loc[registro,'FABRICANTE'],
            'FECHA_INSTALACION': date_f,
            'INTERRUPCION': informacion.loc[registro,'Interrupcion del Servicio'],
            'AFECTACION': informacion.loc[registro,'Afectacion a Terceros'],
            'INFERIOR': informacion.loc[registro,'Posicion Inferior'],
            'SUPERIOR': informacion.loc[registro,'Posicion Superior'],
            'IZQUIERDA': informacion.loc[registro,'Posicion Izquierda'],
            'DERECHA': informacion.loc[registro,'Posicion Derecha'],
            'FECHA_SOLICITUD_1' : antecedentes.loc[0,'FECHA'],
            'MODO_FALLA_1' : antecedentes.loc[0,'MOTIVO'],
            'DESCRIPCION_1' : antecedentes.loc[0,'DESCRIPCION'],
            'O_1' : antecedentes.loc[0,'OT'],
            'ORIGEN_1': antecedentes.loc[0,'ORIGEN'],
            'FECHA_SOLICITUD_2' : antecedentes.loc[1,'FECHA'],
            'MODO_FALLA_2' : antecedentes.loc[1,'MOTIVO'],
            'DESCRIPCION_2' : antecedentes.loc[1,'DESCRIPCION'],
            'O_2' : antecedentes.loc[1,'OT'],
            'ORIGEN_2': antecedentes.loc[1,'ORIGEN'],
            'FECHA_SOLICITUD_3' : antecedentes.loc[2,'FECHA'],
            'MODO_FALLA_3' : antecedentes.loc[2,'MOTIVO'],
            'DESCRIPCION_3' : antecedentes.loc[2,'DESCRIPCION'],
            'O_3' : antecedentes.loc[2,'OT'],
            'ORIGEN_3': antecedentes.loc[2,'ORIGEN'],
            'FECHA_SOLICITUD_4' : antecedentes.loc[3,'FECHA'],
            'MODO_FALLA_4' : antecedentes.loc[3,'MOTIVO'],
            'DESCRIPCION_4' : antecedentes.loc[3,'DESCRIPCION'],
            'O_4' : antecedentes.loc[3,'OT'],
            'ORIGEN_4': antecedentes.loc[3,'ORIGEN'],
            'FECHA_SOLICITUD_5' : antecedentes.loc[4,'FECHA'],
            'MODO_FALLA_5' : antecedentes.loc[4,'MOTIVO'],
            'DESCRIPCION_5' : antecedentes.loc[4,'DESCRIPCION'],
            'O_5' : antecedentes.loc[4,'OT'],
            'ORIGEN_5': antecedentes.loc[4,'ORIGEN'],
            'IMAGEN_1': imagen1,
            'IMAGEN_2': imagen2,
            'IMAGEN_3': imagen3,
            'IMAGEN_4': imagen4,
            'LOCALIZACION': ubicacion_dano,
            'CONDICION_ENTORNO':condicion_entorno,
            'CONDICION_ARBOLES':condiciones_arboles,
            'CONDICION_TRAFICO': condiciones_trafico,
            'CONDICION_DEFORMACIONES': condiciones_deformaciones,
            'CONDICION_INCENDIOS':condiciones_incendios,
            'CONDICION_CLIMA_EXT':condiciones_clima_ext,
            'CONDICION_CLIMA': condiciones_clima,
            'CONDICION_TERCEROS': condiciones_terceros,
            'CONDICION_FREATICO': condiciones_freatico,
            'CONDICION_SUELO': condiciones_suelo,
            'CONDICION_MATERIAL': condiciones_material
            }
    docx_tpl.render(context)
    return(nombre_informe)

#La función mapa_dano se encarga de sacar el mapa de donde ocurrió el daño
def mapa_dano(registro):
    coordenada_x = informacion.loc[registro,'COOR_LAT']
    coordenada_y = informacion.loc[registro,'COOR_LON']
    if coordenada_x > 0:
        ipid = informacion.loc[registro,'Ipid']
        ot = int(informacion.loc[registro,'Numero de OT'])
        m = folium.Map(location=[coordenada_x, coordenada_y], zoom_start=16)
        m.add_child(folium.Marker(location=[coordenada_x,coordenada_y], popup = ipid))
        mapFname = './Informes' +'\\'+str(ot)+'\\'+'dano.html'
        file_png = './Informes' +'\\'+str(ot)+'\\'+'dano.png'
        m.save(mapFname)
        webbrowser.open(mapFname)
        time.sleep(2)
        image = pyautogui.screenshot()
        pyautogui.hotkey('ctrl', 'w')
        image = cv2.cvtColor(np.array(image),cv2.COLOR_RGB2BGR)
        cv2.imwrite(file_png, image)
    else:
        pass

#La función guardar_documento guarda el documento en la ruta de acuerdo con el nombre del documento que es el número de la OT
def guardar_documento(nombre_doc):
    folder = os.path.join('./Informes', nombre_doc)
    file_name = '{}.docx'.format(nombre_doc)
    file = os.path.join(folder, file_name)
    docx_tpl.save(file)
    route = "https://epmco.sharepoint.com/:w:/r/teams/experienciasoperativasmantenimiento/Documentos%20compartidos/PowerApps/Informes" + '//' + nombre_doc
    return(route)

#La función recomendaciones es la que realiza las recomendaciones de informes con similitudes encontradas
def recomendaciones(registro):
    contenido = []
    if informacion.loc[registro,'TIPO_ELEMENTO'] == "Redes Primarias":
        ipid_nuevo = informacion.loc[registro,'Ipid']
        diametro = informacion.loc[registro,'DIAMETRO']
        material_tuberia = informacion.loc[registro,'MATERIAL']
        longitud = informacion.loc[registro,'LONGITUD']

        control_1 = base[base['Ipid']==ipid_nuevo]
        control_2 = base[(base['DIAMETRO']==diametro)&(base['MATERIAL']==material_tuberia)&(base['LONGITUD']==longitud)]
        control_3 = base[(base['DIAMETRO']==diametro)&(base['MATERIAL']==material_tuberia)&(base['LONGITUD']!=longitud)]
        control_4 = base[(base['DIAMETRO']==diametro)&(base['MATERIAL']!=material_tuberia)&(base['LONGITUD']==longitud)]
        control_5 = base[(base['DIAMETRO']==diametro)&(base['MATERIAL']!=material_tuberia)&(base['LONGITUD']!=longitud)]
        control_6 = base[(base['DIAMETRO']!=diametro)&(base['MATERIAL']==material_tuberia)&(base['LONGITUD']==longitud)]
        control_7 = base[(base['DIAMETRO']!=diametro)&(base['MATERIAL']==material_tuberia)&(base['LONGITUD']!=longitud)]
        control_8 = base[(base['DIAMETRO']!=diametro)&(base['MATERIAL']!=material_tuberia)&(base['LONGITUD']==longitud)]

        if control_1.shape[0]>0:
            value = '100%'
            for exp in control_1.index:
                doc_link = control_1.loc[exp,'LINK_INFORME']
                contenido.append(doc_link)
        elif control_2.shape[0]>0:
            value = '80%'
            for exp in control_2.index:
                doc_link = control_2.loc[exp,'LINK_INFORME']
                contenido.append(doc_link)
        elif control_3.shape[0]>0:
            value = '60%'
            for exp in control_3.index:
                doc_link = control_3.loc[exp,'LINK_INFORME']
                contenido.append(doc_link)
        elif control_4.shape[0]>0:
            value = '50%'
            for exp in control_4.index:
                doc_link = control_4.loc[exp,'LINK_INFORME']
                contenido.append(doc_link)
        elif control_5.shape[0]>0:
            value = '40%'
            for exp in control_5.index:
                doc_link = control_5.loc[exp,'LINK_INFORME']
                contenido.append(doc_link)
        elif control_6.shape[0]>0:
            value = '30%'
            for exp in control_6.index:
                doc_link = control_6.loc[exp,'LINK_INFORME']
                contenido.append(doc_link)
        elif control_7.shape[0]>0:
            value = '20%'
            for exp in control_7.index:
                doc_link = control_7.loc[exp,'LINK_INFORME']
                contenido.append(doc_link)
        elif control_8.shape[0]>0:
            value = '10%'
            for exp in control_8.index:
                doc_link = control_8.loc[exp,'LINK_INFORME']
                contenido.append(doc_link)
        else:
            value = '0%'
            doc_link = "No se encontraron informes de experiencia similares"
            contenido.append(doc_link)

                
    elif informacion.loc[registro,'TIPO_ELEMENTO'] == "Valvula":
        ipid_nuevo = informacion.loc[registro,'Ipid']
        diametro = informacion.loc[registro,'DIAMETRO']
        funcion = informacion.loc[registro,'FUNCION_VALVULA']
        tipo = informacion.loc[registro,'TIPO_VALVULA']


        control_1 = base[base['Ipid']==ipid_nuevo]
        control_2 = base[(base['DIAMETRO']==diametro)&(base['FUNCION_VALVULA']==funcion)&(base['TIPO_VALVULA']==tipo)]
        control_3 = base[(base['DIAMETRO']==diametro)&(base['FUNCION_VALVULA']==funcion)&(base['TIPO_VALVULA']!=tipo)]
        control_4 = base[(base['DIAMETRO']==diametro)&(base['FUNCION_VALVULA']!=funcion)&(base['TIPO_VALVULA']==tipo)]
        control_5 = base[(base['DIAMETRO']==diametro)&(base['FUNCION_VALVULA']!=funcion)&(base['TIPO_VALVULA']!=tipo)]
        control_6 = base[(base['DIAMETRO']!=diametro)&(base['FUNCION_VALVULA']==funcion)&(base['TIPO_VALVULA']==tipo)]
        control_7 = base[(base['DIAMETRO']!=diametro)&(base['FUNCION_VALVULA']==funcion)&(base['TIPO_VALVULA']!=tipo)]
        control_8 = base[(base['DIAMETRO']!=diametro)&(base['FUNCION_VALVULA']!=funcion)&(base['TIPO_VALVULA']==tipo)]

        if control_1.shape[0]>0:
            value = '100%'
            for exp in control_1.index:
                doc_link = control_1.loc[exp,'LINK_INFORME']
                contenido.append(doc_link)
        elif control_2.shape[0]>0:
            value = '80%'
            for exp in control_2.index:
                doc_link = control_2.loc[exp,'LINK_INFORME']
                contenido.append(doc_link)
        elif control_3.shape[0]>0:
            value = '60%'
            for exp in control_3.index:
                doc_link = control_3.loc[exp,'LINK_INFORME']
                contenido.append(doc_link)
        elif control_4.shape[0]>0:
            value = '50%'
            for exp in control_4.index:
                doc_link = control_4.loc[exp,'LINK_INFORME']
                contenido.append(doc_link)
        elif control_5.shape[0]>0:
            value = '40%'
            for exp in control_5.index:
                doc_link = control_5.loc[exp,'LINK_INFORME']
                contenido.append(doc_link)
        elif control_6.shape[0]>0:
            value = '30%'
            for exp in control_6.index:
                doc_link = control_6.loc[exp,'LINK_INFORME']
                contenido.append(doc_link)
        elif control_7.shape[0]>0:
            value = '20%'
            for exp in control_7.index:
                doc_link = control_7.loc[exp,'LINK_INFORME']
                contenido.append(doc_link)
        elif control_8.shape[0]>0:
            value = '10%'
            for exp in control_8.index:
                doc_link = control_8.loc[exp,'LINK_INFORME']
                contenido.append(doc_link)
        else:
            value = '0%'
            doc_link = 'No se encontraron registros de experiencias con información similar'
            contenido.append(doc_link)


    elif informacion.loc[registro,'TIPO_ELEMENTO'] == "Redes Secundarias":
        ipid_nuevo = informacion.loc[registro,'Ipid']
        diametro = informacion.loc[registro,'DIAMETRO']
        material_tuberia_sec = informacion.loc[registro,'MATERIAL']
        longitud = informacion.loc[registro,'LONGITUD']
        profundidad = informacion.loc[registro,'PROFUNDIDAD']
    

        control_1 = base[base['Ipid']==ipid_nuevo]
        control_2 = base[(base['DIAMETRO']==diametro)&(base['MATERIAL']==material_tuberia_sec)&(base['LONGITUD']==longitud)&(base['PROFUNDIDAD']==profundidad)]
        control_3 = base[(base['DIAMETRO']==diametro)&(base['MATERIAL']==material_tuberia_sec)&(base['LONGITUD']==longitud)&(base['PROFUNDIDAD']!=profundidad)]
        control_4 = base[(base['DIAMETRO']==diametro)&(base['MATERIAL']==material_tuberia_sec)&(base['LONGITUD']!=longitud)&(base['PROFUNDIDAD']==profundidad)]
        control_5 = base[(base['DIAMETRO']==diametro)&(base['MATERIAL']!=material_tuberia_sec)&(base['LONGITUD']==longitud)&(base['PROFUNDIDAD']==profundidad)]
        control_6 = base[(base['DIAMETRO']==diametro)&(base['MATERIAL']==material_tuberia_sec)&(base['LONGITUD']!=longitud)&(base['PROFUNDIDAD']!=profundidad)]
        control_7 = base[(base['DIAMETRO']==diametro)&(base['MATERIAL']!=material_tuberia_sec)&(base['LONGITUD']!=longitud)&(base['PROFUNDIDAD']==profundidad)]
        control_8 = base[(base['DIAMETRO']==diametro)&(base['MATERIAL']!=material_tuberia_sec)&(base['LONGITUD']!=longitud)&(base['PROFUNDIDAD']!=profundidad)]
        control_9 = base[(base['DIAMETRO']!=diametro)&(base['MATERIAL']==material_tuberia_sec)&(base['LONGITUD']==longitud)&(base['PROFUNDIDAD']==profundidad)]
        control_10 = base[(base['DIAMETRO']!=diametro)&(base['MATERIAL']==material_tuberia_sec)&(base['LONGITUD']==longitud)&(base['PROFUNDIDAD']!=profundidad)]
        control_11 = base[(base['DIAMETRO']!=diametro)&(base['MATERIAL']==material_tuberia_sec)&(base['LONGITUD']!=longitud)&(base['PROFUNDIDAD']==profundidad)]
        control_12 = base[(base['DIAMETRO']!=diametro)&(base['MATERIAL']!=material_tuberia_sec)&(base['LONGITUD']==longitud)&(base['PROFUNDIDAD']==profundidad)]
        control_13 = base[(base['DIAMETRO']!=diametro)&(base['MATERIAL']==material_tuberia_sec)&(base['LONGITUD']!=longitud)&(base['PROFUNDIDAD']!=profundidad)]
        control_14 = base[(base['DIAMETRO']!=diametro)&(base['MATERIAL']!=material_tuberia_sec)&(base['LONGITUD']!=longitud)&(base['PROFUNDIDAD']==profundidad)]

        if control_1.shape[0]>0:
            value = '100%'
            for exp in control_1.index:
                doc_link = control_1.loc[exp,'LINK_INFORME']
                contenido.append(doc_link)
        elif control_2.shape[0]>0:
            value = '80%'
            for exp in control_2.index:
                doc_link = control_2.loc[exp,'LINK_INFORME']
                contenido.append(doc_link)
        elif control_3.shape[0]>0:
            value = '70%'
            for exp in control_3.index:
                doc_link = control_3.loc[exp,'LINK_INFORME']
                contenido.append(doc_link)
        elif control_4.shape[0]>0:
            value = '60%'
            for exp in control_4.index:
                doc_link = control_4.loc[exp,'LINK_INFORME']
                contenido.append(doc_link)
        elif control_5.shape[0]>0:
            value = '50%'
            for exp in control_5.index:
                doc_link = control_5.loc[exp,'LINK_INFORME']
                contenido.append(doc_link)
        elif control_6.shape[0]>0:
            value = '40%'
            for exp in control_6.index:
                doc_link = control_6.loc[exp,'LINK_INFORME']
                contenido.append(doc_link)
        elif control_7.shape[0]>0:
            value = '35%'
            for exp in control_7.index:
                doc_link = control_7.loc[exp,'LINK_INFORME']
                contenido.append(doc_link)
        elif control_8.shape[0]>0:
            value = '30%'
            for exp in control_8.index:
                doc_link = control_8.loc[exp,'LINK_INFORME']
                contenido.append(doc_link)
        elif control_9.shape[0]>0:
            value = '25%'
            for exp in control_9.index:
                doc_link = control_9.loc[exp,'LINK_INFORME']
                contenido.append(doc_link)
        elif control_10.shape[0]>0:
            value = '20%'
            for exp in control_10.index:
                doc_link = control_10.loc[exp,'LINK_INFORME']
                contenido.append(doc_link)
        elif control_11.shape[0]>0:
            value = '15%'
            for exp in control_11.index:
                doc_link = control_11.loc[exp,'LINK_INFORME']
                contenido.append(doc_link)
        elif control_12.shape[0]>0:
            value = '15%'
            for exp in control_12.index:
                doc_link = control_12.loc[exp,'LINK_INFORME']
                contenido.append(doc_link)
        elif control_13.shape[0]>0:
            value = '10%'
            for exp in control_13.index:
                doc_link = control_13.loc[exp,'LINK_INFORME']
                contenido.append(doc_link)
        elif control_14.shape[0]>0:
            value = '10%'
            for exp in control_14.index:
                doc_link = control_14.loc[exp,'LINK_INFORME']
                contenido.append(doc_link)
        else:
            value = '0%'
            doc_link = 'No se encontraron registros de experiencias con información similar'
            contenido.append(doc_link)
    else:
        value = '0%'
        contenido.append('IPID erroneo')
        
    return(value, contenido)


##La función enviar_correo se encarga de enviar al correo indicado la notificación de que se creó el informe de experiencia operativa, con la respectiva ruta, 
#y con las similitudes encontradas con sus respectivos enlaces
def enviar_correo(registro, link_informeexp):
    nombre = informacion.loc[registro, 'Funcionario que Registra']
    correo = directorio_correos['CORREO'][directorio_correos['NOMBRE']==nombre]
    if correo.empty:
        correo = 'SERGIO.QUINTERO.RAMIREZZ@EPM.COM.CO'
    valor_similitud = recomendaciones(registro)[0]
    links_antecedentes = recomendaciones(registro)[1]
    num_antecedentes = len(links_antecedentes)

    email_sender = 'experiencias.operativas.epm@gmail.com'
    email_password = 'njfwwqdzonwmxnxa'
    email_receiver = correo

    subject = "Informe experiencias operativas Agua y Saneamiento"
    if num_antecedentes == 1:
        body ='''\
            Se ha creado el informe de experiencias operativas, al cual podrá acceder para terminar de diligenciar en la siguiente ruta: "\n
            {ruta}.
            \n
            Se han encontrado los siguientes antecedentes con similaridad al evento actual:\n
            Con una similitud de: {valor}
            \n
            {antecedentes}
            '''.format(ruta=link_informeexp, valor=valor_similitud, antecedentes=links_antecedentes[0])
    elif num_antecedentes == 2:
        body ='''\
            Se ha creado el informe de experiencia operativa, el cual podrá entrar y terminar de diligenciar en la siguiente ruta: "\n
            {ruta}.
            \n
            Se han encontrado los siguientes antecedentes con similaridad al evento actual:\n
            Con una similitud de: {valor}
            \n
            {antecedente1}
            \n
            {antecedente2}
            '''.format(ruta=link_informeexp, valor=valor_similitud, antecedente1=links_antecedentes[0],antecedente2=links_antecedentes[1])
    else:
        body ='''\
            Se ha creado el informe de experiencia operativa, el cual podrá entrar y terminar de diligenciar en la siguiente ruta: "\n
            {ruta}.
            \n
            Se han encontrado los siguientes antecedentes con similaridad al evento actual:\n
            Con una similitud de: {valor}
            \n
            {antecedente1}
            \n
            {antecedente2}
            \n
            {antecedente3}
            '''.format(ruta=link_informeexp, valor=valor_similitud, antecedente1=links_antecedentes[0],antecedente2=links_antecedentes[1],antecedente3=links_antecedentes[2])


    em = EmailMessage()
    em['From'] = email_sender
    em['To'] = email_receiver
    em['Subject'] = subject
    em.set_content(body)

    context = ssl.create_default_context()

    with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
        smtp.login(email_sender, email_password)
        smtp.sendmail(email_sender, email_receiver, em.as_string())

if __name__ == "__main__":
    principal()
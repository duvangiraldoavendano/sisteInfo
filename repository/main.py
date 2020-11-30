import csv
from datetime import datetime
from openpyxl import load_workbook

from cedulas import *

#Funcion que obtiene la tabla de PROFESORES LISTA para CIENCIAS POLÍTICAS
from profesores_listado import *
#Funcion que obtiene la tabla de PROFESORES PROGRAMA
from profesore_programa import *
#Funcion que toma los profesores que son comunes tienen horas en CP y DERECHO y discrimina sus horas de trabajo por pregrado.
from discriminacion_horas_trabajo import *
##Funcion que obtiene la tabla de PROFESORES LISTA para DERECHO
from cienciaspoliticas import *

'''
#----------------------------------------------------PARA EL SEMESTRE 20192------------------------------------------------------------

anio=2019
semest=2

#------------------------------DERECHO-------------------------------
prof_institu="20192/reporte_titulos_instituto_20192.xlsx"
programacion='20192/SISTEMATIZACIÓN20192.LIS'
cedula_cp,cedulas_dere,lineas_cedulas=cedulas_progra(programacion)
profes_cat='20192/reporte_titulos_cat_facultad_20192.xlsx'
archivo_resul="20192/profe_der20192.xlsx"
profes_vincu="20192/reporte_titulos_facultad_20192.xlsx"
tipo_cc="20192/reporte_catedras_facultad_20192.xlsx"
reporte_vinculados="20192/reporte_profesores_facultad_20192.xlsx"
planes_derecho="20192/Planes_2019_2_Derecho.xlsx"


#-------------------CIENCIAS POLÍTICAS-------------------------------
prof_institu_cp="20192/reporte_titulos_instituto_20192.xlsx"
programacion_cp='20192/SISTEMATIZACIÓN20192.LIS'
profes_cat_cp='20192/reporte_titulos_cat_facultad_20192.xlsx'
archivo_resul_cp="20192/profe_cp20192.xlsx"
profes_vincu_cp="20192/reporte_titulos_facultad_20192.xlsx"
tipo_cc_cp="20192/reporte_catedras_facultad_20192.xlsx"
tipo_cc_vincu_cp="20192/reporte_profesores_facultad_20192.xlsx"
tipo_cc_inst_cp="20192/reporte_profesores_instituto_20192.xlsx"


discriminacion_horas_trabajo(cedulas_dere, cedula_cp, prof_institu, profes_vincu, profes_cat, prof_institu_cp, profes_vincu_cp, profes_cat_cp, lineas_cedulas)

ciencias_politicas(profes_cat_cp,profes_vincu_cp,prof_institu_cp,archivo_resul_cp,cedula_cp,tipo_cc_cp,anio,semest,tipo_cc_vincu_cp,tipo_cc_inst_cp,lineas_cedulas,reporte_vinculados)

catedra(profes_cat,profes_vincu,prof_institu,archivo_resul,cedulas_dere,tipo_cc,anio,semest,reporte_vinculados,planes_derecho,lineas_cedulas)

#Estadística para los profesores de derecho
estadistica_profesores(anio,semest,archivo_resul,tabla_profe_progra_derecho)

#Estadística para los profesores de ciencias politicas
estadistica_profesores(anio,semest,archivo_resul_cp,tabla_profe_progra_cp)
'''

#----------------------------------------------------PARA EL SEMESTRE 20191------------------------------------------------------------
anio=2019
semest=1

#------------------------------DERECHO-------------------------------
prof_institu="20191/reporte_titulos_instituto_20191.xlsx"
programacion='20191/SISTEMATIZACIÓN20191.LIS'
cedula_cp,cedulas_dere,lineas_cedulas,cedulas_DAD=cedulas_progra(programacion)
profes_cat='20191/reporte_titulos_cat_facultad_20191.xlsx'
archivo_resul="20191/profe_der20191.xlsx"
profes_vincu="20191/reporte_titulos_facultad_20191.xlsx"
tipo_cc="20191/reporte_catedras_facultad_20191.xlsx"
reporte_vinculados="20191/reporte_profesores_facultad_20191.xlsx"
planes_derecho="20191/Planes_2019_1_Derecho.xlsx"
tabla_profe_progra_derecho="estadistica/profesores_programa_derecho.xlsx"

#-------------------CIENCIAS POLÍTICAS-------------------------------
prof_institu_cp="20191/reporte_titulos_instituto_20191.xlsx"
programacion_cp='20191/SISTEMATIZACIÓN20191.LIS'
profes_cat_cp='20191/reporte_titulos_cat_facultad_20191.xlsx'
archivo_resul_cp="20191/profe_cp20191.xlsx"
profes_vincu_cp="20191/reporte_titulos_facultad_20191.xlsx"
tipo_cc_cp="20191/reporte_catedras_facultad_20191.xlsx"
tipo_cc_vincu_cp="20191/reporte_profesores_facultad_20191.xlsx"
tipo_cc_inst_cp="20191/reporte_profesores_instituto_20191.xlsx"
tabla_profe_progra_cp="estadistica/profesores_programa_cp.xlsx"
excel_discriminacion="horas_docentes_Derecho_2017_2019.xlsx"

cate_cpa_horas_cat,cate_cpa_horas_plan,vincu_cpa_horas_cat,vincu_cpa_horas_plan,cate_der_horas_cat,cate_der_horas_plan,vincu_der_horas_cat,vincu_der_horas_plan=discriminacion_horas_trabajo(cedulas_dere, cedula_cp, prof_institu, profes_vincu, profes_cat, prof_institu_cp, profes_vincu_cp, profes_cat_cp,lineas_cedulas,excel_discriminacion,anio,semest)

ciencias_politicas(profes_cat_cp,profes_vincu_cp,
                   prof_institu_cp,archivo_resul_cp,
                   cedula_cp,tipo_cc_cp,
                   anio,semest,tipo_cc_vincu_cp,
                   tipo_cc_inst_cp,lineas_cedulas,
                   reporte_vinculados,planes_derecho,
                   cate_cpa_horas_cat,cate_cpa_horas_plan,
                   vincu_cpa_horas_cat,vincu_cpa_horas_plan, excel_discriminacion,
                   cate_der_horas_cat,cate_der_horas_plan,vincu_der_horas_cat,vincu_der_horas_plan)

catedra(profes_cat,profes_vincu,
            prof_institu,archivo_resul,
            cedulas_dere,tipo_cc,anio,
            semest,reporte_vinculados,
            planes_derecho,lineas_cedulas,
            cate_der_horas_cat,cate_der_horas_plan,
            vincu_der_horas_cat,vincu_der_horas_plan,excel_discriminacion,
            cate_cpa_horas_cat,cate_cpa_horas_plan,vincu_cpa_horas_cat,vincu_cpa_horas_plan)

#Estadística para los profesores de derecho
limpiar(tabla_profe_progra_derecho)
estadistica_profesores(anio,semest,archivo_resul,tabla_profe_progra_derecho)

#Estadística para los profesores de ciencias politicas
limpiar(tabla_profe_progra_cp)
estadistica_profesores(anio,semest,archivo_resul_cp,tabla_profe_progra_cp)

#-----------------------POSGRADOS----------------------------------
archivo_resul_DAD="20191/DAD/profe_DAD20191.xlsx"
#catedra(profes_cat,profes_vincu,prof_institu,archivo_resul_DAD,cedulas_DAD,tipo_cc,anio,semest,reporte_vinculados,planes_derecho,lineas_cedulas)

#----------------------------------------------------PARA EL SEMESTRE 20182------------------------------------------------------------
anio=2018
semest=2

#------------------------------DERECHO-------------------------------
prof_institu="20182/reporte_titulos_instituto_20182.xlsx"
programacion='20182/SISTEMATIZACIÓN20182.LIS'
cedula_cp,cedulas_dere,lineas_cedulas,cedulas_DAD=cedulas_progra(programacion)
profes_cat='20182/reporte_titulos_cat_facultad_20182.xlsx'
archivo_resul="20182/profe_der20182.xlsx"
profes_vincu="20182/reporte_titulos_facultad_20182.xlsx"
tipo_cc="20182/reporte_catedras_facultad_20182.xlsx"
reporte_vinculados="20182/reporte_profesores_facultad_20182.xlsx"
planes_derecho="20182/Planes_2018_2_Derecho.xlsx"

#-------------------CIENCIAS POLÍTICAS-------------------------------
prof_institu_cp="20182/reporte_titulos_instituto_20182.xlsx"
programacion_cp='20182/SISTEMATIZACIÓN20182.LIS'
profes_cat_cp='20182/reporte_titulos_cat_facultad_20182.xlsx'
archivo_resul_cp="20182/profe_cp20182.xlsx"
profes_vincu_cp="20182/reporte_titulos_facultad_20182.xlsx"
tipo_cc_cp="20182/reporte_catedras_facultad_20182.xlsx"
tipo_cc_vincu_cp="20182/reporte_profesores_facultad_20182.xlsx"
tipo_cc_inst_cp="20182/reporte_profesores_instituto_20182.xlsx"

cate_cpa_horas_cat,cate_cpa_horas_plan,vincu_cpa_horas_cat,vincu_cpa_horas_plan,cate_der_horas_cat,cate_der_horas_plan,vincu_der_horas_cat,vincu_der_horas_plan=discriminacion_horas_trabajo(cedulas_dere, cedula_cp, prof_institu, profes_vincu, profes_cat, prof_institu_cp, profes_vincu_cp, profes_cat_cp,lineas_cedulas,excel_discriminacion,anio,semest)

ciencias_politicas(profes_cat_cp,profes_vincu_cp,
                   prof_institu_cp,archivo_resul_cp,
                   cedula_cp,tipo_cc_cp,
                   anio,semest,tipo_cc_vincu_cp,
                   tipo_cc_inst_cp,lineas_cedulas,
                   reporte_vinculados,planes_derecho,
                   cate_cpa_horas_cat,cate_cpa_horas_plan,
                   vincu_cpa_horas_cat,vincu_cpa_horas_plan,excel_discriminacion,
                   cate_der_horas_cat,cate_der_horas_plan,vincu_der_horas_cat,vincu_der_horas_plan)

catedra(profes_cat,profes_vincu,
            prof_institu,archivo_resul,
            cedulas_dere,tipo_cc,anio,
            semest,reporte_vinculados,
            planes_derecho,lineas_cedulas,
            cate_der_horas_cat,cate_der_horas_plan,
            vincu_der_horas_cat,vincu_der_horas_plan,excel_discriminacion,
            cate_cpa_horas_cat,cate_cpa_horas_plan,vincu_cpa_horas_cat,vincu_cpa_horas_plan)

#Estadística para los profesores de derecho
estadistica_profesores(anio,semest,archivo_resul,tabla_profe_progra_derecho)

#Estadística para los profesores de ciencias politicas
estadistica_profesores(anio,semest,archivo_resul_cp,tabla_profe_progra_cp)

#---------------------------POSGRADOS--------------------------------
archivo_resul_DAD="20182/DAD/profe_DAD20182.xlsx"
#catedra(profes_cat,profes_vincu,prof_institu,archivo_resul_DAD,cedulas_DAD,tipo_cc,anio,semest,reporte_vinculados,planes_derecho,lineas_cedulas)

#----------------------------------------------------PARA EL SEMESTRE 20181------------------------------------------------------------
anio=2018
semest=1

#------------------------------DERECHO-------------------------------
prof_institu="20181/reporte_titulos_instituto_20181.xlsx"
programacion='20181/SISTEMATIZACIÓN20181.LIS'
cedula_cp,cedulas_dere,lineas_cedulas,cedulas_DAD=cedulas_progra(programacion)
profes_cat='20181/reporte_titulos_cat_facultad_20181.xlsx'
archivo_resul="20181/profe_der20181.xlsx"
profes_vincu="20181/reporte_titulos_facultad_20181.xlsx"
tipo_cc="20181/reporte_catedras_facultad_20181.xlsx"
reporte_vinculados="20181/reporte_profesores_facultad_20181.xlsx"
planes_derecho="20181/Planes_2018_1_Derecho.xlsx"

#-------------------CIENCIAS POLÍTICAS-------------------------------
prof_institu_cp="20181/reporte_titulos_instituto_20181.xlsx"
programacion_cp='20181/SISTEMATIZACIÓN20181.LIS'
profes_cat_cp='20181/reporte_titulos_cat_facultad_20181.xlsx'
archivo_resul_cp="20181/profe_cp20181.xlsx"
profes_vincu_cp="20181/reporte_titulos_facultad_20181.xlsx"
tipo_cc_cp="20181/reporte_catedras_facultad_20181.xlsx"
tipo_cc_vincu_cp="20181/reporte_profesores_facultad_20181.xlsx"
tipo_cc_inst_cp="20181/reporte_profesores_instituto_20181.xlsx"

cate_cpa_horas_cat,cate_cpa_horas_plan,vincu_cpa_horas_cat,vincu_cpa_horas_plan,cate_der_horas_cat,cate_der_horas_plan,vincu_der_horas_cat,vincu_der_horas_plan=discriminacion_horas_trabajo(cedulas_dere, cedula_cp, prof_institu, profes_vincu, profes_cat, prof_institu_cp, profes_vincu_cp, profes_cat_cp,lineas_cedulas,excel_discriminacion,anio,semest)

ciencias_politicas(profes_cat_cp,profes_vincu_cp,
                   prof_institu_cp,archivo_resul_cp,
                   cedula_cp,tipo_cc_cp,
                   anio,semest,tipo_cc_vincu_cp,
                   tipo_cc_inst_cp,lineas_cedulas,
                   reporte_vinculados,planes_derecho,
                   cate_cpa_horas_cat,cate_cpa_horas_plan,
                   vincu_cpa_horas_cat,vincu_cpa_horas_plan,excel_discriminacion,
                   cate_der_horas_cat,cate_der_horas_plan,vincu_der_horas_cat,vincu_der_horas_plan)

catedra(profes_cat,profes_vincu,
            prof_institu,archivo_resul,
            cedulas_dere,tipo_cc,anio,
            semest,reporte_vinculados,
            planes_derecho,lineas_cedulas,
            cate_der_horas_cat,cate_der_horas_plan,
            vincu_der_horas_cat,vincu_der_horas_plan,excel_discriminacion,
            cate_cpa_horas_cat,cate_cpa_horas_plan,vincu_cpa_horas_cat,vincu_cpa_horas_plan)

#Estadística para los profesores de derecho
estadistica_profesores(anio,semest,archivo_resul,tabla_profe_progra_derecho)

#Estadística para los profesores de ciencias politicas
estadistica_profesores(anio,semest,archivo_resul_cp,tabla_profe_progra_cp)

#----------------------------------------------------PARA EL SEMESTRE 20172------------------------------------------------------------
anio=2017
semest=2

#------------------------------DERECHO-------------------------------
prof_institu="20172/reporte_titulos_instituto_20172.xlsx"
programacion='20172/SISTEMATIZACIÓN20172DERCP.LIS'
cedula_cp,cedulas_dere,lineas_cedulas,cedulas_DAD=cedulas_progra(programacion)
profes_cat='20172/reporte_titulos_cat_facultad_20172.xlsx'
archivo_resul="20172/profe_der20172.xlsx"
profes_vincu="20172/reporte_titulos_facultad_20172.xlsx"
tipo_cc="20172/reporte_catedras_facultad_20172.xlsx"
reporte_vinculados="20172/reporte_profesores_facultad_20172.xlsx"
planes_derecho="20172/Planes_2017_2_Derecho.xlsx"

#-------------------CIENCIAS POLÍTICAS-------------------------------
prof_institu_cp="20172/reporte_titulos_instituto_20172.xlsx"
programacion_cp='20172/SISTEMATIZACIÓN20172DERCP.LIS'
profes_cat_cp='20172/reporte_titulos_cat_facultad_20172.xlsx'
archivo_resul_cp="20172/profe_cp20172.xlsx"
profes_vincu_cp="20172/reporte_titulos_facultad_20172.xlsx"
tipo_cc_cp="20172/reporte_catedras_facultad_20172.xlsx"
tipo_cc_vincu_cp="20172/reporte_profesores_facultad_20172.xlsx"
tipo_cc_inst_cp="20172/reporte_profesores_instituto_20172.xlsx"

cate_cpa_horas_cat,cate_cpa_horas_plan,vincu_cpa_horas_cat,vincu_cpa_horas_plan,cate_der_horas_cat,cate_der_horas_plan,vincu_der_horas_cat,vincu_der_horas_plan=discriminacion_horas_trabajo(cedulas_dere, cedula_cp, prof_institu, profes_vincu, profes_cat, prof_institu_cp, profes_vincu_cp, profes_cat_cp,lineas_cedulas,excel_discriminacion,anio,semest)

ciencias_politicas(profes_cat_cp,profes_vincu_cp,
                   prof_institu_cp,archivo_resul_cp,
                   cedula_cp,tipo_cc_cp,
                   anio,semest,tipo_cc_vincu_cp,
                   tipo_cc_inst_cp,lineas_cedulas,
                   reporte_vinculados,planes_derecho,
                   cate_cpa_horas_cat,cate_cpa_horas_plan,
                   vincu_cpa_horas_cat,vincu_cpa_horas_plan,excel_discriminacion,
                   cate_der_horas_cat,cate_der_horas_plan,vincu_der_horas_cat,vincu_der_horas_plan)

catedra(profes_cat,profes_vincu,
            prof_institu,archivo_resul,
            cedulas_dere,tipo_cc,anio,
            semest,reporte_vinculados,
            planes_derecho,lineas_cedulas,
            cate_der_horas_cat,cate_der_horas_plan,
            vincu_der_horas_cat,vincu_der_horas_plan,excel_discriminacion,
            cate_cpa_horas_cat,cate_cpa_horas_plan,vincu_cpa_horas_cat,vincu_cpa_horas_plan)

#Estadística para los profesores de derecho
estadistica_profesores(anio,semest,archivo_resul,tabla_profe_progra_derecho)

#Estadística para los profesores de ciencias politicas
estadistica_profesores(anio,semest,archivo_resul_cp,tabla_profe_progra_cp)

#----------------------------------------------------PARA EL SEMESTRE 20171------------------------------------------------------------
anio=2017
semest=1

#------------------------------DERECHO-------------------------------
prof_institu="20171/reporte_titulos_instituto_20171.xlsx"
programacion='20171/SISTEMATIZACIÓN20171.LIS'
cedula_cp,cedulas_dere,lineas_cedulas,cedulas_DAD=cedulas_progra(programacion)
profes_cat='20171/reporte_titulos_cat_facultad_20171.xlsx'
archivo_resul="20171/profe_der20171.xlsx"
profes_vincu="20171/reporte_titulos_facultad_20171.xlsx"
tipo_cc="20171/reporte_catedras_facultad_20171.xlsx"
reporte_vinculados="20171/reporte_profesores_facultad_20171.xlsx"
planes_derecho="20171/Planes_2017_1_Derecho.xlsx"

#-------------------CIENCIAS POLÍTICAS-------------------------------
prof_institu_cp="20171/reporte_titulos_instituto_20171.xlsx"
programacion_cp='20171/SISTEMATIZACIÓN20171.LIS'
profes_cat_cp='20171/reporte_titulos_cat_facultad_20171.xlsx'
archivo_resul_cp="20171/profe_cp20171.xlsx"
profes_vincu_cp="20171/reporte_titulos_facultad_20171.xlsx"
tipo_cc_cp="20171/reporte_catedras_facultad_20171.xlsx"
tipo_cc_vincu_cp="20171/reporte_profesores_facultad_20171.xlsx"
tipo_cc_inst_cp="20171/reporte_profesores_instituto_20171.xlsx"

cate_cpa_horas_cat,cate_cpa_horas_plan,vincu_cpa_horas_cat,vincu_cpa_horas_plan,cate_der_horas_cat,cate_der_horas_plan,vincu_der_horas_cat,vincu_der_horas_plan=discriminacion_horas_trabajo(cedulas_dere, cedula_cp, prof_institu, profes_vincu, profes_cat, prof_institu_cp, profes_vincu_cp, profes_cat_cp,lineas_cedulas,excel_discriminacion,anio,semest)

ciencias_politicas(profes_cat_cp,profes_vincu_cp,
                   prof_institu_cp,archivo_resul_cp,
                   cedula_cp,tipo_cc_cp,
                   anio,semest,tipo_cc_vincu_cp,
                   tipo_cc_inst_cp,lineas_cedulas,
                   reporte_vinculados,planes_derecho,
                   cate_cpa_horas_cat,cate_cpa_horas_plan,
                   vincu_cpa_horas_cat,vincu_cpa_horas_plan,excel_discriminacion,
                   cate_der_horas_cat,cate_der_horas_plan,vincu_der_horas_cat,vincu_der_horas_plan)

catedra(profes_cat,profes_vincu,
            prof_institu,archivo_resul,
            cedulas_dere,tipo_cc,anio,
            semest,reporte_vinculados,
            planes_derecho,lineas_cedulas,
            cate_der_horas_cat,cate_der_horas_plan,
            vincu_der_horas_cat,vincu_der_horas_plan,excel_discriminacion,
            cate_cpa_horas_cat,cate_cpa_horas_plan,vincu_cpa_horas_cat,vincu_cpa_horas_plan)

#Estadística para los profesores de derecho
estadistica_profesores(anio,semest,archivo_resul,tabla_profe_progra_derecho)

#Estadística para los profesores de ciencias politicas
estadistica_profesores(anio,semest,archivo_resul_cp,tabla_profe_progra_cp)

#-----------------------POSGRADOS----------------------------------
archivo_resul_DAD="20171/DAD/profe_DAD20171.xlsx"
#catedra(profes_cat,profes_vincu,prof_institu,archivo_resul_DAD,cedulas_DAD,tipo_cc,anio,semest,reporte_vinculados,planes_derecho,lineas_cedulas)

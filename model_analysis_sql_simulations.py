# -*- coding: utf-8 -*-
"""
Created on Tue Nov  3 16:43:45 2020

@author: carriles
"""



##############################################################################
## SIMULACIÓN DEL CASO BASE .SAV Y GUARDADO EN SQL
## lECTURA DE ESCENARIOS EN SQL, SIMULACIÓN Y GUARDADO DE DATOS
##############################################################################



##############################################################################
## Version de Python                                                                    ##
##############################################################################
import sys
print(sys.version)

##############################################################################
## Instrucciones que permiten la ejecucion de los comandos de Python sin
## arrancar PSS/E
##############################################################################
PSSEVERSION=34

import os
import sys
import warnings
warnings.simplefilter("error")
warnings.filterwarnings('ignore', category=PendingDeprecationWarning)
import numpy as np
import pandas as pd
from pandas import DataFrame
import xlsxwriter
import pyodbc 
import math


##############################################################################
## Definir ruta de la libreria de python y de los ejecutables de PSS/E                  ##
##############################################################################
sys_path_PSSE=r'C:\Program Files (x86)\PTI\PSSE34\PSSPY27'
#or where else you find the psspy.pyc
sys.path.append(sys_path_PSSE)
os_path_PSSE=r'C:\Program Files (x86)\PTI\PSSE34\PSSBIN'
# or where else you find the psse.exe
os.environ['PATH'] += ';' + os_path_PSSE
import psspy
import redirect

redirect.psse2py()
psspy.psseinit(50000)
_i=psspy.getdefaultint()
_f=psspy.getdefaultreal()
_s=psspy.getdefaultchar()


##############################################################################
## Instrucciones para leer los casos de estudio y para simular                        ##
##############################################################################
#ruta= r"C:\mario\trabajos2\viesgo_applus_escenarios_red\simulacion\psse"
ruta= r"C:\David\PSSE_Viesgo"
#CASOraw= ruta + r"RDF_Caso_Base_2019_Pcc_wind10_2.raw"
CASOsav= ruta + r"\RDF_Caso_Base_2019_Pcc_wind10_2.sav"
SALIDALINEA= ruta + r"\Salida.txt"
psspy.case(CASOsav)
#psspy.fnsl([0,0,0,1,1,1,99,0])
#psspy.save(CASOsav)

#Volvemos a simular
psspy.fdns([0,0,0,1,1,1,99,0])
U=psspy.solv()


archivo_lineas = ruta + r"\datos_lineas132_zona_Viesgo.csv"

#############################################################################
## lectura de archivos con la información de las líneas.
#############################################################################
df_lineas132_viesgo = pd.read_csv(archivo_lineas, encoding='Latin9', header=0, sep=';', quotechar='\"')
           


#############################################################################
## Recorrer las líneas de 132 kV de la zona Viesgo y guardar los resultados en
## las tablas SQL
#############################################################################
 #Definición de la conexión con la DB
conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=193.144.190.81;'
                      'Database=Simulaciones_2020;'
                      'UID=user;'
                      'PWD=1234')
                      #'Trusted_Connection=yes;')

cursor = conn.cursor()

#for lineas in df_lineas132_viesgo.itertuples():
for idx, lineas in df_lineas132_viesgo.iterrows():
#    print(lineas)
    #No entiendo porque con itertuples lineas[0] es un índice.
#    print(int(lineas['Linea ID']))
#    print(type(lineas[0]))
#    ierr_AMPS, rval_AMPS = psspy.brnmsc(int(lineas[1]),int(lineas[4]), str(lineas[7]), 'AMPS')
    ierr_AMPS, rval_AMPS = psspy.brnmsc(int(lineas['From Bus Number']),int(lineas['To Bus Number']), str(lineas['Linea ID']), 'AMPS')
    ierr_P, rval_P = psspy.brnmsc(int(lineas['From Bus Number']),int(lineas['To Bus Number']), str(lineas['Linea ID']), 'P')
    ierr_Q, rval_Q = psspy.brnmsc(int(lineas['From Bus Number']),int(lineas['To Bus Number']), str(lineas['Linea ID']), 'Q')
    #ierr_MVA, rval_MVA = psspy.brnmsc(int(lineas['From Bus Number']),int(lineas['To Bus Number']), str(lineas['Linea ID']), 'PCTCPA')
    rate_MVA_Perc = 0
    if int(lineas['Rate1 (MVA)']) is not 0:
        rate_MVA_Perc = ((math.sqrt(3)*132000*rval_AMPS)/1000000)/int(lineas['Rate1 (MVA)'])*100
#    print(ierr_AMPS,rval_AMPS)
#    print(ierr_P, rval_P )
#    print(ierr_Q, rval_Q )
    
#    nombre_tabla = "OUTPUT_" + lineas['Subestacion 1'].strip(' ').replace(' ','_').replace('.','_') + "_" + lineas['Subestacion 2'].strip(' ').replace(' ','_').replace('.','_') + "_" + str(lineas['Linea ID'])
#    instruccion_insert = "INSERT INTO " + nombre_tabla + " (ID_Escenario, AMPS, P, Q) VALUES (-1, " + str(rval_AMPS) + ", " + str(rval_P) + ", " + str(rval_Q) + ");"
#    print(instruccion_insert)
    nombre_tabla = "OUTPUT_AMPS_P_Q_RATE_LAT_132"
    instruccion_insert = "INSERT INTO " + nombre_tabla + " (ID_Escenario, From_Bus_Number, Subestacion_1, To_Bus_Number, Subestacion_2, Linea_ID, AMPS, P_MW, Q_MVAR, RATE_PERC) VALUES (-1, " + str(lineas['From Bus Number']) + ", '" + str(lineas['Subestacion 1'].strip(' ').replace(' ','_').replace('.','_')) + "', " + str(lineas['To Bus Number']) + ", '" + str(lineas['Subestacion 2'].strip(' ').replace(' ','_').replace('.','_')) + "', '" + str(lineas['Linea ID']) + "', " + str(rval_AMPS) + ", " + str(rval_P) + ", " + str(rval_Q) + ", " + str(rate_MVA_Perc) + ");"#str(rval_AMPS*100/lineas['Rate1 (MVA)']) + ");"
    
    cursor.execute(instruccion_insert)
    conn.commit()
    
    




#############################################################################
## Recorrer las líneas de 132 kV de la zona Viesgo y obtener la tensión en todos
## los buses
#############################################################################
TENSION_BUSES_132 = []
for idx, lineas in df_lineas132_viesgo.iterrows():
    ierr_orig, rval_orig = psspy.busdat(lineas['From Bus Number'], 'KV')
    ierr_dest, rval_dest = psspy.busdat(lineas['To Bus Number'], 'KV')
    TENSION_BUSES_132.append([int(lineas['From Bus Number']), str(lineas['Subestacion 1']), ierr_orig, rval_orig])
    TENSION_BUSES_132.append([int(lineas['To Bus Number']), str(lineas['Subestacion 2']), ierr_dest, rval_dest])
#Se eliminan los valores duplicados
#TENSION_BUSES_132 = list(dict.fromkeys(TENSION_BUSES_132))
#TENSION_BUSES_132 = list(set(TENSION_BUSES_132))
temp = []
[temp.append(x) for x in TENSION_BUSES_132 if x not in temp]
#Se ordena por número de bus
TENSION_BUSES_132 = sorted(temp, key=lambda num_bus : num_bus[0])


#Se recorre la lista de buses para guardar las tensiones en el SQL
for nb in TENSION_BUSES_132:
#    print(nb[0])
    nombre_tabla = "OUTPUT_KV_BUSES_132"
    instruccion_insert = "INSERT INTO " + nombre_tabla + " (ID_Escenario, Bus_Number, Subestacion, KV) VALUES (-1, " + str(nb[0]) + ", '" + str(nb[1].strip(' ').replace(' ','_').replace('.','_')) + "', " + str(nb[3]) + ");"
#    print(instruccion_insert)
    cursor.execute(instruccion_insert)
    conn.commit()


cursor.close()
del cursor








#############################################################################
## Leer valores de potencia del SQL, cambiarlos en los generadores, simular y 
##guardar resultados en el SQL.
#############################################################################

#Extracción de datos del SQL
cursor = conn.cursor()
nombre_tabla_info = "INPUT_ESCENARIOS_GENERADORES"
num_escenarios = 3 #[1, 2, 3]
for id_esc in range(0, num_escenarios):#num_escenarios:
    SQL_Query = pd.read_sql_query('''SELECT ALL [ID_Escenario],[Bus_Number],[Bus_Name],[ID_Generador],[Porc_simulacion],[P_MW_simulacion] FROM [Simulaciones_2020].[dbo].[''' + nombre_tabla_info + '''] where ID_Escenario = ''' + str(id_esc), conn)
    df_escenario = pd.DataFrame(SQL_Query, columns=['ID_Escenario','Bus_Number','BUS_Name','ID_Generador','P_MW_simulacion'])
    
    #Se comprueba si no hay escenario para es id_escenario y la tabla está vacía
    if not df_escenario.empty:
        #Cambio de potencia en los generadores
        for idx, generadores in df_escenario.iterrows():
            print(str(generadores['Bus_Number']), str(generadores['ID_Generador']),str(generadores['P_MW_simulacion']))
#            filtro = GENERADORES_132_DF.loc[(GENERADORES_132_DF['Bus_Number'] == generadores['Bus_Number']) & (GENERADORES_132_DF['ID_Generador'] == str(generadores['ID_Generador']))]
#            pot_nueva = filtro['PMAX_MW'][0] * generadores['P_MW_simulacion']
            psspy.machine_data_2(generadores['Bus_Number'],str(generadores['ID_Generador']),[psspy.getdefaultint()],[generadores['P_MW_simulacion']])
        
        #Volvemos a simular
        psspy.fdns([0,0,0,1,1,1,99,0])
        U=psspy.solv()
        
        
        #Se obtienen los valores de corriente y potencias y se guardan en el SQL
        for idx, lineas in df_lineas132_viesgo.iterrows():
        #    print(lineas)
            #No entiendo porque con itertuples lineas[0] es un índice.
        #    print(int(lineas['Linea ID']))
        #    print(type(lineas[0]))
        #    ierr_AMPS, rval_AMPS = psspy.brnmsc(int(lineas[1]),int(lineas[4]), str(lineas[7]), 'AMPS')
            ierr_AMPS, rval_AMPS = psspy.brnmsc(int(lineas['From Bus Number']),int(lineas['To Bus Number']), str(lineas['Linea ID']), 'AMPS')
            ierr_P, rval_P = psspy.brnmsc(int(lineas['From Bus Number']),int(lineas['To Bus Number']), str(lineas['Linea ID']), 'P')
            ierr_Q, rval_Q = psspy.brnmsc(int(lineas['From Bus Number']),int(lineas['To Bus Number']), str(lineas['Linea ID']), 'Q')
            #ierr_MVA, rval_MVA = psspy.brnmsc(int(lineas['From Bus Number']),int(lineas['To Bus Number']), str(lineas['Linea ID']), 'PCTCPA')
            #rval_MVA = ((math.sqrt(3)*132000*rval_AMPS)/1000000)/int(lineas['Rate1 (MVA)'])*100
            rate_MVA_Perc = 0
            if int(lineas['Rate1 (MVA)']) is not 0:
                rate_MVA_Perc = ((math.sqrt(3)*132000*rval_AMPS)/1000000)/int(lineas['Rate1 (MVA)'])*100
        #    print(ierr_AMPS,rval_AMPS)
        #    print(ierr_P, rval_P )
        #    print(ierr_Q, rval_Q )
            
        #    nombre_tabla = "OUTPUT_" + lineas['Subestacion 1'].strip(' ').replace(' ','_').replace('.','_') + "_" + lineas['Subestacion 2'].strip(' ').replace(' ','_').replace('.','_') + "_" + str(lineas['Linea ID'])
        #    instruccion_insert = "INSERT INTO " + nombre_tabla + " (ID_Escenario, AMPS, P, Q) VALUES (-1, " + str(rval_AMPS) + ", " + str(rval_P) + ", " + str(rval_Q) + ");"
        #    print(instruccion_insert)
            nombre_tabla = "OUTPUT_AMPS_P_Q_RATE_LAT_132"
    #        instruccion_insert = "INSERT INTO " + nombre_tabla + " (ID_Escenario, From_Bus_Number, To_Bus_Number, Linea_ID, AMPS, P_MW, Q_MVAR, RATE_PERC) VALUES (" + str(id_esc) + ", " + str(lineas['From Bus Number']) + ", " + str(lineas['To Bus Number']) + ", '" + str(lineas['Linea ID']) + "', " + str(rval_AMPS) + ", " + str(rval_P) + ", " + str(rval_Q) + ", " + str(rate_MVA_Perc) + ");"
            instruccion_insert = "INSERT INTO " + nombre_tabla + " (ID_Escenario, From_Bus_Number, Subestacion_1, To_Bus_Number, Subestacion_2, Linea_ID, AMPS, P_MW, Q_MVAR, RATE_PERC) VALUES (" + str(id_esc) + ", " + str(lineas['From Bus Number']) + ", '" + str(lineas['Subestacion 1'].strip(' ').replace(' ','_').replace('.','_')) + "', " + str(lineas['To Bus Number']) + ", '" + str(lineas['Subestacion 2'].strip(' ').replace(' ','_').replace('.','_')) + "', '" + str(lineas['Linea ID']) + "', " + str(rval_AMPS) + ", " + str(rval_P) + ", " + str(rval_Q) + ", " + str(rate_MVA_Perc) + ");"#str(rval_AMPS*100/lineas['Rate1 (MVA)']) + ");"
            
            cursor.execute(instruccion_insert)
            conn.commit()
        
        #Se recorre la lista de buses para guardar las tensiones en el SQL
        for nb in TENSION_BUSES_132:
        #    print(nb[0])
            nombre_tabla = "OUTPUT_KV_BUSES_132"
    #        instruccion_insert = "INSERT INTO " + nombre_tabla + " (ID_Escenario, Bus_Number, KV) VALUES (" + str(id_esc) + ", " + str(nb[0]) + ", " + str(nb[3]) + ");"
            instruccion_insert = "INSERT INTO " + nombre_tabla + " (ID_Escenario, Bus_Number, Subestacion, KV) VALUES (" + str(id_esc) + ", " + str(nb[0]) + ", '" + str(nb[1].strip(' ').replace(' ','_').replace('.','_')) + "', " + str(nb[3]) + ");"
        #    print(instruccion_insert)
            cursor.execute(instruccion_insert)
            conn.commit()
        

cursor.close()
del cursor
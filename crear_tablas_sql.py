# -*- coding: utf-8 -*-
"""
Created on Fri Oct 23 11:26:55 2020

@author: carrilesd
"""

##############################################################################
## Crear tablas SQL Simulaciones PSS octubre 2020                  ##
##############################################################################


#Definición de la conexión con la DB
import pyodbc 
import pandas as pd
from datetime import datetime,timedelta
import random


conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=193.144.190.81;'
                      'Database=Simulaciones_2020;'
                      'UID=user;'
                      'PWD=1234')
                      #'Trusted_Connection=yes;')
                      
#conn = pyodbc.connect('Driver={SQL Server};'
#                      'Server=193.144.195.16;'
#                      'Database=Simulaciones_2020;'
#                      'UID=Cuentapersonas;'
#                      'PWD=Cuentapersonas_pgob16')
#                      #'Trusted_Connection=yes;')

cursor = conn.cursor()


#ruta= r"F:\GTEA\PSS-Moises\PSSE_Viesgo"
ruta= r"C:\David\PSSE_Viesgo"
#CASOraw= ruta + r"RDF_Caso_Base_2019_Pcc_wind10_2.raw"
archivo_lineas = ruta + r"\datos_lineas132_zona_Viesgo.csv"
archivo_generadores = ruta + r"\datos_generadores132_zona_Viesgo.csv"



##############################################################################
## DATOS DE ENTRADA. ESCENARIOS DE SIMULACIÓN

## Lectura del csv con el listado de generadores de 132 kV y creación de tablas                  ##
##############################################################################
df_generadores132_viesgo = pd.read_csv(archivo_generadores, encoding='Latin9', header=0, sep=';', quotechar='\"')

### Se crea una tabla general con la información de los generadores y se rellena.
nombre_tabla_info = "INPUT_GENERADORES_132_Viesgo_INFO"
#instruccion_create = '''CREATE TABLE ''' + nombre_tabla_info + ''' ('''
#for i in range(0,len(df_generadores132_viesgo.columns)):
instruccion_create = '''CREATE TABLE [Simulaciones_2020].[dbo].[''' + nombre_tabla_info + '''] (Bus_Number INT, Bus_Name TEXT, Area_Num INT, ID_Generador TEXT, ierr_P INT, P_MW FLOAT, PMAX_MW FLOAT, PMIN_MW FLOAT, Q_MVAR FLOAT, QMAX_MVAR FLOAT, QMIN_MVAR FLOAT, MBASE_MVA FLOAT);'''

cursor.execute(instruccion_create)

for idx,generadores in df_generadores132_viesgo.iterrows():
    instruccion_insert = "INSERT INTO " + nombre_tabla_info + " (Bus_Number, Bus_Name, Area_Num, ID_Generador, ierr_P, P_MW, PMAX_MW, PMIN_MW, Q_MVAR, QMAX_MVAR, QMIN_MVAR, MBASE_MVA) VALUES (" + str(generadores['Bus_Number']) + ", '" + str(generadores['Bus_Name'].strip(' ').replace(' ','_').replace('.','_')) + "', " + str(generadores['Area_Num']) + ", '" + str(generadores['ID_Generador']) + "', " + str(generadores['ierr_P']) + ", " + str(generadores['P_MW']).replace(',','.') + ", " + str(generadores['PMAX_MW']).replace(',','.') + ", " + str(generadores['PMIN_MW']).replace(',','.') + ", " + str(generadores['Q_MVAR']).replace(',','.') + ", " + str(generadores['QMAX_MVAR']).replace(',','.') + ", " + str(generadores['QMIN_MVAR']).replace(',','.') + ", " + str(generadores['MBASE_MVA']).replace(',','.') + ");"
    cursor.execute(instruccion_insert)
    conn.commit()
#    print(instruccion_insert)
    
    


##############################################################################
## Tabla de escenarios de simulación.                  ##
##############################################################################   
    
nombre_tabla = "INPUT_ESCENARIOS_GENERADORES"
instruccion_create = '''CREATE TABLE [Simulaciones_2020].[dbo].[''' + nombre_tabla + '''] (ID_Escenario INT, Bus_Number INT, Bus_Name TEXT, ID_Generador TEXT, Porc_simulacion INT, P_MW_simulacion FLOAT);'''
cursor.execute(instruccion_create)
conn.commit()
#print(instruccion_create)

#Se rellena la tabla con 3 escenarios aleatorios
## ESCENARIO 1: Todos los generadores al 90% de PGen
#factor_escalado = 0.9
#for idx,generadores in df_generadores132_viesgo.iterrows():
##    print(str(generadores['P_MW']).replace(',','.'), str(float(generadores['P_MW'].replace(',','.'))*factor_escalado).replace(',','.'))
#    instruccion_insert = "INSERT INTO " + nombre_tabla + " (ID_Escenario, Bus_Number, Bus_Name, ID_Generador, P_MW_simulacion) VALUES (1, " + str(generadores['Bus_Number']) + ", '" + str(generadores['Bus_Name'].strip(' ').replace(' ','_').replace('.','_')) + "', '" + str(generadores['ID_Generador']) + "', " + str(float(generadores['P_MW'].replace(',','.'))*factor_escalado).replace(',','.') + ");"
##    print(instruccion_insert)
#    cursor.execute(instruccion_insert)
#    conn.commit()
#
### ESCENARIO 2: Todos los generadores al 110% de PGen
#factor_escalado = 1.1
#for idx,generadores in df_generadores132_viesgo.iterrows():
#    instruccion_insert = "INSERT INTO " + nombre_tabla + " (ID_Escenario, Bus_Number, BUS_Name, ID_Generador, P_MW_simulacion) VALUES (2, " + str(generadores['Bus_Number']) + ", '" + str(generadores['Bus_Name'].strip(' ').replace(' ','_').replace('.','_')) + "', '" + str(generadores['ID_Generador']) + "', " + str(float(generadores['P_MW'].replace(',','.'))*factor_escalado).replace(',','.') + ");"
#    cursor.execute(instruccion_insert)
#    conn.commit()
#
#
### ESCENARIO 3: Todos los generadores al 130% de PGen
#factor_escalado = 1.3
#for idx,generadores in df_generadores132_viesgo.iterrows():
#    instruccion_insert = "INSERT INTO " + nombre_tabla + " (ID_Escenario, Bus_Number, BUS_Name, ID_Generador, P_MW_simulacion) VALUES (3, " + str(generadores['Bus_Number']) + ", '" + str(generadores['Bus_Name'].strip(' ').replace(' ','_').replace('.','_')) + "', '" + str(generadores['ID_Generador']) + "', " + str(float(generadores['P_MW'].replace(',','.'))*factor_escalado).replace(',','.') + ");"
#    cursor.execute(instruccion_insert)
#    conn.commit()
    
    
    
    
##############################################################################  
##Generación de escenarios mediante simulaciones de Montecarlo
##############################################################################  
# Para cada generador PC se genera un valor aleatorio entre 0 y 100 que representa
# el porcentage de generación.
cursor = conn.cursor()
nombre_tabla = "INPUT_ESCENARIOS_GENERADORES"
num_escenarios = 3
for id_esc in range(0,num_escenarios):
    for idx, generadores in df_generadores132_viesgo.iterrows():
#    for generadores in df_generadores132_viesgo.itertuples():
        #Se calcula el cambio de potencia unicamente para los generadores PC
        if generadores['ID_Generador'] == 'PC':
            factor_escalado = random.randrange(0,100,1)
            pot_gen_esc = float(generadores['PMAX_MW'].replace(',','.'))*factor_escalado/100
#            print(factor_escalado, generadores['PMAX_MW'], pot_gen_esc)
            instruccion_insert = "INSERT INTO " + nombre_tabla + " (ID_Escenario, Bus_Number, Bus_Name, ID_Generador, Porc_simulacion, P_MW_simulacion) VALUES (" + str(id_esc) + ", " + str(generadores['Bus_Number']) + ", '" + str(generadores['Bus_Name'].strip(' ').replace(' ','_').replace('.','_')) + "', '" + str(generadores['ID_Generador']) + "', " + str(factor_escalado) + ", " + str(pot_gen_esc).replace(',','.') + ");"
#            print(instruccion_insert)
            cursor.execute(instruccion_insert)
            conn.commit()
#cursor.close()
#del cursor  
        




##############################################################################
## Lectura del csv con el listado de líneas de 132 kV y creación de tablas de salida                  ##
##############################################################################

df_lineas132_viesgo = pd.read_csv(archivo_lineas, encoding='Latin9', header=0, sep=';', quotechar='\"')

### Se crea una tabla general con la información de las líneas y se rellena después.
nombre_tabla_info = "OUTPUT_LAT_132_Viesgo_INFO"
instruccion_create = '''CREATE TABLE ''' + nombre_tabla_info + ''' (From_Bus_Number INT, Subestacion_1 TEXT, Area_Num_1 INT, To_Bus_Number INT, Subestacion_2 TEXT, Area_Num_2 INT, Linea_ID TEXT, Length_km FLOAT, RATE_1_MVA FLOAT);'''
#    print(instruccion_create)
cursor.execute(instruccion_create)


for idx,lineas in df_lineas132_viesgo.iterrows():
    # Guardar información de la línea en la tabla general
    instruccion_insert = "INSERT INTO " + nombre_tabla_info + " (From_Bus_Number, Subestacion_1, Area_Num_1, To_Bus_Number, Subestacion_2, Area_Num_2, Linea_ID, Length_km, RATE_1_MVA) VALUES (" + str(lineas['From Bus Number']) + ", '" + str(lineas['Subestacion 1'].strip(' ').replace(' ','_').replace('.','_')) + "', " + str(lineas['Area Num 1']) + ", " + str(lineas['To Bus Number']) + ", '" + str(lineas['Subestacion 2'].strip(' ').replace(' ','_').replace('.','_')) + "', " + str(lineas['Area Num 2']) + ", '" + str(lineas['Linea ID']) + "', " + str(lineas['Length (km)']).replace(',','.') + "," + str(lineas['Rate1 (MVA)']).replace(',','.') + ");"
#    print(instruccion_insert)
    cursor.execute(instruccion_insert)
    conn.commit()
   
#cursor.close()
#del cursor   
    

print('Estas dos instrucciones CREATE TABLE es necesario ejecutarlas directamente desde el SQL, desde Python se queda bloqueado el SQL del servidor y hay que reiniciar el equipo \n')
#cursor = conn.cursor()
#Crear la tabla de resultados de corriente y potencias
nombre_tabla_info = "OUTPUT_AMPS_P_Q_RATE_LAT_132"
instruccion_create = '''CREATE TABLE [Simulaciones_2020].[dbo].[''' + nombre_tabla_info + '''] (ID_Escenario INT, From_Bus_Number INT, Subestacion_1 TEXT, To_Bus_Number INT, Subestacion_2 TEXT, Linea_ID TEXT, AMPS FLOAT, P_MW FLOAT, Q_MVAR FLOAT, RATE_PERC FLOAT);'''
#cursor.execute(instruccion_create)
print(instruccion_create)


#Crear la tabla de resultados de tensión en los buses
nombre_tabla_info = "OUTPUT_KV_BUSES_132"
instruccion_create = '''CREATE TABLE [Simulaciones_2020].[dbo].[''' + nombre_tabla_info + '''] (ID_Escenario INT, Bus_Number INT, Subestacion TEXT, KV FLOAT);'''
#cursor.execute(instruccion_create)
print(instruccion_create)


   













###### OLD ######


##############################################################################
## OPCIÓN 2 SQL. 
##Lectura del csv con el listado de líneas de 132 kV y creación de UNA TABLA POR LÍNEA                  ##
##############################################################################

#df_lineas132_viesgo = pd.read_csv(archivo_lineas, encoding='Latin9', header=0, sep=';', quotechar='\"')
#
#### Se crea una tabla general con la información de las líneas y se rellena después.
#nombre_tabla_info = "OUTPUT_LAT_132_Viesgo_INFO"
#instruccion_create = '''CREATE TABLE ''' + nombre_tabla_info + ''' (From_Bus_Number INT, Subestacion_1 TEXT, Area_Num_1 INT, To_Bus_Number INT, Subestacion_2 TEXT, Area_Num_2 INT, Linea_ID TEXT, Length_km FLOAT);'''
##    print(instruccion_create)
#cursor.execute(instruccion_create)
#
#
#
### Creación de las tablas de resultados. Una tabla para cada línea con el formato:
### OUTPUT_Origen_Fin_IDLinea
#for idx,lineas in df_lineas132_viesgo.iterrows():
#    nombre_tabla = "OUTPUT_" + lineas['Subestacion 1'].strip(' ').replace(' ','_').replace('.','_') + "_" + lineas['Subestacion 2'].strip(' ').replace(' ','_').replace('.','_') + "_" + str(lineas['Linea ID'])
##    print(nombre_tabla)
##    print(lineas['Subestacion 1'])
##    print(lineas['Subestacion 2'])
##    print(lineas['Linea ID'])
#    instruccion_create = '''CREATE TABLE ''' + nombre_tabla + ''' (ID_Escenario int, AMPS float, P float, Q float);'''
##    print(instruccion_create)
#    cursor.execute(instruccion_create)
#    
#    # Guardar información de la línea en la tabla general
#    instruccion_insert = "INSERT INTO " + nombre_tabla_info + " (From_Bus_Number, Subestacion_1, Area_Num_1, To_Bus_Number, Subestacion_2, Area_Num_2, Linea_ID, Length_km) VALUES (" + str(lineas['From Bus Number']) + ", '" + str(lineas['Subestacion 1'].strip(' ').replace(' ','_').replace('.','_')) + "', " + str(lineas['Area Num 1']) + ", " + str(lineas['To Bus Number']) + ", '" + str(lineas['Subestacion 2'].strip(' ').replace(' ','_').replace('.','_')) + "', " + str(lineas['Area Num 2']) + ", '" + str(lineas['Linea ID']) + "', " + str(lineas['Length (km)']).replace(',','.') + ");"
##    print(instruccion_insert)
#    cursor.execute(instruccion_insert)
##    cursor.execute('''
##                   CREATE TABLE''' + nombre_tabla + '''(
##                   AMPS float,
##                   P float,
##                   Q float,
##                   );
##                   ''')
#    conn.commit()
#cursor.close()
#del cursor   
    
    



    
    
###### CÓDIGO PARA BORRAR LAS TABLAS
#cursor = conn.cursor()
#for idx,lineas in df_lineas132_viesgo.iterrows():
#    nombre_tabla = "OUTPUT_" + lineas['Subestacion 1'].strip(' ').replace(' ','_').replace('.','_') + "_" + lineas['Subestacion 2'].strip(' ').replace(' ','_').replace('.','_') + "_" + str(lineas['Linea ID'])  
#    cursor.execute('DROP TABLE ' + nombre_tabla)
#    conn.commit()
#cursor.close()
#del cursor










#import mysql.connector
#
#mydb = mysql.connector.connect(
#  host="193.144.190.81",
#  user="use",
#  password="1234",
#  database="Simulaciones_2020"
#)
#
#mycursor = mydb.cursor()
#
#mycursor.execute("CREATE TABLE [Simulaciones_2020].[dbo].[probando2] (AMPS float, P float, Q float)")
##                 CREATE TABLE customers (name VARCHAR(255), address VARCHAR(255))")
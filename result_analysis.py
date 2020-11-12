# -*- coding: utf-8 -*-
"""
Created on Wed Nov  4 09:49:57 2020

@author: carrilesd
"""


##############################################################################
## ANÁLISIS DE RESULTADOS DE LA SIMULACIÓN. QUERYs SQL
##############################################################################

import pyodbc
import pandas as pd

#SELECT COUNT( *) as "Number of Rows"
#FROM [Simulaciones_2020].[dbo].[OUTPUT_AMPS_P_Q_RATE_LAT_132]
#where [From_Bus_Number] = 411 and [To_Bus_Number] = 412 and [Linea_ID] = '1' and [Indice_Carga] > 120;
#
#SELECT DISTINCT From_Bus_Number,To_Bus_Number, Linea_ID
#FROM [Simulaciones_2020].[dbo].[OUTPUT_AMPS_P_Q_RATE_LAT_132]
#order by From_Bus_Number asc;
#
#/*No se que hace esto:*/
#SELECT COUNT( *) as dt
#FROM (SELECT DISTINCT From_Bus_Number, To_Bus_Number
#FROM [Simulaciones_2020].[dbo].[OUTPUT_AMPS_P_Q_RATE_LAT_132] where Indice_Carga > 120) as dt1;

      
ruta= r"C:\David\PSSE_Viesgo"
archivo_config = ruta + r"\Config.txt"
f = open (archivo_config,'r')
ip_server = f.readline().splitlines()[0]
db_server = f.readline().splitlines()[0]
usr_server = f.readline().splitlines()[0]
pwd_server = f.readline().splitlines()[0]
f.close()


#Definición de la conexión con la DB
conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=' + ip_server + ';'
                      'Database=' + db_server + ';'
                      'UID=' + usr_server + ';'
                      'PWD=' + pwd_server)
                      #'Trusted_Connection=yes;')
cursor = conn.cursor()


nombre_tabla_info = "OUTPUT_AMPS_P_Q_RATE_LAT_132"
perc_sobrecarga = 125
#SQL_Query = pd.read_sql_query('''SELECT ALL [ID_Escenario],[Bus_Number],[Bus_Name],[ID_Generador],[Porc_simulacion],[P_MW_simulacion] FROM [Simulaciones_2020].[dbo].[''' + nombre_tabla_info + '''] where ID_Escenario = ''' + str(id_esc), conn)
SQL_Query = pd.read_sql_query('''SELECT ALL [ID_Escenario],[From_Bus_Number],[Subestacion_1],[To_Bus_Number],[Subestacion_2],[Linea_ID],[AMPS],[P_MW],[Q_MVAR],[Indice_Carga] FROM [Simulaciones_2020].[dbo].[''' + nombre_tabla_info + '''] where Indice_Carga > ''' + str(perc_sobrecarga), conn)

df_sobrecarga_lineas = pd.DataFrame(SQL_Query, columns=['ID_Escenario','From_Bus_Number','Subestacion_1','To_Bus_Number','Subestacion_2','Linea_ID','AMPS','P_MW','Q_MVAR','Indice_Carga'])
    
escenarios_sobrecarga_lineas = df_sobrecarga_lineas['ID_Escenario'].drop_duplicates(keep = 'first').reset_index(drop=True)

SQL_Query = pd.read_sql_query('''SELECT ALL [ID_Escenario] FROM [Simulaciones_2020].[dbo].[''' + nombre_tabla_info + ''']''', conn)
indice_escenarios =  pd.DataFrame(SQL_Query, columns=['ID_Escenario']).drop_duplicates(keep = 'first').sort_values('ID_Escenario').reset_index(drop=True)

cursor.close()
del cursor

#Escenarios donde no se supera el porcentaje de sobrecarga definido.
escenarios_sin_sobrecarga = []
escenarios_sin_sobrecarga = indice_escenarios['ID_Escenario'].append(escenarios_sobrecarga_lineas).drop_duplicates(keep = False).reset_index(drop=True)
#escenarios_sin_sobrecarga2 = escenarios_sin_sobrecarga.drop_duplicates(keep = False).reset_index(drop=True)
#escenarios_sin_sobrecarga2 = list(dict.fromkeys(escenarios_sin_sobrecarga)) #Eliminar duplicados


#Se considera una tolerancia de un número máximo de líneas sobrecargadas.
escenarios_tolerancia = []
tolerancia_lineas = 5
for row in escenarios_sobrecarga_lineas:
    temp = []
    temp = df_sobrecarga_lineas[df_sobrecarga_lineas['ID_Escenario'] == row]
    if len(temp) <= tolerancia_lineas:
        escenarios_tolerancia.append(int(row))

print('Escenarios sin sobrecarga: ' + str(len(escenarios_sin_sobrecarga)))
print(escenarios_sin_sobrecarga)
print('Escenarios con tolerancia de ' + str(tolerancia_lineas) + ' líneas: ' + str(len(escenarios_tolerancia)))
print(escenarios_tolerancia)






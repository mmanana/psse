##############################################################################
##############################################################################
#######                                                                #######
#######     SIMULACION DE ESCENARIOS SIN CONTINGENCIAS                 #######
#######                                                                #######
##############################################################################
##############################################################################

## FICHERO: escenarios_sin_con_v1.py
## AUTOR: GTEA-UC
## FECHA ULTIMA ACTUALIZACION: 26/9/2020 (Mario Manana)

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


#bus_i = 798
#bus_f = 734
#ierr, line_rx = psspy.brndt2( bus_i, bus_f, '1', 'RXZ')

##############################################################################
## Lee los buses del modelo
##############################################################################
ierr, busnumbers = psspy.abusint(sid=-1, string="NUMBER")
ierr, busnames = psspy.abuschar(sid=-1, string="NAME")
ierr, busvoltages = psspy.abusreal(sid=-1, string="BASE")
#ierr, bus_shunts = psspy.abuscplx(sid=-1, string="SHUNTACT")




##############################################################################
## Instrucciones para recorrer todos los buses y almacenar aquellos que
## verifiquen el criterio de tension base = 132 kV
##############################################################################
BUSLIST=[]
BUSNAME=[]
#Nbus = len( busnumbers[0])
for nb in busnumbers[0]:
    bus_index = busnumbers[0].index(nb)
    voltage = busvoltages[0][bus_index]
    name = busnames[0][bus_index]
    if  voltage == 132.0:
        print( busnames[0][bus_index])
        BUSLIST.append(nb)
        BUSNAME.append(psspy.notona(nb)[1])
    nb=nb+1
print '\n\n---------- TODOS LOS BUSES DE 132 kV ----------\n'
print BUSLIST
print '\n\n-----------------------------------------------\n'

##############################################################################
## Instrucciones para obtener las líneas conectadas a los buses de 132 kV
##############################################################################
p=0
RAMAS=[]
BUSZONA=BUSLIST
while p<len(BUSZONA):
    j=p
    while j<len(BUSZONA):
        if 0 == psspy.brnstt(BUSZONA[p],BUSZONA[j],'1')[0]:
            RAMAS.append([BUSZONA[p],BUSZONA[j],'1'])
        elif 0 == psspy.brnstt(BUSZONA[p],BUSZONA[j],'2')[0]:
            RAMAS.append([BUSZONA[p],BUSZONA[j],'2'])
        j=j+1
    p=p+1
print '\n\n---------- LINEAS DE 132 kV  ----------\n'
print RAMAS
print '\n\n---------------------------------------\n'



##############################################################################
## Instrucciones para obtener los valores de las líneas de 132 kV
##############################################################################
total_kmlinea132 = 0
kmlinea132_sindatos = 0
Nlinea132_sindatos = 0
Ntotal_linea132 = 0
kmlinea132_sindatosZ = 0
Nlinea132_sindatosZ = 0


# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook(ruta + r"\lineas_1.xlsx")
worksheet = workbook.add_worksheet()

# Start from the first cell. Rows and columns are zero indexed.
row = 0
col = 0

worksheet.write(row, col, 'From Bus Number')
col += 1
worksheet.write(row, col, 'Subestacion 1')
col += 1

worksheet.write(row, col, 'Area Num 1')
col += 1

worksheet.write(row, col, 'To Bus Number')
col += 1
worksheet.write(row, col, 'Subestacion 2')
col += 1

worksheet.write(row, col, 'Area Num 2')
col += 1

worksheet.write(row, col, 'Linea ID')
col += 1

worksheet.write(row, col, 'Length (km)')
col += 1
worksheet.write(row, col, 'R(pu)')
col += 1
worksheet.write(row, col, 'X(pu)')
col += 1
worksheet.write(row, col, 'B(pu)')
col += 1
worksheet.write(row, col, 'RZ(pu)')
col += 1
worksheet.write(row, col, 'XZ(pu)')


for linea in RAMAS:
    row += 1
    col = 0
    bus_i = linea[0]
    bus_index = busnumbers[0].index(bus_i)
    voltage_i = busvoltages[0][bus_index]
    name_i = busnames[0][bus_index]
    
    bus_i_area_num = psspy.busint(bus_i, 'AREA')
    
    bus_f = linea[1]
    ckt = linea[2]
    bus_index = busnumbers[0].index(bus_f)
    voltage_f = busvoltages[0][bus_index]
    name_f = busnames[0][bus_index]
    
    bus_f_area_num = psspy.busint(bus_f, 'AREA')
    
    R_line = 0
    X_line = 0
    RZ_line = 0
    XZ_line = 0
    line_length = 0.0
    ierr, line_length = psspy.brndat( bus_i, bus_f, ckt, 'LENGTH')
    ierr, line_B = psspy.brndat( bus_i, bus_f, ckt, 'CHARG')
    ierr, line_rx = psspy.brndt2( bus_i, bus_f, ckt, 'RX')
    if ierr == 0:
        R_line = line_rx.real
        X_line = line_rx.imag
    ierr, line_rxz = psspy.brndt2( bus_i, bus_f, ckt, 'RXZ')
    if ierr == 0:
        RZ_line = line_rxz.real
        XZ_line = line_rxz.imag
     # print( str(bus_i) + ' ' + name_i + ' ; ' + str(bus_f) + ' ' + name_f + ' ' +  str(line_rx) + ' ' + str(line_length))
    worksheet.write(row, col, bus_i)
    col += 1
    worksheet.write(row, col, name_i)
    col += 1
    
    worksheet.write(row, col, bus_i_area_num[1])
    col += 1
    
    worksheet.write(row, col, bus_f)
    col += 1
    worksheet.write(row, col, name_f)
    col += 1
    
    worksheet.write(row, col, bus_f_area_num[1])
    col += 1
    
    worksheet.write(row, col, linea[2]) #ID de la línea
    col += 1
    
    worksheet.write(row, col, round(line_length,2))
    col += 1
    worksheet.write(row, col, R_line)
    col += 1
    worksheet.write(row, col, X_line)
    col += 1
    worksheet.write(row, col, line_B)
    total_kmlinea132 += line_length
    Ntotal_linea132 += 1
    if R_line == 0:
        kmlinea132_sindatos += line_length
        Nlinea132_sindatos += 1
    if RZ_line == 0:
        kmlinea132_sindatosZ += line_length
        Nlinea132_sindatosZ += 1


workbook.close()

print('Total km líneas de 132 kV: ' + str(total_kmlinea132))
print('Número total de líneas de 132 kV: ' + str(Ntotal_linea132))
print('Total km líneas de 132 kV sin datos: ' + str(kmlinea132_sindatos))
print('Número total de líneas de 132 kV sin datos: ' + str(Nlinea132_sindatos))
print('Número total de líneas de 132 kV sin datos de secuencia cero: ' + str(Nlinea132_sindatosZ))


#############################################################################
## Lectura de datos disponibles sobre las líneas de 132 kV
#############################################################################
datos_lineas= ruta + r"\PSSE_Listado_LATs_132kV.xlsx"

### Posible error. Instalar paquete xlrd: conda install -c anaconda xlrd (habiendo activado el entorno python2)

df_lineas132 = pd.read_excel(datos_lineas, sheet_name='Hoja1' )
Num_lineas = len(df_lineas132)
Num_NaN = df_lineas132['R-Zero [pu]'].isna().sum()
print('Datos disponibles de secuencia cero: ' + str(Num_lineas-Num_NaN))



#############################################################################
## Lectura de datos del archivo generado con líneas de 132 kV.
## Comparativa de datos disponibles para incluir las mediciones en campo
#############################################################################
datos_lineas_modelo = ruta + r"\lineas_1.xlsx"
df_lineas132_modelo = pd.read_excel(datos_lineas_modelo) #, sheet_name='Sheet1' )
#Se añaden los valores leídos del listado proporcionado en columnas nuevas
df_lineas132_modelo["R[pu]_med"] = ""
df_lineas132_modelo["Delta_R[pu]_med"] = ""
df_lineas132_modelo["X[pu]_med"] = ""
df_lineas132_modelo["Delta_X[pu]_med"] = ""
df_lineas132_modelo["B[pu]_med"] = ""
df_lineas132_modelo["Delta_B[pu]_med"] = ""
df_lineas132_modelo["RZ[pu]_med"] = ""
# df_lineas132_modelo["Delta_RZ[pu]_med"] = ""
df_lineas132_modelo["XZ[pu]_med"] = ""
# df_lineas132_modelo["Delta_XZ[pu]_med"] = ""
df_lineas132_modelo["BZ[pu]_med"] = ""
# df_lineas132_modelo["Delta_BZ[pu]_med"] = ""


for idx, row in df_lineas132_modelo.iterrows():
    #print(row["R(pu)"])
    #row["R[pu]_med"] = 
    filtro_from_to = df_lineas132.loc[(df_lineas132['From Bus  Number'] == row['From Bus Number']) & (df_lineas132['To Bus  Number'] == row['To Bus Number'])].reset_index(drop=True)
    # print(len(filtro_from_to))
    
    if len(filtro_from_to) > 0:
        # print(len(filtro_from_to))
        # print(filtro_from_to['R [pu]'][0])
        # print(filtro_from_to['X [pu]'][0])
        # print(filtro_from_to['B(pu).1'][0])
        # print(filtro_from_to['R-Zero [pu]'][0])
        # print(filtro_from_to['X-Zero [pu]'][0])
        # print(filtro_from_to['B-Zero(pu)'][0])

    
        df_lineas132_modelo.loc[idx, 'R[pu]_med'] = str(filtro_from_to['R [pu]'][0]).replace('.',',')
        df_lineas132_modelo.loc[idx, 'X[pu]_med'] = str(filtro_from_to['X [pu]'][0]).replace('.',',')
        df_lineas132_modelo.loc[idx, 'B[pu]_med'] = str(filtro_from_to['B(pu).1'][0]).replace('.',',')
        df_lineas132_modelo.loc[idx, 'RZ[pu]_med'] = str(filtro_from_to['R-Zero [pu]'][0]).replace('.',',')
        df_lineas132_modelo.loc[idx, 'XZ[pu]_med'] = str(filtro_from_to['X-Zero [pu]'][0]).replace('.',',')
        df_lineas132_modelo.loc[idx, 'BZ[pu]_med'] = str(filtro_from_to['B-Zero(pu)'][0]).replace('.',',')
        #Se calcula y guarda el delta de cada variable
        if row['R(pu)'] > 0:
            df_lineas132_modelo.loc[idx, 'Delta_R[pu]_med'] = str(filtro_from_to['R [pu]'][0] / row['R(pu)']).replace('.',',')
        if row['X(pu)'] > 0:
            df_lineas132_modelo.loc[idx, 'Delta_X[pu]_med'] = str(filtro_from_to['X [pu]'][0] / row['X(pu)']).replace('.',',')
        if row['B(pu)'] > 0:    
            df_lineas132_modelo.loc[idx, 'Delta_B[pu]_med'] = str(filtro_from_to['B(pu).1'][0] / row['B(pu)']).replace('.',',')
        # if row['RZ(pu)'] > 0: 
            # df_lineas132_modelo.loc[idx, 'RZ[pu]_med'] = str(filtro_from_to['R-Zero [pu]'][0] / row['RZ(pu)']).replace('.',',')
        # if row['XZ(pu)'] > 0: 
            # df_lineas132_modelo.loc[idx, 'XZ[pu]_med'] = str(filtro_from_to['X-Zero [pu]'][0] / row['XZ(pu)']).replace('.',',')
        # if row['XB(pu)'] > 0: 
            
        # print(df_lineas132_modelo.loc[idx])
        

#Se guarda el resultado en un csv
datos_lineas_mod = ruta + r"\datos_lineas132_mod.csv"        
df_lineas132_modelo.to_csv(datos_lineas_mod,  decimal=',', sep=';', encoding='latin9', index=False, header=True)



#############################################################################
## Filtrado de las líneas de 132 kV para dejar únicamente las de zona Viesgo
## Se dejan también las líneas que empiezan o acaban en otra zona.
#############################################################################
df_lineas132_viesgo = []
df_lineas132_viesgo = df_lineas132_modelo[(df_lineas132_modelo['Area Num 1'] == 10) | (df_lineas132_modelo['Area Num 2'] == 10)]
datos_lineas_mod = ruta + r"\datos_lineas132_zona_Viesgo.csv" 
df_lineas132_viesgo.to_csv(datos_lineas_mod,  decimal=',', sep=';', encoding='latin9', index=False, header=True)





#############################################################################
## Recorrer los buses para identificar los que tienen conectado un generador, 
## son de 132 kV y son de la zona Viesgo.
#############################################################################
BUSESGEN_132=[]
BUSESGEN_NAME_132=[]

#bus_i_area_num = psspy.busint(bus_i, 'AREA')

for nb in busnumbers[0]:
    bus_index = busnumbers[0].index(nb)
    voltage = busvoltages[0][bus_index]
    bus_i_area_num = psspy.busint(nb, 'AREA')
    ierr, cmpval = psspy.gendat(nb)
    
    #if (ierr == 0 and voltage == 132.0):# or ierr == 4):
    if (ierr == 0 or ierr == 4) and voltage == 132.0 and bus_i_area_num[1] == 10:
#        print(bus_i_area_num[1],nb,  bus_index, voltage)
#         print( busnames[0][bus_index])
        BUSESGEN_132.append(nb)
        BUSESGEN_NAME_132.append([nb, psspy.notona(nb)[1], bus_i_area_num[1], voltage, cmpval])
        
#Columnas de la lista BUSESGEN_NAME_132:
#['Bus_Number','BUS_NAME', 'Base_kV', 'Area_Num', 'Complex_generation']



#############################################################################
## Recorrer los buses de 132 kV con generador conectado y obtener el número de 
## generadores y el id de cada uno.
#############################################################################
# Diferentes IDs de generadores localizados en zona Viesgo (10): 
#1, 2, 3, 4, 5, 6, 7
#A, B, C, D, E
#PC, EO

GENERADORES_132=[]
lista_ids = ['1', '2', '3', '4', '5', '6', '7', 'A', 'B', 'C', 'D', 'E', 'PC', 'EO']
i=0
for nb in BUSESGEN_132:
    for id_gen in lista_ids:
        ierr_P, rval_P = psspy.macdat(nb, id_gen, 'P')
        ierr_PMAX, rval_PMAX = psspy.macdat(nb, id_gen, 'PMAX')
        ierr_PMIN, rval_PMIN = psspy.macdat(nb, id_gen, 'PMIN')
        ierr_Q, rval_Q = psspy.macdat(nb, id_gen, 'Q')
        ierr_QMAX, rval_QMAX = psspy.macdat(nb, id_gen, 'QMAX')
        ierr_QMIN, rval_QMIN = psspy.macdat(nb, id_gen, 'QMIN')
        ierr_MBASE, rval_MBASE = psspy.macdat(nb, id_gen, 'MBASE')
        bus_i_area_num = psspy.busint(nb, 'AREA')
        if (ierr_P == 0 or ierr_P == 4):
            GENERADORES_132.append([nb, psspy.notona(nb)[1], bus_i_area_num[1], id_gen, ierr_P, rval_P, rval_PMAX, rval_PMIN, rval_Q, rval_QMAX, rval_QMIN, rval_MBASE])
#    print(i, nb, ierr, rval)
#Columnas de la lista GENERADORES_132: 
#['Bus_Number', 'BUS_NAME', 'ID_generator', 'ierr', 'Machine_loading_MW']
#Se convierte la lista GENERADORES_132 a DataFrame para guardarlo como csv con columnas
GENERADORES_132_DF = DataFrame(GENERADORES_132, columns=['Bus_Number', 'BUS_Name', 'Area_Num', 'ID_Generador', 'ierr_P', 'P_MW', 'PMAX_MW', 'PMIN_MW', 'Q_MVAR', 'QMAX_MVAR', 'QMIN_MVAR', 'MBASE_MVA'])
GENERADORES_132_DF.to_csv(ruta + r"\datos_generadores132_zona_Viesgo.csv",  decimal=',', sep=';', encoding='latin9', index=False, header=True)
    

            
            
            
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
#    print(ierr_AMPS,rval_AMPS)
#    print(ierr_P, rval_P )
#    print(ierr_Q, rval_Q )
    
#    nombre_tabla = "OUTPUT_" + lineas['Subestacion 1'].strip(' ').replace(' ','_').replace('.','_') + "_" + lineas['Subestacion 2'].strip(' ').replace(' ','_').replace('.','_') + "_" + str(lineas['Linea ID'])
#    instruccion_insert = "INSERT INTO " + nombre_tabla + " (ID_Escenario, AMPS, P, Q) VALUES (-1, " + str(rval_AMPS) + ", " + str(rval_P) + ", " + str(rval_Q) + ");"
#    print(instruccion_insert)
    nombre_tabla = "OUTPUT_AMPS_P_Q_LAT_132"
    instruccion_insert = "INSERT INTO " + nombre_tabla + " (ID_Escenario, From_Bus_Number, To_Bus_Number, AMPS, P_MW, Q_MVAR) VALUES (-1, " + str(lineas['From Bus Number']) + ", " + str(lineas['To Bus Number']) + ", " + str(rval_AMPS) + ", " + str(rval_P) + ", " + str(rval_Q) + ");"
    
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
    instruccion_insert = "INSERT INTO " + nombre_tabla + " (ID_Escenario, Bus_Number, KV) VALUES (-1, " + str(nb[0]) + ", " + str(nb[3]) + ");"
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
num_escenario = [1, 2, 3] #"2"
#id_esc = 3
for id_esc in num_escenario:
    SQL_Query = pd.read_sql_query('''SELECT ALL [ID_Escenario],[Bus_Number],[Bus_Name],[ID_Generador],[P_MW_simulacion] FROM [Simulaciones_2020].[dbo].[''' + nombre_tabla_info + '''] where ID_Escenario = ''' + str(id_esc), conn)
    df_escenario = pd.DataFrame(SQL_Query, columns=['ID_Escenario','Bus_Number','BUS_Name','ID_Generador','P_MW_simulacion'])
    
    #Cambio de potencia en los generadores
    for idx, generadores in df_escenario.iterrows():
        print(str(generadores['Bus_Number']), str(generadores['ID_Generador']),str(generadores['P_MW_simulacion']))
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
    #    print(ierr_AMPS,rval_AMPS)
    #    print(ierr_P, rval_P )
    #    print(ierr_Q, rval_Q )
        
    #    nombre_tabla = "OUTPUT_" + lineas['Subestacion 1'].strip(' ').replace(' ','_').replace('.','_') + "_" + lineas['Subestacion 2'].strip(' ').replace(' ','_').replace('.','_') + "_" + str(lineas['Linea ID'])
    #    instruccion_insert = "INSERT INTO " + nombre_tabla + " (ID_Escenario, AMPS, P, Q) VALUES (-1, " + str(rval_AMPS) + ", " + str(rval_P) + ", " + str(rval_Q) + ");"
    #    print(instruccion_insert)
        nombre_tabla = "OUTPUT_AMPS_P_Q_LAT_132"
        instruccion_insert = "INSERT INTO " + nombre_tabla + " (ID_Escenario, From_Bus_Number, To_Bus_Number, AMPS, P_MW, Q_MVAR) VALUES (" + str(id_esc) + ", " + str(lineas['From Bus Number']) + ", " + str(lineas['To Bus Number']) + ", " + str(rval_AMPS) + ", " + str(rval_P) + ", " + str(rval_Q) + ");"
        
        cursor.execute(instruccion_insert)
        conn.commit()
    
    #Se recorre la lista de buses para guardar las tensiones en el SQL
    for nb in TENSION_BUSES_132:
    #    print(nb[0])
        nombre_tabla = "OUTPUT_KV_BUSES_132"
        instruccion_insert = "INSERT INTO " + nombre_tabla + " (ID_Escenario, Bus_Number, KV) VALUES (" + str(id_esc) + ", " + str(nb[0]) + ", " + str(nb[3]) + ");"
    #    print(instruccion_insert)
        cursor.execute(instruccion_insert)
        conn.commit()
    

cursor.close()
del cursor



















#print('Info inicial. Tensión bus inicio y fin, corriente línea y potencia generador.')
#print(psspy.busdat(109, 'KV'), 'ierr, kV inicio')
#print(psspy.busdat(941, 'KV'), 'ierr, kV fin')
#print(psspy.brnmsc(109, 941, '1', 'AMPS'), 'ierr, A linea')
#print(psspy.macdat(109, '1', 'P'), 'ierr, MW')
#
#print('Info tras el cambio. Tensión bus inicio y fin, corriente línea y potencia generador.')
#print(psspy.busdat(109, 'KV'), 'ierr, kV inicio')
#print(psspy.busdat(941, 'KV'), 'ierr, kV fin')
#print(psspy.brnmsc(109, 941, '1', 'AMPS'), 'ierr, A linea')
#print(psspy.macdat(109, '1', 'P'), 'ierr, MW')
#
#
#psspy.machine_data_2(109,'1',[psspy.getdefaultint()],[142.25999450683594])
#psspy.fdns([0,0,0,1,1,1,99,0])
#U=psspy.solv()
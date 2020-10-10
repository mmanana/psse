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
import xlsxwriter


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
ruta= r"C:\mario\trabajos2\viesgo_applus_escenarios_red\simulacion\psse"
#ruta= r"C:\David\PSSE_Viesgo"
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
## Instrucciones para obtener los valore de las líneas de 132 kV
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
worksheet.write(row, col, 'To Bus Number')
col += 1
worksheet.write(row, col, 'Subestacion 2')
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
    bus_f = linea[1]
    ckt = linea[2]
    bus_index = busnumbers[0].index(bus_f)
    voltage_f = busvoltages[0][bus_index]
    name_f = busnames[0][bus_index]
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
    worksheet.write(row, col, bus_f)
    col += 1
    worksheet.write(row, col, name_f)
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

print('********************************************************************')
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
## Comparativa de datos disponibles
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
## Recorrido de los buses para identificar los que tienen conectado un generador.
#############################################################################
BUSESGEN=[]
BUSESGEN_NAME=[]

for nb in busnumbers[0]:
    ierr, cmpval = psspy.gendat(nb)
    if (ierr == 0):# or ierr == 4):
        # print( busnames[0][bus_index])
        BUSESGEN.append(nb)
        BUSESGEN_NAME.append([nb, psspy.notona(nb)[1], cmpval])

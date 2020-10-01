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
ruta= r"""C:\mario\trabajos2\viesgo_applus_escenarios_red\simulacion\escenarios_sin_con\ """
CASOraw= ruta + r"RDF_Caso_Base_2019_Pcc_wind10_2.raw"
CASOsav= ruta + r"RDF_Caso_Base_2019_Pcc_wind10_2.sav"
SALIDALINEA= ruta + r"Salida.txt"
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
workbook = xlsxwriter.Workbook('lines_1.xlsx')
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

print('Total km líneas de 132 kV: ' + str(total_kmlinea132))
print('Número total de líneas de 132 kV: ' + str(Ntotal_linea132))
print('Total km líneas de 132 kV sin datos: ' + str(kmlinea132_sindatos))
print('Número total de líneas de 132 kV sin datos: ' + str(Nlinea132_sindatos))
print('Número total de líneas de 132 kV sin datos de secuencia cero: ' + str(Nlinea132_sindatosZ))


#############################################################################
## Lectura de datos disponibles sobre las líneas de 132 kV
#############################################################################
ruta_datos= r"""C:\mario\trabajos2\viesgo_applus_escenarios_red\simulacion"""
datos_lineas= ruta_datos + r"\PSSE_Listado_LATs_132kV.xlsx"

df_lineas132 = pd.read_excel(datos_lineas, sheet_name='Hoja1' )
Num_lineas = len(df_lineas132)
Num_NaN = df_lineas132['R-Zero [pu]'].isna().sum()
print('Datos disponibles de secuencia cero: ' + str(Num_lineas-Num_NaN))



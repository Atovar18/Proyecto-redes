import pandas as pd
import numpy as np
import math, sys, openpyxl
from decimal import Decimal 
datos = open("C:/Users/Usuario/Desktop/Hola mundo/Proyecto-redes/data_io_copia.xlsx")

#almcenamos las paginas de la tabla de excel
F_and = pd.read_excel ("C:/Users/Usuario/Desktop/Hola mundo/Proyecto-redes/data_io_copia.xlsx" , "f_and_ouput", header= None)
V_fuente = pd.read_excel ("C:/Users/Usuario/Desktop/Hola mundo/Proyecto-redes/data_io_copia.xlsx", "V_fuente")
I_fuente = pd.read_excel ("C:/Users/Usuario/Desktop/Hola mundo/Proyecto-redes/data_io_copia.xlsx", "V_fuente")
Z = pd.read_excel("C:/Users/Usuario/Desktop/Hola mundo/Proyecto-redes/data_io_copia.xlsx", "Z")
VTH_AND_ZTH = pd.read_excel("C:/Users/Usuario/Desktop/Hola mundo/Proyecto-redes/data_io_copia.xlsx", "VTH_AND_ZTH")
Sfuente = pd.read_excel("C:/Users/Usuario/Desktop/Hola mundo/Proyecto-redes/data_io_copia.xlsx", "Sfuente")
S_Z = pd.read_excel("C:/Users/Usuario/Desktop/Hola mundo/Proyecto-redes/data_io_copia.xlsx", "S_Z")
Balance_S = pd.read_excel("C:/Users/Usuario/Desktop/Hola mundo/Proyecto-redes/data_io_copia.xlsx", "V_fuente")

V_nominal = np.array(V_fuente.iloc[:, 2])

#calculamos el valor de w para poder hacer el wt y obtener el angulo de desfase
if F_and[1][0] == 60: #si la frecuncia es 60Hz, el valor de w es directamente 60
    w = 377

else: #si la frecuencia no es 60, entonces se calcula manualmente el valor de w y se redondea a su entero mas cercano
    w = round(2 * math.pi * F_and[1][0], 4)


capacitores = np.array(V_fuente.iloc[:,25])                 #Escogemos la columna de los capacitores del archivo.

Zcap = (-1/(w*(capacitores*(10**-6))))                      #Calculamos las imperancias de los capacitores.

inductores=np.array (V_fuente.iloc[:, 5])                   #Escogemos la columna de los inductores del archivo.

Zinduc = (w*(inductores*(10**-3)))                          #Calculamos las imperancias de los inductores.

resistencias=np.array (V_fuente.iloc[:, 4])                 #Escogemos la columna de las resistencias.

desfa=np.array (V_fuente.iloc[:, 3])                        #Escogemos la columna para del tiempo de desfasaje.

ang = (desfa * w)                                           #Calculamos el wt que usaremos para el angulo de desfasaje.


Vrms = np.array (V_nominal / np.sqrt(2))                    #Calculo del Vrms pero con dudas sobre todos los decimales.
for i in range (len(Vrms)):
    Vrms [i] = round (Vrms [i], 4)                          #Aprroximamos Vrms a 4 decimales.    

V_fasorial = Vrms * (np.cos(ang)) + np.complex_(Vrms * np.sin(ang) * 1j) 
V_fasorial = np.round(V_fasorial, 4)                        #Aproximamos el V en su forma rentangular a 4 decimales.

print (Vrms)
print ()
print (ang)
print ()
print (V_fasorial)
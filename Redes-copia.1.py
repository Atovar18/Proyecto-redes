import pandas as pd
import numpy as np
import math, sys, openpyxl
from decimal import Decimal 
from itertools import combinations
#Listas para usadas - pagina2.
ListZcap = []
ListZinduc = []
ListResistencias = []


#Listas usadas - pagina3.

ListResistencias_I = []
ListZcap_I = []
ListZinduc_I = []
SeriesFi = []

datos = open("C:/Users/Usuario/Desktop/Hola mundo/Proyecto-redes/data_io_copia.xlsx")

#almcenamos las paginas de la tabla de excel
F_and = pd.read_excel ("C:/Users/Usuario/Desktop/Hola mundo/Proyecto-redes/data_io_copia.xlsx" , "f_and_ouput", header= None)
V_fuente = pd.read_excel ("C:/Users/Usuario/Desktop/Hola mundo/Proyecto-redes/data_io_copia.xlsx", "V_fuente")
I_fuente = pd.read_excel ("C:/Users/Usuario/Desktop/Hola mundo/Proyecto-redes/data_io_copia.xlsx", "I_fuente")
Z = pd.read_excel("C:/Users/Usuario/Desktop/Hola mundo/Proyecto-redes/data_io_copia.xlsx", "Z")
VTH_AND_ZTH = pd.read_excel("C:/Users/Usuario/Desktop/Hola mundo/Proyecto-redes/data_io_copia.xlsx", "VTH_AND_ZTH")
Sfuente = pd.read_excel("C:/Users/Usuario/Desktop/Hola mundo/Proyecto-redes/data_io_copia.xlsx", "Sfuente")
S_Z = pd.read_excel("C:/Users/Usuario/Desktop/Hola mundo/Proyecto-redes/data_io_copia.xlsx", "S_Z")
Balance_S = pd.read_excel("C:/Users/Usuario/Desktop/Hola mundo/Proyecto-redes/data_io_copia.xlsx", "V_fuente")

V_nominal = np.array(V_fuente.iloc[:, 2])
I_nominal = np.array(I_fuente["I pico f (A)"])
#Hoja 1.

#calculamos el valor de w para poder hacer el wt y obtener el angulo de desfase
if F_and[1][0] == 60: #si la frecuncia es 60Hz, el valor de w es directamente 60
    w = 377

else: #si la frecuencia no es 60, entonces se calcula manualmente el valor de w y se redondea a su entero mas cercano
    w = round(2 * math.pi * F_and[1][0], 4)

#Hoja 2.
capacitores = np.array(V_fuente["Cf (uF)"])                 #Escogemos la columna de los capacitores del archivo.

Zcap = ((-1/(w*(capacitores*(10**-6)))))                    #Calculamos las imperancias de los capacitores.
ListZcap.extend (Zcap)
inductores=np.array (V_fuente.iloc[:, 5])                   #Escogemos la columna de los inductores del archivo.

Zinduc = (w*(inductores*(10**-3)))                          #Calculamos las imperancias de los inductores.
ListZinduc.extend(Zinduc)

resistencias=np.array (V_fuente.iloc[:, 4])                 #Escogemos la columna de las resistencias.
ListResistencias.extend(resistencias)

desfa=np.array (V_fuente.iloc[:, 3])                        #Escogemos la columna para del tiempo de desfasaje.

ang = (desfa * w)                                           #Calculamos el wt que usaremos para el angulo de desfasaje.


Vrms = np.array (V_nominal / np.sqrt(2))                    #Calculo del Vrms pero con dudas sobre todos los decimales.
for i in range (len(Vrms)):
    Vrms [i] = round (Vrms [i], 4)                          #Aprroximamos Vrms a 4 decimales.    

V_fasorial = Vrms * (np.cos(ang)) + np.complex_(Vrms * np.sin(ang) * 1j) 
V_fasorial = np.round(V_fasorial, 4)                        #Aproximamos el V en su forma rentangular a 4 decimales.



busi = np.array(V_fuente.iloc[:, 0])
for i, z in combinations(range(len(busi)), 2):             #Hacemos que lea todos la columna de busi, para determinar los nodos en series.
    if busi[i] == busi[z]:
        casa = [resistencias[i] + resistencias[z]]         #Sumamos los terminos en series.
        del ListResistencias[busi[i]] 
        del ListResistencias[busi[z]-1]
        ListResistencias.extend(casa)                      #Agregamos estos valores a la lista de resistenciias y eliminamos los elemntos sumados.
        if busi [i] ==busi [z]:
            paralI = [V_fasorial [i] + V_fasorial [z]]
            del SeriesFi [busi [i]]
            del SeriesFi [busi [z]-1]
            SeriesFi.extend (paralI) 
            if busi[i] == busi[z]:                             #Repetimos el proceso para los capacitores.
                Zc = [Zcap[i] + Zcap[z]]
                del ListZcap[busi[i]]
                del ListZcap[busi[z]-1]
                ListZcap.extend(Zc)
                if busi[i] == busi[z]:                         #Repetimos el proceso para los inductores.
                    Zinducx = Zinduc[i] + Zinduc[z]
                    del ListZinduc[busi[i]]
                    del ListZinduc[busi[z]-1]
                    if isinstance(Zinducx, np.float64):
                        ListZinduc.extend([Zinducx])
                    else:
                        ListZinduc.extend(Zinducx)



#Hoja 3.

DatosResistencias_I = np.array (I_fuente["Rf (ohms)"])          #Seleccionamos las resistencia y las agregamos a la lista.
ListResistencias_I.extend (DatosResistencias_I)

datoscapacitores_I = np.array (I_fuente ["Cf (uF)"])       #Selecionamos los capacitores y calculamos sus imperancias.
Zcapi = (-1 / (w*((datoscapacitores_I)*(10**-6))))         
ListZcap_I.extend (Zcapi)                                  #Guardamos en una lista.

datosinductores_I = np.array (I_fuente ["Lf (mH)"])        #Selecionamos los inductores y calculamos sus imperancias. 
Zinduc_I = (w * (datosinductores_I)*(10**-3))
ListZinduc_I.extend(Zinduc_I)                              #Agregamos en una lista.

corri = np.array (I_fuente ["Corriemento to (seg)"])       #Calculamos el angulo de desfasaje de las fuentes de corrientes.
wt_i = w * corri

Irms = np.array (I_nominal / np.sqrt(2))                    #Calculo del Irms pero con dudas sobre todos los decimales.
for i in range (len(Irms)):
    Irms [i] = round (Irms [i], 4)                          #Aprroximamos Irms a 4 decimales.    

I_fasorial = Irms * (np.cos(wt_i)) + np.complex_(Irms * np.sin(wt_i) * 1j) 
print (I_fasorial)
I_fasorial = np.round(I_fasorial, 4)                        #Aproximamos el I en su forma rentangular a 4 decimales.
SeriesFi.extend (I_fasorial)


busii = np.array(I_fuente["Bus i"])
for i, z in combinations(range(len(busii)), 2):             #Hacemos que lea todos la columna de busi, para determinar los nodos en series.
    if busii[i] == busii[z]:
        cosa = [DatosResistencias_I[i] + DatosResistencias_I[z]]         #Sumamos los terminos en series.
        del ListResistencias_I[busii[i]] 
        del ListResistencias_I[busii[z]-1]
        ListResistencias_I.extend(cosa)                      #Agregamos estos valores a la lista de resistenciias y eliminamos los elemntos sumados.
        if busii [i] ==busii [z]:
            paralI = [I_fasorial [i] + I_fasorial [z]]
            del SeriesFi [busii [i]]
            del SeriesFi [busii [z]-1]
            SeriesFi.extend (paralI) 
            if busii[i] == busii[z]:                             #Repetimos el proceso para los capacitores.
                Zc_I = [Zcapi [i] + Zcapi[z]]
                del ListZcap_I[busii[i]]
                del ListZcap_I[busii[z]-1]
                ListZcap_I.extend(Zc_I)
                if busi[i] == busii[z]:                         #Repetimos el proceso para los inductores.
                    Zinduc_i = Zinduc[i] + Zinduc[z]
                    del ListZinduc_I [busii[i]]
                    del ListZinduc_I[busii[z]-1]
                    if isinstance(Zinducx, np.float64):
                        ListZinduc_I.extend([Zinduc_i])
                    else:
                        ListZinduc_I.extend(Zinduc_i)


print ()
print (ListResistencias_I)
print ()
print (SeriesFi)   
print () 
import pandas as pd
import numpy as np
import math, sys, openpyxl
from decimal import Decimal 
from itertools import combinations
import pandas as pd
#Listas usadas - pagina2.
ListZcap = []
ListZinduc = []
ListResistencias = []
SeriesFv = []
Zgen = []

#Listas usadas - pagina3.

ListResistencias_I = []
ListZcap_I = []
ListZinduc_I = []
SeriesFi = []
Zcapi = []

#Listas usadas - pagina4.

historia_resis = []
historia_induc = []
historia_capa = []

datos = open("C:/Users/Usuario/Desktop/Hola mundo/Proyecto-redes/data_io_copia.xlsx")

#almacenamos las paginas de la tabla de excel
F_and = pd.read_excel ("C:/Users/Usuario/Desktop/Hola mundo/Proyecto-redes/data_io_copia.xlsx" , "f_and_ouput", header= None)
V_fuente = pd.read_excel ("C:/Users/Usuario/Desktop/Hola mundo/Proyecto-redes/data_io_copia.xlsx", "V_fuente")
I_fuente = pd.read_excel ("C:/Users/Usuario/Desktop/Hola mundo/Proyecto-redes/data_io_copia.xlsx", "I_fuente")
Z = pd.read_excel("C:/Users/Usuario/Desktop/Hola mundo/Proyecto-redes/data_io_copia.xlsx", "Z")
VTH_AND_ZTH = pd.read_excel("C:/Users/Usuario/Desktop/Hola mundo/Proyecto-redes/data_io_copia.xlsx", "VTH_AND_ZTH")
Sfuente = pd.read_excel("C:/Users/Usuario/Desktop/Hola mundo/Proyecto-redes/data_io_copia.xlsx", "Sfuente")
S_Z = pd.read_excel("C:/Users/Usuario/Desktop/Hola mundo/Proyecto-redes/data_io_copia.xlsx", "S_Z")
Balance_S = pd.read_excel("C:/Users/Usuario/Desktop/Hola mundo/Proyecto-redes/data_io_copia.xlsx", "V_fuente")

# Ordenamos las filas por la primera columna para ordenar los nodos y la almacenamos de nuevos en las mismas variables.

Z = Z.sort_values(by=Z.columns[0]).reset_index(drop=True)
V_fuente = V_fuente.sort_values(by=V_fuente.columns[0]).reset_index(drop=True)
I = I_fuente.sort_values(by=I_fuente.columns[0]).reset_index(drop=True)

V_nominal = np.array(V_fuente.iloc[:, 2])
I_nominal = np.array(I_fuente["I pico f (A)"])


#Hoja 1.

#calculamos el valor de w para poder hacer el wt y obtener el angulo de desfase
if F_and[1][0] == 60: #si la frecuncia es 60Hz, el valor de w es directamente 60
    w = 377

else: #si la frecuencia no es 60Hz, entonces se calcula manualmente el valor de w y se redondea a su entero mas cercano
    w = round(2 * math.pi * F_and[1][0], 4)

#Hoja 2.
capacitores = np.array(V_fuente.iloc[:, 25])                 #Escogemos la columna de los capacitores del archivo.

Zcap = ((-1/(w*(capacitores*(10**-6)))*1j))                  #Calculamos las imperancias de los capacitores.
Zcap = np.round (Zcap, 4)
ListZcap.extend (Zcap)


inductores=np.array (V_fuente.iloc[:, 5])                   #Escogemos la columna de los inductores del archivo.
Zinduc = (w*(inductores*(10**-3))*1j)                       #Calculamos las imperancias de los inductores.
ListZinduc.extend(Zinduc)

resistencias=np.array (V_fuente["Rf (ohms)"])                 #Escogemos la columna de las resistencias.
ListResistencias.extend(resistencias)

desfa=np.array (V_fuente.iloc[:, 3])                        #Escogemos la columna para del tiempo de desfasaje.

ang = (desfa * w)                                           #Calculamos el producto de wt que usaremos para el angulo de desfasaje.


Vrms = np.array (V_nominal / np.sqrt(2))                    #Calculo del Vrms pero con dudas sobre todos los decimales.
for i in range (len(Vrms)):
    Vrms [i] = round (Vrms [i], 4)                          #Aproximamos Vrms a 4 decimales.    

V_fasorial = Vrms * (np.cos(ang)) + np.complex_(Vrms * np.sin(ang) * 1j) 
V_fasorial = np.complex_(V_fasorial)
V_fasorial = np.round(V_fasorial, 4)                        #Aproximamos el V en su forma rentangular a 4 decimales.

SeriesFv.extend (V_fasorial)
busi = np.array(V_fuente["Bus i"])
for i, z in combinations(range(len(busi)), 2):             #Hacemos que lea todos la columna de busi, para determinar los nodos en series.
    if busi[i] == busi[z]:
        casa = [ListResistencias[i] + ListResistencias[z]]         #Sumamos los terminos en series.
        del ListResistencias[busi[i]] 
        del ListResistencias[busi[z]-1]
        ListResistencias.extend (casa)
        prueba1 = ListResistencias.pop ()
        ListResistencias.insert (i, prueba1)                  #Agregamos estos valores a la lista de resistenciias y eliminamos los elemntos sumados
        if busi[i] == busi[z]:                             #Repetimos el proceso para los capacitores.
            Zc = [Zcap[i] + Zcap[z]]
            del ListZcap[busi[i]]
            del ListZcap[busi[z]-1]
            ListZcap.extend (Zc)
            prueba2 = ListZcap.pop ()
            ListZcap.insert(i, prueba2)
            if busi[i] == busi[z]:
                Zinducx = Zinduc[i] + Zinduc[z]
                del ListZinduc[busi[i]]
                del ListZinduc[busi[z]-1]
                if isinstance(Zinducx, np.float64):
                    ListZinduc.extend([Zinducx])
                    Prueba5 = ListZinduc.pop ()
                    ListZinduc.insert (i, Prueba5)
                else:
                    ListZinduc.extend([Zinducx]) 
                    Prueba6 = ListZinduc.pop ()
                    ListZinduc.insert (i, Prueba6)
for i,z in combinations (range(len(busi)), 2):
                    if busi [i] == busi [z]:
                        paralv = [V_fasorial [i] + V_fasorial [z]]     #Sumamos la fuentes de voltaje que estan en series.
                        del SeriesFv [busi [i]]
                        del SeriesFv [busi [z]-1]
                        SeriesFv.extend (paralv)
                        prueba4 = SeriesFv.pop()
                        SeriesFv.insert(i, prueba4)
                        

#Hoja 3.

DatosResistencias_I = np.array (I_fuente["Rf (ohms)"])          #Seleccionamos las resistencia y las agregamos a la lista.
ListResistencias_I.extend (DatosResistencias_I)

datoscapacitores_I = np.array (I_fuente ["Cf (uF)"])       #Selecionamos los capacitores y calculamos sus imperancias.


for cap in datoscapacitores_I:
     if cap == 0:
          Zcapi = 0
     else:
          Zcapi = (-1 / (w*((datoscapacitores_I)*(10**-6))))*1j 
          Zcapi = np.round (Zcapi, 4)         
ListZcap_I.extend (Zcapi)                                   #Guardamos en una lista.


datosinductores_I = np.array (I_fuente ["Lf (mH)"])        #Selecionamos los inductores y calculamos sus imperancias. 
Zinduc_I = (w * (datosinductores_I)*(10**-3))*1j
Zinduc_I = np.round (Zinduc_I, 4)
ListZinduc_I.extend(Zinduc_I)                              #Agregamos en una lista.

corri = np.array (I_fuente ["Corriemento to (seg)"])       #Calculamos el angulo de desfasaje de las fuentes de corrientes.
wt_i = w * corri

Irms = np.array (I_nominal / np.sqrt(2))                    #Calculo del Irms pero con dudas sobre todos los decimales.
for i in range (len(Irms)):
    Irms [i] = round (Irms [i], 4)                          #Aprroximamos Irms a 4 decimales.    

I_fasorial = Irms * (np.cos(wt_i)) + np.complex_(Irms * np.sin(wt_i) * 1j) 
I_fasorial = np.round(I_fasorial, 4)                        #Aproximamos el I en su forma rentangular a 4 decimales.
SeriesFi.extend (I_fasorial)

busii = np.array(I_fuente["Bus i"])
for i, z in combinations(range(len(busii)), 2):              #operacion para elementos en series I_Fuente.
    if busii[i] == busii[z]:
        cosa = [DatosResistencias_I[i] + DatosResistencias_I[z]]        
        del ListResistencias_I[busii[i]] 
        del ListResistencias_I[busii[z]-1]
        ListResistencias_I.extend(cosa)                      #Agregamos estos valores a la lista de resistenciias y eliminamos los elementos sumados.
        prueba3 = ListResistencias_I.pop()
        ListResistencias_I.insert (i, prueba3)
        if busii[i] == busii[z]:                             #Repetimos el proceso para los capacitores.
            Zc_I = [Zcapi [i] + Zcapi[z]]
            del ListZcap_I[busii[i]]
            del ListZcap_I[busii[z]-1]
            ListZcap_I.extend (Zc_I)
            prueba7 = ListZcap_I.pop()
            ListZcap_I.insert (i, prueba7)
            if busii[i] == busii[z]:                         #Repetimos el proceso para los inductores.
                Zinduc_i = Zinduc[i] + Zinduc[z]
                del ListZinduc_I [busii[i]]
                del ListZinduc_I[busii[z]-1]
                ListZinduc_I.insert (i , Zinduc_i)

for b, t in combinations(range(len(busii)), 2):              #Operacion para fuente de corrientes en series.
    if busii [b] == busii [t]:                               
            paralI = [I_fasorial [b] + I_fasorial [t]]       #Juntamos la fuentes de corrientes en series y agregamos a final de la lista.
            del SeriesFi [busii [b]]
            del SeriesFi [busii [t]-1]
            SeriesFi.extend(paralI)
            prueba8 = SeriesFi.pop ()
            SeriesFi.insert (i, prueba8)

#Hoja4.

# Obtenemos los valores de resistencias, inductancias y capacitancias de los datos de entrada
r = np.array(Z["R (ohms)"])
historia_resis.extend(r)

l = np.array(Z["L (uH)"])
historia_induc.extend((w *(((l))*10**-6))*1j)
historia_induc = np.round (historia_induc, 4)

c = np.array(Z["C (uF)"])
for cap in c:
     if cap == 0:
          historia_capa.append(0)
     else:
          historia_capa.append(-1j / (w*(cap*(10**-6))))

historia_capa = np.array(historia_capa)
historia_capa = np.round (historia_capa, 4)

bus_i = np.array(Z["Bus i"])
bus_j = np.array(Z["Bus j"])

# Buscamos elementos en paralelo
for i, j in combinations(range(len(bus_i)), 2):
    if bus_i[i] == bus_i[j] and bus_j[i] == bus_j[j]:                                                    # Calculamos las equivalentes en paralelo

        #Calculamos la o las resistencias en paralelo.
        
        historia_resis = np.where(historia_resis == 0, 1e+15, historia_resis)
        HellomotoRE = (1/(1/historia_resis [j] + 1/historia_resis [i]))
        historia_resis = np.delete (historia_resis, [i , j])                                                                  #Eliminamos las resistencias utilizadas para el paralelo.
        historia_resis = np.insert (historia_resis, i , HellomotoRE)
        historia_resis = np.round (historia_resis, 4)                                                 #Agregamos el resultados de la resistencia en paralelo en la lista de las resistencias.

        #Calculamos la o las inductancias en paralelo.

        historia_induc = np.where(historia_induc == 0, 1e+15, historia_induc)
        NicolaInduc = (1 / (1 / historia_induc[j] + 1 / historia_induc[i]))
        historia_induc = np.delete(historia_induc, [i , j])                                     #Eliminamos las Inductancias utilizadas para calcular el paralelo.
        historia_induc = np.insert(historia_induc, i , NicolaInduc)                             #Agregamos el resultado del paralelo en la lista de inductancias.
        historia_induc = np.round (historia_induc, 4)      

        #Calculemos la o las capacitancias en paralelo. 
        historia_capa = np.where(historia_capa == 0, 1e+15, historia_capa)
        Motomami = (1 / (1 / historia_capa[j] + 1 / historia_capa[i]))
        historia_capa = np.delete(historia_capa, [i , j])                                       #Eliminamos las Inductancias utilizadas para calcular el paralelo. 
        historia_capa = np.insert(historia_capa, i , Motomami)                                  ##Agregamos el resultado del paralelo en la lista de inductancias.
        historia_capa = np.round (historia_capa, 4)

Busisnodos = np.append(bus_i, bus_j)
busjnodos = np.append(busii, busi)
busto = np.append(Busisnodos, busjnodos)
nodos = np.unique (busto)

if nodos[0] == 0:
    nodos = np.delete (nodos , 0)

    c = len(nodos)
elif nodos [0] != 0:
     c = len(nodos)
else:
    print ("Pega un warnign")
         
    #Admitancias de la hoja Z


    #Llamamos las listas ListResistencias, ListZcap, ListZinduc para calcular las Z del generador calcular las corrientes inyectadas por las fuentes de voltaje.
    Zgen = np.sum((ListResistencias,ListZcap, ListZinduc,),axis=0)  
    AdmitanciasGenV = (1 / Zgen)
    AdmitanciasGenV = np.round (AdmitanciasGenV , 4)
    Iinyectadas = np.divide (SeriesFv,Zgen)
    Iinyectadas = np.round (Iinyectadas, 4)



    #Llamamos las listas ListResistencias, Listcap, ListZinduc para calcular las Z de las fuentes de corrientes hacia el circuito.
    Zigen = np.sum((ListResistencias_I,ListZcap_I, ListZinduc_I),axis=0)  
    Ifinyectadas = np.divide (SeriesFi,Zigen)
    Ifinyectadas = np.round (Ifinyectadas, 4)

    #Llamamos las listas Elemncapaci, Elemeninduc, Elemenresis para calcular las Z de los elementos conectados al circuito.

    #Aqui saque las admitancias de la pagina Z

    AdminHistoria_Resis = 1/ np.array (historia_resis) 
    AdminHistoria_induc = 1/np.array (historia_induc)

    # Reemplazar valores cero por 1e-15
    historia_capa_sin_cero = np.where(historia_capa == 0, 1e+15, historia_capa)

    # Aplicar la función de inversión
    AdminHistoria_capa = 1 / historia_capa_sin_cero

    AdminHistoria_capa = np.round (AdminHistoria_capa, 4)

#Admitancias de 


#PLanteamos una matriz que nula, para fabricar la nueva matriz de admitancias.
Yelemnt = np.zeros((c, c), dtype= np.complex64)
AdminZelement = np.sum ((AdminHistoria_Resis, AdminHistoria_induc, AdminHistoria_capa), axis = 0)
print (AdminZelement)
print ()
np.fill_diagonal(Yelemnt, AdminZelement )

print(Yelemnt)

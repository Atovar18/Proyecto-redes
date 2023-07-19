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
ListZcap_I = []
Zcapi = []

#Listas usadas - pagina4.

historia_resis = []
historia_induc = []
historia_capa = []

datos = pd.read_excel("data_io_copia.xlsx")

#almacenamos las paginas de la tabla de excel
F_and = pd.read_excel ("data/data_io_copia.xlsx" , "f_and_ouput", header= None)
V_fuente = pd.read_excel ("C:/Users/Usuario/Desktop/Hola mundo/Proyecto-redes/data_io_copia.xlsx", "V_fuente")
I_fuente = pd.read_excel ("C:/Users/Usuario/Desktop/Hola mundo/Proyecto-redes/data_io_copia.xlsx", "I_fuente")
Z = pd.read_excel("C:/Users/Usuario/Desktop/Hola mundo/Proyecto-redes/data_io_copia.xlsx", "Z")
VTH_AND_ZTH = pd.read_excel("C:/Users/Usuario/Desktop/Hola mundo/Proyecto-redes/data_io_copia.xlsx", "VTH_AND_ZTH")
Sfuente = pd.read_excel("C:/Users/Usuario/Desktop/Hola mundo/Proyecto-redes/data_io_copia.xlsx", "Sfuente")
S_Z = pd.read_excel("C:/Users/Usuario/Desktop/Hola mundo/Proyecto-redes/data_io_copia.xlsx", "S_Z")
Balance_S = pd.read_excel("C:/Users/Usuario/Desktop/Hola mundo/Proyecto-redes/data_io_copia.xlsx", "V_fuente")

# Ordenamos las filas por la primera columna para ordenar los nodos y la almacenamos de nuevos en las mismas variables.

Z = Z.sort_values(by=Z.columns[0]).reset_index(drop=True)
print (Z)
V_fuente = V_fuente.sort_values(by=V_fuente.columns[0]).reset_index(drop=True)
I = I_fuente.sort_values(by=I_fuente.columns[0]).reset_index(drop=True)

# def sort_excel(data : pd.DataFrame):
#      pass


V_nominal = np.array(V_fuente.iloc[:, 2])
I_nominal = np.array(I_fuente["I pico f (A)"])


#Hoja 1.

w = round(2 * math.pi * F_and[1][0], 4)

#Hoja 2.
capacitores = np.array(V_fuente.iloc[:, 25])                 #Escogemos la columna de los capacitores del archivo.
ListZcap = []
print (capacitores)
print ()
for cap in capacitores:
    if cap == 0:                                             #Analizamos los datos de la lista para que el código no realice la división si el capacitor = 0
        Zcap = 0
    else:
        Zcap = (-1/(w*(cap*(10**-6)))*1j)                    #Calculamos las impedancias de los capacitores.
    ListZcap.append(np.round(Zcap, 4))

print (ListZcap)
print ()


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
            Zc = [ListZcap[i] + ListZcap[z]]
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


print (datoscapacitores_I)
print ()
for comp in enumerate(datoscapacitores_I):
    if comp == 0:
        Zcompi = 0
    else:
        Zcompi = (-1 / (w * ((comp)*(10**-6)))) * 1j 
        Zcompi = np.round(Zcompi, 4)         
    ListZcap_I[comp] = Zcompi                                         #Guardamos en una lista.


print (ListZcap_I)
print ()


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
historia_capa = [] 

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
        
        HellomotoRE = (1/(1/historia_resis [j] + 1/historia_resis [i]))
        historia_resis.pop (i)                                                                  #Eliminamos las resistencias utilizadas para el paralelo.
        historia_resis.pop (j)
        historia_resis.insert (i , HellomotoRE)                                                 #Agregamos el resultados de la resistencia en paralelo en la lista de las resistencias.

        #Calculamos la o las inductancias en paralelo.

        NicolaInduc = (1 / (1 / historia_induc[j] + 1 / historia_induc[i]))
        historia_induc = np.delete(historia_induc, [i , j])                                     #Eliminamos las Inductancias utilizadas para calcular el paralelo.
        historia_induc = np.insert(historia_induc, i , NicolaInduc)                             #Agregamos el resultado del paralelo en la lista de inductancias.
        historia_induc = np.round (historia_induc, 4)    

        #Calculemos la o las capacitancias en paralelo. 
        Motomami = (1 / (1 / historia_capa[j] + 1 / historia_capa[i]))
        historia_capa = np.delete(historia_capa, [i , j])                                       #Eliminamos las Inductancias utilizadas para calcular el paralelo. 
        historia_capa = np.insert(historia_capa, i , Motomami)                                  ##Agregamos el resultado del paralelo en la lista de inductancias.
        historia_capa = np.round (historia_capa, 4)

Busisnodos = np.append(bus_i, bus_j)
busjnodos = np.append(busii, busi)
busto = np.append(Busisnodos, busjnodos)
nodos = np.unique (Busisnodos)

if nodos[0] == 0:
    nodos = np.delete (nodos , 0)
    c = len(nodos)
    print ("toma")
    print ()
elif nodos [0] != 0:
     c = len(nodos)
else:
    print ("Pega un warnign")
            

                                 
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

historia_resis_sin_cero = np.where (historia_resis == 0, 1e+15, historia_resis)

AdminHistoria_Resis = 1/ np.array (historia_resis_sin_cero) 
AdminHistoria_Resis = np.round(AdminHistoria_Resis, 4)

# Reemplazar valores cero por 1e-15
historia_induc_sin_cero = np.where(historia_induc == 0, 1e+15, historia_induc)

# Aplicar la función de inversión
AdminHistoria_induc = 1/ historia_induc_sin_cero
AdminHistoria_induc = np.round (AdminHistoria_induc, 4)

# Reemplazar valores cero por 1e-15
historia_capa_sin_cero = np.where(historia_capa == 0, 1e+15, historia_capa)

# Aplicar la función de inversión
AdminHistoria_capa = 1 / historia_capa_sin_cero

AdminHistoria_capa = np.round (AdminHistoria_capa, 4)

#Admitancias de V_fuente.

Z_totV = np.sum((ListZcap, ListZinduc, resistencias), axis = 0)
Z_V = np.where(Z_totV == 0, 1e+15, Z_totV)

Y_V = 1/Z_V
Y_V = np.round(Y_V, 4)

#Admitancia de I_fuente:

Z_totI = np.sum((Zcapi, ListZinduc_I, ListResistencias_I), axis = 0)

Z_I = np.where(Z_totI == 0, 1e+15, Z_totI)
Y_I = 1/Z_totI
Y_I = np.round(Y_I, 4)

Y_eq = np.sum((AdminHistoria_capa, AdminHistoria_induc, AdminHistoria_Resis), axis = 0)


#Planteamos la Ybus:
bus_jp = np.where(bus_j == 0, bus_i, bus_j)

def matriz_admittancia_Z(Y_eqq, bus__i, bus__j):       # Definimos la matriz de admittancia nodal
    n = max(max(nodos), max(nodos))              # Definimos el tamaño de la matriz en función de la cantidad de nodos.
    Y_Z = np.zeros((n, n), dtype=complex)
    for i in range(len(Y_eq)):
        Y_eqq = Y_eq[i]
        bus__i = bus_i[i] - 1                    # Restamos 1 para ajustar al índice de Python
        bus__j = bus_jp[i] - 1
        Y_Z[bus__i, bus__j] = (-1*Y_eqq)                 # Definimos las admitancias y su signo según su posición en la matriz
        Y_Z[bus__j, bus__i] = (-1*Y_eqq)
        
        Y_Z[bus__j, bus__j] = Y_eqq
    return Y_Z
    
Y_Z = matriz_admittancia_Z(Y_eq, bus_i, bus_j)

print()
print("Admitancias de la hoja Z")
print(Y_Z)

n = max(max(nodos), max(nodos))
Admitt_V = np.zeros((n, n), dtype= np.complex64)

def matriz_admittancia_V(Y_V, busi1):      
    n = max(max(nodos), max(nodos))             
    Admitt_V = np.zeros((n, n), dtype=complex)
    for i in range(len(Y_V)):
        Y_Vq = Y_V[i]
        busi1 = busi[i] - 1                    
        Admitt_V[busi1, busi1] = Y_Vq
    return Admitt_V
Admitt_V = matriz_admittancia_V(Y_V, busi)
print()
print("Matriz hoja de Vfuente")
print(Admitt_V,)
print ()

def matriz_admittancia_I(Y_I, busii1):      
    n = max(max(nodos), max(nodos))             
    Admitt_I = np.zeros((n, n), dtype=complex)
    for i in range(len(Y_I)):
        Y_Iq = Y_I[i]
        busii1 = busii[i] - 1                    
        Admitt_I[busii1, busii1] = Y_Iq
    return Admitt_I
Admitt_I = matriz_admittancia_I(Y_I, busii)



Y_bus = np.sum((Y_Z, Admitt_V, Admitt_I), axis = 0)

print (Y_bus)


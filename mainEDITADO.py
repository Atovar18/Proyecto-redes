import pandas as pd
import numpy as np
import impedancia
import ybus
import potencia
import compensadores
import sys

#Lector de archivo .xlsx
df_gen = pd.read_excel("C:\\Users\\PC\\Documents\\USB\\Estudios\\Prj\\USB\\Humberto - Sora\\ProyectoSEP\\AnalisiSEP\\data_io.xlsx","GENERATION")                 #
df_lines = pd.read_excel("C:\\Users\\PC\\Documents\\USB\\Estudios\\Prj\\USB\\Humberto - Sora\\ProyectoSEP\\AnalisiSEP\\data_io.xlsx","LINES")                    #
df_load = pd.read_excel("C:\\Users\\PC\\Documents\\USB\\Estudios\\Prj\\USB\\Humberto - Sora\\ProyectoSEP\\AnalisiSEP\\data_io.xlsx","LOAD")                      #
df_vnom = pd.read_excel("C:\\Users\\PC\\Documents\\USB\\Estudios\\Prj\\USB\\Humberto - Sora\\ProyectoSEP\\AnalisiSEP\\data_io.xlsx", "V_NOM")                    #

#N° de barras del SEP
Barras_i = max(df_lines.iloc[:,0])
Barras_j = max(df_lines.iloc[:,1])
Numero_Barras = int(max(Barras_i,Barras_j))



# DATOS
#GENERADOR:
Generador_resist = np.array(df_gen.iloc[:,4])                     #Impedancia resistiva
Generador_react = np.array(df_gen.iloc[:,5])                      #Impedancia reactiva
Teta = np.array(df_gen.iloc[:,3],dtype="float_")                  #Angulo° fasor voltaje
Voltaje = np.array(df_gen.iloc[:,2])                              #V de fuente
G_barra_i = np.array(df_gen.iloc[:,0])                            #Barra de conexion i
G_barra_j = np.full((len(df_gen.iloc[:,0])),0)                    #Barra de conexion j




#CARGAS:
Carga_resist = np.array(df_load.iloc[:,9])                       #Zresistiva
Carga_react = np.array(df_load.iloc[:,10])                      #Zreactiva
Carga_type = np.array(df_load.iloc[:,2])                            #Tipo de carga
Barra_Carga_i = np.array(df_load.iloc[:,0])                         
Barra_Carga_j = np.full((len(df_load.iloc[:,0])),0)                
CargaxD = np.concatenate(([Barra_Carga_i],[Barra_Carga_j]))     #Matriz de conexion de las cargas
CargaxD = np.transpose(CargaxD)


#LINEAS
#Datos LINES:
df_lines.dropna()
Linea_resist = np.array(df_lines.iloc[:,5])                      #Impedancia resistiva
Linea_react = np.array(df_lines.iloc[:,6])                      #Impedancia reactiva
p_shunt = np.array(df_lines.iloc[:,7])                              #admitancia del efecto capacitivo de las lineas
Long = np.array(df_lines.iloc[:,4])                             #largo de las lineas
Linea_Barra_i = np.array(df_lines.iloc[:,0])                        #Barra de conexion i
Linea_Barra_j = np.array(df_lines.iloc[:,1])                        #Barra de conexion j
LineaxD = np.concatenate(([Linea_Barra_i],[barra_linea_j]))     #Matriz de conexion de las cargas
LineaxD = np.transpose(LineaxD)


#VOLTAJE NOMINAL
#Datos V_NOM
Voltaje_n = float(df_vnom.iloc[0,1])                                #Voltaje nominal
Voltaje_max = float(df_vnom.iloc[0,3])                                    #Voltaje máximo segun COVENIN 159
Voltaje_min = float(df_vnom.iloc[0,2])                                    #Voltaje mínimo segun COVENIN 159
#Warning:
if Voltaje_n<0 or Voltaje_min<0 or Voltaje_max<0:
    mensaje= "Valor negativo"
    print("\n[*] Error 159: Valor de datos ingresados no es reconocido")
    print("\tRevisar la casilla warning del excel __data_io.xlsx__ en la hoja V_NOM\n")
    sys.exit(-1)
elif df_vnom.iloc[0,1:4].isnull().sum() > 0:
    mensaje = "Casilla vacia"
    print("\n[*] Error 159: Valor de datos ingresados no es reconocido")
    print("\tRevisar la casilla warning del excel __data_io.xlsx__ en la hoja V_NOM\n")
    sys.exit(-1)
        



def run():
#---------------------------------------------- CALCULO DE LAS IMPEDANCIAS---------------------------------------------- 
    #Generadores
    ZP_g = impedancia.generador(Generador_resist,Generador_react)
    1_ge =  np.concatenate(([G_barra_i],[G_barra_j],[ZP_g]), axis=0)
    1_ge = np.transpose(1_gen)
    
    
    #Cargas
    ZP_c = impedancia.carga(Carga_resist, Carga_react, Carga_type)
    carga = np.concatenate(([Barra_Carga_i],[Barra_Carga_j],[ZP_c]),axis=0)
    carga = np.transpose(carga)
    #print(imp_carga)


    #Lineas
    ZP_L, y_shunt = impedancia.linea(Linea_resist,Linea_react,Long, b_shunt)
    linea = np.concatenate(([Linea_Barra_i],[Linea_Barra_j],[imp_linea]),axis=0)      
    linea = np.transpose(linea)
    #print(imp_linea)
    linea_d = np.concatenate(([Linea_resist],[Linea_react]),axis=0)
    linea_d = np.transpose(linea_d)

   
#------------------------------------------------- CALCULO DE YBUS, VTH Y ZTH --------------------------------------------------
    #Corrientes inyectadas
    Iny_corrientes = impedancia.corrientes(voltaje,teta,ZP_g,Numero_Barras, barra_gen_i)

    #Y bus
    Y_bus = ybus.ybus(1_gen, carga, linea, Numero_Barras,Linea_Barra_i,Linea_Barra_j,y_shunt, longitud)
    Y_bus = np.round(Y_bus,4)
    #print(y_bus)


    #Z de thevenin
    Zth, Zbus = ybus.Zth(Y_bus)

    #Voltajes de thevenin
    Vth,Vth_rect = ybus.Vth(Zbus,Iny_corrientes,Numero_Barras)
    #print(vth)
    
    #GBUS
    g_bus = ybus.gbus(y_bus,Numero_Barras)

    #BBUS
    b_bus = ybus.bbus(y_bus,Numero_Barras)

#------------------------------------------------ COMPENSADORES ------------------------------------------------------------

    compe1, compe2 = compensadores.test_compen(vth, Voltaje_max, Voltaje_max, Numero_Barras, Voltaje_n)
    #print(check_com1,check_com2)

    compeX = compensadores.compensador_pasivo(Numero_Barras,vth,Voltaje_min, Voltaje_max, zbus, Voltaje_n)
    #print(x_comp)

#------------------------------------------------ CALCULO DE LAS POTENCIAS ------------------------------------------------

    #Potencia del generador
    Pot_generador, Q_generador= potencia.generador(ZP_g, voltaje, teta, vth_rect, Barras_i)
    #print(p_gen)

    #Potencia de la carga
    s_carga, p_carga, q_carga = potencia.Cargas(ZP_c, vth_rect, Barra_Carga_i)

    #Lineflow (Flujo de potencias)
    p_ij, q_ij, p_ji, q_ji = potencia.lineflow(LineaxD, linea_d, Long, vth_rect)

    #Balance de potencias
    delta_p, delta_q = potencia.balance(Pot_generador, Q_generador, p_carga, q_carga)
    #print(delta_p)
    #print(delta_q)




if __name__ == "__main__":
    run()

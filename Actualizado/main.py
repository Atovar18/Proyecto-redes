import pandas as pd
import numpy as np
import impedancia
import ybus
import potencia
import compensadores


df_gen = pd.read_excel("C:\\Users\\Usuario\\Desktop\\Hola mundo\\Proyecto-redes\\rEscAtA sIstEmAs\\data_io.xlsx","GENERATION")                 #
df_lines = pd.read_excel("C:\\Users\\Usuario\\Desktop\\Hola mundo\\Proyecto-redes\\rEscAtA sIstEmAs\\data_io.xlsx","LINES")                    #
df_load = pd.read_excel("C:\\Users\\Usuario\\Desktop\\Hola mundo\\Proyecto-redes\\rEscAtA sIstEmAs\\data_io.xlsx","LOAD")                      #
df_vnom = pd.read_excel("C:\\Users\\Usuario\\Desktop\\Hola mundo\\Proyecto-redes\\rEscAtA sIstEmAs\\data_io.xlsx", "V_NOM")                    #
df_comp = pd.read_excel("C:\\Users\\Usuario\\Desktop\\Hola mundo\\Proyecto-redes\\rEscAtA sIstEmAs\\data_io.xlsx", "REACTIVE_COMP")            #



#N° de barras del SEP
Barras_i = max(df_lines.iloc[:,0])
Barras_j = max(df_lines.iloc[:,1])
Numero_Barras = int(max(Barras_i,Barras_j))



#GENERADOR:
Generador_resist = np.array(df_gen.iloc[:,4])                     #Impedancia resistiva
Generador_react = np.array(df_gen.iloc[:,5])                      #Impedancia reactiva
Teta = np.array(df_gen.iloc[:,3],dtype="float_")                  #Angulo° fasor voltaje
Voltaje = np.array(df_gen.iloc[:,2])                              #V de fuente
G_barra_i = np.array(df_gen.iloc[:,0])                            #Barra de conexion i
G_barra_j = np.full((len(df_gen.iloc[:,0])),0)                    #Barra de conexion j




#CARGAS:
Carga_resist = np.array(df_load.iloc[:,9])                      #Zresistiva
Carga_react = np.array(df_load.iloc[:,10])                      #Zreactiva
Carga_type = np.array(df_load.iloc[:,2])                        #Tipo de carga
Barra_Carga_i = np.array(df_load.iloc[:,0])                         
Barra_Carga_j = np.full((len(df_load.iloc[:,0])),0)                
CargaxD = np.concatenate(([Barra_Carga_i],[Barra_Carga_j]))     #Matriz de conexion de las cargas
CargaxD = np.transpose(CargaxD)


#LINEAS
#Datos LINES:
df_lines.dropna()
Linea_resist = np.array(df_lines.iloc[:,5])                     #Impedancia resistiva
Linea_react = np.array(df_lines.iloc[:,6])                      #Impedancia reactiva
p_shunt = np.array(df_lines.iloc[:,7])                          #admitancia del efecto capacitivo de las lineas
Long = np.array(df_lines.iloc[:,4])                             #largo de las lineas
Linea_Barra_i = np.array(df_lines.iloc[:,0])                        #Barra de conexion i
Linea_Barra_j = np.array(df_lines.iloc[:,1])                        #Barra de conexion j
LineaxD = np.concatenate(([Linea_Barra_i],[Linea_Barra_j]))     #Matriz de conexion de las cargas
LineaxD = np.transpose(LineaxD)


#VOLTAJE NOMINAL
#Datos V_NOM
Voltaje_n = float(df_vnom.iloc[0,1])                                      #Voltaje nom
Voltaje_max = float(df_vnom.iloc[0,3])                                    #Voltaje máx COVENIN 159
Voltaje_min = float(df_vnom.iloc[0,2])                                    #Voltaje mín COVENIN 159

        
#COMPENSADORES
Barra_p_compensador_i = np.array(df_comp.iloc[:,0])
Barra_p_compensador_j = np.full((len(df_comp.iloc[:,0])),0)
Compensador_type = np.array(df_comp.iloc[:,2])
Volt_compensador = np.array(df_comp.iloc[:,3])
Compensador_Q = np.array(df_comp.iloc[:,4])
Compensador_X = np.array(df_comp.iloc[:,5])

def run():
    #Generadores
    ZP_g = impedancia.generador(Generador_resist,Generador_react)
    Gzp = np.concatenate(([G_barra_i],[G_barra_j],[ZP_g]), axis=0)
    Gzp = np.transpose(Gzp)
    
    
    #Cargas
    ZP_c = impedancia.carga(Carga_resist, Carga_react, Carga_type)
    carga = np.concatenate(([Barra_Carga_i],[Barra_Carga_j],[ZP_c]),axis=0)
    carga = np.transpose(carga)
    

    #Lineas
    ZP_L, y_shunt = impedancia.linea(Linea_resist,Linea_react,Long, p_shunt)
    linea = np.concatenate(([Linea_Barra_i],[Linea_Barra_j],[ZP_L]),axis=0)      
    linea = np.transpose(linea)
    
    linea_d = np.concatenate(([Linea_resist],[Linea_react]),axis=0)
    linea_d = np.transpose(linea_d)

    y_comp, barra_bus = impedancia.compensadores(Compensador_X,Barra_p_compensador_i,Compensador_type)     #compensadores


    Iny_corrientes = impedancia.corrientes(Voltaje,Teta,ZP_g,Numero_Barras, G_barra_i)

    Y_bus = ybus.ybus(Gzp, carga, linea, Numero_Barras,Linea_Barra_i,Linea_Barra_j,y_shunt, Long, barra_bus , y_comp)
    Y_bus = np.round(Y_bus,4)


    Zth, Zbus = ybus.Zth(Y_bus)
    Vth,Vth_rect = ybus.Vth(Zbus,Iny_corrientes,Numero_Barras)
    g_bus = ybus.gbus(Y_bus,Numero_Barras)
    b_bus = ybus.bbus(Y_bus,Numero_Barras)


    compe1, compe2 = compensadores.test_compen(Vth, Voltaje_max, Voltaje_max, Numero_Barras, Voltaje_n)

 
    verificador = list(filter(lambda x: 'CAP' in x, compe2))                                                       #revisa si se necesita compesacion
    verificador2 = list(filter(lambda x: 'IND' in x, compe2))
    print(verificador)
    print(verificador2)

    if len(verificador) == 0 or len(verificador2) == 0:
        print("\n  [*] No es necesario compensar, según COVENIN 159")
    else:
        compeX = compensadores.compensador_pasivo(Numero_Barras,Vth,Voltaje_min, Voltaje_max, Zbus, Voltaje_n)
        print("\n  [*] Se requiere compensación en las barras:")
        for i in range(len(compe2)):
            print("\t", compe1[i],"---",compe2[i])
        print("\n  [*] Se recomiendan los siguientes compensadores: ")
        for i in range(len(compeX)):
            print("\t",compeX[i])

    

    #Potencias:
    Pot_generador, Q_generador= potencia.generador(ZP_g, Voltaje, Teta, Vth_rect, G_barra_i)

    p_carga, q_carga = potencia.Cargas(ZP_c, Vth, Barra_Carga_i)
   
    p_ij, q_ij, p_ji, q_ji = potencia.lineflow(LineaxD, linea_d, Long, Vth_rect)

    Balance_P, Balance_q = potencia.balance(Pot_generador, Q_generador, p_carga, q_carga)



if __name__ == "__main__":
    run()

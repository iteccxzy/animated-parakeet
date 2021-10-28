#!/usr/bin/env python
# coding: utf-8

# In[ ]:
import os
import pandas as pd
import datetime


main_dict = {}
cols      = ['fecha', 'entrada', 'salida', 'hs_trab', 'observaciones' , 'nombre']

def iter_files(list_files, df_obs, downloads):    
    for i in range(len(cols)):
                main_dict[cols[i]] = []  
    for l in list_files:
        files     = os.path.join(downloads, l)
        data_file = pd.read_excel(files)

        nombre    = data_file['Nombre'][1]  
        fh_series = data_file['Fecha/Hora']   

        #contar si marca 1 o 2 veces
        cont = {}
        for r in fh_series:
            fecha, _ = str(r).split(' ')
            if fecha in cont:
                cont[fecha] = 2
            else:
                cont[fecha] = 1

        #si marco una vez agregar otra marca
        for i, j in cont.items():    
            if j == 1:     
                dt_string = f'{i} 00:00:00'
                i      = datetime.datetime.strptime(dt_string, "%Y-%m-%d %H:%M:%S")           
                s1     = pd.Series(i)            
                fh_series = fh_series.append(s1)            

        #ordernar columna luego de agregar fila        
        fh_series = fh_series.sort_values()

        #agregar salto de linea
        fila  = [ '', '', '' , '', '', '']
        fila1 = [ 'Nombre', f'{nombre}', '' , '', '', '']
        for i in range(len(cols)):
            main_dict[cols[i]].append(fila[i])
            main_dict[cols[i]].append(fila1[i]) 
            main_dict[cols[i]].append(cols[i])

        #recorrer la columna fecha-hora
        temp = {}  
        for row in fh_series:
            fecha, hora = str(row).split(' ')
            try:
                if hora == '00:00:00':
                    hora = 'None'
                else:               
                    hora = datetime.datetime.strptime(hora, '%H:%M:%S')               
            except:
                hora = 'None'    

            #si es el primer registro de esa fecha        
            if fecha not in temp:
                temp[fecha] = []
                temp[fecha].append(fecha)
                temp[fecha].append(hora)

            #si es el segundo registro de esa fecha
            else:
                temp[fecha].append(hora)
                fecha   = temp[fecha][0]
                entrada = str(temp[fecha][1])[11:16]
                salida  = str(temp[fecha][2])[11:16]  
                obs     = '----'
                try:
                    horas_t = str(temp[fecha][2] - temp[fecha][1])[:4]
                except:
                    horas_t = 'None'  

                #recorrer observaciones
                p_gen = df_obs.iterrows()
                for o in p_gen:
                    nan = o[1][1]
                    fecha2, _ = str(o[1][2]).split(' ')                
                    if  nombre == o[1][0]  and fecha == fecha2  and not pd.isna(nan):
                        obs = o[1][1]                                               

                #agregamos el registro a main_dict
                fila = [fecha, entrada, salida , horas_t, obs, nombre]            
                for i in range(len(cols)):
                    main_dict[cols[i]].append(fila[i]) 
        
        #agregar las observaciones de los dias que no marcó
        p_gen = df_obs.iterrows()        
        for o in p_gen:
            nan = o[1][1]
            fecha2, _ = str(o[1][2]).split(' ')
            if fecha2 not in temp.keys() and nombre == o[1][0]:
                obs = o[1][1]
                
               
                fila = [fecha2, '----', '----' , '----', obs, nombre]
                for i in range(len(cols)):
                    main_dict[cols[i]].append(fila[i])
    return main_dict


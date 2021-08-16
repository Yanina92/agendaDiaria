# -*- coding: utf-8 -*-
"""
Created on Mon Oct  5 21:21:09 2020

@author: Yanina Vanesa Gonzalez

"""
#%%

import pandas as pd # Librería para el manejo de datos.

import tkinter as tk # Librería la cual me permite crear interfases para que interactue el usuario. Importa la librería tkinter con todas las clases.

from tkinter import * #Llamo a todas las funciones dentro de tkinter

import os # (operating system) Librería que se comunica con el sistema operativo para indicarle que realice la acción que quiero.

from datetime import datetime #Librería para manejo de tiempo

import plotly.express as px #Librería para graficar el diagrama de gantt

#%%

file_path = "C:/"

#%%

def hola():         #Esto es para probar los botones
    print ("Hola!")
    
#Fin hola

#%%
 
def ventana_DIA():
    
    def hor_DIA():

        #Obtengo las variables de mi interfaz gráfica
        
        DIA = int(Entrada_12.get())
        
        MES = Entrada_22.get()
          
        AN = Entrada_32.get()
        
        #Abro el archivo donde está mi dataframe
        
        notas_4=pd.read_excel(file_path + "Cronograma Actividades/Mensual/"+str(AN)+"/Agenda "+str(MES)+" "+str(AN)+".xlsx")
        
        notas_42= notas_4[(notas_4["día"] == DIA)] #Arma un nuevo dataframe donde se toman los datos de la PC que necesito.
        
        notas_42.to_excel(file_path + "Cronograma Actividades/Diarios/"+str(AN)+"/Agenda "+str(DIA)+" "+str(MES)+" "+str(AN)+".xlsx", index=False)
        
        os.startfile(file_path + "Cronograma Actividades/Diarios/"+str(AN)+"/Agenda "+str(DIA)+" "+str(MES)+" "+str(AN)+".xlsx")
   
    # Fin hor_DIA
    
    def diag_DIA():
      
        #Obtengo las variables de mi interfaz gráfica
        
        DIA = int(Entrada_12.get())
        
        MES = Entrada_22.get()
          
        AN = Entrada_32.get()
        
        #Abro el archivo donde está mi dataframe
        
        notas_4=pd.read_excel(file_path + "Cronograma Actividades/Mensual/"+str(AN)+"/Agenda "+str(MES)+" "+str(AN)+".xlsx")
        
        notas_42= notas_4[(notas_4["día"] == DIA)] #Arma un nuevo dataframe donde se toman los datos de la PC que necesito.
        
        Fecha=[]
        
        #Arma un vector con formato de fecha
        
        for i in range (0,len(notas_42),1):
            
            f_modif= datetime(notas_42.Año[i] , notas_42.Mes[i] , notas_42.día[i])
            
            Fecha.append(f_modif.strftime("%Y-%m-%d"))
            
        #Armo mejor los vectores horario.    
        
        h_ini=notas_42["Hora de inicio"]
        
        h_fin=notas_42["Hora de finalización"]
        
        INI_1=[]
        
        FIN_1=[]
        
        for i in range (0,len(notas_42),1):
            
            date= Fecha[i]+" "+h_ini[i]
            
            INI_1.append(date)
            
        for i in range (0,len(notas_42),1):
            
            date_1= Fecha[i]+" "+h_fin[i]
            
            FIN_1.append(date_1)   
             
        #Armo el dataframe general que va a usar el diagrama de Gantt
        
        notas_43= pd.DataFrame(INI_1)
        
        notas_43["Final"]=FIN_1
    
        notas_43=notas_43.rename(columns={0:"Inicio"}) #Cambia de nombre la columna
        
        notas_43["Actividad:"]=notas_42["Actividad"]
        
        notas_43["Observaciones"]=notas_42["Observaciones"]
        
        notas_43["Día"]=Fecha
        
        #Arma el diagrama de Gantt
        
        fig = px.timeline(notas_43, x_start="Inicio", x_end="Final", y="Día", color="Actividad:") #Armo el dograma indicando todas las partes necesarias
    
        fig.update_yaxes(visible=False,showticklabels=False) #Escondo los marcadores del eje y
        
        fig.update_layout(legend=dict(orientation="h", y=1.1)) ##Cambio la orientación de la leyenda
        
        fig.show() #Muestra el diagrama utilizando el explorador de internet predeterminado.
        
    # Fin diag_DIA    
    
    #Creo una ventana
    
    window_2 = tk.Tk()
    
    #Le aplico un título a mi ventana
    
    window_2.title("Ver horario por día")
    
    #Le Agrego el ícono a la ventana

    window_2.iconbitmap("ilustracion.ico") # https://www.online-convert.com/es/
    
    #Determino el tamaño de mi ventana
    
    window_2.geometry("600x230") #(largo, alto)
    
    #Evito que el usuario me cambie las dimensiones de la ventana

    window_2.resizable(0,False) #ancho y alto 0 o 1 también
    
    #Le cambio el color al findo de mi ventana
    
    window_2.configure(background="Pink")
    
    #Ingreso texto en mi ventana
    
    tk.Label(window_2,text="Ingresa los valores solicitados:", bg="Pink", fg="white", font="none 16 bold"). grid(row=1,column=3, sticky="W")
    
    #Ahora genero los cuadros para que se ingresen datos.
    
    Entrada_12= tk.Entry(window_2, width=20, bg="white")
    Entrada_12.grid(row=4, column=4)
    
    Entrada_22= tk.Entry(window_2, width=20, bg="white")
    Entrada_22.grid(row=5, column= 4)
    
    Entrada_32= tk.Entry(window_2, width=20, bg="white")
    Entrada_32.grid(row=6, column= 4)
    

    #Agrego las etiquetas correspondientes a dichos cuadros
    
    Etiqueta_12=tk.Label(window_2,text="Día", bg="Pink", fg="black", font="none 16 bold").grid(row=4, column= 3)
    
    Etiqueta_22=tk.Label(window_2,text="Mes", bg="Pink", fg="black", font="none 16 bold").grid(row=5,column= 3)
    
    Etiqueta_32=tk.Label(window_2,text="Año", bg="Pink", fg="black", font="none 16 bold").grid(row=6, column= 3)
        
    #Etiqueta de espacio
    
    Etiqueta_82=tk.Label(window_2,text=" ", bg="Pink").grid(row=11, column= 3)
    
    Etiqueta_92=tk.Label(window_2,text="          ", bg="Pink").grid(row=0, column= 0)
    
    #Ingreso un botón para confirmar los datos y comenzar el cáculo.
    
    Boton_12=tk.Button(window_2, text="Ver archivo diario", width=15, command= hor_DIA).grid(row=12, column= 4)
    
    Boton_12=tk.Button(window_2, text="Ver diagrama diario", width=15, command= diag_DIA).grid(row=12, column= 3)
    
    #Le indico a mi programa que ya realicé todas las configuraciones relacionadas con mi interfase y ya puede mostrar la ventana. Va a ser el bucle principal por lo que va a mantenerse abierto.
    
    window_2.mainloop()

    
#%%

def crear_arch_mult_ACT(): #Esta es la función que se podrá ejecutar desde esta pantalla

    #Obtengo las variables de mi interfaz gráfica
    
    ACT = Entrada_1.get()
    
    AN = Entrada_2.get()
      
    MES = Entrada_3.get()
    
    DIA = Entrada_4.get()
    
    INICIO = Entrada_5.get()  
    
    FINAL= Entrada_6.get()
    
    OBS = Entrada_7.get()
                
    #Armo una serie donde ingresarán mis valores
      
    notas_4 = pd.DataFrame(columns=('Hora de inicio','Hora de finalización','día','Mes','Año','Actividad', 'Observaciones'))
    
    notas_4.loc[len(notas_4)]=[INICIO, FINAL, DIA, MES, AN, ACT, OBS]
    
    notas_4.to_excel(file_path + "Cronograma Actividades/Mensual/"+str(AN)+"/Agenda "+str(MES)+" "+str(AN)+".xlsx", index=False)
      
    Mensaje_saliente_ventana= tk.Label(window, text = "¡¡Creación de archivo Exitosa!!", bg="Pink", fg="black", font="none 12 bold").grid(row=16,column=5, sticky="W")    

#Fin función crear_arch_mult_ACT()

#%%

def agregar_arch_mult_ACT(): #Esta es la función que se podrá ejecutar desde esta pantalla

    #Obtengo las variables de mi interfaz gráfica
    
    ACT = Entrada_1.get()
    
    AN = Entrada_2.get()
      
    MES = Entrada_3.get()
    
    DIA = Entrada_4.get()
    
    INICIO = Entrada_5.get()  
    
    FINAL= Entrada_6.get()
    
    OBS = Entrada_7.get()
                
    #Abro el archivo donde está mi dataframe
    
    notas_4=pd.read_excel(file_path + "Cronograma Actividades/Mensual/"+str(AN)+"/Agenda "+str(MES)+" "+str(AN)+".xlsx")
    
    notas_4.loc[len(notas_4)]=[INICIO, FINAL, DIA, MES, AN, ACT, OBS]
    
    notas_4.to_excel(file_path + "Cronograma Actividades/Mensual/"+str(AN)+"/Agenda "+str(MES)+" "+str(AN)+'.xlsx', index=False, sheet_name="Actividades")
    
    Mensaje_saliente_ventana= tk.Label(window, text = "¡¡Grabación Exitosa!!", bg="Pink", fg="black", font="none 12 bold").grid(row=16,column=5, sticky="W")    

#Fin función agregar_arch_mult_ACT()

#%%    

#Creación de la interfaz gráfica principal (Ventana principal)
    
#Creo una ventana
    
window = tk.Tk()

#Le aplico un título a mi ventana

window.title("Agenda Diaria")

#Le Agrego el ícono a la ventana

window.iconbitmap("ilustracion.ico") # https://www.online-convert.com/es/

#Determino el tamaño de mi ventana

window.geometry("720x400") #(largo, alto)

#Evito que el usuario me cambie las dimensiones de la ventana

window.resizable(0,False) #ancho y alto (0/Falso o 1/verdadero) esto permite al usuario aumentar o diminuir la ventana.

#Le cambio el color al findo de mi ventana

window.configure(background="Pink")

#Ingreso texto en mi ventana

tk.Label(window,text="Bienvenidos:", bg="Pink", fg="white", font="none 16 bold"). grid(row=1,column=3, sticky="W")

#Ahora genero los cuadros para que se ingresen datos.

Entrada_1= tk.Entry(window, width=20, bg="white")
Entrada_1.grid(row=4, column=5)

Entrada_2= tk.Entry(window, width=20, bg="white")
Entrada_2.grid(row=5, column= 5)

Entrada_3= tk.Entry(window, width=20, bg="white")
Entrada_3.grid(row=6, column= 5)

Entrada_4= tk.Entry(window, width=20, bg="white")
Entrada_4.grid(row=7, column= 5)

Entrada_5= tk.Entry(window, width=20, bg="white")
Entrada_5.grid(row=8, column= 5)

Entrada_6= tk.Entry(window, width=20, bg="white")
Entrada_6.grid(row=9, column= 5)

Entrada_7= tk.Entry(window, width=20, bg="white")
Entrada_7.grid(row=10, column= 5)

#Agrego las etiquetas correspondientes a dichos cuadros

Etiqueta_1=tk.Label(window,text="Actividad", bg="Pink", fg="black", font="none 16 bold").grid(row=4, column= 4)

Etiqueta_2=tk.Label(window,text="Año", bg="Pink", fg="black", font="none 16 bold").grid(row=5,column= 4)

Etiqueta_3=tk.Label(window,text="Mes", bg="Pink", fg="black", font="none 16 bold").grid(row=6, column= 4)

Etiqueta_4=tk.Label(window,text="Día", bg="Pink", fg="black", font="none 16 bold").grid(row=7,column= 4)

Etiqueta_5=tk.Label(window,text="Hora de inicio", bg="Pink", fg="black", font="none 16 bold").grid(row=8, column= 4)

Etiqueta_6=tk.Label(window,text="Hora de finalización", bg="Pink", fg="black", font="none 16 bold").grid(row=9, column= 4)

Etiqueta_7=tk.Label(window,text="Observaciones", bg="Pink", fg="black", font="none 16 bold").grid(row=10,column= 4)

#Agrego una imagen

miImagen=PhotoImage(file="agenda.gif") #sin colocar la ruta cuando esta dentro de la carpeta de trabajo.

Etiqueta_Imagen=tk.Label(window, image=miImagen).place(x=540, y=20) #Pongo la imagen en una etiqueta

#Etiqueta de espacio

Etiqueta_8=tk.Label(window,text=" ", bg="Pink").grid(row=11, column= 3)

Etiqueta_9=tk.Label(window,text=" ", bg="Pink").grid(row=13, column= 4)

#Ingreso un botón para confirmar los datos y comenzar el cáculo.

Boton_1=tk.Button(window, text="Agregar al horario", width=15, command= agregar_arch_mult_ACT).grid(row=12, column= 4)

Boton_2=tk.Button(window, text="Crear archivo mensual", width=20, command= crear_arch_mult_ACT).grid(row=14, column= 4)

# Crear el menu principal

menu_barra = tk.Menu(window) #Indica que crearé un menú

window.config(menu=menu_barra) # Muestra el menú

opciones1 = tk.Menu(menu_barra) #Indica que trabajaré con una cascada (crearé un botón cascada)

menu_barra.add_cascade(label="Ver horarios", menu=opciones1) #Nombre de la primera cascada

#Opciones de la primera cascada

opciones1.add_command(label="Ver horario del día", command=ventana_DIA)

opciones1.add_separator() #Separador

opciones1.add_command(label="Borrar actividad", command=hola)

#Le indico a mi programa que ya realicé todas las configuraciones relacionadas con mi interfase y ya puede mostrar la ventana. Va a ser el bucle principal por lo que va a mantenerse abierto.

window.mainloop()

#Fin

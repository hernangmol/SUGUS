import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import mysql.connector
from datetime import date
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib
import logging
matplotlib.use("TkAgg")
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import BDNP as bd
from datetime import datetime
import datetime as dt
import math
############################################ Ventanas ##########################################
# w1 - Menú principal
# w2 - Inicio de sesión
# w3 - Visualización de pedidos
# w4 - Ingreso de pedido
# w5 - Cambio de contraseña
# w6 - Alta de usuario
# w7 - Formulario de asignación
# w8 - Modificación de documentos y normas 
# w9 - Ver asignaciones a documentos y normas
# WA - Cambio de privilegios
# wB - Ventana de visualización de muestras
# wC - Ver estadísticas
# wD - Agenda de tareas
# wE - Visualización de pedidos asignados a proyectos, desarrollos, asistencias
# wF - Modificación de proyectos
# wG - Modificación de tareas
# wH - Modificación de ordenes de laboratorio
# wI - Modificación de ordenes de servicio
# wJ - Modificación de muestras
# wK - Ventana de ingreso de ordenes internas
# wL - Ventana de gestión de recursos
'''
error 5119 Unread result found
error 5550 Unread result found
error 5561 Unread result found
error 5312 Unread result found
'''
###########################################################################
def main():
    global my_conn
    while True:
        try:
            my_conn = mysql.connector.connect(host = '192.168.100.105',
                                            port = 3306,
                                            #database = "BDNP",  # base de produccion
                                            database = "BDNP_t",  # base de test
                                            user = "admin",
                                            password = "cenadif2023")
            break
            #print("conexion OK!")
        except Exception as e:
            if messagebox.askretrycancel(message="Error de conexión con base de datos.", title="SUGUS - Frontend BDNP"):
                pass
            logging.error(e, exc_info=True)
            exit(1)
            #print("error", e)

    refresh_conn(my_conn)
    ######### ventana de MENU PRINCIPAL ###############################################################################
    global w1
    w1=tk.Tk()
    global photo
    photo = tk.PhotoImage(file = "Recursos/iconoCDF.png")
    w1.withdraw()
    w1.geometry("600x650")
    w1.iconphoto(False, photo)
    w1.title("CENADIF - Base de datos")

    image = tk.PhotoImage(file="Recursos/CENADIF.png")
    image_sub=image.subsample(4)
    # Insertarla en una etiqueta.
    L10 =tk.Label(w1, image=image_sub)
    L10.place(x=8, y=-85)
    ### Botón de estadísticas ###############################
    B1A=tk.Button(w1,text="Estadísticas", width=12, command=V_estad)
    B1A.place(x=450, y=45)

    L11 = tk.Label(w1, text = "Menú principal")
    L11.place(x=265, y=105)
    # Primer piso
    L12 = tk.Label(w1, text = "Solicitudes de trabajo")
    L12.place(x=75, y=135)
    global B11
    B11=tk.Button(w1,text="Ver listado", width=12, command=V_viewRequest)
    B11.place(x=210, y=135)
    B11['state'] = tk.DISABLED
    global B12
    B12=tk.Button(w1,text="Ing. pedido", width=12, command = ingresar)
    B12.place(x=320, y=135)
    B12['state'] = tk.DISABLED
    global B1B
    B1B=tk.Button(w1,text="Ord.Int.Lab", width=12, command = V_oIntLab)
    B1B.place(x=430, y=135)
    B1B['state'] = tk.DISABLED
    # Segundo piso
    L13 = tk.Label(w1, text = "Gestión de trabajos")
    L13.place(x=75, y=185)
    global B13
    B13=tk.Button(w1,text="Asignar", width=12, command=assignment_form)
    B13.place(x=210, y=185)
    B13['state'] = tk.DISABLED
    ### Botón de gestión de recursos ###############################
    global B1C
    B1C=tk.Button(w1,text="Recursos", width=12, command=V_recursos)
    B1C.place(x=320, y=185)
    B1C['state'] = tk.DISABLED
    # Tercer piso ######################
    L14 = tk.Label(w1, text = "Documentación técnica")
    L14.place(x=75, y=235)
    global B14
    B14=tk.Button(w1,text="Ver pedidos", width=12, command=viewDocs)
    B14.place(x=210, y=235)
    B14['state'] = tk.DISABLED
    global B15
    B15=tk.Button(w1,text="Activos", width=12 , command=modDocs)
    B15.place(x=320, y=235)
    B15['state'] = tk.DISABLED
    # Cuarto piso ######################
    L20 = tk.Label(w1, text = "Proyectos")
    L20.place(x=75, y=285)
    global B26
    B26=tk.Button(w1,text="Ver pedidos", width=12, command = lambda: V_viewPro('proy'))
    B26.place(x=210, y=285)
    global B27
    B27=tk.Button(w1,text="Activos", width=12, command = lambda: V_modPro('proy'))
    B27.place(x=320, y=285)
    global B29
    B29=tk.Button(w1,text="Inactivos", width=12, command = lambda: V_modPro('proy_c'))
    B29.place(x=430, y=285)
    # Quinto piso ######################
    L15 = tk.Label(w1, text = "Desarrollos")
    L15.place(x=75, y=335)
    global B16
    B16=tk.Button(w1,text="Ver pedidos", width=12, command = lambda: V_viewPro('des') )
    B16.place(x=210, y=335)
    #B16['state'] = tk.DISABLED
    global B17
    B17=tk.Button(w1,text="Activos", width=12, command = lambda: V_modPro('des'))
    B17.place(x=320, y=335)
    #B17['state'] = tk.DISABLED
    global B30
    B30=tk.Button(w1,text="Inactivos", width=12, command = lambda: V_modPro('des_c'))
    B30.place(x=430, y=335)
    # Sexto piso ######################
    L16 = tk.Label(w1, text = "Asistencia técnica")
    L16.place(x=75, y=385)
    global B18
    B18=tk.Button(w1,text="Ver pedidos", width=12, command = lambda: V_viewPro('ate'))
    B18.place(x=210, y=385)
    #B18['state'] = tk.DISABLED
    global B19
    B19=tk.Button(w1,text="Activas", width=12, command = lambda: V_modPro('ate'))
    B19.place(x=320, y=385)
    #B19['state'] = tk.DISABLED
    global B31
    B31=tk.Button(w1,text="Inactivas", width=12, command = lambda: V_modPro('ate_c'))
    B31.place(x=430, y=385)
    # Septimo piso ######################
    L17 = tk.Label(w1, text = "Laboratorio")
    L17.place(x=75, y=435)
    global B20
    B20=tk.Button(w1,text="Ord. trabajo", width=12, command = lambda: V_modLab('T'))
    B20.place(x=210, y=435)
    B20['state'] = tk.DISABLED
    global B21
    B21=tk.Button(w1,text="Ord. Internas", width=12, command = lambda: V_modLab('I'))
    B21.place(x=320, y=435)
    B21['state'] = tk.DISABLED
    global B34
    B34=tk.Button(w1,text="Ver muestras", width=12, command = V_viewSamples)
    B34.place(x=430, y=435)
    #B34['state'] = tk.DISABLED
    # Octavo piso ######################
    L18 = tk.Label(w1, text = "Admin. sistema")
    L18.place(x=75, y=485)
    global B22
    B22=tk.Button(w1,text="Alta usuario", width=12, command = addUser)
    B22.place(x=210, y=485)
    B22['state'] = tk.DISABLED
    global B23
    B23=tk.Button(w1,text="Permisos", width=12, command = privChange)
    B23.place(x=320, y=485)
    B23['state'] = tk.DISABLED
    # Noveno piso ######################
    L19 = tk.Label(w1, text = "Usuario")
    L19.place(x=75, y=535)
    global B24
    B24=tk.Button(w1,text="Contraseña", width=12, command=passChange)
    B24.place(x=210, y=535)
    global B25
    B25=tk.Button(w1,text="Agenda", width=12, command = V_accSchedule)
    B25.place(x=320, y=535)
    #B25['state'] = tk.DISABLED
    global B28
    B28=tk.Button(w1,text="Tareas", width=12, command = lambda: V_modTar('user'))
    B28.place(x=430, y=535)

    # boton salir
    Bsalir=tk.Button(w1,text="Salir", width=12, command = lambda: salirConAviso())
    Bsalir.place(x=460, y=585)

    accessForm() 
    w1.mainloop()

#################### Ventana listado de pedidos #########################################################
def V_viewRequest(geom = ''):
    w1.withdraw()
    global w3
    w3=tk.Toplevel()
    if len(geom)==0:
        w3.state('zoomed')
    else:
        w3.geometry(geom)
    w3.iconphoto(False, photo)
    w3.title("Pedidos de trabajo")
    w3.frame3 = tk.Frame(w3)
    w3.frame3.grid(rowspan=2, column=1, row=1)
    w3.tabla = ttk.Treeview(w3.frame3, height=35)
    w3.tabla.tag_configure('abiertos', background="red", foreground="white") 
    w3.tabla.grid(column=1, row=1)
    #w3.tabla.place(x=500, y=500)
    ladox = tk.Scrollbar(w3.frame3, orient = tk.VERTICAL, command= w3.tabla.yview)
    ladox.grid(column=0, row = 1, sticky='ew') 
    ladoy = tk.Scrollbar(w3.frame3, orient = tk.HORIZONTAL, command = w3.tabla.yview)
    ladoy.grid(column = 1, row = 0, sticky='ns')

    w3.tabla.configure(xscrollcommand = ladox.set, yscrollcommand = ladoy.set)
    w3.tabla['columns'] = ('nombre', 'fecha solicitud', 'descripcion', 'solicitante','contacto', 'correo', 'telefono')
    w3.tabla.column('#0', minwidth=50, width=50, anchor='center')
    #w3.tabla.column('id_solicitud', minwidth=100, width=120, anchor='center')
    w3.tabla.column('nombre', minwidth=100, width=300 , anchor='w')
    w3.tabla.column('fecha solicitud', minwidth=100, width=120, anchor='center' )
    w3.tabla.column('descripcion', minwidth=300, width=400 , anchor='w')
    w3.tabla.column('solicitante', minwidth=100, width=150, anchor='center')
    w3.tabla.column('contacto', minwidth=120, width=140, anchor='center')
    w3.tabla.column('correo', minwidth=150, width=190, anchor='center')
    w3.tabla.column('telefono', minwidth=100, width=140, anchor='center')

    w3.tabla.heading('#0', text='ID', anchor ='e')
    #w3.tabla.heading('id_solicitud', text='ID', anchor ='center')
    w3.tabla.heading('nombre', text='Nombre', anchor ='center')
    w3.tabla.heading('fecha solicitud', text='Fecha solicitud', anchor ='center')
    w3.tabla.heading('descripcion', text='Descripción', anchor ='center')
    w3.tabla.heading('solicitante', text='Area solicitante', anchor ='center')
    w3.tabla.heading('contacto', text='Contacto', anchor ='center')
    w3.tabla.heading('correo', text='Correo', anchor ='center')
    w3.tabla.heading('telefono', text='Teléfono', anchor ='center')

    # boton Salir
    B31=tk.Button(w3,text="Salir", width=12, command = lambda: salirConAviso())
    B31.place(x=850, y=760)

    # boton Actualizar
    B32=tk.Button(w3,text="Actualizar", width=12, command = actualizar)
    B32.place(x=735, y=760)

    # boton Volver
    B33=tk.Button(w3,text="Volver", width=12, command = lambda: menuFrom(w3))
    B33.place(x=620, y=760)

    w3.tabla.delete(*w3.tabla.get_children())

    refresh_conn(my_conn)
    try:
            my_cursor = my_conn.cursor()
            statement = "SELECT * FROM solicitudes ORDER BY ID_solicitud DESC LIMIT 55"
            my_cursor.execute(statement)
            resultados = my_cursor.fetchall()
    except Exception as e:
        print("error", e)
    
    for fila in resultados:
        if fila[7] is None:
            w3.tabla.insert('',index = 'end', iid=None, text = str(fila[0]), values = [fila[2], fila[1], fila[3], fila[5], fila[4], fila[11], fila[12]],tags='abiertos')
        else:
            w3.tabla.insert('',index = 'end', iid=None, text = str(fila[0]), values = [fila[2], fila[1], fila[3], fila[5], fila[4], fila[11], fila[12]])
           
    w3.after(1, lambda: w3.focus_force())
##############################################################################################################################
def ingresar():
    w1.withdraw()
    #print("Ingresando...")
    global w4
    w4=tk.Toplevel()
    w4.iconphoto(False, photo)
    w4.geometry("500x650")
    w4.title("CENADIF - Base de datos")
    L41 = tk.Label(w4, text = "Ingreso de pedido")
    L41.place(x=210, y=25)
    
    # Fecha de ingreso
    L42 = tk.Label(w4, text = "Fecha de ingreso: ")
    L42.place(x=25, y=75)
    global E41
    E41 = tk.Entry(w4)
    E41.place(x=125, y=75)
    E41.insert(0, date.today()) 
    L46 = tk.Label(w4, text = "(Formato: AAAA-MM-DD)")
    L46.place(x=275, y=75)
    
    # Cliente
    L43 = tk.Label(w4, text = "Nombre cliente: ")
    L43.place(x=25, y=125)
    global E42
    E42 = tk.Entry(w4,width = 25)
    E42.place(x=125, y=125)

    
    # Filiación
    L44 = tk.Label(w4, text = "Área cliente: ")
    L44.place(x=25, y=175)
    global E43
    E43 = tk.Entry(w4,width = 25)
    E43.place(x=125, y=175)
    
    # Correo
    L48 = tk.Label(w4, text = "Correo: ")
    L48.place(x=25, y=225)
    global E48
    E48 = tk.Entry(w4, width = 25)
    E48.place(x=125, y=225)

    # Teléfono
    L49 = tk.Label(w4, text = "Teléfono: ")
    L49.place(x=25, y=275)
    global E49
    E49 = tk.Entry(w4, width = 25)
    E49.place(x=125, y=275)

    # Asunto
    L45 = tk.Label(w4, text = "Asunto: ")
    L45.place(x=25, y=325)
    global E44
    E44 = tk.Entry(w4, width = 56)
    E44.place(x=125, y=325)

    # Descripción
    L45 = tk.Label(w4, text = "Descripción: ")
    L45.place(x=25, y=375)
    global T41
    T41 = tk.Text(w4, width = 42, height = 8)
    T41.place(x=125, y=375)

     # Ingresó
    L47 = tk.Label(w4, text = "Ingresó: ")
    L47.place(x=25, y=530)
    
    list = userList()
    try:
        act_user_index = list.index((actualUser,))
    except:
        act_user_index = 0
    #print(act_user_index)
    global C45
    C45 = ttk.Combobox(w4, state="readonly", width = 17, values = list)
    C45.place(x=125, y=530)
    C45.current(act_user_index) # setea por default el usuario actual en el combobox 
    # boton Volver
    B41=tk.Button(w4,text="Volver", command = lambda: menuFrom(w4))
    B41.place(x=125, y=580)

    # boton Ingresar
    B42=tk.Button(w4,text="Ingresar", command=ingresar_pedido)
    B42.place(x=400, y=580)

    w4.after(1, lambda: w4.focus_force())
######################################################################################################################
def assignment_form():
    w1.withdraw()
    global w7
    w7=tk.Toplevel()
    w7.geometry("1000x700")
    w7.iconphoto(False, photo)
    w7.title("Pedidos de trabajo - Asignación")
    w7.frame7 = tk.Frame(w7)
    w7.frame7.grid(rowspan=2, column=1, row=1)
    w7.tabla = ttk.Treeview(w7.frame7, height=15)
    w7.tabla.grid(column=1, row=1)
    ladox = tk.Scrollbar(w7.frame7, orient = tk.VERTICAL, command= w7.tabla.yview)
    ladox.grid(column=0, row = 1, sticky='ew') 
    ladoy = tk.Scrollbar(w7.frame7, orient = tk.HORIZONTAL, command = w7.tabla.yview)
    ladoy.grid(column = 1, row = 0, sticky='ns')
    w7.tabla.configure(xscrollcommand = ladox.set, yscrollcommand = ladoy.set)
    w7.tabla['columns'] = ('nombre', 'fecha solicitud', 'descripcion', 'solicitante','contacto')
    w7.tabla.column('#0', minwidth=50, width=50, anchor='center')
    w7.tabla.column('nombre', minwidth=100, width=210 , anchor='w')
    w7.tabla.column('fecha solicitud', minwidth=100, width=120, anchor='center' )
    w7.tabla.column('descripcion', minwidth=100, width=250 , anchor='w')
    w7.tabla.column('solicitante', minwidth=100, width=150, anchor='center')
    w7.tabla.column('contacto', minwidth=100, width=150, anchor='center')
    w7.tabla.heading('#0', text='ID', anchor ='e')
    w7.tabla.heading('nombre', text='Nombre', anchor ='center')
    w7.tabla.heading('fecha solicitud', text='Fecha solicitud', anchor ='center')
    w7.tabla.heading('descripcion', text='Descripción', anchor ='center')
    w7.tabla.heading('solicitante', text='Area solicitante', anchor ='center')
    w7.tabla.heading('contacto', text='Contacto', anchor ='center')

# Tipo de salida
    L72 = tk.Label(w7, text = "Tipo de salida *")
    L72.place(x=25, y=380)

# Tipo de salida (combobox)
    # registering the observer
    my_var = tk.StringVar()
    my_var.trace('w', lambda *_ , var = my_var : traceCB(*_,var=var))
    global C70
    C70 = ttk.Combobox(w7, state="readonly", textvariable = my_var, width = 17, values = ('Documentacion', 'Proyectos', 'Desarrollos', 'Asistencias', 'Laboratorio', 'Cierre'))
    C70.place(x=143, y=380) # menos 8?
    
# Fecha de asignación  
    L73 = tk.Label(w7, text = "Fecha de asignación: * ")
    L73.place(x=25, y=430)
    global E73
    E73 = tk.Entry(w7)
    E73.place(x=145, y=430)
    E73.insert(0, date.today())
    L74 = tk.Label(w7, text = "(Formato: AAAA-MM-DD)")
    L74.place(x=295, y=430)

# Responsable (solo para Proyectos)
    
    global L76
    L76 = tk.Label(w7, text = "Responsable: * ")

    #global E74
    #E74 = tk.Entry(w7)
    list = userList()
    global C74
    C74 = ttk.Combobox(w7, state="readonly", width = 17, values = list)

# Descripción

    L75 = tk.Label(w7, text = "Descripción: ")
    L75.place(x=25, y=480)
    global T71
    T71 = tk.Text(w7, width = 42, height = 8)
    T71.place(x=145, y=480)

# boton Salir
    B71=tk.Button(w7,text="Salir", width=12,  command = lambda: salirConAviso())
    B71.place(x=510, y=530)

# boton Volver
    B71=tk.Button(w7,text="Volver", width=12,  command = lambda: menuFrom(w7))
    B71.place(x=510, y=480)

# boton Asignar
    B72=tk.Button(w7,text="Asignar", width=12, command=B_assign)
    B72.place(x=510, y=580)

    refresh_conn(my_conn)
    try:
        my_cursor = my_conn.cursor()
        statement = "SELECT * FROM solicitudes WHERE tipo_salida IS NULL "
        my_cursor.execute(statement, )
        resultados = my_cursor.fetchall()
#       print(resultados) 
    except Exception as e:
        print("error", e)
    
    for fila in resultados:
        w7.tabla.insert('',index = fila[0], iid=None, text = str(fila[0]), values = [fila[2], fila[1], fila[3], fila[5], fila[4]])

    w7.after(1, lambda: w7.focus_force())

# defining the callback function (observer)
def traceCB(*args, var):
    tipoSal = var.get()
    if tipoSal == 'Proyectos':
        L76.place(x=295, y=380)
        C74.place(x=395, y=380)
    else:
        L76.place_forget()
        C74.place_forget()

#################### Ver pedidos asignados a doc ################################################
def viewDocs():
    w1.withdraw()
    global w9
    w9=tk.Toplevel() 
    #style = ttk.Style(w9)
    #style.theme_use('clam')
    w9.state("zoomed")
    #w9.geometry("1100x650")
    w9.iconphoto(False, photo)
    w9.title("Pedidos de trabajo asignados a DyN")
    w9.frame = tk.Frame(w9)
    w9.frame.grid(rowspan=3, column=2, row=1, pady = 2)
    #w9.frame.place(x=20, y=20)
    w9.tabla = ttk.Treeview(w9.frame, height=32)
    w9.tabla.grid(column=2, row=1)
    #w9.tabla.place(x=50, y=50)
    ladox = tk.Scrollbar(w9.frame, orient = tk.VERTICAL, command= w9.tabla.yview)
    ladox.grid(column=3, row = 1, sticky='ew') 
    ladoy = tk.Scrollbar(w9.frame, orient =tk.HORIZONTAL, command = w9.tabla.xview)
    ladoy.grid(column = 2, row = 0, sticky='ns')

    w9.tabla.configure(xscrollcommand = ladox.set, yscrollcommand = ladoy.set)
       
    w9.tabla['columns'] = ('id_doc','nombre', 'fecha solicitud', 'descripcion', 'solicitante','contacto')

    w9.tabla.column('#0', minwidth=50, width=50, anchor='center')
    w9.tabla.column('id_doc', minwidth=50, width=70, anchor='center')
    w9.tabla.column('nombre', minwidth=150, width=300 , anchor='w')
    w9.tabla.column('fecha solicitud', minwidth=100, width=120, anchor='center' )
    w9.tabla.column('descripcion', minwidth=100, width=400 , anchor='w')
    w9.tabla.column('solicitante', minwidth=100, width=300, anchor='center')
    w9.tabla.column('contacto', minwidth=100, width=150, anchor='center')

    w9.tabla.heading('#0', text='Pedido', anchor ='e')
    w9.tabla.heading('id_doc', text='N° DyN', anchor ='center')
    w9.tabla.heading('nombre', text='Nombre', anchor ='center')
    w9.tabla.heading('fecha solicitud', text='Fecha solicitud', anchor ='center')
    w9.tabla.heading('descripcion', text='Descripción', anchor ='center')
    w9.tabla.heading('solicitante', text='Area solicitante', anchor ='center')
    w9.tabla.heading('contacto', text='Contacto', anchor ='center')

    # boton Salir
    B91=tk.Button(w9,text="Salir", width=12, command = lambda: salirConAviso())
    B91.place(x=850, y=730)

    # boton Volver
    B93=tk.Button(w9,text="Volver", width=12, command = lambda: menuFrom(w9))
    B93.place(x=620, y=730)

    w9.tabla.delete(*w9.tabla.get_children())

    refresh_conn(my_conn)
    try:
            my_cursor = my_conn.cursor()
            statement = "SELECT * FROM solicitudes WHERE tipo_salida = 'documentacion' "
            my_cursor.execute(statement)
            resultados = my_cursor.fetchall()
#            print(resultados) 
    except Exception as e:
        print("error", e)
    
    refresh_conn(my_conn)
    for fila in resultados:
        #w9.tabla.insert('',index = fila[0], iid=None, text = str(fila[0]), values = [fila[10], fila[2], fila[1], fila[3], fila[5], fila[4]])
        try:
            my_cursor = my_conn.cursor()
            statement = "SELECT ID_trabajo FROM documentacion WHERE ID_solicitud = %s" 
            values = (fila[0],)
            my_cursor.execute(statement, values)
            aux = my_cursor.fetchall()
            #print(aux) 
        except Exception as e:
            print("error", e)
        w9.tabla.insert('',index = fila[0], iid=None, text = str(fila[0]), values = [aux[0], fila[2], fila[1], fila[9], fila[5], fila[4]])    

    w9.after(1, lambda: w9.focus_force())

#################### Modificar Documentación #####################################################
def modDocs(geom = ''):
    w1.withdraw()
    global w8
    w8=tk.Toplevel()
    #w8.state("zoomed")
    if len(geom)==0:
        w8.state('zoomed')
    else:
        w8.geometry(geom)
    w8.iconphoto(False, photo)
    w8.title("CENADIF Base de datos - Documentación y Normas - Trabajos pendientes")
    w8.frame8 = tk.Frame(w8)
    w8.frame8.grid(rowspan=2, column=1, row=1)
    w8.tabla = ttk.Treeview(w8.frame8, height=15)
    w8.tabla.grid(column=1, row=1)
    #w8.tabla.place(x=500, y=500)
    ladox = tk.Scrollbar(w8.frame8, orient = tk.VERTICAL, command= w8.tabla.yview)
    ladox.grid(column=0, row = 1, sticky='ew') 
    ladoy = tk.Scrollbar(w8.frame8, orient =tk.HORIZONTAL, command = w8.tabla.yview)
    ladoy.grid(column = 1, row = 0, sticky='ns')

    w8.tabla.configure(xscrollcommand = ladox.set, yscrollcommand = ladoy.set)
       
    w8.tabla['columns'] = ('pedido','tipo_doc', 'emisor', 'nro_doc', 'titulo', 'fuente','precio', 'u_monet','cantidad')

    w8.tabla.column('#0', minwidth=50, width=50, anchor='center')
    w8.tabla.column('pedido', minwidth=100, width=200 , anchor='w')
    w8.tabla.column('tipo_doc', minwidth=100, width=100 , anchor='w')
    w8.tabla.column('emisor', minwidth=100, width=100 , anchor='w')
    w8.tabla.column('nro_doc', minwidth=100, width=120, anchor='center' )
    w8.tabla.column('titulo', minwidth=100, width=350 , anchor='w')
    w8.tabla.column('fuente', minwidth=100, width=100, anchor='center')
    w8.tabla.column('precio', minwidth=100, width=100, anchor='center')
    w8.tabla.column('u_monet', minwidth=100, width=100, anchor='center')
    w8.tabla.column('cantidad', minwidth=100, width=100, anchor='center')

    w8.tabla.heading('#0', text='N° DyN', anchor ='e')
    w8.tabla.heading('pedido', text='Pedido', anchor ='center')
    w8.tabla.heading('tipo_doc', text='Tipo', anchor ='center')
    w8.tabla.heading('emisor', text='Emisor', anchor ='center')
    w8.tabla.heading('nro_doc', text='Número', anchor ='center')
    w8.tabla.heading('titulo', text='Título', anchor ='center')
    w8.tabla.heading('fuente', text='Fuente', anchor ='center')
    w8.tabla.heading('precio', text='Precio', anchor ='center')
    w8.tabla.heading('u_monet', text='Un. Mon.', anchor ='center')
    w8.tabla.heading('cantidad', text='Cant.', anchor ='center')

# N° DyN
    L81 = tk.Label(w8, text = "N° DyN")
    L81.place(x=75, y=380)

    global E81
    E81 = tk.Entry(w8)
    E81.place(x=195, y=380)

# Tipo de documento
    L82 = tk.Label(w8, text = "Tipo de documento")
    # L82.place(x=25, y=430)
    L82.place(x=400, y=380)

    #global E82
    #E82 = tk.Entry(w8)
    #E82.place(x=525, y=380)

    global C81
    C81 = ttk.Combobox(w8, state="write", values = ['Norma', 'Plano', 'Esp. Técnica', 'Inst. Técnico', 'Informe', 'Manual'])
    C81.place(x=525, y=380)

# N° de documento  
    L83 = tk.Label(w8, text = "N° de documento ")
    L83.place(x=1050, y=380)

    global E83
    E83 = tk.Entry(w8)
    E83.place(x=1175, y=380)

# Título
    L84 = tk.Label(w8, text = "Título")
    #L84.place(x=350, y=380)
    L84.place(x=400, y=430)

    global E84
    E84 = tk.Entry(w8)
    #E84.place(x=475, y=380, width=450)
    E84.place(x=525, y=430, width=450)

# Fuente
    L85 = tk.Label(w8, text = "Fuente")
    #L85.place(x=350, y=430)
    L85.place(x=75, y=480)
    global C82
    C82 = ttk.Combobox(w8, state="write", values = ['IRAM colección', 'AENOR MAS', 'UIC', 'Arch. Gral. Ferrov.', 'Base de datos', 'Internet'])
    C82.place(x=195, y=480)

# Emisor  
    L87 = tk.Label(w8, text = "Emisor ")
    #L87.place(x=350, y=480)
    L87.place(x=725, y=380)
    global E87
    E87 = tk.Entry(w8)
    #E87.place(x=475, y=480)
    E87.place(x=850, y=380)

# Fecha de resolución
    L88 = tk.Label(w8, text = "Fecha de Resolución")
    L88.place(x=1050, y=480)
    global E88
    E88 = tk.Entry(w8)
    E88.place(x=1175, y=480)
    L84 = tk.Label(w8, text = "(Formato: AAAA-MM-DD)")
    L84.place(x=1165, y=505)

# Cantidad
    L89 = tk.Label(w8, text = "Cantidad")
    L89.place(x=75, y=430)
    global E89
    E89 = tk.Entry(w8)
    E89.place(x=195, y=430)

# Respondió  
    L8A = tk.Label(w8, text = "Respondió ")
    L8A.place(x=1050, y=430)
    global E8A
    E8A = tk.Entry(w8)
    E8A.place(x=1175, y=430)
    E8A.insert(0, actualUser)

# Precio
    L8D = tk.Label(w8, text = "Precio")
    # L8D.place(x=675, y=430)
    L8D.place(x=400, y=480)

    global E8D
    E8D = tk.Entry(w8)
    #E8D.place(x=800, y=430)
    E8D.place(x=525, y=480)

# Unidad monetaria  
    L8E = tk.Label(w8, text = "Unidad monetaria")
    L8E.place(x=725, y=480)

    #global E8E
    #E8E = tk.Entry(w8)
    #E8E.place(x=850, y=480)

    global C83
    C83 = ttk.Combobox(w8, state="write", values = ['ARS', 'USD', 'EUR'])
    C83.place(x=850, y=480)


# Descripción

    L8B = tk.Label(w8, text = "Observaciones: ")
    L8B.place(x=75, y=530)

    global T81
    T81 = tk.Text(w8, width = 81, height = 8)
    T81.place(x=195, y=530)

# Botones
    # boton Precarga
    B84=tk.Button(w8,text="Precarga", width=12, command = precDoc)
    B84.place(x=865, y=585)

    # boton Recursos
    B85=tk.Button(w8, text="Recursos", width=12, command = lambda: V_empleos('doc'))
    B85.place(x=1000, y=585)

    # boton Modificar
    B82=tk.Button(w8,text="Modificar", width=12, command = B_modDoc)
    B82.place(x=1135, y=585)

    # boton Volver####
    B83=tk.Button(w8,text="Volver", width=12, command = lambda: menuFrom(w8))
    B83.place(x=1000, y=635)

    # boton Salir###
    B81=tk.Button(w8,text="Salir", width=12, command = lambda: salirConAviso())
    B81.place(x=1135, y=635)

    w8.tabla.delete(*w8.tabla.get_children())

    refresh_conn(my_conn)
    try:
        my_cursor = my_conn.cursor()
        statement = "SELECT * FROM documentacion WHERE fecha_resolucion IS NULL"
        my_cursor.execute(statement)
        resultados = my_cursor.fetchall()
        #print(resultados) 
    except Exception as e:
        print("error", e)
    

    for fila in resultados:

        try:
            my_cursor = my_conn.cursor()
            statement = "SELECT asunto FROM solicitudes WHERE ID_solicitud = %s"
            values = (fila[1],)
            my_cursor.execute(statement, values)
            pedido = my_cursor.fetchone()
            #print(pedido) 
        except Exception as e:
            print("error", e)

        w8.tabla.insert('',index = fila[0], iid=None, text = str(fila[0]), values = [pedido[0], fila[2], fila[8], fila[3], fila[4], fila[5], fila[6], fila[7], fila[11]])

    try:
        my_cursor = my_conn.cursor()
        statement = "SELECT * FROM documentacion WHERE fecha_resolucion = 0"
#        values = (fila[1],)
        my_cursor.execute(statement)
        resultados = my_cursor.fetchall()
#        #print(resultados) 
    except Exception as e:
        print("error", e)

    for fila in resultados:

        try:
            my_cursor = my_conn.cursor()
            statement = "SELECT asunto FROM solicitudes WHERE ID_solicitud = %s"
            values = (fila[1],)
            my_cursor.execute(statement, values)
            pedido = my_cursor.fetchone()
#            print(pedido) 
        except Exception as e:
            print("error", e)

        w8.tabla.insert('',index = fila[0], iid=None, text = str(fila[0]), values = [pedido[0], fila[2], fila[8], fila[3], fila[4], fila[5], fila[6], fila[7], fila[11]])

    w8.after(1, lambda: w8.focus_force())

############# ventana de visualización de proyectos ##########################################
def V_viewPro(modo):
    # Parte común a todos los modos
    w1.withdraw()
    global wE
    wE=tk.Toplevel()
    wE.geometry("1240x650")
    wE.iconphoto(False, photo)
    # Titulo de la ventana (dependiente de modo)
    if modo == 'proy':
        wE.title("Pedidos de trabajo asignados a Proyectos")
    if modo == 'des':
        wE.title("Pedidos de trabajo asignados a Desarrollos")
    if modo == 'ate':
        wE.title("Pedidos de trabajo asignados a Asistencias técnicas")
    # Parte común a todos los modos
    wE.frame = tk.Frame(wE)
    wE.frame.grid(rowspan=2, column=1, row=1)
    wE.tabla = ttk.Treeview(wE.frame, height=25)
    wE.tabla.grid(column=1, row=1)
    ladox = tk.Scrollbar(wE.frame, orient = tk.VERTICAL, command= wE.tabla.yview)
    ladox.grid(column=0, row = 1, sticky='ew') 
    ladoy = tk.Scrollbar(wE.frame, orient =tk.HORIZONTAL, command = wE.tabla.yview)
    ladoy.grid(column = 1, row = 0, sticky='ns')

    wE.tabla.configure(xscrollcommand = ladox.set, yscrollcommand = ladoy.set)
       
    wE.tabla['columns'] = ('num_proy','nombre', 'fecha solicitud', 'descripcion', 'solicitante','contacto', 'responsable')

    wE.tabla.column('#0', minwidth=50, width=50, anchor='center')
    wE.tabla.column('num_proy', minwidth=100, width=120, anchor='center')
    wE.tabla.column('nombre', minwidth=100, width=210 , anchor='w')
    wE.tabla.column('fecha solicitud', minwidth=100, width=120, anchor='center' )
    wE.tabla.column('descripcion', minwidth=100, width=250 , anchor='w')
    wE.tabla.column('solicitante', minwidth=100, width=150, anchor='center')
    wE.tabla.column('contacto', minwidth=100, width=150, anchor='center')
    wE.tabla.column('responsable', minwidth=100, width=150, anchor='center')

    wE.tabla.heading('#0', text='Pedido', anchor ='e')
    wE.tabla.heading('num_proy', text='N° PROY', anchor ='center')
    wE.tabla.heading('nombre', text='Nombre', anchor ='center')
    wE.tabla.heading('fecha solicitud', text='Fecha solicitud', anchor ='center')
    wE.tabla.heading('descripcion', text='Descripción', anchor ='center')
    wE.tabla.heading('solicitante', text='Area solicitante', anchor ='center')
    wE.tabla.heading('contacto', text='Contacto', anchor ='center')
    wE.tabla.heading('responsable', text='Responsable', anchor ='center')

    # boton Salir
    BE1=tk.Button(wE,text="Salir", width=12, command = lambda: salirConAviso())
    BE1.place(x=850, y=560)

    # boton Volver
    BE3=tk.Button(wE,text="Volver", width=12, command = lambda: menuFrom(wE))
    BE3.place(x=620, y=560)

    wE.tabla.delete(*wE.tabla.get_children())

    # Contenido de la tabla (dependiente de modo)
    refresh_conn(my_conn)
    try:
        my_cursor = my_conn.cursor()
        if modo == 'proy':
            statement = "SELECT * FROM solicitudes WHERE tipo_salida = 'Proyectos' "
        if modo == 'des':
            statement = "SELECT * FROM solicitudes WHERE tipo_salida = 'Desarrollos' "
        if modo == 'ate':
            statement = "SELECT * FROM solicitudes WHERE tipo_salida = 'Asistencias' "        
        my_cursor.execute(statement)
        resultados = my_cursor.fetchall()
        #print(resultados) 
    except Exception as e:
        print("error", e)

    # Parte común a todos los modos
    for fila in resultados:
        #w9.tabla.insert('',index = fila[0], iid=None, text = str(fila[0]), values = [fila[10], fila[2], fila[1], fila[3], fila[5], fila[4]])
        try:
            my_cursor = my_conn.cursor()
            statement = "SELECT numero_cdf, responsable FROM proyectos WHERE ID_solicitud = %s" 
            values = (fila[0],)
            #print(values)
            my_cursor.execute(statement, values)
            aux = my_cursor.fetchone()
            #print(aux) 
        except Exception as e:
            print("error 1083", e)
        wE.tabla.insert('',index = fila[0], iid=None, text = str(fila[0]), values = [aux[0], fila[2], fila[1], fila[9], fila[5], fila[4], aux[1]])    

    wE.after(1, lambda: wE.focus_force())

############# Ventana de modificación de proyectos ##########################################
def V_modPro(modo, geom =''):
    # Parte común a todos los modos
    w1.withdraw()
    global wF
    wF=tk.Toplevel()
    if len(geom)==0:
        wF.state('zoomed')
    else:
        wF.geometry(geom)
    wF.iconphoto(False, photo)
    # Titulo de la ventana (dependiente de modo)
    if modo == 'proy':
        wF.title("CENADIF Base de datos - Proyectos en curso")
    if modo == 'proy_c':
        wF.title("CENADIF Base de datos - Proyectos cerrados")    
    if modo == 'des':
        wF.title("CENADIF Base de datos - Desarrollos en curso")
    if modo == 'des_c':
        wF.title("CENADIF Base de datos - Desarrollos cerrados")
    if modo == 'ate':
        wF.title("CENADIF Base de datos - Asistencias técnicas en curso")
    if modo == 'ate_c':
        wF.title("CENADIF Base de datos - Asistencias técnicas cerradas")
    # Parte común a todos los modos
    wF.frame = tk.Frame(wF)
    wF.frame.grid(rowspan=2, column=1, row=1)
    wF.tabla = ttk.Treeview(wF.frame, height=15)
    wF.tabla.grid(column=1, row=1)

    ladox = tk.Scrollbar(wF.frame, orient = tk.VERTICAL, command= wF.tabla.yview)
    ladox.grid(column=0, row = 1, sticky='ew') 
    ladoy = tk.Scrollbar(wF.frame, orient =tk.HORIZONTAL, command = wF.tabla.xview)
    ladoy.grid(column = 1, row = 0, sticky='ns')
    wF.tabla.configure(xscrollcommand = ladox.set, yscrollcommand = ladoy.set)
       
    wF.tabla['columns'] = ('nombre_proy','estado', 'responsable', 'descripción', 'alcance', 'num_pedido','pedido')
    wF.tabla.column('#0', minwidth=50, width=60, anchor='center')
    wF.tabla.column('nombre_proy', minwidth=100, width=350, anchor='center')
    wF.tabla.column('estado', minwidth=80, width=100 , anchor='center')
    wF.tabla.column('responsable', minwidth=100, width=100 , anchor='center')
    wF.tabla.column('descripción', minwidth=100, width=400, anchor='center' )
    wF.tabla.column('alcance', minwidth=100, width=150 , anchor='center')
    wF.tabla.column('num_pedido', minwidth=100, width=100, anchor='center')
    wF.tabla.column('pedido', minwidth=90, width=200, anchor='center')
    # Titulo del número CENADIF y nombre (dependientes del modo)
    if modo == 'proy' or modo == 'proy_c':
        wF.tabla.heading('#0', text='N° PROY', anchor ='e')
        wF.tabla.heading('nombre_proy', text='Nombre del proyecto', anchor ='center')
    if modo == 'des' or modo == 'des_c':
        wF.tabla.heading('#0', text='N° DES', anchor ='e')
        wF.tabla.heading('nombre_proy', text='Nombre del desarrollo', anchor ='center')
    if modo == 'ate' or modo == 'ate_c':
        wF.tabla.heading('#0', text='N° ATE', anchor ='e')
        wF.tabla.heading('nombre_proy', text='Nombre de la asistencia', anchor ='center')
    # Parte común a todos los modos
    wF.tabla.heading('estado', text='Estado', anchor ='center')
    wF.tabla.heading('responsable', text='Responsable', anchor ='center')
    wF.tabla.heading('descripción', text='Descripción', anchor ='center')
    wF.tabla.heading('alcance', text='Alcance', anchor ='center')
    wF.tabla.heading('num_pedido', text='Num pedido', anchor ='center')
    wF.tabla.heading('pedido', text='Pedido', anchor ='center')

# N° PROY
    # Titulo del número CENADIF y nombre (dependientes del modo)
    if modo == 'proy' or modo == 'proy_c':
        LF1 = tk.Label(wF, text = "N° PROY")
    if modo == 'des' or modo == 'des_c':
        LF1 = tk.Label(wF, text = "N° DES")
    if modo == 'ate' or modo == 'ate_c':
        LF1 = tk.Label(wF, text = "N° ATE")
    # Parte común a todos los modos
    LF1.place(x=50, y=380)

    global EF1
    EF1 = tk.Entry(wF)
    EF1.place(x=175, y=380)

# Responsable
    LFD = tk.Label(wF, text = "Responsable")
    LFD.place(x=350, y=380)

#    global EFB
#    EFB = tk.Entry(wF)
#    EFB.place(x=475, y=380)

    list = userList()
    global CFB
    CFB = ttk.Combobox(wF, state="readonly", width = 17, values = list)
    CFB.place(x=475, y=380)
    
# Nombre del proyecto
    # Nombre del... (dependiente del modo)
    if modo == 'proy' or modo == 'proy_c':
        LF4 = tk.Label(wF, text = "Nombre del proyecto")
    if modo == 'des' or modo == 'des_c':
        LF4 = tk.Label(wF, text = "Nombre del desarrollo")
    if modo == 'ate' or modo == 'ate_c':
        LF4 = tk.Label(wF, text = "Nombre de asistencia")
    # Parte común a todos los modos
    LF4.place(x=50, y=430)

    global EF4
    EF4 = tk.Entry(wF)
    EF4.place(x=175, y=430, width=425)

# Estado
    LF5 = tk.Label(wF, text = "Estado")
    LF5.place(x=50, y=480)
    
    global CF9
    CF9 = ttk.Combobox(wF, state="readonly", width = 17, values = ['En curso', 'En pausa', 'Cancelado', 'Finalizado'])
    CF9.place(x=175, y=480)

# Alcance
    LF8 = tk.Label(wF, text = "Alcance")
    LF8.place(x=350, y=480)

    global CF8
    CF8 = ttk.Combobox(wF, width = 17, values = ['Ingeniería', 'Planos', 'Prototipo', 'Prototipo homologado', 'Proveedores', 'Informe', 'Instructivo técnico', 'Especificación técnica', 'Producto homologado', 'Presentación'])
    CF8.place(x=475, y=480)

# Descripción

    LFB = tk.Label(wF, text = "Descripción: ")
    LFB.place(x=50, y=530)

    global TF1
    if modo == 'des':
        TF1 = tk.Text(wF, width = 53, height = 3)
    else:
        TF1 = tk.Text(wF, width = 53, height = 7)
    TF1.place(x=175, y=530)

    if modo == 'des':
    # Prioridad
        LFC = tk.Label(wF, text = "Prioridad")
        LFC.place(x=50, y=610)

        #global EFC
        #EFC = tk.Entry(wF)
        #EFC.place(x=175, y=610)

        global CFC
        CFC = ttk.Combobox(wF, state="readonly", width = 17, values = ['Alta', 'Normal', 'Baja'])
        CFC.place(x=175, y=610)

    # Clasificación
        LFD = tk.Label(wF, text = "Clasificación")
        LFD.place(x=350, y=610)

        global EFD
        EFD = tk.Entry(wF)
        EFD.place(x=475, y=610)

    # NUM
        LFE = tk.Label(wF, text = "NUM:")
        LFE.place(x=50, y=650)

        global EFE
        EFE = tk.Entry(wF)
        EFE.place(x=175, y=650)

#   BOTONES #####################
    # boton Precarga
    BF4=tk.Button(wF, text="Seleccionar", width=12, command = lambda: B_selProy(modo))
    BF4.place(x=175, y=680)

    # boton Modificar
    global BF2
    BF2=tk.Button(wF,text="Modificar", width=12, state = tk.DISABLED, command = lambda: B_modProy(modo))
    BF2.place(x=340, y=680)
    
    # boton Tareas
    BF3=tk.Button(wF, text="Tareas", width=12, command = lambda: V_modTar(modo))
    BF3.place(x=510, y=680)

    # boton Volver
    BF3=tk.Button(wF, text="Volver", width=12, command=lambda: menuFrom(wF))
    BF3.place(x=340, y=730)

    # boton Salir
    BF1=tk.Button(wF, text="Salir", width=12, command = lambda: salirConAviso())
    BF1.place(x=510, y=730)

    wF.tabla.delete(*wF.tabla.get_children())

    # proceso de los permisos
    refresh_conn(my_conn)
    if modo == 'proy':
        if actualRol & 16384: # modificación proyectos habilitado?
            try:
                my_cursor = my_conn.cursor()
                statement = "SELECT * FROM proyectos WHERE area = 'Proyectos' AND estado = 'En curso' ORDER BY numero_cdf DESC" # proyectos activos
                my_cursor.execute(statement)
                resultados = my_cursor.fetchall()
                #print(resultados) 
            except Exception as e:
                print("error 982", e)
        else:
            try:
                my_cursor = my_conn.cursor()
                statement = "SELECT * FROM proyectos WHERE area = 'Proyectos' AND estado = 'En curso' and responsable = %s ORDER BY numero_cdf DESC" # proyectos del usuario
                values = (actualUser,)
                my_cursor.execute(statement, values)
                resultados = my_cursor.fetchall()
                #print(resultados) 
            except Exception as e:
                print("error 1004", e)
    if modo == 'proy_c':
        try:
            my_cursor = my_conn.cursor()
            statement = "SELECT * FROM proyectos WHERE area = 'Proyectos' AND estado <> 'En curso'ORDER BY numero_cdf DESC" # proyectos cerrados
            my_cursor.execute(statement)
            resultados = my_cursor.fetchall()
            #print(resultados) 
        except Exception as e:
            print("error 982", e)
        
    if modo == 'des':
        if actualRol & 64: # modificación desarrollos habilitado?
            try:
                my_cursor = my_conn.cursor()
                statement = "SELECT * FROM proyectos WHERE area = 'Desarrollos' AND estado = 'En curso' ORDER BY numero_cdf DESC" # proyectos activos
                my_cursor.execute(statement)
                resultados = my_cursor.fetchall()
                #print(resultados) 
            except Exception as e:
                print("error 1020", e)
        else:
            try:
                my_cursor = my_conn.cursor()
                statement = "SELECT * FROM proyectos WHERE area = 'Desarrollos' AND estado = 'En curso' and responsable = %s ORDER BY numero_cdf DESC" # proyectos del usuario
                values = (actualUser,)
                my_cursor.execute(statement, values)
                resultados = my_cursor.fetchall()
                #print(resultados) 
            except Exception as e:
                print("error 1014", e)
    if modo == 'des_c':
        try:
            my_cursor = my_conn.cursor()
            statement = "SELECT * FROM proyectos WHERE area = 'Desarrollos' AND estado <> 'En curso' ORDER BY numero_cdf DESC" # proyectos cerrados
            my_cursor.execute(statement)
            resultados = my_cursor.fetchall()
            #print(resultados) 
        except Exception as e:
            print("error 982", e)
    if modo == 'ate':
        if actualRol & 256: # modificación asistencias habilitado?
            try:
                my_cursor = my_conn.cursor()
                statement = "SELECT * FROM proyectos WHERE area = 'Asistencias' AND estado = 'En curso' ORDER BY numero_cdf DESC" # proyectos activos
                my_cursor.execute(statement)
                resultados = my_cursor.fetchall()
                #print(resultados) 
            except Exception as e:
                print("error 1029", e)
        else:
            try:
                my_cursor = my_conn.cursor()
                statement = "SELECT * FROM proyectos WHERE area = 'Asistencias' AND estado = 'En curso' and responsable = %s ORDER BY numero_cdf DESC" # proyectos del usuario
                values = (actualUser,)
                my_cursor.execute(statement, values)
                resultados = my_cursor.fetchall()
                #print(resultados) 
            except Exception as e:
                print("error 1014", e)
    if modo == 'ate_c':
        try:
            my_cursor = my_conn.cursor()
            statement = "SELECT * FROM proyectos WHERE area = 'Asistencias' AND estado <> 'En curso' ORDER BY numero_cdf DESC" # proyectos cerrados
            my_cursor.execute(statement)
            resultados = my_cursor.fetchall()
            #print(resultados) 
        except Exception as e:
            print("error 982", e)
    # Parte común a todos los modos
    for fila in resultados:

        # muestra el nombre de la solicitud, por referencia
        try:
            my_cursor = my_conn.cursor()
            statement = "SELECT asunto FROM solicitudes WHERE ID_solicitud = %s"
            values = (fila[8],)
            my_cursor.execute(statement, values)
            pedido = my_cursor.fetchone()
            #print(pedido) 
        except Exception as e:
            print("error", e)

        wF.tabla.insert('',index = 'end', iid=None, text = str(fila[7]), values = [fila[3], fila[5], fila[1], fila[4], fila[6],fila[8], pedido[0]])
    wF.after(1, lambda: wF.focus_force())

#################### Alta de usuario #########################################################
def addUser():
    w1.withdraw()
    global w6
    w6=tk.Toplevel()
    w6.geometry("320x500")
    w6.iconphoto(False, photo)
    w6.title("CENADIF - Base de datos")

    L61 = tk.Label(w6, text = "Alta de usuario")
    L61.place(x=105, y=25)
    
    # ID_usuario
    L62 = tk.Label(w6, text = "ID de usuario: ")
    L62.place(x=25, y=75)

    global E61
    E61 = tk.Entry(w6)
    E61.place(x=155, y=75)
    
     # nombre_usuario
    L63 = tk.Label(w6, text = "Nombre de usuario: ")
    L63.place(x=25, y=125)

    global E62
    E62 = tk.Entry(w6)
    E62.place(x=155, y=125)

     # Nombres 
    L64 = tk.Label(w6, text = "Nombres: ")
    L64.place(x=25, y=175)

    global E63
    E63 = tk.Entry(w6)
    E63.place(x=155, y=175)

     # Apellidos 
    L65 = tk.Label(w6, text = "Apellidos: ")
    L65.place(x=25, y=225)

    global E64
    E64 = tk.Entry(w6)
    E64.place(x=155, y=225)

     # correo
    L66 = tk.Label(w6, text = "Correo electrónico: ")
    L66.place(x=25, y=275)

    global E65
    E65 = tk.Entry(w6)
    E65.place(x=155, y=275)

    # Perfil
    L68 = tk.Label(w6, text = "Perfil: ")
    L68.place(x=25, y=325)

    global E68
    E68 = tk.Entry(w6)
    E68.place(x=155, y=325)

    # Perfil
    L69 = tk.Label(w6, text = "Se asigna contraseña predeterminada")
    L69.place(x=25, y=375)

    # boton Volver
    B61=tk.Button(w6,text="Volver", command = lambda: menuFrom(w6))
    B61.place(x=120, y=425)

    # boton Ingresar
    B62=tk.Button(w6,text="Ingresar", command = userConfirm)
    B62.place(x=230, y=425)
    
    w6.after(1, lambda: w6.focus_force())

############################### ventana de cambio de privilegios ###############################
def privChange():
    w1.withdraw()
    global wA
    wA=tk.Toplevel()
    wA.geometry("500x600")
    wA.iconphoto(False, photo)
    wA.title("CENADIF Base de datos - Gestión de permisos")
    # tk.Entry Usuario
    LA1 = tk.Label(wA, text = "Usuario: ")
    LA1.place(x=50, y=40)

    list = userList()
    #print(list)
    global CEA1
    CEA1 = ttk.Combobox(wA, state="readonly", width = 17, values = list)
    CEA1.place(x=110, y=40)

    # Boton precargar
    BA1=tk.Button(wA,text="Precargar", width=12, command = precUsu)
    BA1.place(x=280, y=40)

    LA2 = tk.Label(wA, text = "Permisos:")
    LA2.place(x=50, y=80)

    # Radiotk.buttons (1x permiso)
    # Ver pedidos
    LA3 = tk.Label(wA, text = "Ver pedidos")
    LA3.place(x=50, y=120)
    global verPed
    verPed = tk.IntVar()
    global CA1
    CA1=tk.Checkbutton(wA, variable = verPed, onvalue = 1, offvalue = 0)
    CA1.place(x=190, y=120)
    # Ingresar pedidos
    LA4 = tk.Label(wA, text = "Ingresar pedidos")
    LA4.place(x=270, y=120)
    global ingPed
    ingPed = tk.IntVar()
    global CA2
    CA2=tk.Checkbutton(wA, variable = ingPed, onvalue = 1, offvalue = 0)
    CA2.place(x=410, y=120)
    # Asignar
    LA5 = tk.Label(wA, text = "Asignar")
    LA5.place(x=50, y=160)
    global asiPed
    asiPed = tk.IntVar()
    global CA3
    CA3=tk.Checkbutton(wA, variable = asiPed, onvalue = 1, offvalue = 0)
    CA3.place(x=190, y=160)
    # Ver documentación 
    LA6 = tk.Label(wA, text = "Ver documentación")
    LA6.place(x=50, y=200)
    global verDoc
    verDoc = tk.IntVar()
    global CA4
    CA4=tk.Checkbutton(wA, variable = verDoc, onvalue = 1, offvalue = 0)
    CA4.place(x=190, y=200)
    # Modificar documentación 
    LA7 = tk.Label(wA, text = "Modificar documentación")
    LA7.place(x=270, y=200)
    global modDoc
    modDoc = tk.IntVar()
    global CA5
    CA5=tk.Checkbutton(wA, variable = modDoc, onvalue = 1, offvalue = 0)
    CA5.place   (x=410, y=200)
    # Ver desarrollos 
    LA8 = tk.Label(wA, text = "Ver desarrollos")
    LA8.place(x=50, y=240)
    global verDes
    verDes = tk.IntVar()
    global CA6
    CA6=tk.Checkbutton(wA, variable = verDes, onvalue = 1, offvalue = 0)
    CA6.place(x=190, y=240)
    # Modificar desarrollos 
    LA9 = tk.Label(wA, text = "Modificar desarrollos")
    LA9.place(x=270, y=240)
    global modDes
    modDes = tk.IntVar()
    global CA7
    CA7=tk.Checkbutton(wA, variable = modDes, onvalue = 1, offvalue = 0)
    CA7.place(x=410, y=240)
    # Ver asistencias 
    LA10 = tk.Label(wA, text = "Ver asistencias")
    LA10.place(x=50, y=280)
    global verAsi
    verAsi = tk.IntVar()
    global CA8
    CA8=tk.Checkbutton(wA, variable = verAsi, onvalue = 1, offvalue = 0)
    CA8.place(x=190, y=280)
    # Modificar asistencias 
    LA11 = tk.Label(wA, text = "Modificar asistencias")
    LA11.place(x=270, y=280)
    global modAsi
    modAsi = tk.IntVar()
    global CA9
    CA9=tk.Checkbutton(wA, variable = modAsi, onvalue = 1, offvalue = 0)
    CA9.place(x=410, y=280)
    # Ver laboratorio 
    LA12 = tk.Label(wA, text = "Ver laboratorio")
    LA12.place(x=50, y=320)
    global verLab
    verLab = tk.IntVar()
    global CA10
    CA10=tk.Checkbutton(wA, variable = verLab, onvalue = 1, offvalue = 0)
    CA10.place(x=190, y=320)
    # Modificar laboratorio 
    LA13 = tk.Label(wA, text = "Modificar laboratorio")
    LA13.place(x=270, y=320)
    global modLab
    modLab = tk.IntVar()
    global CA11
    CA11=tk.Checkbutton(wA, variable = modLab, onvalue = 1, offvalue = 0)
    CA11.place(x=410, y=320)
    # Alta usuario
    LA14 = tk.Label(wA, text = "Alta de usuario")
    LA14.place(x=50, y=360)
    global altUsu
    altUsu = tk.IntVar()
    global CA12
    CA12=tk.Checkbutton(wA, variable = altUsu, onvalue = 1, offvalue = 0)
    CA12.place(x=190, y=360)
    # Gestionar permisos
    LA15 = tk.Label(wA, text = "Gestionar permisos")
    LA15.place(x=270, y=360)
    global gesPer
    gesPer = tk.IntVar()
    global CA13
    CA13=tk.Checkbutton(wA, variable = gesPer, onvalue = 1, offvalue = 0)
    CA13.place(x=410, y=360)
    # Ver Proyectos
    LA16 = tk.Label(wA, text = "Ver proyectos")
    LA16.place(x=50, y=400)
    global verPro
    verPro = tk.IntVar()
    global CA14
    CA14=tk.Checkbutton(wA, variable = verPro, onvalue = 1, offvalue = 0)
    CA14.place(x=190, y=400)
    # Modificar Proyectos
    LA17 = tk.Label(wA, text = "Modificar proyectos")
    LA17.place(x=270, y=400)
    global modPro
    modPro = tk.IntVar()
    global CA15
    CA15=tk.Checkbutton(wA, variable = modPro, onvalue = 1, offvalue = 0)
    CA15.place(x=410, y=400)
    # Modificar Proyectos
    LA18 = tk.Label(wA, text = "Orden int. Laboratorio")
    LA18.place(x=50, y=440)
    global OILab
    OILab = tk.IntVar()
    global CA18
    CA18=tk.Checkbutton(wA, variable = OILab, onvalue = 1, offvalue = 0)
    CA18.place(x=190, y=440)
    # Modificar Recursos
    LA19 = tk.Label(wA, text = "Modificar recursos")
    LA19.place(x=270, y=440)
    global modRec
    modRec = tk.IntVar()
    global CA19
    CA19=tk.Checkbutton(wA, variable = modRec, onvalue = 1, offvalue = 0)
    CA19.place(x=410, y=440)

    # Boton Volver
    BA2=tk.Button(wA,text="Volver", width=12, command = lambda: menuFrom(wA))
    BA2.place(x=250, y=510)

    # Boton 
    global BA3
    BA3=tk.Button(wA,text="Modificar", width=12, command= modPriv, state = tk.DISABLED)
    BA3.place(x=360, y=510)

    # Boton 
    global BA4
    BA4=tk.Button(wA,text="Reset PW", width=12, command = resetPW, state = tk.DISABLED)
    BA4.place(x=50, y=510)

    wA.focus_force()

#################### Cambio de contraseña #########################################################
def passChange():
    ### esconde la pantalla de llamado
    try:
        w1.withdraw()
    except:
        pass
    try:
        w2.withdraw()
    except:
        pass
    global w5
    w5=tk.Toplevel()
    w5.geometry("330x300")
    w5.iconphoto(False, photo)
    w5.title("CENADIF - Base de datos")

    L51 = tk.Label(w5, text = "Cambio de contraseña")
    L51.place(x=105, y=25)
    
    # Contraseña actual
    L52 = tk.Label(w5, text = "Contraseña actual: ")
    L52.place(x=25, y=75)

    global E51
    #E51 = tk.Entry(w5)
    E51 = tk.Entry(w5, show = "*")
    E51.place(x=135, y=75)
     
     # Contraseña nueva
    L53 = tk.Label(w5, text = "Contraseña nueva: ")
    L53.place(x=25, y=125)

    global E52
    #E52 = tk.Entry(w5)
    E52 = tk.Entry(w5, show = "*")
    E52.place(x=135, y=125)

     # Segunda contraseña nueva 
    L54 = tk.Label(w5, text = "Contraseña nueva: ")
    L54.place(x=25, y=175)

    global E53
    #E53 = tk.Entry(w5)
    E53 = tk.Entry(w5, show = "*")
    E53.place(x=135, y=175)

    # boton Volver
    B51=tk.Button(w5,text="Volver", command = lambda: menuFrom(w5))
    B51.place(x=110, y=225)

    # boton Ingresar
    B52=tk.Button(w5,text="Cambiar", command = changeConfirm)
    B52.place(x=220, y=225)
    
    w5.after(1, lambda: w5.focus_force())

############################### Ventana de agenda de tareas ########################################################
def V_accSchedule():
    #print("Agenda")
    w1.withdraw()
    global wD
    wD=tk.Toplevel()
    wD.geometry("560x610")
    wD.iconphoto(False, photo)
    wD.title("Agenda de tareas")
    wD.frame = tk.Frame(wD)
    wD.frame.grid(rowspan=2, column=1, row=1)
    wD.tabla = ttk.Treeview(wD.frame, height=25)
    wD.tabla.grid(column=1, row=1)
    ladox = tk.Scrollbar(wD.frame, orient = tk.VERTICAL, command= wD.tabla.yview)
    ladox.grid(column=0, row = 1, sticky='ew') 
    ladoy = tk.Scrollbar(wD.frame, orient =tk.HORIZONTAL, command = wD.tabla.yview)
    ladoy.grid(column = 1, row = 0, sticky='ns')

    wD.tabla.configure(xscrollcommand = ladox.set, yscrollcommand = ladoy.set)
    wD.tabla['columns'] = ('tipo', 'numero', 'fecha_prog', 'estado','correo')
    wD.tabla.column('#0', minwidth=0, width=1, anchor='center')
    wD.tabla.column('tipo', minwidth=100, width=100 , anchor='center')
    wD.tabla.column('numero', minwidth=100, width=100, anchor='center' )
    wD.tabla.column('fecha_prog', minwidth=100, width=120 , anchor='center')
    wD.tabla.column('estado', minwidth=100, width=100, anchor='center')
    wD.tabla.column('correo', minwidth=100, width=100, anchor='center')

    #wD.tabla.heading('#0', text='ID', anchor ='e')
    wD.tabla.heading('tipo', text='Tipo', anchor ='center')
    wD.tabla.heading('numero', text='Número', anchor ='center')
    wD.tabla.heading('fecha_prog', text='Fecha programada', anchor ='center')
    wD.tabla.heading('estado', text='Estado', anchor ='center')
    wD.tabla.heading('correo', text='Correo aviso', anchor ='center')

    # boton Salir
    BD1=tk.Button(wD,text="Salir", width=12, command = lambda: salirConAviso())
    BD1.place(x=430, y=560)

    # boton Actualizar
    #BD2=tk.Button(wD,text="Actualizar", width=12)
    #BD2.place(x=735, y=560)

    # boton Volver
    BD3=tk.Button(wD,text="Volver", width=12, command = lambda: menuFrom(wD))
    BD3.place(x=300, y=560)

    wD.tabla.delete(*wD.tabla.get_children())

    refresh_conn(my_conn)
    try:
            my_cursor = my_conn.cursor()
            statement = "SELECT * FROM agenda WHERE nombre_usuario = %s "
            #values=(actualID)
            my_cursor.execute(statement, (actualUser,))
            resultados = my_cursor.fetchall()
            #print(f"Agenda del usuario: {resultados}") 
    except Exception as e:
        print("error", e)
    
    for fila in resultados:
        try:
            my_cursor = my_conn.cursor()
            statement = "SELECT estado FROM tareas WHERE ID_tarea = %s "
            #values=(actualID)
            my_cursor.execute(statement, (fila[3],))
            estado = my_cursor.fetchone() 
        except Exception as e:
            print("error", e)
        
        wD.tabla.insert('',index = fila[2], iid=None, values = [fila[1], fila[2], fila[5], estado[0], fila[6]])

    wD.after(1, lambda: wD.focus_force())

########################################################################################################
def actualizar():
    #print("Actualizando")
    win_geometry = w3.winfo_geometry()
    w3.destroy()
    V_viewRequest(win_geometry)

########################################################################################################
def menuFrom(window):
    window.destroy()
    w1.deiconify()

########################################################################################################
def OrdenesFrom(window, modo):
    window.destroy()
    V_modLab(modo)
    #V_modLab(tipo, geom = ''):

########################################################################################################
def ingresar_pedido():
    fechaIng = E41.get()
    if date_check(fechaIng) == False:
        return
    cliente = E42.get()
    filiacion = E43.get()
    asunto = E44.get()
    descripcion = T41.get("1.0",'end-1c')
    operador = C45.get()
    correo = E48.get()
    telefono = E49.get()

    refresh_conn(my_conn)
    if(len(cliente) != 0 and len(filiacion) != 0 and len(asunto) != 0 and len(correo) != 0 ):
        try:  
            my_cursor = my_conn.cursor()
            statement = '''INSERT INTO solicitudes (asunto, fecha_ingreso, filiacion_cliente, cliente, descripcion_asunto, operador, correo, telefono) 
            VALUES('{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}')'''.format(asunto, fechaIng, filiacion, cliente, descripcion, operador, correo, telefono )
            my_cursor.execute(statement)
            my_conn.commit()
        #    my_conn.close()
            messagebox.showinfo(message="Nuevo pedido registrado.", title="Aviso del sistema")

        except Exception as e:
            print("error", e)
            logging.error(e, exc_info=True)
            messagebox.showinfo(message="pedido NO registrado.", title="Aviso del sistema")
        else:
            menuFrom(w4)    

    else:
         messagebox.showinfo(message="Campos vacios.", title="Aviso del sistema")

################### Función Lista desplegable de usuarios #############################
def userList():
    refresh_conn(my_conn)
    try:
        my_cursor = my_conn.cursor()
        statement = "SELECT nombre_usuario FROM usuarios WHERE nombre_usuario <> 'admin' ORDER BY nombre_usuario ASC"
        my_cursor.execute(statement)
        list = my_cursor.fetchall()
        return list
        #print(resultados) 
    except Exception as e:
        print("error", e)

##################### funcion de asignar ########################################################
def B_assign():
    curItem = w7.tabla.focus()
    idSol = w7.tabla.item(curItem).get('text')
    #idSol = E71.get()
    if len(idSol) != 0:
        tipoSal = C70.get()
        fechaSal = E73.get()
        if date_check(fechaSal) == False:
            return
        descripcion = T71.get("1.0",'end-1c')
        responsable = C74.get()

        # chequeo de campos no nulos
        refresh_conn(my_conn)
        if (len(tipoSal) != 0 and len(fechaSal) != 0 and (len(responsable)!=0 or (tipoSal != 'Proyectos') )):
        
            # Documentación
            if tipoSal == 'Documentacion':
                #print('Tipo de salida Documentacion')
                # crea entrada tabla documentación
                try:    
                    my_cursor = my_conn.cursor()
                    statement = '''INSERT INTO documentacion (ID_solicitud) VALUES('{}')'''.format(idSol )
                    my_cursor.execute(statement)
                    my_conn.commit() 

                except Exception as e:
                    print("error", e)  
            
                # actualiza tabla de pedidos
                try:
                    my_cursor = my_conn.cursor()
                    statement = "UPDATE solicitudes SET tipo_salida = %s  , fecha_salida = %s , descripcion_salida = %s WHERE ID_solicitud = %s" 
                    values = (tipoSal, fechaSal, descripcion, idSol )
                    my_cursor.execute(statement, values)
                    my_conn.commit() 

                except Exception as e:
                    print("error", e)

                messagebox.showinfo(message="Pedido asignado al área de Documentación y Normas.", title="Aviso del sistema")  
                menuFrom(w7)

            elif tipoSal == 'Proyectos': 
                # crea entrada tabla proyectos
                try:    
                    my_cursor = my_conn.cursor()
                    statement = '''INSERT INTO proyectos (ID_solicitud, responsable, area, estado, numero_cdf) VALUES('{}', '{}','{}','{}','{}')'''.format(idSol, responsable, 'Proyectos', 'En curso', 0)
                    #statement = '''INSERT INTO proyectos (ID_solicitud, responsable, area, estado, numero_cdf) VALUES('{}', '{}','{}','{}','{}') ON DUPLICATE KEY UPDATE numero_cdf = numero_cdf - 1.format(idSol, responsable, 'Proyectos', 'En curso', 0)'''
                    my_cursor.execute(statement)
                    my_conn.commit() 

                except Exception as e:
                    print("error 1428", e)  
            
                # actualiza tabla de pedidos
                try:
                    my_cursor = my_conn.cursor()
                    statement = "UPDATE solicitudes SET tipo_salida = %s  , fecha_salida = %s , descripcion_salida = %s WHERE ID_solicitud = %s" 
                    values = (tipoSal, fechaSal, descripcion, idSol )
                    my_cursor.execute(statement, values)
                    my_conn.commit() 

                except Exception as e:
                    print("error 1439", e)

                messagebox.showinfo(message="Pedido asignado a un proyecto.", title="Aviso del sistema")
                menuFrom(w7)

            elif tipoSal == 'Desarrollos':
                # crea entrada tabla proyectos
                try:    
                    my_cursor = my_conn.cursor()
                    statement = '''INSERT INTO proyectos (ID_solicitud, responsable, area, estado, numero_cdf) VALUES('{}', '{}','{}','{}','{}')'''.format(idSol, responsable, 'Desarrollos', 'En curso', 0)
                    my_cursor.execute(statement)
                    my_conn.commit() 

                except Exception as e:
                    print("error 2424", e)
                # actualiza tabla de pedidos
                try:
                    my_cursor = my_conn.cursor()
                    statement = "UPDATE solicitudes SET tipo_salida = %s  , fecha_salida = %s , descripcion_salida = %s WHERE ID_solicitud = %s" 
                    values = (tipoSal, fechaSal, descripcion, idSol )
                    my_cursor.execute(statement, values)
                    my_conn.commit() 

                except Exception as e:
                    print("error 2434", e)

                messagebox.showinfo(message="Pedido asignado al área de Desarrollos.", title="Aviso del sistema")
                menuFrom(w7)

            elif tipoSal == 'Asistencias': 
                # crea entrada tabla asistencias
                try:    
                    my_cursor = my_conn.cursor()
                    statement = '''INSERT INTO proyectos (ID_solicitud, responsable, area, estado, numero_cdf) VALUES('{}', '{}','{}','{}','{}')'''.format(idSol, responsable, 'Asistencias', 'En curso', 0)
                    my_cursor.execute(statement)
                    my_conn.commit() 

                except Exception as e:
                    print("error", e)  
            
                # actualiza tabla de pedidos
                try:
                    my_cursor = my_conn.cursor()
                    statement = "UPDATE solicitudes SET tipo_salida = %s  , fecha_salida = %s , descripcion_salida = %s WHERE ID_solicitud = %s" 
                    values = (tipoSal, fechaSal, descripcion, idSol )
                    my_cursor.execute(statement, values)
                    my_conn.commit() 

                except Exception as e:
                    print("error", e)

                messagebox.showinfo(message="Pedido asignado al área de Asistencias técnicas.", title="Aviso del sistema")
                menuFrom(w7)

            elif tipoSal == 'Laboratorio':
                try:
                    my_cursor = my_conn.cursor()
                    statement = "SELECT numero_cdf FROM ordenes WHERE tipo = 'T'"
                    my_cursor.execute(statement)
                    num_OT = my_cursor.fetchall()
                    #print(num_OI)
                except Exception as e:
                    print("error 1725", e)

                max = 0
                for n in num_OT:
                    if n[0] > max:
                        max = n[0]
                if max < 10000:
                    max = 10000
                elif max >= 10000:
                    max += 1
                else:
                    pass # To Do manejo de error 
                # crea entrada tabla ordenes
                try:    
                    my_cursor = my_conn.cursor()
                    statement = '''INSERT INTO ordenes (ID_solicitud, tipo, numero_cdf)
                        VALUES('{}', '{}','{}')'''.format(idSol, 'T', max )
                    my_cursor.execute(statement)
                    my_conn.commit() 

                except Exception as e:
                    print("error 1712", e)  
            
                # actualiza tabla de pedidos
                try:
                    my_cursor = my_conn.cursor()
                    statement = "UPDATE solicitudes SET tipo_salida = %s  , fecha_salida = %s , descripcion_salida = %s WHERE ID_solicitud = %s" 
                    values = (tipoSal, fechaSal, descripcion, idSol )
                    my_cursor.execute(statement, values)
                    my_conn.commit() 

                except Exception as e:
                    print("error", e)

                messagebox.showinfo(message="Pedido asignado al área de Laboratorio.", title="Aviso del sistema")
                menuFrom(w7)


            elif tipoSal == 'Cierre':  
                        # actualiza tabla de pedidos
                try:
                    my_cursor = my_conn.cursor()
                    statement = "UPDATE solicitudes SET tipo_salida = %s  , fecha_salida = %s , descripcion_salida = %s WHERE ID_solicitud = %s" 
                    values = (tipoSal, fechaSal, descripcion, idSol )
                    my_cursor.execute(statement, values)
                    my_conn.commit() 

                except Exception as e:
                    print("error", e)

                messagebox.showinfo(message="Pedido cerrado.", title="Aviso del sistema")
                menuFrom(w7)

            else:
                print('Error de Tipo de salida')
        else:
            messagebox.showinfo(message="Campos obligatorios vacios.", title="Aviso del sistema")
    else:
        messagebox.showinfo(message="Seleccionar un pedido.", title="Aviso del sistema")

##################### Botón de modificar documentos ############################################
def B_modDoc():
    #print('Modificando docs')
    numDyN = E81.get()
    #print(numDyN)
    tipoDoc = C81.get()
    numDoc = E83.get()
    titulo = E84.get()
    fuente = C82.get()
    emisor = E87.get()
    fechaRes = E88.get()
    if len(fechaRes)!=0: # realiza chequeo solo si se modificó la fecha
        if date_check(fechaRes) == False:
            return
    cantDoc = E89.get()
    respo = E8A.get()
    precio = E8D.get()
    uniMon = C83.get()
    observ = T81.get("1.0",'end-1c')
    #print(tipoDoc)
    #print(len(tipoDoc))
    refresh_conn(my_conn)
    if len(tipoDoc) != 0:
        try:
            my_cursor = my_conn.cursor()
            statement = "UPDATE documentacion SET tipo_doc = %s WHERE ID_trabajo = %s"
            val = (tipoDoc, numDyN)
            my_cursor.execute(statement,val)
            my_conn.commit()
            #    my_conn.close()

        except Exception as e:
            print("error", e)

    if len(numDoc) != 0:
        try:
            my_cursor = my_conn.cursor()
            statement = "UPDATE documentacion SET numero_doc = %s WHERE ID_trabajo = %s"
            val = (numDoc, numDyN)
            my_cursor.execute(statement,val)
            my_conn.commit()
            #    my_conn.close()

        except Exception as e:
            print("error", e)

    if len(titulo) != 0:
        try:
            my_cursor = my_conn.cursor()
            statement = "UPDATE documentacion SET titulo = %s WHERE ID_trabajo = %s"
            val = (titulo, numDyN)
            my_cursor.execute(statement,val)
            my_conn.commit()
            #    my_conn.close()

        except Exception as e:
            print("error", e)
    
    if len(fuente) != 0:
        try:
            my_cursor = my_conn.cursor()
            statement = "UPDATE documentacion SET fuente = %s WHERE ID_trabajo = %s"
            val = (fuente, numDyN)
            my_cursor.execute(statement,val)
            my_conn.commit()
            #    my_conn.close()

        except Exception as e:
            print("error", e)

    if len(emisor) != 0:
        try:
            my_cursor = my_conn.cursor()
            statement = "UPDATE documentacion SET emisor = %s WHERE ID_trabajo = %s"
            val = (emisor, numDyN)
            my_cursor.execute(statement,val)
            my_conn.commit()
            #    my_conn.close()

        except Exception as e:
            print("error", e)

    if len(fechaRes) != 0:
        try:
            my_cursor = my_conn.cursor()
            statement = "UPDATE documentacion SET fecha_resolucion = %s WHERE ID_trabajo = %s"
            val = (fechaRes, numDyN)
            my_cursor.execute(statement,val)
            my_conn.commit()
            #    my_conn.close()

        except Exception as e:
            print("error", e)

    if len(cantDoc) != 0:
        #print('actualizando cantidad')
        try:
            my_cursor = my_conn.cursor()
            statement = "UPDATE documentacion SET cantidad = %s WHERE ID_trabajo = %s"
            val = (cantDoc, numDyN)
            my_cursor.execute(statement,val)
            my_conn.commit()
            #    my_conn.close()

        except Exception as e:
            print("error", e)

    if len(respo) != 0:
        try:
            my_cursor = my_conn.cursor()
            statement = "UPDATE documentacion SET respondio = %s WHERE ID_trabajo = %s"
            val = (respo, numDyN)
            my_cursor.execute(statement,val)
            my_conn.commit()
            #    my_conn.close()

        except Exception as e:
            print("error", e)

    if len(precio) != 0:
        try:
            my_cursor = my_conn.cursor()
            statement = "UPDATE documentacion SET precio = %s WHERE ID_trabajo = %s"
            val = (precio, numDyN)
            my_cursor.execute(statement,val)
            my_conn.commit()
            #    my_conn.close()

        except Exception as e:
            print("error", e)

    if len(uniMon) != 0:
        try:
            my_cursor = my_conn.cursor()
            statement = "UPDATE documentacion SET u_monetaria = %s WHERE ID_trabajo = %s"
            val = (uniMon, numDyN)
            my_cursor.execute(statement,val)
            my_conn.commit()
            #    my_conn.close()

        except Exception as e:
            print("error", e) 

    if len(observ) != 0:
        try:
            my_cursor = my_conn.cursor()
            statement = "UPDATE documentacion SET observaciones = %s WHERE ID_trabajo = %s"
            val = (observ, numDyN)
            my_cursor.execute(statement,val)
            my_conn.commit()
            #    my_conn.close()

        except Exception as e:
            print("error", e) 
    
    win_geometry = w8.winfo_geometry()
    w8.destroy()
    modDocs(win_geometry)
    
############################### Precargar documento ############################################
def precDoc():
    numDyN = E81.get()
    refresh_conn(my_conn)
    # actualiza tabla de pedidos
    try:
        my_cursor = my_conn.cursor()
        statement = "SELECT * FROM documentacion WHERE ID_trabajo = %s" 
        values = (numDyN,)
        my_cursor.execute(statement, values)
        resultado = my_cursor.fetchone()
        #print(resultado)
        #my_conn.commit() 

    except Exception as e:
        print("error", e)  

    #E82.delete(0, tk.END)              # tipo
    C81.set('')              # Tipo
    if resultado[2] is not None:
        C81.set(resultado[2])
    E83.delete(0, tk.END)              # número  
    if resultado[3] is not None:
        E83.insert(0, resultado[3]) 
    E84.delete(0, tk.END)              # título
    if resultado[4] is not None:
        E84.insert(0, resultado[4])  
    C82.set('')                     # fuente
    if resultado[5] is not None:
        C82.set(resultado[5])
    E8D.delete(0, tk.END)              # precio
    if resultado[6] is not None:
         E8D.insert(0, resultado[6]) 
    C83.set('')                     # unidad monetaria
    if resultado[7] is not None:
        C83.set(resultado[7])
    E87.delete(0, tk.END)              # emisor
    if resultado[8] is not None:
         E87.insert(0, resultado[8]) 
    E88.delete(0, tk.END)              # fecha resolución
    if resultado[9] is not None:
        fecha = resultado[9].strftime('%Y-%m-%d')
        #print(fecha)
        E88.insert(0,fecha)
    T81.delete("1.0","end")         # Observaciones
    if resultado[10] is not None:
         T81.insert("1.0", resultado[10]) 
    E89.delete(0, tk.END)              # cantidad
    if resultado[11] is not None:
         E89.insert(0, resultado[11])
    E8A.delete(0, tk.END)              # respondió
    if resultado[12] is not None:
         E8A.insert(0, resultado[12]) 

############# Botón seleccionar proyecto #####################################################
def B_selProy(modo):
    curItem = wF.tabla.focus() # diccionario de la fila seleccionada en la tabla
    #print(wF.tabla.item(curItem))
    aux = wF.tabla.item(curItem).get('values')
    #global currNcdf ##########################################################################
    #currNcdf = wF.tabla.item(curItem).get('text') ############################################
    #print(currProy)
    refresh_conn(my_conn)
    if len(aux)!=0:
        try:
            my_cursor = my_conn.cursor()
            statement = "SELECT * FROM proyectos WHERE ID_solicitud = %s"
            values = (aux[5],)  # aux[5] es el ID_solicitud
            my_cursor.execute(statement, values)
            resultado = my_cursor.fetchone()

            gantt(resultado[0], wF)
            # Precarga de los tk.entry
            EF1.delete(0, tk.END)              # N° Proy
            if resultado[7] is not None:
                EF1.insert(0, resultado[7])
            CFB.set('')                      # Responsable
            if resultado[1] is not None:
                CFB.set(resultado[1])
            EF4.delete(0, tk.END)              # nombre de proyecto
            if resultado[3] is not None:
                EF4.insert(0, resultado[3]) 
            CF9.set('')                       # estado
            if resultado[5] is not None:
                CF9.set(resultado[5]) 
            CF8.set('')                      # alcance
            if resultado[6] is not None: 
                CF8.set(resultado[6])
            if modo == 'des':
                CFC.set('')                   # prioridad
                if resultado[9] is not None:
                    CFC.set(resultado[9])
                EFD.delete(0, tk.END)              # clasificación
                if resultado[10] is not None:
                    EFD.insert(0, resultado[10]) 
                EFE.delete(0, tk.END)              # NUM
                if resultado[11] is not None:
                    EFE.insert(0, resultado[11]) 
            TF1.delete("1.0","end")         # descripción
            if resultado[4] is not None:
                TF1.insert("1.0", resultado[4])
            # habilitar botón modificar 
            if modo != 'proy_c' and modo != 'des_c' and modo != 'ate_c':
                BF2['state'] = tk.NORMAL 
        except Exception as e:
            print("error 890", e)
    else:
        if modo == 'proy':
            messagebox.showinfo(message="Seleccione un proyecto.", title="Aviso del sistema")
        if modo == 'des':
            messagebox.showinfo(message="Seleccione un desarrollo.", title="Aviso del sistema")
        if modo == 'ate':
            messagebox.showinfo(message="Seleccione una asistencia.", title="Aviso del sistema")
        
############# Botón modificar proyecto #####################################################
def B_modProy(modo):
    if modo == 'proy':
        area = 'Proyectos'
    if modo == 'des':
        area = 'Desarrollos'
    if modo == 'ate':
        area = 'Asistencias'
    curItem = wF.tabla.focus()  ##########################################################
    numCDF = wF.tabla.item(curItem).get('text') ##########################################
    #numCDF = currNcdf
    refresh_conn(my_conn)
    if len(numCDF)!=0:
        ######### Actualización del numero de proyecto
        NnumCDF = EF1.get()
        if numCDF != NnumCDF:
            # chequea que no exista NnumCDF
            try:
                my_cursor = my_conn.cursor()
                statement = "SELECT COUNT(numero_cdf) FROM proyectos WHERE area = %s and numero_cdf = %s "
                val = (area, NnumCDF)
                my_cursor.execute(statement,val)
                flg_exists = my_cursor.fetchall()
                #print (results)
            except Exception as e:
                print("error 2093", e)
            if flg_exists[0][0]:
                messagebox.showinfo(message="El ID escogido ya existe .", title="Aviso del sistema")
                return(1)
            else:                
                try:
                    my_cursor = my_conn.cursor()
                    statement = "UPDATE proyectos SET numero_cdf = %s WHERE area = %s and numero_cdf = %s"
                    val = (NnumCDF, area, numCDF)
                    my_cursor.execute(statement,val)
                    my_conn.commit()
                    numCDF = NnumCDF       
                except Exception as e:
                    print("error 2101", e)
        ######### Actualización del responsable
        respo = CFB.get()
        if len(respo) != 0:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE proyectos SET responsable = %s WHERE area = %s and numero_cdf = %s"
                val = (respo, area, numCDF)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error", e)
        ######### Actualización del nombre de proyecto
        nomProy = EF4.get()
        if len(nomProy) != 0:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE proyectos SET nombre_proy = %s WHERE area = %s and numero_cdf = %s"
                val = (nomProy, area, numCDF)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error", e)
        ######### Actualización del estado
        estado = CF9.get()
        if len(estado) != 0:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE proyectos SET estado = %s WHERE area = %s and numero_cdf = %s"
                val = (estado, area, numCDF)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error", e)
        ######### Actualización del alcance
        alcance = CF8.get()
        if len(alcance) != 0:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE proyectos SET alcance = %s WHERE area = %s and numero_cdf = %s"
                val = (alcance, area, numCDF)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error", e)
        if modo == 'des':
            ######### Actualización de la prioridad
            prioridad = CFC.get()
            if len(prioridad) != 0:
                try:
                    my_cursor = my_conn.cursor()
                    statement = "UPDATE proyectos SET prioridad = %s WHERE area = %s and numero_cdf = %s"
                    val = (prioridad, area, numCDF)
                    my_cursor.execute(statement,val)
                    my_conn.commit()
                except Exception as e:
                    print("error", e)
            ######### Actualización de la clasificación
            clasificacion = EFD.get()
            if len(clasificacion) != 0:
                try:
                    my_cursor = my_conn.cursor()
                    statement = "UPDATE proyectos SET clasificacion = %s WHERE area = %s and numero_cdf = %s"
                    val = (clasificacion, area, numCDF)
                    my_cursor.execute(statement,val)
                    my_conn.commit()
                except Exception as e:
                    print("error", e)
            ######### Actualización de NUM
            NUM = EFE.get()
            if len(NUM) != 0:
                try:
                    my_cursor = my_conn.cursor()
                    statement = "UPDATE proyectos SET NUM = %s WHERE area = %s and numero_cdf = %s"
                    val = (NUM, area, numCDF)
                    my_cursor.execute(statement,val)
                    my_conn.commit()
                except Exception as e:
                    print("error", e)
        ######### Actualización de descripción
        descr = TF1.get("1.0",'end-1c')
        if len(descr) != 0:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE proyectos SET descripcion = %s WHERE area = %s and numero_cdf = %s"
                val = (descr, area, numCDF)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error", e)
        ############ Refresco de la ventana
        win_geometry = wF.winfo_geometry()
        wF.destroy() 
        V_modPro(modo, win_geometry)
    else:
        if modo == 'proy':
            messagebox.showinfo(message="Seleccione un proyecto.", title="Aviso del sistema")
        if modo == 'des':
            messagebox.showinfo(message="Seleccione un desarrollo.", title="Aviso del sistema")
        if modo == 'ate':
            messagebox.showinfo(message="Seleccione una asistencia.", title="Aviso del sistema")

############# Ventana de modificación de tareas ##########################################
def V_modTar(modo, geom = ''):
    ### esconde la pantalla de llamado
    try:
        w1.withdraw()
    except:
        pass
    try:
        wF.withdraw()
    except:
        pass
    global wG
    wG=tk.Toplevel()
    if len(geom)==0:
        wG.state('zoomed')
    else:
        wG.geometry(geom)
    wG.iconphoto(False, photo)
    wG.title("CENADIF Base de datos - Tareas del proyecto ")
    refresh_conn(my_conn)
    if modo != 'user':
        # toma numero cenadif del proyecto seleccionado en la tabla de proyectos
        curItem = wF.tabla.focus()
        numCDF = wF.tabla.item(curItem).get('text')
        if len(numCDF)!=0:
            try:
                my_cursor = my_conn.cursor(buffered=True)
                if modo == 'proy' or modo == 'proy_c':
                    statement = "SELECT * FROM proyectos WHERE area = 'Proyectos' and numero_cdf = %s" 
                if modo == 'des' or modo == 'des_c':
                    statement = "SELECT * FROM proyectos WHERE area = 'Desarrollos' and numero_cdf = %s" 
                if modo == 'ate' or modo == 'ate_c': 
                    statement = "SELECT * FROM proyectos WHERE area = 'Asistencias' and numero_cdf = %s" 
                values = (numCDF,)
                my_cursor.execute(statement, values)
                proyData = my_cursor.fetchone()
                #print(proyData)
                #print(proyData[0]) 
            except Exception as e:
                print("error 392", e)

            wG.frame = tk.Frame(wG)
            wG.frame.grid(rowspan=2, column=1, row=1)
            wG.tabla = ttk.Treeview(wG.frame, height=1)
            wG.tabla.grid(column=1, row=1)

            ladox = tk.Scrollbar(wG.frame, orient = tk.VERTICAL, command= wG.tabla.yview)
            ladox.grid(column=0, row = 2, sticky='ew')
            #ladox.pack(side ='right', fill ='x') 
            ladoy = tk.Scrollbar(wG.frame, orient =tk.HORIZONTAL, command = wG.tabla.xview)
            ladoy.grid(column = 1, row = 0, sticky='ns')
            wG.tabla.configure(xscrollcommand = ladox.set, yscrollcommand = ladoy.set)
        
            wG.tabla['columns'] = ('nombre_proy','estado', 'responsable', 'descripción', 'alcance')
            wG.tabla.column('#0', minwidth=50, width=60, anchor='center')
            wG.tabla.column('nombre_proy', minwidth=100, width=350, anchor='center')
            wG.tabla.column('estado', minwidth=80, width=100 , anchor='center')
            wG.tabla.column('responsable', minwidth=100, width=100 , anchor='center')
            wG.tabla.column('descripción', minwidth=100, width=700, anchor='center' )
            wG.tabla.column('alcance', minwidth=100, width=150 , anchor='center')

            wG.tabla.heading('#0', text='N° PROY', anchor ='e')
            wG.tabla.heading('nombre_proy', text='Nombre del proyecto', anchor ='center')
            wG.tabla.heading('estado', text='Estado', anchor ='center')
            wG.tabla.heading('responsable', text='Responsable', anchor ='center')
            wG.tabla.heading('descripción', text='Descripción', anchor ='center')
            wG.tabla.heading('alcance', text='Alcance', anchor ='center')            
            wG.tabla.insert('',index = proyData[7], iid=None, text = str(proyData[7]), values = [proyData[3], proyData[5], proyData[1], proyData[4], proyData[6]])

            gantt(proyData[0], wG)

            try:
                my_cursor = my_conn.cursor(buffered=True)
                statement = "SELECT * FROM tareas WHERE ID_proyecto = %s" 
                values = (proyData[0],)
                my_cursor.execute(statement, values)
                resultados = my_cursor.fetchall()
                    #print(resultados) 
            except Exception as e:
                print("error 2159", e)

            # boton Eliminar
            global BG6
            BG6=tk.Button(wG, text="Eliminar", width=12, state = tk.DISABLED, command = lambda: B_elimTar(modo))
            BG6.place(x=375, y=705)

            # boton Nueva
            BG3=tk.Button(wG, text="Nueva", width=12, state = tk.DISABLED, command = lambda: B_nuevaTar(proyData[0], modo))
            BG3.place(x=505, y=705)
            if modo != 'proy_c' and modo != 'des_c' and modo != 'ate_c':
                BG3['state'] = tk.NORMAL
        else:
            wG.withdraw()
            wF.state("zoomed")
            wF.deiconify()
            messagebox.showinfo(message="Seleccione un proyecto.", title="Aviso del sistema")
### Tabla de tareas ########################################################
    if modo == 'user' :
        #print('algo')
        try:
            my_cursor = my_conn.cursor()
            statement = "SELECT * FROM tareas WHERE responsable = %s" 
            values = (actualUser,)
            my_cursor.execute(statement, values)
            resultados = my_cursor.fetchall()
            #print(resultados) 
        except Exception as e:
            print("error 480", e)

#if modo == 'proy' or modo == 'user' or modo == 'des' or modo == 'ate' :
    wG.frame2 = tk.Frame(wG)
    wG.frame2.grid(rowspan=2, column=1, row=2)
    wG.frame2.place(y=70)
    wG.tabla2 = ttk.Treeview(wG.frame2, height=12)
    wG.tabla2.grid(column=1, row=1)

    ladox2 = tk.Scrollbar(wG.frame2, orient = tk.VERTICAL, command= wG.tabla2.yview)
    ladox2.grid(column=0, row = 1, sticky='ew') 
    ladoy2 = tk.Scrollbar(wG.frame2, orient =tk.HORIZONTAL, command = wG.tabla2.xview)
    ladoy2.grid(column = 1, row = 0, sticky='ns')
    #wG.tabla2.configure(xscrollcommand = ladox.set, yscrollcommand = ladoy.set)

    wG.tabla2['columns'] = ('nombre', 'estado', 'responsable', 'descripción', 'fecha_inicio', 'fecha_fin', 'avance')
    wG.tabla2.column('#0', minwidth=50, width=60, anchor='center')
    wG.tabla2.column('nombre', minwidth=100, width=350, anchor='center')
    wG.tabla2.column('estado', minwidth=80, width=100 , anchor='center')
    wG.tabla2.column('responsable', minwidth=100, width=100 , anchor='center')
    wG.tabla2.column('descripción', minwidth=100, width=400, anchor='center' )
    wG.tabla2.column('fecha_inicio', minwidth=100, width=150 , anchor='center')
    wG.tabla2.column('fecha_fin', minwidth=100, width=150 , anchor='center')
    wG.tabla2.column('avance', minwidth=100, width=150 , anchor='center')

    wG.tabla2.heading('#0', text='Tarea', anchor ='e')
    wG.tabla2.heading('nombre', text='Nombre de la tarea', anchor ='center')
    wG.tabla2.heading('estado', text='Estado', anchor ='center')
    wG.tabla2.heading('responsable', text='Responsable', anchor ='center')
    wG.tabla2.heading('descripción', text='Descripción', anchor ='center')
    wG.tabla2.heading('fecha_inicio', text='Inicio', anchor ='center')
    wG.tabla2.heading('fecha_fin', text='Finaliz.', anchor ='center')
    wG.tabla2.heading('avance', text='Avance(%)', anchor ='center')

    
    try:
        for fila in resultados:
            wG.tabla2.insert('',index = fila[0], iid=None, text = str(fila[0]), values = [fila[3], fila[16], fila[2], fila[4], fila[14], fila[15], fila[6]])
    except:
        pass
    #wF.after(1, lambda: wF.focus_force())

## PRIMER PISO ##########################
# Nombre de la tarea
    LG1 = tk.Label(wG, text = "Nombre Tarea")
    LG1.place(x=20, y=380)

    global EG1
    EG1 = tk.Entry(wG)
    EG1.place(x=110, y=380, width=320)

# Estado
    LG2 = tk.Label(wG, text = "Estado")
    LG2.place(x=430, y=380)

    global CG2
    CG2 = ttk.Combobox(wG, state="readonly", width = 17, values = ('Planificada','En curso', 'Cerrada', 'Pausada', 'Finalizada'))
    CG2.place(x=480, y=380)

## SEGUNDO PISO #########################
    # Responsable
    LG3 = tk.Label(wG, text = "Responsable")
    LG3.place(x=20, y=430)

    list = userList()
    global CG3
    CG3 = ttk.Combobox(wG, state="readonly", width = 17, values = list)
    CG3.place(x=110, y=430)

# inicio
    LG4 = tk.Label(wG, text = "Inicio")
    LG4.place(x=250, y=430)

    global EG4
    EG4 = tk.Entry(wG)
    EG4.place(x=300, y=430)

# Fin  
    LG5 = tk.Label(wG, text = "Fin")
    LG5.place(x=430, y=430)

    global EG5
    EG5 = tk.Entry(wG)
    EG5.place(x=480, y=430)

## TERCER PISO #########################
    # Avance
    LG6 = tk.Label(wG, text = "Avance")
    LG6.place(x=20, y=480)

    global EG6
    EG6 = tk.Entry(wG)
    EG6.place(x=110, y=480)

# H/H prog
    LG7 = tk.Label(wG, text = "hs.prog")
    LG7.place(x=250, y=480)

    global EG7
    EG7 = tk.Entry(wG)
    EG7.place(x=300, y=480)

# H/H dedicadas  
    LG8 = tk.Label(wG, text = "hs.dedic")
    LG8.place(x=430, y=480)

    global EG8
    EG8 = tk.Entry(wG)
    EG8.place(x=480, y=480)

## CUARTO PISO #########################
    # Integrante 1
    LG9 = tk.Label(wG, text = "Asignados: 1")
    LG9.place(x=20, y=530)

    #global EG9
    #EG9 = tk.Entry(wG)
    #EG9.place(x=110, y=530)

    global CG9
    CG9 = ttk.Combobox(wG, state="readonly", width = 17, values = list)
    CG9.place(x=110, y=530)

    if modo == 'proy' or modo == 'proy_c' or modo == 'ate':
    # Integrante 2
        LGA = tk.Label(wG, text = "      2")
        LGA.place(x=250, y=530)

        global CGA
        CGA = ttk.Combobox(wG, state="readonly", width = 17, values = list)
        CGA.place(x=300, y=530)

    # Integrante 3  
        LGB = tk.Label(wG, text = "      3")
        LGB.place(x=430, y=530)

        global CGB
        CGB = ttk.Combobox(wG, state="readonly", width = 17, values = list)
        CGB.place(x=480, y=530)

    global TG1
    if modo == 'ate':
        # Integrante 4
        LGC = tk.Label(wG, text =  "4")
        LGC.place(x=80, y=580)

        global CGC
        CGC = ttk.Combobox(wG, state="readonly", width = 17, values = list)
        CGC.place(x=110, y=580)

        # Integrante 5
        LGD = tk.Label(wG, text =  "      5")
        LGD.place(x=250, y=580)

        global CGD
        CGD = ttk.Combobox(wG, state="readonly", width = 17, values = list)
        CGD.place(x=300, y=580)

        # Integrante 6
        LGE = tk.Label(wG, text =  "      6")
        LGE.place(x=430, y=580)

        global CGE
        CGE = ttk.Combobox(wG, state="readonly", width = 17, values = list)
        CGE.place(x=480, y=580)

    # Descripción (ATE)
        LGB = tk.Label(wG, text = "Descripción: ")
        LGB.place(x=20, y=630)

        TG1 = tk.Text(wG, width = 61, height = 3)
        TG1.place(x=110, y=630)
    else:
    # Descripción
        LGB = tk.Label(wG, text = "Descripción: ")
        LGB.place(x=20, y=580)

        TG1 = tk.Text(wG, width = 61, height = 6)
        TG1.place(x=110, y=580)


#   BOTONES #####################
    # boton Precarga
    BG4=tk.Button(wG, text="Seleccionar", width=12, command = lambda: B_selTar(modo))
    BG4.place(x=115, y=705)

    # boton Recursos
    BG7=tk.Button(wG, text="Recursos", width=12, command = lambda: V_empleos('tar'))
    BG7.place(x=115, y=750)

    # boton Modificar
    global BG2
    BG2=tk.Button(wG,text="Modificar", width=12, command = lambda: B_modTar(modo), state = tk.DISABLED)
    BG2.place(x=245, y=705)

    # boton Volver
    BG5=tk.Button(wG, text="Volver", width=12, command=lambda: menuFrom(wG))
    BG5.place(x=375, y=750)

    # boton Salir
    BG1=tk.Button(wG, text="Salir", width=12, command = lambda: salirConAviso())
    BG1.place(x=505, y=750)

    #else:
    #    print('error de modo ')

#############################################################################################################################
def userConfirm():
    ID_user = (E61.get())
    n_user = (E62.get())
    name = (E63.get())
    surname = (E64.get())
    email = (E65.get())
    rol = (E68.get())

    refresh_conn(my_conn)
    if (len(ID_user) != 0 and len(n_user) != 0 and len(name) != 0 and len(surname) != 0 and len(email) != 0 and len(rol) != 0):
        aux = bd.F_hash('1234')
        try:
            my_cursor = my_conn.cursor()
            statement = '''INSERT INTO usuarios (ID_usuarios, nombre_usuario, nombres, apellidos, hash, correo, rol) 
            VALUES('{}', '{}', '{}', '{}', '{}', '{}', '{}')'''.format(ID_user, n_user, name, surname, aux, email, rol )
            my_cursor.execute(statement)
            my_conn.commit() 

        except Exception as e:
            print("error", e)

        messagebox.showinfo(message="Nuevo usuario ingresado. La contraseña por defecto es 1234", title="Aviso del sistema")
    else:
        messagebox.showinfo(message="Campos obligatorios vacios.", title="Aviso del sistema")

############################### precargar usuario ##############################################
def precUsu():
    editUser = CEA1.get()

    # mySQL bajar rol del usuario
    refresh_conn(my_conn)
    try:
        my_cursor = my_conn.cursor()
        statement = "SELECT rol FROM usuarios WHERE nombre_usuario = %s" 
        values = (editUser,)
        my_cursor.execute(statement, values)
        resultado = my_cursor.fetchone()
        #print(resultado)
        #my_conn.commit() 

    except Exception as e:
        print("error", e)  
    # proceso de los permisos
    if resultado[0] & 1:
       CA1.select()
    if resultado[0] & 2:
       CA2.select()
    if resultado[0] & 4:
       CA3.select()
    if resultado[0] & 8:
       CA4.select()
    if resultado[0] & 16:
       CA5.select()
    if resultado[0] & 32:
       CA6.select()
    if resultado[0] & 64:
       CA7.select()
    if resultado[0] & 128:
       CA8.select()
    if resultado[0] & 256:
       CA9.select()
    if resultado[0] & 512:
       CA10.select()
    if resultado[0] & 1024:
       CA11.select()
    if resultado[0] & 2048:
       CA12.select()
    if resultado[0] & 4096:
       CA13.select()
    if resultado[0] & 8192:
       CA14.select()
    if resultado[0] & 16384:
       CA15.select()
    if resultado[0] & 32768:
       CA18.select()
    if resultado[0] & 65536:
       CA19.select()

    BA3['state'] = tk.NORMAL
    BA4['state'] = tk.NORMAL

############################### modificar privilegios ##########################################
def modPriv():
    #print('modificando permisos...')
    editUser = CEA1.get()
    privileges = tk.IntVar()
    privileges = 0
    if verPed.get() == 1:
        privileges |= 1
    if ingPed.get() == 1:
        privileges |= 2
    if asiPed.get() == 1:
        privileges |= 4
    if verDoc.get() == 1:
        privileges |= 8
    if modDoc.get() == 1:
        privileges |= 16
    if verDes.get() == 1: 
        privileges |= 32
    if modDes.get() == 1:
        privileges |= 64
    if verAsi.get() == 1:
        privileges |= 128
    if modAsi.get() == 1:
        privileges |= 256
    if verLab.get() == 1:
        privileges |= 512
    if modLab.get() == 1:
        privileges |= 1024
    if altUsu.get() == 1:
        privileges |= 2048
    if gesPer.get() == 1:
        privileges |= 4096
    if verPro.get() ==1:
        privileges |= 8192
    if modPro.get() ==1:
        privileges |= 16384
    if OILab.get() ==1:
        privileges |= 32768
    if modRec.get() ==1:
        privileges |= 65536
    

    refresh_conn(my_conn)
    try:
        my_cursor = my_conn.cursor()
        statement = "UPDATE usuarios SET rol = %s WHERE nombre_usuario = %s" 
        values = (privileges, editUser)
        my_cursor.execute(statement, values)
        #resultado = my_cursor.fetchone()
        #print(resultado)
        my_conn.commit() 
        messagebox.showinfo(message="Permisos actualizados.", title="Aviso del sistema")

    except Exception as e:
        print("error", e)  

############################### resetear contraseña ##########################################
def resetPW():
    editUser = CEA1.get()
    try:
        my_cursor = my_conn.cursor()
        statement = "UPDATE usuarios SET hash = %s WHERE nombre_usuario = %s"
        val = (bd.F_hash('1234'), editUser)
        my_cursor.execute(statement,val)
        my_conn.commit()

    except Exception as e:
        print("error", e)
    messagebox.showinfo(message="Contraseña actualizada (valor por defecto).", title="Aviso del sistema")
    

#######################################################################################################################
def changeConfirm():
    #print("Cambio de password")
    password = (E51.get())
    newPass1 = (E52.get())
    newPass2 = (E53.get())
    
    if(len(password) == 0 or len(newPass1) == 0 or len(newPass2) == 0):
        messagebox.showinfo(message="Campos vacios.", title="Aviso del sistema")
    elif newPass1 != newPass2:
        messagebox.showinfo(message="Entradas inconsistentes.", title="Aviso del sistema")
    else:
        refresh_conn(my_conn)
        try:
            my_cursor = my_conn.cursor()
            statement = "SELECT nombre_usuario, hash FROM usuarios WHERE nombre_usuario = %s"
            my_cursor.execute(statement, (actualUser,))
            resultados = my_cursor.fetchone()
            #print(resultados) 
        except Exception as e:
            print("error", e)

        aux = tuple([actualUser, bd.F_hash(password)])
        #print(aux)
        if  aux == resultados:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE usuarios SET hash = %s WHERE nombre_usuario = %s"
                val = (bd.F_hash(newPass1), actualUser)
                my_cursor.execute(statement,val)
                my_conn.commit()

            except Exception as e:
                print("error", e)
            messagebox.showinfo(message="Contraseña actualizada.", title="Aviso del sistema")
        else:
            messagebox.showinfo(message="Contraseña incorrecta.", title="Aviso del sistema")

############# Función Gantt ################################################################
def gantt(proy, window):
    # Create the data for the Gantt chart
    #refresh_conn(my_conn)
    try:
        my_cursor = my_conn.cursor()
        statement = "SELECT nombre_proy FROM proyectos WHERE ID_proyecto = %s"
        values = (proy,) 
        my_cursor.execute(statement, values)
        titulo = my_cursor.fetchone()
        #print(resultados) 
    except Exception as e:
        print("error", e)
    
    try:
        my_cursor = my_conn.cursor()
        statement = "SELECT nombre_tarea, fecha_inicio, fecha_fin, avance FROM tareas WHERE ID_proyecto = %s"
        values = (proy,) 
        my_cursor.execute(statement, values)
        resultados = my_cursor.fetchall()
        #print(resultados) 
    except Exception as e:
        print("error", e)
    tsk = [] # lista de tareas
    str = [] # lista de comienzos
    end = [] # lista de finales
    adv = [] # lista de avances
    flag_data = 0

    for fila, index in zip(resultados, range(0, len(resultados))):
        if fila[1] != None and fila[2] != None:
            tsk.append(repr(index+1) + '.' + fila[0][:7])
            str.append(fila[1])
            end.append(fila[2]) 
            try:
                adv.append(fila[3]/100)
            except:
                adv.append(0)
            flag_data = 1 
    global fig
    if flag_data == 0:
        canvas = FigureCanvasTkAgg(master=window)  
        canvas.draw()
        canvas.get_tk_widget().place(x=650, y=380, width=825, height=390)
        return 0

    df = pd.DataFrame({'task': tsk, 'start': pd.to_datetime(str), 'end': pd.to_datetime(end), 'completion_frac': adv})
    #def gantt(proy, window):
    df['days_to_start'] = (df['start'] - df['start'].min()).dt.days
    df['days_to_end'] = (df['end'] - df['start'].min()).dt.days
    df['task_duration'] = df['days_to_end'] - df['days_to_start'] + 1  # to include also the end date
    df['completion_days'] = df['completion_frac'] * df['task_duration']
    #print(df)

    span = df['days_to_end'].max()
    if span < 10:
        cell = 1
    elif span < 50:
        cell = 7
    else:
        #cell = int(round(span/8,0))
        cell = int(math.trunc(span/8))
    #print(span)
    #print(cell)
    xticks = np.arange(5, df['days_to_end'].max() + 2, cell)
    if cell < 30:
        xticklabels = pd.date_range(start=df['start'].min() + dt.timedelta(days=4), end=df['end'].max()).strftime("%d/%m")
    else:
        xticklabels = pd.date_range(start=df['start'].min() + dt.timedelta(days=4), end=df['end'].max()).strftime("%m/%y")

    ##########################################################################################################
    
    #global fig
    fig, ax = plt.subplots()
    ### ESTE BLOQUE SIRVE PARA FIGURA SEPARADA##########
    #fig, ax = plt.subplots(num = 'CENADIF - Diagrama Gannt de proyecto')
    #thismanager = plt.get_current_fig_manager()
    #img = tk.PhotoImage(file= "Recursos/iconoCDF.png")
    #thismanager.window.tk.call('wm', 'iconphoto', thismanager.window._w, img)
    ####################################################

    for index, row in df.iterrows():
        plt.barh(y=row['task'], width=row['task_duration'], left=row['days_to_start'] + 1, color= 'c', alpha=0.4)
        plt.barh(y=row['task'], width=row['completion_days'], left=row['days_to_start'] + 1, color= 'c')

    plt.title(titulo[0], fontsize=15)
    plt.gca().invert_yaxis()
    ax.set_xticks(xticks)
    ax.set_xticklabels(xticklabels[::cell])
    ax.xaxis.grid(True, alpha=0.5)
    #ax.legend(handles=patches, labels=team_colors.keys(), fontsize=11)

    # Marking the current date on the chart
    today = datetime.now()
    x_today = (today.date() - df['start'].min().date()).days
    if x_today <= span:
        ax.axvline(x=x_today, color='r', linestyle='dashed')
        ax.text(x=x_today+1, y=0, s=today.date(), color='r')

    #plt.show()
    canvas = FigureCanvasTkAgg(fig, master=window)  
    canvas.draw()
    canvas.get_tk_widget().place(x=650, y=380, width=825, height=390)
    return(1)

############################### Función del botón Eliminar tarea #############################################
def B_elimTar(modo):
    curItem = wG.tabla2.focus()
    numTar = wG.tabla2.item(curItem).get('text')
    if len(numTar)!=0:
        if (messagebox.askokcancel(message="Eliminar tarea?", title="Confirmación de acción")):
            refresh_conn(my_conn)
            try:
                my_cursor = my_conn.cursor()
                statement = "DELETE FROM tareas WHERE ID_tarea = %s"
                val = (numTar,)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error", e)
            
            win_geometry = wG.winfo_geometry()
            wG.destroy()
            V_modTar(modo, win_geometry)
        ############ Refresco de la ventana
    else:
        messagebox.showinfo(message="Seleccione una tarea.", title="Aviso del sistema")        

############################### Función del botón Nueva tarea ############################################
def B_nuevaTar(proy, modo):
#    print('nueva tarea')
    refresh_conn(my_conn)
    try:
        my_cursor = my_conn.cursor()
        statement = "INSERT INTO tareas SET nombre_tarea = %s, ID_proyecto = %s"
        val = ('Nueva tarea', proy)
        my_cursor.execute(statement,val)
        my_conn.commit()
    except Exception as e:
        print("error", e)
    win_geometry = wG.winfo_geometry()
    wG.destroy()
    V_modTar(modo, win_geometry)

############# Botón seleccionar tarea #####################################################
def B_selTar(modo):
    curItem = wG.tabla2.focus()
    global numTar
    numTar = wG.tabla2.item(curItem).get('text')
    #print(numTar)
    if len(numTar)!=0:
        #gantt(numpro, wF)
        refresh_conn(my_conn)
        try:
            my_cursor = my_conn.cursor()
            statement = "SELECT * FROM tareas WHERE ID_tarea = %s"
            values = (numTar,) 
            my_cursor.execute(statement, values)
            resultado = my_cursor.fetchone()
            #print(resultado)
            # Precarga de los tk.entry
            EG1.delete(0, tk.END)              # Nombre de tarea
            if resultado[3] is not None:
                EG1.insert(0, resultado[3])
            CG2.set('')                     # Estado
            if resultado[16] is not None:
                CG2.set( resultado[16])
            CG3.set('')                     # Responsable
            if resultado[2] is not None:
                CG3.set(resultado[2])
            EG4.delete(0, tk.END)              # Inicio
            if resultado[14] is not None:
                EG4.insert(0, resultado[14])
            EG5.delete(0, tk.END)              # Fin
            if resultado[15] is not None:
                EG5.insert(0, resultado[15])
            EG6.delete(0, tk.END)              # Avance
            if resultado[6] is not None:
                EG6.insert(0, resultado[6])
            EG7.delete(0, tk.END)              # hs/H programadas
            if resultado[5] is not None:
                EG7.insert(0, resultado[5])
            EG8.delete(0, tk.END)              # hs/H dedicadas
            if resultado[7] is not None:
                EG8.insert(0, resultado[7])
            CG9.set('')                     # Integrante 1
            if resultado[8] is not None:
                CG9.set(resultado[8])
            if modo == 'proy' or modo == 'proy_c' or modo == 'ate':
                CGA.set('')                     # Integrante 2
                if resultado[9] is not None:
                    CGA.set(resultado[9])
                CGB.set('')                     # Integrante 3
                if resultado[10] is not None:
                    CGB.set(resultado[10])
            if modo == 'ate':
                CGC.set('')                     # Integrante 4
                if resultado[11] is not None:
                    CGC.set(resultado[11])
                CGD.set('')                     # Integrante 5
                if resultado[12] is not None:
                    CGD.set(resultado[12])
                CGE.set('')                     # Integrante 6
                if resultado[13] is not None:
                    CGE.set(resultado[13])
            TG1.delete("1.0","end")             # Descripción
            if resultado[4] is not None:
                TG1.insert("1.0", resultado[4])
            
            if modo != 'proy_c'and modo != 'des_c' and modo != 'ate_c':
                BG2['state'] = tk.NORMAL # habilitación del botón Modificar
                BG6['state'] = tk.NORMAL # habilitación del botón Eliminar
            #####################################################################     
            
        except Exception as e:
            print("error", e)
    else:
        messagebox.showinfo(message="Seleccione una tarea.", title="Aviso del sistema")

############################### Función del botón Modificar tarea ############################################
def B_modTar(modo):
    #curItem = wG.tabla2.focus()
    #numTar = wG.tabla2.item(curItem).get('text')

    if len(numTar)!=0:
        refresh_conn(my_conn)
        ######### Actualización del nombre de tarea
        nomTar = EG1.get()
        if len(nomTar) != 0:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE tareas SET nombre_tarea = %s WHERE ID_tarea = %s"
                val = (nomTar, numTar)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error", e)
        ######### Actualización del estado de tarea
        estado = CG2.get()
        if len(estado) != 0:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE tareas SET estado = %s WHERE ID_tarea = %s"
                val = (estado, numTar)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error", e)
        ######### Actualización del estado de tarea
        responsable = CG3.get()
        if len(responsable) != 0:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE tareas SET responsable = %s WHERE ID_tarea = %s"
                val = (responsable, numTar)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error", e)
        ######### Actualización de Fecha de inicio
        inicio = EG4.get()
        if len(inicio) != 0:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE tareas SET fecha_inicio = %s WHERE ID_tarea = %s"
                val = (inicio, numTar)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error", e)
        ######### Actualización de Fecha de finalizacion
        fin = EG5.get()
        if len(fin) != 0:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE tareas SET fecha_fin = %s WHERE ID_tarea = %s"
                val = (fin, numTar)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error", e)
        ######### Actualización de avance
        avance = EG6.get()
        if len(avance) != 0:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE tareas SET avance = %s WHERE ID_tarea = %s"
                val = (avance, numTar)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error", e)
        ######### Actualización de horas programadas
        HH_prog = EG7.get()
        if len(HH_prog) != 0:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE tareas SET HH_prog = %s WHERE ID_tarea = %s"
                val = (HH_prog, numTar)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error", e)
        ######### Actualización de horas dedicadas
        HH_dedic = EG8.get()
        if len(HH_dedic) != 0:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE tareas SET HH_dedicadas = %s WHERE ID_tarea = %s"
                val = (HH_dedic, numTar)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error", e)
        ######### Actualización del Integrante 1
        integ_1 = CG9.get()
        if len(integ_1) != 0:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE tareas SET integrante_1 = %s WHERE ID_tarea = %s"
                val = (integ_1, numTar)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error", e)
        if modo == 'proy' or modo == 'ate':
            ######### Actualización del Integrante 2
            integ_2 = CGA.get()
            if len(integ_2) != 0:
                try:
                    my_cursor = my_conn.cursor()
                    statement = "UPDATE tareas SET integrante_2 = %s WHERE ID_tarea = %s"
                    val = (integ_2, numTar)
                    my_cursor.execute(statement,val)
                    my_conn.commit()
                except Exception as e:
                    print("error", e)
            ######### Actualización del Integrante 3
            integ_3 = CGB.get()
            if len(integ_3) != 0:
                try:
                    my_cursor = my_conn.cursor()
                    statement = "UPDATE tareas SET integrante_3 = %s WHERE ID_tarea = %s"
                    val = (integ_3, numTar)
                    my_cursor.execute(statement,val)
                    my_conn.commit()
                except Exception as e:
                    print("error", e)
        if modo == 'ate':
            ######### Actualización del Integrante 4
            integ_4 = CGC.get()
            if len(integ_4) != 0:
                try:
                    my_cursor = my_conn.cursor()
                    statement = "UPDATE tareas SET integrante_4 = %s WHERE ID_tarea = %s"
                    val = (integ_4, numTar)
                    my_cursor.execute(statement,val)
                    my_conn.commit()
                except Exception as e:
                    print("error", e)
            ######### Actualización del Integrante 5
            integ_5 = CGD.get()
            if len(integ_5) != 0:
                try:
                    my_cursor = my_conn.cursor()
                    statement = "UPDATE tareas SET integrante_5 = %s WHERE ID_tarea = %s"
                    val = (integ_5, numTar)
                    my_cursor.execute(statement,val)
                    my_conn.commit()
                except Exception as e:
                    print("error", e)
            ######### Actualización del Integrante 6
            integ_6 = CGE.get()
            if len(integ_6) != 0:
                try:
                    my_cursor = my_conn.cursor()
                    statement = "UPDATE tareas SET integrante_6 = %s WHERE ID_tarea = %s"
                    val = (integ_6, numTar)
                    my_cursor.execute(statement,val)
                    my_conn.commit()
                except Exception as e:
                    print("error", e)               

        ######### Actualización de descripción
        descr = TG1.get("1.0",'end-1c')
        if len(descr) != 0:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE tareas SET descripcion = %s WHERE ID_tarea = %s"
                val = (descr, numTar)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error", e)
        ############ actualizacion de agenda
        if modo == 'ate':
            F_actualizarAgenda(numTar)
            #pass
        ############ Refresco de la ventana
        try:
            if wG.tabla.winfo_exists():
                win_geometry = wG.winfo_geometry()
                wG.destroy()
                V_modTar(modo, win_geometry)
        except:         
            V_modTar('user', win_geometry)
    else:
        messagebox.showinfo(message="Seleccione una tarea.", title="Aviso del sistema")

##############################Función del botón Enviar (Inicio de sesión)#################################
#
# Pendientes:
#   - Limitar el número de intentos de ingreso a 3 y cerrar app
# 
# ######################################################################################################### 
def verify(event=None):
        global actualUser, actualID, actualRol
        actualUser = (E1.get()) 
        password = (E2.get())
        
        refresh_conn(my_conn)
        try:
            my_cursor = my_conn.cursor()
            statement = "SELECT hash, rol, ID_usuarios FROM usuarios WHERE nombre_usuario = %s"
            my_cursor.execute(statement, (actualUser,))
            resultados = my_cursor.fetchone()
        except Exception as e:
            print("error 3050", e)
        #aux = tuple([user, password])
        #print(aux)
        try:
            if  bd.F_hash(password) == resultados[0]:
                actualID = resultados[2]
                actualRol = resultados[1]
                #print(actualRol)
                if int(resultados[1]) & 1:
                    B11['state'] = tk.NORMAL # Ver listado de solicitudes
                if int(resultados[1]) & 2:
                    B12['state'] = tk.NORMAL # Ingreso de pedido
                if int(resultados[1]) & 4:
                    B13['state'] = tk.NORMAL # Asignar
                if int(resultados[1]) & 8:
                    B14['state'] = tk.NORMAL # ver pedidos asignados a documentación
                if int(resultados[1]) & 16:
                    B15['state'] = tk.NORMAL # modificar documentación
                if int(resultados[1]) & 32:
                    B16['state'] = tk.NORMAL # ver pedidos asignados a desarrollos
                if int(resultados[1]) & 64:
                    B17['state'] = tk.NORMAL # modificar desarrollos activos
                if int(resultados[1]) & 128:
                    B18['state'] = tk.NORMAL # ver pedidos asignados a asistencias
                if int(resultados[1]) & 256:
                    B19['state'] = tk.NORMAL # modificar Asistencias activas
                if int(resultados[1]) & 512:
                    B20['state'] = tk.NORMAL # ord. de trabajo de lab
                if int(resultados[1]) & 1024:
                    B21['state'] = tk.NORMAL # gestión ord. internas de lab
                if int(resultados[1]) & 2048:
                    B22['state'] = tk.NORMAL # Alta de usuario
                if int(resultados[1]) & 4096:
                    B23['state'] = tk.NORMAL # Permisos
                if int(resultados[1]) & 32768:
                    B1B['state'] = tk.NORMAL # ingreso ord. internas de lab
                if int(resultados[1]) & 65536:
                    B1C['state'] = tk.NORMAL # gestión de reccursos
                #w1.deiconify()
                if bd.F_hash('1234') == resultados[0]:
                    passChange()
                else:
                    w2.withdraw()
                    w1.deiconify()
            else:
                messagebox.showinfo(message="Nombre de usuario y/o contraseña incorrectos.", title="Aviso del sistema")
        except:
            messagebox.showinfo(message="Nombre de usuario y/o contraseña incorrectos.", title="Aviso del sistema")        

def accessForm():
    global w2
    w2=tk.Toplevel()
    w2.geometry("350x150")
    w2.iconphoto(False, photo)
    w2.title("CENADIF - Base de datos")    

    L1 = tk.Label(w2, text = "Formulario de acceso")#.pack(pady=10)
    L1.grid(row = 0, column = 1, padx = 5, pady= 5)
    
    # Username

    L2 = tk.Label(w2, text = "Usuario: ")
    L2.grid(row = 1, column = 0, padx = 5, pady= 5)

    global E1
    E1 = tk.Entry(w2)
    E1.grid(row = 1, column = 1, padx = 5, pady= 5)
    #E1.insert(0, "hgomezmolino") ############################### SACAR!!!!!!!!!!!!!!!!!!!!!!!!!

    # Password

    L3 = tk.Label(w2, text = "Contraseña: ")
    L3.grid(row = 2, column = 0, padx = 5, pady= 5)

    global E2
    E2 = tk.Entry(w2, show = "*")
    E2.grid(row = 2, column = 1, padx = 5, pady= 5)
    E2.bind('<Return>', verify)
    
    B1=tk.Button(w2,text="Enviar", command = verify)
    B1.grid(row = 3, column = 2, padx = 5, pady= 5)

    w2.focus_force()
###########################################################################################################
def V_estad():
    # Parte común a todos los modos
    #w1.withdraw()
    global wC
    wC=tk.Toplevel()
    wC.geometry("1500x800")
    wC.state("zoomed")
    wC.iconphoto(False, photo)
    # Titulo de la ventana 
    wC.title("CENADIF Base de datos - Estadísticas del sistema")
    # datos para gráficos #################
        # Gráfico 1 - pedidos
    areas = ["Documentación","Proyectos","Desarrollos","Asistencias","Laboratorio"]  
    pedidos = [0, 0, 0, 0, 0]                                                       
    
    refresh_conn(my_conn)
    try:
        my_cursor = my_conn.cursor()
        statement = "SELECT COUNT(*) AS count FROM documentacion;"
        my_cursor.execute(statement,)
        resultado = my_cursor.fetchall()
        pedidos[0] = resultado[0][0]
        for i in range(1,4):
            statement = "SELECT COUNT(*) AS count FROM proyectos where area = %s;"
            values = (areas[i],) 
            my_cursor.execute(statement, values)
            resultado = my_cursor.fetchall()
            #print(resultado)
            pedidos[i] = resultado[0][0]
        statement = "SELECT COUNT(*) AS count FROM ordenes;"
        my_cursor.execute(statement,)
        resultado = my_cursor.fetchall()
        pedidos[4] = resultado[0][0]
        #print(pedidos)
    except Exception as e:
        print("error 3157", e)
 
    total = sum(pedidos)
    colors = ['violet', 'blue', 'yellowgreen', 'gold', 'red'] 
 
    # Gráfico 2.1 - Desarrollos
    desarrollos = [0, 0, 0, 0]
    P_estados = ["En curso","Finalizados","Cancelados","Pausados"]
    sql_estados = ['En curso', 'Finalizado', 'Cancelado', 'Pausado']
    D_colors = ['limegreen', 'forestgreen', 'palegreen', 'springgreen']
    for i in range(4):
        try:
            my_cursor = my_conn.cursor()
            statement = "SELECT COUNT(*) AS count FROM proyectos where area = 'Desarrollos' and estado = %s;"
            values = (sql_estados[i],) 
            my_cursor.execute(statement,values)
            resultado = my_cursor.fetchall()
            desarrollos[i] = resultado[0][0]
        except Exception as e:
            print("error 3143", e)
    # Gráfico 2.2
    asistencias = [0, 0, 0, 0]
    A_colors = ['orange', 'goldenrod', 'antiquewhite', 'moccasin']
    for i in range(4):
        try:
            my_cursor = my_conn.cursor()
            statement = "SELECT COUNT(*) AS count FROM proyectos where area = 'Asistencias' and estado = %s;"
            values = (sql_estados[i],) 
            my_cursor.execute(statement,values)
            resultado = my_cursor.fetchall()
            asistencias[i] = resultado[0][0]
        except Exception as e:
            print("error 3156", e)
    # Gráfico 2.3
    proyectos = [0, 0, 0, 0]
    P_colors = ['blue', 'navy', 'aquamarine', 'turquoise']
    for i in range(4):
        try:
            my_cursor = my_conn.cursor()
            statement = "SELECT COUNT(*) AS count FROM proyectos where area = 'Proyectos' and estado = %s;"
            values = (sql_estados[i],) 
            my_cursor.execute(statement,values)
            resultado = my_cursor.fetchall()
            proyectos[i] = resultado[0][0]
        except Exception as e:
            print("error 3169", e)
    # Gráfico 2.4
    laboratorio = [0, 0]
    L_colors = ['red', 'darkred']
    lab_tipos = ['T', 'I']
    lab_labels = ['Ord.Trab', 'Ord.Int']
    for i in range(2):
        try:
            my_cursor = my_conn.cursor()
            statement = "SELECT COUNT(*) AS count FROM ordenes where tipo = %s;"
            values = (lab_tipos[i],) 
            my_cursor.execute(statement,values)
            resultado = my_cursor.fetchall()
            laboratorio[i] = resultado[0][0]
        except Exception as e:
            print("error 3169", e)
    # gráficos ###########################
    fig1 = plt.figure() # create a figure object
    fig1.tight_layout(h_pad=3, w_pad=3)
    fig1.suptitle('Tablero de control', fontsize=16)
    ax1 = fig1.add_subplot(131) # add an Axes to the figure
    ax1.pie(pedidos, radius= 1, labels=areas, colors=colors, autopct=lambda p: '{:.0f}'.format(p * total / 100))
    ax1.set_title('Solicitudes')
    ax2 = fig1.add_subplot(232) # add an Axes to the figure
    ax2.pie(desarrollos, labels=P_estados, colors=D_colors, autopct='%1.1f%%')
    ax2.set_title('Desarrollos')
    ax3 = fig1.add_subplot(233) # add an Axes to the figure
    ax3.pie(asistencias, labels=P_estados, colors=A_colors, autopct='%1.1f%%')
    ax3.set_title('Asistencias')
    ax4 = fig1.add_subplot(235) # add an Axes to the figure
    ax4.pie(proyectos, labels=P_estados, colors=P_colors, autopct='%1.1f%%')
    ax4.set_title('Proyectos')
    if pedidos[4] != 0:
        ax5 = fig1.add_subplot(236) # add an Axes to the figure
        ax5.pie(laboratorio, labels=lab_labels, colors=L_colors, autopct='%1.1f%%')
        ax5.set_title('Laboratorio')

    chart1 = FigureCanvasTkAgg(fig1,wC)
    chart1.get_tk_widget().pack(fill = tk.BOTH, expand = tk.TRUE)

############################### Función de actualización de agenda #################################################

def F_actualizarAgenda(num_tar):
    refresh_conn(my_conn)
    ### Query de numero proyecto #############
    try:
        my_cursor = my_conn.cursor()
        statement = "SELECT ID_proyecto FROM tareas WHERE ID_tarea = %s "
        my_cursor.execute(statement,(num_tar,))
        proy_AT = my_cursor.fetchone()
        #print(f"num proyecto: {proy_AT}")
    except Exception as e:
            print("error 3207", e)
    ### Query de numero de AT #############
    try:
        my_cursor = my_conn.cursor()
        statement = "SELECT numero_cdf FROM proyectos WHERE ID_proyecto = %s "
        my_cursor.execute(statement,proy_AT)
        num_AT = my_cursor.fetchone()
        #print(f"num AT: {num_AT}")
    except Exception as e:
            print("error 3207", e)
    ### Query de fecha planificada #############
    
    try:
        my_cursor = my_conn.cursor()
        statement = "SELECT responsable, integrante_1, integrante_2, integrante_3, integrante_4, integrante_5, integrante_6, fecha_inicio FROM tareas WHERE ID_tarea = %s "
        my_cursor.execute(statement,(num_tar,))
        equipo = my_cursor.fetchone()
        #print(f"equipo: {equipo}")
    except Exception as e:
            print("error 3207", e)
    ### Query de la agenda ##############################################
    
    try: 
        my_cursor = my_conn.cursor()
        statement = "SELECT nombre_usuario, fecha_programada FROM agenda WHERE ID_tarea = %s "
        my_cursor.execute(statement, (num_tar,))
        agenda = my_cursor.fetchall()
        #print(f"agenda: {agenda}") 
    except Exception as e:
            print("error 3232", e)

    for registro in agenda:
        flag_found = 0
    #    print(registro[0])
        for i in range(6):
            if registro[0] != None and registro[0]== equipo[i]:
                flag_found = 1
                # chequear fecha 
                #print (equipo[6], registro[1])
                if equipo[7] != registro[1]: # si la fecha de inicio (tarea) es distinta a fecha programda (agenda)
                    try:
                        my_cursor = my_conn.cursor()
                        statement = "UPDATE agenda SET fecha_programada = %s WHERE ID_tarea = %s AND nombre_usuario = %s "
                        values = (equipo[7], num_tar, equipo[i])
                        my_cursor.execute(statement, values)
                        my_conn.commit()
                        #equipo = my_cursor.fetchone()
                        #print(equipo)
                    except Exception as e:
                        print("error 3252", e)

        if flag_found == 0:
            #print('no encontrado')
            # elimino el registro de la agenda
            try:
                my_cursor = my_conn.cursor()
                statement = "DELETE FROM agenda WHERE ID_tarea = %s AND nombre_usuario = %s "
                values = (num_tar, registro[0])
                my_cursor.execute(statement, values)
                my_conn.commit()
                #equipo = my_cursor.fetchone()
                #print(equipo)
            except Exception as e:
                print("error 3266", e)

    for i in range(6):
        #print(persona)
        flag_found = 0
        for registro in agenda:
            if equipo[i] == registro[0]:
               flag_found = 1
        if flag_found == 0 and equipo[i] != None:
            #print('INSERT INTO, %s', equipo[i])
            try:  
                my_cursor = my_conn.cursor() 
                statement = '''INSERT INTO agenda (nombre_usuario, tipo, numero, ID_tarea, fecha_programada)
                  VALUES('{}', '{}', '{}', '{}', '{}')'''.format(equipo[i], 'ATE', num_AT[0], num_tar, equipo[7] )
                my_cursor.execute(statement)
                my_conn.commit()
            #    my_conn.close()

            except Exception as e:
                print("error3291", e)

def refresh_conn(conn):
    
    #pass
    try:
        cursor = my_conn.cursor() 
        statement = "SELECT 1"
        cursor.execute(statement)
        cursor.fetchall()
    except:
        conn.reconnect(attempts=3, delay=0)
        #print('Reconnect')
    else:
        pass
        #print('Conn OK -')

def V_modLab(tipo, geom = ''):
    w1.withdraw()
    global wH
    wH=tk.Toplevel()
    if len(geom)==0:
        wH.state('zoomed')
    else:
        wH.geometry(geom)
    wH.iconphoto(False, photo)

    wH.frame = tk.Frame(wH)
    wH.frame.grid(rowspan=2, column=1, row=1)
    wH.tabla = ttk.Treeview(wH.frame, height=23)
    wH.tabla.grid(column=1, row=1)

    ladox = tk.Scrollbar(wH.frame, orient = tk.VERTICAL, command= wH.tabla.yview)
    ladox.grid(column=0, row = 1, sticky='ew') 
    ladoy = tk.Scrollbar(wH.frame, orient =tk.HORIZONTAL, command = wH.tabla.xview)
    ladoy.grid(column = 1, row = 0, sticky='ns')
    wH.tabla.configure(xscrollcommand = ladox.set, yscrollcommand = ladoy.set)
       
    if tipo == 'T':
        wH.title("CENADIF - Laboratorio - Ordenes de trabajo")
        wH.tabla['columns'] = ('descripcion', 'responsable', 'fin', 'informe', 'pedido','cliente', 'ing_pedido')
        wH.tabla.column('#0', minwidth=50, width=60, anchor='center')
        wH.tabla.column('descripcion', minwidth=100, width=450, anchor='center')
        wH.tabla.column('responsable', minwidth=80, width=100 , anchor='center')
        wH.tabla.column('fin', minwidth=100, width=100 , anchor='center')
        wH.tabla.column('informe', minwidth=100, width=250, anchor='center' )
        wH.tabla.column('pedido', minwidth=50, width=50 , anchor='center')
        wH.tabla.column('cliente', minwidth=100, width=200, anchor='center')
        wH.tabla.column('ing_pedido', minwidth=150, width=250, anchor='center')   

       # Titulos 
        wH.tabla.heading('#0', text='N° O/T', anchor ='e')
        wH.tabla.heading('descripcion', text='Descripción', anchor ='center')
        wH.tabla.heading('responsable', text='Responsable', anchor ='center')
        wH.tabla.heading('fin', text='Fecha fin', anchor ='center')
        wH.tabla.heading('informe', text='Informe', anchor ='center')
        wH.tabla.heading('pedido', text='Pedido', anchor ='center')
        wH.tabla.heading('cliente', text='Cliente', anchor ='center')
        wH.tabla.heading('ing_pedido', text='Ing. pedido', anchor ='center')
 
    if tipo == 'I':
        wH.title("CENADIF - Laboratorio - Ordenes internas")
        wH.tabla['columns'] = ('descripcion', 'solicitante', 'inicio', 'fin', 'informe', 'responsable')
        wH.tabla.column('#0', minwidth=50, width=60, anchor='center')
        wH.tabla.column('descripcion', minwidth=100, width=550, anchor='center')
        wH.tabla.column('solicitante', minwidth=80, width=100 , anchor='center')
        wH.tabla.column('inicio', minwidth=100, width=100 , anchor='center')
        wH.tabla.column('fin', minwidth=100, width=100, anchor='center' )
        wH.tabla.column('informe', minwidth=100, width=200 , anchor='center')
        wH.tabla.column('responsable', minwidth=100, width=200, anchor='center')
        
       # Titulo
        wH.tabla.heading('#0', text='N° O/I', anchor ='e')
        wH.tabla.heading('descripcion', text='Descripción', anchor ='center')
        wH.tabla.heading('solicitante', text='Solicitante', anchor ='center')
        wH.tabla.heading('inicio', text='Fecha inicio', anchor ='center')
        wH.tabla.heading('fin', text='Fecha fin ', anchor ='center')
        wH.tabla.heading('informe', text='Informe', anchor ='center')
        wH.tabla.heading('responsable', text='Responsable', anchor ='center')
       

    ## PRIMER PISO ##########################
    F1 = 550 # altura fila 1
    F2 = 620
    # N° O/T:
    if tipo =='T':
        LH6 = tk.Label(wH, text = "N° O/T:")
    elif tipo =='I':
        LH6 = tk.Label(wH, text = "N° O/I:")
    else:
        pass # To Do: manejo de error de tipo
    LH6.place(x=20, y=F1)

    global EH6
    EH6 = tk.Entry(wH)
    EH6.place(x=90, y=F1)

# Nombre de la tarea
    LH1 = tk.Label(wH, text = "Descripción")
    LH1.place(x=250, y=F1)

    global EH1
    EH1 = tk.Entry(wH)
    EH1.place(x=325, y=F1, width=320)

    # Responsable
    LH3 = tk.Label(wH, text = "Responsable")
    LH3.place(x=20, y=F2)

    list = userList()
    global CH3
    CH3 = ttk.Combobox(wH, state="readonly", width = 17, values = list)
    CH3.place(x=110, y=F2)

# Fin  
    LH5 = tk.Label(wH, text = "Fin")
    LH5.place(x=270, y=F2)

    global EH5
    EH5 = tk.Entry(wH)
    EH5.place(x=320, y=F2)

# Informe
    LH4 = tk.Label(wH, text = "Informe")
    LH4.place(x=470, y=F2)

    global EH4
    EH4 = tk.Entry(wH)
    EH4.place(x=520, y=F2)

    ##################################################### 
    # Descripción (ATE)
    LH7 = tk.Label(wH, text = "Observaciones: ")
    LH7.place(x=700, y=F1)

    global TH1
    TH1 = tk.Text(wH, width = 61, height = 7)
    TH1.place(x=790, y=F1)
    

    #   BOTONES #####################
    # boton Precarga
    BH4=tk.Button(wH, text="Seleccionar", width=12, command=lambda:B_selecOrden(tipo))
    BH4.place(x=170, y=700)

    # boton Modificar
    global BH2
    BH2=tk.Button(wH,text="Modificar", width=12, state = tk.DISABLED, command=lambda:B_modOrden(tipo))
    BH2.place(x=340, y=700)
    
    # boton Ordenes de servicio
    BH3=tk.Button(wH, text="Ord. Servicio", width=12, command=lambda:V_modServ(tipo))
    BH3.place(x=510, y=700)

    # boton Muestras
    BH6=tk.Button(wH, text="Muestras", width=12, command=lambda:V_modMue(tipo))
    BH6.place(x=680, y=700)

    # boton Volver
    BH5=tk.Button(wH, text="Volver", width=12, command=lambda: menuFrom(wH))
    BH5.place(x=1020, y=700)

    # boton Salir
    BH1=tk.Button(wH, text="Salir", width=12, command = lambda: salirConAviso())
    BH1.place(x=1190, y=700)

    wH.tabla.delete(*wH.tabla.get_children())
    # preparacion de los datos
    refresh_conn(my_conn)
    if tipo == 'T':    
        try:
            my_cursor = my_conn.cursor()
            statement = "SELECT * FROM ordenes WHERE tipo = 'T' ORDER BY numero_cdf DESC LIMIT 50" # ordenes de trabajo
            my_cursor.execute(statement)
            resultados = my_cursor.fetchall()
            #print(resultados) 
        except Exception as e:
            print("error 3430", e)
        for fila in resultados:
            try:
                my_cursor = my_conn.cursor(buffered=True)
                statement = "SELECT filiacion_cliente, asunto FROM solicitudes WHERE ID_solicitud = %s" # ordenes de trabajo
                values = (fila[8],)
                my_cursor.execute(statement, values)
                solic = my_cursor.fetchone()
                #print(fila[8])
                #print(solic)
                #cliente = solic[0]
                #desc_solic = solic[1]
                
            except Exception as e:
                print("error 3435", e)
    # muestra de los datos
            wH.tabla.insert('', index = 'end', iid=None, text = str(fila[2]), values = [fila[4], fila[3], fila[6],fila[5], fila[8], solic[0], solic[1]])
    elif tipo == 'I':    
        try:
            my_cursor = my_conn.cursor()
            statement = "SELECT * FROM ordenes WHERE tipo = 'I' ORDER BY numero_cdf DESC LIMIT 50 " # ordenes internas
            my_cursor.execute(statement)
            resultados = my_cursor.fetchall()
            #print(resultados) 
        except Exception as e:
            print("error 3574)", e)
    # muestra de los datos
        for fila in resultados:
            wH.tabla.insert('',index = 'end', iid=None, text = str(fila[2]), values = [fila[4], fila[9], fila[10], fila[6], fila[5], fila[3]])
    else:
        pass # To Do: insertar manejo de error de tipo 
    wH.tabla.bind("<Double-1>", lambda evt, x = tipo: B_selecOrden(x)) # habilita doble click para seleccionar

def B_selecOrden(modo):
    curItem = wH.tabla.focus()
    numOrd = wH.tabla.item(curItem).get('text')
    if len(numOrd)!=0:
        refresh_conn(my_conn)
        try:
            my_cursor = my_conn.cursor()
            statement = "SELECT * FROM ordenes WHERE tipo = %s AND numero_cdf = %s"
            values = (modo, numOrd,) 
            my_cursor.execute(statement, values)
            resultado = my_cursor.fetchone()
            #print(resultado)
        except Exception as e:
            print("error", e)
            # Precarga de los tk.entry
        EH6.delete(0, tk.END)              # Numero de orden
        if resultado[2] is not None:
            EH6.insert(0, resultado[2])
        EH1.delete(0, tk.END)              # Descripción
        if resultado[4] is not None:
            EH1.insert(0, resultado[4])
        CH3.set('')                     # Responsable
        if resultado[3] is not None:
            CH3.set( resultado[3])
        EH5.delete(0, tk.END)              # Fecha finalización
        if resultado[6] is not None:
            EH5.insert(0, resultado[6])
        EH4.delete(0, tk.END)              # Informe
        if resultado[5] is not None:
            EH4.insert(0, resultado[5])
        TH1.delete("1.0","end")         # Observaciones
        if resultado[7] is not None:
            TH1.insert("1.0", resultado[7])
        BH2['state'] = tk.NORMAL           # habilitación del botón Modificar
            #####################################################################     
                     
    else:
        messagebox.showinfo(message="Seleccione una orden.", title="Aviso del sistema")

def B_modOrden(modo):
    curItem = wH.tabla.focus()
    numOrd = wH.tabla.item(curItem).get('text')
    refresh_conn(my_conn)
    if len(numOrd)!=0:
        ######### Actualización del numero de proyecto
        NnumOrd = EH6.get()
        if NnumOrd != numOrd:
            # chequea si el ID elegido existe
            try:
                my_cursor = my_conn.cursor()
                statement = "SELECT COUNT(numero_cdf) FROM ordenes WHERE tipo = %s and numero_cdf = %s "
                val = (modo, NnumOrd)
                my_cursor.execute(statement,val)
                flg_exists = my_cursor.fetchall()
                #print (results)
            except Exception as e:
                print("error 3687", e)
            if flg_exists[0][0]:
                messagebox.showinfo(message="El ID escogido ya existe .", title="Aviso del sistema")
                return(1)
            else:
                try:
                    my_cursor = my_conn.cursor()
                    statement = "UPDATE ordenes SET numero_cdf = %s WHERE tipo = %s and numero_cdf = %s"
                    val = (NnumOrd, modo, numOrd)
                    my_cursor.execute(statement,val)
                    my_conn.commit()
                    numOrd = NnumOrd       
                except Exception as e:
                    print("error", e)
        ######### Actualización de la descripción de la orden
        descOrd = EH1.get()
        if len(descOrd) != 0:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE ordenes SET descripcion = %s WHERE tipo = %s and numero_cdf = %s"
                val = (descOrd, modo, numOrd)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error", e)
        ######### Actualización del responsable
        respo = CH3.get()
        if len(respo) != 0:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE ordenes SET responsable = %s WHERE tipo= %s and numero_cdf = %s"
                val = (respo, modo, numOrd)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error", e)
        ######### Actualización de fecha de finalización
        fechaFin = EH5.get()
        if len(fechaFin) != 0:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE ordenes SET fecha_fin = %s WHERE tipo = %s and numero_cdf = %s"
                val = (fechaFin, modo, numOrd)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error", e)
        ######### Actualización del informe
        infor = EH4.get()
        if len(infor) != 0:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE ordenes SET informe = %s WHERE tipo = %s and numero_cdf = %s"
                val = (infor, modo, numOrd)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error", e)
        ######### Actualización de observaciones
        observ = TH1.get("1.0",'end-1c')
        if len(observ) != 0:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE ordenes SET observaciones = %s WHERE tipo =%s and numero_cdf = %s"
                val = (observ, modo, numOrd)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error", e)
        ############ Refresco de la ventana
        win_geometry = wH.winfo_geometry()
        wH.destroy() 
        V_modLab(modo, win_geometry)
    else:
        messagebox.showinfo(message="Seleccione una orden.", title="Aviso del sistema")
############################################# Ventana de modificación de ordenes de servicio #######################################################
def V_modServ(modo, geom = ''):
    curItem = wH.tabla.focus()
    numOrd = wH.tabla.item(curItem).get('text')
    if len(numOrd)!=0:
        ### esconde la pantalla de llamado
        wH.withdraw()
        global wI
        wI=tk.Toplevel()
        if len(geom)==0:
            wI.state('zoomed')
        else:
            wI.geometry(geom)
        wI.iconphoto(False, photo)
        wI.title("CENADIF Base de datos - Ordenes de servicio ")
        refresh_conn(my_conn)
        
        try:
            my_cursor = my_conn.cursor(buffered=True)            
            statement = "SELECT * FROM ordenes WHERE tipo = %s and numero_cdf = %s" 
            values = (modo, numOrd,)
            my_cursor.execute(statement, values)
            ordData = my_cursor.fetchone()
            #print(ordData)
            #print(ordData[0]) 
        except Exception as e:
            print("error 3624", e)

        try:
            my_cursor = my_conn.cursor(buffered=True)
            statement = "SELECT filiacion_cliente FROM solicitudes WHERE ID_solicitud = %s" # ordenes de trabajo
            values = (ordData[8],)
            my_cursor.execute(statement, values)
            cliente = my_cursor.fetchone()
            #print(cliente)
        except Exception as e:
            print("error 3634", e)

        wI.frame = tk.Frame(wI)
        wI.frame.grid(rowspan=2, column=1, row=1)
        wI.tabla = ttk.Treeview(wI.frame, height=1)
        wI.tabla.grid(column=1, row=1)

        ladox = tk.Scrollbar(wI.frame, orient = tk.VERTICAL, command= wI.tabla.yview)
        ladox.grid(column=0, row = 2, sticky='ew')
        #ladox.pack(side ='right', fill ='x') 
        ladoy = tk.Scrollbar(wI.frame, orient =tk.HORIZONTAL, command = wI.tabla.xview)
        ladoy.grid(column = 1, row = 0, sticky='ns')
        wI.tabla.configure(xscrollcommand = ladox.set, yscrollcommand = ladoy.set)

        wI.tabla['columns'] = ('descripcion', 'responsable', 'fin', 'informe', 'pedido','cliente')
        wI.tabla.column('#0', minwidth=50, width=60, anchor='center')
        wI.tabla.column('descripcion', minwidth=100, width=450, anchor='center')
        wI.tabla.column('responsable', minwidth=80, width=100 , anchor='center')
        wI.tabla.column('fin', minwidth=100, width=100 , anchor='center')
        wI.tabla.column('informe', minwidth=100, width=250, anchor='center' )
        wI.tabla.column('pedido', minwidth=100, width=150 , anchor='center')
        wI.tabla.column('cliente', minwidth=100, width=200, anchor='center') 

        # Titulo del número CENADIF y nombre (dependientes del modo)
        if modo == 'T':
            wI.tabla.heading('#0', text='N° O/T', anchor ='e')
        if modo == 'I':
            wI.tabla.heading('#0', text='N° O/I', anchor ='e')
        
        # Parte común a todos los modos
        wI.tabla.heading('descripcion', text='Descripción', anchor ='center')
        wI.tabla.heading('responsable', text='Responsable', anchor ='center')
        wI.tabla.heading('fin', text='Fecha fin', anchor ='center')
        wI.tabla.heading('informe', text='Informe', anchor ='center')
        wI.tabla.heading('pedido', text='Pedido', anchor ='center')
        wI.tabla.heading('cliente', text='Cliente', anchor ='center')

        # muestra de los datos
        wI.tabla.insert('',index = 'end', iid=None, text = str(ordData[2]), values = [ordData[4], ordData[3], ordData[6], ordData[5], ordData[8], cliente,])
        
        ### Tabla de ordenes de servicio ########################################################
        try:
            my_cursor = my_conn.cursor()
            statement = "SELECT * FROM ord_serv WHERE ID_orden = %s" 
            values = (ordData[0],)
            my_cursor.execute(statement, values)
            resultados = my_cursor.fetchall()
            #print(resultados) 
        except Exception as e:
            print("error 3683", e)

        wI.frame2 = tk.Frame(wI)
        wI.frame2.grid(rowspan=2, column=1, row=2)
        wI.frame2.place(y=70)
        wI.tabla2 = ttk.Treeview(wI.frame2, height=15)
        wI.tabla2.grid(column=1, row=1)

        ladox2 = tk.Scrollbar(wI.frame2, orient = tk.VERTICAL, command= wI.tabla2.yview)
        ladox2.grid(column=0, row = 1, sticky='ew') 
        ladoy2 = tk.Scrollbar(wI.frame2, orient =tk.HORIZONTAL, command = wI.tabla2.xview)
        ladoy2.grid(column = 1, row = 0, sticky='ns')
        #wG.tabla2.configure(xscrollcommand = ladox.set, yscrollcommand = ladoy.set)

        wI.tabla2['columns'] = ('tarea', 'tipo', 'responsable', 'operador', 'fecha_inicio', 'fecha_fin')
        wI.tabla2.column('#0', minwidth=50, width=60, anchor='w')
        wI.tabla2.column('tarea', minwidth=200, width=650, anchor='center')
        wI.tabla2.column('tipo', minwidth=80, width=100 , anchor='center')
        wI.tabla2.column('responsable', minwidth=100, width=150 , anchor='center')
        wI.tabla2.column('operador', minwidth=100, width=150, anchor='center' )
        wI.tabla2.column('fecha_inicio', minwidth=100, width=100 , anchor='center')
        wI.tabla2.column('fecha_fin', minwidth=100, width=100 , anchor='center')

        wI.tabla2.heading('#0', text='OS', anchor ='center')
        wI.tabla2.heading('tarea', text='Tarea', anchor ='center')
        wI.tabla2.heading('tipo', text='Tipo', anchor ='center')
        wI.tabla2.heading('responsable', text='Responsable', anchor ='center')
        wI.tabla2.heading('operador', text='Operador', anchor ='center')
        wI.tabla2.heading('fecha_inicio', text='Inicio', anchor ='center')
        wI.tabla2.heading('fecha_fin', text='Finaliz.', anchor ='center')
        
        try:
            for fila in resultados:
                wI.tabla2.insert('',index = fila[0], iid=None, text = str(fila[0]), values = [fila[3], fila[2], fila[1], fila[5], fila[6], fila[7]])        
        except:
            print("error 3722")
        #wF.after(1, lambda: wF.focus_force())

    ## ALTURAS
        F1 = 460
        F2 = 510
        F3 = 560
        F4 = 610
       
    ## PRIMER PISO ##########################

    # Nombre de la tarea
        LI1 = tk.Label(wI, text = "Tarea")
        LI1.place(x=20, y=F1)

        global EI1
        EI1 = tk.Entry(wI)
        EI1.place(x=110, y=F1, width=500)

    ## SEGUNDO PISO #########################
        # Responsable
        LI3 = tk.Label(wI, text = "Responsable")
        LI3.place(x=20, y=F2)

        list = userList()
        global CI3
        CI3 = ttk.Combobox(wI, state="readonly", width = 17, values = list)
        CI3.place(x=110, y=F2)

    # inicio
        LI4 = tk.Label(wI, text = "Inicio")
        LI4.place(x=250, y=F2)

        global EI4
        EI4 = tk.Entry(wI)
        EI4.place(x=300, y=F2)

    # Fin  
        LI5 = tk.Label(wI, text = "Fin")
        LI5.place(x=430, y=F2)

        global EI5
        EI5 = tk.Entry(wI)
        EI5.place(x=480, y=F2)

    ## TERCER PISO #########################
    
    # H/H prog
        LI7 = tk.Label(wI, text = "horas")
        LI7.place(x=250, y=F3)

        global EI7
        EI7 = tk.Entry(wI)
        EI7.place(x=300, y=F3)

        # Operador
        LI9 = tk.Label(wI, text = "Operador")
        LI9.place(x=20, y=F3)

        #global EG9
        #EG9 = tk.Entry(wG)
        #EG9.place(x=110, y=530)

        global CI9
        CI9 = ttk.Combobox(wI, state="readonly", width = 17, values = list)
        CI9.place(x=110, y=F3)
    # Tipo
        LI2 = tk.Label(wI, text = "Tipo")
        LI2.place(x=430, y=F3)

        global CI2
        CI2 = ttk.Combobox(wI, state="readonly", width = 17, values = ('Laboratorio', 'Taller'))
        CI2.place(x=480, y=F3)

    ## CUARTO PISO #########################
    # Equipos
        LIA = tk.Label(wI, text = "Equipamiento")
        LIA.place(x=20, y=F4)

        global EIA
        EIA = tk.Entry(wI)
        EIA.place(x=110, y=F4, width=500)
    ##################################################### 
        # Descripción (ATE)
        LIB = tk.Label(wI, text = "Observaciones: ")
        LIB.place(x=650, y=F2)

        global TI1
        TI1 = tk.Text(wI, width = 61, height = 7)
        TI1.place(x=740, y=F2)
     
    #   BOTONES #####################
        H_bot = 680
        # boton Precarga
        BI4=tk.Button(wI, text="Seleccionar", width=12, command = lambda: B_selServ(modo))
        BI4.place(x=115, y=H_bot)

        # boton Modificar
        global BI2
        BI2=tk.Button(wI,text="Modificar", width=12, command = lambda: B_modServ(modo), state = tk.DISABLED)
        BI2.place(x=245, y=H_bot)

        # boton Eliminar
        global BI6
        BI6=tk.Button(wI, text="Eliminar", width=12, command = lambda: B_elimServ(modo))
        BI6.place(x=375, y=H_bot)

        # boton Nueva
        BI3=tk.Button(wI, text="Nueva", width=12, command = lambda: B_nuevoServ(ordData[0], modo))
        BI3.place(x=505, y=H_bot)

        # boton Recursos
        BI7=tk.Button(wI, text="Recursos", width=12, command = lambda: V_empleos('o_serv'))
        BI7.place(x=635, y=H_bot)

        # boton Volver
        BI5=tk.Button(wI, text="Volver", width=12, command=lambda: OrdenesFrom(wI, modo))
        BI5.place(x=895, y=H_bot)

        # boton Salir
        BI1=tk.Button(wI, text="Salir", width=12, command = lambda: salirConAviso())
        BI1.place(x=1025, y=H_bot)

        wI.tabla2.bind("<Double-1>", lambda evt, x = modo: B_selServ(x)) # habilita doble click para seleccionar

    else:
        messagebox.showinfo(message="Seleccione una orden.", title="Aviso del sistema")

def B_selServ(modo):
    curItem = wI.tabla2.focus()
    numServ = wI.tabla2.item(curItem).get('text')
    if len(numServ)!=0:
        refresh_conn(my_conn)
        try:
            my_cursor = my_conn.cursor()
            statement = "SELECT * FROM ord_serv WHERE ID_servicio = %s"
            values = (numServ,) 
            my_cursor.execute(statement, values)
            resultado = my_cursor.fetchone()
        except Exception as e:
            print("error", e)
        # Precarga de los tk.entry
        EI1.delete(0, tk.END)              # Tarea
        if resultado[3] is not None:
            EI1.insert(0, resultado[3])
        CI3.set('')                     # Responsable
        if resultado[1] is not None:
            CI3.set( resultado[1])
        EI4.delete(0, tk.END)              # Inicio
        if resultado[6] is not None:
            EI4.insert(0, resultado[6])
        EI5.delete(0, tk.END)              # Fin
        if resultado[7] is not None:
            EI5.insert(0, resultado[7])
        EI7.delete(0, tk.END)              # Horas
        if resultado[8] is not None:
            EI7.insert(0, resultado[8])
        CI9.set('')                     # Operador
        if resultado[5] is not None:
            CI9.set( resultado[5])
        CI2.set('')                     # Tipo
        if resultado[2] is not None:
            CI2.set( resultado[2])
        EIA.delete(0, tk.END)              # Equipamiento
        if resultado[4] is not None:
            EIA.insert(0, resultado[4])
        TI1.delete("1.0","end")         # Descripción
        if resultado[9] is not None:
            TI1.insert("1.0", resultado[9])
        BI2['state'] = tk.NORMAL # habilitación del botón Modificar

    else:
        messagebox.showinfo(message="Seleccione una orden.", title="Aviso del sistema")

def B_modServ(modo): 
    curItem = wI.tabla2.focus()
    numServ = wI.tabla2.item(curItem).get('text')
    refresh_conn(my_conn)
    if len(numServ)!=0:
        ######### Actualización de la tarea
        tareaServ = EI1.get()
        if len(tareaServ) != 0:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE ord_serv SET tarea = %s WHERE ID_servicio = %s"
                val = (tareaServ, numServ)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error", e)
                ######### Actualización de la tarea
        respoServ = CI3.get()
        if len(respoServ) != 0:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE ord_serv SET responsable = %s WHERE ID_servicio = %s"
                val = (respoServ, numServ)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error", e)
        ######### Actualización de la fecha de inicio
        iniServ = EI4.get()
        if len(iniServ) != 0:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE ord_serv SET fecha_inicio = %s WHERE ID_servicio = %s"
                val = (iniServ, numServ)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error", e)
        ######### Actualización de la fecha de finalización
        finServ = EI5.get()
        if len(finServ) != 0:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE ord_serv SET fecha_fin = %s WHERE ID_servicio = %s"
                val = (finServ, numServ)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error", e)
        ######### Actualización de las horas
        horasServ = EI7.get()
        if len(horasServ) != 0:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE ord_serv SET horas = %s WHERE ID_servicio = %s"
                val = (horasServ, numServ)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error", e)
        ######### Actualización del operador
        operServ = CI9.get()
        if len(operServ) != 0:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE ord_serv SET operador = %s WHERE ID_servicio = %s"
                val = (operServ, numServ)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error", e)
        ######### Actualización del tipo
        tipoServ = CI2.get()
        if len(tipoServ) != 0:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE ord_serv SET tipo = %s WHERE ID_servicio = %s"
                val = (tipoServ, numServ)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error", e)
        ######### Actualización del equipamiento
        equipServ = EIA.get()
        if len(equipServ) != 0:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE ord_serv SET equipos = %s WHERE ID_servicio = %s"
                val = (equipServ, numServ)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error", e)
        ######### Actualización del equipamiento
        obsServ = TI1.get("1.0",'end-1c')
        if len(obsServ) != 0:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE ord_serv SET observaciones = %s WHERE ID_servicio = %s"
                val = (obsServ, numServ)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error", e)
        
        # TI1 Descripción
        win_geometry = wI.winfo_geometry()
        wI.destroy() 
        V_modServ(modo, win_geometry)

    else:
        messagebox.showinfo(message="Seleccione una orden de servicio.", title="Aviso del sistema")

def B_elimServ(modo):
    curItem = wI.tabla2.focus()
    numServ = wI.tabla2.item(curItem).get('text')
    if len(numServ)!=0:
        if (messagebox.askokcancel(message="Eliminar orden?", title="Confirmación de acción")):
            refresh_conn(my_conn)
            try:
                my_cursor = my_conn.cursor()
                statement = "DELETE FROM ord_serv WHERE ID_servicio = %s"
                val = (numServ,)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error", e)
            win_geometry = wI.winfo_geometry()
            wI.destroy() 
            V_modServ(modo, win_geometry)
        ############ Refresco de la ventana
    else:
        messagebox.showinfo(message="Seleccione una tarea.", title="Aviso del sistema")  

############################### Función del botón Nueva orden de servicio ############################################
def B_nuevoServ(orden, modo):
#    print('nueva tarea')
    refresh_conn(my_conn)
    try:
        my_cursor = my_conn.cursor()
        statement = "INSERT INTO ord_serv SET tarea = %s, ID_orden = %s"
        val = ('Nueva orden de servicio', orden)
        my_cursor.execute(statement,val)
        my_conn.commit()
    except Exception as e:
        print("error", e)
    win_geometry = wI.winfo_geometry()
    wI.destroy() 
    V_modServ(modo, win_geometry)
##########################################################################################################################################
def V_modMue(modo, geom = ''):
    curItem = wH.tabla.focus()
    numOrd = wH.tabla.item(curItem).get('text')
    if len(numOrd)!=0:
        ### esconde la pantalla de llamado
        wH.withdraw()
        global wJ
        wJ=tk.Toplevel()
        if len(geom)==0:
            wJ.state('zoomed')
        else:
            wJ.geometry(geom)
        wJ.iconphoto(False, photo)
        wJ.title("CENADIF Base de datos - Muestras ")
        refresh_conn(my_conn)

        try:
            my_cursor = my_conn.cursor(buffered=True)            
            statement = "SELECT * FROM ordenes WHERE tipo = %s and numero_cdf = %s" 
            values = (modo, numOrd,)
            my_cursor.execute(statement, values)
            ordData = my_cursor.fetchone()
            #print(ordData)
            #print(ordData[0]) 
        except Exception as e:
            print("error 3624", e)

        try:
            my_cursor = my_conn.cursor(buffered=True)
            statement = "SELECT filiacion_cliente FROM solicitudes WHERE ID_solicitud = %s" # ordenes de trabajo
            values = (ordData[4],)
            my_cursor.execute(statement, values)
            cliente = my_cursor.fetchone()
            #print(cliente)
        except Exception as e:
            print("error 3634", e)

        wJ.frame = tk.Frame(wJ)
        wJ.frame.grid(rowspan=2, column=1, row=1)
        wJ.tabla = ttk.Treeview(wJ.frame, height=1)
        wJ.tabla.grid(column=1, row=1)

        ladox = tk.Scrollbar(wJ.frame, orient = tk.VERTICAL, command= wJ.tabla.yview)
        ladox.grid(column=0, row = 2, sticky='ew')
        #ladox.pack(side ='right', fill ='x') 
        ladoy = tk.Scrollbar(wJ.frame, orient =tk.HORIZONTAL, command = wJ.tabla.xview)
        ladoy.grid(column = 1, row = 0, sticky='ns')
        wJ.tabla.configure(xscrollcommand = ladox.set, yscrollcommand = ladoy.set)

        wJ.tabla['columns'] = ('descripcion', 'responsable', 'fin', 'informe', 'pedido','cliente')
        wJ.tabla.column('#0', minwidth=50, width=60, anchor='center')
        wJ.tabla.column('descripcion', minwidth=100, width=450, anchor='center')
        wJ.tabla.column('responsable', minwidth=80, width=100 , anchor='center')
        wJ.tabla.column('fin', minwidth=100, width=100 , anchor='center')
        wJ.tabla.column('informe', minwidth=100, width=250, anchor='center' )
        wJ.tabla.column('pedido', minwidth=100, width=150 , anchor='center')
        wJ.tabla.column('cliente', minwidth=100, width=200, anchor='center') 

        # Titulo del número CENADIF y nombre (dependientes del modo)
        if modo == 'T':
            wJ.tabla.heading('#0', text='N° O/T', anchor ='e')
        if modo == 'I':
            wJ.tabla.heading('#0', text='N° O/I', anchor ='e')
        
        # Parte común a todos los modos
        wJ.tabla.heading('descripcion', text='Descripción', anchor ='center')
        wJ.tabla.heading('responsable', text='Responsable', anchor ='center')
        wJ.tabla.heading('fin', text='Fecha fin', anchor ='center')
        wJ.tabla.heading('informe', text='Informe', anchor ='center')
        wJ.tabla.heading('pedido', text='Pedido', anchor ='center')
        wJ.tabla.heading('cliente', text='Cliente', anchor ='center')

        # muestra de los datos
        wJ.tabla.insert('',index = ordData[2], iid=None, text = str(ordData[2]), values = [ordData[4], ordData[3], ordData[6], ordData[5], ordData[8], cliente,])

        ### Tabla de muestras ########################################################
        try:
            my_cursor = my_conn.cursor()
            statement = "SELECT * FROM muestras WHERE ID_orden = %s" 
            values = (ordData[0],)
            my_cursor.execute(statement, values)
            resultados = my_cursor.fetchall()
            #print(resultados) 
        except Exception as e:
            print("error 3683", e)

        wJ.frame2 = tk.Frame(wJ)
        wJ.frame2.grid(rowspan=2, column=1, row=2)
        wJ.frame2.place(y=70)
        wJ.tabla2 = ttk.Treeview(wJ.frame2, height=15)
        wJ.tabla2.grid(column=1, row=1)

        ladox2 = tk.Scrollbar(wJ.frame2, orient = tk.VERTICAL, command= wJ.tabla2.yview)
        ladox2.grid(column=0, row = 1, sticky='ew') 
        ladoy2 = tk.Scrollbar(wJ.frame2, orient =tk.HORIZONTAL, command = wJ.tabla2.xview)
        ladoy2.grid(column = 1, row = 0, sticky='ns')

        wJ.tabla2['columns'] = ('ID', 'corresp', 'observ', 'ingreso')
        wJ.tabla2.column('#0', minwidth=50, width=60, anchor='w')
        wJ.tabla2.column('ID', minwidth=100, width=150, anchor='center')
        wJ.tabla2.column('corresp', minwidth=150, width=300 , anchor='center')
        wJ.tabla2.column('observ', minwidth=100, width=650 , anchor='center')
        wJ.tabla2.column('ingreso', minwidth=100, width=150, anchor='center' )

        wJ.tabla2.heading('#0', text='Tipo', anchor ='center')
        wJ.tabla2.heading('ID', text='ID muestra', anchor ='center')
        wJ.tabla2.heading('corresp', text='Corresponde a', anchor ='center')
        wJ.tabla2.heading('observ', text='Observaciones', anchor ='center')
        wJ.tabla2.heading('ingreso', text='Ingresó', anchor ='center')
        
        try:
            for fila in resultados:
                #wJ.tabla2.insert('',index = fila[0], iid=None, text = str(fila[0]), values = [fila[3], fila[2], fila[1], fila[5], fila[6], fila[6]])
                wJ.tabla2.insert('',index = fila[0], iid=None, text = str(fila[5]), values = [fila[0], fila[3],fila[4], fila[2],])        
        except:
            print("error 4154")
    
    ## ALTURAS
        F1 = 450
       
    ## PRIMER PISO ##########################

    # Observaciones
        LJ1 = tk.Label(wJ, text = "Observaciones")
        LJ1.place(x=560, y=F1)

        global EJ1
        EJ1 = tk.Entry(wJ)
        EJ1.place(x=650, y=F1, width=500)

    ## SEGUNDO PISO #########################
        # Responsable
        LJ3 = tk.Label(wJ, text = "Tipo")
        LJ3.place(x=20, y=F1)

        list = ['MC', 'MP', 'MR', 'ML', 'MG', 'MM']
        global CJ3
        CJ3 = ttk.Combobox(wJ, state="readonly", width = 17, values = list)
        CJ3.place(x=80, y=F1)

    # inicio
        LJ4 = tk.Label(wJ, text = "Corresponde a:")
        LJ4.place(x=230, y=F1)

        global EJ4
        EJ4 = tk.Entry(wJ)
        EJ4.place(x=320, y=F1, width=230)

    # Fecha ingreso 
        LJ5 = tk.Label(wJ, text = "Ingresó")
        LJ5.place(x=1150, y=F1)

        global EJ5
        EJ5 = tk.Entry(wJ)
        EJ5.place(x=1200, y=F1)

    #   BOTONES #####################
        H_bot = 550
        # boton Precarga
        BJ4=tk.Button(wJ, text="Seleccionar", width=12, command = lambda: B_selMue())
        #BJ4=tk.Button(wJ, text="Seleccionar", width=12)
        BJ4.place(x=115, y=H_bot)

        # boton Modificar
        global BJ2
        BJ2=tk.Button(wJ,text="Modificar", width=12, command = lambda: B_modMue(modo), state = tk.DISABLED)
        #BJ2=tk.Button(wJ,text="Modificar", width=12)
        BJ2.place(x=245, y=H_bot)

        # boton Eliminar
        global BJ6
        BJ6=tk.Button(wJ, text="Eliminar", width=12, command = lambda: B_elimMue(modo))
        BJ6.place(x=375, y=H_bot)

        # boton Nueva
        BJ3=tk.Button(wJ, text="Nueva", width=12, command = lambda: B_nuevaMue(modo, ordData[0]))
        BJ3.place(x=505, y=H_bot)

        # boton Volver
        BJ5=tk.Button(wJ, text="Volver", width=12, command=lambda: OrdenesFrom(wJ, modo))
        BJ5.place(x=765, y=H_bot)

        # boton Salir
        BJ1=tk.Button(wJ, text="Salir", width=12, command = lambda: salirConAviso())
        BJ1.place(x=895, y=H_bot)

    else:
        messagebox.showinfo(message="Seleccione una orden.", title="Aviso del sistema")
################## Boton de selección de muestra #############################################################################
def B_selMue():
    curItem = wJ.tabla2.focus()
    numMue = wJ.tabla2.item(curItem).get('values')
    if len(numMue)!=0:
        refresh_conn(my_conn)
        try:
            my_cursor = my_conn.cursor()
            statement = "SELECT * FROM muestras WHERE ID_muestra = %s"
            values = (numMue[0],) 
            my_cursor.execute(statement, values)
            resultado = my_cursor.fetchone()
            #print(resultado[5])
        except Exception as e:
            print("error", e)
            # Precarga de los tk.entry
        CJ3.set('')                     # Tipo
        if resultado[5] is not None:
            CJ3.set( resultado[5])
        EJ1.delete(0, tk.END)              # Observaciones
        if resultado[4] is not None:
            EJ1.insert(0, resultado[4])
        EJ4.delete(0, tk.END)              # Correspondencia
        if resultado[3] is not None:
            EJ4.insert(0, resultado[3])
        EJ5.delete(0, tk.END)              # fecha de ingreso
        if resultado[2] is not None:
            EJ5.insert(0, resultado[2])
        BJ2['state'] = tk.NORMAL # habilitación del botón Modificar
    else:
        messagebox.showinfo(message="Seleccione una muestra.", title="Aviso del sistema")

############################### Función del botón Modificar muestra ############################################
def B_modMue(modo):
    curItem = wJ.tabla2.focus()
    numMue = wJ.tabla2.item(curItem).get('values')
    if len(numMue)!=0:
        refresh_conn(my_conn)
        ######### Actualización del tipo de muestra
        tipoMue = CJ3.get()
        if len(tipoMue) != 0:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE muestras SET tipo = %s WHERE ID_muestra = %s"
                val = (tipoMue, numMue[0])
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error", e)
        ######### Actualización de observaciones
        obsMue = EJ1.get()
        if len(obsMue) != 0:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE muestras SET observaciones = %s WHERE ID_muestra = %s"
                val = (obsMue, numMue[0])
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error", e)
        ######### Actualización de correspondencia
        corMue = EJ4.get()
        if len(corMue) != 0:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE muestras SET correspondencia = %s WHERE ID_muestra = %s"
                val = (corMue, numMue[0])
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error", e)
        ######### Actualización de fecha de ingreso
        fechaMue = EJ5.get()
        if len(fechaMue) != 0:
            if date_check(fechaMue) == True:
            #    return
                try:
                    my_cursor = my_conn.cursor()
                    statement = "UPDATE muestras SET ingreso = %s WHERE ID_muestra = %s"
                    val = (fechaMue, numMue[0])
                    my_cursor.execute(statement,val)
                    my_conn.commit()
                except Exception as e:
                    print("error", e)
                    ############ Refresco de la ventana
        win_geometry = wJ.winfo_geometry()
        wJ.destroy() 
        V_modMue(modo, win_geometry)
    else:
        messagebox.showinfo(message="Seleccione una muestra.", title="Aviso del sistema")

############################### Función del botón Nueva muestra ############################################
def B_nuevaMue(modo, orden):
    refresh_conn(my_conn)
    try:
        my_cursor = my_conn.cursor()
        #statement = "INSERT INTO muestras SET tipo = 'XX', observaciones = 'Nueva muestra', ID_orden = %s"
        statement = "INSERT INTO muestras (ID_orden, tipo, observaciones) VALUES(%s, %s, %s)"
        val = (orden, 'XX', 'Nueva muestra')
        my_cursor.execute(statement,val)
        my_conn.commit()
    except Exception as e:
        print("error", e)
    win_geometry = wJ.winfo_geometry()
    wJ.destroy() 
    V_modMue(modo, win_geometry)
############################### Función del botón Eliminar muestra  #############################################
def B_elimMue(modo):
    curItem = wJ.tabla2.focus()
    numMue = wJ.tabla2.item(curItem).get('values')
    if len(numMue)!=0:
        if (messagebox.askokcancel(message="Eliminar muestra?", title="Confirmación de acción")):
            refresh_conn(my_conn)
            try:
                my_cursor = my_conn.cursor()
                statement = "DELETE FROM muestras WHERE ID_muestra = %s"
                val = (numMue[0],)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error", e)
            win_geometry = wJ.winfo_geometry()
            wJ.destroy() 
            V_modMue(modo, win_geometry)
########################## Función de ingreso de orden interna de lab ###########################################################
def V_oIntLab():
    w1.withdraw()
    #print("Ingresando...")
    global wK
    wK=tk.Toplevel()
    wK.iconphoto(False, photo)
    wK.geometry("500x450")
    wK.title("CENADIF - Base de datos")

    LK1 = tk.Label(wK, text = "Ingreso de O/I - Laboratorio")
    LK1.place(x=190, y=25)
   
    F1 = 80
    F2 = 200
    F3 = 280
    F4 = 380
    C1 = 20
    C2 = 100

# Observaciones
    LK2 = tk.Label(wK, text = "Descripción")
    LK2.place(x=C1, y=F1)

    global TK2
    TK2 = tk.Text(wK, width = 42, height = 4)
    TK2.place(x= C2, y=F1)

## SEGUNDO PISO #########################
    # Responsable
    LK3 = tk.Label(wK, text = "Solicitante")
    LK3.place(x=C1, y=F2)

    list = userList()
    global CK3
    CK3 = ttk.Combobox(wK, state="readonly", width = 17, values = list)
    CK3.place(x=C2, y=F2)

# inicio
    LK4 = tk.Label(wK, text = "Fecha")
    LK4.place(x=C1, y=F3)

    global EK4
    EK4 = tk.Entry(wK)
    EK4.place(x=C2, y=F3)
# BOTONES #########################################################
    # boton Volver
    BK1=tk.Button(wK,text="Volver", command = lambda: menuFrom(wK))
    BK1.place(x=125, y=F4)
    
     # boton Ingresar
    BK2=tk.Button(wK,text="Ingresar", command=ingresar_oInt)
    #BK2=tk.Button(wK,text="Ingresar")
    BK2.place(x=325, y=F4)

    wK.after(1, lambda: wK.focus_force())

############# ventana de visualización de muestras ##########################################
def V_viewSamples():
    # Parte común a todos los modos
    w1.withdraw()
    global wB
    wB=tk.Toplevel()
    #wB.geometry("1240x650")
    wB.state("zoomed")
    wB.iconphoto(False, photo)
    wB.title("Laboratorio - Listado de muestras")
    wB.frame = tk.Frame(wB)
    wB.frame.grid(rowspan=2, column=1, row=1)
    wB.tabla = ttk.Treeview(wB.frame, height=33)
    wB.tabla.grid(column=1, row=1)
    ladox = tk.Scrollbar(wB.frame, orient = tk.VERTICAL, command= wB.tabla.yview)
    ladox.grid(column=0, row = 1, sticky='ew') 
    ladoy = tk.Scrollbar(wB.frame, orient =tk.HORIZONTAL, command = wB.tabla.yview)
    ladoy.grid(column = 1, row = 0, sticky='ns')

    wB.tabla.configure(xscrollcommand = ladox.set, yscrollcommand = ladoy.set)
    wB.tabla['columns'] = ('ID', 'tipo_ord', 'num_cdf', 'corresp', 'observ', 'ingreso')
    wB.tabla.column('#0', minwidth=50, width=50, anchor='w')
    wB.tabla.column('ID', minwidth=50, width=70, anchor='center')
    wB.tabla.column('tipo_ord', minwidth=50, width=70, anchor='center')
    wB.tabla.column('num_cdf', minwidth=50, width=70, anchor='center')
    wB.tabla.column('corresp', minwidth=150, width=350 , anchor='center')
    wB.tabla.column('observ', minwidth=100, width=650 , anchor='center')
    wB.tabla.column('ingreso', minwidth=100, width=150, anchor='center' )

    wB.tabla.heading('#0', text='Tipo', anchor ='center')
    wB.tabla.heading('ID', text='ID muestra', anchor ='center')
    wB.tabla.heading('tipo_ord', text='Tipo_orden', anchor ='center')
    wB.tabla.heading('num_cdf', text='Num_orden', anchor ='center')
    wB.tabla.heading('corresp', text='Corresponde a', anchor ='center')
    wB.tabla.heading('observ', text='Observaciones', anchor ='center')
    wB.tabla.heading('ingreso', text='Ingresó', anchor ='center')

    # boton Salir
    BB1=tk.Button(wB,text="Salir", width=12, command = lambda: salirConAviso())
    BB1.place(x=850, y=740)

    # boton Volver
    BB3=tk.Button(wB,text="Volver", width=12, command = lambda: menuFrom(wB))
    BB3.place(x=620, y=740)

    wB.tabla.delete(*wB.tabla.get_children())

    refresh_conn(my_conn)
    
    try:
        my_cursor = my_conn.cursor()
        statement = "SELECT ordenes.ID_orden, ordenes.tipo, ordenes.numero_cdf, muestras.ID_muestra, muestras.ingreso, muestras.correspondencia, muestras.observaciones, muestras.tipo FROM ordenes INNER JOIN muestras ON ordenes.ID_orden = muestras.ID_orden ORDER BY muestras.ID_muestra DESC"        
        my_cursor.execute(statement)
        resultados = my_cursor.fetchall()
        #print(resultados) 
    except Exception as e:
        print("error", e)
    
    # Parte común a todos los modos
    for fila in resultados:
        wB.tabla.insert('',index = 'end', iid=None, text = str(fila[7]), values = [fila[3], fila[1], fila[2], fila[5], fila[6], fila[4]])

    wB.after(1, lambda: wB.focus_force())
    
##################################### Chequeo de  formato de fecha ###########################################################
def date_check(date_req):
    try:
        year, month, day = date_req.split('-')
        year = year # To Do insertar chequeo de formato de año
        month = month # To Do insertar chequeo de formato de mes
        day = day # To Do insertar chequeo de formato de día
        return True

    except:
        if (messagebox.showinfo(message="Formato de fecha invalido \n Utilizar AAAA-MM-DD (ISO 8601)", title="Aviso del sistema")):
            return False
        
########################### Boton ingresar orden interna ###############################################################         
def ingresar_oInt():
    # chequeo de descripción
    desc = TK2.get("1.0",'end-1c')
    if len(desc) == 0:
        if (messagebox.showinfo(message="El campo Descripción no puede quedar vacío.", title="Aviso del sistema")):
            return
    # chequeo de solicitante
    solic = CK3.get()
    if len(solic) == 0:
        if (messagebox.showinfo(message="El campo Solicitante no puede quedar vacío.", title="Aviso del sistema")):
            return
    # chequeo de fecha
    fecha = EK4.get()
    if date_check(fecha) == False:
        return
    # update de BDNP
    refresh_conn(my_conn)
    try:
        my_cursor = my_conn.cursor()
        statement = "SELECT numero_cdf FROM ordenes WHERE tipo = 'I'"
        #val = ('I')
        #my_cursor.execute(statement,val)
        my_cursor.execute(statement)
        num_OI = my_cursor.fetchall()
        #print(num_OI)
    except Exception as e:
        print("error 4497", e)

    max = 0
    for n in num_OI:
        if n[0] > max:
            max = n[0]
    if max < 10000:
        max = 10000
    elif max >= 10000:
        max += 1
    else:
        pass # To Do manejo de error

    try:
        my_cursor = my_conn.cursor()
        statement = "INSERT INTO ordenes (tipo, descripcion, solicitante, fecha_inicio, numero_cdf) VALUES(%s, %s, %s, %s, %s )"
        val = ('I', desc, solic, fecha, max)
        my_cursor.execute(statement,val)
        my_conn.commit()
    except Exception as e:
        print("error 4506", e)
    win_geometry = wK.winfo_geometry()
    wK.destroy() 
    V_modLab('I', win_geometry)
############### Función de boton Salir #############################################################
def salirConAviso():
    if (messagebox.askokcancel(message="Desea cerrar el programa?", title="Confirmación de acción")):   
        my_conn.close()
        w1.quit()
    else:
        return
############# Ventana de gestión de empleos ##########################################
def V_empleos(modo, geom = ''):
    if modo == 'tar':
        tabla = wG.tabla2
    elif modo == 'doc':
        tabla = w8.tabla
    elif modo == 'o_serv':
        tabla = wI.tabla2
    else:
        print('error de modo')
        return
    curItem = tabla.focus()
    numTar = tabla.item(curItem).get('text')
    ### Cierre  de la ventana llamante
    if len(numTar)!=0:
        try:
            wG.withdraw()
        except:
            pass
        try:
            w8.withdraw()
        except:
            pass
        try:
            wI.withdraw()
        except:
            pass
        ### Preparación de datos para tabla 1
        refresh_conn(my_conn)
        global wL
        if modo == 'tar':
            try:
                my_cursor = my_conn.cursor(buffered=True)
                statement = "SELECT proyectos.area, proyectos.numero_cdf, tareas.nombre_tarea FROM proyectos INNER JOIN tareas ON proyectos.ID_proyecto = tareas.ID_proyecto WHERE ID_tarea = %s" 
                values = (numTar,)
                my_cursor.execute(statement, values)
                tarData = my_cursor.fetchone()
                #print(tarData) 
            except Exception as e:
                print("error 4755", e)
            try:
                my_cursor = my_conn.cursor(buffered=True)
                statement = "SELECT * FROM empleos WHERE tipo_contexto = %s and ID_contexto = %s" 
                values = (modo, numTar,)
                my_cursor.execute(statement, values)
                resultados = my_cursor.fetchall()
                #print(proyData[0]) 
            except Exception as e:
                print("error 4761", e)
            ### Generación de la ventana 
            wL=tk.Toplevel()
            if len(geom)==0:
                wL.state('zoomed')
            else:
                wL.geometry(geom)
            wL.iconphoto(False, photo)
            wL.title("CENADIF Base de datos - Recursos ")
            ### Tabla 1 (modo doc)
            wL.frame = tk.Frame(wL)
            wL.frame.grid(rowspan=2, column=1, row=1)
            wL.tabla = ttk.Treeview(wL.frame, height=1)
            wL.tabla.grid(column=1, row=1)

            ladox = tk.Scrollbar(wL.frame, orient = tk.VERTICAL, command= wL.tabla.yview)
            ladox.grid(column=0, row = 2, sticky='ew')
            ladoy = tk.Scrollbar(wL.frame, orient =tk.HORIZONTAL, command = wL.tabla.xview)
            ladoy.grid(column = 1, row = 0, sticky='ns')
            wL.tabla.configure(xscrollcommand = ladox.set, yscrollcommand = ladoy.set)

            wL.tabla['columns'] = ('tipo', 'ID', 'tarea')
            wL.tabla.column('#0', minwidth=10, width=10, anchor='center')
            wL.tabla.column('tipo', minwidth=100, width=100 , anchor='center')
            wL.tabla.column('ID', minwidth=100, width=100 , anchor='center')
            wL.tabla.column('tarea', minwidth=200, width=240, anchor='center' )
            
            #wL.tabla.heading('#0', text='N° DyN', anchor ='e')
            wL.tabla.heading('tipo', text='Tipo', anchor ='center')
            wL.tabla.heading('ID', text='ID', anchor ='center')
            wL.tabla.heading('tarea', text='Tarea', anchor ='center')
             
            wL.tabla.insert('',index = 'end', iid=None, values = [tarData[0], tarData[1], tarData[2]])

        elif modo == 'doc':
            try:
                my_cursor = my_conn.cursor(buffered=True)
                statement = "SELECT * FROM documentacion WHERE ID_trabajo = %s" 
                values = (numTar,)
                my_cursor.execute(statement, values)
                docData = my_cursor.fetchone()
                #print(docData) 
            except Exception as e:
                print("error 4755", e)
            try:
                my_cursor = my_conn.cursor(buffered=True)
                statement = "SELECT * FROM empleos WHERE tipo_contexto = %s and ID_contexto = %s" 
                values = (modo, numTar,)
                my_cursor.execute(statement, values)
                resultados = my_cursor.fetchall()
                #print(proyData[0]) 
            except Exception as e:
                print("error 4761", e)
            ### Generación de la ventana 
            wL=tk.Toplevel()
            if len(geom)==0:
                wL.state('zoomed')
            else:
                wL.geometry(geom)
            wL.iconphoto(False, photo)
            wL.title("CENADIF Base de datos - Recursos ")
            ### Tabla 1 (modo doc)
            wL.frame = tk.Frame(wL)
            wL.frame.grid(rowspan=2, column=1, row=1)
            wL.tabla = ttk.Treeview(wL.frame, height=1)
            wL.tabla.grid(column=1, row=1)

            ladox = tk.Scrollbar(wL.frame, orient = tk.VERTICAL, command= wL.tabla.yview)
            ladox.grid(column=0, row = 2, sticky='ew')
            ladoy = tk.Scrollbar(wL.frame, orient =tk.HORIZONTAL, command = wL.tabla.xview)
            ladoy.grid(column = 1, row = 0, sticky='ns')
            wL.tabla.configure(xscrollcommand = ladox.set, yscrollcommand = ladoy.set)

            wL.tabla['columns'] = ('tipo_doc', 'emisor', 'nro_doc', 'titulo', 'fuente','precio', 'u_monet','cantidad')
            wL.tabla.column('#0', minwidth=50, width=50, anchor='center')
            wL.tabla.column('tipo_doc', minwidth=100, width=100 , anchor='center')
            wL.tabla.column('emisor', minwidth=100, width=100 , anchor='center')
            wL.tabla.column('nro_doc', minwidth=100, width=120, anchor='center' )
            wL.tabla.column('titulo', minwidth=100, width=350 , anchor='w')
            wL.tabla.column('fuente', minwidth=100, width=100, anchor='center')
            wL.tabla.column('precio', minwidth=100, width=100, anchor='center')
            wL.tabla.column('u_monet', minwidth=100, width=100, anchor='center')
            wL.tabla.column('cantidad', minwidth=100, width=100, anchor='center')

            wL.tabla.heading('#0', text='N° DyN', anchor ='e')
            wL.tabla.heading('tipo_doc', text='Tipo', anchor ='center')
            wL.tabla.heading('emisor', text='Emisor', anchor ='center')
            wL.tabla.heading('nro_doc', text='Número', anchor ='center')
            wL.tabla.heading('titulo', text='Título', anchor ='center')
            wL.tabla.heading('fuente', text='Fuente', anchor ='center')
            wL.tabla.heading('precio', text='Precio', anchor ='center')
            wL.tabla.heading('u_monet', text='Un. Mon.', anchor ='center')
            wL.tabla.heading('cantidad', text='Cant.', anchor ='center') 
            wL.tabla.insert('',index = 'end', iid=None, text = str(docData[0]), values = [docData[2], docData[8], docData[3], docData[4], docData[5], docData[6], docData[7], docData[11]])
        elif modo == 'o_serv':
            try:
                my_cursor = my_conn.cursor(buffered=True)
                statement = "SELECT * FROM ord_serv WHERE ID_servicio = %s" 
                values = (numTar,)
                my_cursor.execute(statement, values)
                servData = my_cursor.fetchone()
                #print(servData) 
            except Exception as e:
                print("error 4755", e)
            try:
                my_cursor = my_conn.cursor(buffered=True)
                statement = "SELECT * FROM empleos WHERE tipo_contexto = %s and ID_contexto = %s" 
                values = (modo, numTar,)
                my_cursor.execute(statement, values)
                resultados = my_cursor.fetchall()
                #print(proyData[0]) 
            except Exception as e:
                print("error 4761", e)
            ### Generación de la ventana 
            wL=tk.Toplevel()
            if len(geom)==0:
                wL.state('zoomed')
            else:
                wL.geometry(geom)
            wL.iconphoto(False, photo)
            wL.title("CENADIF Base de datos - Recursos ")
            ### Tabla 1 (modo doc)
            wL.frame = tk.Frame(wL)
            wL.frame.grid(rowspan=2, column=1, row=1)
            wL.tabla = ttk.Treeview(wL.frame, height=1)
            wL.tabla.grid(column=1, row=1)

            ladox = tk.Scrollbar(wL.frame, orient = tk.VERTICAL, command= wL.tabla.yview)
            ladox.grid(column=0, row = 2, sticky='ew')
            ladoy = tk.Scrollbar(wL.frame, orient =tk.HORIZONTAL, command = wL.tabla.xview)
            ladoy.grid(column = 1, row = 0, sticky='ns')
            wL.tabla.configure(xscrollcommand = ladox.set, yscrollcommand = ladoy.set)
 
            wL.tabla['columns'] = ('tarea', 'tipo', 'responsable', 'operador', 'fecha_inicio', 'fecha_fin')
            wL.tabla.column('#0', minwidth=50, width=60, anchor='w')
            wL.tabla.column('tarea', minwidth=200, width=650, anchor='center')
            wL.tabla.column('tipo', minwidth=80, width=100 , anchor='center')
            wL.tabla.column('responsable', minwidth=100, width=150 , anchor='center')
            wL.tabla.column('operador', minwidth=100, width=150, anchor='center' )
            wL.tabla.column('fecha_inicio', minwidth=100, width=100 , anchor='center')
            wL.tabla.column('fecha_fin', minwidth=100, width=100 , anchor='center')

            wL.tabla.heading('#0', text='OS', anchor ='center')
            wL.tabla.heading('tarea', text='Tarea', anchor ='center')
            wL.tabla.heading('tipo', text='Tipo', anchor ='center')
            wL.tabla.heading('responsable', text='Responsable', anchor ='center')
            wL.tabla.heading('operador', text='Operador', anchor ='center')
            wL.tabla.heading('fecha_inicio', text='Inicio', anchor ='center')
            wL.tabla.heading('fecha_fin', text='Finaliz.', anchor ='center')

            wL.tabla.insert('',index = 'end', iid=None, text = str(servData[0]), values = [servData[3], servData[2], servData[1], servData[5], servData[6], servData[7]])        
           
        else:
            print('error de modo')
            return           

        ### Tabla 2
        wL.frame2 = tk.Frame(wL)
        wL.frame2.grid(rowspan=2, column=1, row=2)
        wL.frame2.place(y=70)
        wL.tabla2 = ttk.Treeview(wL.frame2, height=12)
        wL.tabla2.grid(column=1, row=1)

        ladox2 = tk.Scrollbar(wL.frame2, orient = tk.VERTICAL, command= wL.tabla2.yview)
        ladox2.grid(column=0, row = 1, sticky='ew') 
        ladoy2 = tk.Scrollbar(wL.frame2, orient =tk.HORIZONTAL, command = wL.tabla2.xview)
        ladoy2.grid(column = 1, row = 0, sticky='ns')
        #wG.tabla2.configure(xscrollcommand = ladox.set, yscrollcommand = ladoy.set)

        wL.tabla2['columns'] = ('tipo', 'categoria', 'cantidad', 'unidad', 'detalle', 'fecha', 'responsable')
        wL.tabla2.column('#0', minwidth=70, width=70, anchor='center')
        wL.tabla2.column('tipo', minwidth=100, width=150, anchor='center')
        wL.tabla2.column('categoria', minwidth=80, width=200 , anchor='center')
        wL.tabla2.column('cantidad', minwidth=100, width=100 , anchor='center')
        wL.tabla2.column('unidad', minwidth=100, width=100, anchor='center' )
        wL.tabla2.column('detalle', minwidth=100, width=550 , anchor='center')
        wL.tabla2.column('fecha', minwidth=100, width=150 , anchor='center')
        wL.tabla2.column('responsable', minwidth=100, width=150 , anchor='center')

        wL.tabla2.heading('#0', text='ID_empleo', anchor ='e')
        wL.tabla2.heading('tipo', text='Tipo recurso', anchor ='center')
        wL.tabla2.heading('categoria', text='Categoria', anchor ='center')
        wL.tabla2.heading('cantidad', text='Cantidad', anchor ='center')
        wL.tabla2.heading('unidad', text='Unidad med.', anchor ='center')
        wL.tabla2.heading('detalle', text='Detalle', anchor ='center')
        wL.tabla2.heading('fecha', text='Fecha', anchor ='center')
        wL.tabla2.heading('responsable', text='Responsable', anchor ='center')
        for fila in resultados:
            wL.tabla2.insert('',index = 'end', iid=None, text = str(fila[0]), values = [fila[3], fila[4], fila[5], fila[6], fila[7], fila[8], fila[9]])
       
        # boton Volver
        BL2=tk.Button(wL, text="Volver", width=12, command=lambda: menuFrom(wL))
        BL2.place(x=375, y=750)
        # boton Salir
        BL1=tk.Button(wL, text="Salir", width=12, command = lambda: salirConAviso())
        BL1.place(x=505, y=750)
        # boton Eliminar
        global BL3
        BL3=tk.Button(wL, text="Eliminar", width=12, command = lambda: B_elimEmpleo(modo))
        BL3.place(x=375, y=705)
        # boton Nueva
        BL4=tk.Button(wL, text="Nuevo", width=12, command = lambda: B_nuevoEmpleo(modo, numTar))
        BL4.place(x=505, y=705)
        
    #   BOTONES #####################
        # boton Precarga
        global BL5
        BL5=tk.Button(wL, text="Seleccionar", width=12, command = lambda: B_selEmpleo(modo))
        BL5.place(x=115, y=705)

        # boton Modificar
        global BL6
        BL6=tk.Button(wL,text="Modificar", width=12, command = lambda: B_modEmpleo(modo, numTar), state = tk.DISABLED)
        BL6.place(x=245, y=705)
    ## PRIMER PISO ##########################
    # Tipo de recurso
        LL1 = tk.Label(wL, text = "Tipo")
        LL1.place(x=20, y=380)

        refresh_conn(my_conn)
        try:
            my_cursor = my_conn.cursor()
            statement = "SELECT DISTINCT tipo FROM recursos "
            my_cursor.execute(statement)
            list = my_cursor.fetchall()
            #print(list) 
        except Exception as e:
            print("error 4994", e)
        # registro del observador
        rec_tipo = tk.StringVar()
        rec_tipo.trace('w', lambda *_ , var = rec_tipo : traceCB1(*_,var=var))
        global CL1
        CL1 = ttk.Combobox(wL, state="readonly", textvariable = rec_tipo, width = 17, values = list)
        CL1.place(x=110, y=380)

    # Categoría
        LL2 = tk.Label(wL, text = "Categoría")
        LL2.place(x=390, y=380)

        rec_cat = tk.StringVar()
        rec_cat.trace('w', lambda *_ , var = rec_cat : traceCB2(*_,var=var))

        global CL2
        CL2 = ttk.Combobox(wL, state="readonly", textvariable = rec_cat, width = 17)
        CL2.place(x=480, y=380)

        ## SEGUNDO PISO #########################
        # Cantidad
        LL3 = tk.Label(wL, text = "Cantidad")
        LL3.place(x=20, y=430)

        global EL3
        EL3 = tk.Entry(wL)
        EL3.place(x=110, y=430)

    # Unidad de medida 
        LL5 = tk.Label(wL, text = "Un. Medida")
        LL5.place(x=390, y=430)

        global EL5
        EL5 = tk.Entry(wL)
        EL5.place(x=480, y=430)
        
        ## TERCER PISO #########################
        
    # Detalle
        LL6 = tk.Label(wL, text = "Detalle: ")
        LL6.place(x=20, y=480)
        global TL6
        TL6 = tk.Text(wL, width = 61, height = 3)
        TL6.place(x=110, y=480)
        
        
    ## TERCER PISO #########################
        # Fecha
        LL7 = tk.Label(wL, text = "Fecha")
        LL7.place(x=20, y=560)

        global EL7
        EL7 = tk.Entry(wL)
        EL7.place(x=110, y=560)

    # Responsable
        LL8 = tk.Label(wL, text = "Responsable")
        LL8.place(x=390, y=560)

        list = userList()
        global CL8
        CL8 = ttk.Combobox(wL, state="readonly", width = 17, values = list)
        CL8.place(x=480, y=560)

    else:
        if modo == 'tar':
            messagebox.showinfo(message="Seleccione una tarea.", title="Aviso del sistema")
        elif modo == 'doc':
            messagebox.showinfo(message="Seleccione un documento.", title="Aviso del sistema")
        elif modo == 'o_serv':
            messagebox.showinfo(message="Seleccione una orden.", title="Aviso del sistema")
        else:
            print('error de modo')
            return

    #else:
    #    print('error de modo ')
########## Trace de tipo ##################################################################
def traceCB1(*args, var):
    # busca la lista de tipos de recurso
    tipoRec = var.get()
    refresh_conn(my_conn)
    try:
        my_cursor = my_conn.cursor()
        statement = "SELECT DISTINCT categoria FROM recursos WHERE tipo = %s"
        values = (tipoRec,)
        my_cursor.execute(statement, values)
        list = my_cursor.fetchall()
        #print(list) 
    except Exception as e:
        print("error 5094", e)
    # pone la lista en el combobox si existe
    try:    
        CL2['values'] = list
    except:
        pass
    try:
        CM2['values'] = list
    except:
        pass
    
########### Trace de categoria ###########################################################
def traceCB2(*args, var):
    catRec = var.get()
    #print(catRec)
    refresh_conn(my_conn)
    try:
        my_cursor = my_conn.cursor()
        statement = "SELECT unidad_medida FROM recursos WHERE categoria = %s"
        values = (catRec,)
        my_cursor.execute(statement, values)
        uniMed = my_cursor.fetchone()
        #print(uniMed)
        ### Solo busca por categoria, si existiesen dos tipos de recursos con una categoria que se llame igual generaría conflicto #### To Do 
    except Exception as e:
        print("error 5119", e)
    
    try:
        EL5.delete(0, 'end')
        if uniMed is not None:
            EL5.insert(0, uniMed[0])
    except:
        pass

############# Botón nuevo empleo #########################################################
def B_nuevoEmpleo(tipo, ID):
    try:
        my_cursor = my_conn.cursor()
        statement = "INSERT INTO empleos (tipo_contexto, ID_contexto, fecha) VALUES(%s, %s, %s )"
        val = (tipo, ID, date.today())
        my_cursor.execute(statement,val)
        my_conn.commit()
    except Exception as e:
        print("error 5109", e)
    win_geometry = wL.winfo_geometry()
    wL.destroy() 
    V_empleos(tipo, win_geometry)   

############# Botón eliminar empleo #####################################################
def B_elimEmpleo(tipo):
    curItem = wL.tabla2.focus()
    numEmp = wL.tabla2.item(curItem).get('text')
    #print(numEmp)
    if len(numEmp)!=0:
        if (messagebox.askokcancel(message="Eliminar recurso?", title="Confirmación de acción")):
            refresh_conn(my_conn)
            try:
                my_cursor = my_conn.cursor()
                statement = "DELETE FROM empleos WHERE ID_empleo = %s"
                val = (numEmp,)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error 5127", e)
            win_geometry = wL.winfo_geometry()
            wL.destroy() 
            V_empleos(tipo, win_geometry)
    else:
        messagebox.showinfo(message="Seleccione un recurso.", title="Aviso del sistema")   

############# Botón seleccionar empleo #####################################################
def B_selEmpleo(modo):
    curItem = wL.tabla2.focus()
    numEmp = wL.tabla2.item(curItem).get('text')
    if len(numEmp)!=0:
        #gantt(numpro, wF)
        refresh_conn(my_conn)
        try:
            my_cursor = my_conn.cursor()
            statement = "SELECT * FROM empleos WHERE ID_empleo = %s"
            values = (numEmp,) 
            my_cursor.execute(statement, values)
            resultado = my_cursor.fetchone()
            # Precarga de los tk.entry
            CL1.set('')                     # Tipo de recurso
            if resultado[3] is not None:
                CL1.set( resultado[3])
            CL2.set('')                     # Categoría
            if resultado[4] is not None:
                CL2.set( resultado[4])
            EL3.delete(0, tk.END)           # Cantidad
            if resultado[5] is not None:
                EL3.insert(0, resultado[5])
            EL5.delete(0, tk.END)              # Unidad de medida
            if resultado[6] is not None:
                EL5.insert(0, resultado[6])
            TL6.delete("1.0","end")             # Descripción
            if resultado[7] is not None:
                TL6.insert("1.0", resultado[7])
            EL7.delete(0, tk.END)              # Fecha
            if resultado[8] is not None:
                EL7.insert(0, resultado[8])
            CL8.set('')                     # Responsable
            if resultado[9] is not None:
                CL8.set(resultado[9])
            BL6['state'] = tk.NORMAL # habilitación del botón Eliminar
        except Exception as e:
            print("error 5173", e)
    else:
        messagebox.showinfo(message="Seleccione un recurso.", title="Aviso del sistema") 

############################### Función del botón Modificar empleo ############################################
def B_modEmpleo(modo, ID):
    curItem = wL.tabla2.focus()
    numEmp = wL.tabla2.item(curItem).get('text')
    if len(numEmp)!=0:
        refresh_conn(my_conn)
        #print(modo)
        #print(ID)
        #print(numEmp)
        ######### Actualización del tipo de recurso
        tipo = CL1.get()
        if len(tipo) != 0:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE empleos SET tipo = %s WHERE ID_empleo = %s"
                val = (tipo, numEmp)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error 5199", e)
        ######### Actualización de la categoria del recurso
        categoria = CL2.get()
        if len(categoria) != 0:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE empleos SET categoria = %s WHERE ID_empleo = %s"
                val = (categoria, numEmp)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error 5210", e)
        ######### Actualización de la cantidad
        cantidad = EL3.get()
        if len(cantidad) != 0:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE empleos SET cantidad = %s WHERE ID_empleo = %s"
                val = (cantidad, numEmp)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error 5221", e)
        ######### Actualización de la unidad de medida
        uniMedida = EL5.get()
        if len(uniMedida) != 0:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE empleos SET unidad_medida = %s WHERE ID_empleo = %s"
                val = (uniMedida, numEmp)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error 5232", e)
        ######### Actualización de detalle
        detalle = TL6.get("1.0",'end-1c')
        if len(detalle) != 0:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE empleos SET detalle = %s WHERE ID_empleo = %s"
                val = (detalle, numEmp)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error 5243", e)
        ######### Actualización de la fecha
        fecha = EL7.get()
        if len(fecha) != 0:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE empleos SET fecha = %s WHERE ID_empleo = %s"
                val = (fecha, numEmp)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error 5254", e)
        ######### Actualización del responsable
        responsable = CL8.get()
        if len(responsable) != 0:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE empleos SET responsable = %s WHERE ID_empleo = %s"
                val = (responsable, numEmp)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error 5265", e)
        ############ Refresco de la ventana
        win_geometry = wL.winfo_geometry()
        wL.destroy()
        V_empleos(modo, win_geometry)
    else:
        messagebox.showinfo(message="Seleccione un recurso.", title="Aviso del sistema")
### Ventana de gestión de recursos #####################################################################################
def V_recursos( geom = ''):
#    curItem = wH.tabla.focus()
#    numOrd = wH.tabla.item(curItem).get('text')
#    if len(numOrd)!=0:
    ### esconde la pantalla de llamado
    try:
        w1.withdraw()
    except:
        pass
    global wM
    wM=tk.Toplevel()
    if len(geom)==0:
        wM.state('zoomed')
    else:
        wM.geometry(geom)
    wM.iconphoto(False, photo)
    wM.title("CENADIF Base de datos - Gestión de recursos ")
    refresh_conn(my_conn)
    try:
        my_cursor = my_conn.cursor(buffered=True)            
        statement = "SELECT * FROM recursos" 
        my_cursor.execute(statement)
        resultados = my_cursor.fetchall()
        #print(resultados) 
    except Exception as e:
        print("error 5312", e)

    wM.frame = tk.Frame(wM)
    wM.frame.grid(rowspan=2, column=1, row=1)
    wM.tabla = ttk.Treeview(wM.frame, height=21)
    wM.tabla.grid(column=1, row=1)

    ladox = tk.Scrollbar(wM.frame, orient = tk.VERTICAL, command= wM.tabla.yview)
    ladox.grid(column=0, row = 2, sticky='ew')
    #ladox.pack(side ='right', fill ='x') 
    ladoy = tk.Scrollbar(wM.frame, orient =tk.HORIZONTAL, command = wM.tabla.xview)
    ladoy.grid(column = 1, row = 0, sticky='ns')
    wM.tabla.configure(xscrollcommand = ladox.set, yscrollcommand = ladoy.set)

    wM.tabla['columns'] = ('tipo', 'categoria', 'precio_unitario', 'unidad_monetaria', 'validez','unidad_medida')
    wM.tabla.column('#0', minwidth=50, width=60, anchor='center')
    wM.tabla.column('tipo', minwidth=100, width=225, anchor='center')
    wM.tabla.column('categoria', minwidth=80, width=225 , anchor='center')
    wM.tabla.column('precio_unitario', minwidth=100, width=175 , anchor='center')
    wM.tabla.column('unidad_monetaria', minwidth=100, width=175, anchor='center' )
    wM.tabla.column('validez', minwidth=100, width=175 , anchor='center')
    wM.tabla.column('unidad_medida', minwidth=100, width=175, anchor='center') 
     
    wM.tabla.heading('#0', text='ID', anchor ='e')
    # Parte común a todos los modos
    wM.tabla.heading('tipo', text='Tipo', anchor ='center')
    wM.tabla.heading('categoria', text='Categoria', anchor ='center')
    wM.tabla.heading('precio_unitario', text='Precio unitario', anchor ='center')
    wM.tabla.heading('unidad_monetaria', text='Unidad monetaria', anchor ='center')
    wM.tabla.heading('validez', text='Validez', anchor ='center')
    wM.tabla.heading('unidad_medida', text='U. medida', anchor ='center')
    # muestra de los datos
    for rec in resultados:
        #print(rec)
        wM.tabla.insert('',index = 'end', iid=None, text = rec[0], values = [rec[1], rec[2], rec[3], rec[4], rec[5], rec[6]])
    

## ALTURAS
    F3 = 540
    F4 = 610
    F5 = 680
    
## PRIMER PISO ##########################
# Tipo de recurso
    LM1 = tk.Label(wM, text = "Tipo")
    LM1.place(x=20, y=F3)
    
    refresh_conn(my_conn)
    try:
        my_cursor = my_conn.cursor()
        statement = "SELECT DISTINCT tipo FROM recursos "
        my_cursor.execute(statement)
        list = my_cursor.fetchall()
        #print(list) 
    except Exception as e:
        print("error 5366", e)

    # registro del observador
    rec_tipo = tk.StringVar()
    rec_tipo.trace('w', lambda *_ , var = rec_tipo : traceCB1(*_,var=var))
    global CM1
    CM1 = ttk.Combobox(wM, textvariable = rec_tipo, width = 17, values = list)
    CM1.place(x=110, y=F3)

    # Categoría
    LM2 = tk.Label(wM, text = "Categoría")
    LM2.place(x=390, y=F3)

    rec_cat = tk.StringVar()
    rec_cat.trace('w', lambda *_ , var = rec_cat : traceCB2(*_,var=var))

    global CM2
    CM2 = ttk.Combobox(wM, textvariable = rec_cat, width = 17)
    CM2.place(x=480, y=F3)

    ## SEGUNDO PISO #########################
    # Precio unitar

    LM3 = tk.Label(wM, text = "Precio unitario")
    LM3.place(x=20, y=F4)

    global EM3
    EM3 = tk.Entry(wM)
    EM3.place(x=110, y=F4)

    # Unidad de medida 
    LM5 = tk.Label(wM, text = "Un. Monetaria")
    LM5.place(x=390, y=F4)

    global EM5
    EM5 = tk.Entry(wM)
    EM5.place(x=480, y=F4)
        
    ## TERCER PISO #########################
        
    # Validez
    LM6 = tk.Label(wM, text = "Validez: ")
    LM6.place(x=20, y=F5)
    global EM6
    EM6 = tk.Entry(wM)
    EM6.place(x=110, y=F5)
    
    
    # Unidad de medida
    LM7 = tk.Label(wM, text = "U. medida")
    LM7.place(x=390, y=F5)

    global EM7
    EM7 = tk.Entry(wM)
    EM7.place(x=480, y=F5)
    
#   BOTONES #####################
    # boton Precarga
    BM4=tk.Button(wM, text="Seleccionar", width=12, command = lambda: B_selRec())
    BM4.place(x=700, y=F4)

    # boton Modificar
    global BM2
    BM2=tk.Button(wM,text="Modificar", width=12, command = lambda: B_modRecurso(), state = tk.DISABLED)
    BM2.place(x=830, y=F4)

    # boton Eliminar
    global BM6
    BM6=tk.Button(wM, text="Eliminar", width=12, command = lambda: B_elimRecurso())
    BM6.place(x=960, y=F4)

    # boton Nueva
    BM3=tk.Button(wM, text="Nuevo", width=12, command = lambda: B_nuevoRecurso())
    BM3.place(x=1090, y=F4)

    # boton Volver
    BM5=tk.Button(wM, text="Volver", width=12, command=lambda: menuFrom(wM))
    BM5.place(x=960, y=F5)

    # boton Salir
    BM1=tk.Button(wM, text="Salir", width=12, command = lambda: salirConAviso())
    BM1.place(x=1090, y=F5)

    #wI.tabla2.bind("<Double-1>", lambda evt, x = modo: B_selServ(x)) # habilita doble click para seleccionar
#    else:
#        messagebox.showinfo(message="Seleccione una orden.", title="Aviso del sistema")

############# Botón seleccionar recurso ##############################################################################
def B_selRec():
    curItem = wM.tabla.focus()
    numRec = wM.tabla.item(curItem).get('text')
    if len(str(numRec)) != 0:
        refresh_conn(my_conn)
        try:
            my_cursor = my_conn.cursor()
            statement = "SELECT * FROM recursos WHERE ID_recurso = %s"
            values = (numRec,) 
            my_cursor.execute(statement, values)
            resultado = my_cursor.fetchone()
            #print(resultado)
            ### Precarga de los tk.entry
            CM1.set('')                     # Tipo de recurso
            if resultado[1] is not None:
                CM1.set( resultado[1])
            CM2.set('')                     # Categoría
            if resultado[2] is not None:
                CM2.set( resultado[2])
            EM3.delete(0, tk.END)           # Precio unitario
            if resultado[3] is not None:
                EM3.insert(0, resultado[3])
            EM5.delete(0, tk.END)              # Unidad monetaria
            if resultado[4] is not None:
                EM5.insert(0, resultado[4])
            EM6.delete(0, tk.END)              # Validez
            if resultado[5] is not None:
                EM6.insert(0, resultado[5])
            EM7.delete(0, tk.END)              # Unidad de medida
            if resultado[6] is not None:
                EM7.insert(0, resultado[6])
            BM2['state'] = tk.NORMAL 
        except Exception as e:
            print("error 5498", e)
    else:
        messagebox.showinfo(message="Seleccione un recurso.", title="Aviso del sistema") 

############# Botón nuevo recurso #########################################################
def B_nuevoRecurso():
    try:
        my_cursor = my_conn.cursor()
        statement = "INSERT INTO recursos (tipo, categoria) VALUES('Nuevo_tipo', 'Nueva_categoria' )"
        my_cursor.execute(statement)
        my_conn.commit()
    except Exception as e:
        print("error 5109", e)
    win_geometry = wM.winfo_geometry()
    wM.destroy() 
    V_recursos(win_geometry)   

############# Botón eliminar empleo #####################################################
def B_elimRecurso():
    curItem = wM.tabla.focus()
    numRec = wM.tabla.item(curItem).get('text')
    if len(str(numRec)) != 0:
        if (messagebox.askokcancel(message="Eliminar recurso?", title="Confirmación de acción")):
            refresh_conn(my_conn)
            try:
                my_cursor = my_conn.cursor()
                statement = "DELETE FROM recursos WHERE ID_recurso = %s"
                val = (numRec,)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error 5527", e)
            win_geometry = wM.winfo_geometry()
            wM.destroy() 
            V_recursos(win_geometry)
    else:
        messagebox.showinfo(message="Seleccione un recurso.", title="Aviso del sistema")   

############################### Función del botón Modificar recurso ############################################
def B_modRecurso():
    curItem = wM.tabla.focus()
    numRec = wM.tabla.item(curItem).get('text')
    if len(str(numRec)) != 0:
        refresh_conn(my_conn)
        ######### Actualización del tipo de recurso
        tipo = CM1.get()
        if len(tipo) != 0:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE recursos SET tipo = %s WHERE ID_recurso = %s"
                val = (tipo, numRec)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error 5550", e)
        ######### Actualización de la categoria del recurso
        categoria = CM2.get()
        if len(categoria) != 0:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE recursos SET categoria = %s WHERE ID_recurso = %s"
                val = (categoria, numRec)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error 5561", e)
        ######### Actualización de la cantidad
        precio = EM3.get()
        if len(precio) != 0:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE recursos SET precio_unitario = %s WHERE ID_recurso = %s"
                val = (precio, numRec)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error 5221", e)
        ######### Actualización de la unidad de medida
        uniMone = EM5.get()
        if len(uniMone) != 0:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE recursos SET unidad_monetaria = %s WHERE ID_recurso = %s"
                val = (uniMone, numRec)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error 5584", e)
        ######### Actualización de detalle
        validez = EM6.get()
        if len(validez) != 0:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE recursos SET validez = %s WHERE ID_recurso = %s"
                val = (validez, numRec)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error 5595", e)
        ######### Actualización de la fecha
        uniMed = EM7.get()
        if len(uniMed) != 0:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE recursos SET unidad_Medida = %s WHERE ID_recurso = %s"
                val = (uniMed, numRec)
                my_cursor.execute(statement,val)
                my_conn.commit()
            except Exception as e:
                print("error 5606", e)
        win_geometry = wM.winfo_geometry()
        wM.destroy()
        V_recursos(win_geometry)
    else:
        messagebox.showinfo(message="Seleccione un recurso.", title="Aviso del sistema")

########################################################################################################################
LOG_FILENAME = 'SUGUS_log.out'
logging.basicConfig(filename=LOG_FILENAME, filemode = 'w', level=logging.ERROR)
#logging.basicConfig(filename=LOG_FILENAME, filemode = 'w', level=logging.INFO)

if __name__ == "__main__":
    try:
        main() 
    except Exception as e:
        logging.error(e, exc_info=True)
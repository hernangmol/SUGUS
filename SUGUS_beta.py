from tkinter import *
from tkinter import ttk
from tkinter.ttk import * ######################
from tkinter import messagebox
import mysql.connector
from datetime import date
import tkinter as tk
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import matplotlib.dates as mdates
import matplotlib
matplotlib.use("TkAgg")
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
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
# wB - Ver asignaciones a Asistencias
# wC - Modificar Asistencias
# wD - Agenda de tareas
# wE - Visualización de proyectos
# wF - Modificación de proyectos
# wG - Modificación de tareas
# wH - Modificación de ordenes de laboratorio
###########################################################################
def main():
    global my_conn
    try:
        my_conn = mysql.connector.connect(host = '192.168.100.105',
                                        port = 3306,
                                        #database = "BDNP",  # base de produccion
                                        database = "BDNP_t",  # base de test
                                        user = "admin",
                                        password = "cenadif2023")
        #print("conexion OK!")
    except Exception as e:
        print("error", e)

    refresh_conn(my_conn)
    ######### MENU PRINCIPAL ################
    global w1
    w1=Tk()
    global photo
    photo = PhotoImage(file = "Recursos/iconoCDF.png")
    w1.withdraw()
    w1.geometry("600x650")
    w1.iconphoto(False, photo)
    w1.title("CENADIF - Base de datos")

    image = PhotoImage(file="Recursos/CENADIF.png")
    image_sub=image.subsample(4)
    # Insertarla en una etiqueta.
    L10 = Label(w1, image=image_sub)
    L10.place(x=8, y=-85)

    L11 = Label(w1, text = "Menú principal")
    L11.place(x=205, y=105)
    # Primer piso
    L12 = Label(w1, text = "Solicitudes de trabajo")
    L12.place(x=75, y=135)
    global B11
    B11=Button(w1,text="Ver listado", width=12, command=V_viewRequest)
    B11.place(x=210, y=135)
    B11['state'] = DISABLED
    global B12
    B12=Button(w1,text="Ing. pedido", width=12, command = ingresar)
    B12.place(x=320, y=135)
    B12['state'] = DISABLED
    #B12['state'] = NORMAL
    # Segundo piso
    L13 = Label(w1, text = "Asignación de trabajo")
    L13.place(x=75, y=185)
    global B13
    B13=Button(w1,text="Asignar", width=12, command=assignment_form)
    B13.place(x=210, y=185)
    B13['state'] = DISABLED
    B1A=Button(w1,text="Estadísticas", width=12, command=V_estad)
    B1A.place(x=320, y=185)
    #B1A['state'] = DISABLED
    # Tercer piso ######################
    L14 = Label(w1, text = "Documentación técnica")
    L14.place(x=75, y=235)
    global B14
    B14=Button(w1,text="Ver pedidos", width=12, command=viewDocs)
    B14.place(x=210, y=235)
    B14['state'] = DISABLED
    global B15
    B15=Button(w1,text="Activos", width=12 , command=modDocs)
    B15.place(x=320, y=235)
    B15['state'] = DISABLED
    # Cuarto piso ######################
    L20 = Label(w1, text = "Proyectos")
    L20.place(x=75, y=285)
    global B26
    B26=Button(w1,text="Ver pedidos", width=12, command = lambda: V_viewPro('proy'))
    B26.place(x=210, y=285)
    global B27
    B27=Button(w1,text="Activos", width=12, command = lambda: V_modPro('proy'))
    B27.place(x=320, y=285)
    global B29
    B29=Button(w1,text="Inactivos", width=12, command = lambda: V_modPro('proy_c'))
    B29.place(x=430, y=285)
    # Quinto piso ######################
    L15 = Label(w1, text = "Desarrollos")
    L15.place(x=75, y=335)
    global B16
    B16=Button(w1,text="Ver pedidos", width=12, command = lambda: V_viewPro('des') )
    B16.place(x=210, y=335)
    B16['state'] = DISABLED
    global B17
    B17=Button(w1,text="Activos", width=12, command = lambda: V_modPro('des'))
    B17.place(x=320, y=335)
    B17['state'] = DISABLED
    global B30
    B30=Button(w1,text="Inactivos", width=12, command = lambda: V_modPro('des_c'))
    B30.place(x=430, y=335)
    # Sexto piso ######################
    L16 = Label(w1, text = "Asistencia técnica")
    L16.place(x=75, y=385)
    global B18
    B18=Button(w1,text="Ver pedidos", width=12, command = lambda: V_viewPro('ate'))
    B18.place(x=210, y=385)
    B18['state'] = DISABLED
    global B19
    B19=Button(w1,text="Activas", width=12, command = lambda: V_modPro('ate'))
    B19.place(x=320, y=385)
    B19['state'] = DISABLED
    global B31
    B31=Button(w1,text="Inactivas", width=12, command = lambda: V_modPro('ate_c'))
    B31.place(x=430, y=385)
    # Septimo piso ######################
    L17 = Label(w1, text = "Laboratorio")
    L17.place(x=75, y=435)
    global B20
    B20=Button(w1,text="Ord. trabajo", width=12, command = lambda: V_modLab('T'))
    B20.place(x=210, y=435)
    B20['state'] = DISABLED
    global B21
    B21=Button(w1,text="Ord. Internas", width=12)
    B21.place(x=320, y=435)
    B21['state'] = DISABLED
    # Octavo piso ######################
    L18 = Label(w1, text = "Admin. sistema")
    L18.place(x=75, y=485)
    global B22
    B22=Button(w1,text="Alta usuario", width=12, command = addUser)
    B22.place(x=210, y=485)
    B22['state'] = DISABLED
    global B23
    B23=Button(w1,text="Permisos", width=12, command = privChange)
    B23.place(x=320, y=485)
    B23['state'] = DISABLED
    # Noveno piso ######################
    L19 = Label(w1, text = "Usuario")
    L19.place(x=75, y=535)
    global B24
    B24=Button(w1,text="Contraseña", width=12, command=passChange)
    B24.place(x=210, y=535)
    global B25
    B25=Button(w1,text="Agenda", width=12, command = V_accSchedule)
    B25.place(x=320, y=535)
    #B25['state'] = DISABLED
    global B28
    B28=Button(w1,text="Tareas", width=12, command = lambda: V_modTar('user'))
    B28.place(x=430, y=535)

    # boton salir
    #Bsalir=Button(w1,text="Salir", width=12, command = w1.destroy)
    Bsalir=Button(w1,text="Salir", width=12, command = w1.quit)
    Bsalir.place(x=460, y=585)

    accessForm() 
    w1.mainloop()

#################### Ventana listado de pedidos #########################################################
def V_viewRequest():
    w1.withdraw()
    global w3
    w3=Toplevel()
    w3.state("zoomed")
    w3.iconphoto(False, photo)
    w3.title("Pedidos de trabajo")
    w3.frame3 = Frame(w3)
    w3.frame3.grid(rowspan=2, column=1, row=1)
    w3.tabla = ttk.Treeview(w3.frame3, height=35)
    w3.tabla.tag_configure('abiertos', background="red", foreground="white") #####################>>>>>>>>>>>>>>>>>>>>>>
    w3.tabla.grid(column=1, row=1)
    #w3.tabla.place(x=500, y=500)
    ladox = Scrollbar(w3.frame3, orient = VERTICAL, command= w3.tabla.yview)
    ladox.grid(column=0, row = 1, sticky='ew') 
    ladoy = Scrollbar(w3.frame3, orient =HORIZONTAL, command = w3.tabla.yview)
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
    B31=Button(w3,text="Salir", width=12, command = w1.destroy)
    B31.place(x=850, y=760)

    # boton Actualizar
    B32=Button(w3,text="Actualizar", width=12, command = actualizar)
    B32.place(x=735, y=760)

    # boton Volver
    B33=Button(w3,text="Volver", width=12, command = lambda: menuFrom(w3))
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
        if fila[8] is None:
            w3.tabla.insert('',index = 'end', iid=None, text = str(fila[0]), values = [fila[2], fila[1], fila[3], fila[5], fila[4], fila[11], fila[12]],tags='abiertos')
        else:
            w3.tabla.insert('',index = 'end', iid=None, text = str(fila[0]), values = [fila[2], fila[1], fila[3], fila[5], fila[4], fila[11], fila[12]])
           
    w3.after(1, lambda: w3.focus_force())
##############################################################################################################################
def ingresar():
    w1.withdraw()
    #print("Ingresando...")
    global w4
    w4=Toplevel()
    w4.iconphoto(False, photo)
    w4.geometry("500x650")
    w4.title("CENADIF - Base de datos")

    L41 = Label(w4, text = "Ingreso de pedido")
    L41.place(x=210, y=25)
    
    # Fecha de ingreso
    L42 = Label(w4, text = "Fecha de ingreso: ")
    L42.place(x=25, y=75)
    

    global E41
    E41 = Entry(w4)
    E41.place(x=125, y=75)
    E41.insert(0, date.today()) 

    L46 = Label(w4, text = "(Formato: AAAA-MM-DD)")
    L46.place(x=275, y=75)
    
    # Cliente

    L43 = Label(w4, text = "Nombre cliente: ")
    L43.place(x=25, y=125)

    global E42
    E42 = Entry(w4,width = 25)
    E42.place(x=125, y=125)

    
    # Filiación
    L44 = Label(w4, text = "Área cliente: ")
    L44.place(x=25, y=175)

    global E43
    E43 = Entry(w4,width = 25)
    E43.place(x=125, y=175)
    
    # Correo

    L48 = Label(w4, text = "Correo: ")
    L48.place(x=25, y=225)

    global E48
    E48 = Entry(w4, width = 25)
    E48.place(x=125, y=225)


    # Teléfono

    L49 = Label(w4, text = "Teléfono: ")
    L49.place(x=25, y=275)

    global E49
    E49 = Entry(w4, width = 25)
    E49.place(x=125, y=275)

    # Asunto

    L45 = Label(w4, text = "Asunto: ")
    L45.place(x=25, y=325)

    global E44
    E44 = Entry(w4, width = 56)
    E44.place(x=125, y=325)

    # Descripción

    L45 = Label(w4, text = "Descripción: ")
    L45.place(x=25, y=375)

    global T41
    T41 = Text(w4, width = 42, height = 8)
    T41.place(x=125, y=375)

     # Ingresó

    L47 = Label(w4, text = "Ingresó: ")
    L47.place(x=25, y=530)

    global E45
    E45 = Entry(w4, width = 25)
    E45.place(x=125, y=530)
    E45.insert(0, actualUser)

        # boton Volver
    B41=Button(w4,text="Volver", command = lambda: menuFrom(w4))
    B41.place(x=125, y=580)

        # boton Ingresar
    B42=Button(w4,text="Ingresar", command=ingresar_pedido)
    B42.place(x=400, y=580)

    w4.after(1, lambda: w4.focus_force())
######################################################################################################################
def assignment_form():
    w1.withdraw()
    global w7
    w7=Toplevel()
    w7.geometry("1000x700")
    w7.iconphoto(False, photo)
    w7.title("Pedidos de trabajo - Asignación")
    w7.frame7 = Frame(w7)
    w7.frame7.grid(rowspan=2, column=1, row=1)
    w7.tabla = ttk.Treeview(w7.frame7, height=15)
    w7.tabla.grid(column=1, row=1)
    ladox = Scrollbar(w7.frame7, orient = VERTICAL, command= w7.tabla.yview)
    ladox.grid(column=0, row = 1, sticky='ew') 
    ladoy = Scrollbar(w7.frame7, orient =HORIZONTAL, command = w7.tabla.yview)
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
    L72 = Label(w7, text = "Tipo de salida *")
    L72.place(x=25, y=380)

# Tipo de salida (combobox)
    # registering the observer
    my_var = StringVar()
    my_var.trace('w', lambda *_ , var = my_var : traceCB(*_,var=var))

    global C70
    C70 = ttk.Combobox(w7, state="readonly", textvariable = my_var, width = 17, values = ('Documentacion', 'Proyectos', 'Desarrollos', 'Asistencias', 'Laboratorio', 'Cierre'))
    C70.place(x=143, y=380) # menos 8?
    
# Fecha de asignación  
    L73 = Label(w7, text = "Fecha de asignación: * ")
    L73.place(x=25, y=430)

    global E73
    E73 = Entry(w7)
    E73.place(x=145, y=430)
    E73.insert(0, date.today())

    L74 = Label(w7, text = "(Formato: AAAA-MM-DD)")
    L74.place(x=295, y=430)

# Responsable (solo para Proyectos)
    
    global L76
    L76 = Label(w7, text = "Responsable: * ")

    #global E74
    #E74 = Entry(w7)

    list = userList()
    global C74
    C74 = ttk.Combobox(w7, state="readonly", width = 17, values = list)

# Descripción

    L75 = Label(w7, text = "Descripción: ")
    L75.place(x=25, y=480)

    global T71
    T71 = Text(w7, width = 42, height = 8)
    T71.place(x=145, y=480)

# boton Salir
    B71=Button(w7,text="Salir", width=12,  command = w1.destroy)
    B71.place(x=510, y=530)

# boton Volver
    B71=Button(w7,text="Volver", width=12,  command = lambda: menuFrom(w7))
    B71.place(x=510, y=480)

# boton Asignar
    B72=Button(w7,text="Asignar", width=12, command=B_assign)
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
    w9=Toplevel() 
    #style = ttk.Style(w9)
    #style.theme_use('clam')
    w9.state("zoomed")
    #w9.geometry("1100x650")
    w9.iconphoto(False, photo)
    w9.title("Pedidos de trabajo asignados a DyN")
    w9.frame = Frame(w9)
    w9.frame.grid(rowspan=3, column=2, row=1, pady = 2)
    #w9.frame.place(x=20, y=20)
    w9.tabla = ttk.Treeview(w9.frame, height=32)
    w9.tabla.grid(column=2, row=1)
    #w9.tabla.place(x=50, y=50)
    ladox = Scrollbar(w9.frame, orient = VERTICAL, command= w9.tabla.yview)
    ladox.grid(column=3, row = 1, sticky='ew') 
    ladoy = Scrollbar(w9.frame, orient =HORIZONTAL, command = w9.tabla.xview)
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
    B91=Button(w9,text="Salir", width=12, command = w1.destroy)
    B91.place(x=850, y=730)

    # boton Volver
    B93=Button(w9,text="Volver", width=12, command = lambda: menuFrom(w9))
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

###########################################################################

#################### Modificar Documentación #####################################################
def modDocs():
    w1.withdraw()
    global w8
    w8=Toplevel()
    w8.state("zoomed")
    #w8.geometry("1400x750")
    w8.iconphoto(False, photo)
    w8.title("CENADIF Base de datos - Documentación y Normas - Trabajos pendientes")
    w8.frame8 = Frame(w8)
    w8.frame8.grid(rowspan=2, column=1, row=1)
    w8.tabla = ttk.Treeview(w8.frame8, height=15)
    w8.tabla.grid(column=1, row=1)
    #w8.tabla.place(x=500, y=500)
    ladox = Scrollbar(w8.frame8, orient = VERTICAL, command= w8.tabla.yview)
    ladox.grid(column=0, row = 1, sticky='ew') 
    ladoy = Scrollbar(w8.frame8, orient =HORIZONTAL, command = w8.tabla.yview)
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
    L81 = Label(w8, text = "N° DyN")
    L81.place(x=75, y=380)

    global E81
    E81 = Entry(w8)
    E81.place(x=195, y=380)

# Tipo de documento
    L82 = Label(w8, text = "Tipo de documento")
    # L82.place(x=25, y=430)
    L82.place(x=400, y=380)

    #global E82
    #E82 = Entry(w8)
    #E82.place(x=525, y=380)

    global C81
    C81 = ttk.Combobox(w8, state="write", values = ['Norma', 'Plano', 'Esp. Técnica', 'Inst. Técnico', 'Informe', 'Manual'])
    C81.place(x=525, y=380)

# N° de documento  
    L83 = Label(w8, text = "N° de documento ")
    L83.place(x=1050, y=380)

    global E83
    E83 = Entry(w8)
    E83.place(x=1175, y=380)

# Título
    L84 = Label(w8, text = "Título")
    #L84.place(x=350, y=380)
    L84.place(x=400, y=430)

    global E84
    E84 = Entry(w8)
    #E84.place(x=475, y=380, width=450)
    E84.place(x=525, y=430, width=450)

# Fuente
    L85 = Label(w8, text = "Fuente")
    #L85.place(x=350, y=430)
    L85.place(x=75, y=480)

    #global E86
    #E86 = Entry(w8)
    #E86.place(x=195, y=480)

    global C82
    C82 = ttk.Combobox(w8, state="write", values = ['IRAM colección', 'AENOR MAS', 'UIC', 'Arch. Gral. Ferrov.', 'Base de datos', 'Internet'])
    C82.place(x=195, y=480)

# Emisor  
    L87 = Label(w8, text = "Emisor ")
    #L87.place(x=350, y=480)
    L87.place(x=725, y=380)

    global E87
    E87 = Entry(w8)
    #E87.place(x=475, y=480)
    E87.place(x=850, y=380)

# Fecha de resolución
    L88 = Label(w8, text = "Fecha de Resolución")
    L88.place(x=1050, y=480)

    global E88
    E88 = Entry(w8)
    E88.place(x=1175, y=480)

    L84 = Label(w8, text = "(Formato: AAAA-MM-DD)")
    L84.place(x=1165, y=505)

# Cantidad
    L89 = Label(w8, text = "Cantidad")
    L89.place(x=75, y=430)

    global E89
    E89 = Entry(w8)
    E89.place(x=195, y=430)

# Respondió  
    L8A = Label(w8, text = "Respondió ")
    L8A.place(x=1050, y=430)

    global E8A
    E8A = Entry(w8)
    E8A.place(x=1175, y=430)
    E8A.insert(0, actualUser)

# Precio
    L8D = Label(w8, text = "Precio")
    # L8D.place(x=675, y=430)
    L8D.place(x=400, y=480)

    global E8D
    E8D = Entry(w8)
    #E8D.place(x=800, y=430)
    E8D.place(x=525, y=480)

# Unidad monetaria  
    L8E = Label(w8, text = "Unidad monetaria")
    L8E.place(x=725, y=480)

    #global E8E
    #E8E = Entry(w8)
    #E8E.place(x=850, y=480)

    global C83
    C83 = ttk.Combobox(w8, state="write", values = ['ARS', 'USD', 'EUR'])
    C83.place(x=850, y=480)


# Descripción

    L8B = Label(w8, text = "Observaciones: ")
    L8B.place(x=75, y=530)

    global T81
    T81 = Text(w8, width = 81, height = 8)
    T81.place(x=195, y=530)

    # boton Salir
    B81=Button(w8,text="Salir", width=12, command = w1.destroy)
    B81.place(x=1135, y=585)

    # boton Modificar
    B82=Button(w8,text="Modificar", width=12, command = B_modDoc)
    B82.place(x=1135, y=635)

    # boton Volver
    B83=Button(w8,text="Volver", width=12, command = lambda: menuFrom(w8))
    B83.place(x=1000, y=585)

    # boton Precarga
    B84=Button(w8,text="Precarga", width=12, command = precDoc)
    B84.place(x=1000, y=635)

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
    wE=Toplevel()
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
    wE.frame = Frame(wE)
    wE.frame.grid(rowspan=2, column=1, row=1)
    wE.tabla = ttk.Treeview(wE.frame, height=25)
    wE.tabla.grid(column=1, row=1)
    ladox = Scrollbar(wE.frame, orient = VERTICAL, command= wE.tabla.yview)
    ladox.grid(column=0, row = 1, sticky='ew') 
    ladoy = Scrollbar(wE.frame, orient =HORIZONTAL, command = wE.tabla.yview)
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
    BE1=Button(wE,text="Salir", width=12, command = w1.destroy)
    BE1.place(x=850, y=560)

    # boton Volver
    BE3=Button(wE,text="Volver", width=12, command = lambda: menuFrom(wE))
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

###########################################################################


############# Ventana de modificación de proyectos ##########################################
def V_modPro(modo):
    # Parte común a todos los modos
    w1.withdraw()
    global wF
    wF=Toplevel()
    #wF.geometry("1500x800")
    wF.state("zoomed")
    wF.iconphoto(False, photo)
    # Titulo de la ventana (dependiente de modo)
    if modo == 'proy':
        wF.title("CENADIF Base de datos - Proyectos abiertos")
    if modo == 'proy_c':
        wF.title("CENADIF Base de datos - Proyectos cerrados")    
    if modo == 'des':
        wF.title("CENADIF Base de datos - Desarrollos abiertos")
    if modo == 'des_c':
        wF.title("CENADIF Base de datos - Desarrollos cerrados")
    if modo == 'ate':
        wF.title("CENADIF Base de datos - Asistencias técnicas abiertas")
    if modo == 'ate_c':
        wF.title("CENADIF Base de datos - Asistencias técnicas cerradas")
    # Parte común a todos los modos
    wF.frame = Frame(wF)
    wF.frame.grid(rowspan=2, column=1, row=1)
    wF.tabla = ttk.Treeview(wF.frame, height=15)
    wF.tabla.grid(column=1, row=1)

    ladox = Scrollbar(wF.frame, orient = VERTICAL, command= wF.tabla.yview)
    ladox.grid(column=0, row = 1, sticky='ew') 
    ladoy = Scrollbar(wF.frame, orient =HORIZONTAL, command = wF.tabla.xview)
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
        LF1 = Label(wF, text = "N° PROY")
    if modo == 'des' or modo == 'des_c':
        LF1 = Label(wF, text = "N° DES")
    if modo == 'ate' or modo == 'ate_c':
        LF1 = Label(wF, text = "N° ATE")
    # Parte común a todos los modos
    LF1.place(x=50, y=380)

    global EF1
    EF1 = Entry(wF)
    EF1.place(x=175, y=380)

# Responsable
    LFD = Label(wF, text = "Responsable")
    LFD.place(x=350, y=380)

#    global EFB
#    EFB = Entry(wF)
#    EFB.place(x=475, y=380)

    list = userList()
    global CFB
    CFB = ttk.Combobox(wF, state="readonly", width = 17, values = list)
    CFB.place(x=475, y=380)
    
# Nombre del proyecto
    # Nombre del... (dependiente del modo)
    if modo == 'proy' or modo == 'proy_c':
        LF4 = Label(wF, text = "Nombre del proyecto")
    if modo == 'des' or modo == 'des_c':
        LF4 = Label(wF, text = "Nombre del desarrollo")
    if modo == 'ate' or modo == 'ate_c':
        LF4 = Label(wF, text = "Nombre de asistencia")
    # Parte común a todos los modos
    LF4.place(x=50, y=430)

    global EF4
    EF4 = Entry(wF)
    EF4.place(x=175, y=430, width=425)

# Estado
    LF5 = Label(wF, text = "Estado")
    LF5.place(x=50, y=480)
    
    global CF9
    CF9 = ttk.Combobox(wF, state="readonly", width = 17, values = ['Abierto', 'En pausa', 'Cancelado', 'Finalizado'])
    CF9.place(x=175, y=480)

# Alcance
    LF8 = Label(wF, text = "Alcance")
    LF8.place(x=350, y=480)

    global CF8
    CF8 = ttk.Combobox(wF, width = 17, values = ['Ingeniería', 'Planos', 'Prototipo', 'Prototipo homologado', 'Proveedores'])
    CF8.place(x=475, y=480)

# Descripción

    LFB = Label(wF, text = "Descripción: ")
    LFB.place(x=50, y=530)

    global TF1
    if modo == 'des':
        TF1 = Text(wF, width = 53, height = 3)
    else:
        TF1 = Text(wF, width = 53, height = 7)
    TF1.place(x=175, y=530)

    if modo == 'des':
    # Prioridad
        LFC = Label(wF, text = "Prioridad")
        LFC.place(x=50, y=610)

        #global EFC
        #EFC = Entry(wF)
        #EFC.place(x=175, y=610)

        global CFC
        CFC = ttk.Combobox(wF, state="readonly", width = 17, values = ['Alta', 'Normal', 'Baja'])
        CFC.place(x=175, y=610)

    # Clasificación
        LFD = Label(wF, text = "Clasificación")
        LFD.place(x=350, y=610)

        global EFD
        EFD = Entry(wF)
        EFD.place(x=475, y=610)

    # NUM
        LFE = Label(wF, text = "NUM:")
        LFE.place(x=50, y=650)

        global EFE
        EFE = Entry(wF)
        EFE.place(x=175, y=650)

#   BOTONES #####################
    # boton Precarga
    BF4=Button(wF, text="Seleccionar", width=12, command = lambda: B_selProy(modo))
    BF4.place(x=175, y=680)

    # boton Modificar
    global BF2
    BF2=Button(wF,text="Modificar", width=12, state = DISABLED, command = lambda: B_modProy(modo))
    BF2.place(x=340, y=680)
    
    # boton Tareas
    BF3=Button(wF, text="Tareas", width=12, command = lambda: V_modTar(modo))
    BF3.place(x=510, y=680)

    # boton Volver
    BF3=Button(wF, text="Volver", width=12, command=lambda: menuFrom(wF))
    BF3.place(x=340, y=730)

    # boton Salir
    BF1=Button(wF, text="Salir", width=12, command = w1.quit)
    BF1.place(x=510, y=730)

    wF.tabla.delete(*wF.tabla.get_children())

    # proceso de los permisos
    refresh_conn(my_conn)
    if modo == 'proy':
        if actualRol & 16384: # modificación proyectos habilitado?
            try:
                my_cursor = my_conn.cursor()
                statement = "SELECT * FROM proyectos WHERE area = 'Proyectos' AND estado = 'Abierto'" # proyectos activos
                my_cursor.execute(statement)
                resultados = my_cursor.fetchall()
                #print(resultados) 
            except Exception as e:
                print("error 982", e)
        else:
            try:
                my_cursor = my_conn.cursor()
                statement = "SELECT * FROM proyectos WHERE area = 'Proyectos' AND estado = 'Abierto' and responsable = %s" # proyectos del usuario
                values = (actualUser,)
                my_cursor.execute(statement, values)
                resultados = my_cursor.fetchall()
                #print(resultados) 
            except Exception as e:
                print("error 1004", e)
    if modo == 'proy_c':
        try:
            my_cursor = my_conn.cursor()
            statement = "SELECT * FROM proyectos WHERE area = 'Proyectos' AND estado <> 'Abierto'" # proyectos cerrados
            my_cursor.execute(statement)
            resultados = my_cursor.fetchall()
            #print(resultados) 
        except Exception as e:
            print("error 982", e)
        
    if modo == 'des':
        if actualRol & 64: # modificación desarrollos habilitado?
            try:
                my_cursor = my_conn.cursor()
                statement = "SELECT * FROM proyectos WHERE area = 'Desarrollos' AND estado = 'Abierto'" # proyectos activos
                my_cursor.execute(statement)
                resultados = my_cursor.fetchall()
                #print(resultados) 
            except Exception as e:
                print("error 1020", e)
        else:
            try:
                my_cursor = my_conn.cursor()
                statement = "SELECT * FROM proyectos WHERE area = 'Desarrollos' AND estado = 'Abierto' and responsable = %s" # proyectos del usuario
                values = (actualUser,)
                my_cursor.execute(statement, values)
                resultados = my_cursor.fetchall()
                #print(resultados) 
            except Exception as e:
                print("error 1014", e)
    if modo == 'des_c':
        try:
            my_cursor = my_conn.cursor()
            statement = "SELECT * FROM proyectos WHERE area = 'Desarrollos' AND estado <> 'Abierto'" # proyectos cerrados
            my_cursor.execute(statement)
            resultados = my_cursor.fetchall()
            #print(resultados) 
        except Exception as e:
            print("error 982", e)
    if modo == 'ate':
        if actualRol & 256: # modificación asistencias habilitado?
            try:
                my_cursor = my_conn.cursor()
                statement = "SELECT * FROM proyectos WHERE area = 'Asistencias' AND estado = 'Abierto'" # proyectos activos
                my_cursor.execute(statement)
                resultados = my_cursor.fetchall()
                #print(resultados) 
            except Exception as e:
                print("error 1029", e)
        else:
            try:
                my_cursor = my_conn.cursor()
                statement = "SELECT * FROM proyectos WHERE area = 'Asistencias' AND estado = 'Abierto' and responsable = %s" # proyectos del usuario
                values = (actualUser,)
                my_cursor.execute(statement, values)
                resultados = my_cursor.fetchall()
                #print(resultados) 
            except Exception as e:
                print("error 1014", e)
    if modo == 'ate_c':
        try:
            my_cursor = my_conn.cursor()
            statement = "SELECT * FROM proyectos WHERE area = 'Asistencias' AND estado <> 'Abierto'" # proyectos cerrados
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

        wF.tabla.insert('',index = fila[7], iid=None, text = str(fila[7]), values = [fila[3], fila[5], fila[1], fila[4], fila[6],fila[8], pedido[0]])
        #wF.tabla.insert('', index = fila[8], iid=None, text = str(fila[7]), values = [fila[3], fila[5], fila[1], fila[4], fila[6],fila[8], pedido[0]])
    wF.after(1, lambda: wF.focus_force())

#################### Alta de usuario #########################################################
def addUser():
    w1.withdraw()
    global w6
    w6=Toplevel()
    w6.geometry("320x500")
    w6.iconphoto(False, photo)
    w6.title("CENADIF - Base de datos")

    L61 = Label(w6, text = "Alta de usuario")
    L61.place(x=105, y=25)
    
    # ID_usuario
    L62 = Label(w6, text = "ID de usuario: ")
    L62.place(x=25, y=75)

    global E61
    E61 = Entry(w6)
    E61.place(x=155, y=75)
    
     # nombre_usuario
    L63 = Label(w6, text = "Nombre de usuario: ")
    L63.place(x=25, y=125)

    global E62
    E62 = Entry(w6)
    E62.place(x=155, y=125)

     # Nombres 
    L64 = Label(w6, text = "Nombres: ")
    L64.place(x=25, y=175)

    global E63
    E63 = Entry(w6)
    E63.place(x=155, y=175)

     # Apellidos 
    L65 = Label(w6, text = "Apellidos: ")
    L65.place(x=25, y=225)

    global E64
    E64 = Entry(w6)
    E64.place(x=155, y=225)

     # correo
    L66 = Label(w6, text = "Correo electrónico: ")
    L66.place(x=25, y=275)

    global E65
    E65 = Entry(w6)
    E65.place(x=155, y=275)

         # fecha de ingreso
#    L67 = Label(w6, text = "Fecha ingreso: ")
#    L67.place(x=25, y=325)

#    global E67
#    E67 = Entry(w6)
#    E67.place(x=135, y=325)

    # Perfil
    L68 = Label(w6, text = "Perfil: ")
    L68.place(x=25, y=325)

    global E68
    E68 = Entry(w6)
    E68.place(x=155, y=325)

    # Perfil
    L69 = Label(w6, text = "Se asigna contraseña predeterminada")
    L69.place(x=25, y=375)

    # boton Volver
    B61=Button(w6,text="Volver", command = lambda: menuFrom(w6))
    B61.place(x=120, y=425)

    # boton Ingresar
    B62=Button(w6,text="Ingresar", command = userConfirm)
    B62.place(x=230, y=425)
    
    w6.after(1, lambda: w6.focus_force())

############################### ventana de cambio de privilegios ###############################
def privChange():
    w1.withdraw()
    global wA
    wA=Toplevel()
    wA.geometry("500x560")
    wA.iconphoto(False, photo)
    wA.title("CENADIF Base de datos - Gestión de permisos")
    # Entry Usuario
    LA1 = Label(wA, text = "Usuario: ")
    LA1.place(x=50, y=40)

    global EA1
    EA1 = Entry(wA)
    EA1.place(x=110, y=40)

    # Boton precargar
    BA1=Button(wA,text="Precargar", width=12, command = precUsu)
    BA1.place(x=280, y=40)

    LA2 = Label(wA, text = "Permisos:")
    LA2.place(x=50, y=80)

    # Radiobuttons (1x permiso)
    # Ver pedidos
    LA3 = Label(wA, text = "Ver pedidos")
    LA3.place(x=50, y=120)
    global verPed
    verPed = IntVar()
    global CA1
    CA1=tk.Checkbutton(wA, variable = verPed, onvalue = 1, offvalue = 0)
    CA1.place(x=190, y=120)
    # Ingresar pedidos
    LA4 = Label(wA, text = "Ingresar pedidos")
    LA4.place(x=270, y=120)
    global ingPed
    ingPed = IntVar()
    global CA2
    CA2=tk.Checkbutton(wA, variable = ingPed, onvalue = 1, offvalue = 0)
    CA2.place(x=410, y=120)
    # Asignar
    LA5 = Label(wA, text = "Asignar")
    LA5.place(x=50, y=160)
    global asiPed
    asiPed = IntVar()
    global CA3
    CA3=tk.Checkbutton(wA, variable = asiPed, onvalue = 1, offvalue = 0)
    CA3.place(x=190, y=160)
    # Ver documentación 
    LA6 = Label(wA, text = "Ver documentación")
    LA6.place(x=50, y=200)
    global verDoc
    verDoc = IntVar()
    global CA4
    CA4=tk.Checkbutton(wA, variable = verDoc, onvalue = 1, offvalue = 0)
    CA4.place(x=190, y=200)
    # Modificar documentación 
    LA7 = Label(wA, text = "Modificar documentación")
    LA7.place(x=270, y=200)
    global modDoc
    modDoc = IntVar()
    global CA5
    CA5=tk.Checkbutton(wA, variable = modDoc, onvalue = 1, offvalue = 0)
    CA5.place   (x=410, y=200)
    # Ver desarrollos 
    LA8 = Label(wA, text = "Ver desarrollos")
    LA8.place(x=50, y=240)
    global verDes
    verDes = IntVar()
    global CA6
    CA6=tk.Checkbutton(wA, variable = verDes, onvalue = 1, offvalue = 0)
    CA6.place(x=190, y=240)
    # Modificar desarrollos 
    LA9 = Label(wA, text = "Modificar desarrollos")
    LA9.place(x=270, y=240)
    global modDes
    modDes = IntVar()
    global CA7
    CA7=tk.Checkbutton(wA, variable = modDes, onvalue = 1, offvalue = 0)
    CA7.place(x=410, y=240)
    # Ver asistencias 
    LA10 = Label(wA, text = "Ver asistencias")
    LA10.place(x=50, y=280)
    global verAsi
    verAsi = IntVar()
    global CA8
    CA8=tk.Checkbutton(wA, variable = verAsi, onvalue = 1, offvalue = 0)
    CA8.place(x=190, y=280)
    # Modificar asistencias 
    LA11 = Label(wA, text = "Modificar asistencias")
    LA11.place(x=270, y=280)
    global modAsi
    modAsi = IntVar()
    global CA9
    CA9=tk.Checkbutton(wA, variable = modAsi, onvalue = 1, offvalue = 0)
    CA9.place(x=410, y=280)
    # Ver laboratorio 
    LA12 = Label(wA, text = "Ver laboratorio")
    LA12.place(x=50, y=320)
    global verLab
    verLab = IntVar()
    global CA10
    CA10=tk.Checkbutton(wA, variable = verLab, onvalue = 1, offvalue = 0)
    CA10.place(x=190, y=320)
    # Modificar laboratorio 
    LA13 = Label(wA, text = "Modificar laboratorio")
    LA13.place(x=270, y=320)
    global modLab
    modLab = IntVar()
    global CA11
    CA11=tk.Checkbutton(wA, variable = modLab, onvalue = 1, offvalue = 0)
    CA11.place(x=410, y=320)
    # Alta usuario
    LA14 = Label(wA, text = "Alta de usuario")
    LA14.place(x=50, y=360)
    global altUsu
    altUsu = IntVar()
    global CA12
    CA12=tk.Checkbutton(wA, variable = altUsu, onvalue = 1, offvalue = 0)
    CA12.place(x=190, y=360)
    # Gestionar permisos
    LA15 = Label(wA, text = "Gestionar permisos")
    LA15.place(x=270, y=360)
    global gesPer
    gesPer = IntVar()
    global CA13
    CA13=tk.Checkbutton(wA, variable = gesPer, onvalue = 1, offvalue = 0)
    CA13.place(x=410, y=360)
    # Ver Proyectos
    LA16 = Label(wA, text = "Ver proyectos")
    LA16.place(x=50, y=400)
    global verPro
    verPro = IntVar()
    global CA14
    CA14=tk.Checkbutton(wA, variable = verPro, onvalue = 1, offvalue = 0)
    CA14.place(x=190, y=400)
    # Modificar Proyectos
    LA17 = Label(wA, text = "Modificar proyectos")
    LA17.place(x=270, y=400)
    global modPro
    modPro = IntVar()
    global CA15
    CA15=tk.Checkbutton(wA, variable = modPro, onvalue = 1, offvalue = 0)
    CA15.place(x=410, y=400)


    # Boton Volver
    BA2=Button(wA,text="Volver", width=12, command = lambda: menuFrom(wA))
    BA2.place(x=250, y=470)

    # Boton 
    global BA3
    BA3=Button(wA,text="Modificar", width=12, command= modPriv, state = DISABLED)
    BA3.place(x=360, y=470)

    wA.focus_force()

#################### Cambio de contraseña #########################################################
def passChange():
    w1.withdraw()
    global w5
    w5=Toplevel()
    w5.geometry("330x300")
    w5.iconphoto(False, photo)
    w5.title("CENADIF - Base de datos")

    L51 = Label(w5, text = "Cambio de contraseña")
    L51.place(x=105, y=25)
    
    # Contraseña actual
    L52 = Label(w5, text = "Contraseña actual: ")
    L52.place(x=25, y=75)

    global E51
    #E51 = Entry(w5)
    E51 = Entry(w5, show = "*")
    E51.place(x=135, y=75)
    
     # Contraseña nueva
    L53 = Label(w5, text = "Contraseña nueva: ")
    L53.place(x=25, y=125)

    global E52
    #E52 = Entry(w5)
    E52 = Entry(w5, show = "*")
    E52.place(x=135, y=125)

     # Segunda contraseña nueva 
    L54 = Label(w5, text = "Contraseña nueva: ")
    L54.place(x=25, y=175)

    global E53
    #E53 = Entry(w5)
    E53 = Entry(w5, show = "*")
    E53.place(x=135, y=175)

    # boton Volver
    B51=Button(w5,text="Volver", command = lambda: menuFrom(w5))
    B51.place(x=110, y=225)

    # boton Ingresar
    B52=Button(w5,text="Cambiar", command = changeConfirm)
    B52.place(x=220, y=225)
    
    w5.after(1, lambda: w5.focus_force())

############################### Ventana de agenda de tareas ########################################################
def V_accSchedule():
    #print("Agenda")
    w1.withdraw()
    global wD
    wD=Toplevel()
    wD.geometry("560x610")
    wD.iconphoto(False, photo)
    wD.title("Agenda de tareas")
    wD.frame = Frame(wD)
    wD.frame.grid(rowspan=2, column=1, row=1)
    wD.tabla = ttk.Treeview(wD.frame, height=25)
    wD.tabla.grid(column=1, row=1)
    ladox = Scrollbar(wD.frame, orient = VERTICAL, command= wD.tabla.yview)
    ladox.grid(column=0, row = 1, sticky='ew') 
    ladoy = Scrollbar(wD.frame, orient =HORIZONTAL, command = wD.tabla.yview)
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
    BD1=Button(wD,text="Salir", width=12, command = w1.destroy)
    BD1.place(x=430, y=560)

    # boton Actualizar
    #BD2=Button(wD,text="Actualizar", width=12)
    #BD2.place(x=735, y=560)

    # boton Volver
    BD3=Button(wD,text="Volver", width=12, command = lambda: menuFrom(wD))
    BD3.place(x=300, y=560)

    wD.tabla.delete(*wD.tabla.get_children())

    refresh_conn(my_conn)
    try:
            my_cursor = my_conn.cursor()
            statement = "SELECT * FROM agenda WHERE nombre_usuario = %s "
            #values=(actualID)
            my_cursor.execute(statement, (actualUser,))
            resultados = my_cursor.fetchall()
            print(f"Agenda del usuario: {resultados}") 
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
     w3.destroy()
     V_viewRequest()

########################################################################################################
def menuFrom(window):
    window.destroy()
    w1.deiconify()

########################################################################################################
def ingresar_pedido():
    fechaIng = E41.get()
    cliente = E42.get()
    filiacion = E43.get()
    asunto = E44.get()
    descripcion = T41.get("1.0",'end-1c')
    operador = E45.get()
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
            messagebox.showinfo(message="pedido NO registrado.", title="Aviso del sistema")

    else:
         messagebox.showinfo(message="Campos vacios.", title="Aviso del sistema")

################### Función Lista desplegable de usuarios #############################
def userList():
    refresh_conn(my_conn)
    try:
        my_cursor = my_conn.cursor()
        statement = "SELECT nombre_usuario FROM usuarios WHERE nombre_usuario <> 'admin' "
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
                    statement = '''INSERT INTO proyectos (ID_solicitud, responsable, area, estado, numero_cdf) VALUES('{}', '{}','{}','{}','{}')'''.format(idSol, responsable, 'Proyectos', 'Abierto', 0)
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
                    statement = '''INSERT INTO proyectos (ID_solicitud, responsable, area, estado, numero_cdf) VALUES('{}', '{}','{}','{}','{}')'''.format(idSol, responsable, 'Desarrollos', 'Abierto', 0)
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
                    statement = '''INSERT INTO proyectos (ID_solicitud, responsable, area, estado, numero_cdf) VALUES('{}', '{}','{}','{}','{}')'''.format(idSol, responsable, 'Asistencias', 'Abierto', 0)
                    #statement = '''INSERT INTO asistencias (ID_solicitud) VALUES('{}')'''.format(idSol )
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
                # crea entrada tabla ordenes
                try:    
                    my_cursor = my_conn.cursor()
                    statement = '''INSERT INTO ordenes (ID_solicitud, tipo, numero_cdf)
                        VALUES('{}', '{}')'''.format(idSol, 'T', 0 )
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
    
    # w8.update()
    # w8.update_idletasks()
    w8.destroy()
    modDocs()
    
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

    #E82.delete(0, END)              # tipo
    C81.set('')              # Tipo
    if resultado[2] is not None:
        C81.set(resultado[2])
    E83.delete(0, END)              # número  
    if resultado[3] is not None:
        E83.insert(0, resultado[3]) 
    E84.delete(0, END)              # título
    if resultado[4] is not None:
        E84.insert(0, resultado[4])  
    C82.set('')                     # fuente
    if resultado[5] is not None:
        C82.set(resultado[5])
    E8D.delete(0, END)              # precio
    if resultado[6] is not None:
         E8D.insert(0, resultado[6]) 
    C83.set('')                     # unidad monetaria
    if resultado[7] is not None:
        C83.set(resultado[7])
    E87.delete(0, END)              # emisor
    if resultado[8] is not None:
         E87.insert(0, resultado[8]) 
    E88.delete(0, END)              # fecha resolución
    if resultado[9] is not None:
        fecha = resultado[9].strftime('%Y-%m-%d')
        #print(fecha)
        E88.insert(0,fecha)
    T81.delete("1.0","end")         # Observaciones
    if resultado[10] is not None:
         T81.insert("1.0", resultado[10]) 
    E89.delete(0, END)              # cantidad
    if resultado[11] is not None:
         E89.insert(0, resultado[11])
    E8A.delete(0, END)              # respondió
    if resultado[12] is not None:
         E8A.insert(0, resultado[12]) 

############# Botón seleccionar proyecto #####################################################
def B_selProy(modo):
    curItem = wF.tabla.focus() # diccionario de la fila seleccionada en la tabla
    #print(wF.tabla.item(curItem))
    aux = wF.tabla.item(curItem).get('values')
    refresh_conn(my_conn)
    if len(aux)!=0:
        try:
            my_cursor = my_conn.cursor()
            statement = "SELECT * FROM proyectos WHERE ID_solicitud = %s"
            values = (aux[5],)  # aux[5] es el ID_solicitud
            my_cursor.execute(statement, values)
            resultado = my_cursor.fetchone()

            gantt(resultado[0], wF)
            # Precarga de los entry
            EF1.delete(0, END)              # N° Proy
            if resultado[7] is not None:
                EF1.insert(0, resultado[7])
            CFB.set('')                      # Responsable
            if resultado[1] is not None:
                CFB.set(resultado[1])
            EF4.delete(0, END)              # nombre de proyecto
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
                EFD.delete(0, END)              # clasificación
                if resultado[10] is not None:
                    EFD.insert(0, resultado[10]) 
                EFE.delete(0, END)              # NUM
                if resultado[11] is not None:
                    EFE.insert(0, resultado[11]) 
            TF1.delete("1.0","end")         # descripción
            if resultado[4] is not None:
                TF1.insert("1.0", resultado[4])
            # habilitar botón modificar 
            if modo != 'proy_c' and modo != 'des_c' and modo != 'ate_c':
                BF2['state'] = NORMAL 
        except Exception as e:
            print("error 890", e)
    else:
#        numPro = EF1.get()
#        if len(numPro) != 0 :
#            if modo == 'proy' or modo == 'proy_c':
#                area = 'Proyectos'
#            elif modo == 'des':
#                area = 'Desarrollos'
#            elif modo == 'ate':
#                area = 'Asistencias'
#            else:
#                pass # manejo de error de modo
#            try:
#                my_cursor = my_conn.cursor()
#                statement = "SELECT * FROM proyectos WHERE numero_CDF = %s and area = %s"
#                values = (numPro, area)  # aux[5] es el ID_solicitud
#                my_cursor.execute(statement, values)
#                resultado = my_cursor.fetchone()

#                gantt(resultado[0], wF)
                # Precarga de los entry
#                EF1.delete(0, END)              # N° Proy
#                if resultado[7] is not None:
#                    EF1.insert(0, resultado[7])
#                CFB.set('')                      # Responsable
#                if resultado[1] is not None:
#                    CFB.set(resultado[1])
#                EF4.delete(0, END)              # nombre de proyecto
#                if resultado[3] is not None:
#                    EF4.insert(0, resultado[3]) 
#                CF9.set('')                       # estado
#                if resultado[5] is not None:
#                    CF9.set(resultado[5]) 
#                CF8.set('')                      # alcance
#                if resultado[6] is not None: 
#                    CF8.set(resultado[6])
#                if modo == 'des':
#                    CFC.set('')                   # prioridad
#                    if resultado[9] is not None:
#                        CFC.set(resultado[9])
#                    EFD.delete(0, END)              # clasificación
#                    if resultado[10] is not None:
#                        EFD.insert(0, resultado[10]) 
#                    EFE.delete(0, END)              # NUM
#                    if resultado[11] is not None:
#                        EFE.insert(0, resultado[11]) 
#                TF1.delete("1.0","end")         # descripción
#                if resultado[4] is not None:
#                    TF1.insert("1.0", resultado[4])
                # habilitar botón modificar 
#                if modo != 'proy_c' and modo != 'des_c' and modo != 'ate_c':
#                    BF2['state'] = NORMAL 
#            except Exception as e:
#                print("error 2026", e)
#        else:     
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
    curItem = wF.tabla.focus()
    numCDF = wF.tabla.item(curItem).get('text')
    refresh_conn(my_conn)
    if len(numCDF)!=0:
        ######### Actualización del numero de proyecto
        NnumCDF = EF1.get()
        if len(NnumCDF) != 0:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE proyectos SET numero_cdf = %s WHERE area = %s and numero_cdf = %s"
                val = (NnumCDF, area, numCDF)
                my_cursor.execute(statement,val)
                my_conn.commit()
                numCDF = NnumCDF       
            except Exception as e:
                print("error", e)
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
        wF.destroy() 
        V_modPro(modo)
    else:
        if modo == 'proy':
            messagebox.showinfo(message="Seleccione un proyecto.", title="Aviso del sistema")
        if modo == 'des':
            messagebox.showinfo(message="Seleccione un desarrollo.", title="Aviso del sistema")
        if modo == 'ate':
            messagebox.showinfo(message="Seleccione una asistencia.", title="Aviso del sistema")

############# Ventana de modificación de tareas ##########################################
def V_modTar(modo):
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
    wG=Toplevel()
    #wG.geometry("1500x800")
    wG.state("zoomed")
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

            wG.frame = Frame(wG)
            wG.frame.grid(rowspan=2, column=1, row=1)
            wG.tabla = ttk.Treeview(wG.frame, height=1)
            wG.tabla.grid(column=1, row=1)

            ladox = Scrollbar(wG.frame, orient = VERTICAL, command= wG.tabla.yview)
            ladox.grid(column=0, row = 2, sticky='ew')
            #ladox.pack(side ='right', fill ='x') 
            ladoy = Scrollbar(wG.frame, orient =HORIZONTAL, command = wG.tabla.xview)
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
            BG6=Button(wG, text="Eliminar", width=12, state = DISABLED, command = lambda: B_elimTar(modo))
            BG6.place(x=375, y=705)

            # boton Nueva
            BG3=Button(wG, text="Nueva", width=12, state = DISABLED, command = lambda: B_nuevaTar(proyData[0], modo))
            BG3.place(x=505, y=705)
            if modo != 'proy_c' and modo != 'des_c' and modo != 'ate_c':
                BG3['state'] = NORMAL
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
    wG.frame2 = Frame(wG)
    wG.frame2.grid(rowspan=2, column=1, row=2)
    wG.frame2.place(y=70)
    wG.tabla2 = ttk.Treeview(wG.frame2, height=12)
    wG.tabla2.grid(column=1, row=1)

    ladox2 = Scrollbar(wG.frame2, orient = VERTICAL, command= wG.tabla2.yview)
    ladox2.grid(column=0, row = 1, sticky='ew') 
    ladoy2 = Scrollbar(wG.frame2, orient =HORIZONTAL, command = wG.tabla2.xview)
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
    LG1 = Label(wG, text = "Nombre Tarea")
    LG1.place(x=20, y=380)

    global EG1
    EG1 = Entry(wG)
    EG1.place(x=110, y=380, width=320)

# Estado
    LG2 = Label(wG, text = "Estado")
    LG2.place(x=430, y=380)

    global CG2
    CG2 = ttk.Combobox(wG, state="readonly", width = 17, values = ('Planificada','En curso', 'Cerrada', 'Pausada', 'Finalizada'))
    CG2.place(x=480, y=380)

## SEGUNDO PISO #########################
    # Responsable
    LG3 = Label(wG, text = "Responsable")
    LG3.place(x=20, y=430)

    list = userList()
    global CG3
    CG3 = ttk.Combobox(wG, state="readonly", width = 17, values = list)
    CG3.place(x=110, y=430)

# inicio
    LG4 = Label(wG, text = "Inicio")
    LG4.place(x=250, y=430)

    global EG4
    EG4 = Entry(wG)
    EG4.place(x=300, y=430)

# Fin  
    LG5 = Label(wG, text = "Fin")
    LG5.place(x=430, y=430)

    global EG5
    EG5 = Entry(wG)
    EG5.place(x=480, y=430)

## TERCER PISO #########################
    # Avance
    LG6 = Label(wG, text = "Avance")
    LG6.place(x=20, y=480)

    global EG6
    EG6 = Entry(wG)
    EG6.place(x=110, y=480)

# H/H prog
    LG7 = Label(wG, text = "hs.prog")
    LG7.place(x=250, y=480)

    global EG7
    EG7 = Entry(wG)
    EG7.place(x=300, y=480)

# H/H dedicadas  
    LG8 = Label(wG, text = "hs.dedic")
    LG8.place(x=430, y=480)

    global EG8
    EG8 = Entry(wG)
    EG8.place(x=480, y=480)

## CUARTO PISO #########################
    # Integrante 1
    LG9 = Label(wG, text = "Asignados: 1")
    LG9.place(x=20, y=530)

    #global EG9
    #EG9 = Entry(wG)
    #EG9.place(x=110, y=530)

    global CG9
    CG9 = ttk.Combobox(wG, state="readonly", width = 17, values = list)
    CG9.place(x=110, y=530)

    if modo == 'proy' or modo == 'proy_c' or modo == 'ate':
    # Integrante 2
        LGA = Label(wG, text = "      2")
        LGA.place(x=250, y=530)

        global CGA
        CGA = ttk.Combobox(wG, state="readonly", width = 17, values = list)
        CGA.place(x=300, y=530)

    # Integrante 3  
        LGB = Label(wG, text = "      3")
        LGB.place(x=430, y=530)

        global CGB
        CGB = ttk.Combobox(wG, state="readonly", width = 17, values = list)
        CGB.place(x=480, y=530)

    global TG1
    if modo == 'ate':
        # Integrante 4
        LGC = Label(wG, text =  "4")
        LGC.place(x=80, y=580)

        global CGC
        CGC = ttk.Combobox(wG, state="readonly", width = 17, values = list)
        CGC.place(x=110, y=580)

        # Integrante 5
        LGD = Label(wG, text =  "      5")
        LGD.place(x=250, y=580)

        global CGD
        CGD = ttk.Combobox(wG, state="readonly", width = 17, values = list)
        CGD.place(x=300, y=580)

        # Integrante 6
        LGE = Label(wG, text =  "      6")
        LGE.place(x=430, y=580)

        global CGE
        CGE = ttk.Combobox(wG, state="readonly", width = 17, values = list)
        CGE.place(x=480, y=580)

    # Descripción (ATE)
        LGB = Label(wG, text = "Descripción: ")
        LGB.place(x=20, y=630)

        TG1 = Text(wG, width = 61, height = 3)
        TG1.place(x=110, y=630)
    else:
    # Descripción
        LGB = Label(wG, text = "Descripción: ")
        LGB.place(x=20, y=580)

        TG1 = Text(wG, width = 61, height = 6)
        TG1.place(x=110, y=580)


#   BOTONES #####################
    # boton Precarga
    BG4=Button(wG, text="Seleccionar", width=12, command = lambda: B_selTar(modo))
    BG4.place(x=115, y=705)

    # boton Modificar
    global BG2
    BG2=Button(wG,text="Modificar", width=12, command = lambda: B_modTar(modo), state = DISABLED)
    BG2.place(x=245, y=705)

    # boton Volver
    BG5=Button(wG, text="Volver", width=12, command=lambda: menuFrom(wG))
    BG5.place(x=375, y=750)

    # boton Salir
    BG1=Button(wG, text="Salir", width=12, command = w1.quit)
    BG1.place(x=505, y=750)

    #gantt(proyData[0], wG)

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
        try:
            my_cursor = my_conn.cursor()
            statement = '''INSERT INTO usuarios (ID_usuarios, nombre_usuario, nombres, apellidos, contraseña, correo, rol) 
            VALUES('{}', '{}', '{}', '{}', '{}', '{}', '{}')'''.format(ID_user, n_user, name, surname, '1234',email, rol )
            my_cursor.execute(statement)
            my_conn.commit() 

        except Exception as e:
            print("error", e)

        messagebox.showinfo(message="Nuevo usuario ingresado. La contraseña por defecto es 1234", title="Aviso del sistema")
    else:
        messagebox.showinfo(message="Campos obligatorios vacios.", title="Aviso del sistema")

############################### precargar usuario ##############################################
def precUsu():
    editUser = EA1.get()

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

    BA3['state'] = NORMAL

############################### modificar privilegios ##########################################
def modPriv():
    #print('modificando permisos...')
    editUser = EA1.get()
    privileges = IntVar()
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
            statement = "SELECT nombre_usuario, contraseña FROM usuarios WHERE nombre_usuario = %s"
            my_cursor.execute(statement, (actualUser,))
            resultados = my_cursor.fetchone()
            #print(resultados) 
        except Exception as e:
            print("error", e)

        aux = tuple([actualUser, password])
        #print(aux)
        if  aux == resultados:
            try:
                my_cursor = my_conn.cursor()
                statement = "UPDATE usuarios SET contraseña = %s WHERE nombre_usuario = %s"
                val = (newPass1, actualUser)
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
    refresh_conn(my_conn)
    try:
        my_cursor = my_conn.cursor()
        statement = "SELECT nombre_tarea, fecha_inicio, fecha_fin FROM tareas WHERE ID_proyecto = %s"
        values = (proy,) 
        my_cursor.execute(statement, values)
        resultados = my_cursor.fetchall()
        #print(resultados) 
    except Exception as e:
        print("error", e)
    
    tasks=[]
    start_dates=[]
    lengths =[]
    end_dates =[]
    flag_data = 0
    for fila in resultados:
        if fila[1] != None and fila[2] != None:
            tasks.append(fila[0][:7]+".") # primeros 7 caracteres del nombre de tareas    
            start_dates.append(fila[1])
            lengths.append((fila[2]-fila[1]).days)
            #lengths.append(5)
            end_dates.append(fila[2])
            flag_data = 1 
    if flag_data == 0:
        return 0

    # Initialize the figure and axis
    global fig
    fig, ax = plt.subplots()

    # Set y-axis tick labels
    ax.set_yticks(np.arange(len(tasks)))
    ax.set_yticklabels(tasks)

    # Plot each task as a horizontal bar
    for i in range(len(tasks)):
        start_date = pd.to_datetime(start_dates[i])
        end_date = start_date + pd.DateOffset(days=lengths[i])
        ax.barh(i, end_date - start_date, left=start_date, height=0.5, align='center')

    # Set x-axis limits
    min_date = pd.to_datetime(min(start_dates)) - pd.DateOffset(1)
    #max_date = pd.to_datetime(max(start_dates)) + pd.DateOffset(days=max(durations))
    max_date = pd.to_datetime(max(end_dates)) + pd.DateOffset(1)
    ax.set_xlim(min_date, max_date)

    # Customize the chart
    ax.xaxis_date()
    ax.xaxis.set_major_locator(mdates.WeekdayLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%m-%d'))
    ax.set_xlabel('Fecha[mes-día]')
    # Ajuste dinámico de la cuadrícula
    if((max_date - min_date).days>70):
        ax.xaxis.set_major_locator(mdates.DayLocator(interval=14))
    if((max_date - min_date).days>140):
        #ax.xaxis.set_major_locator(mdates.DayLocator(interval=28))
        ax.xaxis.set_major_locator(mdates.MonthLocator())
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%y-%m'))
        ax.set_xlabel('Fecha[año-mes]')
    if((max_date - min_date).days>350):
        ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
    if((max_date - min_date).days>700):
        ax.xaxis.set_major_locator(mdates.MonthLocator(interval=4))
    if((max_date - min_date).days>1500):
        ax.xaxis.set_major_locator(mdates.MonthLocator(interval=8))
#    ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
    ax.set_ylabel('Tareas')
    ax.set_title('Planificación del proyecto')
    #plt.iconphoto(False, photo)

    # Display the chart
    plt.grid(True)
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
            wG.destroy() 
            V_modTar(modo)
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
    wG.destroy() 
    V_modTar(modo)

############# Botón seleccionar tarea #####################################################
def B_selTar(modo):
    curItem = wG.tabla2.focus()
    numTar = wG.tabla2.item(curItem).get('text')
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
            # Precarga de los entry
            EG1.delete(0, END)              # Nombre de tarea
            if resultado[3] is not None:
                EG1.insert(0, resultado[3])
            CG2.set('')                     # Estado
            if resultado[16] is not None:
                CG2.set( resultado[16])
            CG3.set('')                     # Responsable
            if resultado[2] is not None:
                CG3.set(resultado[2])
            EG4.delete(0, END)              # Inicio
            if resultado[14] is not None:
                EG4.insert(0, resultado[14])
            EG5.delete(0, END)              # Fin
            if resultado[15] is not None:
                EG5.insert(0, resultado[15])
            EG6.delete(0, END)              # Avance
            if resultado[6] is not None:
                EG6.insert(0, resultado[6])
            EG7.delete(0, END)              # hs/H programadas
            if resultado[5] is not None:
                EG7.insert(0, resultado[5])
            EG8.delete(0, END)              # hs/H dedicadas
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
            TG1.delete("1.0","end")         # Descripción
            if resultado[4] is not None:
                TG1.insert("1.0", resultado[4])
            
            if modo != 'proy_c'and modo != 'des_c' and modo != 'ate_c':
                BG2['state'] = NORMAL # habilitación del botón Modificar
                BG6['state'] = NORMAL # habilitación del botón Eliminar
            #####################################################################     
            
        except Exception as e:
            print("error", e)
    else:
        messagebox.showinfo(message="Seleccione una tarea.", title="Aviso del sistema")

############################### Función del botón Modificar tarea ############################################
def B_modTar(modo):
    curItem = wG.tabla2.focus()
    numTar = wG.tabla2.item(curItem).get('text')
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
                wG.destroy() 
                V_modTar(modo)
        except:         
            V_modTar('user')
    else:
        messagebox.showinfo(message="Seleccione una tarea.", title="Aviso del sistema")

##############################Función del botón Enviar (Inicio de sesión)#################################
#
# Pendientes:
#   - Limitar el número de intentos de ingreso a 3 y cerrar app
# 
# ######################################################################################################### 
def verify():
        global actualUser, actualID, actualRol
        actualUser = (E1.get()) 
        password = (E2.get())
        
        refresh_conn(my_conn)
        try:
            my_cursor = my_conn.cursor()
            statement = "SELECT contraseña, rol, ID_usuarios FROM usuarios WHERE nombre_usuario = %s"
            my_cursor.execute(statement, (actualUser,))
            resultados = my_cursor.fetchone()

        except Exception as e:
            print("error", e)

        #aux = tuple([user, password])
        #print(aux)
        if  password == resultados[0]:
            actualID = resultados[2]
            actualRol = resultados[1]
            #print(actualRol)
            w2.withdraw()
            w1.deiconify()
            if int(resultados[1]) & 1:
                B11['state'] = NORMAL
            if int(resultados[1]) & 2:
                B12['state'] = NORMAL
            if int(resultados[1]) & 4:
                B13['state'] = NORMAL
            if int(resultados[1]) & 8:
                B14['state'] = NORMAL
            if int(resultados[1]) & 16:
                B15['state'] = NORMAL
            if int(resultados[1]) & 32:
                B16['state'] = NORMAL
            if int(resultados[1]) & 64:
                B17['state'] = NORMAL
            if int(resultados[1]) & 128:
                B18['state'] = NORMAL
            if int(resultados[1]) & 256:
                B19['state'] = NORMAL
            if int(resultados[1]) & 1024:
                B20['state'] = NORMAL
            if int(resultados[1]) & 2048:
                B21['state'] = NORMAL
            if int(resultados[1]) & 4096:
                B22['state'] = NORMAL
            if int(resultados[1]) & 4096:
                B23['state'] = NORMAL
        else:
            messagebox.showinfo(message="Nombre de usuario y/o contraseña incorrectos.", title="Aviso del sistema")
        #print("FIN")        

def accessForm():
    global w2
    w2=Toplevel()
    w2.geometry("350x150")
    w2.iconphoto(False, photo)
    w2.title("CENADIF - Base de datos")    

    L1 = Label(w2, text = "Formulario de acceso")#.pack(pady=10)
    L1.grid(row = 0, column = 1, padx = 5, pady= 5)
    
    # Username

    L2 = Label(w2, text = "Usuario: ")
    L2.grid(row = 1, column = 0, padx = 5, pady= 5)

    global E1
    E1 = Entry(w2)
    E1.grid(row = 1, column = 1, padx = 5, pady= 5)
    #E1.insert(0, "hgomezmolino") ############################### SACAR!!!!!!!!!!!!!!!!!!!!!!!!!

    # Password

    L3 = Label(w2, text = "Contraseña: ")
    L3.grid(row = 2, column = 0, padx = 5, pady= 5)

    global E2
    E2 = Entry(w2, show = "*")
    E2.grid(row = 2, column = 1, padx = 5, pady= 5)
    
    B1=Button(w2,text="Enviar", command = verify)
    B1.grid(row = 3, column = 2, padx = 5, pady= 5)

    w2.focus_force()
###########################################################################################################
def V_estad():
    # Parte común a todos los modos
    #w1.withdraw()
    global wC
    wC=Toplevel()
    wC.geometry("1500x800")
    wC.state("zoomed")
    wC.iconphoto(False, photo)
    # Titulo de la ventana 
    wC.title("CENADIF Base de datos - Estadísticas del sistema")
    # datos para gráficos #################
        # Gráfico 1 - pedidos
    #areas = ["Documentación","Proyectos","Desarrollos","Asistencias","Laboratorio"]  # descomentar para agregar Laboratorio
    #pedidos = [0, 0, 0, 0, 0]                                                        # descomentar para agregar Laboratorio
    areas = ["Documentación","Proyectos","Desarrollos","Asistencias"]
    pedidos = [0, 0, 0, 0]
    
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
            pedidos[i] = resultado[0][0]

    except Exception as e:
        print("error 3135", e)
    
    total = sum(pedidos)
    # colors = ['violet', 'blue', 'yellowgreen', 'gold', 'red']      # descomentar para agregar Laboratorio
    colors = ['violet', 'blue', 'yellowgreen', 'gold']
    # Gráfico 2.1 - Desarrollos
    desarrollos = [0, 0, 0, 0]
    P_estados = ["Abiertos","Finalizados","Cancelados","Pausados"]
    sql_estados = ['Abierto', 'Finalizado', 'Cancelado', 'En pausa']
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

    chart1 = FigureCanvasTkAgg(fig1,wC)
    chart1.get_tk_widget().pack(fill=BOTH, expand=TRUE)

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
    if (conn.is_connected()):
        #print("Connected")
        pass
    else:
        try:
            conn.reconnect(attempts=1, delay=0)
        except:

            w1.quit    ################## NO FUNCIONA ##################################################

def V_modLab(tipo):
    w1.withdraw()
    global wH
    wH=Toplevel()
    wH.state("zoomed")
    wH.iconphoto(False, photo)
    # Titulo de la ventana (dependiente de modo)
    if tipo == 'T':
        wH.title("CENADIF - Laboratorio - Ordenes de trabajo")
    if tipo == 'I':
        wF.title("CENADIF - Laboratorio - Ordenes internas") 
    wH.frame = Frame(wH)
    wH.frame.grid(rowspan=2, column=1, row=1)
    wH.tabla = ttk.Treeview(wH.frame, height=15)
    wH.tabla.grid(column=1, row=1)

    ladox = Scrollbar(wH.frame, orient = VERTICAL, command= wH.tabla.yview)
    ladox.grid(column=0, row = 1, sticky='ew') 
    ladoy = Scrollbar(wH.frame, orient =HORIZONTAL, command = wH.tabla.xview)
    ladoy.grid(column = 1, row = 0, sticky='ns')
    wH.tabla.configure(xscrollcommand = ladox.set, yscrollcommand = ladoy.set)
       
    wH.tabla['columns'] = ('descripcion', 'responsable', 'fin', 'informe', 'pedido','cliente')
    wH.tabla.column('#0', minwidth=50, width=60, anchor='center')
    wH.tabla.column('descripcion', minwidth=100, width=450, anchor='center')
    wH.tabla.column('responsable', minwidth=80, width=100 , anchor='center')
    wH.tabla.column('fin', minwidth=100, width=100 , anchor='center')
    wH.tabla.column('informe', minwidth=100, width=250, anchor='center' )
    wH.tabla.column('pedido', minwidth=100, width=150 , anchor='center')
    wH.tabla.column('cliente', minwidth=100, width=200, anchor='center')   

       # Titulo del número CENADIF y nombre (dependientes del modo)
    if tipo == 'T':
        wH.tabla.heading('#0', text='N° O/T', anchor ='e')
    if tipo == 'I':
        wH.tabla.heading('#0', text='N° O/I', anchor ='e')
    
    # Parte común a todos los modos
    wH.tabla.heading('descripcion', text='Descripción', anchor ='center')
    wH.tabla.heading('responsable', text='Responsable', anchor ='center')
    wH.tabla.heading('fin', text='Fecha fin', anchor ='center')
    wH.tabla.heading('informe', text='Informe', anchor ='center')
    wH.tabla.heading('pedido', text='Pedido', anchor ='center')
    wH.tabla.heading('cliente', text='Cliente', anchor ='center')

    #   BOTONES #####################
    # boton Precarga
    BH4=Button(wH, text="Seleccionar", width=12)
    BH4.place(x=175, y=680)

    # boton Modificar
    global BH2
    BH2=Button(wH,text="Modificar", width=12, state = DISABLED)
    BH2.place(x=340, y=680)
    
    # boton Tareas
    BH3=Button(wH, text="Tareas", width=12)
    BH3.place(x=510, y=680)

    # boton Volver
    BH3=Button(wH, text="Volver", width=12, command=lambda: menuFrom(wH))
    BH3.place(x=340, y=730)

    # boton Salir
    BH1=Button(wH, text="Salir", width=12, command = w1.quit)
    BH1.place(x=510, y=730)

    wH.tabla.delete(*wH.tabla.get_children())
    # preparacion de los datos
    refresh_conn(my_conn)
    if tipo == 'T':    
        try:
            my_cursor = my_conn.cursor()
            statement = "SELECT * FROM ordenes WHERE tipo = 'T'" # ordenes de trabajo
            my_cursor.execute(statement)
            resultados = my_cursor.fetchall()
            #print(resultados) 
        except Exception as e:
            print("error 3430", e)
        for fila in resultados:
            try:
                my_cursor = my_conn.cursor(buffered=True)
                statement = "SELECT filiacion_cliente FROM solicitudes WHERE ID_solicitud = %s" # ordenes de trabajo
                values = (fila[4],)
                my_cursor.execute(statement, values)
                cliente = my_cursor.fetchone()
                #print(cliente)
            except Exception as e:
                print("error 3435", e)
    # muestra de los datos
            wH.tabla.insert('',index = fila[2], iid=None, text = str(fila[2]), values = [fila[5], fila[3], fila[7],fila[6], fila[4], cliente,])
    elif tipo == 'I':    
        try:
            my_cursor = my_conn.cursor()
            statement = "SELECT * FROM ordenes WHERE tipo = 'I'" # ordenes de trabajo
            my_cursor.execute(statement)
            resultados = my_cursor.fetchall()
            print(resultados) 
        except Exception as e:
            print("error 3439)", e)
    else:
        pass # To Do: insertar manejo de error de tipo 

main()
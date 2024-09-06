import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
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
#import datetime as dt
import math
import SUGUS_conn as sc
import re

#######################################################################################################################
# Reporte de horas trabajadas en 2024 de desarrollos finalizados en 2024.
#######################################################################################################################
def main():
    desarrollos = 0
    tareas = 0
    hs_total = 0  
    connList = sc.connect_db()
    my_conn = connList [0]
    my_cursor = connList[1]

    ################ busqueda de proyecto de Desarrollos y Finalizados en 2024 ###############################
    try:
        statement = "SELECT proyectos.numero_cdf, proyectos.ID_proyecto, proyectos.ID_solicitud, tareas.ID_tarea, tareas.fecha_fin, tareas.fecha_inicio, tareas.HH_dedicadas  FROM tareas inner join proyectos on proyectos.ID_proyecto = tareas.ID_proyecto WHERE area = 'Desarrollos' and proyectos.estado = 'Finalizado' and tareas.fecha_fin > '2023-12-31' ORDER BY proyectos.ID_proyecto"
        my_cursor.execute(statement, )
        resultados = my_cursor.fetchall()
        #print(resultados)
    except Exception as e:
        print("error", e)
    ############### Filtra los que son de SOFSE ##############################################################
    filtro = []
    filtro_proy = []
    for des in resultados:
        #print(des)
        '''
        try:
            statement = "SELECT filiacion_cliente FROM solicitudes where ID_solicitud = %s"
            values = (des[2],)
            my_cursor.execute(statement, values)
            filiac = my_cursor.fetchone()
            #print(filiac)
        except Exception as e:
            print("error", e)
        #if es_SOFSE(filiac[0]):
        '''
        filtro.append(des)
        filtro_proy. append(des[1])
    sc.disconnect_db(my_conn, my_cursor)
    ########### Calculo de horas ###############################################################
    for tar in filtro:
        #print(tar)
        #print('_tarea:', tar[0])
        tareas +=1
        dur_m = diff_month( tar[4], tar[5])
        if dur_m > 0:
            if tar[6] is not None:
                horas = float(tar[6] * tar[4].month /dur_m)
            else:
                horas = 0
        else:
            horas = float(tar[6])
        #print('horas de tarea', horas)
        hs_total = hs_total + horas
    ############ calculo de desarrollos ##############################################################
    while len(filtro_proy) > 0:
        if filtro_proy.count(filtro_proy[0]) == 1:
            desarrollos +=1
        filtro_proy.remove(filtro_proy[0])
    print('desarrollos finalizados:', desarrollos)
    print('tareas:', tareas)
    print('total_horas:', hs_total)

def es_SOFSE(lista):
    x = re.search("SOFSE", lista)
    return x

def diff_month(d1, d2):
    return (d1.year - d2.year) * 12 + d1.month - d2.month
    
if __name__ == "__main__":
        main() 
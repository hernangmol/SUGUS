import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from datetime import date
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib
matplotlib.use("TkAgg")
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import BDNP as bd
from datetime import datetime
#import datetime as dt
import SUGUS_conn as sc

#######################################################################################################################
# Reporte de horas dedicadas a SOFSE. 
# Desarrollos, en curso
#   Tareas -> hs ponderadas
#
#######################################################################################################################
def report():
    asistencias = 0
    tareas = 0
    hs_total = 0  
    connList = sc.connect_db()
    my_conn = connList [0]
    my_cursor = connList[1]
    ################ busqueda de proyecto de Desarrollos y En curso ###############################
    try:
        statement = "SELECT numero_cdf, ID_proyecto FROM proyectos WHERE area = 'Asistencias' and estado = 'En curso' ORDER BY numero_cdf DESC"
        my_cursor.execute(statement, )
        resultados = my_cursor.fetchall()
    except Exception as e:
        print("error", e)
    #print(resultados)
    for at in resultados:
        try:
            statement = "SELECT ID_tarea FROM tareas WHERE ID_proyecto = %s"
            values = (at[1],)
            my_cursor.execute(statement, values)
            resultados = my_cursor.fetchall()
        except Exception as e:
            print("error", e)
        if len(resultados) == 0:
            print(at[0])
            #print(resultados)
    '''
    ################ filtrado de desarrollos de SOFSE #############################################
    #sofse = filter(es_SOFSE, resultados)
    #print(sofse)
    for ate in resultados:
        print(ate)
        asistencias +=1
        try:
            statement = "SELECT ID_tarea, fecha_inicio, avance, HH_dedicadas FROM tareas WHERE ID_proyecto = %s" 
            values = (ate[0],)
            my_cursor.execute(statement, values)
            aux = my_cursor.fetchall()
        except Exception as e:
            print("error", e)
        for tar in aux:
            print('_tarea:', tar[0])
            tareas +=1
            dur_m = diff_month( datetime.now(), aux[0][1])
            if tar[3] is not None:
                horas = float(tar[3] * datetime.now().month /dur_m)
            else:
                horas = 0
            #print('horas de tarea', horas)
            hs_total = hs_total + horas
            #print('horas totales', hs_total)  
    sc.disconnect_db(my_conn, my_cursor)
    print('asistencias técnicas en curso:', asistencias)
    print('tareas:', tareas)
    print('total_horas:', hs_total)
    '''

#def es_SOFSE(lista):
#    x = re.search("SOFSE", lista[3])
#    return x


#def es_SOFSE(lista):
#    x = re.search("SOFSE", lista[3])
#    return x

def diff_month(d1, d2):
    return (d1.year - d2.year) * 12 + d1.month - d2.month

def main():
    report()

if __name__ == "__main__":
    main() 
import pandas as pd
import mysql.connector
from tkinter import messagebox

def AT_charge(at):
    # Extracción datos de excel ########################
#    df = pd.read_excel("fuente.xlsx")
    data = df.iloc[at]
    # Preparación de datos para BDNP ###################
    pedido = {} # P
    asistencia = {} # A
    tarea = {} # T

    # fecha_ingreso(P), fecha_salida(P), fecha_inicio(T)
    if pd.isnull(data.iloc[7]):
        pedido['fecha_ingreso'] = '0000-00-00'
    else:     
        aux = pd.to_datetime(data.iloc[7], errors='ignore')
        pedido['fecha_ingreso'] = aux.strftime("%Y-%m-%d")
    pedido['fecha_salida'] = pedido['fecha_ingreso']
    tarea['fecha_inicio'] = pedido['fecha_ingreso']
    #print(data)
    # asunto(P), nombre_proy(D), nombre_tarea(T)
    pedido['asunto'] = data.iloc[1]
    asistencia['nombre_proy'] = pedido['asunto']
    tarea['nombre_tarea'] = pedido['asunto']

    # descripcion_asunto(P), descripcion(D y T), descripción_salida(p)
    pedido['descripcion_asunto'] = 'Dato migrado automaticamente de DNT.'
    asistencia['descripcion'] = 'Dato migrado automaticamente de DNT.'
    tarea['descripcion'] = 'Dato migrado automaticamente de DNT.'
    pedido['descripcion_salida'] = 'Dato migrado automaticamente de DNT.'
    # cliente(P)
    if pd.isnull(data.iloc[11]):
        pedido['cliente'] = ''
    else: 
        pedido['cliente'] = data.iloc[11].strip()
    # filiacion_cliente(P)
    if pd.isnull(data.iloc[12]):
        pedido['filiacion'] = ''
    else: 
        pedido['filiacion'] = data.iloc[12].strip()   

    #filiacion = data.iloc[12]
    #pedido['filiacion'] = filiacion.strip()
    
    # operador(p)
    pedido['operador'] ='SISTEMA'
    # tipo_salida(p), area(d)
    pedido['tipo_salida'] = 'Asistencias'
    asistencia['area'] = 'Asistencias'

    # responsable(A)
    if pd.isnull(data.iloc[6]):
        asistencia['responsable'] = ''
    else: 
        asistencia['responsable'] = data.iloc[6][:20]

    # estado(D), estado(T)
    asistencia['estado'] = data.iloc[2]
    tarea['estado'] = data.iloc[2]
    # alcance(D)
    asistencia['alcance'] = data.iloc[3]
    # numero_cdf(D)
    aux=data.iloc[0]
    asistencia['numero_cdf'] = int(aux[-4:])
    print('**************** AT: ', asistencia['numero_cdf']) 
    # prioridad(D)
    asistencia['prioridad'] = data.iloc[4]
    # clasificacion(D)
    asistencia['clasificacion'] = data.iloc[11]
    # responsable(T)
    if pd.isna(data.iloc[5]):
        tarea['responsable'] = ''
    else:
        tarea['responsable'] = data.iloc[5].split()[0]
        try:
            tarea['integrante_1'] = data.iloc[5].split()[1]
        except:
            pass
        try:
            tarea['integrante_2'] = data.iloc[5].split()[2]
        except:
            pass
        try:
            tarea['integrante_3'] = data.iloc[5].split()[3]
        except:
            pass
        try:
            tarea['integrante_4'] = data.iloc[5].split()[4]
        except:
            pass
        try:
            tarea['integrante_5'] = data.iloc[5].split()[5]
        except:
            pass
        try:
            tarea['integrante_6'] = data.iloc[5].split()[6]
        except:
            pass
    # avance(T)
    if pd.isna(data.iloc[9]):
        tarea['avance'] = 0
    else:
        tarea['avance'] = int(data.iloc[9])
    # fecha_fin(t)
    if pd.isnull(data.iloc[8]):
        tarea['fecha_fin'] = '0000-00-00'
    else:
        aux = pd.to_datetime(data.iloc[8], errors='ignore')
        tarea['fecha_fin'] = aux.strftime("%Y-%m-%d")
    
    #print('Pedido: ', pedido)
    #print('Asistencia: ', asistencia)
    #print('Tarea: ',  tarea)

 # actualización de la base de datos ############    
    global my_conn
    # conexión con la base de datos
    try:
        my_conn = mysql.connector.connect(host = '192.168.100.105',
                                        port = 3306,
                                        #database = "BDNP_t",  # base de test
                                        database = "BDNP",  # base de produccion
                                        user = "admin",
                                        password = "cenadif2023")
        print("conexion OK!")
    except Exception as e:
        print("error", e)
# actualizacion de tablas 'solicitudes'
    try:
        my_cursor = my_conn.cursor()
        statement = '''INSERT INTO solicitudes (fecha_ingreso, asunto, descripcion_asunto, cliente, filiacion_cliente, operador, tipo_salida, fecha_salida, descripcion_salida) 
        VALUES('{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}')'''.format( pedido['fecha_ingreso'], pedido['asunto'], pedido['descripcion_asunto'],
                                                                                pedido['cliente'], pedido['filiacion'], pedido['operador'], 
                                                                                pedido['tipo_salida'], pedido['fecha_salida'],pedido['descripcion_salida'])
        my_cursor.execute(statement)
        my_conn.commit()

    except Exception as e:
        print("error", e)
        messagebox.showinfo(message="pedido NO registrado.", title="Aviso del sistema")
        exit('error de pedido')

# actualización de tabla proyectos(desarrollo)

    try:
            my_cursor = my_conn.cursor()
            statement = "SELECT ID_solicitud FROM solicitudes WHERE asunto = %s"
            values = (pedido['asunto'],)
            my_cursor.execute(statement, values)
            ID_pedido = my_cursor.fetchone()
            print('ID de pedido:',ID_pedido[0])
    except Exception as e:
        print("error 124", e)
    try:
        my_cursor = my_conn.cursor()
        statement = '''INSERT INTO proyectos (responsable, area, nombre_proy, descripcion, estado, alcance, numero_cdf, ID_solicitud, prioridad, clasificacion) 
        VALUES('{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}')'''.format( asistencia['responsable'], asistencia['area'], asistencia['nombre_proy'],
                                                                                asistencia['descripcion'], asistencia['estado'], asistencia['alcance'], 
                                                                                asistencia['numero_cdf'], ID_pedido[0], asistencia['prioridad'],
                                                                                asistencia['clasificacion'])
        my_cursor.execute(statement)
        my_conn.commit()

    except Exception as e:
        print("error", e)
        messagebox.showinfo(message="asistencia NO registrada.", title="Aviso del sistema")
        exit('error de desarrollo')

# actualización de tabla tareas

    try:
            my_cursor = my_conn.cursor()
            statement = "SELECT ID_proyecto FROM proyectos WHERE ID_solicitud = %s"
            values = (ID_pedido[0],)
            my_cursor.execute(statement, values)
            ID_proyecto = my_cursor.fetchone()
            print('ID de proyecto:',ID_proyecto[0])
    except Exception as e:
        print("error 124", e)
    try:
        my_cursor = my_conn.cursor()
        statement = '''INSERT INTO tareas (ID_proyecto, responsable, nombre_tarea, descripcion, avance, fecha_inicio, fecha_fin, estado)
        VALUES('{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}')'''.format( ID_proyecto[0], tarea['responsable'], tarea['nombre_tarea'],
                                                                                tarea['descripcion'], tarea['avance'], 
                                                                                tarea['fecha_inicio'], tarea['fecha_fin'], tarea['estado'])
        my_cursor.execute(statement)
        my_conn.commit()

    except Exception as e:
        print("error", e)
        messagebox.showinfo(message="tarea NO registrada.", title="Aviso del sistema")
        exit('error de tarea')
    try:
            my_cursor = my_conn.cursor()
            statement = "SELECT ID_tarea FROM tareas WHERE ID_proyecto = %s"
            values = (ID_proyecto[0],)
            my_cursor.execute(statement, values)
            ID_tarea = my_cursor.fetchone()
            print('ID de tarea:',ID_tarea[0])
    except Exception as e:
        print("error 176", e)

    my_conn.close()
    messagebox.showinfo(message="Nueva asistencia registrada.", title="Aviso del sistema")

def main():
    global df
    df = pd.read_excel("fuente_ATs.xlsx")
    
    AT_charge(0)
    AT_charge(1)
    AT_charge(2)
    AT_charge(3)
    #AT_charge(4)
    
if __name__ == '__main__':
    main()
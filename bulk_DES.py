import pandas as pd
import mysql.connector
from tkinter import messagebox

def DES_charge(des):
    # Extracción datos de excel ########################
#    df = pd.read_excel("fuente.xlsx")
    data = df.iloc[des]
    #print(data)
    # Preparación de datos para BDNP ###################
    pedido = {} # P
    desarrollo = {} # D
    tarea = {} # T
    # fecha_ingreso(P), fecha_salida(P), fecha_inicio(T)

    if pd.isnull(data.iloc[10]):
        pedido['fecha_ingreso'] = '0000-00-00'
    else:     
        aux = pd.to_datetime(data.iloc[10], errors='ignore')
        pedido['fecha_ingreso'] = aux.strftime("%Y-%m-%d")
    pedido['fecha_salida'] = pedido['fecha_ingreso']
    tarea['fecha_inicio'] = pedido['fecha_ingreso']
    # asunto(P), nombre_proy(D), nombre_tarea(T)
    pedido['asunto'] = data.iloc[3]
    desarrollo['nombre_proy'] = pedido['asunto']
    tarea['nombre_tarea'] = pedido['asunto']
    # descripcion_asunto(P), descripcion(D y T), descripción_salida(p)
    pedido['descripcion_asunto'] = 'Dato migrado automaticamente de DNT.'
    desarrollo['descripcion'] = 'Dato migrado automaticamente de DNT.'
    tarea['descripcion'] = 'Dato migrado automaticamente de DNT.'
    pedido['descripcion_salida'] = 'Dato migrado automaticamente de DNT.'
    # cliente(P), filiacion_cliente(P)
    try:
        cliente, filiacion = data.iloc[16].split('-')
        pedido['cliente'] = cliente.strip()
        pedido['filiacion'] = 'SOFSE - ' + filiacion.strip()
    except:
        pedido['cliente'] = data.iloc[16]
        pedido['filiacion'] = data.iloc[16]
    # operador(p)
    pedido['operador'] ='SISTEMA'
    # tipo_salida(p), area(d)
    pedido['tipo_salida'] = 'Desarrollos'
    desarrollo['area'] = 'Desarrollos'
    # responsable(D)
    if pd.isnull(data.iloc[9]):
        desarrollo['responsable'] = ''
    else: 
        desarrollo['responsable'] = data.iloc[9][:20]

    # estado(D), estado(T)
    desarrollo['estado'] = data.iloc[4]
    tarea['estado'] = data.iloc[4]
    # alcance(D)
    desarrollo['alcance'] = data.iloc[5]
    # numero_cdf(D)
    aux=data.iloc[1]
    desarrollo['numero_cdf'] = int(aux[-4:])
    print('DES: ',desarrollo['numero_cdf']) 
    # prioridad(D)
    desarrollo['prioridad'] = data.iloc[7]
    # clasificacion(D)
    desarrollo['clasificacion'] = data.iloc[15]
    # NUM(D)
    if pd.isnull(data.iloc[18]):
        desarrollo['NUM'] = 0
    else: 
        desarrollo['NUM'] = data.iloc[18]
    # responsable(T)
    if pd.isna(data.iloc[8]):
        tarea['responsable'] = ''
    else:
        tarea['responsable'] = data.iloc[8][:20]
    # avance(T)
    if pd.isna(data.iloc[12]):
        tarea['avance'] = 0
    else:
        tarea['avance'] = int(data.iloc[12])

    # HH_dedic(t)
    match data.iloc[13]:
        case 'Baja':
            tarea['HH_dedic'] = 75
        case 'Media':
            tarea['HH_dedic'] = 200
        case 'Alta':
            tarea['HH_dedic'] = 625
        case _ :
            tarea['HH_dedic'] = 0
    # fecha_fin(t)
    if pd.isnull(data.iloc[11]):
        tarea['fecha_fin'] = '0000-00-00'
    else:
        aux = pd.to_datetime(data.iloc[11], errors='ignore')
        tarea['fecha_fin'] = aux.strftime("%Y-%m-%d")

    #print(desarrollo)

 # actualización de la base de datos ############    
    global my_conn
    # conexión con la base de datos
    try:
        my_conn = mysql.connector.connect(host = '192.168.100.105',
                                        port = 3306,
                                        database = "BDNP",  # base de produccion
                                        #database = "BDNP_t",  # base de test
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
        statement = '''INSERT INTO proyectos (responsable, area, nombre_proy, descripcion, estado, alcance, numero_cdf, ID_solicitud, prioridad, clasificacion, NUM) 
        VALUES('{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}')'''.format( desarrollo['responsable'], desarrollo['area'], desarrollo['nombre_proy'],
                                                                                desarrollo['descripcion'], desarrollo['estado'], desarrollo['alcance'], 
                                                                                desarrollo['numero_cdf'], ID_pedido[0], desarrollo['prioridad'],
                                                                                desarrollo['clasificacion'], desarrollo['NUM'])
        my_cursor.execute(statement)
        my_conn.commit()

    except Exception as e:
        print("error", e)
        messagebox.showinfo(message="desarrollo NO registrado.", title="Aviso del sistema")
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
        statement = '''INSERT INTO tareas (ID_proyecto, responsable, nombre_tarea, descripcion, avance, HH_dedicadas, fecha_inicio, fecha_fin, estado)
        VALUES('{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}')'''.format( ID_proyecto[0], tarea['responsable'], tarea['nombre_tarea'],
                                                                                tarea['descripcion'], tarea['avance'], tarea['HH_dedic'], 
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
    messagebox.showinfo(message="Nuevo desarrollo registrado.", title="Aviso del sistema")

def main():
    global df
    df = pd.read_excel("fuente.xlsx")
    DES_charge(6)
    '''
    DES_charge(7)
    DES_charge(8)
    DES_charge(9)
    DES_charge(10)
    DES_charge(11)
    DES_charge(12)
    DES_charge(13)
    DES_charge(14)
    DES_charge(15)
    '''
if __name__ == '__main__':
    main()
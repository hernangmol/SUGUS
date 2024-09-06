from tkinter import messagebox
import mysql.connector
from mysql.connector import RefreshOption
from datetime import datetime
import time

status = 'init'
times_on = 0
times_off = 0

def connect_db():
    global conn
    global status
    try:
        conn = mysql.connector.connect(host = '192.168.100.105',
                                        port = 3306,
                                        #database = "BDNP",  # base de produccion
                                        database = "BDNP_t",  # base de test
                                        user = "admin",
                                        password = "cenadif2023",
                                        )
    except Exception as e:
        #if messagebox.askretrycancel(message="Error de conexión con base de datos.", title="SUGUS - Frontend BDNP"):
        #status = 'off'
        #print("error connecting", e)
        #exit(1)
        #return()
        pass
    else:
        cursor = conn.cursor()
        status = 'on'
        return (cursor) 


def main():
    global conn ######################################### poner en main (una vez)
    global times_on
    global times_off
    global status
    while True:
        if status == 'init':
            print('Init:', datetime.now())
            status = 'on'
        elif status == 'on':
            try:
                cursor = connect_db() ########################### poner (cada vez)
                statement = "SELECT 1"
                cursor.execute(statement)
                resultado = cursor.fetchone()
                #print (resultado[0])
                if resultado[0] == 1:
                    times_on +=1
                    #print(' ')
                    print('  ', times_on, end = '\r')
                    disconnect_db(cursor) ########################### poner (cada vez)
            except Exception as e:
                #print("error", e)
                print('times_on:', times_on)
                print('Disconnected:', datetime.now())
                status = 'off'
                #exit(1)
        elif status == 'off':
            try:
                cursor = connect_db() ########################### poner (cada vez)
                statement = "SELECT 1"
                cursor.execute(statement)
                resultado = cursor.fetchone()
                if resultado[0] == 1:
                    print('times_off:', times_off)
                    print('Reconnected:', datetime.now())
                    exit(1)
                
            except Exception as e:
                times_off +=1
                print('  ', times_off, end = '\r')
                #print("error", e)
                #exit(1)
        #print(datetime.now())
        time.sleep(1)

def disconnect_db(cursor):
    cursor.close()
    conn.close()

if __name__ == "__main__":
        main() 
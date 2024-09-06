import time
from threading import Thread
import mysql.connector
from datetime import datetime
import datetime as dt
from tkinter import messagebox
import os

def test(s):
    print(s)
    while True:
        try:
            #my_cursor = my_conn.cursor()
            statement = "SELECT 1"
            my_cursor.execute(statement)
            resultados = my_cursor.fetchall()
        except Exception as e:
            print("error THREAD", e)
            print(datetime.now())
            #exit(1)
            os._exit(1)
        print('THREAD - BDNP: Ok')
        time.sleep(10)

def main():

    global my_conn
    global my_cursor
    try:
        my_conn = mysql.connector.connect(host = '192.168.100.105',
                                        port = 3306,
                                        database = "BDNP",  # base de produccion
                                        #database = "BDNP_t",  # base de test
                                        user = "admin",
                                        password = "cenadif2023")
        print("conexion OK!", datetime.now())
        #break
    except Exception as e:
        if messagebox.askretrycancel(message="Error de conexión con base de datos.", title="SUGUS - Frontend BDNP"):
            pass
        print("error INIT", e)
        exit(1)
    my_cursor = my_conn.cursor()
    print("- MAIN INIT ENDED")
    time.sleep(5)
    t = Thread(target=test, args=("- THREAD STARTED",)) # always keep the , even when there is only one argument!!
    t.start()


    while True:
        try:
            statement = "SELECT 1"
            my_cursor.execute(statement)
            resultados = my_cursor.fetchall()
        except Exception as e:
            print("error MAIN", e)
            print(datetime.now())
            #exit(1)
            os._exit(1)
        print('MAIN - BDNP = Ok')
        time.sleep(60)

if __name__ == '__main__':
    main()
from tkinter import messagebox
import mysql.connector
import time

def connect_db():
    global conn
    try:
        conn = mysql.connector.connect(host = '192.168.5.92',
                                        port = 3306,
                                        database = "BDNP",  # base de produccion
                                        #database = "BDNP_t",  # base de test
                                        user = "admin",
                                        password = "cenadif2023",
                                        )
    except Exception as e:
        if messagebox.askretrycancel(message="Error de conexión con base de datos.", title="SUGUS - Frontend BDNP"): 
            pass
        else:
            exit(1)
    else:
        cursor = conn.cursor(buffered=True)
        #print ('SUGUS_conn - connection OK')
        return [conn, cursor] 

def disconnect_db(conn, cursor):
    cursor.close()
    conn.close()

def main():
    global conn ######################################### poner en main (una vez)
    print('Connection test running - "Crtl + C" to stop')
    while True:
        cursor = connect_db() ########################### poner (cada vez)
        try:
            statement = "SELECT 1"
            cursor.execute(statement)
            resultado = cursor.fetchone()
            disconnect_db(cursor) ########################### poner (cada vez)
        except Exception as e:
            print('SUGUS_conn - DB NOK')
        else:
            print('SUGUS_conn - DB OK')   
                
        time.sleep(1)

if __name__ == "__main__":
        main() 
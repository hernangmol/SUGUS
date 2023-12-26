import BDNP as bd
import mysql.connector

def main():
    global my_conn

    try:
        my_conn = mysql.connector.connect(host = '192.168.100.105',
                                        port = 3306,
                                        #database = "BDNP",  # base de produccion
                                        database = "BDNP_t",  # base de test
                                        user = "admin",
                                        password = "cenadif2023")
    except Exception as e:
        print("error de conexion!")
            
    try:
            my_cursor = my_conn.cursor()
            statement = "SELECT ID_usuarios, contraseña FROM usuarios"
            my_cursor.execute(statement)
            resultados = my_cursor.fetchall()
            #print(resultados)
    except Exception as e:
        print("error", e)

    for u in resultados:
        aux = bd.F_hash(u[1])
        print(aux)
        try:
            my_cursor = my_conn.cursor()
            statement = "UPDATE usuarios SET hash = %s WHERE  ID_usuarios = %s"
            values = (aux, u[0])
            my_cursor.execute(statement, values)
            my_conn.commit()
        except Exception as e:
            print("error", e)

if __name__ == '__main__':
     main()
import hashlib
from screeninfo import get_monitors

class display:
    def __init__(self):
        m = get_monitors()
        self.width = m[0].width
        self.height = m[0].height
        self.full_geometry = str(self.width) + 'x' + str(self.height) + '+0+0'  
    
class usuario:

    def __init__(self, ID_usuario, nombre_usuario, nombres, apellidos, correo, rol):
        if 10000000 < ID_usuario < 100000000:
            self.ID_usuario = ID_usuario
        else:
            pass
            # mensaje de error y return    
        if len(nombre_usuario) <= 20:
            self.nombre_usuario = nombre_usuario
        else:
            pass
            # mensaje de error y return
        if len(nombres) <= 20:
            self.nombres = nombres
        else:
            pass
            # mensaje de error y return
        
        self.apellidos = apellidos
        self.correo = correo
        #self.contraseña = '1234'
        self.hash = F_hash('1234')
        self.rol = rol
    
def F_hash(password):
    
    # adding 5gz as password
    salt = "5gz"
    
    # Adding salt at the last of the password
    dataBase_password = password+salt
    # Encoding the password
    hashed = hashlib.sha256(dataBase_password.encode())
    
    # returning the Hash
    return(hashed.hexdigest())
def main():
    ###### pruebas
    
    # display
    d = display()
    print(d.full_geometry)

    pass

if __name__  == "__main__":
    main()
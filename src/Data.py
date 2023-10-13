import pymysql


class Data():

    def __init__(self, host, user, password, database):
        self.conexion = pymysql.connect(
            host=host,
            user=user,
            password=password,
            database=database
        )

    def mostrar(self):
        cursor = self.conexion.cursor()
        sql = "SELECT * FROM Registros "
        cursor.execute(sql)
        registro = cursor.fetchall()
        cursor.close()
        self.conexion.commit()
        return registro

    def insertar(self, cedula, estatus, nombre, estado, municipio, parroquia, centro, direccion, id):
        cur = self.conexion.cursor()
        sql = '''INSERT INTO Registros (Id, Cedula, Estatus, Nombre, Estado, Municipio, Parroquia, Centro, Direccion)
        VALUES( '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}', '{}')'''.format(id, cedula, estatus, nombre, estado, municipio, parroquia, centro, direccion)
        cur.execute(sql)
        self.conexion.commit()
        cur.close()

    def buscar(self, cedula):
        cur = self.conexion.cursor()
        sql = "SELECT * FROM Registros WHERE cedula = {}".format(cedula)
        cur.execute(sql)
        act = cur.fetchall()
        self.conexion.commit()
        cur.close()
        return act

    def buscar_e(self, estatus):
        cur = self.conexion.cursor()
        sql = "SELECT * FROM Registros WHERE ESTATUS = %s"
        cur.execute(sql, (estatus,))
        resultados = cur.fetchall()
        self.conexion.commit()
        cur.close()
        return resultados

    def eliminar(self, estatus):
        cur = self.conexion.cursor()
        sql = '''DELETE FROM Registros WHERE ESTATUS = {}'''.format(estatus)
        cur.execute(sql)
        self.conexion.commit()
        cur.close()

    def actualizar(self, id, cedula, estatus, nombre, estado, municipio, parroquia, centro, direccion):
        cur = self.conexion.cursor()
        sql = '''UPDATE Registros SET Cedula = %s, Estatus = %s, Nombre = %s, Estado = %s, Municipio = %s,
        Parroquia = %s, Centro = %s, Direccion = %s WHERE ID = %s'''
        cur.execute(sql, (cedula, estatus, nombre, estado,
                    municipio, parroquia, centro, direccion, id))
        a = cur.rowcount
        self.conexion.commit()
        cur.close()
        return a

    def obtener_estatus(self):
        cursor = self.conexion.cursor()
        sql = "SELECT DISTINCT ESTATUS FROM Registros"
        cursor.execute(sql)
        estatus = cursor.fetchall()
        self.conexion.commit()
        cursor.close()

        # Obtener solo los valores de la columna ESTATUS
        estatus = [e[0] for e in estatus]

        return estatus

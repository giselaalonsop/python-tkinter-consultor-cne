import subprocess
import pymysql
import tkinter as tk
from tkinter import messagebox
import os
import urllib.request
from tkinter.ttk import Progressbar
from tqdm import tqdm
import os
import psutil
import stat


class Conexion:
    def __init__(self):
        self.config = {
            'host': 'localhost',
            'port': 3306,
            'user': 'root',
            'password': '',
            'database': 'Data'
        }

    def check_connection(self):
        try:
            connection = pymysql.connect(
                host=self.config['host'],
                user=self.config['user'],
                password=self.config['password'],
                database=self.config['database']
            )
            connection.close()
            return True
        except pymysql.err.OperationalError:
            return False

    def install_xampp(self):
        try:
            # Verificar si XAMPP está instalado en el sistema
            if not self.check_xampp_installed():
                # Descargar el instalador de XAMPP
                url = "https://sourceforge.net/projects/xampp/files/XAMPP%20Windows/8.2.4/xampp-windows-x64-8.2.4-0-VS16-installer.exe/download"
                response = urllib.request.urlopen(url)
                total_size = int(response.headers['Content-Length'])
                script_dir = os.path.dirname(os.path.abspath(__file__))
                output_file = os.path.join(script_dir, "xampp-installer.exe")

                # Verificar los permisos del directorio
                if not os.access(script_dir, os.W_OK):
                    # Cambiar los permisos del directorio para permitir escritura
                    os.chmod(script_dir, stat.S_IRWXU)

                # Crear la ventana emergente para mostrar el progreso
                progress_window = tk.Tk()
                progress_window.title("Descargando XAMPP")

                # Obtener las dimensiones de la pantalla
                screen_width = progress_window.winfo_screenwidth()
                screen_height = progress_window.winfo_screenheight()

                # Calcular las coordenadas para centrar la ventana emergente
                window_width = 300
                window_height = 100
                x = (screen_width - window_width) // 2
                y = (screen_height - window_height) // 2

                # Establecer las coordenadas de la ventana emergente
                progress_window.geometry(
                    f"{window_width}x{window_height}+{x}+{y}")

                # Crear la barra de progreso centrada en la ventana
                progress_bar = Progressbar(
                    progress_window, length=250, mode='determinate')
                progress_bar.pack(pady=20)

                # Descargar el archivo con progreso
                with open(output_file, 'wb') as file:
                    pbar = tqdm(total=total_size, unit='B', unit_scale=True)
                    while True:
                        chunk = response.read(1024)
                        if not chunk:
                            break
                        file.write(chunk)
                        pbar.update(len(chunk))

                        # Actualizar la barra de progreso en la ventana
                        progress = int(pbar.n / pbar.total * 100)
                        progress_bar['value'] = progress
                        progress_window.update()

                    pbar.close()

                # Establecer los permisos del archivo descargado
                os.chmod(output_file, stat.S_IRUSR |
                         stat.S_IWUSR | stat.S_IXUSR)

                # Ejecutar el instalador de XAMPP
                subprocess.run([output_file])

                messagebox.showinfo("Instalación exitosa",
                                    "XAMPP se ha instalado correctamente.")
            else:
                messagebox.showinfo("XAMPP ya está instalado",
                                    "XAMPP ya está instalado en el sistema.")
        except Exception as e:
            messagebox.showerror("Error de instalación",
                                 f"No se pudo instalar XAMPP: {str(e)}")

    def check_xampp_installed(self):
        xampp_paths = [
            'C:\\xampp',
            'C:\\Archivos de programa\\xampp',
            'C:\\Archivos de programa (x86)\\xampp',
            'C:\\Program Files\\xampp',
            'C:\\Program Files (x86)\\xampp',
            '/opt/lampp',
            '/usr/local/lampp',
            'C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\XAMPP'
        ]

        for path in xampp_paths:
            if os.path.exists(path):
                return True

        return False

    def start_mysql_service(self):
        try:
            # Verificar si el servicio de MySQL está activo
            if not self.check_mysql_service_active():
                # Obtener la ruta del archivo mysql_start.bat en la instalación de XAMPP
                xampp_path = self.get_xampp_installation_path()
                mysql_start_path = os.path.join(xampp_path, 'mysql_start.bat')

                # Ejecutar el archivo mysql_start.bat para iniciar el servicio de MySQL
                subprocess.run(mysql_start_path, shell=True)

                messagebox.showinfo("Servicio de MySQL iniciado",
                                    "El servicio de MySQL se ha iniciado correctamente.")
            else:
                messagebox.showinfo("Servicio de MySQL ya está activo",
                                    "El servicio de MySQL ya está activo.")
        except Exception as e:
            messagebox.showerror("Error de inicio de servicio",
                                 f"No se pudo iniciar el servicio de MySQL: {str(e)}")

    def get_xampp_installation_path(self):
        possible_paths = [
            'C:\\xampp',
            'C:\\Archivos de programa\\xampp',
            'C:\\Archivos de programa (x86)\\xampp',
            'C:\\Program Files\\xampp',
            'C:\\Program Files (x86)\\xampp',
            '/opt/lampp',
            '/usr/local/lampp',
            'C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\XAMPP'
        ]

        for path in possible_paths:
            if os.path.exists(path):
                return path

        raise Exception("La ruta de instalación de XAMPP no se encontró.")

    def create_database(self):
        try:
            # Establecer conexión sin seleccionar una base de datos
            connection = pymysql.connect(
                host=self.config['host'],
                port=self.config['port'],
                user=self.config['user'],
                password=self.config['password']
            )

            # Crear base de datos
            with connection.cursor() as cursor:
                cursor.execute(
                    f"CREATE DATABASE IF NOT EXISTS {self.config['database']}")

            connection.commit()
            connection.close()

            messagebox.showinfo("Base de datos creada",
                                "La base de datos se ha creado correctamente.")

            # Guardar el indicador de conexión en un archivo o registro
            with open('database_setup.txt', 'w') as file:
                file.write('connection_setup')

        except pymysql.Error as e:
            messagebox.showerror("Error de base de datos",
                                 f"No se pudo crear la base de datos: {str(e)}")

    def get_connection_data(self):
        if self.check_connection():
            return self.config['user'], self.config['host'], self.config['password'], self.config['database']
        else:
            # No hay conexión
            # No hay conexión, verificar si XAMPP está instalado
            if not self.check_xampp_installed():
                self.install_xampp()

            # Verificar si el servicio de MySQL está activo
            if not self.check_mysql_service_active():
                self.start_mysql_service()

            # Crear la base de datos y establecer la conexión
            self.create_database()
            self.create_tables()

            return self.config['user'], self.config['host'], self.config['password'], self.config['database']

    def check_mysql_service_active(self):
        for process in psutil.process_iter(['name']):
            if process.info['name'] == 'mysqld.exe':
                return True

        return False

    def create_tables(self):
        try:
            # Establecer conexión a la base de datos
            connection = pymysql.connect(
                host=self.config['host'],
                port=self.config['port'],
                user=self.config['user'],
                password=self.config['password'],
                database=self.config['database']
            )

            # Crear tabla Cedulas
            create_table_query1 = '''
            CREATE TABLE IF NOT EXISTS Cedulas (
                id INT AUTO_INCREMENT PRIMARY KEY,
                cedula VARCHAR(20)
            )
            '''
            with connection.cursor() as cursor:
                cursor.execute(create_table_query1)

            # Crear tabla Registros
            create_table_query2 = '''
            CREATE TABLE IF NOT EXISTS Registros (
                id INT AUTO_INCREMENT PRIMARY KEY,
                cedula VARCHAR(20),
                estatus VARCHAR(50),
                nombre VARCHAR(50),
                estado VARCHAR(50),
                municipio VARCHAR(50),
                parroquia VARCHAR(50),
                centro VARCHAR(50),
                direccion VARCHAR(100)
            )
            '''
            with connection.cursor() as cursor:
                cursor.execute(create_table_query2)

            connection.commit()
            connection.close()

            messagebox.showinfo(
                "Tablas creadas", "Las tablas se han creado correctamente.")
        except pymysql.Error as e:
            messagebox.showerror(
                "Error de tablas", f"No se pudieron crear las tablas: {str(e)}")

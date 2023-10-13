from tkinter import Tk, Button, Entry, Label, ttk, PhotoImage
from tkinter import StringVar, Scrollbar, Frame
from principal import*
from Data import Data
import time
from click import progressbar
from tkinter import messagebox
from tkinter import ttk
import textwrap
from tkinter import Tk, Text, Scrollbar
import sys
from Conexion import Conexion


# https://colorhunt.co/palette/f0e5cff7f6f2c8c6c64b6587


class Ventana(Frame):
    def __init__(self, master, *args):
        super().__init__(master, *args)
        self.menu = True
        self.color = True
        self.cedula = StringVar()
        self.nombre = StringVar()
        self.estatus = StringVar()
        self.estado = StringVar()
        self.municipio = StringVar()
        self.parroquia = StringVar()
        self.centro = StringVar()
        self.direccion = StringVar()
        self.buscar_actualiza = StringVar()
        self.id = StringVar()
        host, user, password, database = Conexion().get_connection_data()
        self.base_datos = Data(user, host, password, database)

        self.frame_inicio = Frame(self.master, bg='white', width=50, height=45)
        self.frame_inicio.grid_propagate(0)
        self.frame_inicio.grid(column=0, row=0, sticky='nsew')
        self.frame_menu = Frame(self.master, bg='white', width=50)
        self.frame_menu.grid_propagate(0)
        self.frame_menu.grid(column=0, row=1, sticky='nsew')
        self.frame_top = Frame(self.master, bg='white', height=50)
        self.frame_top.grid(column=1, row=0, sticky='nsew')
        self.frame_principal = Frame(self.master, bg='white')
        self.frame_principal.grid(column=1, row=1, sticky='nsew')
        self.master.columnconfigure(1, weight=1)
        self.master.rowconfigure(1, weight=1)
        self.frame_principal.columnconfigure(0, weight=1)
        self.frame_principal.rowconfigure(0, weight=1)
        self.widgets()

    def vaciar(self):
        tabla1 = "registros"
        tabla2 = "cedulas"
        vaciar_tabla(tabla1)
        vaciar_tabla(tabla2)
        # Mostrar un mensaje en una ventana emergente
        self.mensange = messagebox.showinfo(
            "Data", "Operacion exitosa, importe un nuevo archivo de cedulas para realizar mas consultas")
        estatus = self.base_datos.obtener_estatus()
        self.selector_estatus['values'] = estatus

    def importar(self):

        # Crear una barra de progreso debajo del botón de la clase Ventana
        self.progress_bar = ttk.Progressbar(
            self.frame_tres, length=200, mode='determinate')
        self.progress_bar.grid(columnspan=2, column=0, row=11, pady=10)

        # Llamar a la función de importación y pasar la instancia de la clase Ventana
        importar_data(self)
        self.progress_bar.destroy()

    def consulta_(self):
        self.progress_bar = ttk.Progressbar(
            self.frame_tres, length=200, mode='determinate')
        self.progress_bar.grid(columnspan=2, column=0, row=11, pady=10)
        self.progreso_text = Entry(self.frame_tres, width=12)
        self.progreso_text.grid(column=0, row=12, pady=10, columnspan=3)

        # Llamar a la función de importación y pasar la instancia de la clase Ventana
        consulta(self)
        self.progress_bar.destroy()
        self.progreso_text.destroy()

    def pantalla_inicial(self):
        self.paginas.select([self.frame_uno])

    def pantalla_datos(self):
        self.paginas.select([self.frame_dos])
        self.frame_dos.columnconfigure(0, weight=1)
        self.frame_dos.columnconfigure(1, weight=1)
        self.frame_dos.rowconfigure(2, weight=1)
        self.frame_tabla_uno.columnconfigure(0, weight=1)
        self.frame_tabla_uno.rowconfigure(0, weight=1)

    def pantalla_escribir(self):
        self.paginas.select([self.frame_tres])
        self.frame_tres.columnconfigure(0, weight=1)
        self.frame_tres.columnconfigure(1, weight=1)

    def pantalla_actualizar(self):
        self.paginas.select([self.frame_cuatro])
        self.frame_cuatro.columnconfigure(0, weight=1)
        self.frame_cuatro.columnconfigure(1, weight=1)

    def pantalla_buscar(self):
        self.paginas.select([self.frame_cinco])
        self.frame_cinco.columnconfigure(0, weight=1)
        self.frame_cinco.columnconfigure(1, weight=1)
        self.frame_cinco.columnconfigure(2, weight=1)
        self.frame_cinco.rowconfigure(2, weight=1)
        self.frame_tabla_dos.columnconfigure(0, weight=1)
        self.frame_tabla_dos.rowconfigure(0, weight=1)

    def pantalla_ajustes(self):
        self.paginas.select([self.frame_seis])

    def menu_lateral(self):
        if self.menu is True:
            for i in range(50, 190, 10):
                self.frame_menu.config(width=i)
                self.frame_inicio.config(width=i)
                self.frame_menu.update()
                clik_inicio = self.bt_cerrar.grid_forget()
                if clik_inicio is None:
                    self.bt_inicio.grid(column=0, row=0, padx=10, pady=10)
                    self.bt_inicio.grid_propagate(0)
                    self.bt_inicio.config(width=i)
                    self.pantalla_inicial()
            self.menu = False
        else:
            for i in range(170, 50, -10):
                self.frame_menu.config(width=i)
                self.frame_inicio.config(width=i)
                self.frame_menu.update()
                clik_inicio = self.bt_inicio.grid_forget()
                if clik_inicio is None:
                    self.frame_menu.grid_propagate(0)
                    self.bt_cerrar.grid(column=0, row=0, padx=10, pady=10)
                    self.bt_cerrar.grid_propagate(0)
                    self.bt_cerrar.config(width=i)
                    self.pantalla_inicial()
            self.menu = True

    def widgets(self):
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(
            os.path.abspath(__file__)))
        logo_path1 = os.path.join(base_path, "menu.png")
        logo_path2 = os.path.join(base_path, "datos.png")
        logo_path3 = os.path.join(base_path, "escribir.png")
        logo_path4 = os.path.join(base_path, "actualizar.png")
        logo_path5 = os.path.join(base_path, "buscar.png")
        logo_path6 = os.path.join(base_path, "configuracion.png")
        self.imagen_inicio = PhotoImage(
            file=logo_path1)
        self.imagen_menu = PhotoImage(
            file=logo_path1)
        self.imagen_datos = PhotoImage(
            file=logo_path2)
        self.imagen_registrar = PhotoImage(
            file=logo_path3)
        self.imagen_actualizar = PhotoImage(
            file=logo_path4)
        self.imagen_buscar = PhotoImage(
            file=logo_path5)
        self.imagen_ajustes = PhotoImage(
            file=logo_path6)

        # self.dia = PhotoImage(file='dia.png')
        # self.noche = PhotoImage(file='noche.png')
        self.bt_inicio = Button(self.frame_inicio, image=self.imagen_inicio,
                                bg='white', activebackground='white', bd=0, command=self.menu_lateral)
        self.bt_inicio.grid(column=0, row=0, padx=5, pady=10)
        self.bt_cerrar = Button(self.frame_inicio, image=self.imagen_menu,
                                bg='white', activebackground='white', bd=0, command=self.menu_lateral)
        self.bt_cerrar.grid(column=0, row=0, padx=5, pady=10)
        # BOTONES Y ETIQUETAS DEL MENU LATERAL
        Button(self.frame_menu, image=self.imagen_datos, bg='white', activebackground='white',
               bd=0, command=self.pantalla_datos).grid(column=0, row=1, pady=20, padx=10)
        Button(self.frame_menu, image=self.imagen_registrar, bg='white', activebackground='white',
               bd=0, command=self.pantalla_escribir).grid(column=0, row=2, pady=20, padx=10)
        Button(self.frame_menu, image=self.imagen_actualizar, bg='white', activebackground='white',
               bd=0, command=self.pantalla_actualizar).grid(column=0, row=3, pady=20, padx=10)
        Button(self.frame_menu, image=self.imagen_buscar, bg='white', activebackground='white',
               bd=0, command=self.pantalla_buscar).grid(column=0, row=4, pady=20, padx=10)
        Button(self.frame_menu, image=self.imagen_ajustes, bg='white', activebackground='white',
               bd=0, command=self.pantalla_ajustes).grid(column=0, row=5, pady=20, padx=10)

        Label(self.frame_menu, text='DATA', bg='white', fg='#C8C6C6', font=(
            'Lucida Sans', 12, 'bold')).grid(column=1, row=1, pady=20, padx=2)
        Label(self.frame_menu, text='REGISTRAR', bg='white', fg='#C8C6C6', font=(
            'Lucida Sans', 12, 'bold')).grid(column=1, row=2, pady=20, padx=2)
        Label(self.frame_menu, text='ACTUALIZAR', bg='white', fg='#C8C6C6', font=(
            'Lucida Sans', 12, 'bold')).grid(column=1, row=3, pady=20, padx=2)
        Label(self.frame_menu, text='FILTRAR', bg='white', fg='#C8C6C6', font=(
            'Lucida Sans', 12, 'bold')).grid(column=1, row=4, pady=20, padx=2)
        Label(self.frame_menu, text='MANUAL', bg='white', fg='#C8C6C6', font=(
            'Lucida Sans', 12, 'bold')).grid(column=1, row=5, pady=20, padx=2)

        #############################  CREAR  PAGINAS  ##############################
        estilo_paginas = ttk.Style()
        estilo_paginas.configure(
            "TNotebook", background='#FBF8F1', foreground='white', padding=0, borderwidth=0)
        estilo_paginas.theme_use('default')
        estilo_paginas.configure(
            "TNotebook", background='white', borderwidth=0)
        estilo_paginas.configure(
            "TNotebook.Tab", background="white", borderwidth=0)
        estilo_paginas.map("TNotebook", background=[("selected", 'white')])
        estilo_paginas.map("TNotebook.Tab", background=[
                           ("selected", 'white')], foreground=[("selected", 'white')])

        # CREACCION DE LAS PAGINAS
        self.paginas = ttk.Notebook(
            self.frame_principal, style='TNotebook')  # , style = 'TNotebook'
        self.paginas.grid(column=0, row=0, sticky='nsew')
        self.frame_uno = Frame(self.paginas, bg='#4B6587')
        self.frame_dos = Frame(self.paginas, bg='white')
        self.frame_tres = Frame(self.paginas, bg='white')
        self.frame_cuatro = Frame(self.paginas, bg='white')
        self.frame_cinco = Frame(self.paginas, bg='white')
        self.frame_seis = Frame(self.paginas, bg='white')
        self.paginas.add(self.frame_uno)
        self.paginas.add(self.frame_dos)
        self.paginas.add(self.frame_tres)
        self.paginas.add(self.frame_cuatro)
        self.paginas.add(self.frame_cinco)
        self.paginas.add(self.frame_seis)
        ##############################         PAGINAS       #############################################

        ######################## FRAME TITULO #################
        self.titulo = Label(self.frame_top, text='APLICACION DE ESCRITORIO ',
                            bg='white', fg='#4B6587', font=('Imprint MT Shadow', 15, 'bold'))
        self.titulo.pack(expand=1)
        ######################## VENTANA PRINCIPAL #################
        Label(self.frame_uno, text='Consulta de Cedulas en el CNE',
              bg='#4B6587', fg='white', font=('Freehand521 BT', 20, 'bold')).pack(expand=1)
        Label(self.frame_uno,  bg='#4B6587').pack(expand=1)
        Button(self.frame_dos, text='ACTUALIZAR', fg='#4B6587', font=('Arial', 11, 'bold'),
               command=self.datos_totales, bg='white', bd=2, borderwidth=2).grid(column=1, row=0, pady=5),

        Button(self.frame_dos, text='Exportar Data', fg='#4B6587', font=('Arial', 11, 'bold'),
               command=exportar_a_excel, bg='white', bd=2, borderwidth=2).grid(column=1, row=1, pady=5)
        # ESTILO DE LAS TABLAS DE DATOS TREEVIEW
        estilo_tabla = ttk.Style()
        estilo_tabla.configure("Treeview", font=('Helvetica', 10, 'bold'),
                               foreground='black',  background='white')  # , fieldbackground='yellow'
        estilo_tabla.map('Treeview', background=[
                         ('selected', '#4B6587')], foreground=[('selected', 'black')])
        estilo_tabla.configure('Heading', background='white',
                               foreground='navy', padding=3, font=('Arial', 10, 'bold'))
        estilo_tabla.configure('Item', foreground='white',
                               focuscolor='#4B6587')
        estilo_tabla.configure('TScrollbar', arrowcolor='#4B6587',
                               bordercolor='black', troughcolor='white', background='white')
        # TABLA UNO
        self.frame_tabla_uno = Frame(self.frame_dos, bg='gray90')
        self.frame_tabla_uno.grid(columnspan=3, row=2, sticky='nsew')
        self.tabla_uno = ttk.Treeview(self.frame_tabla_uno)
        self.tabla_uno.grid(column=0, row=0, sticky='nsew')
        ladox = ttk.Scrollbar(self.frame_tabla_uno,
                              orient='horizontal', command=self.tabla_uno.xview)
        ladox.grid(column=0, row=1, sticky='ew')
        ladoy = ttk.Scrollbar(self.frame_tabla_uno,
                              orient='vertical', command=self.tabla_uno.yview)
        ladoy.grid(column=1, row=0, sticky='ns')

        self.tabla_uno.configure(
            xscrollcommand=ladox.set, yscrollcommand=ladoy.set)
        self.tabla_uno['columns'] = ('Id',
                                     'Cedula', 'Estatus', 'Nombre', 'Estado', 'Municipio', 'Parroquia', 'Centro', 'Direccion')
        self.tabla_uno.column('#0', width=0, stretch=False)
        self.tabla_uno.column('Id', minwidth=100,
                              width=105, anchor='center')
        self.tabla_uno.column('Cedula', minwidth=100,
                              width=130, anchor='center')
        self.tabla_uno.column('Estatus', minwidth=100,
                              width=120, anchor='center')
        self.tabla_uno.column('Nombre', minwidth=100,
                              width=120, anchor='center')
        self.tabla_uno.column('Estado', minwidth=100,
                              width=105, anchor='center')
        self.tabla_uno.column('Municipio', minwidth=100,
                              width=105, anchor='center')
        self.tabla_uno.column('Parroquia', minwidth=100,
                              width=105, anchor='center')
        self.tabla_uno.column('Centro', minwidth=100,
                              width=105, anchor='center')
        self.tabla_uno.column('Direccion', minwidth=100,
                              width=105, anchor='center')

        self.tabla_uno.heading("Id", text="Id", anchor='center')
        self.tabla_uno.heading('Cedula', text='Cedula', anchor='center')
        self.tabla_uno.heading('Estatus', text='Estatus', anchor='center')
        self.tabla_uno.heading("Nombre", text="Nombre", anchor='center')
        self.tabla_uno.heading('Estado', text='Estado', anchor='center')
        self.tabla_uno.heading("Municipio", text="Municipio", anchor='center')
        self.tabla_uno.heading("Parroquia", text="Parroquia", anchor='center')
        self.tabla_uno.heading("Centro", text="Centro", anchor='center')
        self.tabla_uno.heading("Direccion", text="Direccion", anchor='center')

        self.tabla_uno.bind("<<TreeviewSelect>>", self.obtener_fila)

        ######################## Busqueda  #################

        Label(self.frame_tres, text=' Nueva Busqueda ', fg='#4B6587', bg='white', font=(
            'Kaufmann BT', 24, 'bold')).grid(columnspan=2, column=0, row=0, pady=5)

        self.button_importar_cedulas = Button(self.frame_tres, text='EXTRAER CEDULAS', command=self.importar, font=(
            'Rockwell', 13, 'bold'), bg='white', fg='black', relief='groove', bd=3)

        self.button_importar_cedulas.grid(
            columnspan=2, column=0, row=10, pady=20)

        self.button_importar_cedulas = Button(self.frame_tres, text='CONSULTAR CEDULAS', command=self.consulta_, font=(
            'Rockwell', 13, 'bold'), bg='white', fg='black', relief='groove', bd=3)
        self.button_importar_cedulas.grid(
            columnspan=2, column=0, row=11, pady=20)

        Label(self.frame_tres,  bg='white').grid(
            column=3, rowspan=5, row=0, padx=50)
        self.aviso_guardado = Label(
            self.frame_tres, bg='white', font=('Comic Sans MS', 12), fg='black')
        self.aviso_guardado.grid(columnspan=2, column=0, row=6, padx=5)

        # Actualizar Datos
        ######################## ACTUALIZAR DATOS #################

        Label(self.frame_cuatro, text='Actualizar Datos', fg='black', bg='white', font=(
            'Kaufmann BT', 24, 'bold')).grid(columnspan=4, row=0)
        Label(self.frame_cuatro, text='Ingrese la cedula que desea  actualizar',
              fg='black', bg='white', font=('Roboto Condensed', 12, 'bold')).grid(column=0, row=1, pady=15)

        self.aviso_actualizado = Label(
            self.frame_cuatro, fg='black', bg='white', font=('Arial', 12, 'bold'))
        self.aviso_actualizado.grid(columnspan=2, row=7, pady=10, padx=5)

        Label(self.frame_cuatro, text='Id', fg='black', bg='white', font=(
            'Rockwell', 13, 'bold')).grid(column=0, row=2, pady=15)
        Label(self.frame_cuatro, text='Cedula', fg='black', bg='white', font=(
            'Rockwell', 13, 'bold')).grid(column=0, row=3, pady=15, padx=10)
        Label(self.frame_cuatro, text='Estatus', fg='black', bg='white',
              font=('Rockwell', 13, 'bold')).grid(column=0, row=4, pady=15)
        Label(self.frame_cuatro, text='Nombre', fg='black', bg='white',
              font=('Rockwell', 13, 'bold')).grid(column=0, row=5, pady=15)
        Label(self.frame_cuatro, text='Estado', fg='black', bg='white',
              font=('Rockwell', 13, 'bold')).grid(column=0, row=6, pady=15)
        Label(self.frame_cuatro, text='Municipio', fg='black', bg='white', font=(
            'Rockwell', 13, 'bold')).grid(column=0, row=7, pady=15)
        Label(self.frame_cuatro, text='Parroquia', fg='black', bg='white', font=(
            'Rockwell', 13, 'bold')).grid(column=0, row=8, pady=15)
        Label(self.frame_cuatro, text='Centro', fg='black', bg='white', font=(
            'Rockwell', 13, 'bold')).grid(column=0, row=9, pady=15)
        Label(self.frame_cuatro, text='Direccion', fg='black', bg='white', font=(
            'Rockwell', 13, 'bold')).grid(column=0, row=10, pady=15)

        Entry(self.frame_cuatro, textvariable=self.buscar_actualiza, font=('Comic Sans MS', 12),
              highlightbackground="#4B6587", width=12, highlightthickness=5).grid(column=1, row=1, pady=30)

        Entry(self.frame_cuatro, textvariable=self.id, font=('Comic Sans MS', 12),
              highlightbackground="#4B6587", highlightthickness=5).grid(column=1, row=2)

        Entry(self.frame_cuatro, textvariable=self.cedula, font=('Comic Sans MS', 12),
              highlightbackground="#4B6587", highlightthickness=5).grid(column=1, row=3)
        Entry(self.frame_cuatro, textvariable=self.estatus, font=('Comic Sans MS', 12),
              highlightbackground="#4B6587", highlightthickness=5).grid(column=1, row=4)
        Entry(self.frame_cuatro, textvariable=self.nombre, font=('Comic Sans MS', 12),
              highlightbackground="#4B6587", highlightthickness=5).grid(column=1, row=5)
        Entry(self.frame_cuatro, textvariable=self.estado, font=('Comic Sans MS', 12),
              highlightbackground="#4B6587", highlightthickness=5).grid(column=1, row=6)
        Entry(self.frame_cuatro, textvariable=self.municipio, font=('Comic Sans MS', 12),
              highlightbackground="#4B6587", highlightthickness=5).grid(column=1, row=7)
        Entry(self.frame_cuatro, textvariable=self.parroquia, font=('Comic Sans MS', 12),
              highlightbackground="#4B6587", highlightthickness=5).grid(column=1, row=8)
        Entry(self.frame_cuatro, textvariable=self.centro, font=('Comic Sans MS', 12),
              highlightbackground="#4B6587", highlightthickness=5).grid(column=1, row=9)
        Entry(self.frame_cuatro, textvariable=self.direccion, font=('Comic Sans MS', 12),
              highlightbackground="#4B6587", highlightthickness=5).grid(column=1, row=10)

        Button(self.frame_cuatro, command=self.actualizar_datos, text='BUSCAR', font=(
            'Arial', 12, 'bold'), bg='white').grid(column=2, columnspan=1, row=1, pady=10, padx=100)
        Button(self.frame_cuatro, command=self.actualizar_tabla, text='ACTUALIZAR', font=(
            'Arial', 12, 'bold'), bg='white').grid(column=2, columnspan=1, row=7, pady=10, padx=100)
        Label(self.frame_cuatro,  bg='white').grid(
            column=2, columnspan=2, rowspan=5, row=1, padx=2)
        Button(self.frame_cuatro, command=self.vaciar, text='Borrar data', font=(
            'Arial', 12, 'bold'), bg='white').grid(column=2, columnspan=1, row=8, pady=10, padx=100)

        ######################## BUSCAR  #################

        Label(self.frame_cinco, text='Filtro de busqueda', fg='black', bg='white', font=(
            'Kaufmann BT', 24, 'bold')).grid(columnspan=4, row=0, sticky='nsew', padx=2)

        self.selector_estatus = ttk.Combobox(
            self.frame_cinco, state='readonly')
        self.selector_estatus['font'] = ('Comic Sans MS', 12)

        # Crear un estilo personalizado y configurar el color de fondo
        style = ttk.Style()
        style.configure('Custom.TCombobox', background='white',
                        selectbackground='black')

        # Asignar el estilo personalizado al combobox
        self.selector_estatus['style'] = 'Custom.TCombobox'

        self.selector_estatus.grid(column=0, row=1, sticky='nsew', padx=2)
        # Obtener los estatus desde la base de datos

        estatus = self.base_datos.obtener_estatus()
        self.selector_estatus['values'] = estatus

        Button(self.frame_cinco, command=self.buscar_estatus, text='BUSCAR POR ESTATUS', font=(
            'Arial', 8, 'bold'), bg='#C8C6C6').grid(column=1, row=1, sticky='nsew', padx=2)

        # TABLA DOS
        self.frame_tabla_dos = Frame(self.frame_cinco, bg='gray90')
        self.frame_tabla_dos.grid(columnspan=4, row=2, sticky='nsew')
        self.tabla_dos = ttk.Treeview(self.frame_tabla_dos)
        self.tabla_dos.grid(column=0, row=0, sticky='nsew')
        ladox = ttk.Scrollbar(self.frame_tabla_dos,
                              orient='horizontal', command=self.tabla_dos.xview)
        ladox.grid(column=0, row=1, sticky='ew')
        ladoy = ttk.Scrollbar(self.frame_tabla_dos,
                              orient='vertical', command=self.tabla_dos.yview)
        ladoy.grid(column=1, row=0, sticky='ns')

        self.tabla_dos.configure(
            xscrollcommand=ladox.set, yscrollcommand=ladoy.set,)
        self.tabla_dos['columns'] = (
            'Cedula', 'Estatus', 'Nombre', 'Estado', 'Municipio', 'Parroquia', 'Centro', 'Direccion', 'Id')
        self.tabla_dos.column('#0', minwidth=100, width=120, anchor='center')
        self.tabla_dos.column('Cedula', minwidth=100,
                              width=120, anchor='center')
        self.tabla_dos.column('Estatus', minwidth=100,
                              width=120, anchor='center')
        self.tabla_dos.column('Nombre', minwidth=100,
                              width=120, anchor='center')
        self.tabla_dos.column('Estado', minwidth=100,
                              width=120, anchor='center')
        self.tabla_dos.column('Municipio', minwidth=100,
                              width=120, anchor='center')
        self.tabla_dos.column('Parroquia', minwidth=100,
                              width=120, anchor='center')
        self.tabla_dos.column('Centro', minwidth=100,
                              width=120, anchor='center')
        self.tabla_dos.column('Direccion', minwidth=100,
                              width=120, anchor='center')
        self.tabla_dos.column('Id', minwidth=100,
                              width=120, anchor='center')
        self.tabla_dos.heading('#0', text='Codigo', anchor='center')
        self.tabla_dos.heading('Cedula', text='Cedula', anchor='center')
        self.tabla_dos.heading('Estatus', text='Estatus', anchor='center')
        self.tabla_dos.heading('Nombre', text='Nombre', anchor='center')
        self.tabla_dos.heading('Estado', text='Estado', anchor='center')
        self.tabla_dos.heading('Municipio', text='Municipio', anchor='center')
        self.tabla_dos.heading('Parroquia', text='Parroquia', anchor='center')
        self.tabla_dos.heading('Centro', text='Centro', anchor='center')
        self.tabla_dos.heading('Direccion', text='Direccion', anchor='center')
        self.tabla_dos.heading('Id', text='Id', anchor='center')
        self.tabla_dos.bind("<<TreeviewSelect>>", self.obtener_fila)
        ######################## AJUSTES #################

        # Crear el contenedor Frame
        frame_texto = Frame(self.frame_seis, borderwidth=1,
                            relief='solid', bg="#4B6587")
        frame_texto.pack(side='left', expand=True, fill='both', padx=5, pady=5)

        # Crear el widget de texto dentro del contenedor Frame
        texto_manual = Text(frame_texto, wrap='word', font=(
            "Helvetica", 12), spacing1=2, spacing3=2)
        texto_manual.pack(expand=True)

        # Configurar el borde del contenedor Frame para crear las líneas en la parte inferior y derecha
        frame_texto.grid_propagate(False)

        # Scrollbar para el texto
        scrollbar = Scrollbar(self.frame_seis)
        scrollbar.pack(side='right', fill='y')

        # Asociar el scrollbar al texto
        texto_manual.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=texto_manual.yview)

        # Agregar contenido al manual
        contenido = """                                  MANUAL DE USO 

            Bienvenido al manual de uso del software de extracción, consulta y organización masiva de
            cédulas desde Excel.

            1. Introducción:
                El software ha sido diseñado específicamente para la Alcaldía de San Diego, con el
                objetivo de agilizar y mejorar la gestión de información crucial. Permite extraer
                datos de cédulas desde archivos de Excel y utilizar los servicios del CNE como fuente
                de datos.

            2. Requisitos del sistema:
                Antes de utilizar el software, asegúrese de cumplir con los siguientes requisitos
                del sistema:
                •	 Sistema operativo: Windows 10 o superior
                •	 Microsoft Excel instalado
                •	Conexión a Internet


            3. Interfaz principal:
                La interfaz principal del software consta de los siguientes elementos:
                •	Menú de opciones: Permite acceder a las diferentes funcionalidades del software.
                •	Área de carga de archivos: Permite cargar archivos de Excel que contengan las
                    cédulas a procesar.
                •	Botones de extracción y consulta: Permiten iniciar los procesos de extracción y
                    consulta de cédulas.
                •	Resultados y visualización de datos: Muestra los resultados obtenidos y la
                    visualización de los datos extraídos.

            4. Funcionalidades principales:
                A continuación se describen las principales funcionalidades del software:
                •	Extracción de datos: Permite extraer información de cédulas específicas desde
                    archivos de Excel,
                •	utilizando los servicios del CNE.
                •	Consulta de cédulas: Permite realizar consultas en línea de cédulas específicas
                    utilizando los servicios del CNE.
                •	Organización masiva de datos: Permite organizar los datos extraídos en un formato
                    estructurado y exportarlos a otros formatos, Excel en este caso.
                •	Filtro para la búsqueda de registros con un estatus específico
                •	Edición de registro: Escriba la cedula correspondiente al registro deseado para poder
                    editarlo

            5. Uso del software:
                Para utilizar el software, siga los siguientes pasos:
                1.	Cargue el archivo de Excel que contiene las cédulas a procesar.
                2.	Seleccione la funcionalidad deseada: extracción de datos o consulta de cédulas.
                3.	Inicie el proceso correspondiente y espere a que se completen las operaciones.
                4.	Revise los resultados obtenidos y la visualización de los datos.
                5.	Organice y exporte los datos si es necesario.
                6.	Seleccione la funcionalidad adicional deseada: Filtrar, editar o borrar data.


            6. Recomendaciones y consejos:
                •	Asegúrese de tener una conexión estable a Internet para utilizar las funcionalidades
                    de consulta y extracción.
                •	Revise los resultados y verifique la exactitud de los datos extraídos antes de utilizarlos.
                •	Guarde y respalde regularmente los archivos de Excel utilizados para evitar pérdida de
                    información.

            7. Soporte técnico:
                Si tiene algún problema o duda sobre el uso del software, no dude en contactarme

            ¡Gracias por utilizar nuestro software! Esperamos que esta herramienta sea de gran utilidad
                    para agilizar y mejorar la gestión de información en la Alcaldía de San Diego!.
        """
        margen_izquierdo = 60
        margen_derecho = 50
        margen_superior = 50
        margen_inferior = 100
        texto_manual.tag_configure(
            "margen", lmargin1=margen_izquierdo, lmargin2=margen_izquierdo, rmargin=margen_derecho)
        texto_manual.tag_add("margen", '1.0', 'end')

        # Configurar márgenes inferiores y superiores mediante espaciado
        contenido_formateado = f"\n{' ' * margen_superior}{contenido.strip()}\n{' ' * margen_inferior}"
        texto_manual.insert('1.0', contenido_formateado)

        texto_manual.config(state='disabled')

       #
    def datos_totales(self):
        registros = self.base_datos.mostrar()

        # Limpiar la tabla
        self.tabla_uno.delete(*self.tabla_uno.get_children())

        # Insertar los registros en la tabla
        for registro in registros:
            self.tabla_uno.insert('', 'end', text='', values=registro)
            estatus = self.base_datos.obtener_estatus()
            self.selector_estatus['values'] = estatus

    def actualizar_datos(self):
        dato = self.buscar_actualiza.get()
        dato = str("'" + dato + "'")
        datos_buscados = self.base_datos.buscar(dato)
        if datos_buscados == []:
            self.aviso_actualizado['text'] = 'No existe'
            self.aviso_actualizado.update()
            time.sleep(1)
            self.limpiar_datos()
            self.aviso_actualizado['text'] = ''
        else:
            i = -1
            for dato in datos_buscados:
                i = i + 1
                self.id.set(datos_buscados[i][0])
                self.cedula.set(datos_buscados[i][1])
                self.estatus.set(datos_buscados[i][2])
                self.nombre.set(datos_buscados[i][3])
                self.estado.set(datos_buscados[i][4])
                self.municipio.set(datos_buscados[i][5])
                self.parroquia.set(datos_buscados[i][6])
                self.centro.set(datos_buscados[i][7])
                self.direccion.set(datos_buscados[i][8])

    def actualizar_tabla(self):
        id = self.id.get()
        cedula = self.cedula.get()
        estatus = self.estatus.get()
        nombre = self.nombre.get()
        estado = self.estado.get()
        municipio = self.municipio.get()
        parroquia = self.parroquia.get()
        centro = self.centro.get()
        direccion = self.direccion.get()
        self.base_datos.actualizar(id,
                                   cedula, estatus, nombre, estado, municipio, parroquia, centro, direccion)
        time.sleep(3)
        self.mensage = messagebox.showinfo(
            "Actualizacion", "Sus cambios fueron procesados correctamente")

        time.sleep(1)
        self.limpiar_datos()
        self.buscar_actualiza.set('')
        estatus = self.base_datos.obtener_estatus()
        self.selector_estatus['values'] = estatus

    def limpiar_datos(self):
        self.id.set('')
        self.cedula.set('')
        self.estatus.set('')
        self.nombre.set('')
        self.estado.set('')
        self.municipio.set('')
        self.parroquia.set('')
        self.centro.set('')
        self.direccion.set('')

    def buscar_estatus(self):
        estatus_seleccionado = self.selector_estatus.get()

        estatus_buscado = self.base_datos.buscar_e(estatus_seleccionado)
        if estatus_buscado == []:
            self.indica_busqueda['text'] = 'No existe'
            self.indica_busqueda.update()
            time.sleep(1)
            self.indica_busqueda['text'] = ''
        else:
            self.tabla_dos.delete(*self.tabla_dos.get_children())
            i = -1
            for dato in estatus_buscado:
                i = i + 1
                self.tabla_dos.insert(
                    '', i, text=estatus_buscado[i][0], values=estatus_buscado[i][1:8])

    def eliminar_fila(self):
        fila = self.tabla_dos.selection()
        if len(fila) != 0:
            self.tabla_dos.delete(fila)
            estatus = ("'" + str(self.estatus_borrar) + "'")
            self.base_datos.eliminar(estatus)
            self.indica_busqueda['text'] = 'Eliminado'
            self.indica_busqueda.update()
            self.tabla_dos.delete(*self.tabla_dos.get_children())
            time.sleep(1)
            self.indica_busqueda['text'] = ''
            self.limpiar_datos()
        else:
            self.indica_busqueda['text'] = 'No se Eliminó'
            self.indica_busqueda.update()
            self.tabla_dos.delete(*self.tabla_dos.get_children())
            time.sleep(1)
            self.indica_busqueda['text'] = ''
            self.buscar.set('')
            self.limpiar_datos()

    def obtener_fila(self, event):
        current_item = self.tabla_uno.focus()

        if not current_item:
            return
        data = self.tabla_uno.item(current_item)
        self.nombre_borrar = data['values'][0]

    def guardar_registro(self):
        archivo = open('registros.txt', 'a')
        id = self.id.get()
        cedula = self.cedula.get()
        estatus = self.estatus.get()
        nombre = self.nombre.get()
        estado = self.estado.get()
        municipio = self.municipio.get()
        parroquia = self.parroquia.get()
        centro = self.centro.get()
        direccion = self.direccion.get()
        registro = f"{id},{cedula},{estatus},{nombre},{estado},{municipio},{parroquia},{centro},{direccion}\n"
        archivo.write(registro)
        archivo.close()
        self.aviso_guardado['text'] = 'Registro guardado en registros.txt'
        self.aviso_guardado.update()


# Ejecutar el bucle de la interfaz gráfica
if __name__ == "__main__":
    ventana = Tk()
    ventana.title('')
    ventana.minsize(height=475, width=795)
    ventana.geometry('1000x700+180+80')
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(
        os.path.abspath(__file__)))
    logo_path = os.path.join(base_path, "logo.png")
    ventana.call('wm', 'iconphoto', ventana._w, PhotoImage(
        file=logo_path))
    ventana.title("@autor: Gisela Alonso - Desarrollado en Python")

    app = Ventana(ventana)
    app.mainloop()

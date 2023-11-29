from tkinter import messagebox
from tkinter import *
import pyodbc
import os
import sys
import ldap3 as ldap
from functions import Functions

# Definir colores y tipos
black = '#2C2A29'
white = 'white'
grey = '#E7E6E6'
yellow = '#FDDA24'
font = 'CIBFont Sans'
font_bold = 'CIBFont Sans Bold'

# Dimensiones widgets
dims = {
    'ventana_autenticación': (500, round(500 / (4398 / 2475))),
    'ventana_front': (810, round(810 / (4398 / 2475)))
}


class Authenticator(Functions):

    def __init__(self):
        super().__init__()
        self.n_frame = None
        self.n_window = None
        self.principal_frame = None
        self.username_entry = None
        self.password_entry = None
        self.password_entry_i = None
        self.user_pwd = {} # Diccionario que va a guardar los usuarios con las contraseñas


    def connect_AS400(self, user, pwd, maq = 'NACIONAL'):
        """Realiza conexión con AS400"""
        if maq == 'NACIONAL':
            conn = pyodbc.connect('driver={iSeries Access ODBC Driver};system=10.9.2.201;UID='+str(user)+';PWD='+str(pwd)+';Trusted_Connection=no;autocommit=True', autocommit = True)
        elif maq == 'MEDELLIN':
            conn = pyodbc.connect('driver={iSeries Access ODBC Driver};system=MEDELLINET01;UID='+str(user)+';PWD='+str(pwd)+';Trusted_Connection=no')
        else:
            raise Exception('La máquina especificada no es válida')

        return conn


    def clear_frame(self):
        """limpia el formulario"""
        self.principal_frame.pack_forget()
        self.principal_frame.destroy()

    
    def fin(func):
        # Continuar cuando ya no queden más autenticaciones por realizar
        def wrapper(self, *args, **kwargs):
            try:
                user, pwd = func(self, *args, **kwargs)
            except Exception as e:
                messagebox.showerror(title = 'Error', message = 'No fue posible conectarse, error: ' + str(e))
            else:
                try:
                    messagebox.showinfo(title = 'Ingreso exitoso', message = 'Has ingresado exitosamente')
                    self.user_pwd[func.__name__ + '_user'] = user
                    self.user_pwd[func.__name__ + '_pwd'] = pwd
                    self.clear_frame()
                    next(app)
                except StopIteration:
                    self.logger.info('Autenticación exitosa')
                    if self.n_frame == None:
                        self.n_window.destroy()
                    else:
                        self.n_frame.pack(expand = True, fill = 'both')
                        self.n_window.geometry(f"{dims['ventana_front'][0]}x{dims['ventana_front'][1]}")

        return wrapper


    @fin
    def login_RED(self):
        """valida autenticación con usuario de red"""
        user = self.username_entry.get()
        pwd = self.password_entry.get()
        ldap_user = user + '@bancolombia.corp'
        ldap_server = 'Ldap.bancolombia.corp'
        server = ldap.Server(ldap_server, get_info = ldap.ALL)
        conn = ldap.Connection(server, user = ldap_user, password = pwd, auto_bind = True)

        return user, pwd
        

    @fin
    def login_NACIONAL(self):
        """Valida autenticación con usuario de NACIONAL"""
        user = self.username_entry.get()
        pwd = self.password_entry.get()
        self.connect_AS400(user, pwd)

        return user, pwd
        

    @fin
    def login_MEDELLIN(self):
        """Valida autenticación con usuario de MEDELLIN"""
        user = self.username_entry.get()
        pwd = self.password_entry.get()
        self.connect_AS400(user, pwd, 'MEDELLIN')

        return user, pwd
        
    
    def authenticate(self, applications, fr, window, **kwargs):
        """Se autentica en las aplicaciones definidas"""

        # Definir diccionario de funciones
        funcs = {
            'RED': [self.login_RED, os.getlogin()],
            'NACIONAL': [self.login_NACIONAL, ''],
            'MEDELLIN': [self.login_MEDELLIN, '']
        }

        # Definir el dataframe con las opciones
        self.n_window = window
        self.n_frame = fr

        # Definir función de formulario de autenticación
        def aut(application):

            window.geometry(f"{dims['ventana_autenticación'][0]}x{dims['ventana_autenticación'][1]}")
            
            # Crear frame principal
            self.principal_frame = Frame(bg = white)
            func = funcs[application][0]

            # Frame título
            frame1 = Frame(self.principal_frame, bg = 'white')
            login_logo = kwargs['login_logo']
            label = Label(frame1, image = login_logo, background = 'white')
            label.grid(row = 0, column = 0)

            if application == 'RED':
                text = 'Ingresa tu usuario y contraseña\nde red'
            else:
                text = f'Ingresa tu usuario y contraseña\nde AS400 {application}'
                
            title = Label(frame1, text = text, font = (font, 15), fg = black, bg = white)
            title.grid(row = 0, column = 1, pady = 10, padx=0)

            # Frame credenciales
            frame2 = Frame(self.principal_frame, bg = 'white', highlightthickness = 0)
            data_entry_box = kwargs['data_entry_box']
            username_label = Label(frame2, text = 'Usuario:', bg = white, font = (font, 14))
            password_label = Label(frame2, text = 'Contraseña:', bg = white, font = (font, 14))
            username_entry_i = Label(frame2, image = data_entry_box, borderwidth = 0, bg = white)
            self.password_entry_i = Label(frame2, image = data_entry_box, borderwidth = 0, bg = white)
            self.username_entry = Entry(frame2, font = (font, 14),  highlightthickness = 0, borderwidth = 0, bg = grey, fg = black)
            self.password_entry = Entry(frame2, font = (font, 14), show = '*', highlightthickness = 0, borderwidth = 0, bg = grey, fg = black)
            
            username_label.grid(row = 0, column = 0, sticky = 'esn')
            password_label.grid(row = 1, column = 0)
            username_entry_i.grid(row = 0, column = 1, padx = 10, pady = 10)
            self.username_entry.grid(row = 0, column = 1)
            self.password_entry_i.grid(row = 1, column = 1)
            self.password_entry.grid(row = 1, column = 1, padx = 10)

            # Frame botones
            frame3 = Frame(self.principal_frame, bg = 'white')
            aut_button = kwargs['aut_button']
            login_button_i = Label(frame3, image = aut_button, borderwidth = 0, bg = white)
            cancel_button_i = Label(frame3, image = aut_button, borderwidth = 0, bg = white)
            login_button = Button(frame3, text = 'Procesar', font = (font_bold, 12), bg = yellow, fg = black, borderwidth = 0, activebackground = yellow, activeforeground = white, command = func)
            cancel_button = Button(frame3, text = 'Cancelar', font = (font_bold, 12), bg = yellow, fg = black, borderwidth = 0, activebackground = yellow, activeforeground = white, command = lambda: sys.exit())
            
            cancel_button_i.grid(row = 0, column = 0, padx = 10)
            cancel_button.grid(row = 0, column = 0)
            login_button_i.grid(row = 0, column = 1)
            login_button.grid(row = 0, column = 1)
            
            # Texto por defecto
            self.username_entry.insert(0, funcs[application][1])

            # Ubicar widgets
            frame1.pack(pady =15)
            frame2.pack(pady = 0)
            frame3.pack(pady = 20)
            self.principal_frame.place(relx=0.5, rely=0.5, anchor=CENTER)

            # Ejecutar la función cuando se da enter y poner el mouse en la primer entrada
            window.bind("<Return>", lambda event: func())
            self.username_entry.focus_set()

        # Crear generador para las aplicaciones
        def app_gen():
            for application in applications:
                yield aut(application)

        global app
        app = app_gen()

        next(app)

        return self.user_pwd
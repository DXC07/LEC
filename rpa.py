from functions import Functions
from datetime import datetime
import os
import tkinter as tk
from PIL import Image
import json
from authentication import Authenticator
import sys
from time import sleep
from tabulate import tabulate
from front import Front
from tkcalendar import DateEntry
from consolidation import Consolidation
from conciliacion import Conciliacion
from lectura import Lectura
from embargos import Embargos
from classification import Classification
from comprobantes import Comprobantes
from report import Report
from getpass import getpass
from cryptography.fernet import Fernet

# Cargar imágenes y hacer resize
def load_image(path, resize = None):
    image = Image.open(path)

    if resize != None:
        new_image = image.resize(resize)
        new_image.save(path)

    image = tk.PhotoImage(file = path)

    return image

def tkinter_button(frame, image, row, column, padx = 0, pady = 0, **kwargs):
    button = tk.Button(frame, image = image, borderwidth = 0, bg = 'white')
    button.grid(row = row, column = column, padx = padx, pady = pady)
    button = tk.Button(frame, **kwargs)
    button.grid(row = row, column = column, padx = padx, pady = pady)

class RPA(Functions):

    @Functions.log_errors
    def __init__(self):
        super().__init__()
        self.start_date = None
        self.end_date = self._get_proc_data()
        self.credentials = None

        if not os.path.exists('./temp'):
            os.makedirs('./temp')
        
        self.app_authentication()
    
    def app_authentication(self):
        """Valida que la contraseña para ingresar a la aplicación sea correcta"""
        log_pwd = getpass("Ingrese contraseña EUC: ")

        try:
            pwd = "gAAAAABlpvQcKRUO0rjRrrfYQAwUb-AuniAdsLwhdSr9fo9czsUNImDR7RtXxe0nxjSZP2NuwrgkfcwOjUsAPai4Sq2gcAT4BQ=="
            key = open("mykey.key", mode = 'r', encoding = 'utf8').read()
            f = Fernet(key)
            pwd = f.decrypt(pwd).decode('utf-8')
        except:
            raise Exception("No se encontró archivo mykey.key o este no corresponde")
        
        if log_pwd == pwd:
            self.logger.info("Ingreso aplicación exitoso")
        else:
            raise Exception("Autenticación fallida: contraseña incorrecta")
        
        return None
    
    def _print_sum(self, _list):
        """Imprime un resumen con el las estadísticas de las tareas en formato de tabla"""
        # Crear lista vacía para guardar las filas
        table_rows = []

        # Process log data and add to table_rows
        for entry in _list:
            table_rows.append([entry['Tarea'], entry['Tiempo'], entry['Registros']])

        # Imprimir tabla
        table_headers = ["Tarea", "Tiempo", "Registros"]
        table = tabulate(table_rows, headers = table_headers, colalign = ['center'] * len(table_headers), tablefmt = "grid", stralign = 'left')
        self.logger.info('\n' * 2 + table + '\n')
    
    def _get_proc_data(self):
        """Obtiene datos que se recopilan de la ejecución anterior en caso de que la aplicación se cierre"""
        # Abrir archivo procesamiento
        with open('./procesamiento.json', mode = 'r', encoding = 'utf8') as jsonfile:
            data = json.load(jsonfile)

        try:
            date = data['fecha_ciclo']

            # Validar si corresponde al procesamiento de hoy
            if datetime.strptime(date, '%Y%m%d').date() < datetime.now().date():
                raise KeyError
            else:
                try:
                    end_date = datetime.strptime(data['fecha_final_rechazos'], '%Y%m%d').date()

                    return end_date
                except TypeError:
                    raise KeyError
        except KeyError:
            # Actualizar archivo procesamiento
            self.update_json(fecha_ciclo = datetime.now().strftime('%Y%m%d'), fecha_final_rechazos = None, archivo_embargos_radicado = None, radicado_lectura = None, ruta_consolidado_rechazos = None)
            
            return None

    def run_front(self, *funcs):
        """Despliega el front definido para la aplicación"""

        def on_closing():
            window.destroy()
            sys.exit()

        dims = {
            'ventana': (810, round(810 / (4398 / 2475))),
            'botones_ejecutar': (140, round(130 / (1.74 / 1.15))),
            'logo_inicio_sesion': (70, 70),
            'cuadro_texto_aut': (210, round(210 / (7.25 / 0.93))),
            'botones_aut': (95, round(95 / (4.1 / 1.35)))
        }

        front = Front()

        window = tk.Tk()
        window.protocol("WM_DELETE_WINDOW", on_closing)
        window.resizable(0,0)
        window.config(bg = 'white')
        window.title(self.settings['nombre_proceso'])
        p1 = tk.PhotoImage(file = './static/Imágenes/logo.png')
        window.iconphoto(False, p1)
        self.window = window

        # Cargar imágenes
        bg = load_image('./static/Imágenes/fondo.png', dims['ventana'])
        run_button = load_image('./static/Imágenes/ejecutar.png', dims['botones_ejecutar'])
        login_logo = load_image('./static/Imágenes/logo_inicio_sesion.png', dims['logo_inicio_sesion'])
        data_entry_box = load_image('./static/Imágenes/entrada_texto.png', dims['cuadro_texto_aut'])
        aut_button = load_image('./static/Imágenes/iniciar_sesion.png', dims['botones_aut'])

        # Cargar frame del front
        front_frame = front.run(window, *funcs, bg = bg, run_button = run_button)
        
        # Llamar módulo autenticación
        self.credentials = Authenticator().authenticate(['RED', 'NACIONAL'], front_frame, window, login_logo = login_logo, data_entry_box = data_entry_box, aut_button = aut_button)

        window.mainloop()

        return None

    def get_dates(self):
        """Muestra ventana para escoger fecha inicial y fecha final del proceso"""
        # Crear función para seleccionar funciones
        def get_date():
            # Validar que no sean fechas futuras o que la inicial sea mayor a la final
            if cal1.get_date() > cal2.get_date():
                self.show_messages('La fecha inicial es mayor a la fecha final', 'error')
            elif (cal1.get_date() >= datetime.now().date()) or (cal2.get_date() >= datetime.now().date()):
                self.show_messages('Fechas no válidas', 'error')
            else:
                # Obtener fechas
                self.start_date = cal1.get_date()
                self.end_date = cal2.get_date()
                # Cerrar ventana
                root.destroy()

                # Ejecutar función
                self.consolidar()

        # Crear ventana
        root = tk.Toplevel(self.window)
        root.configure(bg="white")
        root.title("")
        p1 = tk.PhotoImage(file = './static/Imágenes/logo.png')
        root.iconphoto(False, p1)

        # Definir mensaje título entradas calendario
        tk.Label(root, text='Por favor escoja las fechas', bg = 'white', font = ('CIBFont Sans', 12)).pack(padx=10, pady=10)
        cal1 = DateEntry(root, width=12, background='black',
                        foreground='white', borderwidth=2, locale = 'es_CO', date_pattern='dd/mm/yyyy', font = ('CIBFont Sans', 12))
        
        cal2 = DateEntry(root, width=12, background='black',
                        foreground='white', borderwidth=2, locale = 'es_CO', date_pattern='dd/mm/yyyy', font = ('CIBFont Sans', 12))
        cal1.delete(0, "end")
        cal2.delete(0, "end")

        # Desplegar entradas calendario y botón de aceptar
        cal1.pack(padx=10, pady=10)
        cal2.pack(padx=10, pady=10)
        accept_button = tk.Button(root, text = 'Aceptar', command = get_date, bg = '#2C2A29', fg = 'white', font = ('CIBFont Sans', 12))
        accept_button.pack(padx=10, pady=10)

        root.mainloop()

    @Functions.log_errors
    @Functions.save_execution_time
    def consolidar(self):
        print('Ejecutando botón 1')
        self.val_permission()
        consolidation = Consolidation(self.credentials, self.start_date, self.end_date)
        consolidation.run()

    @Functions.log_errors
    @Functions.save_execution_time
    def conciliacion(self):
        print('Ejecutando botón 2')
        conciliacion = Conciliacion(self.credentials, self.end_date)
        conciliacion.run()
    
    @Functions.log_errors
    @Functions.save_execution_time
    def lectura(self):
        print('Ejecutando botón 3')
        lectura = Lectura(self.credentials, self.end_date)
        lectura.run()

    @Functions.log_errors
    @Functions.save_execution_time
    def embargos(self):
        print('Ejecutando botón 4')
        embargos = Embargos(self.credentials ,self.end_date, self.window)
        embargos.run()
    
    @Functions.log_errors
    @Functions.save_execution_time
    def clasificacion(self):
        print('Ejecutando botón 5')
        classification = Classification(self.credentials, self.end_date)
        classification.run()
    
    @Functions.log_errors
    @Functions.save_execution_time
    def comprobantes(self):
        print('Ejecutando botón 6')
        comprobantes = Comprobantes(self.end_date)
        comprobantes.run()
    
    @Functions.log_errors
    @Functions.save_execution_time
    def informe(self):
        print('Ejecutando botón 7')
        report = Report(self.end_date)
        report.run()

    def run_RPA(self):
        """Orquesta la ejecución del RPA"""
        self.logger.info('Iniciando RPA')

        self.run_front(self.get_dates, self.conciliacion, self.lectura, self.embargos, self.clasificacion, self.comprobantes, self.informe)




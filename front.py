import tkinter as tk
import json
from PIL import Image
from functions import Functions

# Dimensiones widgets
dims = {
    'ventana': (810, round(810 / (4398 / 2475)))
}

# Colores y tipo de funetes
black = '#2C2A29'
white = 'white'
font_bold = 'CIBFont Sans Bold'
font = 'CIBFont Sans'

def tkinter_button(frame, image, row, column, padx = 0, pady = 0, **kwargs):
    button = tk.Button(frame, image = image, borderwidth = 0, bg = 'white')
    button.grid(row = row, column = column, padx = padx, pady = pady)
    button = tk.Button(frame, **kwargs)
    button.grid(row = row, column = column, padx = padx, pady = pady)

class Front(Functions):

    def run(self, window ,*funcs, **kwargs):
        # Configuración tamaño
        window.geometry(f"{dims['ventana'][0]}x{dims['ventana'][1]}")

        # Crear frame principal
        principal_frame = tk.Frame(window, bg = 'white', width = dims['ventana'][0], height = dims['ventana'][1])

        # Fondo de pantalla
        bg = kwargs['bg']
        tk.Label(principal_frame, image = bg).place(x = 0, y = 0)

        # Cargar imágenes de los botones
        run_button = kwargs['run_button']

        # Crear frame con los widgets
        frame = tk.Frame(principal_frame, bg = white)
        frame1 = tk.Frame(frame, bg = white)
        frame2 = tk.Frame(frame, bg = white)

        # Título
        principal_frame.grid_columnconfigure(0, weight=1)
        title = tk.Label(principal_frame, text = self.settings['nombre_proceso'], font = (font, 28), fg = 'black', bg = white)
        title.grid(row = 0, column = 0, pady = 7)
        subtitle = tk.Label(principal_frame, text = 'Sección Servicios Financieros', font = (font, 18), fg = 'black', bg = white)
        subtitle.grid(row = 1, column = 0, pady = 0)

        # Texto botones
        texts = self.settings['botones']

        # Desplegar botones primera fila
        tkinter_button(frame1, run_button, 1, 1, padx = 5, pady = 5, text = texts[0], fg = white, bg = black, borderwidth = 0, activebackground = black, activeforeground = white, command = funcs[0], font = ('CIBFont Sans', 12), height = 3, width = 14)
        tkinter_button(frame1, run_button, 1, 2, padx = 5, pady = 5, text = texts[1], fg = white, bg = black, borderwidth = 0, activebackground = black, activeforeground = white, command = funcs[1], font = ('CIBFont Sans', 12), height = 3, width = 14)
        tkinter_button(frame1, run_button, 1, 3, padx = 5, pady = 5, text = texts[2], fg = white, bg = black, borderwidth = 0, activebackground = black, activeforeground = white, command = funcs[2], font = ('CIBFont Sans', 12), height = 3, width = 14)
        tkinter_button(frame1, run_button, 1, 4, padx = 5, pady = 5, text = texts[3], fg = white, bg = black, borderwidth = 0, activebackground = black, activeforeground = white, command = funcs[3], font = ('CIBFont Sans', 12), height = 3, width = 14)

        # Desplegar botones primera fila
        tkinter_button(frame2, run_button, 1, 1, padx = 5, pady = 5, text = texts[4], fg = white, bg = black, borderwidth = 0, activebackground = black, activeforeground = white, command = funcs[4], font = ('CIBFont Sans', 12), height = 3, width = 14)
        tkinter_button(frame2, run_button, 1, 2, padx = 5, pady = 5, text = texts[5], fg = white, bg = black, borderwidth = 0, activebackground = black, activeforeground = white, command = funcs[5], font = ('CIBFont Sans', 12), height = 3, width = 14)
        tkinter_button(frame2, run_button, 1, 3, padx = 5, pady = 5, text = texts[6], fg = white, bg = black, borderwidth = 0, activebackground = black, activeforeground = white, command = funcs[6], font = ('CIBFont Sans', 12), height = 3, width = 14)

        # Desplegar frame
        frame.grid(row = 2, column = 0, pady= 30)
        frame1.grid(row = 1, column = 0, pady= 0)
        frame2.grid(row = 2, column = 0, pady= 0)

        return principal_frame


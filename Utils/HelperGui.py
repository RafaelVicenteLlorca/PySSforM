import tkinter as tk
import os


# Crea una ventana y la devuelve para su uso
def crearVentana(titulo):
    ventana = tk.Tk()
    ventana.geometry("1000x700")  # tamaño e la ventana
    # ventana.resizable(False, False) #que no se pueda modificar el tamaño de la ventana
    ventana.minsize("800", "500")  # tamaño minimo de la ventana (anchura, altura)
    # ventana.maxsize("700", "800")   tamaño maximo de la ventana (anchura, altura)
    ventana.state('zoomed')
    ventana.attributes('-alpha', 1)  # trasparencia (mas cerca 0 = mas transparente)
    ventana.configure(background="grey")
    icono = tk.PhotoImage(file=os.path.join(os.path.dirname(__file__), "../Imagenes/PySSforM-logo3.png"))
    ventana.iconphoto(True, icono)  # icono de la ventana
    ventana.title(titulo)  # titulo de la ventana
    return ventana


def cambioHexARGB(color):
    rgb = tuple(int(color[i:i + 2], 16) for i in (2, 4, 6))
    colorRGB = '#%02x%02x%02x' % rgb
    return colorRGB


def separarRutaArchivo(ruta_archivo):
    ruta, archivo = os.path.split(ruta_archivo)
    return ruta, archivo


def tieneExtension(nombre_archivo):
    _, extension = os.path.splitext(nombre_archivo)
    return bool(extension)

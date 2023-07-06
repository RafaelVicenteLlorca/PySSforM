
import os
from BackEnd import Back
from Utils import HelperGui
import tkinter as tk
from tkinter import filedialog

colorFondo = "grey"

class GuiClass:

    def __init__(self):
        self.lista_textbox = []
        self.lista_labels = []
        self.lista_indices = []

        self.ventana = None
        self.botonAnterior = None
        self.botonSiguiente = None

        self.posicion = 1
        self.ordenado = 0
        self.orden = True
        self.max_filas = 0
        self.back = Back.BackClass()

    def reinicioVariables(self):

        self.lista_textbox = []
        self.lista_labels = []
        self.lista_indices = []

        self.ventana.destroy()
        self.ventana = None
        self.botonAnterior = None
        self.botonSiguiente = None

        self.posicion = 1
        self.ordenado = 0
        self.orden = True
        self.max_filas = 0
        self.back = Back.BackClass()



    """
            FUNCIONES VENTANA
    """


    def inicioVentana(self):
        titulo = "PySSforM"
        self.ventana = HelperGui.crearVentana(titulo)

        # Frame principal
        frame_principal = tk.Frame(self.ventana,background= colorFondo)
        frame_principal.pack(expand=True)

        imagen_logo = tk.PhotoImage(file=os.path.join(os.path.dirname(__file__), "../Imagenes/PySSforM-logo3.png"))
        frame_logo = tk.Frame(frame_principal, background=colorFondo)
        etiqueta_logo = tk.Label(frame_logo, image=imagen_logo, background=colorFondo)
        etiqueta_logo.pack(pady=10)
        frame_logo.pack()
        # Crear contenedor Frame
        frame = tk.Frame(frame_principal, background=colorFondo)

        # Podemos pedir el path por teclado

        # Etiqueta del path
        etiqueta = tk.Label(frame, text="Inserte el path al archivo", background=colorFondo, font=("Arial", 12, "bold"))
        etiqueta.pack(side=tk.LEFT)
        # TextBox del path
        ruta = tk.Entry(frame)
        ruta.pack(side=tk.LEFT)

        # Crear contenedor Frame

        # Etiqueta del path
        etiqueta2 = tk.Label(frame, text="Inserte el archivo", background=colorFondo, font=("Arial", 12, "bold"))
        etiqueta2.pack(side=tk.LEFT)
        # TextBox del archivo
        archivo = tk.Entry(frame)

        archivo.pack(side=tk.LEFT)

        # Crear contenedor Frame
        frame2 = tk.Frame(frame_principal,background=colorFondo)
        # Etiqueta del path
        etiqueta3 = tk.Label(frame2, text="Inserte la hoja (si no inserta nada por defecto sera Hoja1", background=colorFondo, font=("Arial", 12, "bold"))
        etiqueta3.pack(side=tk.LEFT)
        # TextBox del path
        hoja = tk.Entry(frame2)
        hoja.pack(side=tk.LEFT)

        frame.pack()

        frame2.pack()

        # Boton para usar los textbox y buscar el archivo
        frame_Botones = tk.Frame(self.ventana, background=colorFondo)
        frame_Botones.pack(side="bottom", anchor="center")
        botonIniciador = tk.Button(frame_Botones, text="Cambiar ventana",
                          command=lambda: self.iniciarExcelManejador(ruta.get(), archivo.get(), hoja.get()))

        # Tambien se puede usar el buscador
        # Boton para buscar el archivo mediante el explorador de archivos
        botonBuscador = tk.Button(frame_Botones, text="Buscar archivo",
                          command=lambda: self.buscadorArchivo(ruta.get(), hoja.get()))

        botonIniciador.grid(row=0, column=0, padx=10, pady=5)
        botonBuscador.grid(row=0, column=1, padx=10, pady=5)

        self.ventana.mainloop()

    def buscadorArchivo(self, ruta, hoja):
        root = tk.Tk()
        root.withdraw()
        # Obtenemos el path
        ruta_archivo = filedialog.askopenfilename(initialdir=ruta)

        ruta, archivo = HelperGui.separarRutaArchivo(ruta_archivo)
        self.iniciarExcelManejador(ruta, archivo, hoja)

        # manejador del boton de inicio

    def iniciarExcelManejador(self, ruta, archivo, hoja):

        nombrearchivo = str(archivo)
        if len(self.lista_indices) == 0:

            ruta_final = str(ruta) + "\\" + nombrearchivo
            if not (HelperGui.tieneExtension(nombrearchivo)):
                ruta_final += ".xlsx"
            self.back.abrirExcel(ruta_final, hoja)
            self.back.crearIndice()
            self.lista_indices = self.back.getIndice()
            self.ordenarIndice()
            self.max_filas = self.back.maximoFilas()

        if self.ventana is not None:
            self.ventana.destroy()

        self.ventana = HelperGui.crearVentana(nombrearchivo)


        # Boton para volver al principio
        salirboton = tk.Button(self.ventana, text="Salir", command=lambda: self.salirBoton())
        salirboton.pack(side="top", anchor="ne")

        self.formatearVentana()
        # Contenedor para los botones de la izquierda
        frame_izquierda = tk.Frame(self.ventana, background=colorFondo)
        frame_izquierda.pack(side="left", anchor="sw")

        nuevoboton = tk.Button(frame_izquierda, text="Nuevo", command=lambda: self.nuevoBoton())
        borrarboton = tk.Button(frame_izquierda, text="Borrar", command=lambda: self.borrarBoton())

        nuevoboton.pack(side="left", padx=5, pady=5)
        borrarboton.pack(side="left", padx=5, pady=5)


        # Contenedor para los botones de la derecha
        frame_derecha = tk.Frame(self.ventana, background=colorFondo)
        frame_derecha.pack(side="right", anchor="se")

        modificarboton = tk.Button(frame_derecha, text="Modificar", command=lambda: self.modificarBoton())
        vaciarboton = tk.Button(frame_derecha, text="Vaciar", command=lambda: self.vaciarBoton())
        buscarboton = tk.Button(frame_derecha, text="Buscar", command=lambda: self.buscarBoton())

        modificarboton.pack(side="right", padx=5, pady=5)
        vaciarboton.pack(side="right", padx=5, pady=5)
        buscarboton.pack(side="right", padx=5, pady=5)


        # Contenedor para los botones centrales
        frame_centro = tk.Frame(self.ventana, background=colorFondo)
        frame_centro.pack(side="bottom", anchor="center")

        self.botonSiguiente = tk.Button(frame_centro, text="Siguiente", command=lambda: self.siguienteBoton())
        self.botonAnterior = tk.Button(frame_centro, text="Anterior", command=lambda: self.anteriorBoton())
        self.botonAnterior.configure(state="disabled")

        self.botonAnterior.grid(row=0, column=0, padx=10, pady=5)
        self.botonSiguiente.grid(row=0, column=1, padx=10, pady=5)




        self.ventana.mainloop()

    def ordenarIndice(self):

        lista= []
        for celda in self.lista_indices:
            inserted= False
            encontrado = False
            for posicion in range(len(lista)):
                if celda.color == lista[posicion].color:
                    encontrado = True
                elif encontrado:
                    lista.insert(posicion, celda)
                    inserted = True
                    break
            if not inserted:
                lista.append(celda)
        self.lista_indices = lista

    def formatearVentana(self):
        color2 = None
        tam = 0
        fila = 0
        columna = 0
        for indice in self.lista_indices:
            # Obtenemos el color de una seccion
            color1 = indice.color
            if color1 != color2 :
                # Crear contenedor Frame
                frame = tk.Frame(self.ventana,width=200, height=100, relief=tk.RAISED, borderwidth=1)
                frame.configure(bg=indice.color)
                frame.pack(side="top", anchor="w", fill="both", expand=True)

                frame.propagate(False)
                color2 = color1
                tam = 0
                fila = 0
                columna = 0
            elif color1 == color2 and (tam + indice.tam) > 1:
                fila = fila + 2
                columna = 0

            tam = self.printBox(frame, indice, tam, fila, columna)
            columna = columna + 2
    def printBox(self, frame, indice, tam, fila, columna):

        # Crear etiqueta
        bgcolor = indice.color
        etiqueta = tk.Label(frame, text=indice.titulo, bd=0, relief=tk.FLAT, bg=bgcolor, width=20, font=("Arial", 12, "bold"))

        if not (indice.esBooleano()) and not (indice.esFuncion()):
            etiqueta.configure(cursor="hand2")
            etiqueta.bind("<Button-1>", lambda event: self.ordenar(indice.posicionHorizontal))

        etiqueta.grid(row=fila, column=columna)
        columna = columna +1
        # Crear CheckBox
        if indice.esBooleano():
            checkvar = tk.StringVar()
            textbox = tk.Checkbutton(frame, variable=checkvar)
            textbox.deselect()
            textbox.configure(bg=bgcolor)
            textbox.grid(row=fila, column=columna)
            tam = tam + indice.tam
            self.lista_textbox.append((indice.posicionHorizontal, textbox, checkvar))
        elif indice.esString():
            textbox = tk.Text(frame)
            if indice.tipo == "string":
                textbox.config(width=40, height=1)
                tam = tam + indice.tam
            elif indice.tipo == "long string":
                textbox.config(width=40, height=5)
                tam = tam + indice.tam
            else:
                textbox = tk.Entry(frame, validate="key")
                tam = tam + indice.tam

            textbox.grid(row=fila, column=columna)
            self.lista_textbox.append((indice.posicionHorizontal, textbox))
        else:
            textbox = tk.Entry(frame, validate="key")
            textbox.configure(validate="key")
            # Crear caja de texto para SOLO numeros
            if indice.esInteger():
                textbox.configure(validatecommand=(textbox.register(lambda text: text.isdigit()), "%P"))
                tam = tam + indice.tam
            # Crear caja de texto para SOLO lectura
            if indice.esFuncion():
                textbox.configure(state="readonly")
                tam = tam + indice.tam



            textbox.grid(row=fila, column=columna)
            self.lista_textbox.append((indice.posicionHorizontal, textbox))
        espacio = tk.Frame(frame, width=30, height=1,bg=indice.color)
        espacio.pack(side="left")
        frame.grid_rowconfigure(fila, weight=1, minsize=50)
        return tam



    """
            BOTONES CRUD
    """

    def nuevoBoton(self):

        lista = []

        contador = 0
        for textbox in self.lista_textbox:
            vacio = False
            if isinstance(textbox[1], tk.Checkbutton):
                if textbox[2].get() == "1":
                    valortexbox = "true"
                else:
                    valortexbox = "false"
            else:
                if isinstance(textbox[1], tk.Text):
                    valortexbox = textbox[1].get("1.0", tk.END).rstrip("\n")
                else:
                    valortexbox = textbox[1].get()
            if valortexbox is None or valortexbox == "":
                vacio = True
                contador += 1

            if not vacio:
                lista.append((textbox[0], valortexbox))
        if contador != len(self.lista_textbox):
            self.back.anadirFila(lista)
            self.vaciarTextbox()
            self.reinicioBotones()
            self.max_filas = self.back.maximoFilas()

    def buscarBoton(self):
        linea = None
        posicion = 1
        for textbox in self.lista_textbox:
            if not isinstance(textbox[1], tk.Checkbutton):
                if isinstance(textbox[1], tk.Text):
                    valortexbox = textbox[1].get("1.0", tk.END).rstrip("\n")
                else:
                    valortexbox = textbox[1].get()
                if valortexbox is not None and valortexbox != "":
                    posicion, linea = self.back.obtenerLinea(textbox[0], valortexbox)
                    break
        if linea is not None:
            self.vaciarTextbox()
            self.rellenarTextbox(linea)
            self.reinicioBotones()
            self.posicion = posicion
            self.comprobarBotones()

    def modificarBoton(self):
        if self.posicion != 1:
            lista = []

            contador = 0
            for textbox in self.lista_textbox:

                if isinstance(textbox[1], tk.Checkbutton):
                    if textbox[2].get() == "1":
                        valortexbox = "true"
                    else:
                        valortexbox = "false"
                else:
                    if isinstance(textbox[1], tk.Text):
                        valortexbox = textbox[1].get("1.0", tk.END).rstrip("\n")
                    else:
                        valortexbox = textbox[1].get()
                if valortexbox is None or valortexbox == "":
                    contador += 1

                lista.append((textbox[0], valortexbox))
            if contador != len(self.lista_textbox):
                self.back.modificarFila(self.posicion, lista)
                self.vaciarTextbox()
                self.reinicioBotones()
                self.max_filas = self.back.maximoFilas()

    def borrarBoton(self):
        lista = []

        for textbox in self.lista_textbox:
            if isinstance(textbox[1], tk.Checkbutton):
                if textbox[2].get() == "1":
                    valortexbox = "true"
                else:
                    valortexbox = "false"
            else:
                if isinstance(textbox[1], tk.Text):
                    valortexbox = textbox[1].get("1.0", tk.END).rstrip("\n")
                else:
                    valortexbox = textbox[1].get()
            if valortexbox is not None and valortexbox != "":
                lista.append((textbox[0], valortexbox))

        if len(lista) > 0:
            self.back.borrarFila(lista)
            self.vaciarTextbox()
            self.reinicioBotones()
            self.max_filas = self.back.maximoFilas()

    """
            BOTONES EXTRAS CENTRALES
    """

    def siguienteBoton(self):
        lista = self.back.siguienteLinea(self.posicion)
        if lista is not None:
            self.vaciarTextbox()
            self.rellenarTextbox(lista)
            self.posicion += 1
        self.comprobarBotones()

    def anteriorBoton(self):
        lista = self.back.anteriorLinea(self.posicion)
        if lista is not None:
            self.vaciarTextbox()
            self.rellenarTextbox(lista)
            self.posicion -= 1
        self.comprobarBotones()

    """
            BOTONES EXTRAS
    """

    def vaciarBoton(self):
        self.reinicioBotones()
        self.vaciarTextbox()

    def vaciarTextbox(self):
        for textbox in self.lista_textbox:
            if isinstance(textbox[1], tk.Entry):
                if textbox[1].get().isdigit():
                    textbox[1].configure(validate="none")
                    textbox[1].delete(0, tk.END)
                    textbox[1].configure(validate="key")
                elif textbox[1].cget("state") == "readonly":
                    textbox[1].configure(state="normal")
                    textbox[1].delete(0, tk.END)
                    textbox[1].configure(state="readonly")
                else:
                    textbox[1].delete(0, tk.END)
            if isinstance(textbox[1], tk.Checkbutton):
                textbox[1].deselect()
            if isinstance(textbox[1], tk.Text):
                textbox[1].delete("1.0", tk.END)

    def ordenar(self, posicion):
        if self.ordenado == posicion:
            self.orden = not self.orden
        else:
            self.orden = False
            self.ordenado = posicion
        self.back.ordenarFilas(posicion - 1, self.orden)
        self.reinicioBotones()
        self.vaciarTextbox()

    def salirBoton(self):

        self.back.reinicio()
        self.reinicioVariables()
        self.inicioVentana()


    """
            FUNCIONES DE AYUDA BOTONES   
    """

    def reinicioBotones(self):
        self.posicion = 1
        self.botonSiguiente.configure(state='normal')
        self.botonAnterior.configure(state='disabled')

    def rellenarTextbox(self, linea):
        for valor in linea:
            for textbox in self.lista_textbox:
                if valor.valor != None and valor.valor != "":
                    if textbox[0] == valor.posicionHorizontal:
                        if not isinstance(textbox[1], tk.Checkbutton):
                            if isinstance(textbox[1],  tk.Text):
                                textbox[1].insert(tk.END, valor.valor)
                            else:
                                if textbox[1].cget("state") == "readonly":
                                    textbox[1].configure(state="normal")
                                    textbox[1].insert('0', valor.valor)
                                    textbox[1].configure(state="readonly")
                                else:
                                    textbox[1].insert('0', valor.valor)
                        elif valor.valor == "true":
                            textbox[1].select()
                        else:
                            textbox[1].deselect()
                        break

    def comprobarBotones(self):

        if self.posicion <= 2:
            self.botonAnterior.configure(state='disabled')
        else:
            self.botonAnterior.configure(state='normal')

        if self.posicion >= self.max_filas:
            self.botonSiguiente.configure(state='disabled')
        else:
            self.botonSiguiente.configure(state='normal')

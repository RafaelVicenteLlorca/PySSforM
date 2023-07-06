
from Clases import Cells


class ExcelClass:

    def __init__(self, libro=None, hoja=None, file_path=None):

        if libro is not None and hoja is not None and file_path is not None:
            self.libro = libro
            self.hoja = hoja
            self.ruta = file_path

        else:
            self.libro = None
            self.hoja = None
            self.ruta = None

        self.indice = []

    def reinicio(self):
        self.libro = None
        self.hoja = None
        self.ruta = None
        self.indice = []

    """
            FUNCIONES GET
    """
    def maximoFilas(self):
        return self.hoja.max_row

    def maximoColumnas(self):
        return self.hoja.max_column

    def getHoja(self):
        return self.hoja

    def getIndice(self):
        return self.indice

    def getLibro(self):
        return self.libro

    def getValorIndice(self, pos):
        for objeto in self.indice:
            if objeto.posicionHorizontal == pos:
                return objeto
        return None




    """
            FUNCIONES BUSCAR
    """
    # Funcion que recibe una posicion vertical y devuelve la linea asociada a esa posicion
    def buscarColumna(self, pos):
        columna = []

        # Especificar el rango de celdas que contiene la línea
        for posH in range(self.maximoColumnas()):
            celdaExcel = self.hoja.cell(row=pos, column=posH+1)
            if Cells.esFuncion(celdaExcel):
                valor = celdaExcel.value
                celda = Cells.Celda(celdaExcel.column, valor)
            else:
                celda = Cells.Celda(celdaExcel.column, celdaExcel.value)


            columna.append(celda)

        return columna

    # Funcion que recibe un dato y una posicion horizontal y devuelve la posicion vertical en la que se encuentra
    def buscarDato(self, pos, dato):
        posDato = -1
        # Especificar el rango de celdas que contiene la línea
        for posV in range(self.maximoFilas()):
            celdaExcel = self.hoja.cell(row=posV+1, column=pos)
            if str(celdaExcel.value) == dato:
                posDato = posV+1
                break
        return posDato



    """
            FUNCIONES PARA MODIFICAR
    """

    # Funcion que recibe una posicion vertical y una linea y modifica los valores de esa linea asociada a la posicion
    def modificarColumna(self, linea, posy):
        # Recorremos la linea a introducir
        for celda in linea:
            dato = celda[1]
            posx = celda[0]
            # Comprobamos que no estamos intentando introducir un dato donde es una funcion
            if self.getValorIndice(posx).esInteger():
                self.hoja.cell(row=posy, column=posx).value = int(dato)
            elif not self.getValorIndice(posx).esFuncion():
                # Escribir el valor en la última fila de la columna
                self.hoja.cell(row=posy, column=posx).value = dato

        self.libro.save(self.ruta)

    # Funcion que recibe el string de una formula y las posiciones verticales que debe mover y la ajusta
    def cambiarFormula(self, formula, posiciones):
        resultado = ""
        numero_actual = ""
        caracter_anterior = ""
        for caracter in formula:
            if caracter.isdigit() and caracter_anterior != "$":
                numero_actual += caracter
            else:
                if numero_actual != "":
                    numero_incrementado = str(int(numero_actual) + posiciones)
                    resultado += numero_incrementado
                    numero_actual = ""
                    caracter_anterior = caracter
                resultado += caracter

        if numero_actual != "":
            numero_incrementado = str(int(numero_actual) + 1)
            resultado += numero_incrementado

        return resultado

    """
            FUNCIONES PARA INSERTAR
    """
    # Funcion que añade una Columna nueva al final del excel a traves de los datos asociados a cada celda
    def insertarColumna(self, linea):

        # Determinar la última fila en la columna
        nueva_ultima_fila = self.maximoFilas() + 1
        # Recorremos la linea a introducir
        for celda in linea:
            dato = celda[1]
            posicion = celda[0]
            # Comprobamos que no estamos intentando introducir un dato donde es una funcion
            if self.getValorIndice(posicion).esInteger():
                self.hoja.cell(row=nueva_ultima_fila, column=posicion).value = int(dato)
            else:
                # Escribir el valor en la última fila de la columna
                self.hoja.cell(row=nueva_ultima_fila, column=posicion).value = dato
        self.rellenarFormulas(1)
        self.libro.save(self.ruta)
    # Funcion que recibe una posicion y copia la formula anterior ajustandola
    def rellenarFormulas(self, posicion):
        for pos in range(self.maximoColumnas()):
            if self.getValorIndice(pos+1).esFuncion():
                ultima_fila = self.maximoFilas()
                # obtenemos la formula de un valor por encima modificandolo
                dato = self.cambiarFormula(self.hoja.cell(row=ultima_fila - 1, column=pos+1).value, posicion)
                # la introducimos en la nueva columna
                self.hoja.cell(row=ultima_fila, column=pos+1).value = dato

    #Funcion que inserta los datos con el nuevo orden
    def insertarOrdenados(self, datos_ordenados):
        for fila, valores in enumerate(datos_ordenados, start=2):
            for columna, valor in enumerate(valores, start=1):
                if not self.getValorIndice(columna).esFuncion():
                    self.hoja.cell(row=fila, column=columna, value=valor)

    # Funcion que introduce un nuevo indice en la lista si este no existe
    def setIntoIndice(self, pos, type, color, title, tam):
        nuevo = Cells.Indice(pos, type, color, title, tam)
        if len(self.indice) == 0 or nuevo not in self.indice:
            self.indice.append(nuevo)


    """
                FUNCIONES BORRAR
    """

    # Funcion que recibe una posicion vertical y elimina la linea asociada a esa posicion
    def eliminarColumna(self, pos):
        self.hoja.delete_rows(pos)
        self.libro.save(self.ruta)

    # Funcion que Borrar el contenido de las celdas vacías en la hoja
    def eliminarAntiguos(self, filas):

        for fila in range(2, filas + 1):
            for columna in range(1, self.maximoColumnas() + 1):
                celda = self.hoja.cell(row=fila, column=columna)
                if celda.value is None:
                    celda.value = ''
        self.libro.save(self.ruta)



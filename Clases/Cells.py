
class Celda:
    def __init__(self, pos, value):
        self.posicionHorizontal = pos
        self.valor = value

    def __eq__(self, other):
        return self.posicionHorizontal == other.posicionHorizontal

class Indice:
    def __init__(self, pos, type, color, title, tam):
        self.posicionHorizontal = pos
        self.tipo = type
        self.color = color
        self.titulo = title
        self.tam = tam

    def __eq__(self, other):
        return self.posicionHorizontal == other.posicionHorizontal

    def esBooleano(self):
        return self.tipo == "Booleano"
    def esFuncion(self):
        return self.tipo == "Funcion"

    def esInteger(self):
        return self.tipo == "Integer"

    def esIndeterminado(self):
        return self.tipo == "Indeterminado"

    def esString(self):
        return not ( self.esFuncion()  or self.esIndeterminado() or self.esInteger() or self.esBooleano())

    def tipoString(self):
        if self.tipo == "short string":
            return 0
        elif self.tipo == "string":
            return 1
        elif self.tipo == "long string":
            return 2
        else:
            return -1

def esBoleano(cell):
    if (cell.value == 0 or cell.value == 1 or cell.value == "true" or cell.value == "false"):
        return True
def esFuncion(cell):
    return cell.data_type == "f"

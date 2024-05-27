class Pila:
    def crearPila(self):
        self.items = []
    
    def pilaVacia(self):
        return len(self.items) == 0
    
    def ponerEnPila(self, item):
        return (self.items.append(item))
    
    def sacarDePila(self):
        if len(self.items) == 0:
            return 0
        return self.items.pop()
        
    def verTope(self):
        if len(self.items) == 0:
            return 0
        
        return self.items[-1]
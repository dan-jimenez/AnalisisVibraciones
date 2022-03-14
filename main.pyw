import pandas as pd  # Libreria para abrir archivos excel y importar los datos a la progra
from tkinter import *  # Libreria para crear ventanas
import matplotlib.pyplot as plt  # libreria para crear graficos
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk  # Libreria para importar graficos dentro de la interfaz


# Creacion de la ventana y el archivo como objeto
# ------------------------------------------------


class Graficos:
    # Clase funciona como un machote para crear todos los graficos necesarios dentro de la interfaz
    canvas = None
    figure = None
    data = None
    frame = None

    def __init__(self, root, sheetName, columns,
                 frame):  # Esta funcion se llama metodo constructor, sirve para crear graficos con esta plantilla
        self.data = pd.read_excel(root, sheet_name=sheetName,
                                  usecols=columns)  # Se lee el archivo usando solo ciertas columnas
        self.frame = self.data.plot(columns[0], columns[1])  # Se crea la ventana grande del grafico
        self.figure = plt.Figure(figsize=(6, 5),
                                 dpi=100)  # Se crea el grafico como una figura para insertarla dentro de la interfaz
        self.figure.add_subplot().plot(self.data["Date"], self.data["mm/s"])  # Se agregon los ejes al grafico
        self.canvas = FigureCanvasTkAgg(self.figure, master=frame)  # se coloca el grafico dentro de la interfaz
        self.canvas.draw()  # Se dibuja el grafico

    def open(self):
        # metodo que funciona para abrir un grafico en pantalla grande
        plt.show()

    def placeFigure(self, _x, _y):
        # metodo que funciona para colocar un grafico dentro de la interfaz en un espacio determinado
        self.canvas.get_tk_widget().place(x=_x, y=_y)


class BigFrame:
    # Declaracion de la clase que sirve para implementar la interfaz
    graficos = []  # Tenemos una lista de graficos generales
    leftFrame = None
    rightFrame = None
    bottomFrame = None

    def __init__(self, window):  # Metodo constructor que sirve para crear la ventana general de la interfaz
        window.geometry('1000x700')  # le damos el tamanio a la ventana
        window.title('Pronostico vibracion AV')  # Aqui le agregamos el titulo a la ventana
        window.configure(bg="white")  # le cambiamos el fondo a la ventana en color blanco

        # Para dividir la pantalla en tres partes no iguales, modo de visualizacion de la interfaz en formato Z para mejor enfoque en lo importante (graficos)

        self.leftFrame = Frame(window, background="#3C4FA8", width=690,
                               height=620)  # agregar una seccion dentro de la ventana
        self.leftFrame.place(relx=300, rely=50)  # insertar esa seccion dentro de la ventana (que se dibuje)
        self.leftFrame.grid(row=0, column=0, padx=5, pady=5)  # Colocamos en la posicion 0 la seccion

        self.rightFrame = Frame(window, background="#3C4FA8", width=290, height=690)
        self.rightFrame.place(relx=300, rely=50)
        self.rightFrame.grid(row=0, column=1, pady=5, padx=5, rowspan=2)

        self.bottomFrame = Frame(window, background="#3C4FA8", width=690, height=60)
        self.bottomFrame.place(relx=300, rely=50)
        self.bottomFrame.grid(row=1, column=0, pady=5, padx=5)

        self.graficos.append(Graficos("test.xlsx", "Vent Int A", ["Date", "mm/s"],
                                      self.leftFrame))  # Agregamos un grafico a la lista de graficos general
        self.graficos[0].placeFigure(50, 50)  # Dibujamos el grafico 0 dentro de la interfaz

        self.title = Label(self.leftFrame, text="Grafico de prueba", background="#3C4FA8", foreground="white",
                           font=("Times New Romans", "22", "bold"))  # agregar un texto en la seccion 1
        self.title.place(x=220, y=10)  # colocar el texto en la posicion deseada

        self.bOpen = Button(self.bottomFrame, text="Abrir", command=lambda: self.openGraphic(0))
        self.bOpen.place(x=50, y=10)

    def openGraphic(self, position):
        # Funcion para abrir un grafico en pantalla grande
        self.graficos[position].open()


if __name__ == "__main__":
    # Esta es la funcion principal que ejecuta todo el programa
    window = Tk() # crea la ventana
    bFrame = BigFrame(window) #ejecuta todos los comandos para la personalizacion de la interfaz

    window.mainloop() #se corre el programa

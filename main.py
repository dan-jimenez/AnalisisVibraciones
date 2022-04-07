import tkinter
import tkinter.ttk
from tkinter import ttk

import pandas as pd  # Libreria para abrir archivos excel y importar los datos a la progra
from tkinter import *  # Libreria para crear ventanas
import tkinter as tk
import matplotlib.pyplot as plt  # libreria para crear graficosTiempo
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, \
    NavigationToolbar2Tk  # Libreria para importar graficosTiempo dentro de la interfaz


# Creacion de la ventana y el archivo como objeto
# ------------------------------------------------


class Graficos:
    # Clase funciona como un machote para crear todos los graficosTiempo necesarios dentro de la interfaz
    canvas = None
    figure = None
    data = None
    frame = None

    def __init__(self, root, sheetName, columns,
                 frame):  # Esta funcion se llama metodo constructor, sirve para crear graficosTiempo con esta plantilla
        self.data = pd.read_excel(root, sheet_name=sheetName,
                                  usecols=columns)  # Se lee el archivo usando solo ciertas columnas

        self.frame = self.data.plot(columns[0], columns[1])  # Se crea la ventana grande del grafico
        self.figure = plt.Figure(figsize=(7, 6),
                                 dpi=110)  # Se crea el grafico como una figura para insertarla dentro de la interfaz

        self.figure.add_subplot().plot(self.data[columns[0]], self.data[columns[1]])  # Se agregon los ejes al grafico
        # self.main = Frame(master=frame, background="#3C4FA8", width=300, height=300)
        self.canvas = FigureCanvasTkAgg(self.figure, master=frame)  # se coloca el grafico dentro de la interfaz
        self.canvas.draw()  # Se dibuja el grafico

    def open(self):
        # metodo que funciona para abrir un grafico en pantalla grande
        plt.show(0)

    def placeFigure(self, _x, _y):
        # metodo que funciona para colocar un grafico dentro de la interfaz en un espacio determinado
        self.canvas.get_tk_widget().place(x=_x, y=_y)

    def exit(self):
        plt.close()

    def ocultar(self):
        self.canvas.get_tk_widget().place_forget()


class BigFrame(Frame):
    # Declaracion de la clase que sirve para implementar la interfaz
    graficosTiempo = []  # Tenemos una lista de graficosTiempo generales
    graficosFourier = []
    graficosFiltros = []
    fourier = None
    filtros = False
    graficoPos = []
    leftFrame = None
    rightFrame = None
    bottomFrame = None

    def __init__(self, master=None):  # Metodo constructor que sirve para crear la ventana general de la interfaz
        self.graficoPos = [0, 0]
        self.master = master
        self.master.geometry('1000x700')  # le damos el tamanio a la ventana
        self.master.title('Pronostico vibracion AV')  # Aqui le agregamos el titulo a la ventana
        self.master.configure(bg="white")  # le cambiamos el fondo a la ventana en color blanco
        super().__init__(master)
        self.createInternalFrames()
        self.fourier = tk.BooleanVar(self)

    def createInternalFrames(self):
        # Para dividir la pantalla en tres partes no iguales, modo de visualizacion de la interfaz en formato Z para mejor enfoque en lo importante (graficosTiempo)
        self.leftFrame = Frame(master, background="#3C4FA8", width=690,
                               height=620)  # agregar una seccion dentro de la ventana
        self.leftFrame.grid(row=0, column=0, padx=5, pady=5)  # Colocamos en la posicion 0 la seccion

        self.rightFrame = Frame(master, background="#3C4FA8", width=290, height=690)
        self.rightFrame.grid(row=0, column=1, pady=5, padx=5, rowspan=2)

        self.bottomFrame = Frame(master, background="#3C4FA8", width=690, height=60)
        self.bottomFrame.grid(row=1, column=0, pady=5, padx=5)

        # pestaña de graficosTiempo (superior izquierda)
        self.title = Label(self.leftFrame, background="#3C4FA8", text="Grafico de prueba", foreground="white",
                           font=("Times New Romans", "22", "bold"))  # agregar un texto en la seccion 1
        self.title.place(x=250, y=10)  # colocar el texto en la posicion deseada

        self.uploadGraphics()

        # pestaña de seleccion de graficosTiempo (inferior izquierda)
        self.bOpen = Button(self.bottomFrame, text="Pantalla completa",
                            command=lambda: self.openGraphic(self.graficoPos[0], self.graficoPos[1]))
        self.bOpen.place(x=10, y=15)

        self.cbGraphicType = ttk.Combobox(self.bottomFrame, state="readonly",
                                          values=["Axial", "Vertical", "Horizontal"])
        self.cbGraphicType.current(0)
        self.cbGraphicType.place(x=160, y=15)

        self.cbGraphicUses = ttk.Combobox(self.bottomFrame, state="readonly",
                                          values=["Velocidad", "Aceleracion", "Aceleracion Env"])
        self.cbGraphicUses.current(0)
        self.cbGraphicUses.place(x=380, y=15)

        self.btnAplicar = Button(self.bottomFrame, text="Aplicar",
                                 command=lambda: self.aplicarCambio())
        self.btnAplicar.place(x=600, y=15)

        # pestaña de configuraciones (superior-inferior derecho)

        self.title = Label(self.rightFrame, background="#3C4FA8", text="Funciones", foreground="white",
                           font=("Times New Romans", "22", "bold"))
        self.title.place(x=95, y=12)

        self.btnGFiltros = Button(self.rightFrame, width=20, height=1, text="Grafico Filtros",
                                  command=lambda: self.graficoF())
        self.btnGFiltros.place(x=40, y=80)

        self.btnFourier = Button(self.rightFrame, width=20, height=1, text="Grafico Filtros Fourier",
                                 command=lambda: self.hideGraphic())
        self.btnFourier.place(x=40, y=140)

        self.checkBoxFourier = ttk.Checkbutton(self.rightFrame, text="  Aplicar transformada de Fourier",
                                               variable=self.fourier)
        self.checkBoxFourier.place(x=40, y=200)

    def graficoF(self):
        self.graficosTiempo[self.graficoPos[0]][self.graficoPos[1]].ocultar()
        self.graficosFiltros[0].placeFigure(50, 60)
        self.cbGraphicType.configure(state=tk.DISABLED)
        self.cbGraphicUses.configure(state=tk.DISABLED)

    def hideGraphic(self):
        self.graficosTiempo[0][1].ocultar()

    def getIndex(self):
        if self.cbGraphicUses.get() == "Velocidad":
            if self.cbGraphicType.get() == "Axial":
                return [0, 0]
            elif self.cbGraphicType.get() == "Vertical":
                return [0, 1]
            elif self.cbGraphicType.get() == "Horizontal":
                return [0, 2]
        elif self.cbGraphicUses.get() == "Aceleracion":
            if self.cbGraphicType.get() == "Axial":
                return [1, 0]
            elif self.cbGraphicType.get() == "Vertical":
                return [1, 1]
            elif self.cbGraphicType.get() == "Horizontal":
                return [1, 2]
        elif self.cbGraphicUses.get() == "Aceleracion Env":
            if self.cbGraphicType.get() == "Axial":
                return [2, 0]
            elif self.cbGraphicType.get() == "Vertical":
                return [2, 1]
            elif self.cbGraphicType.get() == "Horizontal":
                return [2, 2]

    def uploadGraphics(self):
        self.graficosTiempo.append([Graficos("general.xlsx", "Vent Int A", ["Date", "mm/s"],
                                             self.leftFrame), Graficos("general.xlsx", "Vent Int A", ["Date", "g"],
                                                                       self.leftFrame),
                                    Graficos("general.xlsx", "Vent Int A", ["Date", "gE"],
                                             self.leftFrame)])  # Agregamos un grafico a la lista de graficosTiempo general
        self.graficosTiempo.append([Graficos("general.xlsx", "Vent Int V", ["Date", "mm/s"],
                                             self.leftFrame), Graficos("general.xlsx", "Vent Int V", ["Date", "g"],
                                                                       self.leftFrame),
                                    Graficos("general.xlsx", "Vent Int V", ["Date", "gE"],
                                             self.leftFrame)])
        self.graficosTiempo.append([Graficos("general.xlsx", "Vent Int H", ["Date", "mm/s"],
                                             self.leftFrame), Graficos("general.xlsx", "Vent Int H", ["Date", "mm/s"],
                                                                       self.leftFrame),
                                    Graficos("general.xlsx", "Vent Int H", ["Date", "mm/s"],
                                             self.leftFrame)])

        self.graficosFiltros.append(Graficos("filtros.xlsx", "filtros", ["Date", "a"], self.leftFrame))

        self.graficosTiempo[0][0].placeFigure(50, 60)

    def aplicarCambio(self):
        self.graficosFiltros[0].ocultar()
        self.cbGraphicType.configure(state=tk.ACTIVE)
        self.cbGraphicUses.configure(state=tk.ACTIVE)
        self.graficosTiempo[self.graficoPos[0]][self.graficoPos[1]].ocultar()
        self.graficoPos = self.getIndex()
        self.graficosTiempo[self.graficoPos[0]][self.graficoPos[1]].placeFigure(50, 60)

    def openGraphic(self, positionX=None, positionY=None):
        # Funcion para abrir un grafico en pantalla grande
        self.graficosTiempo[positionX][positionY].open()


if __name__ == "__main__":
    # Esta es la funcion principal que ejecuta todo el programa
    master = Tk()  # crea la ventana
    app = BigFrame(master)  # ejecuta todos los comandos para la personalizacion de la interfaz
    app.mainloop()  # se corre el programa

from tkinter import Tk, Frame, Button, Label, ttk, BooleanVar, IntVar, ACTIVE, DISABLED
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import scipy
from scipy import fft
import numpy as np


# Creacion de la ventana y el archivo como objeto
# ------------------------------------------------

class Grafico:
    canvas = None
    data = None
    frame = None

    def __init__(self, root, sheetName, columns, frame):
        self.ax = None
        self.window = None
        self.figure = None
        self.data = pd.read_excel(root, sheet_name=sheetName, usecols=columns)
        self.columns = columns
        self.frame = frame

    def fourier(self):
        N = 600
        T = 1 / 800

        dataNP = self.data[self.columns[1]].to_numpy()

        y = scipy.fft.fft(dataNP)
        x = np.linspace(0, 1 / (2 * T), N // 2)
        x = x[:len(y)]

        if self.figure is not None:
            plt.close(self.figure)
            self.canvas.get_tk_widget().destroy()
        self.inicializar()
        self.ax.plot(x, 2/N * np.abs(y[:N//2]))
        self.canvas = FigureCanvasTkAgg(self.figure, master=self.frame)
        self.canvas.draw()
        self.canvas.get_tk_widget().place(x=50, y=60)

    def figureExist(self):
        if self.figure is None:
            return False
        else:
            return True

    def cambiarX(self, limX, step):
        plt.close(self.figure)
        self.canvas.get_tk_widget().destroy()
        self.inicializar()
        self.ax.set_xlim([limX, 280])
        self.crear()
        self.ax.set_xlabel("Tiempo")
        self.ax.set_xticks(range(limX, 280, step))
        self.canvas = FigureCanvasTkAgg(self.figure, master=self.frame)
        self.canvas.draw()
        self.canvas.get_tk_widget().place(x=50, y=60)

    def open(self):
        plt.show()

    def inicializar(self):
        self.figure = plt.figure(dpi=50, figsize=(8, 7))
        self.ax = self.figure.subplots(1)

    def crear(self):
        self.ax.plot(self.data[self.columns[0]], self.data[self.columns[1]])

    def placeFigure(self):
        if self.figure is not None:
            plt.close(self.figure)
            self.canvas.get_tk_widget().destroy()
        self.inicializar()
        self.crear()
        self.ax.set_xlabel("Tiempo")
        self.ax.set_xticks(range(0, 280, 80))
        self.canvas = FigureCanvasTkAgg(self.figure, master=self.frame)
        self.canvas.draw()
        self.canvas.get_tk_widget().place(x=50, y=60)

    def exit(self):
        self.figure.close()

    def ocultar(self):
        self.canvas.get_tk_widget().destroy()


class BigFrame(Frame):
    graficosTiempo = []
    graficosFourier = []
    graficosFiltros = []
    fourier = None
    filtros = False
    graficoPos = []
    leftFrame = None
    rightFrame = None
    bottomFrame = None

    def __init__(self, master=None):
        self.graficoPos = [0, 0]
        self.master = master
        super().__init__(master)
        self.createInternalFrames()
        self.uploadGraphics()

    def createInternalFrames(self):
        self.leftFrame = Frame(master, background="#3C4FA8", width=690,
                               height=620)
        self.leftFrame.grid(row=0, column=0, padx=5, pady=5)

        self.rightFrame = Frame(master, background="#3C4FA8", width=290, height=620)
        self.rightFrame.grid(row=0, column=1, pady=5, padx=5)

        self.bottomFrame = Frame(master, background="#3C4FA8", width=990, height=60)
        self.bottomFrame.grid(row=1, column=0, pady=5, padx=5, columnspan=2)

        # pestaña de graficosTiempo (superior izquierda)
        self.title = Label(self.leftFrame, background="#3C4FA8", text="Grafico Actual", foreground="white",
                           font=("Times New Romans", "22", "bold"))
        self.title.place(x=250, y=10)

        # pestaña de seleccion de graficosTiempo (inferior izquierda-derecha)
        self.bOpen = Button(self.bottomFrame, text="Pantalla completa",
                            command=lambda: self.openGraphic(self.graficoPos[0], self.graficoPos[1]))
        self.bOpen.place(x=15, y=15)
        # self.bOpen.configure(state= DISABLED)

        self.cbGraphicType = ttk.Combobox(self.bottomFrame, state="readonly",
                                          values=["Axial", "Vertical", "Horizontal"])
        self.cbGraphicType.current(0)
        self.cbGraphicType.place(x=200, y=15)

        self.cbGraphicUses = ttk.Combobox(self.bottomFrame, state="readonly",
                                          values=["Velocidad", "Aceleracion", "Aceleracion Env"])
        self.cbGraphicUses.current(0)
        self.cbGraphicUses.place(x=415, y=15)

        self.cbGraphicPart = ttk.Combobox(self.bottomFrame, state="readonly",
                                          values=["Ventilador Int", "Ventilador Ext", "Motor Int"])
        self.cbGraphicPart.current(0)
        self.cbGraphicPart.place(x=630, y=15)

        self.btnAplicar = Button(self.bottomFrame, text="Aplicar",
                                 command=lambda: self.aplicarCambio(), width=6)
        self.btnAplicar.place(x=870, y=15)

        # pestaña de configuraciones (superior derecho)

        self.title = Label(self.rightFrame, background="#3C4FA8", text="Funciones", foreground="white",
                           font=("Times New Romans", "22", "bold"))
        self.title.place(x=95, y=12)

        self.btnGFiltros = Button(self.rightFrame, width=20, height=1, text="Grafico Filtros",
                                  command=lambda: self.graficoF())
        self.btnGFiltros.place(x=40, y=80)

        self.btnFourier = Button(self.rightFrame, width=20, height=1, text="Grafico sin definir",
                                 command=lambda: self.hideGraphic())
        self.btnFourier.place(x=40, y=140)

        self.fourier = BooleanVar(self)
        self.fourier.set(False)
        self.checkBoxFourier = ttk.Checkbutton(self.rightFrame, text=" Aplicar transformada de Fourier",
                                               variable=self.fourier, onvalue= True, offvalue= False)

        self.checkBoxFourier.place(x=40, y=200)

        self.btnConfiguraciones = Button(self.rightFrame, text="Configuraciones", width=20, height=1)
        self.btnConfiguraciones.place(x=40, y=260)

        self.cbTiempo = ttk.Combobox(self.rightFrame, state="readonly", values=["3 Meses", "6 Meses", "12 Meses"],
                                     width=20)
        self.cbTiempo.current(2)
        self.cbTiempo.place(x=40, y=320)

        self.btnCambiar = Button(self.rightFrame, text="Cambiar Tiempo", width=20, height=1,
                                 command=lambda: self.cambiarTiempo())
        self.btnCambiar.place(x=40, y=380)


    def cambiarTiempo(self):
        if self.cbTiempo.get() == "3 Meses":
            self.graficosTiempo[self.graficoPos[0]][self.graficoPos[1]].cambiarX(180, 30)
        elif self.cbTiempo.get() == "6 Meses":
            self.graficosTiempo[self.graficoPos[0]][self.graficoPos[1]].cambiarX(90, 50)
        else:
            self.graficosTiempo[self.graficoPos[0]][self.graficoPos[1]].ocultar()
            self.graficosTiempo[self.graficoPos[0]][self.graficoPos[1]].placeFigure()

    def graficoF(self):
        if self.graficosTiempo[self.graficoPos[0]][self.graficoPos[1]].figureExist():
            self.graficosTiempo[self.graficoPos[0]][self.graficoPos[1]].ocultar()

        self.graficosFiltros[0].placeFigure()
        self.cbGraphicType.configure(state=DISABLED)
        self.cbGraphicUses.configure(state=DISABLED)
        self.cbGraphicPart.configure(state=DISABLED)

    def hideGraphic(self):
        self.graficosTiempo[0][1].ocultar()

    def getIndex(self):
        if self.cbGraphicPart.get() == "Ventilador Int":
            if self.cbGraphicType.get() == "Axial":
                if self.cbGraphicUses.get() == "Velocidad":
                    return [0, 0]
                elif self.cbGraphicUses.get() == "Aceleracion":
                    return [0, 1]
                elif self.cbGraphicUses.get() == "Aceleracion Env":
                    return [0, 2]
            elif self.cbGraphicType.get() == "Vertical":
                if self.cbGraphicUses.get() == "Velocidad":
                    return [1, 0]
                elif self.cbGraphicUses.get() == "Aceleracion":
                    return [1, 1]
                elif self.cbGraphicUses.get() == "Aceleracion Env":
                    return [1, 2]
            elif self.cbGraphicType.get() == "Horizontal":
                if self.cbGraphicUses.get() == "Velocidad":
                    return [2, 0]
                elif self.cbGraphicUses.get() == "Aceleracion":
                    return [2, 1]
                elif self.cbGraphicUses.get() == "Aceleracion Env":
                    return [2, 2]

        elif self.cbGraphicPart.get() == "Ventilador Ext":
            if self.cbGraphicType.get() == "Axial":
                if self.cbGraphicUses.get() == "Velocidad":
                    return [3, 0]
                elif self.cbGraphicUses.get() == "Aceleracion":
                    return [3, 1]
                elif self.cbGraphicUses.get() == "Aceleracion Env":
                    return [3, 2]
            elif self.cbGraphicType.get() == "Vertical":
                if self.cbGraphicUses.get() == "Velocidad":
                    return [4, 0]
                elif self.cbGraphicUses.get() == "Aceleracion":
                    return [4, 1]
                elif self.cbGraphicUses.get() == "Aceleracion Env":
                    return [4, 2]
            elif self.cbGraphicType.get() == "Horizontal":
                if self.cbGraphicUses.get() == "Velocidad":
                    return [5, 0]
                elif self.cbGraphicUses.get() == "Aceleracion":
                    return [5, 1]
                elif self.cbGraphicUses.get() == "Aceleracion Env":
                    return [5, 2]
        else:
            if self.cbGraphicType.get() == "Axial":
                if self.cbGraphicUses.get() == "Velocidad":
                    return [6, 0]
                elif self.cbGraphicUses.get() == "Aceleracion":
                    return [6, 1]
                elif self.cbGraphicUses.get() == "Aceleracion Env":
                    return [6, 2]
            elif self.cbGraphicType.get() == "Vertical":
                if self.cbGraphicUses.get() == "Velocidad":
                    return [7, 0]
                elif self.cbGraphicUses.get() == "Aceleracion":
                    return [7, 1]
                elif self.cbGraphicUses.get() == "Aceleracion Env":
                    return [7, 2]
            elif self.cbGraphicType.get() == "Horizontal":
                if self.cbGraphicUses.get() == "Velocidad":
                    return [8, 0]
                elif self.cbGraphicUses.get() == "Aceleracion":
                    return [8, 1]
                elif self.cbGraphicUses.get() == "Aceleracion Env":
                    return [8, 2]

    def uploadGraphics(self):
        self.graficosTiempo.append([Grafico("general.xlsx", "Vent Int A", ["Date", "mm/s"], self.leftFrame),
                                    Grafico("general.xlsx", "Vent Int A", ["Date", "g"], self.leftFrame),
                                    Grafico("general.xlsx", "Vent Int A", ["Date", "gE"], self.leftFrame)])

        self.graficosTiempo.append([Grafico("general.xlsx", "Vent Int V", ["Date", "mm/s"], self.leftFrame),
                                    Grafico("general.xlsx", "Vent Int V", ["Date", "g"], self.leftFrame),
                                    Grafico("general.xlsx", "Vent Int V", ["Date", "gE"], self.leftFrame)])

        self.graficosTiempo.append([Grafico("general.xlsx", "Vent Int H", ["Date", "mm/s"], self.leftFrame),
                                    Grafico("general.xlsx", "Vent Int H", ["Date", "g"], self.leftFrame),
                                    Grafico("general.xlsx", "Vent Int H", ["Date", "gE"], self.leftFrame)])

        self.graficosTiempo.append([Grafico("general.xlsx", "Vent Ext A", ["Date", "mm/s"], self.leftFrame),
                                    Grafico("general.xlsx", "Vent Ext A", ["Date", "g"], self.leftFrame),
                                    Grafico("general.xlsx", "Vent Ext A", ["Date", "gE"], self.leftFrame)])

        self.graficosTiempo.append([Grafico("general.xlsx", "Vent Ext V", ["Date", "mm/s"], self.leftFrame),
                                    Grafico("general.xlsx", "Vent Ext V", ["Date", "g"], self.leftFrame),
                                    Grafico("general.xlsx", "Vent Ext V", ["Date", "gE"], self.leftFrame)])

        self.graficosTiempo.append([Grafico("general.xlsx", "Vent Ext H", ["Date", "mm/s"], self.leftFrame),
                                    Grafico("general.xlsx", "Vent Ext H", ["Date", "g"], self.leftFrame),
                                    Grafico("general.xlsx", "Vent Ext H", ["Date", "gE"], self.leftFrame)])

        self.graficosTiempo.append([Grafico("general.xlsx", "Motor Int A", ["Date", "mm/s"], self.leftFrame),
                                    Grafico("general.xlsx", "Motor Int A", ["Date", "g"], self.leftFrame),
                                    Grafico("general.xlsx", "Motor Int A", ["Date", "gE"], self.leftFrame)])

        self.graficosTiempo.append([Grafico("general.xlsx", "Motor Int V", ["Date", "mm/s"], self.leftFrame),
                                    Grafico("general.xlsx", "Motor Int V", ["Date", "g"], self.leftFrame),
                                    Grafico("general.xlsx", "Motor Int V", ["Date", "gE"], self.leftFrame)])

        self.graficosTiempo.append([Grafico("general.xlsx", "Motor Int H", ["Date", "mm/s"], self.leftFrame),
                                    Grafico("general.xlsx", "Motor Int H", ["Date", "g"], self.leftFrame),
                                    Grafico("general.xlsx", "Motor Int H", ["Date", "gE"], self.leftFrame)])

        self.graficosFiltros.append(Grafico("filtros.xlsx", "filtrosData", ["date", "meta"], self.leftFrame))

        self.graficosTiempo[0][0].placeFigure()


    def aplicarCambio(self):


        if self.fourier.get() == 1:
            if self.graficosFiltros[0].figureExist():
                self.graficosFiltros[0].ocultar()
                self.cbGraphicType.configure(state=ACTIVE)
                self.cbGraphicUses.configure(state=ACTIVE)
                self.cbGraphicPart.configure(state=ACTIVE)
            self.graficosTiempo[self.graficoPos[0]][self.graficoPos[1]].ocultar()
            self.graficoPos = self.getIndex()
            self.graficosTiempo[self.graficoPos[0]][self.graficoPos[1]].fourier()
            self.cbTiempo.configure(state=DISABLED)
            self.btnCambiar.configure(state=DISABLED)
        else:
            self.cbTiempo.configure(state=ACTIVE)
            self.btnCambiar.configure(state=ACTIVE)
            if self.graficosFiltros[0].figureExist():
                self.graficosFiltros[0].ocultar()
                self.cbGraphicType.configure(state=ACTIVE)
                self.cbGraphicUses.configure(state=ACTIVE)
                self.cbGraphicPart.configure(state=ACTIVE)
            self.graficosTiempo[self.graficoPos[0]][self.graficoPos[1]].ocultar()
            self.graficoPos = self.getIndex()
            self.graficosTiempo[self.graficoPos[0]][self.graficoPos[1]].placeFigure()

    def openGraphic(self, positionX=None, positionY=None):
        self.graficosTiempo[positionX][positionY].open()


if __name__ == "__main__":
    master = Tk()
    # master.attributes('-fullscreen', True)

    master.geometry("1000x700")
    master.title('Pronostico vibracion')
    master.configure(bg="white")
    app = BigFrame(master)
    app.mainloop()

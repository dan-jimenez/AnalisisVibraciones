import tkinter.font

import xlrd
import pandas as pd
from tkinter import *
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk


# Creacion de la ventana y el archivo como objeto
# ------------------------------------------------

data = pd.read_excel("test.xlsx", sheet_name="Vent Int A", usecols=["mm/s", "Date"])
window = Tk()
window.geometry('1000x800')
window.title('Pronostico vibracion AV')
window.configure(bg="white" )


# Declaracion de frame para divisiones de espacios
# -------------------------------------------------
leftFrame = Frame(window, background="#3C4FA8", width=690, height=720)
leftFrame.place(relx=300, rely=50)
leftFrame.grid(row=0, column=0, padx=5, pady=5)

rightFrame = Frame(window, background="#3C4FA8", width=290, height=720)
rightFrame.place(relx=300, rely=50)
rightFrame.grid(row=0, column=1, pady=5, padx=5)

bottomFrame = Frame(window, background="#3C4FA8", width=990, height=60)
bottomFrame.place(relx=300, rely=50)
bottomFrame.grid(row=1, column=0, pady=5, padx=5, columnspan=2)

# Colocar widgets dentro del frame de graficos
# -------------------------------------------------

title = Label(leftFrame, text="Grafico de prueba", background= "#3C4FA8", foreground= "white",
              font = ("Times New Romans", "22", "bold"))
title.place(x= 220,y= 10)



#prueba
figure = plt.Figure(figsize= (5,4), dpi= 100)
figure.add_subplot(111).plot(data["Date"], data["mm/s"])

canvas = FigureCanvasTkAgg(figure, master= leftFrame)
canvas.draw()
canvas.get_tk_widget().place(x = 40, y= 80)

if __name__ == "__main__":
    print(data)
    my_plot = data.plot("Date", "mm/s", kind="scatter")
    plt.show()
    window.mainloop()

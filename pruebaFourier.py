import pandas as pd
import scipy
from scipy import fft
import matplotlib.pyplot as plt
import numpy as np


data = pd.read_excel("test.xlsx", sheet_name= "Motor Int A")

N = 600
T = 1.0 / 800.0
y = data["mm/s"].to_numpy()
#x = np.linspace(0.0, N*T, N)
#y = np.sin(60.0 * 2.0*np.pi*x) + 0.5 * np.sin(90.0 * 2.0*np.pi*x)
y_f = scipy.fft.fft(y)
x_f = np.linspace(0.0, 1.0/(2.0*T), N//2)
P2 = np.abs(y_f/N)
P1 = P2[:(N//2)+1]
P1[2:len(P2)-1] = 2* P1[2:len(P2)-1]
plt.plot(x_f[:len(y_f)], 2.0/N * np.abs(y_f[:N//2]))
#plt.plot(x_f[:len(P1)], P1)
#plt.plot(x,y)

plt.show()




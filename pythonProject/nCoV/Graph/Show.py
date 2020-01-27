import numpy as np
from matplotlib import pyplot as plt
import matplotlib
import pandas as pd
class Show_graph():
    def __init__(self):
        pass
    def show_graph(self):
        x = np.linspace(20, 26, 7)
        y1 = [291, 440, 571, 830, 1287, 1975, 2762]
        plt.scatter(x, y1)
        plt.plot(x, y1, linewidth=1.2)
        plt.show()

# -*- coding: utf-8 -*-
"""
Created on Tue Dec 18 13:11:01 2018

@author: MENGstudents
"""

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Dec 18 12:54:57 2018

@author: Tom1
"""

import matplotlib
from tkinter import ttk
from numpy import arange, sin, pi
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2TkAgg
from matplotlib.figure import Figure
import tkinter as Tk
import pandas as pd
import numpy as np


root = Tk.Tk()
note = ttk.Notebook(root)
note.grid()

home_tab = ttk.Frame(note)
other_tab = ttk.Frame(note)
note.add(home_tab,text='test')
note.add(other_tab, text='other tab')
root.wm_title("Embedding in TK")


f = Figure(figsize=(5, 4), dpi=100)
a = f.add_subplot(111)
t = np.random.normal(3,2,1000)
ss = pd.DataFrame(columns=['a','b'])

_, binss, _ = a.hist(t, facecolor='blue', alpha=0.23)
a.hist(t[::2], facecolor='blue', bins=binss, alpha=0.7)

# a tk.DrawingArea
canvas = FigureCanvasTkAgg(f, master=home_tab)
canvas.draw()
canvas.get_tk_widget().grid()
canvas._tkcanvas.grid()

home_tab.forget(canvas)


Tk.mainloop()

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
t = arange(0.0, 3.0, 0.01)
s = sin(2*pi*t)

a.plot(t, s)

# a tk.DrawingArea
canvas = FigureCanvasTkAgg(f, master=home_tab)
canvas.draw()
canvas.get_tk_widget().grid()
canvas._tkcanvas.grid()


Tk.mainloop()

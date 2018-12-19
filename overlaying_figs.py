# -*- coding: utf-8 -*-
"""
Created on Tue Dec 18 22:03:45 2018

@author: MENGstudents
"""

import matplotlib.pyplot as plt
import numpy as np
#from mod import plotsomefunction
#from diffrentmod import plotsomeotherfunction

def plotsomefunction(ax, x):

    return ax.plot(x, np.sin(x))

def plotsomeotherfunction(ax, x):

    return ax.plot(x,np.cos(x))


fig, ax = plt.subplots(1,1)
x = np.linspace(0,np.pi,1000)
ax.plot(x, np.sin(x))
ax.plot(x, np.cos(x))
plt.show()
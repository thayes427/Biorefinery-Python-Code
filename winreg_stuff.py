# -*- coding: utf-8 -*-
"""
Created on Sun Jan 27 19:41:57 2019

@author: MENGstudents
"""

import winreg
key = winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, 'Apwn.Document')
winreg.SetValue(key, 'CLSID', winreg.REG_SZ, 'hi')
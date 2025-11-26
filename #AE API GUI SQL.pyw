#AE API GUI
import warnings
import pyodbc
from datetime import datetime, date, time, timedelta
from tkinter import *
from tkinter import messagebox, filedialog, ttk
import atexit
import time as ty
import threading
import numpy as np
from tkinter import simpledialog
import ctypes
from icecream import ic
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import subprocess
import os, sys
import glob
import json
import re
from bs4 import BeautifulSoup

# Add the parent directory ('NCC Automations') to the Python path
# This allows us to import the 'PythonTools' package from there.
parent_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(parent_dir)
from PythonTools import CREDS, EMAILS, sql_date_validation, PausableTimer, restart_pc #Both of these Variables are Dictionaries with a single layer that holds Personnel data or app passwords
from PythonTools import CREDS, EMAILS, sql_date_validation, PausableTimer, restart_pc, ToolTip #Both of these Variables are Dictionaries with a single layer that holds Personnel data or app passwords


import pandas as pd
from sklearn.linear_model import LinearRegression


site_widgets ={}

breaker_pulls = 6
meter_pulls = 8

start = ty.perf_counter()
myappid = 'AE.API.Data.GUI'
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

main_color = '#ADD8E6'
root = Tk()
root.title("Site Data")
try:
    root.iconbitmap(r"G:\Shared drives\O&M\NCC Automations\Icons\favicon.ico")
except Exception as e:
    print(f"Error loading icon: {e}")
root.wm_attributes("-topmost", True)
root.configure(bg="#ADD8E6") 

#Date Validation Registration
vcmd_date = (root.register(sql_date_validation), '%P')

checkIns= Toplevel(root)
try:
    checkIns.iconbitmap(r"G:\Shared drives\O&M\NCC Automations\Icons\favicon.ico")
except Exception as e:
    print(f"Error loading icon: {e}")
checkIns.title("Personnel On-Site")

checkIns.wm_attributes("-topmost", True)


timeWin= Toplevel(root)
try:
    timeWin.iconbitmap(r"G:\Shared drives\O&M\NCC Automations\Icons\favicon.ico")
except Exception as e:
    print(f"Error loading icon: {e}")
timeWin.title("Timestamps")
timeWin.wm_attributes("-topmost", True)
timeW = Frame(timeWin)
timeW.pack(side=LEFT)
timeW_notes= Label(timeW, text= "Data Pull Timestamps", font= ("Calibiri", 14))
timeW_notes.grid(row=0, column= 0, columnspan= 3)

# Fit in another 2 columns and make them sticky to eeach other like these below.
time1= Label(timeW, text= "First:", font= ("Calibiri", 12))
time2= Label(timeW, text= "Second:", font= ("Calibiri", 12))
time3= Label(timeW, text= "Third:", font= ("Calibiri", 12))
time4= Label(timeW, text= "Fourth:", font= ("Calibiri", 12))
time5= Label(timeW, text= "Tenth:", font= ("Calibiri", 12))
timeL= Label(timeW, text= "Fifteenth:", font= ("Calibiri", 12))
time1.grid(row=1, column= 0, sticky=E)
time2.grid(row=2, column= 0, sticky=E)
time3.grid(row=3, column = 0, sticky=E)
time4.grid(row=4, column = 0, sticky=E)
time5.grid(row=5, column = 0, sticky=E)
timeL.grid(row=6, column = 0, sticky=E)
time1v= Label(timeW, text= "Time")
time2v= Label(timeW, text= "Time")
time3v= Label(timeW, text= "Time")
time4v= Label(timeW, text= "Time")
time10v= Label(timeW, text= "Time")
timeLv= Label(timeW, text= "Time")
time1v.grid(row=1, column= 1, sticky=W)
time2v.grid(row=2, column= 1, sticky=W)
time3v.grid(row=3, column = 1, sticky=W)
time4v.grid(row=4, column = 1, sticky=W)
time10v.grid(row=5, column = 1, sticky=W)
timeLv.grid(row=6, column = 1, sticky=W)



datalbl= Label(timeW, text= "MsgBox Data:", font= ("Calibiri", 14))
datalbl.grid(row=1, column=2)
inverterT = Label(timeW, text= "Inverters:", font= ("Calibiri", 12))
inverterT.grid(row=2, column=2)
spread15 = Label(timeW, text= "Time")
spread15.grid(row=3, column=2)
breakermeter = Label(timeW, text= """Breakers &
Meters:""", font= ("Calibiri", 12))
breakermeter.grid(row=4, column=2)
spread10 = Label(timeW, text= "Time")
spread10.grid(row=5, column=2)



alertW = Toplevel(root)
alertW.title("Alert Windows Info")
alertW.wm_attributes("-topmost", True)
try:
    alertW.iconbitmap(r"G:\Shared drives\O&M\NCC Automations\Icons\favicon.ico")
except Exception as e:
    print(f"Error loading icon: {e}")

#Top Labels Main Window
siteLabel = Label(root, bg="#ADD8E6", text= "Sites", font=('Tk_defaultFont', 10, 'bold'))
siteLabel.grid(row=0, column= 0, sticky=W)
breakerstatusLabel= Label(root, bg="#ADD8E6", text= "Breaker Status", font=('Tk_defaultFont', 10, 'bold'))
breakerstatusLabel.grid(row=0, column=1)
meterVLabel = Label(root, bg="#ADD8E6", text= "Utility V", font=('Tk_defaultFont', 10, 'bold'))
meterVLabel.grid(row= 0, column=2)
meterkWLabel = Label(root, bg="#ADD8E6", text="Meter kW", font=('Tk_defaultFont', 10, 'bold'))
meterkWLabel.grid(row=0, column=4)
meterratioLabel = Label(root, bg="#ADD8E6", text= "% of Max", font=('Tk_defaultFont', 10, 'bold'))
meterratioLabel.grid(row=0, column= 5)
meterpvsystLabel = Label(root, bg="#ADD8E6", text= "% of PvSyst", font=('Tk_defaultFont', 10, 'bold'))
meterpvsystLabel.grid(row=0, column= 6)
POALabel = Label(root, bg="#ADD8E6", text= "POA", font=('Tk_defaultFont', 10, 'bold'))
POALabel.grid(row=0, column= 7)
SSSLabel = Label(root, bg="#ADD8E6", text= "Site Snap Shot", font=('Tk_defaultFont', 10, 'bold'))

# Configure column 8 to expand, allowing the snapshot frame to fill the available space
root.columnconfigure(8, weight=1)
SSSLabel.grid(row=0, column= 8)

#Windows with multiple pages of Site inv data
solrvr_win = Toplevel(root)
solrvr_win.title("Sol River's Portfolio")
try:
    solrvr_win.iconbitmap(r"G:\Shared drives\O&M\NCC Automations\Icons\favicon.ico")
except Exception as e:
    print(f"Error loading icon: {e}")
solrvrnotebook = ttk.Notebook(solrvr_win)
style = ttk.Style(root)
style.configure("TNotebook.Tab", padding=[90, 2], font=('Tk_defaultFont', 12, 'bold'))
solrvr = ttk.Frame(solrvrnotebook)
solrvr2 = ttk.Frame(solrvrnotebook)

solrvrnotebook.add(solrvr, text="Bulloch 1A - Sunflower")
solrvrnotebook.add(solrvr2, text="Upson - Williams")

solrvrnotebook.pack(expand=True, fill='both')

hst_win = Toplevel(root)
hst_win.title("Harrison Street's Portfolio")
try:
    hst_win.iconbitmap(r"G:\Shared drives\O&M\NCC Automations\Icons\favicon.ico")
except Exception as e:
    print(f"Error loading icon: {e}")
hstnotebook = ttk.Notebook(hst_win)
hst = ttk.Frame(hstnotebook)
hstnotebook.add(hst, text="Bishopville II - Van Buren")
hstnotebook.pack(expand=True, fill='both')

nar_win = Toplevel(root)
nar_win.title("NARENCO's Portfolio")
try:
    nar_win.iconbitmap(r"G:\Shared drives\O&M\NCC Automations\Icons\favicon.ico")
except Exception as e:
    print(f"Error loading icon: {e}")
narnotebook = ttk.Notebook(nar_win)
nar = ttk.Frame(narnotebook)
narnotebook.add(nar, text="Bluebird - Violet")
narnotebook.pack(expand=True, fill='both')

#Static Inv Windows
soltage = Toplevel(root)
soltage.title("Soltage")
soltage.wm_attributes("-topmost", True)

try:
    soltage.iconbitmap(r"G:\Shared drives\O&M\NCC Automations\Icons\favicon.ico")
except Exception as e:
    print(f"Error loading icon: {e}")
ncemc = Toplevel(root)
ncemc.title("NCEMC")
ncemc.wm_attributes("-topmost", True)
try:
    ncemc.iconbitmap(r"G:\Shared drives\O&M\NCC Automations\Icons\favicon.ico")
except Exception as e:
    print(f"Error loading icon: {e}")
#Inverter Windows created


MAP_SITES_HARDWARE_GUI = {
    'Bishopville II': {
        'INV_DICT': {i: f"{(i-1)//9 + 1}.{(i-1)%9 + 1}" for i in range(1, 37)},
        'METER_MAX': 9900000,
        'VAR_NAME': 'bishopvilleII',
        'CUST_ID': hst,
        'PVSYST': None,
        'BREAKER': True,
    },
    'Bluebird': {
        'INV_DICT': {i: f'A{i}' if i <= 12 else f'B{i}' for i in range(1, 25)},
        'METER_MAX': 3000000,
        'VAR_NAME': 'bluebird',
        'CUST_ID': nar,
        'PVSYST': 'BLUEBIRD',
        'BREAKER': False,
    },
    'Bulloch 1A': {
        'INV_DICT': {i: str(i) for i in range(1, 25)},
        'METER_MAX': 3000000,
        'VAR_NAME': 'bulloch1a',
        'CUST_ID': solrvr,
        'PVSYST': 'BULLOCH1A',
        'BREAKER': False,

    },
    'Bulloch 1B': {
        'INV_DICT': {i: str(i) for i in range(1, 25)},
        'METER_MAX': 3000000,
        'VAR_NAME': 'bulloch1b',
        'CUST_ID': solrvr,
        'PVSYST': 'BULLOCH1B',
        'BREAKER': False,

    },
    'Cardinal': {
        'INV_DICT': {i: str(i) for i in range(1, 60)},
        'METER_MAX': 7080000,
        'VAR_NAME': 'cardinal',
        'CUST_ID': nar, 
        'PVSYST': 'CARDINAL',
        'BREAKER': True,

    },
    'CDIA': {
        'INV_DICT': {1: '1'},
        'METER_MAX': 192000,
        'VAR_NAME': 'cdia',
        'CUST_ID': nar,
        'PVSYST': None,
        'BREAKER': False,

    },
    'Cherry Blossom': {
        'INV_DICT': {1: '1', 2: '2', 3: '3', 4: '4'},
        'METER_MAX': 10000000,
        'VAR_NAME': 'cherryblossom',
        'CUST_ID': nar,
        'PVSYST': 'CHERRY BLOSSOM',
        'BREAKER': True,

    },
    'Cougar': {
        'INV_DICT': {
            1: '1-1', 2: '1-2', 3: '1-3', 4: '1-4', 5: '1-5', 6: '2-1', 7: '2-2', 8: '2-3', 9: '2-4', 10: '2-5', 11: '2-6',
            12: '3-1', 13: '3-2', 14: '3-3', 15: '3-4', 16: '3-5', 17: '4-1', 18: '4-2', 19: '4-3', 20: '4-4', 21: '4-5',
            22: '5-1', 23: '5-2', 24: '5-3', 25: '5-4', 26: '5-5', 27: '6-1', 28: '6-2', 29: '6-3', 30: '6-4', 31: '6-5'
        },
        'METER_MAX': 2670000,
        'VAR_NAME': 'cougar',
        'CUST_ID': nar,
        'PVSYST': 'COUGAR',
        'BREAKER': False,

    },
    'Conetoe': {
        'INV_DICT': {i: f"{(i-1)//4 + 1}.{(i-1)%4 + 1}" for i in range(1, 17)},
        'METER_MAX': 5000000,
        'VAR_NAME': 'conetoe1',
        'CUST_ID': soltage,
        'PVSYST': None,
        'BREAKER': False,

    },
    'Duplin': {
        'INV_DICT': {i: f'C-{i}' if i <= 3 else f'S-{i-3}' for i in range(1,22)},
        'METER_MAX': 5040000,
        'VAR_NAME': 'duplin',
        'CUST_ID': soltage,
        'PVSYST': None,
        'BREAKER': False,

    },
    'Elk': {
        'INV_DICT': {
            1: '1', 2: '2', 3: '3', 4: '4', 5: '5', 6: '6', 7: '7', 8: '8', 9: '1-3', 10: '10', 11: '11', 12: '12', 13: '13',
            14: '14', 15: '3-3', 16: '16', 17: '17', 18: '18', 19: '19', 20: '20', 21: '21', 22: '22', 23: '23', 24: '24',
            25: '25', 26: '26', 27: '27', 28: '28', 29: '2-8', 30: '30', 31: '31', 32: '32', 33: '2-13', 34: '34', 35: '3-7',
            36: '3-8', 37: '37', 38: '38', 39: '3-11', 40: '40', 41: '41', 42: '42', 43: '43'
        },
        'METER_MAX': 5380000,
        'VAR_NAME': 'elk',
        'CUST_ID': solrvr,
        'PVSYST': 'ELK',
        'BREAKER': True,

    },
    'Freightliner': {
        'INV_DICT': {i: str(i) for i in range(1, 19)},
        'METER_MAX': 2250000,
        'VAR_NAME': 'freightliner',
        'CUST_ID': ncemc,
        'PVSYST': 'FREIGHTLINE',
        'BREAKER': False,

    },
    'Gray Fox': {
        'INV_DICT': {i: f"{(i-1)//20 + 1}.{(i-1)%20 + 1}" for i in range(1, 41)},
        'METER_MAX': 5000000,
        'VAR_NAME': 'grayfox',
        'CUST_ID': solrvr,
        'PVSYST': 'GRAYFOX',
        'BREAKER': True,

    },
    'Harding': {
        'INV_DICT': {i: str(i) for i in range(1, 25)},
        'METER_MAX': 3000000,
        'VAR_NAME': 'harding',
        'CUST_ID': solrvr,
        'PVSYST': 'HARDING',
        'BREAKER': True,

    },
    'Harrison': {
        'INV_DICT': {i: str(i) for i in range(1, 44)},
        'METER_MAX': 5380000,
        'VAR_NAME': 'harrison', 
        'CUST_ID': nar, 
        'PVSYST': 'HARRISON',
        'BREAKER': True,

    },
    'Hayes': {
        'INV_DICT': {i: str(i) for i in range(1, 27)},
        'METER_MAX': 3240000,
        'VAR_NAME': 'hayes',
        'CUST_ID': nar,
        'PVSYST': 'HAYES',
        'BREAKER': True,

    },
    'Hickory': {
        'INV_DICT': {1: '1', 2: '2'}, 
        'METER_MAX': 5000000,
        'VAR_NAME':'hickory',
        'CUST_ID': nar,
        'PVSYST': 'HICKORY',
        'BREAKER': True,

    },
    'Hickson': {
        'INV_DICT': {i: f"1-{i}" for i in range(1, 17)},
        'METER_MAX': 2000000,
        'VAR_NAME': 'hickson', 
        'CUST_ID': hst,
        'PVSYST': None,
        'BREAKER': True,

    },
    'Holly Swamp': {
        'INV_DICT': {i: str(i) for i in range(1, 17)},
        'METER_MAX': 2000000,
        'VAR_NAME': 'hollyswamp',
        'CUST_ID': ncemc,
        'PVSYST': 'HOLLYSWAMP',
        'BREAKER': False,

    },
    'Jefferson': {
        'INV_DICT': {i: f"{(i - 1) // 16 + 1}.{(i - 1) % 16 + 1}" for i in range(1, 65)},
        'METER_MAX': 8000000,
        'VAR_NAME': 'jefferson',
        'CUST_ID': hst,
        'PVSYST': None,
        'BREAKER': False,

    },
    'Longleaf Pine': {
        'INV_DICT': {i: f"A{i}" if i < 21 else f"B{i - 20}" for i in range(1, 41)},
        'METER_MAX': 5000000,
        'VAR_NAME': 'longleafpine',
        'CUST_ID': solrvr,
        'PVSYST': None,
        'BREAKER': True,

    },
    'Marshall': {
        'INV_DICT': {i: f"1.{i}" for i in range(1, 17)},
        'METER_MAX': 2000000,
        'VAR_NAME': 'marshall',
        'CUST_ID': hst,
        'PVSYST': None,
        'BREAKER': True,

    },
    'McLean': {
        'INV_DICT': {i: str(i) for i in range(1, 41)},
        'METER_MAX': 5000000,
        'VAR_NAME': 'mclean',
        'CUST_ID': solrvr,
        'PVSYST': 'MCLEAN',
        'BREAKER': True,

    },
    'Ogburn': {
        'INV_DICT': {i: f"1-{i}" for i in range(1, 17)},
        'METER_MAX': 2000000,
        'VAR_NAME': 'ogburn',
        'CUST_ID': hst,
        'PVSYST': None,
        'BREAKER': True,

    },
    'PG': {
        'INV_DICT': {i: str(i) for i in range(1, 19)},
        'METER_MAX': 2210000,
        'VAR_NAME': 'pg',
        'CUST_ID': ncemc,
        'PVSYST': 'PG',
        'BREAKER': False,

    },
    'Richmond': {
        'INV_DICT': {i: str(i) for i in range(1, 25)},
        'METER_MAX': 3000000,
        'VAR_NAME': 'richmond',
        'CUST_ID': solrvr,
        'PVSYST': 'RICHMOND',
        'BREAKER': False,

    },
    'Shorthorn': {
        'INV_DICT': {i: str(i) for i in range(1, 73)},
        'METER_MAX': 9000000,
        'VAR_NAME': 'shorthorn',
        'CUST_ID': solrvr,
        'PVSYST': 'SHORTHORN',
        'BREAKER': True,

    },
    'Sunflower': {
        'INV_DICT': {i: str(i) for i in range(1, 81)},
        'METER_MAX': 10000000,
        'VAR_NAME': 'sunflower',
        'CUST_ID': solrvr,
        'PVSYST': 'SUNFLOWER',
        'BREAKER': True,

    },
    'Tedder': {
        'INV_DICT': {i: str(i) for i in range(1, 17)},
        'METER_MAX': 2000000,
        'VAR_NAME': 'tedder',
        'CUST_ID': hst,
        'PVSYST': None,
        'BREAKER': True,

    },
    'Thunderhead': {
        'INV_DICT': {i: str(i) for i in range(1, 17)},
        'METER_MAX': 2000000,
        'VAR_NAME': 'thunderhead',
        'CUST_ID': hst,
        'PVSYST': None,
        'BREAKER': True,

    },
    'Upson': {
        'INV_DICT': {i: str(i) for i in range(1, 25)},
        'METER_MAX': 3000000,
        'VAR_NAME': 'upson',
        'CUST_ID': solrvr2,
        'PVSYST': None,
        'BREAKER': False,

    },
    'Van Buren': {
        'INV_DICT': {i: str(i) for i in range(1, 18)},
        'METER_MAX': 2000000,
        'VAR_NAME': 'vanburen',
        'CUST_ID': hst,
        'PVSYST': 'VAN BUREN',
        'BREAKER': False,

    },
    'Warbler': {
        'INV_DICT': {i: f"{'A' if i <= 16 else 'B'}{i}" for i in range(1, 33)},
        'METER_MAX': 4000000,
        'VAR_NAME': 'warbler',
        'CUST_ID': solrvr2,
        'PVSYST': 'WARBLER',
        'BREAKER': True,

    },
    'Washington': {
        'INV_DICT': {i: str(i) for i in range(1, 41)},
        'METER_MAX': 5000000,
        'VAR_NAME': 'washington',
        'CUST_ID': solrvr2,
        'PVSYST': None,
        'BREAKER': True,

    },
    'Wayne 1': {
        'INV_DICT': {1: '1', 2: '2', 3: '3', 4: '4'},
        'METER_MAX': 5000000,
        'VAR_NAME': 'wayne1',
        'CUST_ID': soltage,
        'PVSYST': None,
        'BREAKER': False,

    },
    'Wayne 2': {
        'INV_DICT': {1: '1', 2: '2', 3: '3', 4: '4'},
        'METER_MAX': 5000000,
        'VAR_NAME': 'wayne2',
        'CUST_ID': soltage,
        'PVSYST': None,
        'BREAKER': False,

    },
    'Wayne 3': {
        'INV_DICT': {1: '1', 2: '2', 3: '3', 4: '4'},
        'METER_MAX': 5000000,
        'VAR_NAME': 'wayne3',
        'CUST_ID': soltage,
        'PVSYST': None,
        'BREAKER': False,

    },
    'Wellons': {
        'INV_DICT': {1: '1-1', 2: '1-2', 3: '2-1', 4: '2-2', 5: '3-1', 6: '3-2'},
        'METER_MAX': 5000000,
        'VAR_NAME': 'wellons',
        'CUST_ID': nar,
        'PVSYST': 'WELLONS',
        'BREAKER': False,

    },
    'Whitehall': {
        'INV_DICT': {i: str(i) for i in range(1, 17)},
        'METER_MAX': 2000000,
        'VAR_NAME': 'whitehall',
        'CUST_ID': solrvr2,
        'PVSYST': 'WHITEHALL',
        'BREAKER': True,

    },
    'Whitetail': {
        'INV_DICT': {i: str(i) for i in range(1, 81)},
        'METER_MAX': 10000000,
        'VAR_NAME': 'whitetail',
        'CUST_ID': solrvr2,
        'PVSYST': None,
        'BREAKER': True,

    },
    'Williams': {
        'INV_DICT': {i: f"{'A' if i <= 20 else 'B'}{i}" for i in range(1, 41)},
        'METER_MAX': 5000000,
        'VAR_NAME': 'williams',
        'CUST_ID': solrvr2,
        'PVSYST': None,
        'BREAKER': True,

    },
    'Violet': {
        'INV_DICT': {1: '1', 2: '2'},
        'METER_MAX': 5000000,
        'VAR_NAME': 'violet',
        'CUST_ID': nar,
        'PVSYST': 'VIOLET',
        'BREAKER': True,

    }
}


all_CBs = []
normal_numbering = {'Bluebird', 'Cardinal', 'Cherry Blossom', 'Cougar', 'Harrison', 'Hayes', 'Hickory', 'Violet', 'HICKSON',
                    'JEFFERSON', 'Marshall', 'OGBURN', 'Tedder', 'Thunderhead', 'Van Buren', 'Bulloch 1A', 'Bulloch 1B', 'Elk', 'Duplin',
                    'Harding', 'Mclean', 'Richmond Cadle', 'Shorthorn', 'Sunflower', 'Upson', 'Warbler', 'Washington', 'Whitehall', 'Whitetail',
                    'Conetoe 1', 'Wayne I', 'Wayne II', 'Wayne III', 'Freight Line', 'Holly Swamp', 'PG'}

number20set = {'Gray Fox'}
number9set = {'BISHOPVILLE'}
number2set = {'Wellons Farm'}

def define_inv_num(site, group, num):
    group = int(group)
    num = int(num)

    if site in normal_numbering:
        return num
    elif site in number20set:
        inv = num+((20*group)-20)
        return inv
    elif site in number9set:
        inv = num+((9*group)-9)
        return inv
    elif site in number2set:
        inv = num+((2*group)-2)
        return inv


def open_wo_tracking(name):
    file = f"G:\\Shared drives\\O&M\\NCC Automations\\Notification System\\WO Tracking\\{name} Open WO's.txt"
    try:
        os.startfile(file)
    except FileNotFoundError:
        print(f"File Not Found: {file}")

#This loop creates all the tkinter widgets for the GUI
col = 1
for ro, (name, site_dictionary) in enumerate(MAP_SITES_HARDWARE_GUI.items(), start=1):
    invdict = site_dictionary['INV_DICT']
    var_name = site_dictionary['VAR_NAME']
    custid = site_dictionary['CUST_ID']
    breakertf = site_dictionary['BREAKER']

    invnum = len(invdict)
    #Site Info
    #Main Color
    globals()[f'{var_name}Label'] = Label(root, bg=main_color, text=name, fg= 'black', font=('Tk_defaultFont', 10, 'bold'))
    globals()[f'{var_name}Label'].grid(row=ro, column= 0, sticky=W)
    if breakertf:
        if name == 'Violet':
            vio_excep = 1
        else:
            vio_excep = ''
        globals()[f'{var_name}{vio_excep}statusLabel'] = Label(root, bg=main_color, text='❌', fg= 'black')
        globals()[f'{var_name}{vio_excep}statusLabel'].grid(row=ro, column= 1)
        if name == 'Violet':
            violet2statusLabel = Label(root, bg=main_color, text='❌', fg= 'black')
            violet2statusLabel.grid(row=ro+1, column= 1)

    if name != 'CDIA':
        #Site Voltage Boolean
        globals()[f'{var_name}meterVLabel'] = Label(root, bg=main_color, text='V', fg= 'black')
        globals()[f'{var_name}meterVLabel'].grid(row=ro, column= 2)

    globals()[f'{var_name}metercbval'] = IntVar()
    all_CBs.append(globals()[f'{var_name}metercbval'])
    globals()[f'{var_name}metercb'] = Checkbutton(root, bg=main_color, variable=globals()[f'{var_name}metercbval'], fg= 'black', cursor='hand2')
    globals()[f'{var_name}metercb'].grid(row=ro, column= 3)
    #Meter Producing Boolean
    globals()[f'{var_name}meterkWLabel'] = Label(root, bg=main_color, text='kW', fg= 'black')
    globals()[f'{var_name}meterkWLabel'].grid(row=ro, column= 4)
    #Meter % of Max capability
    globals()[f'{var_name}meterRatioLabel'] = Label(root, bg=main_color, text='Ratio', fg= 'black')
    globals()[f'{var_name}meterRatioLabel'].grid(row=ro, column= 5)
    #PVSyst Value
    globals()[f'{var_name}meterPvSystLabel'] = Label(root, bg=main_color, text='Ratio', fg= 'black')
    globals()[f'{var_name}meterPvSystLabel'].grid(row=ro, column= 6)

    globals()[f'{var_name}POAcbval'] = IntVar()
    all_CBs.append(globals()[f'{var_name}POAcbval'])
    globals()[f'{var_name}POAcb'] = Checkbutton(root, bg=main_color, text='X', variable=globals()[f'{var_name}POAcbval'], fg= 'black', cursor='hand2')
    globals()[f'{var_name}POAcb'].grid(row=ro, column= 7)
    
    # Create a frame for the site snapshot data
    globals()[f'{var_name}kwdata_frame'] = Frame(root, bg=main_color, bd=1, relief="solid")
    globals()[f'{var_name}kwdata_frame'].grid(row=ro, column=8, sticky='ew')

    # Configure columns to have equal weight, allowing them to share space
    globals()[f'{var_name}kwdata_frame'].columnconfigure(0, weight=1)
    globals()[f'{var_name}kwdata_frame'].columnconfigure(1, weight=1)
    globals()[f'{var_name}kwdata_frame'].columnconfigure(2, weight=1)

    # Initialize the site's widget dictionary
    site_widgets[var_name] = {
        'kwdata_frame': globals()[f'{var_name}kwdata_frame']
    }

    # Create 6 individual labels inside the frame, using sticky='ew' to make them fill their columns
    globals()[f'{var_name}_inv_kw_total'] = Label(globals()[f'{var_name}kwdata_frame'], bg=main_color, text='INV kW', fg='black')
    globals()[f'{var_name}_inv_kw_total'].grid(row=0, column=0, sticky='ew')
    globals()[f'{var_name}_meter_inv_diff'] = Label(globals()[f'{var_name}kwdata_frame'], bg=main_color, text='Meter-INVs', fg='black')
    globals()[f'{var_name}_meter_inv_diff'].grid(row=0, column=1, sticky='ew')
    globals()[f'{var_name}_invs_online'] = Label(globals()[f'{var_name}kwdata_frame'], bg=main_color, text='# INVs ✅', fg='black')
    globals()[f'{var_name}_invs_online'].grid(row=0, column=2, sticky='ew')
    globals()[f'{var_name}_meter_kw'] = Label(globals()[f'{var_name}kwdata_frame'], bg=main_color, text='Meter kW', fg='black')
    globals()[f'{var_name}_meter_kw'].grid(row=1, column=0, sticky='ew')
    globals()[f'{var_name}_invs_no_comms'] = Label(globals()[f'{var_name}kwdata_frame'], bg=main_color, text='No Comms', fg='black')
    globals()[f'{var_name}_invs_no_comms'].grid(row=1, column=1, sticky='ew')
    globals()[f'{var_name}_invs_total'] = Label(globals()[f'{var_name}kwdata_frame'], bg=main_color, text='Total INVs', fg='black')
    globals()[f'{var_name}_invs_total'].grid(row=1, column=2, sticky='ew')

    # Add the new labels to the site_widgets dictionary
    site_widgets[var_name]['inv_kw_total'] = globals()[f'{var_name}_inv_kw_total']
    site_widgets[var_name]['meter_inv_diff'] = globals()[f'{var_name}_meter_inv_diff']
    site_widgets[var_name]['invs_online'] = globals()[f'{var_name}_invs_online']
    site_widgets[var_name]['meter_kw'] = globals()[f'{var_name}_meter_kw']
    site_widgets[var_name]['invs_no_comms'] = globals()[f'{var_name}_invs_no_comms']
    site_widgets[var_name]['invs_total'] = globals()[f'{var_name}_invs_total']
    
    #End
    #INVERTER INFO
    length_limit = 73
    if name != 'CDIA':
        if invnum > length_limit:
            span_col = 6
        else:
            span_col = 3
        globals()[f'{var_name}invsLabel'] = Button(custid, text=name, command=lambda name=name: open_wo_tracking(name), bg=main_color, font=("Tk_defaultFont", 12, 'bold'), cursor='hand2')
        globals()[f'{var_name}invsLabel'].grid(row= 0, column= col, columnspan= span_col, sticky='ew')
    for num in range(1, invnum+1):
        column_offset = 0 if num <= length_limit else 3
        row_offset = num if num <= length_limit else num - length_limit
        if name != 'CDIA':
            if num in invdict:  # Check if the key exists in the dictionary
                inv_val = invdict[num]
            else:
                inv_val = str(num)

            globals()[f'{var_name}inv{inv_val}cbval'] = IntVar()
            all_CBs.append(globals()[f'{var_name}inv{inv_val}cbval'])
            globals()[f'{var_name}inv{inv_val}cb'] = Checkbutton(custid, text=str(inv_val), variable=globals()[f'{var_name}inv{inv_val}cbval'], cursor='hand2')
            globals()[f'{var_name}inv{inv_val}cb'].grid(row= row_offset, column= col+column_offset, sticky=W)

            globals()[f'{var_name}inv{num}WOLabel'] = Label(custid, text='⬜') #intial Setup of WO Placeholder. 
            globals()[f'{var_name}inv{num}WOLabel'].grid(row= row_offset, column= col+1+column_offset)

            if name != "Conetoe":
                globals()[f'{var_name}invup{num}cbval'] = IntVar()
                all_CBs.append(globals()[f'{var_name}invup{num}cbval'])
                globals()[f'{var_name}invup{num}cb'] = Checkbutton(custid, variable=globals()[f'{var_name}invup{num}cbval'], cursor='hand2')
                globals()[f'{var_name}invup{num}cb'].grid(row= row_offset, column= col+2+column_offset, sticky=W)
            else:
                if num < 5:
                    globals()[f'{var_name}invup{num}cbval'] = IntVar()
                    all_CBs.append(globals()[f'{var_name}invup{num}cbval'])
                    globals()[f'{var_name}invup{num}cb'] = Checkbutton(custid, variable=globals()[f'{var_name}invup{num}cbval'], cursor='hand2')
                    globals()[f'{var_name}invup{num}cb'].grid(row= (4*row_offset-3), rowspan= 4, column= col+2+column_offset, sticky=W)
    col += span_col


##########
##########
##########
##########
##########
#TEMPORARY CHECKS TO BE REMOVED WHEN DEVICE IS REPIARED
# check if False then call function to implement
conetoe_check = False
def conetoe_offline():
    global conetoe_check
    try:
        if int(hardingPOAcb.cget("text")) >= 400 and datetime.now().hour >= 9:
            msg= "Call Conetoe Utilities:\nWO 29980, 35307\n757-857-2888\nID: 710R41"
            if not textOnly.get():
                messagebox.showinfo(title="Comm Outage", parent=alertW, message=msg)
            else:
                text_update_Table.append("<br>" + str(msg))
            conetoe_check = True
    except ValueError:
        print("POA Value is not a valid Integer")


##########
##########
##########
##########
##########
##########
def connect_Logbook():
    global cur, lbconnection

    lbconn_str = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=G:\Shared drives\Narenco Projects\O&M Projects\NCC\NCC\NCC 039.accdb;'
    lbconnection = pyodbc.connect(lbconn_str)
    cur = lbconnection.cursor()


def connect_db():
    # Create a connection to the Access database
    globals()['dbconn_str'] = (
                r'DRIVER={ODBC Driver 18 for SQL Server};'
                r'SERVER=localhost\SQLEXPRESS01;'
                r'DATABASE=NARENCO_O&M_AE;'
                r'Trusted_Connection=yes;'
                r'Encrypt=no;'
            )
    globals()['dbconnection'] = pyodbc.connect(dbconn_str)
    globals()['c'] = dbconnection.cursor()

def launch_check():
    tday = datetime.now()
    format_date = tday.strftime('%m/%d/%Y')
    query = """
    SELECT TOP 16 [Timestamp] FROM [Whitetail Meter Data]
    WHERE FORMAT([Timestamp], 'MM/DD/YYYY') = ?
    """
    c.execute(query, (format_date,))
    data = c.fetchall()
    if len(data) == 16:
        #ic(data)
        return True
    else:
        #ic(data)
        return False
    

def last_online(site, inv_num, duplin_except):
    query = f"""
    SELECT TOP 1 [Timestamp] 
    FROM [{site}{duplin_except} INV {inv_num} Data]
    WHERE [Watts] > 2
    ORDER BY [Timestamp] DESC
    """
    c.execute(query)
    data = c.fetchone()
    if data:
        last_producing = f"Last Online: {data[0]}"
        return last_producing
    else:
        return None

def meter_last_online(site):
    query = f"""
    SELECT TOP 1 [Timestamp] 
    FROM [{site} Meter Data]
    WHERE [Watts] > 2
    ORDER BY [Timestamp] DESC
    """
    c.execute(query)
    data = c.fetchone()
    if data:
        last_producing = f"Last Online: {data[0]}"
        return last_producing
    else:
        return None
        
def last_closed(site):
    if site == "Violet":
        query1 = f"""
        SELECT TOP 1 [Timestamp] 
        FROM [{site} Breaker Data 1]
        WHERE [Status] = 1
        ORDER BY [Timestamp] DESC
        """
        c.execute(query1)
        data1 = c.fetchone()
        query2 = f"""
        SELECT TOP 1 [Timestamp] 
        FROM [{site} Breaker Data 2]
        WHERE [Status] = 1
        ORDER BY [Timestamp] DESC
        """
        c.execute(query2)
        data2 = c.fetchone()

        if data1 and data2:
            data = f"Last Closed Breaker 1: {data1[0]} | Breaker 2: {data2[0]}"
            return data
        else:
            return None
    elif site in ['Cardinal', 'Harrison', 'Hayes', 'Warbler']:
        query = f"""
        SELECT TOP 1 [Timestamp]
        FROM [{site} Meter Data]
        WHERE [Amps A] <> 0 AND [Amps B] <> 0 AND [Amps C] <> 0
        ORDER BY [Timestamp] DESC 
        """
        c.execute(query)
        data = c.fetchone()
        if data:
            last_breaker = f"Last Closed: {data[0]}"
            return last_breaker
        else:
            return None
    else:
        query = f"""
        SELECT TOP 1 [Timestamp] 
        FROM [{site} Breaker Data]
        WHERE [Status] = 1
        ORDER BY [Timestamp] DESC
        """
        c.execute(query)
        data = c.fetchone()
        if data:
            last_breaker = f"Last Closed: {data[0]}"
            return last_breaker
        else:
            return None

def check_inv_consecutively_online(alist):
    consecutive_count = 0
    
    for num in alist:
        if num > 0:
            consecutive_count += 1
            if consecutive_count == 2:
                return True
        else:
            consecutive_count = 0
    
    return False

def siteSnapShot_Update(site, var_name, inv_num, meterkW):
    global_vars = globals().copy()
    invs_ON = []
    invs_values = []
    
    for var, var_value in global_vars.items():
        #Filter variables to just the INV ones for site
        if (not isinstance(var_value, (Checkbutton, Label, Button)) or
                any(filt in var for filt in {'val', 'invup'}) or
                not all(filt in var for filt in {'inv', f'{var_name}', 'cb'})):
            continue

        if globals()[f'{var}'].cget('bg') == 'green':
            invs_ON.append(var)

    for inv in range(1, inv_num + 1):
        if site == 'Duplin':
            if inv <= 3:
                duplin_except = ' Central'
                inv_n = inv
            else:
                duplin_except = ' String'
                inv_n = inv -3
        else:
            duplin_except = ''
            inv_n = inv
        data = inv_data[f'{site}{duplin_except} INV {inv_n} Data']
        invmaxkW = max([row[1] for row in data[:meter_pulls]])
        invs_values.append(invmaxkW)
    
    no_0s_invs_values = [value for value in invs_values if value > 0]
    total_INVkW = sum(no_0s_invs_values)
    communicating_INVs = len(invs_ON) #Count of inverters who's bg is green

    if len(no_0s_invs_values) == 0 or meterkW <= 0:
        color = 'black'
    elif communicating_INVs == inv_num: #If all inverters are communicating
        color = main_color
    elif communicating_INVs < inv_num and meterkW >= total_INVkW + ((total_INVkW/len(no_0s_invs_values)) * 0.75)*(inv_num-communicating_INVs): #If theres a non-communicating INV and the Meter value is greater than Total of reporting INV's kW value (plus the avg of the inverters - 10 kW) * # of non Communicating inverters
        color = 'green'
    else: #Otherwise show yellow that theres a none communicating Inverter and it is offline according to meter
        color = 'yellow'

    # Update the individual labels and the frame background color
    site_widgets[var_name]['kwdata_frame'].config(bg=color)
    site_widgets[var_name]['inv_kw_total'].config(text=f"{total_INVkW/1000:.1f} kW", bg=color, font=('Tk_default', 9, 'bold'))
    site_widgets[var_name]['meter_inv_diff'].config(text=f"{round((meterkW - total_INVkW) / 1000, 1)} kW", bg=color, font=('Tk_default', 9, 'bold'))
    site_widgets[var_name]['invs_online'].config(text=f"{communicating_INVs}", bg=color, font=('Tk_default', 9, 'bold'))
    site_widgets[var_name]['meter_kw'].config(text=f"{meterkW/1000:.1f} kW", bg=color, font=('Tk_default', 9, 'bold'))
    site_widgets[var_name]['invs_no_comms'].config(text=f"{inv_num - communicating_INVs}", bg=color, font=('Tk_default', 9, 'bold'))
    site_widgets[var_name]['invs_total'].config(text=f"{inv_num}", bg=color, font=('Tk_default', 9, 'bold'))



def update_data():
    global text_update_Table
    save_cb_state()
    update_data_start = ty.perf_counter()

    now = datetime.now()
    if textOnly.get():
        # Create the email message
        message = MIMEMultipart()
        message["Subject"] = f"GUI Update {now}"
        message["From"] = EMAILS['NCC Desk']
        user = optionTexts.get()
        message["To"] = EMAILS[f'{user}'] 
        password = CREDS['remoteMonitoring']
        sender = EMAILS['NCC Desk']
    
    #How to Capture the Data from the GUI...
    text_update_Table = []
    html_start = """<html><head><div style="color:black; font-size:14pt; font-weight:bold; font-family:sans-serif;">
    GUI Update</div></head><body>"""
    #Initiate HTML Table Title turn these into an actual table so that each notification is a new row. Atleast for now to make sure that everything gets added
    text_update_Table.append(html_start)



    status_all = {}
    #Retireve Current INV status's and store in lists
    for name, site_dictionary in MAP_SITES_HARDWARE_GUI.items():
        invdict = site_dictionary['INV_DICT']
        var_name = site_dictionary['VAR_NAME']

        inverters = len(invdict)
        l = []
        if name != "CDIA":
            for i in range(1, inverters + 1):
                if i in invdict:  # Check if the key exists in the dictionary
                    inv_val = invdict[i]
                else:
                    inv_val = str(i)
                checkbox_var = globals()[f'{var_name}inv{inv_val}cbval']
                if not checkbox_var.get():
                    config_color = globals()[f'{var_name}inv{inv_val}cb'].cget("bg")
                    l.append(config_color)
                else:
                    l.append("green") #Checked inverters are known offline, this logs them in this check as 'online' but thats ok since this process is not the one used to check if the inverter is online or not. If we don't do this the message box reports the wrong index.
            status_all[f'{var_name}'] = l



    tm_now = datetime.now()
    str_tm_now = tm_now.strftime('%H')
    h_tm_now = int(str_tm_now)
    for name, site_dictionary in MAP_SITES_HARDWARE_GUI.items():
        invdict = site_dictionary['INV_DICT']
        var_name = site_dictionary['VAR_NAME']
        custid = site_dictionary['CUST_ID']
        metermax = site_dictionary['METER_MAX']
        pvsyst_name = site_dictionary['PVSYST']
        breakertf = site_dictionary['BREAKER']


        inverters = len(invdict)
        if name == "Violet":
            time_date_compare = (timecurrent - timedelta(hours=4))
        else:
            time_date_compare = (timecurrent - timedelta(hours=2))

        one_day_ago = (timecurrent - timedelta(days=1))
        allinv_kW = []

        if name != "CDIA":
            metercomms = max(comm_data[f'{name} Meter Data'])[0]



        #POA data Update
        if globals()[f'{var_name}POAcbval'].get() == 1:
            poa_noti = False
            poa = 9999
        else:
            poa = POA_data[f'{name} POA Data'][0]
            poa_noti = True
        #POA Comms check
        poa_data = max(comm_data[f'{name} POA Data'])[0]
        strtime_poa = poa_data.strftime('%m/%d/%y | %H:%M')
        if 800 <= poa < 2000:
            poa_color = '#ADD8E6'  # Light Blue
        elif 800 > poa >= 650:
            poa_color = '#87CEEB'  # Sky Blue
        elif 650 > poa >= 500:
            poa_color = '#1E90FF'  # Dodger Blue
        elif 500 > poa >= 350:
            poa_color = '#4682B4'  # Steel Blue
        elif 350 > poa >= 200:
            poa_color = '#4169E1'  # Royal Blue
        elif poa == 0:
            poa_color = 'black'
        else: 
            poa_color = 'gray'
        

        if poa_data < time_date_compare:
            poalbl = globals()[f'{var_name}POAcb'].cget('bg')
            if poalbl != 'pink' and poa_noti:
                msg = f"{name} lost comms with POA sensor at {strtime_poa}"
                if not textOnly.get():
                    messagebox.showwarning(parent= alertW, title=f"{name}, POA Comms Error", message=msg)
                else:
                    text_update_Table.append("<br>" + str(msg))
            globals()[f'{var_name}POAcb'].config(bg='pink', text=poa, font=('Tk_defaultFont', 10, 'bold'))
        else:
            globals()[f'{var_name}POAcb'].config(bg=poa_color, text=poa, font=('Tk_defaultFont', 10, 'bold'))


        master_cb_skips_INV_check = True if globals()[f'{var_name}metercbval'].get() == 0 else False
        #print(name, master_cb_skips_INV_check)


        #Breaker Update
        if breakertf:
            if name == "Violet":
                for two in range(1, 3):
                    breakercomm = max(comm_data[f'{name} Breaker Data {two}'])[0]
                    bk_Ltime = breakercomm.strftime('%m/%d/%y | %H:%M')
                    if breakercomm > time_date_compare:
                        breakerconfig = globals()[f'{var_name}{two}statusLabel'].cget("text")
                        if any(breaker_data[f'{name} Breaker Data {two}'][i][0] == True for i in range(breaker_pulls)):
                            breakerstatus = "✓✓✓"
                            breakerstatuscolor = 'green'
                        else:         
                            if breakerconfig != "❌❌" and master_cb_skips_INV_check:
                                last_operational = last_closed(name)
                                msg= f"{name} Breaker Tripped Open, Please Investigate! {last_operational}"
                                ToolTip( globals()[f'{var_name}{two}statusLabel'], f"{name} Breaker {two} is Open\n{last_operational}")
                                if not textOnly.get():
                                    messagebox.showerror(parent= alertW, title= f"{name}", message= msg)
                                else:
                                    text_update_Table.append("<br>" + str(msg))

                            last_operational = last_closed(name)
                            ToolTip( globals()[f'{var_name}{two}statusLabel'], f"{name} Breaker {two} is Open\n{last_operational}")
                            breakerstatus = "❌❌"
                            breakerstatuscolor = 'red'
                        globals()[f'{var_name}{two}statusLabel'].config(text= breakerstatus, bg=breakerstatuscolor)
                    else:
                        bklbl = globals()[f'{var_name}{two}statusLabel'].cget('bg')
                        globals()[f'{var_name}{two}statusLabel'].config(bg='pink')
                        if bklbl != 'pink' and master_cb_skips_INV_check:
                            last_operational = last_closed(name)
                            ToolTip( globals()[f'{var_name}statusLabel'], f"{name} Breaker Lost Comms\n{last_operational}")
                            msg= f"Breaker Comms lost {bk_Ltime} with the Breaker at {name}! Please Investigate!"
                            if not textOnly.get():
                                messagebox.showerror(parent= alertW, title=f"{name}, Breaker Comms Loss", message=msg)
                            else:
                                text_update_Table.append("<br>" + str(msg))
            elif name in {'Cardinal', 'Harrison', 'Hayes', 'Warbler', 'Hickory'}:
                if metercomms > time_date_compare:
                    rows_w_zeros = 0
                    for i in range(meter_pulls):
                        if any(meter_data[f'{name} Meter Data'][i][j] == 0 for j in range(3,6)):
                            rows_w_zeros += 1
                    if rows_w_zeros >= 2:
                        breakerconfig = globals()[f'{var_name}statusLabel'].cget("text")
                        if breakerconfig != "❌❌" and master_cb_skips_INV_check:
                            last_operational = last_closed(name)
                            msg= f"{name} Breaker Tripped Open, Please Investigate! {last_operational}"
                            if not textOnly.get():
                                messagebox.showerror(parent= alertW, title= f"{name}", message= msg)
                            else:
                                text_update_Table.append("<br>" + str(msg))
                        
                        last_operational = last_closed(name)
                        ToolTip( globals()[f'{var_name}statusLabel'], f"{name} Breaker is Open\n{last_operational}")
                        breakerstatus = "❌❌"
                        breakerstatuscolor = 'red'     
                    else:
                            breakerstatus = "✓✓✓"
                            breakerstatuscolor = 'green'
                
                    ToolTip( globals()[f'{var_name}statusLabel'], f"{name} Breaker is Operational")
                    globals()[f'{var_name}statusLabel'].config(text= breakerstatus, bg=breakerstatuscolor)
                else:
                    metercomms_time = metercomms.strftime('%m/%d/%y | %H:%M')
                    bklbl = globals()[f'{var_name}statusLabel'].cget('bg')
                    globals()[f'{var_name}statusLabel'].config(bg='pink')
                    if bklbl != 'pink' and master_cb_skips_INV_check:
                        last_operational = last_closed(name)
                        ToolTip( globals()[f'{var_name}statusLabel'], f"{name} Breaker Lost Comms\n{last_operational}")
                        msg= f"Meter Comms lost {metercomms_time} with the Meter at {name}! Please Investigate!"
                        if not textOnly.get():
                            messagebox.showerror(parent= alertW, title=f"{name}, Meter Comms Loss", message=msg)   
                        else:
                            text_update_Table.append("<br>" + str(msg))
            else:
                breakercomm = max(comm_data[f'{name} Breaker Data'])[0]
                bk_Ltime = breakercomm.strftime('%m/%d/%y | %H:%M')

                if breakercomm > time_date_compare:
                    breakerconfig = globals()[f'{var_name}statusLabel'].cget("text")

                    if any(breaker_data[f'{name} Breaker Data'][i][0] == True for i in range(breaker_pulls)):
                        breakerstatus = "✓✓✓"
                        breakerstatuscolor = 'green'
                        ToolTip( globals()[f'{var_name}statusLabel'], f"{name} Breaker is Operational")
                    else:         
                        if breakerconfig != "❌❌" and master_cb_skips_INV_check:
                            last_operational = last_closed(name)
                            ToolTip( globals()[f'{var_name}statusLabel'], f"{name} Breaker is Open\n{last_operational}")
                            msg= f"{name} Breaker Tripped Open, Please Investigate! {last_operational}"
                            if not textOnly.get():
                                messagebox.showerror(parent= alertW, title= f"{name}", message= msg)
                            else:
                                text_update_Table.append("<br>" + str(msg))
                        last_operational = last_closed(name)
                        ToolTip( globals()[f'{var_name}statusLabel'], f"{name} Breaker is Open\n{last_operational}")
                        breakerstatus = "❌❌"
                        breakerstatuscolor = 'red'
                    globals()[f'{var_name}statusLabel'].config(text= breakerstatus, bg=breakerstatuscolor)
                else:
                    bklbl = globals()[f'{var_name}statusLabel'].cget('bg')
                    globals()[f'{var_name}statusLabel'].config(bg='pink')
                    if bklbl != 'pink' and master_cb_skips_INV_check:
                        msg= f"Breaker Comms lost {bk_Ltime} with the Breaker at {name}! Please Investigate!"
                        last_operational = last_closed(name)
                        ToolTip( globals()[f'{var_name}statusLabel'], f"{name} Breaker Lost Comms\n{last_operational}")
                        if not textOnly.get():
                            messagebox.showerror(parent= alertW, title=f"{name}, Breaker Comms Loss", message= msg)
                        else:
                            text_update_Table.append("<br>" + str(msg))
        #INVERTER CHECKS            
        if name == "CDIA":
                data = inv_data[f'{name} INV 1 Data']
                current_config = globals()[f'{var_name}meterkWLabel'].cget("bg")
                cbval = globals()[f'{var_name}metercbval'].get()
                duplin_except = ''
                r = 1
                
                #Meter Ratio Check for CDIA
                avg_kW = np.mean([row[1] for row in data])
                meterRatio = avg_kW/metermax
                if meterRatio > .90:
                    ratio_color = '#ADD8E6'  # Light Blue
                elif .90 > meterRatio > .80:
                    ratio_color = '#87CEEB'  # Sky Blue
                elif .80 > meterRatio > .70:
                    ratio_color = '#1E90FF'  # Dodger Blue
                elif .70 > meterRatio > .60:
                    ratio_color = '#4682B4'  # Steel Blue
                elif .60 > meterRatio > .50:
                    ratio_color = '#4169E1'  # Royal Blue
                elif meterRatio < 0.001:
                    ratio_color = 'black'
                else: 
                    ratio_color = 'gray'
                globals()[f'{var_name}meterRatioLabel'].config(text= f"{round(meterRatio*100, 1)}%", bg= ratio_color, font=('Tk_defaultFont', 10, 'bold'))

                avg_dcv = np.mean([row[0] for row in data])
                inv_comm = max(comm_data[f'{name} INV 1 Data'])[0]
                
                inv_Ltime = inv_comm.strftime('%m/%d/%y | %H:%M')
                if inv_comm > time_date_compare:
                    if all(point[1] <= 1 for point in data):
                        if avg_dcv > 100:
                            if current_config not in ["orange", "red"] and ((poa > 250 and begin) or (h_tm_now >= 10 and poa > 100)) and cbval == 0 and master_cb_skips_INV_check:
                                var_key = f'{var_name}statusLabel'
                                if var_key in globals():
                                    if globals()[var_key].cget("bg") == 'green':
                                        online_last = last_online(name, r, duplin_except)
                                        msg= f"{name} | Inverter Offline, Good DC Voltage | {online_last}"
                                        if not textOnly.get():
                                            messagebox.showwarning(title=f"{name}", parent= alertW, message= msg)
                                        else:
                                            text_update_Table.append("<br>" + str(msg))
                                else:
                                    online_last = last_online(name, r, duplin_except)
                                    msg= f"{name} | Inverter Offline, Good DC Voltage | {online_last}"
                                    if not textOnly.get():
                                        messagebox.showwarning(title=f"{name}", parent= alertW, message= msg)
                                    else:
                                        text_update_Table.append("<br>" + str(msg))

                            globals()[f'{var_name}meterkWLabel'].config(text="X✓", bg='orange', font=('Tk_defaultFont', 10, 'bold'))
                            globals()[f'{var_name}Label'].config(bg='orange')

                        else:
                            online_last = last_online(name, inv_num, duplin_except)
                            ToolTip(globals()[f'{var_name}inv{inv_val}cb'], f"Inverter Offline, Bad DC Voltage\n{online_last}")
                            if current_config not in ["orange", "red"] and ((poa > 250 and begin) or (h_tm_now >= 10 and poa > 100)) and cbval == 0 and master_cb_skips_INV_check:
                                var_key = f'{var_name}statusLabel'
                                if var_key in globals():
                                    if globals()[var_key].cget("bg") == 'green':
                                        online_last = last_online(name, r, duplin_except)
                                        msg= f"{name} | Inverter Offline, Bad DC Voltage | {online_last}"
                                        if not textOnly.get():
                                            messagebox.showwarning(title=f"{name}", parent= alertW, message= msg)
                                        else:
                                            text_update_Table.append("<br>" + str(msg))
                                else:
                                    online_last = last_online(name, r, duplin_except)
                                    msg= f"{name} | Inverter Offline, Bad DC Voltage | {online_last}"
                                    if not textOnly.get():
                                        messagebox.showwarning(title=f"{name}", parent= alertW, message= msg)
                                    else:
                                        text_update_Table.append("<br>" + str(msg))

                            globals()[f'{var_name}meterkWLabel'].config(text="❌❌", bg='red')
                            globals()[f'{var_name}Label'].config(bg='red')

                    else:
                        if check_inv_consecutively_online(point[1] for point in data):
                            globals()[f'{var_name}meterkWLabel'].config(text=round(avg_kW/1000, 1), bg='green', font=('Tk_defaultFont', 10, 'bold'))
                            globals()[f'{var_name}Label'].config(bg='#ADD8E6')

                else:
                    invlbl = globals()[f'{var_name}meterkWLabel'].cget('bg')
                    globals()[f'{var_name}meterkWLabel'].config(bg='pink')
                    globals()[f'{var_name}meterRatioLabel'].config(bg='pink')

                    if invlbl != 'pink' and cbval == 0 and master_cb_skips_INV_check:
                        var_key = f'{var_name}statusLabel'
                        if var_key in globals():
                            if globals()[var_key].cget("bg") == 'green':
                                msg= f"INV Comms lost {inv_Ltime} with Inverter {r} at {name}! Please Investigate!"
                                if not textOnly.get():
                                    messagebox.showerror(parent= alertW, title=f"{name}, Inverter Comms Loss", message=msg)
                                else:
                                    text_update_Table.append("<br>" + str(msg))
                        else:
                            msg= f"INV Comms lost {inv_Ltime} with Inverter {r} at {name}! Please Investigate!"
                            if not textOnly.get():
                                messagebox.showerror(parent= alertW, title=f"{name}, Inverter Comms Loss", message=msg)
                            else:
                                text_update_Table.append("<br>" + str(msg))
        else:
            for r in range(1, inverters + 1):
                if r in invdict:  # Check if the key exists in the dictionary
                    inv_val = invdict[r]
                else:
                    inv_val = str(r)
                if name == 'Duplin':
                    if r <= 3:
                        duplin_except = ' Central'
                        inv_num = r
                    else:
                        duplin_except = ' String'
                        inv_num = r -3
                else:
                    duplin_except = ''
                    inv_num = r
                data = inv_data[f'{name}{duplin_except} INV {inv_num} Data']
                invmaxkW = max([row[1] for row in data[:meter_pulls]])
                
                current_config = globals()[f'{var_name}inv{inv_val}cb'].cget("bg")
                cbval = globals()[f'{var_name}inv{inv_val}cbval'].get()
                
                avg_dcv = np.mean([row[0] for row in data])
                inv_comm = max(comm_data[f'{name}{duplin_except} INV {inv_num} Data'])[0]

                allinv_kW.append(invmaxkW if inv_comm > time_date_compare else 0)
                inv_Ltime = inv_comm.strftime('%m/%d/%y | %H:%M')
                if inv_comm > time_date_compare:
                    if all(point[1] < 1 for point in data):
                        if avg_dcv > 100:
                            online_last = last_online(name, inv_num, duplin_except)
                            ToolTip(globals()[f'{var_name}inv{inv_val}cb'], f"Inverter Offline, Good DC Voltage\n{online_last}")
                            if current_config not in ["orange", "red"] and ((poa > 250 and begin) or (h_tm_now >= 10 and poa > 100)) and cbval == 0 and master_cb_skips_INV_check:
                                var_key = f'{var_name}statusLabel'
                                if var_key in globals():
                                    if globals()[var_key].cget("bg") == 'green':
                                        online_last = last_online(name, inv_num, duplin_except)
                                        msg= f"{name} | Inverter {inv_val} Offline, Good DC Voltage | {online_last}"
                                        if not textOnly.get():
                                            messagebox.showwarning(title=f"{name}", parent= alertW, message= msg)
                                        else:
                                            text_update_Table.append("<br>" + str(msg))
                                else:
                                    online_last = last_online(name, inv_num, duplin_except)
                                    msg= f"{name} | Inverter {inv_val} Offline, Good DC Voltage | {online_last}"
                                    if not textOnly.get():
                                        messagebox.showwarning(title=f"{name}", parent= alertW, message= msg)
                                    else:
                                        text_update_Table.append("<br>" + str(msg))
                            globals()[f'{var_name}inv{inv_val}cb'].config(bg='orange')
                        else:
                            online_last = last_online(name, inv_num, duplin_except)
                            ToolTip(globals()[f'{var_name}inv{inv_val}cb'], f"Inverter Offline, Bad DC Voltage\n{online_last}")
                            if current_config not in ["orange", "red"] and ((poa > 250 and begin) or (h_tm_now >= 10 and poa > 100)) and cbval == 0 and master_cb_skips_INV_check:
                                var_key = f'{var_name}statusLabel'
                                if var_key in globals():
                                    if globals()[var_key].cget("bg") == 'green':
                                        online_last = last_online(name, inv_num, duplin_except)
                                        msg= f"{name} | Inverter {inv_val} Offline, Bad DC Voltage | {online_last}"
                                        if not textOnly.get():
                                            messagebox.showwarning(title=f"{name}", parent= alertW, message= msg)
                                        else:
                                            text_update_Table.append("<br>" + str(msg))                                            
                                else:
                                    online_last = last_online(name, inv_num, duplin_except)
                                    msg= f"{name} | Inverter {inv_val} Offline, Bad DC Voltage | {online_last}"
                                    if not textOnly.get():
                                        messagebox.showwarning(title=f"{name}", parent= alertW, message= msg)
                                    else:
                                        text_update_Table.append("<br>" + str(msg))                                        

                            globals()[f'{var_name}inv{inv_val}cb'].config(bg='red')
                    else:
                        if check_inv_consecutively_online(point[1] for point in data):
                            globals()[f'{var_name}inv{inv_val}cb'].config(bg='green')
                            ToolTip(globals()[f'{var_name}inv{inv_val}cb'], "Inverter Online")
                else:
                    online_last = last_online(name, inv_num, duplin_except)
                    ToolTip(globals()[f'{var_name}inv{inv_val}cb'], f"Inverter Comms Lost\n{online_last}")
                    globals()[f'{var_name}inv{inv_val}cb'].config(bg='pink')
                    invlbl = globals()[f'{var_name}inv{inv_val}cb'].cget('bg')
                    if invlbl != 'pink' and cbval == 0 and master_cb_skips_INV_check:
                        var_key = f'{var_name}statusLabel'
                        if var_key in globals():
                            if globals()[var_key].cget("bg") == 'green':
                                msg= f"INV Comms lost {inv_Ltime} with Inverter {inv_val} at {name}! Please Investigate!"

                                if not textOnly.get():
                                    messagebox.showerror(parent= alertW, title=f"{name}, Inverter Comms Loss", message=msg)
                                else:
                                    text_update_Table.append("<br>" + str(msg))                                   
                        else:
                            msg= f"INV Comms lost {inv_Ltime} with Inverter {inv_val} at {name}! Please Investigate!"
                            if not textOnly.get():
                                online_last = last_online(name, inv_num, duplin_except)
                                ToolTip(globals()[f'{var_name}inv{inv_val}cb'], f"Inverter Comms Lost\n{online_last}")
                                messagebox.showerror(parent= alertW, title=f"{name}, Inverter Comms Loss", message=msg)
                            else:
                                text_update_Table.append("<br>" + str(msg))
        #Meter Check
        erroneous_meter_value = 760000000
        if name != "CDIA":
            meter_Ltime = metercomms.strftime('%m/%d/%y | %H:%M')
            if metercomms > time_date_compare:
                meterdata = meter_data[f'{name} Meter Data']
                meterdataVA= None
                meterdataVB= None
                meterdataVC= None 
                meterdataW= None 
                meterdataWM= None 
                meterdataAA= None 
                meterdataAB= None 
                meterdataAC= None

                # Iterate over fetched rows to find the maximum value for each column
                for row in meterdata:
                    # Update maximum values for each column
                    meterdataVA = max(meterdataVA, row[0]) if meterdataVA is not None else row[0]
                    meterdataVB = max(meterdataVB, row[1]) if meterdataVB is not None else row[1]
                    meterdataVC = max(meterdataVC, row[2]) if meterdataVC is not None else row[2]
                meterdataavgVA = np.mean([row[0] for row in meterdata if row[0] is not None])
                meterdataavgVB = np.mean([row[1] for row in meterdata if row[1] is not None])
                meterdataavgVC = np.mean([row[2] for row in meterdata if row[2] is not None])

                if meterdataavgVA != 0 and meterdataavgVB != 0 and meterdataavgVC !=0:
                    percent_difference_AB = ((max(meterdataavgVA, meterdataavgVB) - min(meterdataavgVA, meterdataavgVB)) / np.mean([meterdataavgVA, meterdataavgVB]))  * 100
                    percent_difference_AC = ((max(meterdataavgVA, meterdataavgVC) - min(meterdataavgVA, meterdataavgVC)) / np.mean([meterdataavgVA, meterdataavgVC]))  * 100
                    percent_difference_BC = ((max(meterdataavgVC, meterdataavgVB) - min(meterdataavgVC, meterdataavgVB)) / np.mean([meterdataavgVC, meterdataavgVB]))  * 100
                    #print(f'{name} | AB: {percent_difference_AB} AC: {percent_difference_AC} BC: {percent_difference_BC}')
                else:
                    percent_difference_AB = 0
                    percent_difference_AC = 0
                    percent_difference_BC = 0

                meterVconfig = globals()[f'{var_name}meterVLabel'].cget("text")

                meterdataavgAA = np.mean([row[3] for row in meterdata if row[3] is not None])
                meterdataavgAB = np.mean([row[4] for row in meterdata if row[4] is not None])
                meterdataavgAC = np.mean([row[5] for row in meterdata if row[5] is not None])
                meterdataAA = all(row[3] < 1 for row in meterdata if row[3] is not None)
                meterdataAB = all(row[4] < 1 for row in meterdata if row[4] is not None)
                meterdataAC = all(row[5] < 1 for row in meterdata if row[5] is not None)
                meterdataW = np.mean([row[6] for row in meterdata if row[6] is not None and row[6] < erroneous_meter_value])
                print(name, " ", meterdataW)
                if name == "Wellons":
                    meterdataWM = max(row[6] for row in meterdata if row[6] is not None and row[6] < erroneous_meter_value) if max(row[6] for row in meterdata if row[6] is not None and row[6] < erroneous_meter_value) else 0
                else:
                    meterdataWM = max(row[6] for row in meterdata if row[6] is not None) if max(row[6] for row in meterdata if row[6] is not None) else 0
                print(name, " ", meterdataWM)

                
                #print(f'{name} |  A: {meterdataAA}, B: {meterdataAB}, C: {meterdataAC}')
                #Accounting for Sites reporting Votlage differently
                if name == "Hickory":
                    val = 5
                else:
                    val = 5000

                if name in  ["Wellons", "Cherry Blossom"]:
                    dif = 9
                else:
                    dif = 5

                if meterdataVA < val and meterdataVB < val and meterdataVC < val:
                    meterVstatus= '❌❌'
                    meterVstatuscolor= 'red'
                    if meterVconfig != '❌❌':
                        online = meter_last_online(name)
                        msg= f"Loss of Utility Voltage or Lost Comms with Meter. {online}"
                        if not textOnly.get():
                            messagebox.showerror(parent=alertW, title= f"{name} Meter", message= msg)
                        else:
                            text_update_Table.append("<br>" + str(name) + " | " + str(msg))
                elif meterdataVA < val:
                    meterVstatus= 'X✓✓'
                    meterVstatuscolor= 'orange'
                    if meterVconfig != 'X✓✓':
                        msg= f"Loss of Utility Phase A Voltage or Lost Comms with Meter."
                        if not textOnly.get():
                            messagebox.showerror(parent=alertW, title= f"{name} Meter", message= msg)
                        else:
                            text_update_Table.append("<br>" + str(name) + " | " + str(msg))
                elif meterdataVB < val:
                    meterVstatus= '✓X✓'
                    meterVstatuscolor= 'orange'
                    if meterVconfig != '✓X✓':
                        msg= f"Loss of Utility Phase B Voltage or Lost Comms with Meter."
                        if not textOnly.get():
                            messagebox.showerror(parent=alertW, title= f"{name} Meter", message= msg)
                        else:
                            text_update_Table.append("<br>" + str(name) + " | " + str(msg))
                elif meterdataVC < val:
                    meterVstatus= '✓✓X'
                    meterVstatuscolor= 'orange'
                    if meterVconfig != '✓✓X':
                        msg= f"Loss of Utility Phase C Voltage or Lost Comms with Meter."
                        if not textOnly.get():
                            messagebox.showerror(parent=alertW, title= f"{name} Meter", message= msg)
                        else:
                            text_update_Table.append("<br>" + str(name) + " | " + str(msg))
                elif percent_difference_AB >= dif or percent_difference_AC >= dif or percent_difference_BC >= dif:
                    meterVstatus= '???'
                    meterVstatuscolor= 'orange'
                    if meterVconfig != '???':
                        msg= f"Voltage Imbalance greater than {dif}%"
                        if not textOnly.get():
                            messagebox.showerror(parent=alertW, title= f"{name} Meter", message= msg)
                        else:
                            text_update_Table.append("<br>" + str(name) + " | " + str(msg))
                else:
                    meterVstatus= '✓✓✓'
                    meterVstatuscolor= 'green'
                    if meterVconfig not in  ['✓✓✓', 'V']:
                        msg= "Utility Voltage Restored!!! Close the Breaker"
                        if not textOnly.get():
                            messagebox.showinfo(parent=alertW, title=f"{name} Meter", message= msg)
                        else:
                            text_update_Table.append("<br>" + str(name) + " | " + str(msg))

                globals()[f'{var_name}meterVLabel'].config(text= meterVstatus, bg= meterVstatuscolor)

                total_invkW = sum(allinv_kW)
                #Here is where I would check for meterRatio and set the text of the meterRatiolabel
                meterRatio = meterdataWM/metermax
                if meterRatio > .90:
                    ratio_color = '#ADD8E6'  # Light Blue
                elif .90 > meterRatio > .80:
                    ratio_color = '#87CEEB'  # Sky Blue
                elif .80 > meterRatio > .70:
                    ratio_color = '#1E90FF'  # Dodger Blue
                elif .70 > meterRatio > .60:
                    ratio_color = '#4682B4'  # Steel Blue
                elif .60 > meterRatio > .50:
                    ratio_color = '#4169E1'  # Royal Blue
                elif meterRatio < 0.001:
                    ratio_color = 'black'
                else: 
                    ratio_color = 'gray'
                print(f"{name:<15} | {round(meterRatio*100, 1):<5} | {meterdataWM:<9} | {metermax}")

                globals()[f'{var_name}meterRatioLabel'].config(text= f"{round(meterRatio*100, 1)}%", bg= ratio_color, font=('Tk_defaultFont', 10, 'bold'))


                if (meterdataW < 2 or meterdataAA or meterdataAB or meterdataAC) and begin:
                    print(f"Condition Met: {meterdataW}")
                    if name != 'Van Buren': # This if statement and elif pair juke around the VanBuren being down a phase. 
                        meterkWstatus= '❌❌'
                        meterkWstatuscolor= 'red'
                        meterlbl = globals()[f'{var_name}meterkWLabel'].cget('bg')
                        if meterlbl != 'red' and master_cb_skips_INV_check and poa > 10:
                            online = meter_last_online(name)
                            msg= f"Site: {name}\nMeter Production: {round(meterdataW, 2)}\nMeter Amps:\nA: {round(meterdataavgAA, 2)}\nB: {round(meterdataavgAB, 2)}\nC: {round(meterdataavgAC, 2)}\n{online}"
                            if not textOnly.get():
                                messagebox.showerror(parent= alertW, title=f"{name}, Power Loss", message=msg)
                            else:
                                text_update_Table.append("<br>" + str(msg))  
                    elif meterdataAC or meterdataAB or meterdataW < 2: #This is a continuation of the above 'Juke' The Elif below is yet another continuation as Vanburen gets trapped by the very first if statement
                        meterkWstatus= '❌❌'
                        meterkWstatuscolor= 'red'
                        meterlbl = globals()[f'{var_name}meterkWLabel'].cget('bg')
                        if meterlbl != 'red' and master_cb_skips_INV_check and poa > 10:
                            online = meter_last_online(name)
                            msg= f"Site: {name}\nMeter Production: {round(meterdataW, 2)}\nMeter Amps:\nA: {round(meterdataavgAA, 2)}\nB: {round(meterdataavgAB, 2)}\nC: {round(meterdataavgAC, 2)}\n{online}"
                            if not textOnly.get():
                                messagebox.showerror(parent= alertW, title=f"{name}, Meter Power Loss", message=msg)
                            else:
                                text_update_Table.append("<br>" + str(msg))
     
                else:
                    meter_kw = round(meterdataW/1000, 1)
                    meterkWstatus= meter_kw
                    if meter_kw < 1:
                        meterkWstatuscolor= 'red'
                    else:
                        meterkWstatuscolor= 'green'

                #Below we update the GUI with the above defined text and color
                globals()[f'{var_name}meterkWLabel'].config(text= meterkWstatus, bg= meterkWstatuscolor, font=('Tk_defaultFont', 10, 'bold'))
                
                #PVSYST Ratio Update
                try:
                    pysyst_connect()
                    if meterdataWM and poa and pvsyst_name:
                        performance_ratio, degradation, meter_est = pvsyst_est(meterdataWM, poa, pvsyst_name)
                        if pvsyst_name is not None:
                            print(f'{pvsyst_name:<15} | {round(performance_ratio, 1)}% | {round(degradation*100, 2)}% Loss | {round(meter_est, 2):<13} W or kW?')
                        
                        if performance_ratio != 0:
                            if performance_ratio > 90:
                                pvSyst_color = '#ADD8E6'  # Light Blue
                            elif 90 > performance_ratio > 80:
                                pvSyst_color = '#87CEEB'  # Sky Blue
                            elif 80 > performance_ratio > 70:
                                pvSyst_color = '#1E90FF'  # Dodger Blue
                            elif 70 > performance_ratio > 60:
                                pvSyst_color = '#4682B4'  # Steel Blue
                            elif 60 > performance_ratio > 50:
                                pvSyst_color = '#4169E1'  # Royal Blue
                            else: 
                                pvSyst_color = 'gray'
                            globals()[f'{var_name}meterPvSystLabel'].config(text=f'{round(performance_ratio, 1)}%', bg=pvSyst_color, font=('Tk_defaultFont', 10, 'bold'))
                        else:
                            globals()[f'{var_name}meterPvSystLabel'].config(text='N/A')
                    connect_pvsystdb.close()
                except Exception as erro:
                    print("Error: ", erro)
                    pass
                siteSnapShot_Update(name, var_name, inverters, meterdataWM)


            else:
                meterlbl = globals()[f'{var_name}meterkWLabel'].cget('bg')
                if meterlbl != 'pink' and master_cb_skips_INV_check:
                    msg=f"Meter Comms lost {meter_Ltime} with the Meter at {name}! Please Investigate!"
                    if not textOnly.get():
                        messagebox.showerror(parent= alertW, title=f"{name}, Meter Comms Loss", message=msg)
                    else:
                        text_update_Table.append("<br>" + str(msg))
                globals()[f'{var_name}meterkWLabel'].config(bg='pink')
                globals()[f'{var_name}meterVLabel'].config(bg='pink')
                globals()[f'{var_name}meterRatioLabel'].config(bg='pink')



    

    #conetoe_offline()

    if textOnly.get():
        text_update_Table.append("</body></html>")
        print("Table: ", text_update_Table)
        if text_update_Table != [html_start, "</body></html>"]:
            # Connect to Gmail SMTP server and send email
            with smtplib.SMTP("smtp.gmail.com", 587) as server:
                server.starttls()
                server.login(sender, password)
                msg_list = "".join(text_update_Table)
                soup = BeautifulSoup(msg_list, 'html.parser')
                message.attach(MIMEText(soup.prettify(), 'html'))
                print("\n", text_update_Table)
                print("\n", soup.prettify())
                print("\n",message.as_string())
                print("\n", textOnly.get())
                server.send_message(message)
        else:
            print("No New Updates | Email Passed")


    dbconnection.close()


    def allinv_message_update(num, state):
        with open(f"C:\\Users\\OMOPS\\OneDrive - Narenco\\Documents\\APISiteStat\\Site {num} All INV Msg Stat.txt", "w+") as outfile:
                outfile.write(str(state))

    def allinv_message_check(num):
        global text_update_Table
        file_path = f"C:\\Users\\OMOPS\\OneDrive - Narenco\\Documents\\APISiteStat\\Site {num} All INV Msg Stat.txt"
        try:
            with open(file_path, "r+") as rad:
                allinvstat = rad.read()
                return allinvstat
        except FileNotFoundError:
            with open(file_path, "w") as rad:
                 rad.write('1')
            return '1'
        except Exception as errorr:
            msg = f"Error Reading Site {num} txt file"
            if not textOnly.get():
                messagebox.showerror(parent= alertW, message= errorr, title=msg)
            else:
                text_update_Table += f"<br>{msg} | {errorr}"
            return '1'


    # Post Update Status of Inverters
    poststatus_all = {}
    #Retireve Current INV status's and store in lists

    for name, site_dictionary in MAP_SITES_HARDWARE_GUI.items():
        invdict = site_dictionary['INV_DICT']
        var_name = site_dictionary['VAR_NAME']
        custid = site_dictionary['CUST_ID']
        metermax = site_dictionary['METER_MAX']
        pvsyst_name = site_dictionary['PVSYST']

        inverters = len(invdict)
        l = []
        if name != "CDIA":
            for i in range(1, inverters + 1):
                if i in invdict:  # Check if the key exists in the dictionary
                    inv_val = invdict[i]
                else:
                    inv_val = str(i)
                checkbox_var = globals()[f'{var_name}inv{inv_val}cbval']
                if not checkbox_var.get():
                    config_color = globals()[f'{var_name}inv{inv_val}cb'].cget("bg")
                    l.append(config_color)
                else:
                    l.append("green") #Checked inverters are known offline, this logs them in this check as 'online' but thats ok since this process is not the one used to check if the inverter is online or not. If we don't do this the message box reports the wrong index.
            poststatus_all[f'{var_name}'] = l
    #ic(poststatus_all['vanburen'])
    
    for index, (name, site_dictionary) in enumerate(MAP_SITES_HARDWARE_GUI.items()):
        invdict = site_dictionary['INV_DICT']
        var_name = site_dictionary['VAR_NAME']
        custid = site_dictionary['CUST_ID']
        metermax = site_dictionary['METER_MAX']
        pvsyst_name = site_dictionary['PVSYST']
        inverters = len(invdict)
        if name != "CDIA":
            master_cb_skips_INV_checks = True if globals()[f'{var_name}metercbval'].get() == 0 else False
            if poststatus_all[f'{var_name}']:
                if master_cb_skips_INV_checks and all(status in ['red', 'orange', 'pink'] for status in poststatus_all[f'{var_name}']):
                    #Reasons we want to ignore all the Inverters showing offline
                    if allinv_message_check(index + 1) == '1':
                        if ((POA_data[f'{name} POA Data'][0] > 250 and begin) or (h_tm_now >= 10 and POA_data[f'{name} POA Data'][0] > 75)):
                            print(f'{name} Past', status_all[f'{var_name}'])
                            msg= f"All Inverters Offline, Please Investigate!"
                            if not textOnly.get():
                                messagebox.showerror(parent= alertW, title= f"{name}", message=msg)
                            else:
                                text_update_Table.append("<br>" + str(name) + " | " + str(msg))
                            stat = 0
                            allinv_message_update(index + 1, stat)
                else:
                    stat = 1
                    if allinv_message_check(index + 1) == "0":
                        print(f'{name} Trigger Error', poststatus_all[f'{var_name}'])
                        allinv_message_update(index + 1, stat)  



    def compare_lists(site, before, after, inverter_dictionary):
        global text_update_Table
        #Comapres 2 lists of sites inverters to see what remains online
        changed_indices = []  # List to store indices of changed items
        
        # First, identify any change from not "✓" to "✓"
        changes_detected = any(before_item != "green" and after_item == "green" for before_item, after_item in zip(before, after))
        
        if changes_detected:
            # If changes detected, then identify all items that remain "not ✓"
            changed_indices = [i for i, item in enumerate(after) if item != "green"]
        
        if changed_indices:
            # Look up the corresponding values in the inverter dictionary
            offline_inverters = [inverter_dictionary.get(i + 1, f"Index {i + 1}") for i in changed_indices]
            late_starts = ', '.join(offline_inverters)
            msg=f"Some Inverters just came Online. Inverters: {late_starts} remain Offline."
            if not textOnly.get():
                messagebox.showinfo(parent=alertW, title=site, message=msg)
            else:
                text_update_Table += f"<br>{site} | {msg}"

    
    
    #Comapres all lists of sites inverters to see what remains online
    for name, site_dictionary in MAP_SITES_HARDWARE_GUI.items():
        invdict = site_dictionary['INV_DICT']
        var_name = site_dictionary['VAR_NAME']

        inverters = len(invdict)
        if int(globals()[f'{var_name}POAcb'].cget("text")) > 100:
            if name != "CDIA":
                compare_lists(name, status_all[f'{var_name}'], poststatus_all[f'{var_name}'], invdict)

    update_data_finish = ty.perf_counter()
    print("Update Data Time (secs):", round(update_data_finish - update_data_start, 2))
    global gui_update_timer
    target_time = time(8, 30)
    if textOnly.get():
        gui_update_timer = PausableTimer(420, db_to_dict)
        gui_update_timer.start() 
    elif timecurrent.time() < target_time:
        gui_update_timer = PausableTimer(300, db_to_dict)
        gui_update_timer.start()
    else:
        gui_update_timer = PausableTimer(60, db_to_dict)
        gui_update_timer.start()


    sendTexts.config(state=NORMAL)


def pysyst_connect():
    global cursor_p, connect_pvsystdb
    #Connect to PV Syst DB for Performance expectations
    pvsyst_db = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=G:\Shared drives\O&M\NCC Automations\Notification System\PVsyst (Josephs Edits).accdb;'
    connect_pvsystdb = pyodbc.connect(pvsyst_db)
    cursor_p = connect_pvsystdb.cursor()

def pvsyst_est(meterval, poa_val, pvsyst_name):
    if pvsyst_name == None:
        return (0,0,0)
    if poa_val == 9999:
        return (0,0,0)



    
    if pvsyst_name not in ["WELLONS", "FREIGHTLINE", "WARBLER", "PG", "HOLLYSWAMP"]:
        meterval = meterval/1000


    query = "SELECT [GlobInc_WHSQM], [EGrid_KWH] FROM [PVsystHourly] WHERE [PlantName] = ?"
    date_query = "SELECT [SimulationDate] FROM [PVsystHourly] WHERE [PlantName] = ?"

    with warnings.catch_warnings():
        warnings.simplefilter("ignore", category=UserWarning)
        # Execute the query and read into a DataFrame
        slope_df = pd.read_sql_query(query, connect_pvsystdb, params=[pvsyst_name])
    
    meter = 'EGrid_KWH'
    poa = 'GlobInc_WHSQM'
    # Reshape the data for sklearn
    X = slope_df[poa].values.reshape(-1, 1)
    y = slope_df[meter].values
    
    # Create and fit the model
    model = LinearRegression()
    model.fit(X, y)
    
    # Get the slope (coefficient) and intercept
    slope = model.coef_[0]
    intercept = model.intercept_
    
    meter_est = slope * poa_val + intercept

    cursor_p.execute(date_query, pvsyst_name)
    simulation_date = cursor_p.fetchone()[0]
    
    try:
        difference_in_days = (datetime.now() - simulation_date).days
        difference_in_years = difference_in_days / 365.25  # Using 365.25 to account for leap years
        degradation_percentage = difference_in_years * 0.005
        if pvsyst_name == 'WELLONS':
            print('Wellons p50: ', meter_est, ' | ', degradation_percentage)

        meter_estdegrad = meter_est * (1 - degradation_percentage)
        performance = (meterval/meter_estdegrad)*100 #%
        return (performance, degradation_percentage, meter_est)
    except TypeError as e:
        print(e)
        print("Moving On...")
        return (0,0,0)


def checkin():
    try: 
        for widget in checkIns.winfo_children():
            widget.destroy()

        connect_Logbook()

        cur.execute("SELECT Location, Company, Employee FROM [Checked In]")
        checkedIn = cur.fetchall()

        for row_index, row in enumerate(checkedIn):
            for col_index, value in enumerate(row):
            # Check if the value is a datetime object and format it
                if isinstance(value, datetime):
                    value = value.strftime('%m/%d/%y')
                # Apply different formatting for specific columns
                if row_index in range(1, 100, 2):
                    bg_color = '#90EE90'  # Pale Light Green
                else:
                    bg_color = '#ADD8E6'  # Light Blue of Main Site Data Window
                if col_index == 2:
                    wsize = 24
                elif col_index == 1:
                    wsize = 32
                else:
                    wsize = 23
                label = Label(checkIns, text=value, font=("Calibri", 14), borderwidth=1, relief="solid", width=wsize, bg=bg_color)
                label.grid(row=row_index, column=col_index)

    except pyobdc.Error as err:
        print("Logbook Error: ", err)

    lbconnection.close()
    update_data()


def last_update():
    times = []
    for name in MAP_SITES_HARDWARE_GUI:
        if name != "CDIA":
            c.execute(f"SELECT TOP 1 [Timestamp] FROM [{name} Meter Data] ORDER BY [Timestamp] DESC")
            last_time = c.fetchone()
            times.append(last_time[0])
    most_recent = max(times)
    return most_recent







def time_window():

    global timecurrent, text_update_Table
    #SELECT 15 = 30 Mins
    c.execute("SELECT TOP 16 [Timestamp] FROM [Ogburn Meter Data] ORDER BY [Timestamp] DESC")
    data_timestamps = c.fetchall()
    firsttime = data_timestamps[0][0]
    secondtime = data_timestamps[1][0]
    thirdtime = data_timestamps[2][0]
    fourthtime = data_timestamps[3][0]
    fifthtime = data_timestamps[4][0]
    tenthtime = data_timestamps[9][0]
    lasttime = data_timestamps[14][0]

    hm_firsttime = firsttime.strftime('%H:%M')
    hm_secondtime = secondtime.strftime('%H:%M')
    hm_thirdtime = thirdtime.strftime('%H:%M')
    hm_fourthtime = fourthtime.strftime('%H:%M')
    hm_tenthtime = tenthtime.strftime('%H:%M')
    hm_lasttime = lasttime.strftime('%H:%M')

    time1v.config(text=hm_firsttime, font=("Calibri", 16))
    time2v.config(text=hm_secondtime, font=("Calibri", 16))
    time3v.config(text=hm_thirdtime, font=("Calibri", 16))
    time4v.config(text=hm_fourthtime, font=("Calibri", 16))
    time10v.config(text=hm_tenthtime, font=("Calibri", 16))
    timeLv.config(text=hm_lasttime, font=("Calibri", 16))

    pulls5TD = firsttime - fifthtime
    pulls5TDmins = round(pulls5TD.total_seconds() / 60, 2)
    pulls15TD = firsttime - lasttime
    pulls15TDmins = round(pulls15TD.total_seconds() / 60, 2)
    spread10.config(text=f"5 Pulls\n{pulls5TDmins} Minutes")
    spread15.config(text=f"15 Pulls\n{pulls15TDmins} Minutes")
    
    timecurrent = datetime.now()
    db_update_time = 10
    timecompare = timecurrent - timedelta(minutes=db_update_time)
    recent_update = last_update()
    if recent_update < timecompare and comms_delay.get() == False:
        os.startfile(r"G:\Shared drives\O&M\NCC Automations\Notification System\API Data Pull, Multi SQL.py")
        msg = f"The Database has not been updated in {str(db_update_time)} Minutes and usually updates every 2\nLaunching Data Pull Script in response."
        if not textOnly.get():
            messagebox.showerror(parent=timeW, title="Notification System/GUI", message=msg)
        else:
            try:
                text_update_Table.append("<br>" + str(msg))
            except UnboundLocalError:
                pass
            except NameError:
                pass
        ty.sleep(180)

    tupdate = timecurrent.strftime('%H:%M')

    timmytimeLabel.config(text= tupdate, font= ("Calibiri", 30))
    
    checkin() 

def db_to_dict():
    day_of_week = datetime.today().weekday()
    if day_of_week > 4:
        now = datetime.now()
        if (now.hour == 14 and now.minute >= 55) or (now.hour > 14):
            os.startfile(r"G:\Shared drives\O&M\NCC Automations\Notification System\Email Notification (Breaker).py")
            ty.sleep(5)
            os.startfile(r"G:\Shared drives\O&M\NCC Automations\Emails\Close AE GUI.ahk")
    else:
        now = datetime.now()
        if now.hour >= 20:
            restart_pc()

    query_start = ty.perf_counter()
    sendTexts.config(state=DISABLED)


    connect_db()
    global tables, inv_data, breaker_data, meter_data, comm_data, POA_data, begin
    tables = []
    for tb in c.tables(tableType='TABLE'):
        if 'Data' in tb.table_name:
            tables.append(tb)
    #ic(tables)
    excluded_tables = ["1)Sites", "2)Breakers", "3)Meters", "4)Inverters", "5)POA"]

    tb_file = r"C:\Users\omops\Documents\Automations\Troubleshooting.txt"
    comm_data = {}
    for table in tables:
        table_name = table.table_name
        if table_name not in excluded_tables:
            c.execute(f"SELECT TOP 10 [Last Upload] FROM [{table_name}] ORDER BY Timestamp DESC")
            comm_value = c.fetchall()
            comm_data[table_name] = comm_value

    #ic(comm_data)
    inv_data = {}
    for table in tables:
        table_name = table.table_name
        if "INV" in table_name and table_name not in excluded_tables:
            #SELECT 15 = 30 Mins
            c.execute(f"SELECT TOP 16 [dc V], Watts FROM [{table_name}] ORDER BY Timestamp DESC")
            inv_rows = c.fetchall()
            #ic(inv_rows)
            inv_data[table_name] = inv_rows
    #ic(inv_data)

    meter_data = {}
    for table in tables:
        table_name = table.table_name
        if any(name in table_name for name in ["Hickory", "Whitehall"]) and "Meter" in table_name:
            #SELECT 13 = 17 Mins of Data
            c.execute(f"SELECT TOP 16 [Volts A], [Volts B], [Volts C], [Amps A], [Amps B], [Amps C], Watts FROM [{table_name}] ORDER BY Timestamp DESC")
            meter_rows = c.fetchall()
            meter_data[table_name] = meter_rows 
        elif any(name in table_name for name in ["Wellons",]) and "Meter" in table_name:
            #SELECT 45 = 60 Mins of Data | Wellons has a severe intermittent comms issue.
            c.execute(f"SELECT TOP 60 [Volts A], [Volts B], [Volts C], [Amps A], [Amps B], [Amps C], Watts FROM [{table_name}] ORDER BY Timestamp DESC")
            meter_rows = c.fetchall()
            meter_data[table_name] = meter_rows 
        elif "Meter" in table_name and table_name not in excluded_tables:
            #SELECT 5 = 5.5 Mins of Data
            c.execute(f"SELECT TOP {meter_pulls} [Volts A], [Volts B], [Volts C], [Amps A], [Amps B], [Amps C], Watts FROM [{table_name}] ORDER BY Timestamp DESC")
            meter_rows = c.fetchall()
            meter_data[table_name] = meter_rows

    #ic(meter_data)
    POA_data = {}
    for table in tables:
        table_name = table.table_name
        if "POA" in table_name and table_name not in excluded_tables:
            c.execute(f"SELECT TOP 1 [W/M²] FROM [{table_name}] ORDER BY [Timestamp] DESC")
            POA_rows = c.fetchone()
            POA_data[table_name] = POA_rows
    

    #ic(POA_data)

    breaker_data = {}
    for table in tables:
        table_name = table.table_name
        if "Breaker" in table_name and table_name not in excluded_tables:
            c.execute(f"SELECT TOP {breaker_pulls} [Status] FROM [{table_name}] ORDER BY [Timestamp] DESC")
            breaker_rows = c.fetchall()
            breaker_data[table_name] = breaker_rows
    #ic(breaker_data)

    begin = launch_check()

    query_end = ty.perf_counter()
    print("Query Time (secs):", round(query_end - query_start, 2))
    time_window()



def parse_wo():
    directory = "G:\\Shared drives\\O&M\\NCC Automations\\Notification System\\WO Tracking\\"
    existing_wo_files = glob.glob(os.path.join(directory, "*.txt"))
    for file in existing_wo_files:
        try:
            os.remove(file)
        except Exception as e:
            print(f"Error: {e}")

    open_wo_file = filedialog.askopenfilename(parent=alertW,
        title="Select a file", 
        filetypes=[("Excel Files", "*.xlsx *.xls")],
        initialdir="C:\\Users\\OMOPS\\Downloads")
    wo_data = pd.read_excel(open_wo_file)
    for index, row in wo_data.iterrows():
        var_adjustment = [", LLC", "Farm", "Cadle", "Solar"]
        site = row['Site']
        if pd.isna(site) or site in  {"Charter GM", "Charter RM", "Charter Roof"}:
            continue
        var_name = str(site)
        for phrase in var_adjustment:
            var_name = var_name.replace(phrase, "")
        if 'Wayne' in site:
            var_name = var_name.replace("III", "3")
            var_name = var_name.replace("II", "2")
            var_name = var_name.replace("I", "1")
        var_name = var_name.replace(" ", "").lower()
        var_name = var_name.replace("freightline", "freightliner")
        if site == "BISHOPVILLE":
            var_name = 'bishopvilleII'
        
        status = row['Job Status']
        error_type = row['Fault Code Category']
        device_type = row['Asset Description']
        wo_summary = row['Brief Description']
        notes = row['Work Description']
        wo_date = row['WO Date']
        wo_num = row['WO No.']

        if error_type == 'No Issue':
            continue

        #Inverter check
        if 'inv' in device_type.lower() or 'inv' in wo_summary.lower():
            inv_pattern = r"(?:inverter|inv)\s*(\d+)?(?:-|\.)?(\d+)?"
            matchs = re.search(inv_pattern, wo_summary.lower())
            #Reset Variables
            group = None
            num = None
            if matchs is not None:
                group = int(matchs.group(1)) if matchs.group(1) is not None else None
                num = int(matchs.group(2)) if matchs.group(2) is not None else group
            print(f"{site}  |  {group} | {num}")
            
            #Inverter # not Found
            if group == None or num == None:
                print(f"{site} | G:{group} | Num:{num} WO Parse")
                continue

            inv_num = define_inv_num(site, group, num)
            
            if inv_num is None:
                print(f"Num: {inv_num} | {site} | {wo_num} | {wo_summary}\n")
                # Construct the file path for the text file
                txt_file_path = os.path.join(directory, f"{var_name} Open WO's.txt")
                # Append the row data to the text file
                with open(txt_file_path, 'a+') as file:
                    file.write(f'{inv_num} |  WO: {wo_num:<8}|  {wo_date}  |  {wo_summary}\n')
            else:            
                #Color Assignment Logic
                current_colorstatus = globals()[f'{var_name}inv{inv_num}WOLabel'].cget('text')
                if current_colorstatus == 'gray':
                    continue
                
                if error_type == 'Underperformance':
                    if current_colorstatus == 'black':
                        color = 'gray'
                    else:
                        color = 'blue' 
                elif error_type == 'Equipment Outage':
                    if current_colorstatus == 'blue':
                        color = 'gray'
                    else:
                        color = 'black'
                elif error_type == 'COMMs Outage':
                    color = 'pink'
                else: 
                    color = 'yellow'
                globals()[f'{var_name}inv{inv_num}WOLabel'].config(bg=color)

                # Construct the file path for the text file
                txt_file_path = os.path.join(directory, f"{var_name} Open WO's.txt")
                # Append the row data to the text file
                with open(txt_file_path, 'a+') as file:
                    file.write(f'{inv_num:<5}|  WO: {wo_num:<8}|  {wo_date}  |  {wo_summary}\n')

STATE_FILE = r"G:\Shared drives\O&M\NCC Automations\Notification System\CheckBoxState.json"

def save_cb_state():
    state = [var.get() for var in all_CBs]
    with open(STATE_FILE, 'w') as f:
        json.dump(state, f)

def load_cb_state():
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, 'r') as f:
            try:
                state = json.load(f)
                for var, value in zip(all_CBs, state):
                    var.set(value)
            except json.decoder.JSONDecodeError as e:
                print(f"Error decoding JSON data: {e}")
                # Handle the error or provide appropriate fallback behavior
    else:
        print(f"File {STATE_FILE} does not exist.")
def check_button_notes():
    messagebox.showinfo(parent=alertW, title="Checkbutton Info", message= """The First column of CheckButtons in the Site Data Window turns off all notifications associated with that Site.
                        \nThe POA CB will change the value to 9999 so that no inv outages are filtered by the POA
                        \nThe colored INV CheckButtons are to be selected when a WO is open for that device and will turn off notifications of outages with INV
                        \nThe Box in the middle Represents the Status of that device in Emaint. | ⬜ = NO WO | Black BG = Offline WO Open | Blue BG = Underperformance WO Open | Pink BG = Comms Outage WO Open | Yellow BG = Unknown WO Found |
                        \nThe 3rd Column is a CB for Underperformance tracking. Data range is set by the user but as standard 30 days of data and only between the Hours of 10:00 to 15:00.
                        \nThe first value is calculated by averaging all the values in the data range, then comparing the results to the others in the group.
                        \nThe second value is a total of all the values in the data range, then comparing the results to the others in the group.""")

    
    return True
def open_file():
    os.startfile(r"G:\Shared drives\Narenco Projects\O&M Projects\NCC\Procedures\NCC Tools - Joseph\Also Energy GUI Interactions - How To.docx")
    

cur_time = datetime.now()
tupdate =  cur_time.strftime('%H:%M')
notesFrame = Frame(alertW)
notesFrame.grid(row=0, column=0, sticky=EW)
alertwnotes = Label(notesFrame, text= "1st Checkbox: ✓ = Open WO\n& pauses inv notifications", font= ("Calibiri", 12))
alertwnotes.pack()
tupdateLabel = Label(notesFrame, text= "GUI Last Updated", font= ("Calibiri", 18))
tupdateLabel.pack()
timmytimeLabel = Label(notesFrame, text= tupdate, font= ("Calibiri", 30))
timmytimeLabel.pack()
notes_button = Button(notesFrame, command= lambda: check_button_notes(), text= "Checkbutton Notes", font=("Calibiri", 14), bg=main_color, cursor='hand2')
notes_button.pack(padx= 2, pady= 2, fill=X)
proc_button = Button(notesFrame, command= lambda: open_file(), text= "Procedure Doc", font=("Calibiri", 14), cursor='hand2')
proc_button.pack(padx= 2, pady= 2, fill=X)
wo_button = Button(notesFrame, command= lambda: parse_wo(), text= "Assess Open WO's", font=("Calibiri", 14), cursor='hand2')
wo_button.pack(padx= 2, pady= 2, fill=X)


notificationFrame = Frame(alertW)
notificationFrame.grid(row=0, column=1, sticky=N)
notificationNotes = Label(notificationFrame, text="Notification Settings", font=("Calibiri", 14))
notificationNotes.pack()
textOnly = IntVar()
sendTexts = Checkbutton(notificationFrame, text="Send Emails\n(Disable Local MsgBox's)", cursor='hand2', variable=textOnly)
sendTexts.pack(padx=2)
adminTexts = StringVar()
optionTexts = ttk.Combobox(notificationFrame, textvariable=adminTexts, values=["Joseph Lang", "Brandon Arrowood", "Jacob Budd", "Administrators + NCC", "Administrators Only"], state="readonly")
optionTexts.pack()
optionTexts.current(0)
notes_settings = Label(notificationFrame, text="\nSelect from the Dropdown\nBefore turning the function on\nwith the Checkbox\n")
notes_settings.pack()

comms_delay = BooleanVar()
comms_AE_delay = Checkbutton(notificationFrame, text="Select to pause\nrestart of Data Pull Script\nDo so if AE is in a\nmajor comms delay", variable= comms_delay, cursor='hand2')
comms_AE_delay.pack()



root.after(10, load_cb_state)
launch_end = ty.perf_counter()
print("Launch Time (secs):", launch_end - start)

#Start Update Cycle
root.after(10, db_to_dict)
root.mainloop()


#At Exit Tasks
def allinv_message_reset():
    for num in range(1, 38):
        with open(f"C:\\Users\\OMOPS\\OneDrive - Narenco\\Documents\\APISiteStat\\Site {num} All INV Msg Stat.txt", "w+") as outfile:
            outfile.write("1")



atexit.register(allinv_message_reset)
atexit.register(save_cb_state)
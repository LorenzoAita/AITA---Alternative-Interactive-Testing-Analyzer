# import sys
# from PyQt5.QtCore import *
# from PyQt5.QtWebEngineWidgets import *
# from PyQt5.QtWidgets import QApplication, QDialog
#
# app = QApplication(sys.argv)
#
# web = QWebEngineView()
# web.load(QUrl("S:/@Solar/Reliability Laboratory/3_Reliability Data Test Folder NEW/3_PVS-100(120)-Wave2/19_Letture NTC Bobine/graph/SN109294_2201C.html"))
# web.resize(1800, 1000)
# web.setWindowTitle("AITA - Live Graph")
# web.show()
#
# sys.exit(app.exec_())

# def popupBrowser():
#     w = QDialog()
#     w.resize(600, 500)
#     w.setWindowTitle('Wiki')
#     web = QWebEngineView(w)
#     web.load(QUrl('https://github.com/aerospaceresearch/visma/wiki'))
#     web.resize(600, 500)
#     web.show()
#     w.show()
#
# popupBrowser()


import requests
import pyvisa
import plotly.graph_objects as go
import pandas as pd
import struct
import time
import serial
from pymodbus.client.sync import ModbusSerialClient as ModbusClient  # initialize a serial RTU client instance
from datetime import datetime
from lib.Libs import *

sorensen = 'TCPIP0::169.254.135.192::inst0::INSTR'
lambda_port = 'GPIB9::5::INSTR'

session = requests.Session()
rm = pyvisa.ResourceManager()
# alim_open = rm.open_resource(sorensen)
# alim = Sorrensen(alim_open)
# time.sleep(1)
# alim.power(10)
# time.sleep(5)
# alim.voltage(10)
# time.sleep(5)
# alim.current(1)
# time.sleep(5)
# alim.stato(1)
# time.sleep(5)
# alim.stato(0)

alim_open = rm.open_resource(sorensen)
alim = Sorrensen(alim_open)
time.sleep(1)
alim.power(10)
time.sleep(1)
alim.voltage(10)
time.sleep(1)
alim.current(1)
time.sleep(1)
alim.stato(1)
time.sleep(1)
alim.stato(0)


import datetime
import os
import time
import numpy as np
import pandas as pd
import pyvisa
from subprocess import Popen
from lib.Config import *
from lib.Libs import *
import warnings
warnings.filterwarnings('ignore')

# import plotly.graph_objects as go
# import serial
# import socket
# import keyboard
session = requests.Session()
rm = pyvisa.ResourceManager()

# variabili di sistema
namefile = 'Data.csv'
time_sample = 10
test_time = 3600*24
path_config = 'Config/'
device = list(pd.read_excel(path_config + 'Config.xlsx', sheet_name='Strumenti')['ELENCO STRUMENTI'])
path_save = ''  # '//atp.fimer.com/ATP_ONLINE/LOG/64/'
# ORION
if 'inverter' in device:
    id_device, ip_device, telemetry = config_inverter(path_config)

# DataLogger
if 'logger' in device:
    config_DAQ = pd.read_excel(path_config + 'Config.xlsx', sheet_name='Agilent')

# Wattmetro
if 'wattmeter' in device:
    config_WATT = pd.read_excel(path_config + 'Config.xlsx', sheet_name='Wattmeter')

# Grafico
time_refresh = 100
plot = [
    'Contatto Connettore Interno',
    'Cavo Connettore Interno',
    'Contatto Connettore Esterno',
    'Cavo Splitter',
    'Tamb Interna',
    'Tamb Esterna'


]
dati_stamp = pd.DataFrame()

print('>>> Start Config')
print('>>>')
if os.path.exists(path_save + namefile):
    os.rename(path_save + namefile, path_save + 'Data_'+str(datetime.datetime.now().strftime("%Y_%m_%d__%H_%M")[:])+'.csv')
col = list()
col.append('Data')
if 'inverter' in device:
    obj_com = OrionProtocol(ip=ip_device, device_id=id_device)
    for i in telemetry:
        col.append(i)
if 'logger' in device:
    logger_DAQ, porta_DAQ = config_agilent(rm, path_config)
    logger_DAQ = logger_DAQ.fillna('')
    stamp_daq = list()
    for i in logger_DAQ:
        if i != '':
            col.append(i)
            stamp_daq.append(i)
if 'wattmeter' in device:
    logger_WT, porta_WT = config_wt(rm, path_config, config_WATT['MODELLO'][0])
    logger_WT = logger_WT.fillna('')
    stamp_WT = list()
    for i in logger_WT:
        if i != '':
            col.append(i)
            stamp_WT.append(i)

tot_time = 0
step = 0
sample = 0
print('>>> Start Log')
print('>>>')
while tot_time < test_time:
    sample += 1
    dati = pd.DataFrame()
    telemetries = list()
    telemetries.append(datetime.datetime.now())
    time.sleep(time_sample)
    print(
        '>>> Tempo\t' + str(datetime.datetime.now().strftime("%Y/%m/%d, %H:%M:%S.%f")[:-2]) + '\tStep #' + str(sample))
    if 'inverter' in device:
        print('>>> log inverter\t' + str(ip_device))
        for i in telemetry:
            #  print('>>> log la telemetry\t' + str(i))
            try:
                data_eut, status_code = obj_com.get_data(url=i)
                telemetries.append(data_eut)
            except:
                for j in range(0, len(telemetry)-len(telemetries)+1):
                    telemetries.append(0)
                break
                #  print('>>> la telemetry\t' + str(i) + '\tnon risponde')
    if 'logger' in device:
        i = 0
        data_daq = list()
        print('>>> Apro la Comunicazione con il DAQ\t')
        INSTRUMENT = rm.open_resource(porta_DAQ)
        INSTRUMENT.write('INIT')
        time.sleep(0.05)
        a = INSTRUMENT.query('FETCH?').split(',')
        for j in a:
            if '+' in j or '-' in j:
                telemetries.append(float(j[0:15]))
                print('>>> log la telemetry\t' + str(stamp_daq[i]))
                i += 1
    if 'wattmeter' in device:
        i = 0
        data_WT = list()
        print('>>> Apro la Comunicazione con il ' + config_WATT['MODELLO'][0])
        INSTRUMENT = rm.open_resource(porta_WT)
        if config_WATT['MODELLO'][0] == 'WT230':
            a = INSTRUMENT.query('MEAS:NORM:VAL?').split(',')
        if config_WATT['MODELLO'][0] == 'WT500' or config_WATT['MODELLO'][0] == 'WT3000':
            a = INSTRUMENT.query(':NUM:NORM:VAL?').split(',')
        for j in a:
            if '+' in j or '-' in j:
                telemetries.append(float(j[0:15]))
                print('>>> log la telemetry\t' + str(stamp_WT[i]))
                i += 1

    dati_T = pd.DataFrame()
    dati_T = dati_T.append(telemetries)
    dati = pd.concat([dati, dati_T.T], axis=0, ignore_index=True)
    for i in range(0, len(col)):
        dati.rename(columns={i: col[i]}, inplace=True)
    dati_stamp = pd.concat([dati_stamp, dati], axis=0, ignore_index=True)
    if sample == 1:
        dati.to_csv(path_save + namefile, sep=',', index=False)
    else:
        dati.to_csv(path_save + namefile, sep=',', mode='a', index=False, header=False)
        if tot_time >= step * time_refresh:
            step += 1
            try:
                Popen('taskkill /F /IM chrome.exe', shell=True)
                plot_runtime(step, dati_stamp, plot)
            except:
                plot_runtime(step, dati_stamp, plot)
    tot_time += time_sample
    print('>>>')

print('>>> End')
a = 0

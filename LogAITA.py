import datetime
import os
import time
from lib.Config import *
from lib.Libs import *
import warnings
import pyvisa
import keyboard
import pandas as pd
warnings.filterwarnings('ignore')
session = requests.Session()
rm = pyvisa.ResourceManager()

# variabili di sistema
def main_log():
    path_config = 'Config/'
    device = list(pd.read_excel(path_config + 'Config.xlsx', sheet_name='Strumenti')['ELENCO STRUMENTI'])
    time_sample = pd.read_excel(path_config + 'Config.xlsx', sheet_name='Strumenti')['TEMPO CAMPIONAMENTO'][0]
    test_time = pd.read_excel(path_config + 'Config.xlsx', sheet_name='Strumenti')['TEST TIME'][0]
    namefile = pd.read_excel(path_config + 'Config.xlsx', sheet_name='Strumenti')['NOME OUTPUT'][0]
    if str(namefile) == 'nan':
        namefile = 'Data'
    if str(pd.read_excel(path_config + 'Config.xlsx', sheet_name='Strumenti')['PERCORSO OUTPUT'][0]) == 'nan':
        path_save = ''
    else:
        path_save = pd.read_excel(path_config + 'Config.xlsx', sheet_name='Strumenti')['PERCORSO OUTPUT'][
            0]  # '//atp.fimer.com/ATP_ONLINE/LOG/64/'
        if not os.path.exists(path_save):
            os.makedirs(path_save)
    # ORION
    if 'Inverter' in device:
        id_device, ip_device, telemetry_inv = config_inverter(path_config)

    # Colonnina
    if 'Colonnina' in device:
        telemetry_col, reg, com_colonna, addresses = config_colonnina(path_config)

    # DataLogger
    if 'Agilent' in device:
        config_DAQ = pd.read_excel(path_config + 'Config.xlsx', sheet_name='Agilent')

    # Wattmetro
    if 'Wattmeter' in device:
        config_WATT = pd.read_excel(path_config + 'Config.xlsx', sheet_name='Wattmeter')
    # Wattmetro2
    if 'Wattmeter2' in device:
        config_WATT2 = pd.read_excel(path_config + 'Config.xlsx', sheet_name='Wattmeter2')
    # if 'Grafico' in device:
    #     # Grafico
    #     time_refresh = pd.read_excel(path_config + 'Config.xlsx', sheet_name='Grafico')['REFRESH TIME'][0]
    #     plot = list()
    #     for i in range(0, len(pd.read_excel(path_config + 'Config.xlsx', sheet_name='Grafico')['PLOT'])):
    #         if str(pd.read_excel(path_config + 'Config.xlsx', sheet_name='Grafico')['PLOT'][i]) != 'nan' and str(
    #                 pd.read_excel(path_config + 'Config.xlsx', sheet_name='Grafico')['ASSE'][i]) != 'nan':
    #             plot.append([pd.read_excel(path_config + 'Config.xlsx', sheet_name='Grafico')['PLOT'][i],
    #                          pd.read_excel(path_config + 'Config.xlsx', sheet_name='Grafico')['ASSE'][i]])

    dati_stamp = pd.DataFrame()
    print('>>> Start Config')
    print('>>>')
    if os.path.exists(path_save + namefile + '.csv') and 'Bridge' not in device:
        os.rename(path_save + namefile + '.csv',
                  path_save + 'Data_' + str(datetime.datetime.now().strftime("%Y_%m_%d__%H_%M")[:]) + '.csv')
    col = list()
    col.append('Date')
    if 'Inverter' in device:
        obj_com = OrionProtocol(ip=ip_device, device_id=id_device)
        for i in telemetry_inv:
            col.append(i)
    if 'Colonnina' in device:
        for i in telemetry_col:
            col.append(i)
    if 'Agilent' in device:
        logger_DAQ, porta_DAQ = config_agilent(rm, path_config)
        logger_DAQ = logger_DAQ.fillna('')
        stamp_daq = list()
        for i in logger_DAQ:
            if i != '':
                col.append(i)
                stamp_daq.append(i)
    if 'Wattmeter' in device:
        logger_WT, porta_WT = config_wt(rm, path_config, config_WATT['MODELLO'][0])
        logger_WT = logger_WT.fillna('')
        stamp_WT = list()
        for i in logger_WT:
            if i != '':
                col.append(i)
                stamp_WT.append(i)
    if 'Wattmeter2' in device:
        logger_WT2, porta_WT2 = config_wt(rm, path_config, config_WATT['MODELLO'][0])
        logger_WT2 = logger_WT2.fillna('')
        stamp_WT2 = list()
        for i in logger_WT2:
            if i != '':
                col.append(i)
                stamp_WT2.append(i)
    j = 0
    if 'Bridge' in device:
        logger_bridge, porta_bridge = config_ponte(rm, path_config)
        logger_bridge = logger_bridge.fillna('')
        stamp_bridge = list()
        stamp_bridge.append('Frequenza')
        for i in logger_bridge:
            if i != '':
                if i in ['RX', 'GB']:
                    name1 = i[0:1]
                    name2 = i[1:]  # name1 + '-' + i[2:]
                else:
                    name1 = i[0:2]
                    name2 = i[2:]  # name1 + '-' + i[2:]
                if name1 in stamp_bridge:
                    name1 += '_' + str(j)
                    j += 1
                if name2 in stamp_bridge:
                    name2 += '_' + str(j)
                    j += 1
                # col.append(name1)
                # col.append(name2)
                stamp_bridge.append(name1)
                stamp_bridge.append(name2)

    tot_time = 0
    step = 0
    sample = 0
    print('>>> Start Log')
    print('>>>')
    while tot_time < test_time:
        time.sleep(time_sample)
        sample += 1
        dati = pd.DataFrame()
        telemetries = list()
        telemetries.append(datetime.datetime.now())
        print(
            '>>> Tempo\t' + str(datetime.datetime.now().strftime("%Y/%m/%d, %H:%M:%S.%f")[:-2]) + '\tStep #' + str(sample))
        if 'Inverter' in device:
            print('>>> log inverter\t' + str(ip_device))
            for i in telemetry_inv:
                #  print('>>> log la telemetry\t' + str(i))
                try:
                    data_eut, status_code = obj_com.get_data(url=i)
                    telemetries.append(data_eut)
                except:
                    for j in range(0, len(telemetry_inv) - len(telemetries) + 1):
                        telemetries.append(0)
                    break
                    #  print('>>> la telemetry\t' + str(i) + '\tnon risponde')
        if 'Colonnina' in device:
            print('>>> log colonnina\t' + str(com_colonna))
            for i in range(0, len(telemetry_col)):
                # print('>>> log la telemetry\t' + str(telemetry_col[i]) + '\tal registro\t' + str(reg[i]))
                data_col = ReadCol(reg[i], com_colonna, addresses[i])
                telemetries.append(data_col)
        if 'Agilent' in device:
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
        if 'Wattmeter' in device:
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
        if 'Wattmeter2' in device:
            i = 0
            data_WT2 = list()
            print('>>> Apro la Comunicazione con il ' + config_WATT2['MODELLO'][0])
            INSTRUMENT = rm.open_resource(porta_WT2)
            if config_WATT2['MODELLO'][0] == 'WT230':
                a = INSTRUMENT.query('MEAS:NORM:VAL?').split(',')
            if config_WATT2['MODELLO'][0] == 'WT500' or config_WATT2['MODELLO'][0] == 'WT3000':
                a = INSTRUMENT.query(':NUM:NORM:VAL?').split(',')
            for j in a:
                if '+' in j or '-' in j:
                    telemetries.append(float(j[0:15]))
                    print('>>> log la telemetry\t' + str(stamp_WT2[i]))
                    i += 1
        if 'Bridge' in device:
            i = 0
            data_bridge = pd.DataFrame()
            print('>>> Apro la Comunicazione con il ponte\t')
            INSTRUMENT = rm.open_resource(porta_bridge)
            meas_bridge(INSTRUMENT, stamp_bridge, data_bridge, path_config, path_save)

        if 'Bridge' not in device:
            dati_T = pd.DataFrame()
            dati_T = dati_T.append(telemetries)
            dati = pd.concat([dati, dati_T.T], axis=0, ignore_index=True)
            for i in range(0, len(col)):
                dati.rename(columns={i: col[i]}, inplace=True)
            dati_stamp = pd.concat([dati_stamp, dati], axis=0, ignore_index=True)

            if sample == 1:
                dati.to_csv(path_save + namefile + '.csv', sep=',', index=False)
                #os.system('python VisualDati.py')
            else:
                dati.to_csv(path_save + namefile + '.csv', sep=',', mode='a', index=False, header=False)
                #os.system('python VisualDati.py')
                # if 'Grafico' in device:
                #     if tot_time >= step * time_refresh and plot != []:
                #         step += 1
                #         try:
                #             # Popen('taskkill /F /IM chrome.exe', shell=True)
                #             keyboard.press_and_release('ctrl+w')
                #             plot_runtime(step, dati_stamp, plot)
                #         except:
                #             plot_runtime(step, dati_stamp, plot)

        tot_time += time_sample
        print('>>>')

    print('>>> End')

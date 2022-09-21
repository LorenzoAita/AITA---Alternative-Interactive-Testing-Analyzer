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
    try:
        path_config = 'Config/'
        device = list(pd.read_excel(path_config + 'Config.xlsx', sheet_name='Strumenti')['ELENCO STRUMENTI'])
        time_sample = pd.read_excel(path_config + 'Config.xlsx', sheet_name='Strumenti')['TEMPO CAMPIONAMENTO'][0]
        test_time = pd.read_excel(path_config + 'Config.xlsx', sheet_name='Strumenti')['TEST TIME'][0]
        namefile = pd.read_excel(path_config + 'Config.xlsx', sheet_name='Strumenti')['NOME OUTPUT'][0]
        if str(namefile) == 'nan':
            namefile = 'Data'
        if str(pd.read_excel(path_config + 'Config.xlsx', sheet_name='Strumenti')['PERCORSO OUTPUT'][0]) == 'nan':
            path_save = './Data/'
        else:
            path_save = pd.read_excel(path_config + 'Config.xlsx', sheet_name='Strumenti')['PERCORSO OUTPUT'][0]
            if not os.path.exists(path_save):
                os.makedirs(path_save)
        # ORION
        if 'Inverter' in device:
            ids_device, ips_device, telemetrys_inv, labels_inv = config_inverter(path_config)

        # Colonnina
        if 'Colonnina' in device:
            print('>>> Config AC Station')
            telemetrys_col, regs, coms_colonna, addresses = config_colonnina(path_config)

        # DataLogger
        if 'Datalogger' in device:
            config_DAQ = pd.read_excel(path_config + 'Config.xlsx', sheet_name='Agilent')

        # Wattmetro
        if 'Wattmeter' in device:
            config_WATT = pd.read_excel(path_config + 'Config.xlsx', sheet_name='Wattmeter')

        dati_stamp = pd.DataFrame()
        print('>>> Start Config')
        print('>>>')
        if os.path.exists(path_save + namefile + '.csv') and 'Bridge' not in device:
            os.rename(path_save + namefile + '.csv',
                      path_save + 'Data_' + str(datetime.now().strftime("%Y_%m_%d__%H_%M_%S")) + '.csv')
        col = list()
        col.append('Date')
        if 'Inverter' in device:
            obj_coms = list()
            for i in range(0, len(ips_device)):
                obj_com = OrionProtocol(ip=ips_device[i], device_id=ids_device[i])
                obj_coms.append(obj_com)
            for i in range(0, len(labels_inv)):
                for j in labels_inv[i]:
                    col.append('Inv'+str(i+1)+'_'+j)
        if 'Colonnina' in device:
            for i in range(0, len(telemetrys_col)):
                for j in telemetrys_col[i]:
                    col.append('AC Station'+str(i+1)+'_'+j)
        if 'Datalogger' in device:
            logger_DAQs, porta_DAQs = config_agilent(rm, path_config)
            stamp_daq = list()
            for i in logger_DAQs:
                for j in i:
                    if j != '':
                        col.append(j)
                        stamp_daq.append(j)
        if 'Wattmeter' in device:
            logger_WTs, porta_WTs = config_wt(rm, path_config)
            stamp_WT = list()
            for i in logger_WTs:
                for j in i:
                    if j != '':
                        col.append(j)
                        stamp_WT.append(j)
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
            time.sleep(float(time_sample))
            sample += 1
            dati = pd.DataFrame()
            telemetries = list()
            telemetries.append(datetime.now())
            print(
                '>>> Tempo\t' + str(datetime.now().strftime("%Y/%m/%d, %H:%M:%S.%f")[:-2]) + '\tStep #' + str(sample))
            if 'Inverter' in device:
                for k in range(0, len(obj_coms)):
                    print('>>> log inverter\t', ids_device[k])
                    for i in telemetrys_inv[k]:
                        try:
                            data_eut, status_code = obj_coms[k].get_data(url=i)
                            telemetries.append(data_eut)
                        except:
                            for j in range(0, len(telemetries) - len(telemetrys_inv[k]) + 1):  # forse è al contrario la differenza
                                telemetries.append(0)
                            break
            if 'Colonnina' in device:
                for k in range(0, len(coms_colonna)):
                    print('>>> log colonnina\t' + str(coms_colonna[k]))
                    for i in range(0, len(telemetrys_col[k])):
                        try:
                            data_col = ReadCol(regs[k][i], coms_colonna[k], addresses[k][i])
                            telemetries.append(data_col)
                        except:
                            telemetries.append(0)

            if 'Datalogger' in device:
                for k in porta_DAQs:
                    i = 0
                    print('>>> Apro la Comunicazione con il DAQ\t')
                    INSTRUMENT = rm.open_resource(k)
                    INSTRUMENT.write('INIT')
                    time.sleep(0.05)
                    a = INSTRUMENT.query('FETCH?').split(',')
                    print('>>> log il DAQ\t')
                    for j in a:
                        if '+' in j or '-' in j:
                            telemetries.append(float(j[0:15]))
                            # print('>>> log la telemetry\t' + str(stamp_daq[i]))
                            i += 1
            if 'Wattmeter' in device:
                for k in porta_WTs:
                    i = 0
                    print('>>> Apro la Comunicazione con il ' + config_WATT[config_WATT['PORTA WT'] == k].reset_index()['MODELLO'][0])
                    INSTRUMENT = rm.open_resource(k)
                    if config_WATT[config_WATT['PORTA WT'] == k].reset_index()['MODELLO'][0] in ['WT230', 'WT210']:
                        a = INSTRUMENT.query('MEAS:NORM:VAL?').split(',')
                    if config_WATT[config_WATT['PORTA WT'] == k].reset_index()['MODELLO'][0] in ['WT500', 'WT3000', 'WT1800']:
                        a = INSTRUMENT.query(':NUM:NORM:VAL?').split(',')
                    for j in a:
                        if '+' in j or '-' in j:
                            telemetries.append(float(j[0:15]))
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
                else:
                    dati.to_csv(path_save + namefile + '.csv', sep=',', mode='a', index=False, header=False)

            tot_time += time_sample
            print('>>>')

        print('>>> End')
    except:
        path_watchdog = 'misc/watchdog.txt'
        file_object = open(path_watchdog, "w")
        file_object.write('1')
        file_object.close()


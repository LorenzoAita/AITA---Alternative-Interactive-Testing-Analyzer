import time
import serial
import pandas as pd
import pyvisa
import requests

from lib.Config import config_colonnina
from lib.Libs import *


def main_test():
    try:
        path_config = r'./Config/'
        print('>>> Config Test')
        print('>>>')
        test = pd.read_excel(path_config + 'Config.xlsx', sheet_name='Test')
        cmd = pd.read_excel(path_config + 'Config.xlsx', sheet_name='Comandi')
        if test['TEST'][0] != 'Manuale':
            a = True
        else:
            a = False
        while a == True:
            for i in range(0, len(cmd['TEMPO'])):
                # ora gestisco i relay
                for j in range(0, 8, 1):
                    if str(cmd['ARDU_RELAY' + str(j + 1)][i]) not in ['nan', '2', '2.0']:

                        cmd_relay = int(100 + cmd['ARDU_RELAY' + str(j + 1)][i])
                        num_relay = int(j + 1)
                        if i in [1, 2]:
                            var1 = 1
                            if num_relay < 5:
                                var2 = 3
                            else:
                                var2 = 4
                        else:
                            var1 = 0
                            var2 = 0

                        arduino_cmd = serial.Serial(port=test['ARDUINO_PORTA'][0], baudrate=38400, timeout=.3)
                        time.sleep(0.2)
                        arduino_cmd.write(bytes(
                            '1 ' + str(cmd_relay) + ' ' + str(num_relay) + ' ' + str(var1) + ' ' + str(
                                var2) + ' ' + str(
                                1 + cmd_relay + num_relay + var1 + var2), 'utf-8'))
                        time.sleep(1)

                        data = arduino_cmd.readline()
                        print(data)
                        # if (len(data) > 16) and ('CRC' not in str(data)) and ('nan' not in str(data)) and i != 0 and j in [1, 5]:# and (cmd['EVENTO'][i] == 1):
                        # va chiuso il relay e letta la frequenza. Se c'è, la prendo e chiudo i relay
                        #    try:
                        #        session = requests.Session()
                        #        rm = pyvisa.ResourceManager()
                        #        for k in test['ALIMENTATORE_PORTA']:
                        #            alim = Regatron(porta=k, rm=rm)
                        #            alim.stato(0)
                        #            time.sleep(0.2)
                        #            alim.curva(0)
                        #            time.sleep(0.2)
                        #            alim.power(20000)
                        #            time.sleep(0.2)
                        #            alim.voltage(600)
                        #            time.sleep(0.2)
                        #            alim.current(40)
                        #            time.sleep(0.2)
                        #            alim.stato(1)
                        #        time.sleep(0.2)
                        #    except:
                        #        time.sleep(30)
                        #    # time.sleep(1)
                        #        session = requests.Session()
                        #        rm = pyvisa.ResourceManager()
                        #        for k in test['ALIMENTATORE_PORTA']:
                        #            alim = Regatron(porta=k, rm=rm)
                        #            alim.curva(0)
                        #            time.sleep(0.2)
                        #            alim.power(20000)
                        #            time.sleep(0.2)
                        #            alim.voltage(600)
                        #            time.sleep(0.2)
                        #            alim.current(40)
                        #            time.sleep(0.2)
                        #            alim.stato(1)

                        time.sleep(0.2)
                        arduino_cmd.close()

                # ora gestisco i regatron
                for j in range(0, len(test['ALIMENTATORE_PORTA'])):
                    if str(cmd['ALIM_ON/OFF'][i]) not in ['2', 'nan', '2.0'] and str(
                            test['ALIMENTATORE_PORTA'][j]) != 'nan':
                        print(test['ALIMENTATORE_PORTA'][j])
                        session = requests.Session()
                        rm = pyvisa.ResourceManager()
                        alim_open = rm.open_resource(test['ALIMENTATORE_PORTA'][j])
                        if 'REGATRON' in alim_open.query("*IDN?"):
                            alim = Regatron(alim_open)
                        elif 'N8957APV' in alim_open.query("*IDN?"):
                            alim = Keysight(alim_open)
                        elif 'SORRENSEN' in alim_open.query("*IDN?"):
                            alim = Sorrensen(alim_open)
                        elif 'LAMBDA' in alim_open.query("*IDN?"):
                            alim = Lambda(alim_open)

                        alim.curva(cmd['ALIM_CURVA'][i])
                        time.sleep(0.2)
                        alim.power(cmd['ALIM_POW'][i])
                        time.sleep(0.2)
                        alim.voltage(cmd['ALIM_VOLT'][i])
                        time.sleep(0.2)
                        alim.current(cmd['ALIM_CURR'][i])
                        time.sleep(0.2)
                        alim.stato(cmd['ALIM_ON/OFF'][i])

                # ora gestisco gli inverter
                for j in range(0, len(test['INV_PORTA'])):
                    if str(cmd['INV_VAR'][i]) not in ['2', 'nan', '2.0'] and str(test['INV_PORTA'][j]) != 'nan':
                        df_eut = pd.read_excel(path_config + 'Config.xlsx', sheet_name='Inverter')
                        df_eut = df_eut[df_eut['ID'] == j].reset_index()
                        obj_com =OrionProtocol(ip=df_eut['IP'][0], device_id=j)
                        obj_com.write(test['INV_VAR'][i], test['INV_VAL'][i])
                        del df_eut
                # ora gestisco le colonnine
                for j in range(0, len(test['COL_PORTA'])):
                    if str(cmd['COL_VAR'][i]) not in ['2', 'nan', '2.0'] and str(test['COL_PORTA'][j]) != 'nan':
                        df_col = pd.read_excel(path_config + 'Config.xlsx', sheet_name='Colonnina')
                        df_col = df_col[df_col['IP'] == test['COL_PORTA'][j]].reset_index()
                        if df_col['MODALITA\''][0] == 'RTU':
                            baud_rate = 115200
                            com = ModbusClient(method='rtu',
                                               port=test['COL_PORTA'][j],
                                               baudrate=baud_rate,
                                               timeout=1,
                                               parity='N',
                                               stopbits=1,
                                               strict=False)
                        elif df_col['MODALITA\''][0] == 'TCP':
                            port = int(df_col['PORTA'][0])
                            com = df_col(host=test['COL_PORTA'][j], port=int(port))
                        WriteCol(test['COL_VAR'][i], com, test['COL_VAL'][i], df_col['ADDRESS'][0])
                        del df_col
                # ora gestisco la camera cliamtica
                for j in range(0, len(test['CC_PORTA'])):
                    if str(cmd['CC_ON/OFF'][i]) not in ['2', 'nan', '2.0'] and str(test['CC_PORTA'][j]) != 'nan':
                        if test['CC_ID'][0] == 'Weiss':
                            cc = Weiss(com=test['CC_PORTA'][0])
                        elif test['CC_ID'][0] == 'Angelantoni':
                            cc = Discovery(com=test['CC_PORTA'][0])
                        elif test['CC_ID'][0] == 'Endurance':
                            cc = Endurance(com=test['CC_PORTA'][0])

                        cc.set_temp_hum(cmd['CC_set point T'][i], cmd['CC_set point H'][i], cmd['CC_ON/OFF'][i])

                # ora gestisco il california
                for j in range(0, len(test['AC_PORTA'])):
                    if str(cmd['AC_ON/OFF'][i]) not in ['2', 'nan', '2.0'] and str(
                            test['AC_PORTA'][j]) != 'nan':
                        print(test['AC_PORTA'][j])
                        session = requests.Session()
                        rm = pyvisa.ResourceManager()
                        alim_open = rm.open_resource(test['AC_PORTA'][j])
                        alim = RS90(alim_open)
                        alim.set_freq(0)
                        alim.set_regenerative(cmd['AC_REGEN'][i])
                        time.sleep(0.2)
                        alim.set_volt(cmd['AC_VOUT'][i])
                        time.sleep(0.2)
                        alim.voltage(cmd['AC_FREQ'][i])
                        time.sleep(0.2)
                        alim.set_freq(cmd['AC_ON/OFF'][i])

                time.sleep(cmd['TEMPO'][i])

            if test['CICLO'][0] == 'Yes':
                a = True
            else:
                a = False
    except OSError as err:
        print(err)
        path_watchdog = 'misc/watchdog.txt'
        file_object = open(path_watchdog, "w")
        file_object.write('2')
        file_object.close()

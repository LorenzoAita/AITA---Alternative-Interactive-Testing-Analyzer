import pyvisa
import pandas as pd


def config_agilent(rm, path):
    print('>>> Config Agilent')
    print('>>>')
    path_config = path
    config_DAQ = pd.read_excel(path_config+'Config.xlsx', sheet_name='Agilent')
    porta = config_DAQ['PORTA DAQ'][0]
    start = 100
    stringa_T = ''
    stringa_K = ''
    stringa_VDC = ''
    stringa_VAC = ''
    stringa_IDC = ''
    stringa_IAC = ''
    stringa_Freq = ''
    stringa_Ohm = ''
    for i in range(0, len(config_DAQ['CASSETTO'])):
        if str(config_DAQ['TIPOLOGIA'][i]) == 'T':
            start += 1
            stringa_T += str(config_DAQ['CASSETTO'][i])+','
        if str(config_DAQ['TIPOLOGIA'][i]) == 'K':
            start += 1
            stringa_K += str(config_DAQ['CASSETTO'][i])+','
        if str(config_DAQ['TIPOLOGIA'][i]) == 'VDC':
            start += 1
            stringa_VDC += str(config_DAQ['CASSETTO'][i]) + ','
        if str(config_DAQ['TIPOLOGIA'][i]) == 'VAC':
            start += 1
            stringa_VAC += str(config_DAQ['CASSETTO'][i]) + ','
        if str(config_DAQ['TIPOLOGIA'][i]) == 'IDC':
            start += 1
            stringa_IDC += str(config_DAQ['CASSETTO'][i]) + ','
        if str(config_DAQ['TIPOLOGIA'][i]) == 'IAC':
            start += 1
            stringa_IAC += str(config_DAQ['CASSETTO'][i]) + ','
        if str(config_DAQ['TIPOLOGIA'][i]) == 'HZ':
            start += 1
            stringa_Freq += str(config_DAQ['CASSETTO'][i]) + ','
        if str(config_DAQ['TIPOLOGIA'][i]) == 'OHM':
            start += 1
            stringa_Ohm += str(config_DAQ['CASSETTO'][i]) + ','
        if str(config_DAQ['TIPOLOGIA'][i]) == '':
            start += 1

    stringa_tot = ''
    inst = rm.open_resource(porta)
    print('>>> '+inst.query("*IDN?").split('\n')[0])
    inst.write('*RST')
    # CONFIGURAZIONE TERMOCOPPIE TIPO T
    if stringa_T != '':
        stringa_tot += stringa_T
        inst.write("CONF:TEMP TC,T ,(@"+stringa_T[0:-1]+")")
        inst.write("UNIT:TEMP C ,(@"+stringa_T[0:-1]+")")
    # CONFIGURAZIONE TERMOCOPPIE TIPO K
    if stringa_K != '':
        stringa_tot += stringa_K
        inst.write("CONF:TEMP TC,K ,(@"+stringa_K[0:-1]+")")
        inst.write("UNIT:TEMP C ,(@"+stringa_K[0:-1]+")")
    # CONFIGURAZIONE FREQUENZE
    if stringa_Freq != '':
        stringa_tot += stringa_Freq
        inst.write("CONF:FREQ (@"+stringa_Freq[0:-1]+")")
    # CONFIGURAZIONE VDC
    if stringa_VDC != '':
        stringa_tot += stringa_VDC
        inst.write("CONF:VOLT:DC (@"+stringa_VDC[0:-1]+")")
    # CONFIGURAZIONE VAC
    if stringa_VAC != '':
        stringa_tot += stringa_VAC
        inst.write("CONF:VOLT:AC (@"+stringa_VAC[0:-1]+")")
    # CONFIGURAZIONE IDC
    if stringa_IDC != '':
        stringa_tot += stringa_IDC
        inst.write("CONF:CURR:DC (@"+stringa_IDC[0:-1]+")")
    #inst.write("SENS:CURR:DC:RANG:AUTO ON ,(@"+stringa_IDC[0:-1]+")")
    # CONFIGURAZIONE IAC
    if stringa_IAC != '':
        stringa_tot += stringa_IAC
        inst.write("CONF:CURR:AC (@"+stringa_IAC[0:-1]+")")
    #inst.write("SENS:CURR:DC:RANG:AUTO ON ,(@"+stringa_IAC[0:-1]+")")
    # CONFIGURAZIONE RESISTENZE
    if stringa_Ohm != '':
        stringa_tot += stringa_Ohm
        inst.write("CONF:RES (@"+stringa_Ohm[0:-1]+")")

    inst.write("ROUT:SCAN (@"+stringa_tot[0:-1]+")")    # Imposto i canali
    print('>>> '+inst.query("ROUT:SCAN?").split('\n')[0])
    return config_DAQ['LABEL'], porta


def config_inverter(path):
    print('>>> Config Inverter')
    print('>>>')
    path_config = path
    config_inv = pd.read_excel(path_config+'Config.xlsx', sheet_name='Inverter')
    id = str(config_inv['ID'][0])
    ip = str(config_inv['IP DEVICE'][0])
    telemetrie = list()
    for i in config_inv['TELEMETRIE']:
        telemetrie.append(i)
    return id, ip, telemetrie


def config_wt(rm, path, model):
    print('>>> Config Wattmeter')
    print('>>>')
    path_config = path
    config_WT = pd.read_excel(path_config + 'Config.xlsx', sheet_name='Wattmeter')
    porta = config_WT['PORTA WT'][0]
    inst = rm.open_resource(porta)
    print('>>> ' + inst.query("*IDN?").split('\n')[0])
    if model == 'WT230':
        inst.write('MEAS:NORM:ITEM:PRES CLE')
        for i in range(0, len(config_WT['CHANNEL'])):
            if str(config_WT['CHANNEL'][i]) != 'nan':
                inst.write("MEASURE:NORMAL:ITEM:"+str(config_WT['TIPOLOGIA'][i])+":ELEMENT"+str(int(config_WT['CHANNEL'][i]))+" ON")
    if model == 'WT500' or model == 'WT3000':
        inst.write(':NUM:NORM:CLE ALL')
        if model == 'WT500':
            inst.write(':NUM:FORM ASC')
            inst.write(':NUM:NORM:NUM 200')
        j = 1
        for i in range(0, len(config_WT['CHANNEL'])):
            if str(config_WT['CHANNEL'][i]) != 'nan':
                inst.write(":NUM:NORM:ITEM" + str(j) + " " + str(config_WT['TIPOLOGIA'][i]) + "," + str(int(config_WT['CHANNEL'][i])))
                j += 1
    return config_WT['LABEL'], porta

import pandas as pd
from pymodbus.client.sync import ModbusTcpClient as ModbusClientTCP
from pymodbus.client.sync import ModbusSerialClient as ModbusClient
import json

def config_agilent(rm, path):
    print('>>> Config Agilent')
    print('>>>')
    path_config = path
    config_DAQ = pd.read_excel(path_config+'Config.xlsx', sheet_name='Agilent')
    porta_list = list()
    meas_list = list()
    for k in range(0, len(config_DAQ['PORTA DAQ'])):
        if str(config_DAQ['PORTA DAQ'][k]) != 'nan':
            porta = config_DAQ['PORTA DAQ'][k]
            inst = rm.open_resource(porta)
            start = 100
            stringa_T = ''
            stringa_K = ''
            stringa_VDC = ''
            stringa_VAC = ''
            stringa_IDC = ''
            stringa_IAC = ''
            stringa_Freq = ''
            stringa_Ohm = ''
            df_daq = config_DAQ[config_DAQ['ID_LETTURA DAQ'+str(k+1)] == 1].reset_index()
            for i in range(0, len(df_daq['CASSETTO'])):
                if str(df_daq['TIPOLOGIA'][i]) == 'T':
                    start += 1
                    stringa_T += str(df_daq['CASSETTO'][i])+','
                if str(df_daq['TIPOLOGIA'][i]) == 'K':
                    start += 1
                    stringa_K += str(df_daq['CASSETTO'][i])+','
                if str(df_daq['TIPOLOGIA'][i]) == 'VDC':
                    start += 1
                    stringa_VDC += str(df_daq['CASSETTO'][i]) + ','
                if str(df_daq['TIPOLOGIA'][i]) == 'VAC':
                    start += 1
                    stringa_VAC += str(df_daq['CASSETTO'][i]) + ','
                if str(df_daq['TIPOLOGIA'][i]) == 'IDC':
                    start += 1
                    stringa_IDC += str(df_daq['CASSETTO'][i]) + ','
                if str(df_daq['TIPOLOGIA'][i]) == 'IAC':
                    start += 1
                    stringa_IAC += str(df_daq['CASSETTO'][i]) + ','
                if str(df_daq['TIPOLOGIA'][i]) == 'HZ':
                    start += 1
                    stringa_Freq += str(df_daq['CASSETTO'][i]) + ','
                if str(df_daq['TIPOLOGIA'][i]) == 'OHM':
                    start += 1
                    stringa_Ohm += str(df_daq['CASSETTO'][i]) + ','
                if str(df_daq['TIPOLOGIA'][i]) == '':
                    start += 1

            stringa_tot = ''
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

            porta_list.append(porta)
            meas_list.append(list(df_daq['LABEL']))

    return meas_list, porta_list


def config_inverter(path):
    print('>>> Config Inverter')
    print('>>>')
    path_config = path
    config_inv = pd.read_excel(path_config+'Config.xlsx', sheet_name='Inverter')
    telemetrie_tot = list()
    ips = list()
    ids = list()
    labels = list()
    for k in range(0, len(config_inv['ID'])):
        if str(config_inv['ID'][k]) != 'nan':
            id = str(config_inv['ID'][k])
            if '.' in id:
                ids.append(str(id[:-2]))
            else:
                ids.append(str(id))
            ip = str(config_inv['IP DEVICE'][k])
            ips.append(ip)
            telemetrie = list()
            df_inv = config_inv[config_inv['ID_LETTURA INV'+str(k+1)] == 1].reset_index()
            for i in df_inv['TELEMETRIE']:
                telemetrie.append(i)
            label = list()
            for i in df_inv['LABEL']:
                label.append(i)
            telemetrie_tot.append(telemetrie)
            labels.append(label)
    return ids, ips, telemetrie_tot, labels


def config_wt(rm, path):
    print('>>> Config Wattmeter')
    print('>>>')
    path_config = path
    config_WT = pd.read_excel(path_config + 'Config.xlsx', sheet_name='Wattmeter')
    porta_list = list()
    meas_list = list()
    for k in range(0, len(config_WT['MODELLO'])):
        if str(config_WT['PORTA WT'][k]) != 'nan':
            porta = config_WT['PORTA WT'][k]
            porta_list.append(porta)
            df_wt_model = pd.read_excel(path_config + 'Config.xlsx', sheet_name=config_WT['MODELLO'][k])
            df_wt = df_wt_model[df_wt_model['ID_LETTURA WT' + str(k + 1)] == 1].reset_index()
            inst = rm.open_resource(porta)
            print('>>> ' + inst.query("*IDN?").split('\n')[0])
            model = config_WT['MODELLO'][k]
            if model in ['WT230', 'WT210']:
                inst.write('MEAS:NORM:ITEM:PRES CLE')
                for i in range(0, len(df_wt['CHANNEL'])):
                    if str(config_WT['CHANNEL'][i]) != 'nan':
                        inst.write("MEASURE:NORMAL:ITEM:"+str(df_wt['TIPOLOGIA'][i])+":ELEMENT"+str(int(df_wt['CHANNEL'][i]))+" ON")
            if model in ['WT500', 'WT3000', 'WT1800']:
                inst.write(':NUM:NORM:CLE ALL')
                if model == 'WT500':
                    inst.write(':NUM:FORM ASC')
                    inst.write(':NUM:NORM:NUM 200')
                j = 1
                for i in range(0, len(df_wt['CHANNEL'])):
                    if str(df_wt['CHANNEL'][i]) != 'nan':
                        inst.write(":NUM:NORM:ITEM" + str(j) + " " + str(df_wt['TIPOLOGIA'][i]) + "," + str(int(df_wt['CHANNEL'][i])))
                        j += 1

            meas_list.append(list(df_wt['LABEL']))
    return meas_list, porta_list, df_wt_model, df_wt


def config_ponte(rm, path):
    print('>>> Config Bridge')
    print('>>>')
    path_config = path
    config_bridge = pd.read_excel(path_config + 'Config.xlsx', sheet_name='Bridge')
    porta = config_bridge['PORTA BRIDGE'][0]
    inst = rm.open_resource(porta)
    print('>>> ' + inst.query("*IDN?").split('\n')[0])
    inst.write('*RST')
    # Imposto il LVL del bridge
    lvl = config_bridge['LEVEL'][0]
    if lvl[-1] == 'V':
        inst.write(':VOLT '+str(lvl.split('V')[0]))
    elif lvl[-1] == 'A':
        inst.write(':CURR '+str(lvl.split('A')[0]))

    # Imposto la velocità di misura
    MEAS_SPEED = config_bridge['MEAS TIME'][0]
    inst.write(':APER ' + str(MEAS_SPEED))

    # Imposto la CORREZIONE
    MEAS_CORR = config_bridge['CORRECTION']
    MEAS_CORR = MEAS_CORR.fillna('')
    inst.write(':CORR:OPEN:STATE OFF')
    inst.write(':CORR:SHOR:STATE OFF')
    inst.write(':CORR:LOAD:STATE OFF')
    if 'OFF' not in MEAS_CORR:
        for i in MEAS_CORR:
            if i != '':
                inst.write(':CORR:'+str(i)+':STATE ON')
    inst.write(':CORR:LENG '+str(config_bridge['CORRECTION LENGTH'][0]))

    return config_bridge['TIPOLOGIA'], porta


def config_colonnina(path):
    path_config = path
    config_inv = pd.read_excel(path_config+'Config.xlsx', sheet_name='Colonnina')
    telemetries = list()
    misures = list()
    coms = list()
    addresses_tot = list()
    for k in range(0, len(config_inv['IP'])):
        if str(config_inv['IP'][k]) != 'nan':
            ip = str(config_inv['IP'][k])
            com = 0
            df_col = config_inv[config_inv['ID_LETTURA COL'+str(k+1)]==1]
            if config_inv['MODALITA\''][k] == 'RTU':
                baud_rate = 115200
                com = ModbusClient(method='rtu',
                                               port=ip,
                                               baudrate=baud_rate,
                                               timeout=1,
                                               parity='N',
                                               stopbits=1,
                                               strict=False)
            elif config_inv['MODALITA\''][0] == 'TCP':
                port = int(config_inv['PORTA'][0])
                com = ModbusClientTCP(host=ip, port=int(port))
            coms.append(com)
            add = config_inv['ADDRESS'][k]
            telemetrie = list()
            reg = list()
            addresses = list()
            for j in range(0, int(add)):
                for i in range(0, len(df_col['TELEMETRIE'])):
                    telemetrie.append(df_col['TELEMETRIE'][i]+'_Add'+str(j+1))
                    reg.append(df_col['REGISTRO'][i])
                    addresses.append(j+1)
            telemetries.append(telemetrie)
            addresses_tot.append(addresses)
            misure = list()
            read_js = json.load(open(path_config+'MisureColonnine.json'))
            for i in reg:
                for j in  read_js["holding_registers"]:
                    if i == j["address"]:
                        misure.append(j)
            misures.append(misure)
    return telemetries, misures, coms, addresses_tot

import requests
import plotly.graph_objects as go
import pandas as pd
import struct
import time
import serial


class OrionProtocol:

    def __init__(self, ip, device_id):
        self.ip = ip
        self.ip_connector = '127.0.0.1:5000'
        self.device_id = device_id
        self.timeout = 30

    def get_data(self, url):

        rsp = requests.get(
            'http://' + self.ip_connector + '/' + self.ip + '/connector/v1/orion/data/' + self.device_id + '/' + url,
            timeout=1)
        if rsp.status_code == 200:
            return str(rsp.json()), str(rsp.status_code)
        else:
            return str(0), str(rsp.status_code)

    def write(self, url, data):
        rsp = requests.post(
            'http://' + self.ip_connector + '/' + self.ip + '/connector/v1/orion/data/' + self.device_id + '/' + url,
            json=data, verify=False)


def plot_runtime(step_graph, dati_stamp, plot):
    fig = go.Figure()
    fig.update_layout(
        template='simple_white',
        legend=dict(orientation="v", x=1.1, y=1),
        xaxis=dict(
            domain=[0.2, 0.8]
        ),
        yaxis=dict(
            title='Temperature [°C]',
            titlefont=dict(
                color="#1f77b4"
            ),
            tickfont=dict(
                color="#1f77b4"
            )
        ),
        yaxis3=dict(
            title='Power [W]',
            titlefont=dict(
                color="#EC00FF"
            ),
            tickfont=dict(
                color="#EC00FF"
            ),
            anchor="free",
            overlaying="y",
            side="left",
            position=0.13
        ),
        yaxis4=dict(
            title='Text',
            titlefont=dict(
                color="#d62728"
            ),
            tickfont=dict(
                color="#d62728"
            ),
            anchor="x",
            overlaying="y",
            side="right"
        ),
        yaxis2=dict(
            title='Voltage [V]',
            titlefont=dict(
                color="#000000"
            ),
            tickfont=dict(
                color="#000000"
            ),
            anchor="free",
            overlaying="y",
            side="right",
            position=0.95
        )
    )
    if step_graph < 5:
        for i in range(0, len(plot)):
            if plot[i][1] == 'W':
                fig.add_scatter(x=dati_stamp['Date'], y=dati_stamp[plot[i][0]], mode='lines', yaxis='y4',
                                name=plot[i][0])
            elif plot[i][1] == 'V':
                fig.add_scatter(x=dati_stamp['Date'], y=dati_stamp[plot[i][0]], mode='lines', yaxis='y2',
                                name=plot[i][0])
            elif plot[i][1] == 'P':
                fig.add_scatter(x=dati_stamp['Date'], y=dati_stamp[plot[i][0]], mode='lines', yaxis='y3',
                                name=plot[i][0])
            elif plot[i][1] == 'T':
                fig.add_scatter(x=dati_stamp['Date'], y=dati_stamp[plot[i][0]], mode='lines', yaxis='y',
                                name=plot[i][0])
        fig.show()
    else:
        for i in range(0, len(plot)):
            if plot[i][1] == 'W':
                fig.add_scatter(x=dati_stamp['Date'], y=dati_stamp[plot[i][0]], mode='lines', yaxis='y4',
                                name=plot[i][0])
            elif plot[i][1] == 'V':
                fig.add_scatter(x=dati_stamp['Date'], y=dati_stamp[plot[i][0]], mode='lines', yaxis='y2',
                                name=plot[i][0])
            elif plot[i][1] == 'P':
                fig.add_scatter(x=dati_stamp['Date'], y=dati_stamp[plot[i][0]], mode='lines', yaxis='y3',
                                name=plot[i][0])
            elif plot[i][1] == 'T':
                fig.add_scatter(x=dati_stamp['Date'], y=dati_stamp[plot[i][0]], mode='lines', yaxis='y',
                                name=plot[i][0])
        fig.show()
    return True


def meas_bridge(inst, log, data, path, save):
    config_bridge = pd.read_excel(path + 'Config.xlsx', sheet_name='Bridge')
    freq_start = config_bridge['FREQUENZA'][0]
    freq_end = config_bridge['FREQUENZA'][1]
    # freq_sample = config_bridge['FREQUENZA'][2]
    nome_file = pd.read_excel(path + 'Config.xlsx', sheet_name='Strumenti')['NOME OUTPUT'][0]
    while freq_end + 1 > freq_start:
        if freq_start < 1e6:
            freq_sample = 1e4
        if freq_start < 1e5:
            freq_sample = 1e3
        if freq_start < 1e4:
            freq_sample = 1e2
        if freq_start < 1e3:
            freq_sample = 1e1
        if freq_start < 1e2:
            freq_sample = 1e0
        telemetries = list()
        inst.write('INIT')
        inst.write(':FREQ ' + str(freq_start))
        telemetries.append(freq_start)
        for i in config_bridge['TIPOLOGIA'].fillna(''):
            if i != '':
                inst.write(':FUNC:IMP ' + str(i))
                a = inst.query('FETCH?').split(',')
                if i in ['RX', 'GB']:
                    name1 = i[0:1]
                    name2 = i[1:]  # name1 + '-' + i[2:]
                else:
                    name1 = i[0:2]
                    name2 = i[2:]  # name1 + '-' + i[2:]  # name1 + '-' + i[2:]
                print('>>> log la telemetry\t' + str(name1) + '\talla frequenza\t' + str(freq_start))
                telemetries.append(float(a[0]))
                print('>>> log la telemetry\t' + str(name2) + '\talla frequenza\t' + str(freq_start))
                telemetries.append(float(a[1]))
        freq_start += freq_sample
        dati = pd.DataFrame()
        dati_T = pd.DataFrame()
        dati_T = dati_T.append(telemetries)
        dati = pd.concat([dati, dati_T.T], axis=0, ignore_index=True)
        for i in range(0, len(log)):
            dati.rename(columns={i: log[i]}, inplace=True)
        data = pd.concat([data, dati], axis=0, ignore_index=True)

    print('>>> Salvo i dati\t')
    data.to_csv(save + nome_file + '.csv', sep=',', index=False)
    print('>>> Salvo il grafico\t')
    fig = go.Figure()
    fig.update_layout(
        template='simple_white',
        legend=dict(orientation="v", x=1.1, y=1)
    )
    for i in range(1, len(log)):
        if i % 2 == 0:
            fig.add_scatter(x=data['Frequenza'].astype(float), y=data[log[i]].astype(float), mode='lines', name=log[i],
                            yaxis='y2')
        else:
            fig.add_scatter(x=data['Frequenza'].astype(float), y=data[log[i]].astype(float), mode='lines', name=log[i],
                            yaxis='y')

    fig.update_layout(
        xaxis=dict(
            domain=[0.2, 1],
            title='Frequenza [Hz]'
        ),
        yaxis=dict(
            title='Misura Primaria',
        ),
        yaxis3=dict(
            title='Output Power [W]',
            anchor="free",
            overlaying="y",
            side="left",
            position=0.13
        ),
        yaxis2=dict(
            title='Misura Secondaria',
            anchor="x",
            overlaying="y",
            side="right"
        ),
        yaxis4=dict(
            title='Freq. [Hz]',
            anchor="free",
            overlaying="y",
            side="right",
            position=0.95
        ),
        template='simple_white',
        legend=dict(orientation="v", x=1.1, y=1),
        hovermode=False
    )

    fig.write_image(save + nome_file + '.png', width=1920, height=1080)
    # Add dropdown
    fig.update_layout(
        updatemenus=[
            dict(
                type="buttons",
                # direction="left",
                buttons=list([
                    dict(
                        args=["hovermode", False],
                        label="No Select Data Lines",
                        method="relayout"
                    ),
                    dict(
                        args=["hovermode", "x"],
                        label="Select All Data Lines",
                        method="relayout"
                    ),
                    dict(
                        args=["hovermode", "x unified"],
                        label="Second Select All Data Lines",
                        method="relayout"
                    ),
                    dict(
                        args=["hovermode", "closest"],
                        label="Select Single Data Line",
                        method="relayout"
                    )
                ]),
                pad={"r": 10, "t": 8},
                showactive=True,
                x=-0.11,
                xanchor="left",
                y=1,
                # yanchor="top"
            ),
            # dict(
            # type="buttons",
            # direction="left",
            # buttons=list([
            #   dict(
            #        args=["dragmode", "drawline"],
            #         label="Draw Line",
            #          method="relayout"
            #    ),
            #     dict(
            #           args=["dragmode", "drawcircle"],
            #          label="Draw Circle",
            #           method="relayout"
            #       ),
            #       dict(
            #           args=["dragmode", "drawrect"],
            #         label="Draw Rectangle",
            #           method="relayout"
            #        ),
            #    ]),
            #    pad={"r": 10, "t": 8},
            #    showactive=True,
            #    x=-0.038,
            #    xanchor="left",
            #    y=0.7,
            # yanchor="top"
            # ),
            # dict(
            #    type="buttons",
            # direction="left",
            #    buttons=list([
            #        dict(
            #            args=["yaxis.autorange", True],
            #            args2=["xaxis.autorange", True],
            #            label="Reset Axis",
            #            method="relayout"
            #        ),
            #    ]),
            #    pad={"r": 10, "t": 8},
            #    showactive=True,
            #    x=-0.03,
            #    xanchor="left",
            #    y=0.2,
            # yanchor="top"
            # ),
        ]
    )
    fig.write_html(save + nome_file + '.html', auto_open=False,
                   config={'modeBarButtonsToAdd': [  # 'drawline',
                       # 'drawclosedpath',
                       # 'drawcircle',
                       # 'drawrect',
                       # "hoverClosestCartesian", "hoverCompareCartesian"
                       # 'toggleSpikelines',
                       # 'resetViews',
                       # 'toImage',
                       # 'select',
                       # 'hoverCompareCartesian',
                       # 'hoverCompareData',
                       # 'hovermode'
                   ],
                       'modeBarButtonsToRemove': ['resetScale',
                                                  'pan',
                                                  'zoomIn',
                                                  'zoomOut',
                                                  'zoom'],
                       'toImageButtonOptions': {'height': None, 'width': None, },
                       'displaylogo': False,
                       'displayModeBar': True
                   }
                   )

    return True


def ReadCol(reg, client, add):
    result = client.read_holding_registers(reg['address'],
                                           reg['length'],
                                           unit=add,
                                           timeout=1)  # da modificare con l'addr del registro modbus
    result2 = result.registers
    data_type = reg['type']
    number = 0
    if not result.isError():
        if data_type == "UINT":
            if len(result2) == 1:
                number = result2.pop()
            elif len(result2) == 2:
                number = result2.pop()
                number = number << 16
                number = number | result2.pop()
            elif len(result2) == 4:
                number = result2.pop()
                number = number << 16
                number = number | result2.pop()
                number = number << 16
                number = number | result2.pop()
                number = number << 16
                number = number | result2.pop()
        elif data_type == "hex":
            number = result2.pop()
        elif data_type == "FLOAT":
            number = result2.pop()
            number = number << 16
            number = number | result2.pop()
            number = struct.unpack('>f', number.to_bytes(4, 'big'))[0]
        elif data_type == "DOUBLE":
            number = result2.pop()
            number = number << 16
            number = number | result2.pop()
            number = number << 16
            number = number | result2.pop()
            number = number << 16
            number = number | result2.pop()
            number = struct.unpack('>d', number.to_bytes(8, 'big'))[0]
        elif data_type == "list":
            pass
    return number


def WriteCol(reg, client, value, add):
    result = client.write_registers(reg,
                                    value,
                                    unit=add)  # da modificare con l'addr del registro modbus
    return True


class Regatron:
    def __init__(self, rm):
        self.rm = rm

    def stato(self, stato):
        INSTRUMENT_alim = self.rm
        if stato == 1:
            INSTRUMENT_alim.write("OUTP ON")
        elif stato == 0:
            INSTRUMENT_alim.write("OUTP OFF")

    def power(self, pow):
        INSTRUMENT_alim = self.rm
        INSTRUMENT_alim.write("POW " + str(pow))

    def voltage(self, volt):
        INSTRUMENT_alim = self.rm
        INSTRUMENT_alim.write("VOLT " + str(volt))

    def current(self, cur):
        INSTRUMENT_alim = self.rm
        INSTRUMENT_alim.write("CURR " + str(cur))

    def curva(self, curva):
        INSTRUMENT_alim = self.rm
        INSTRUMENT_alim.write("topc:reg:writ #H5cc7,  " + str(curva))


class Keysight:
    def __init__(self, rm):
        self.rm = rm

    def stato(self, stato):
        INSTRUMENT_alim = self.rm
        if stato == 1:
            INSTRUMENT_alim.write("OUTP ON")
        elif stato == 0:
            INSTRUMENT_alim.write("OUTP OFF")

    def power(self, pow):
        INSTRUMENT_alim = self.rm
        INSTRUMENT_alim.write("POW " + str(pow))

    def voltage(self, volt):
        INSTRUMENT_alim = self.rm
        INSTRUMENT_alim.write("VOLT " + str(volt))

    def current(self, cur):
        INSTRUMENT_alim = self.rm
        INSTRUMENT_alim.write("CURR " + str(cur))

    def curva(self, curva):
        INSTRUMENT_alim = self.rm
        INSTRUMENT_alim.write("topc:reg:writ #H5cc7,  " + str(curva))


class Weiss():
    def __init__(self, com):
        self.com = com
        self.bound_rate = 38400
        self.timeout = 30

    def set_temp_hum(self, value_t, value_h, value_on):
        cc = serial.Serial(port=self.com, baudrate=9600, parity="N", stopbits=1, timeout=.3, bytesize=8)
        cc.write(bytes('$01I\r', 'utf-8'))
        time.sleep(6)
        if value_t < 0:
            if value_t > 0 and value_t < 100:
                cc.write(bytes(
                    '$01E -0' + value_t + '.0 00' + value_h + '.0 0100.0 0005.0 0030.0 0' + value_on + '000000000000000000000000000000\r',
                    'utf-8'))
            else:
                cc.write(bytes(
                    '$01E -' + value_t + '.0 00' + value_h + '.0 0100.0 0005.0 0030.0 0' + value_on + '000000000000000000000000000000\r',
                    'utf-8'))
        if value_t > 0 and value_t < 100:
            cc.write(bytes(
                '$01E 00' + value_t + '.0 00' + value_h + '.0 0100.0 0005.0 0030.0 0' + value_on + '000000000000000000000000000000\r',
                'utf-8'))
        else:
            cc.write(bytes(
                '$01E 0' + value_t + '.0 00' + value_h + '.0 0100.0 0005.0 0030.0 0' + value_on + '000000000000000000000000000000\r',
                'utf-8'))

        cc.close()


class Discovery():
    def __init__(self, com):
        self.com = com

    def set_temp_hum(self, value_t, value_h, value_on):
        cc = serial.Serial(port=self.com, baudrate=9600, parity="N", stopbits=1, timeout=.3, bytesize=8)

        if value_on == '1':
            cc.write(prepare_packet(504, float(value_t)))
            time.sleep(3)
            if value_h < '15':
                status_on_value = 2.3693558e-38  # valore accensione camera + temperatura
            else:
                cc.write(prepare_packet(508, float(value_h)))
                time.sleep(3)
                status_on_value = 3.790969281401877e-37  # valore accensione camera + temp + umidità
        else:
            status_on_value = 0  # spengi tutto
        cc.write(prepare_packet(500, status_on_value))

        cc.close()


def four_numbers(val):
    return list(reversed(list(struct.pack('f', val))))


def crc_modbus(msg):
    data = bytearray(msg)
    crc = 0xFFFF
    for pos in data:
        crc ^= pos
        for i in range(8):
            if (crc & 1) != 0:
                crc >>= 1
                crc ^= 0xA001
            else:
                crc >>= 1
    return crc & 0xFF, crc >> 8


def answer_to_float(answer):
    partial = list(answer)
    result = struct.unpack('f', bytearray([partial[6], partial[5], partial[4], partial[3]]))
    result = float(result[0])
    return round(result, 1)


def register_address(address):
    if address > 256:
        return 1, address - 256
    return 0, address


def prepare_packet(register, value=None):
    address = register_address(register)
    msg = [17, 3 if value is None else 16, address[0], address[1], 0, 2]
    if value is not None:
        msg.append(4)
        msg.extend(four_numbers(value))

    msg.extend(crc_modbus(msg))
    return msg


path_config = r'./Config/'

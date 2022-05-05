import requests
import plotly.graph_objects as go
import pandas as pd
from pymodbus.payload import BinaryPayloadDecoder

class OrionProtocol:

    def __init__(self, ip, device_id):
        self.ip = ip
        self.ip_connector = '127.0.0.1:5000'
        self.device_id = device_id
        self.timeout = 30

    def get_data(self, url):
        """ get data

        Keyword arguments:
        url -- attribute object (str)
        """

        rsp = requests.get(
            'http://' + self.ip_connector + '/' + self.ip + '/connector/v1/orion/data/' + self.device_id + '/' + url,
            timeout=1)
        if rsp.status_code == 200:
            return str(rsp.json()), str(rsp.status_code)
        else:
            return str(0), str(rsp.status_code)

'''
    def write(self, url, data, method="POST"):
        """ nucleic write

        Keyword arguments:
        url -- message (str)
        data -- dict, list of tuples, bytes, or file-like object to send in the body (optional)
        """

        try:
            if data is not None and method == "POST":
                # rsp = requests.post('http://' + self.ip_connector + '/' + self.ip + '/connector' + url, json=data)
                rsp = requests.post('http://' + self.ip_connector + '/' + self.ip + '/connector' + url, json=data, verify=False)
            elif method == 'DELETE':
                rsp = requests.delete('http://' + self.ip_connector + '/' + self.ip + '/connector' + url, verify=False)
            else:
                # rsp = requests.post('http://' + self.ip_connector + '/' + self.ip + '/connector' + url)
                rsp = requests.post('http://' + self.ip_connector + '/' + self.ip + '/connector' + url)
            logger.debug(method + ' >> Orion API response status code is ' + str(rsp.status_code) + ' from url ' + rsp.url)
            logger.debug(method + "'s body >> " + str(data))
            if rsp.status_code == 200:
                return True, rsp.status_code
            else:
                logger.warning(
                    method + ' >> Orion API response status code is ' + str(rsp.status_code) + ' from url ' + rsp.url)
                return False, rsp.status_code
        except Exception as e:
            # uncomment this line below to log error with complete stacktrace
            # logger.error('Eut not responding ip: ' + str(self.ip), exc_info=sys.exc_info())
            logger.error('Eut not responding ip: ' + str(self.ip))
            return False, 500
'''

def plot_runtime(step_graph, dati_stamp, plot):
    fig = go.Figure()
    fig.update_layout(
        template='simple_white',
        legend=dict(orientation="v", x=1.1, y=1)
    )
    if step_graph < 5:
        for i in plot:
            fig.add_scatter(x=dati_stamp['Data'], y=dati_stamp[i], mode='lines+markers', name=i)
        fig.show()
    else:
        for i in plot:
            fig.add_scatter(x=dati_stamp['Data'], y=dati_stamp[i], mode='lines', name=i)
        fig.show()
    return True


def meas_bridge(inst, log, data, path, save):
    config_bridge = pd.read_excel(path + 'Config.xlsx', sheet_name='Bridge')
    freq_start = config_bridge['FREQUENZA'][0]
    freq_end = config_bridge['FREQUENZA'][1]
    freq_sample = config_bridge['FREQUENZA'][2]
    while freq_end + 1 > freq_start:
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
    data.to_csv(save + 'DataBridge.csv', sep=',', index=False)
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

    fig.write_image(save + 'FigBridge.png', width=1920, height=1080)
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
    fig.write_html(save + 'FigBridge.html', auto_open=False,
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



# json_config = None
# def read_regs(connection, start_address, length, timeout_reply=1):
#     if timeout_reply != connection.timeout:
#         connection.close()
#         connection.timeout = timeout_reply
#         connection.connect()
#     response = connection.read_holding_registers(address=start_address,
#                                                   count=length,
#                                                   unit=int(json_config["MODBUS_SLAVE_ADDRESS"]))
#     if not response.isError():
#         print(">>> R ADDRESS " + str(start_address) + ":\t" + list_to_string(response.registers))
#         return response.registers
#     else:
#         print(">>> R FAIL ADDRESS " + str(start_address))
#         raise TimeoutError
#
#
# def list_to_string(s):
#     # initialize an empty string
#     str1 = ""
#     try:
#         for ele in s:
#             if type(ele) is int:
#                 str1 += str(ele)
#             else:
#                 str1 += ele
#             str1 += " "
#     except TypeError:
#         str1 = str(s)
#     except NameError:  # traverse in the string
#         str1 = "none"
#     return str1
#
# def ReadInfo(self, client, info):
#     reg = self.holding_registers[info]
#     result = client.read_holding_registers(reg['address'],
#                                            reg['length'],
#                                            unit=self.slave_address)
#     if not result.isError():
#         if reg['type'] == 'HEXSTRING':
#             decoders = [BinaryPayloadDecoder.fromRegisters([x],
#                                                            byteorder=endianess_converter[
#                                                                self.registers['access_rules']['word_endianness']],
#                                                            wordorder=endianess_converter[
#                                                                self.registers['access_rules'][
#                                                                    'register_endianness']]
#                                                            ) for x in result.registers]
#             value = ""
#             for d in decoders:
#                 v = d.decode_16bit_uint()
#                 for _ in range(0, 4):
#                     c = v >> 12
#                     if c >= 10:
#                         value = value + chr(ord('a') + c - 10)
#                     else:
#                         value = value + str(c)
#                     v = v << 4
#             value = v
#         elif reg['type'] == 'BYTESTREAM':
#             decoder = BinaryPayloadDecoder.fromRegisters(result.registers,
#                                                          byteorder=endianess_converter[
#                                                              self.registers['access_rules']['word_endianness']],
#                                                          wordorder=endianess_converter[
#                                                              self.registers['access_rules']['register_endianness']]
#                                                          )
#             v = []
#             for i in range(0, len(result.registers)):
#                 v.append(decoder.decode_8bit_uint())
#                 v.append(decoder.decode_8bit_uint())
#             byteswapped = bytearray(len(v))
#             if self.registers['access_rules']['register_endianness'] == 'LITTLE_ENDIANNESS':
#                 byteswapped[0::2] = v[1::2]
#                 byteswapped[1::2] = v[0::2]
#             else:
#                 byteswapped[0::1] = v[0::1]
#             value = byteswapped
#         else:
#             decoder = BinaryPayloadDecoder.fromRegisters(result.registers,
#                                                          byteorder=endianess_converter[
#                                                              self.registers['access_rules']['word_endianness']],
#                                                          wordorder=endianess_converter[
#                                                              self.registers['access_rules']['register_endianness']]
#                                                          )
#             if reg['type'] == 'UINT':
#                 if reg['length'] == 1:
#                     value = decoder.decode_16bit_uint()
#                 elif reg['length'] == 2:
#                     value = decoder.decode_32bit_uint()
#                 elif reg['length'] == 4:
#                     value = decoder.decode_64bit_uint()
#                 else:
#                     raise ValueError(f"Not managed UINT with length {reg['length']}")
#             elif reg['type'] == 'FLOAT':
#                 if reg['length'] == 2:
#                     value = decoder.decode_32bit_float()
#                 else:
#                     raise ValueError(f"Not managed FLOAT with length {reg['length']}")
#             elif reg['type'] == 'DOUBLE':
#                 if reg['length'] == 4:
#                     value = decoder.decode_64bit_float()
#                 else:
#                     raise ValueError(f"Not managed DOUBLE with length {reg['length']}")
#
#     else:
#         value = None
#     return value
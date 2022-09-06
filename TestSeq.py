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
        test = pd.read_excel(path_config + 'Config_test.xlsx', sheet_name='Test')
        cmd = pd.read_excel(path_config + 'Config_test.xlsx', sheet_name='Comandi')
        # camera_climatica = pd.read_excel(path_config + 'Config_test.xlsx', sheet_name='Camera Climatica')
        # arduino = pd.read_excel(path_config + 'Config_test.xlsx', sheet_name='Arduino')

        # time_sample = pd.read_excel(path_config + 'Config.xlsx', sheet_name='Strumenti')['TEMPO CAMPIONAMENTO'][0]
        # test_time = pd.read_excel(path_config + 'Config.xlsx', sheet_name='Strumenti')['TEST TIME'][0]

        # class TestApp(Frame):
        #     def __init__(self, parent=None):
        #         self.parent = parent
        #         Frame.__init__(self)
        #         self.main = self.master
        #         # w, h = self.main.winfo_screenwidth(), self.main.winfo_screenheight()
        #         self.main.geometry("1800x950+0+0")
        #         # self.main.attributes('-fullscreen', True)
        #         self.main.title('Table Data')
        #         f = Frame(self.main)
        #         f.pack(fill=BOTH, expand=1)
        #         # df = TableModel.getSampleData()
        #         namefile = pd.read_excel(r'.\Config\Config.xlsx', sheet_name='Strumenti')['NOME OUTPUT'][0]
        #         if str(namefile) == 'nan':
        #             namefile = 'Data'
        #         if str(pd.read_excel(r'.\Config\Config.xlsx', sheet_name='Strumenti')['PERCORSO OUTPUT'][0]) == 'nan':
        #             path_save = ''
        #         else:
        #             path_save = pd.read_excel(r'.\Config\Config.xlsx', sheet_name='Strumenti')['PERCORSO OUTPUT'][0]
        #         self.table = pt = Table(f,
        #                                 dataframe=pd.read_csv(path_save + namefile + '.csv'),
        #                                 showtoolbar=True, showstatusbar=True)
        #         pt.show()
        #         #self.main.after(20000, lambda: self.main.destroy())
        #         return
        #
        #
        # class MyWindow:
        #     def __init__(self, win, width):
        #         self.win = win
        #         self.b1 = Button(win, text='Raw Data', command=self.data, width=width, bg='lightgreen')
        #         self.b2 = Button(win, text='Grafico', command=self.graph, width=width, bg='lightblue')
        #         self.b1.place(x=80, y=25)
        #         self.b2.place(x=270, y=25)
        #
        #     def data(self):
        #         TestApp()
        #
        #     def graph(self):
        #         newWindow = Toplevel(self.win)
        #         newWindow.title("AITA - Graph Config")
        #         newWindow.geometry("1300x750")
        #         yscrollbar = Scrollbar(newWindow)
        #         yscrollbar.pack(side=RIGHT, fill=Y)
        #
        #         # label = Label(newWindow,
        #         #               text="Seleziona cosa vuoi graficare e come:",
        #         #               font=("Times New Roman", 10),
        #         #               padx=10, pady=10)
        #         # label.pack()
        #         self.listbox = Listbox(newWindow, selectmode="multiple",
        #                        yscrollcommand=yscrollbar.set)
        #
        #         # # Widget expands horizontally and
        #         # # vertically by assigning both to
        #         # # fill option
        #         self.listbox.pack(#padx=10, pady=10,
        #                    expand=YES, fill="y"
        #                    )
        #         self.listbox.place(x=10, y=80)
        #         namefile = pd.read_excel(r'.\Config\Config.xlsx', sheet_name='Strumenti')['NOME OUTPUT'][0]
        #         if str(namefile) == 'nan':
        #             namefile = 'Data'
        #         if str(pd.read_excel(r'.\Config\Config.xlsx', sheet_name='Strumenti')['PERCORSO OUTPUT'][0]) == 'nan':
        #             path_save = ''
        #         else:
        #             path_save = pd.read_excel(r'.\Config\Config.xlsx', sheet_name='Strumenti')['PERCORSO OUTPUT'][0]
        #
        #         x = pd.read_csv(path_save + namefile + '.csv').columns
        #         x = x.drop('Date')
        #
        #         for each_item in range(len(x)):
        #             self.listbox.insert(END, x[each_item])
        #             self.listbox.itemconfig(each_item, bg="#f5f6f7")
        #
        #         # Attach self.listbox to vertical scrollbar
        #         yscrollbar.config(command=self.listbox.yview)
        #         self.b1 = Button(newWindow, text='Grafico', command=self.print_d, width=20, bg='lightgreen')
        #         self.b1.place(x=10, y=10)
        #         self.lbl7 = Label(newWindow, text='Asse di riferimento', bg="#f5f6f7")
        #         self.t12 = Text(newWindow, bd=3, height=10, width=20)
        #         self.lbl7.place(x=200, y=10)
        #         self.t12.place(x=200, y=80)
        #
        #     def print_d(self):
        #         value = self.t12.get("1.0", 'end-1c').split('\n')
        #         namefile = pd.read_excel(r'.\Config\Config.xlsx', sheet_name='Strumenti')['NOME OUTPUT'][0]
        #         if str(namefile) == 'nan':
        #             namefile = 'Data'
        #         if str(pd.read_excel(r'.\Config\Config.xlsx', sheet_name='Strumenti')['PERCORSO OUTPUT'][0]) == 'nan':
        #             path_save = ''
        #         else:
        #             path_save = pd.read_excel(r'.\Config\Config.xlsx', sheet_name='Strumenti')['PERCORSO OUTPUT'][0]
        #         plot_df = pd.read_csv(path_save + namefile + '.csv')
        #         # time_refresh = pd.read_excel(path_config + 'Config.xlsx', sheet_name='Grafico')['REFRESH TIME'][0]
        #         plot = list()
        #         for i in range(0, len(pd.read_excel(path_config + 'Config.xlsx', sheet_name='Grafico')['PLOT'])):
        #             if str(pd.read_excel(path_config + 'Config.xlsx', sheet_name='Grafico')['PLOT'][i]) != 'nan' and str(
        #                     pd.read_excel(path_config + 'Config.xlsx', sheet_name='Grafico')['ASSE'][i]) != 'nan':
        #                 plot.append([pd.read_excel(path_config + 'Config.xlsx', sheet_name='Grafico')['PLOT'][i],
        #                              pd.read_excel(path_config + 'Config.xlsx', sheet_name='Grafico')['ASSE'][i]])
        #         fig = go.Figure()
        #         fig.update_layout(
        #             template='simple_white',
        #             legend=dict(orientation="v", x=1.1, y=1),
        #             xaxis=dict(
        #                 domain=[0.2, 0.8]
        #             ),
        #             yaxis=dict(
        #                 title='Temperature [°C]',
        #                 titlefont=dict(
        #                     color="#1f77b4"
        #                 ),
        #                 tickfont=dict(
        #                     color="#1f77b4"
        #                 )
        #             ),
        #             yaxis3=dict(
        #                 title='Power [W]',
        #                 titlefont=dict(
        #                     color="#EC00FF"
        #                 ),
        #                 tickfont=dict(
        #                     color="#EC00FF"
        #                 ),
        #                 anchor="free",
        #                 overlaying="y",
        #                 side="left",
        #                 position=0.13
        #             ),
        #             yaxis4=dict(
        #                 title='Text',
        #                 titlefont=dict(
        #                     color="#d62728"
        #                 ),
        #                 tickfont=dict(
        #                     color="#d62728"
        #                 ),
        #                 anchor="x",
        #                 overlaying="y",
        #                 side="right"
        #             ),
        #             yaxis2=dict(
        #                 title='Voltage [V]',
        #                 titlefont=dict(
        #                     color="#000000"
        #                 ),
        #                 tickfont=dict(
        #                     color="#000000"
        #                 ),
        #                 anchor="free",
        #                 overlaying="y",
        #                 side="right",
        #                 position=0.95
        #             )
        #         )
        #         list_plot=list()
        #         for i in self.listbox.curselection():
        #             list_plot.append(self.listbox.get(i))
        #         for i in range(len(list_plot)):
        #             if value[i] == 'V':
        #                 fig.add_scatter(x=plot_df['Date'], y=plot_df[list_plot[i]], mode='lines', yaxis='y2',
        #                                 name=list_plot[i])
        #             elif value[i] == 'P':
        #                 fig.add_scatter(x=plot_df['Date'], y=plot_df[list_plot[i]], mode='lines', yaxis='y3',
        #                                 name=list_plot[i])
        #             elif value[i] == 'T':
        #                 fig.add_scatter(x=plot_df['Date'], y=plot_df[list_plot[i]], mode='lines', yaxis='y',
        #                                 name=list_plot[i])
        #             else:
        #                 fig.add_scatter(x=plot_df['Date'], y=plot_df[list_plot[i]], mode='lines', yaxis='y4',
        #                                 name=list_plot[i])
        #         fig.show()
        #
        #
        # def main_grid():
        #     width = 20
        #     window3 = Tk()
        #     window3.title('AITA - Run Time Log')
        #     window3.geometry("500x100")
        #     MyWindow(window3, width)
        #     # b1 = Button(window2, text='Raw Data', command=data(window2), width=width, bg='lightgreen')
        #     # b2 = Button(window2, text='Grafico', command=graph(window2), width=width, bg='lightblue')
        #     # b1.place(x=100, y=25)
        #     # b2.place(x=270, y=25)
        #     window3.mainloop()
        #     # while True:
        #     #     app = TestApp()
        #     #     app.mainloop()
        #     #     if 'Grafico' in device:
        #     #          app_w = TestGraph()
        #     #          app_w.mainloop()
        #
        if test['TEST'][0] != 'Manuale':
            a = True
        else:
            a = False
        while a == True:
            for i in range(0, len(cmd['TEMPO'])):
                # ora gestisco i relay
                for j in range(0, 8, 1):
                    if str(cmd['ARDU_RELAY' + str(j+1)][i]) not in ['nan', '2', '2.0']:

                        cmd_relay = int(100 + cmd['ARDU_RELAY' + str(j+1)][i])
                        num_relay = int(j+1)
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
                            '1 ' + str(cmd_relay) + ' ' + str(num_relay) + ' ' + str(var1) + ' ' + str(var2) + ' ' + str(
                                1 + cmd_relay + num_relay + var1 + var2), 'utf-8'))
                        time.sleep(1)

                        data = arduino_cmd.readline()
                        print(data)
                        #if (len(data) > 16) and ('CRC' not in str(data)) and ('nan' not in str(data)) and i != 0 and j in [1, 5]:# and (cmd['EVENTO'][i] == 1):
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
                    if str(cmd['ALIM_ON/OFF'][i]) not in ['2', 'nan', '2.0'] and str(test['ALIMENTATORE_PORTA'][j]) != 'nan':
                        print(test['ALIMENTATORE_PORTA'][j])
                        session = requests.Session()
                        rm = pyvisa.ResourceManager()
                        alim_open = rm.open_resource(test['ALIMENTATORE_PORTA'][j])
                        if 'Regaton' in alim_open.query("*IDN?"):
                            alim = Regatron(alim_open)
                        elif 'N8957APV' in alim_open.query("*IDN?"):
                            alim = Keysight(alim_open)
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
                    if str(cmd['INV_VAR'][j]) not in ['2', 'nan', '2.0'] and str(test['INV_VAR'][j]) != 'nan':
                        OrionProtocol.write(test['INV_VAR'][j], test['INV_VAL'][j])
                # ora gestisco le colonnine
                for j in range(0, len(test['COL_PORTA'])):
                    telemetry_col, reg, com_colonna, addresses = config_colonnina(path_config)
                    if str(cmd['COL_VAR'][j]) not in ['2', 'nan', '2.0'] and str(test['COL_VAR'][j]) != 'nan':
                        for k in addresses:
                            WriteCol(test['COL_VAR'][j], com_colonna, test['COL_VAL'][j], k)
                # ora gestisco la camera cliamtica
                for j in range(0, len(test['CC_PORTA'])):
                    if str(cmd['CC_ON/OFF'][j]) not in ['2', 'nan', '2.0'] and str(test['CC_ON/OFF'][j]) != 'nan':
                        test = pd.read_excel(path_config + 'Config_test.xlsx', sheet_name='Test')
                        cc = Weiss(com=test['CC_PORTA'])
                        cc.set_temp_hum(cmd['CC_set point T'][j], cmd['CC_set point H'][j], cmd['CC_ON/OFF'][j])

                time.sleep(cmd['TEMPO'][i])

            if test['CICLO'][0] == 'Yes':
                a = True
            else:
                a = False
    except:
        path_watchdog = r'S:\@Solar\Reliability Laboratory\0_Stazioni di Test\10_AITA\0_misc/watchdog.txt'
        file_object = open(path_watchdog, "W")
        file_object.write('2')
        file_object.close()
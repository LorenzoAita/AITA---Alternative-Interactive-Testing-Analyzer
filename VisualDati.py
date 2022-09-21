import time
from tkinter import *
import os
from pymodbus.client.sync import ModbusTcpClient as ModbusClientTCP
from pymodbus.client.sync import ModbusSerialClient as ModbusClient
import pyvisa
from pandastable import Table  # , TableModel
from lib.Libs import *
from html2image import Html2Image
from PIL import ImageTk, Image

path_config = r'./Config/'
bg = "#f5f6f7"
width = 20
device = list(pd.read_excel(path_config + 'Config.xlsx', sheet_name='Strumenti')['ELENCO STRUMENTI'])


class TestApp(Frame):
    def __init__(self, parent=None):
        self.parent = parent
        Frame.__init__(self)
        self.main = self.master
        # w, h = self.main.winfo_screenwidth(), self.main.winfo_screenheight()
        self.main.geometry("1800x950+0+0")
        # self.main.attributes('-fullscreen', True)
        self.main.title('Table Data')
        f = Frame(self.main)
        f.pack(fill=BOTH, expand=1)
        # df = TableModel.getSampleData()
        namefile = pd.read_excel(r'.\Config\Config.xlsx', sheet_name='Strumenti')['NOME OUTPUT'][0]
        if str(namefile) == 'nan':
            namefile = 'Data'
        if str(pd.read_excel(r'.\Config\Config.xlsx', sheet_name='Strumenti')['PERCORSO OUTPUT'][0]) == 'nan':
            path_save = 'data/'
        else:
            path_save = pd.read_excel(r'.\Config\Config.xlsx', sheet_name='Strumenti')['PERCORSO OUTPUT'][0]
        self.table = pt = Table(f,
                                dataframe=pd.read_csv(path_save + namefile + '.csv'),
                                showtoolbar=True, showstatusbar=True)
        pt.show()
        # self.main.after(20000, lambda: self.main.destroy())
        return


class MyWindow:
    def __init__(self, win, width, value):
        self.value = value
        self.win = win
        self.b1 = Button(win, text='Raw Data', command=self.data, width=width, bg='#FFDB58')
        self.b2 = Button(win, text='Grafico', command=self.graph, width=width, bg='#FFDB58')
        self.b1.place(x=80, y=25)
        self.b2.place(x=270, y=25)

        self.b3 = Button(win, text='Camera Climatica', command=self.cc, width=width, bg='lightblue')
        self.b3.place(x=80, y=75)

        self.b6 = Button(win, text='DC Supply', command=self.dc_source, width=width, bg='lightblue')
        self.b6.place(x=270, y=75)

        self.b7 = Button(win, text='Orion Inverter', command=self.eut_control, width=width, bg='lightblue')
        self.b7.place(x=80, y=125)

        self.b8 = Button(win, text='Colonnina', command=self.ac_ctrl, width=width, bg='lightblue')
        self.b8.place(x=270, y=125)

        self.b17 = Button(win, text='AC Source', command=self.ac_source, width=width, bg='lightblue')
        self.b17.place(x=80, y=175)

        self.b19 = Button(win, text='Quit', command=self.exit, width=width, bg='red')
        self.b19.place(x=270, y=175)

        self.my_label = '0'
        self.ids_eut = '0'
        self.ips_col = 0

    def exit(self):
        os.system('taskkill /F /IM python.exe')

    def data(self):
        TestApp()

    def graph(self):
        newWindow = Toplevel(self.win)
        newWindow = newWindow
        newWindow.title("AITA - Graph Config")
        newWindow.geometry("630x840")
        newWindow.iconbitmap('img/logo_hEC_icon.ico')
        yscrollbar = Scrollbar(newWindow)
        yscrollbar.pack(side=RIGHT, fill=Y)

        self.listbox = Listbox(newWindow, selectmode="multiple",
                               yscrollcommand=yscrollbar.set, width=70, height=35)
        self.listbox.pack(padx=10, pady=10,
                          expand=YES, fill="y"
                          )
        self.listbox.place(x=10, y=80)
        namefile = pd.read_excel(r'.\Config\Config.xlsx', sheet_name='Strumenti')['NOME OUTPUT'][0]
        if str(namefile) == 'nan':
            namefile = 'Data'
        if str(pd.read_excel(r'.\Config\Config.xlsx', sheet_name='Strumenti')['PERCORSO OUTPUT'][0]) == 'nan':
            path_save = 'data/'
        else:
            path_save = pd.read_excel(r'.\Config\Config.xlsx', sheet_name='Strumenti')['PERCORSO OUTPUT'][0]

        x = pd.read_csv(path_save + namefile + '.csv').columns
        x = x.drop('Date')
        for each_item in range(len(x)):
            self.listbox.insert(END, x[each_item])
            self.listbox.itemconfig(each_item, bg="#f5f6f7")

        yscrollbar.config(command=self.listbox.yview)
        self.b4 = Button(newWindow, text='Grafico', command=self.print_d, width=20, bg='lightgreen')
        self.b4.place(x=10, y=20)

    def print_d(self):
        global window_graph

        if window_graph:
            window_graph.destroy()

        window_graph = Toplevel(self.win)
        window_graph.title("AITA - Graph")
        window_graph.geometry('%dx%d+%d+%d' % (1100, 740, 500, 100))
        window_graph.iconbitmap('img/logo_hEC_icon.ico')
        value = self.value
        namefile = pd.read_excel(r'.\Config\Config.xlsx', sheet_name='Strumenti')['NOME OUTPUT'][0]
        if str(namefile) == 'nan':
            namefile = 'Data'
        if str(pd.read_excel(r'.\Config\Config.xlsx', sheet_name='Strumenti')['PERCORSO OUTPUT'][0]) == 'nan':
            path_save = 'data/'
        else:
            path_save = pd.read_excel(r'.\Config\Config.xlsx', sheet_name='Strumenti')['PERCORSO OUTPUT'][0]

        plot_df = pd.read_csv(path_save + namefile + '.csv')

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
            ),
            showlegend=True
        )
        list_plot = list()
        for i in self.listbox.curselection():
            list_plot.append(self.listbox.get(i))
        for i in range(len(value)):
            for j in list_plot:
                if value[i][0] in j:
                    if value[i][1] in ['V', 'VDC', 'VAC', 'URMS', 'UDC', 'UAC', 'U', 'UMN', 'UPP', 'UMP']:
                        fig.add_scatter(x=plot_df['Date'], y=plot_df[j], mode='lines', yaxis='y2',
                                        name=j)
                    elif value[i][1] in ['P', 'S', 'Q', 'W', 'VA', 'VAR']:
                        fig.add_scatter(x=plot_df['Date'], y=plot_df[j], mode='lines', yaxis='y3',
                                        name=j)
                    elif value[i][1] in ['T', 'K']:
                        fig.add_scatter(x=plot_df['Date'], y=plot_df[j], mode='lines', yaxis='y',
                                        name=j)
                    else:
                        fig.add_scatter(x=plot_df['Date'], y=plot_df[j], mode='lines', yaxis='y4',
                                        name=j)
        # Conversione dei file da html a png e spostato nel punto di interesse
        fig.write_html('img/GraphData.html')
        hti = Html2Image()
        hti.screenshot(
            html_file='./img/GraphData.html', save_as='GraphData.png', size=[(1080, 720)]
        )
        os.replace("GraphData.png", "img/GraphData.png")
        time.sleep(0.1)
        # aggiunta dell'img nel tkinter
        try:
            img = Image.open('img/GraphData.png')
        except:
            img = Image.open('img/ATP.png')
        self.my_img = ImageTk.PhotoImage(img)
        self.my_label = Label(window_graph, image=self.my_img)
        self.my_label.pack()
        self.my_label.place(x=10, y=10)
        window_graph.after(10000, self.print_d)

    def cc(self):  # ____________PANNELLINO CAMERA CLIMATICA!_______________
        newWindow = Toplevel(self.win)
        newWindow.title("AITA - Command Climatic Chamber")
        newWindow.geometry("750x200")
        newWindow.iconbitmap('img/logo_hEC_icon.ico')
        self.lbl1 = Label(newWindow, text='Modello', bg=bg)
        self.lbl2 = Label(newWindow, text='Porta', bg=bg)
        self.lbl5 = Label(newWindow, text='Temperature', bg=bg)
        self.lbl3 = Label(newWindow, text='Humidity', bg=bg)
        self.lbl4 = Label(newWindow, text='ON/OFF', bg=bg)

        OPTIONS = ["Weiss", "Angelantoni", "Endurance"]
        self.WT = StringVar(newWindow)
        self.WT.set(OPTIONS[0])  # default value
        self.t1 = OptionMenu(newWindow, self.WT, *OPTIONS)
        self.t1.pack()  # tipologia camera

        self.t2 = Entry(newWindow, bd=3)  # com
        self.t3 = Entry(newWindow, bd=3)  # Umidità
        self.t4 = Entry(newWindow, bd=3)  # On/Off
        self.t5 = Entry(newWindow, bd=3)  # Temperature

        self.lbl1.place(x=100, y=30)
        self.t1.place(x=100, y=60)
        self.lbl2.place(x=250, y=30)
        self.t2.place(x=250, y=70)

        self.lbl5.place(x=100, y=100)
        self.t5.place(x=100, y=130)
        self.lbl3.place(x=300, y=100)
        self.t3.place(x=300, y=130)
        self.lbl4.place(x=500, y=100)
        self.t4.place(x=500, y=130)

        self.b1 = Button(newWindow, text='Send', command=self.send_cc, width=width, bg='lightgreen')
        self.b1.place(x=500, y=50)

    def dc_source(self):
        dc_source = Toplevel(self.win)
        dc_source.title("AITA - Command DC Source")
        dc_source.geometry("750x200")
        dc_source.iconbitmap('img/logo_hEC_icon.ico')
        self.lbl1 = Label(dc_source, text='Modello')
        self.lbl2 = Label(dc_source, text='Porta')
        self.lbl3 = Label(dc_source, text='Voltage')
        self.lbl4 = Label(dc_source, text='Current')
        self.lbl5 = Label(dc_source, text='Power')
        self.lbl6 = Label(dc_source, text='Output')

        OPTIONS = ["Regatron", "Keysight", "Lambda", "Sorrensen"]
        self.DC = StringVar(dc_source)
        self.DC.set(OPTIONS[0])  # default value
        self.t1 = OptionMenu(dc_source, self.DC, *OPTIONS)
        self.t1.pack()  # tipologia dc supply

        self.t2 = Entry(dc_source, bd=3)  # com
        self.t3 = Entry(dc_source, width=10, bd=3)  # voltage
        self.t4 = Entry(dc_source, width=10, bd=3)  # current
        self.t5 = Entry(dc_source, width=10, bd=3)  # power

        self.lbl1.place(x=100, y=30)
        self.t1.place(x=100, y=60)
        self.lbl2.place(x=250, y=30)
        self.t2.place(x=250, y=60)

        self.lbl3.place(x=100, y=100)
        self.t3.place(x=100, y=130)
        self.lbl4.place(x=200, y=100)
        self.t4.place(x=200, y=130)
        self.lbl5.place(x=300, y=100)
        self.t5.place(x=300, y=130)
        self.lbl6.place(x=520, y=100)

        self.b1 = Button(dc_source, text='Send', command=self.send_dc_source, width=width, bg='lightgreen')
        self.b1.place(x=470, y=45)
        self.b2 = Button(dc_source, text='ON', command=self.output_dc_source_on, width=10, bg='red')
        self.b2.place(x=450, y=130)
        self.b3 = Button(dc_source, text='OFF', command=self.output_dc_source_off, width=10, bg='lightgreen')
        self.b3.place(x=550, y=130)

    def ac_source(self):
        ac_source = Toplevel(self.win)
        ac_source.title("AITA - Command AC Source")
        ac_source.geometry("750x200")
        self.lbl1 = Label(ac_source, text='Modello')
        self.lbl2 = Label(ac_source, text='Porta')
        self.lbl3 = Label(ac_source, text='Voltage')
        self.lbl4 = Label(ac_source, text='Frequency')
        self.lbl5 = Label(ac_source, text='Regenerative')
        self.lbl6 = Label(ac_source, text='Output')

        OPTIONS = ["RS90"]
        self.AC = StringVar(ac_source)
        self.AC.set(OPTIONS[0])  # default value
        self.t1 = OptionMenu(ac_source, self.AC, *OPTIONS)
        self.t1.pack()  # tipologia ac supply

        self.t2 = Entry(ac_source, bd=3)  # com
        self.t3 = Entry(ac_source, width=10, bd=3)  # voltage
        self.t4 = Entry(ac_source, width=10, bd=3)  # frequency
        OPTIONS = ["ON", "OFF"]
        self.AC_reg = StringVar(ac_source)
        self.AC_reg.set(OPTIONS[0])  # default value
        self.t5 = OptionMenu(ac_source, self.AC_reg, *OPTIONS)
        self.t5.pack()  # tipologia ac supply  # regenerative

        self.lbl1.place(x=100, y=30)
        self.t1.place(x=100, y=60)
        self.lbl2.place(x=250, y=30)
        self.t2.place(x=250, y=60)

        self.lbl3.place(x=100, y=100)
        self.t3.place(x=100, y=130)
        self.lbl4.place(x=200, y=100)
        self.t4.place(x=200, y=130)
        self.lbl5.place(x=300, y=100)
        self.t5.place(x=300, y=130)

        # self.get_ac_regen()

        self.lbl6.place(x=520, y=100)

        self.b1 = Button(ac_source, text='Send', command=self.send_ac_source, width=width, bg='lightgreen')
        self.b1.place(x=470, y=45)
        self.b2 = Button(ac_source, text='ON', command=self.output_ac_source_on, width=10, bg='red')
        self.b2.place(x=450, y=130)
        self.b3 = Button(ac_source, text='OFF', command=self.output_ac_source_off, width=10, bg='lightgreen')

    def send_ac_source(self):
        value = [self.AC.get(), self.t2.get(), self.t3.get(),
                 self.t4.get(), self.AC_reg.get()]

        session = requests.Session()
        rm = pyvisa.ResourceManager()
        alim_open = rm.open_resource(value[1])
        if self.AC.get() == 'RS90':
            ac_source = RS90(alim_open)

        ac_source.set_volt(value[2])
        ac_source.set_freq(value[3])
        if value[4] in ['ON', 'On', 'on']:
            value[4] = 1
        if value[4] in ['OFF', 'Off', 'off']:
            value[4] = 0
        ac_source.set_regenerative(value[4])
        # self.get_ac_regen()

    def output_ac_source_on(self):
        value = 1
        session = requests.Session()
        rm = pyvisa.ResourceManager()
        alim_open = rm.open_resource(self.t2.get())
        if self.AC.get() == 'RS90':
            ac_source = RS90(alim_open)

        ac_source.set_output(value)

    def output_ac_source_off(self):
        value = 0
        session = requests.Session()
        rm = pyvisa.ResourceManager()
        alim_open = rm.open_resource(self.t2.get())
        if self.AC.get() == 'RS90':
            ac_source = RS90(alim_open)

        ac_source.set_output(value)

    def eut_control(self):
        eut = Toplevel(self.win)
        eut.title("AITA - Command EUT")
        eut.geometry("750x200")
        eut.iconbitmap('img/logo_hEC_icon.ico')
        self.lbl0 = Label(eut, text='Inverter')
        self.lbl1 = Label(eut, text='Attribute')
        self.lbl2 = Label(eut, text='Value')

        OPTIONS_eut = list()
        for i in list(pd.read_excel(path_config + 'Config.xlsx', sheet_name='Inverter')['ID']):
            if str(i) != 'nan':
                OPTIONS_eut.append(i)
        if len(OPTIONS_eut)>1:
            OPTIONS_eut.append('ALL')
        self.EUT = StringVar(eut)
        self.EUT.set(OPTIONS_eut[0])  # default value
        self.t0 = OptionMenu(eut, self.EUT, *OPTIONS_eut)
        self.t0.pack()  # quale inverter mettere

        self.ids_eut = OPTIONS_eut

        self.t1 = Entry(eut, width=60, bd=3)  # attribute
        self.t2 = Entry(eut, bd=3)  # value

        self.lbl0.place(x=10, y=30)
        self.t0.place(x=10, y=50)
        self.lbl1.place(x=150, y=30)
        self.t1.place(x=150, y=60)
        self.lbl2.place(x=550, y=30)
        self.t2.place(x=550, y=60)

        self.b1 = Button(eut, text='Send', width=width, command=self.send_inv, bg='lightgreen')
        self.b1.place(x=300, y=90)

    def ac_ctrl(self):
        win_col = Toplevel(self.win)
        win_col.title("AITA - Command AC Station")
        win_col.geometry("750x200")
        win_col.iconbitmap('img/logo_hEC_icon.ico')
        self.lbl0 = Label(win_col, text='Colonnina')
        self.lbl1 = Label(win_col, text='Registro')
        self.lbl2 = Label(win_col, text='Value')

        OPTIONS_col = list()
        for i in list(pd.read_excel(path_config + 'Config.xlsx', sheet_name='Colonnina')['IP']):
            if str(i) != 'nan':
                OPTIONS_col.append(i)
        if len(OPTIONS_col) > 1:
            OPTIONS_col.append('ALL')
        self.COL = StringVar(win_col)
        self.COL.set(OPTIONS_col[0])  # default value
        self.t0 = OptionMenu(win_col, self.COL, *OPTIONS_col)
        self.t0.pack()  # quale inverter mettere

        self.ips_col = OPTIONS_col

        self.t1 = Entry(win_col, width=60, bd=3)  # attribute
        self.t2 = Entry(win_col, bd=3)  # value

        self.lbl0.place(x=10, y=30)
        self.t0.place(x=10, y=50)
        self.lbl1.place(x=150, y=30)
        self.t1.place(x=150, y=60)
        self.lbl2.place(x=550, y=30)
        self.t2.place(x=550, y=60)

        self.b1 = Button(win_col, text='Send', width=width, command=self.send_col, bg='lightgreen')
        self.b1.place(x=300, y=90)

    def send_cc(self):
        value = [self.WT.get(), self.t2.get(), self.t3.get(),
                 self.t4.get(), self.t5.get()]
        if value[0] == 'Weiss':
            camera_climatica = Weiss(com=value[1])
        elif value[0] == 'Angelantoni':
            camera_climatica = Discovery(com=value[1])
        elif value[0] == 'Endurance':
            camera_climatica = Endurance(com=value[1])

        camera_climatica.set_temp_hum(value[4], value[2], value[3])

    def send_dc_source(self):
        value = [self.DC.get(), self.t2.get(), self.t3.get(),
                 self.t4.get()]
        session = requests.Session()
        rm = pyvisa.ResourceManager()
        alim_open = rm.open_resource(value[1])
        if value[0] == 'Regatron':
            dc_source = Regatron(alim_open)
        elif value[0] == 'Keysight':
            dc_source = Keysight(alim_open)
        elif value[0] == 'Lambda':
            dc_source = Lambda(alim_open)
        elif value[0] == 'Sorrensen':
            dc_source = Sorrensen(alim_open)

        dc_source.voltage(value[2])
        dc_source.current(value[3])
        dc_source.power(value[4])

    def output_dc_source_on(self):
        value = 1
        session = requests.Session()
        rm = pyvisa.ResourceManager()
        alim_open = rm.open_resource(self.t2.get())
        if self.DC.get() == 'Regatron':
            dc_source = Regatron(alim_open)
        elif self.DC.get() == 'Keysight':
            dc_source = Keysight(alim_open)
        elif self.DC.get() == 'Lambda':
            dc_source = Lambda(alim_open)
        elif self.DC.get() == 'Sorrensen':
            dc_source = Sorrensen(alim_open)

        dc_source.stato(value)

    def output_dc_source_off(self):
        value = 0
        session = requests.Session()
        rm = pyvisa.ResourceManager()
        alim_open = rm.open_resource(self.t2.get())
        if self.DC.get() == 'Regatron':
            dc_source = Regatron(alim_open)
        elif self.DC.get() == 'Keysight':
            dc_source = Keysight(alim_open)
        elif self.DC.get() == 'Lambda':
            dc_source = Lambda(alim_open)
        elif self.DC.get() == 'Sorrensen':
            dc_source = Sorrensen(alim_open)

        dc_source.stato(value)

    def send_inv(self):
        ids = self.ids_eut
        value = [self.EUT.get(), self.t1.get(), self.t2.get()]
        if value[0] != 'ALL':
            df_eut = pd.read_excel(path_config + 'Config.xlsx', sheet_name='Inverter')
            ip = list(df_eut['IP DEVICE'][df_eut['ID'] == value[0]])[0]
            obj_com = OrionProtocol(ip=value[0], device_id=ip)
            obj_com.write(value[1], value[2])
        else:
            for i in ids:
                df_eut = pd.read_excel(path_config + 'Config.xlsx', sheet_name='Inverter')
                ip = df_eut['IP DEVICE'][df_eut['ID'] == i]
                obj_com = OrionProtocol(ip=value[0], device_id=ip)
                obj_com.write(value[1], value[2])
        del df_eut

    def send_col(self):
        ids = self.ips_col
        value = [self.COL.get(), self.t1.get(), self.t2.get()]
        if value[0] != 'ALL':
            df_eut = pd.read_excel(path_config + 'Config.xlsx', sheet_name='Colonnina')
            ip = str(df_eut['IP'][df_eut['IP']==value[0]][0])
            if df_eut['MODALITA\''][df_eut['IP']==value[0]][0] == 'RTU':
                baud_rate = 115200
                com = ModbusClient(method='rtu',
                                   port=ip,
                                   baudrate=baud_rate,
                                   timeout=1,
                                   parity='N',
                                   stopbits=1,
                                   strict=False)
            elif df_eut['MODALITA\''][df_eut['IP']==value[0]][0] == 'TCP':
                port = int(df_eut['PORTA'][df_eut['IP']==value[0]][0])
                com = ModbusClientTCP(host=ip, port=int(port))

            WriteCol(value[1], com, value[2], df_eut['ADDRESS'][df_eut['IP']==value[0]][0])

        else:
            for i in ids:
                df_eut = pd.read_excel(path_config + 'Config.xlsx', sheet_name='Colonnina')
                ip = str(df_eut['IP'][df_eut['IP'] == i][0])
                if df_eut['MODALITA\''][df_eut['IP'] == i][0] == 'RTU':
                    baud_rate = 115200
                    com = ModbusClient(method='rtu',
                                       port=ip,
                                       baudrate=baud_rate,
                                       timeout=1,
                                       parity='N',
                                       stopbits=1,
                                       strict=False)
                elif df_eut['MODALITA\''][df_eut['IP'] == i][0] == 'TCP':
                    port = int(df_eut['PORTA'][df_eut['IP'] == i][0])
                    com = ModbusClientTCP(host=ip, port=int(port))

                WriteCol(value[1], com, value[2], df_eut['ADDRESS'][df_eut['IP'] == i][0])
        del df_eut


def main_grid():
    try:
        col = list(pd.read_excel(r'.\Config\Config.xlsx', sheet_name='Strumenti')['ELENCO STRUMENTI'])
        value = list()
        if 'Datalogger' in col:
            strumento = pd.read_excel(r'.\Config\Config.xlsx', sheet_name='Agilent')
            for i in range(0, len(strumento['LABEL'])):
                value.append([strumento['LABEL'][i], strumento['TIPOLOGIA'][i]])
        if 'Wattmeter' in col:
            strumento = pd.read_excel(r'.\Config\Config.xlsx', sheet_name='Wattmeter')
            for i in range(0, len(strumento['LABEL'])):
                value.append([strumento['LABEL'][i], strumento['TIPOLOGIA'][i]])
        if 'Inverter' in col:
            strumento = pd.read_excel(r'.\Config\Config.xlsx', sheet_name='Inverter')
            for i in range(0, len(strumento['LABEL'])):
                value.append([strumento['LABEL'][i], strumento['TYPE TELEMETRIA'][i]])
        if 'Colonnina' in col:
            strumento = pd.read_excel(r'.\Config\Config.xlsx', sheet_name='Colonnina')
            for i in range(0, len(strumento['TELEMETRIE'])):
                value.append([strumento['TELEMETRIE'][i], strumento['TYPE TELEMETRIA'][i]])
        try:
            del strumento, col
        except:
            print('nessuno strumento')
        width = 20
        window2 = Tk()
        window2.title('AITA - Run Time Log')
        window2.geometry("500x250")
        window2.iconbitmap('img/logo_hEC_icon.ico')
        MyWindow(window2, width, value)
        window2.mainloop()

    except OSError as err:
        print(err)
        path_watchdog = 'misc/watchdog.txt'
        file_object = open(path_watchdog, "w")
        file_object.write('3')
        file_object.close()


window_graph = None
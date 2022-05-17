from tkinter import *
import os
import subprocess
import pandas as pd


# funzioni per  il PLC
class MyWindow:
    def __init__(self, win):
        # self.lbl1 = Label(win, text='Config. Agilent', bg=bg)
        # self.lbl2 = Label(win, text='Config. Inverter', bg=bg)
        # self.lbl3 = Label(win, text='Config. Wattmeter2', bg=bg)
        # self.lbl4 = Label(win, text='Config. Inverter', bg=bg)
        # self.lbl5 = Label(win, text='Config. Wattmeter', bg=bg)
        # self.lbl6 = Label(win, text='Config. Wattmeter2', bg=bg)
        # self.lbl7 = Label(win, text='Config. Bridge', bg=bg)
        # self.lbl8 = Label(win, text='Config. Colonnina', bg=bg)
        # self.lbl9 = Label(win, text='', bg=bg)
        # self.t1 = Entry(bd=3)
        # self.t2 = Entry(bd=3)
        # self.t3 = Entry(bd=3)
        # self.t4 = Entry(bd=3)
        # self.t5 = Entry(bd=3)
        # self.t6 = Entry(bd=3)
        # self.t7 = Entry(bd=3)
        # self.t8 = Entry(bd=3)
        # self.t9 = Entry(bd=3, state='disabled')
        # self.t10 = Entry(bd=3, state='disabled')
        # self.t11 = Entry(bd=3, state='disabled')
        # self.btn1 = Button(win, text='Chiusura Interruttori')
        # self.btn2 = Button(win, text='Set V DC')
        # self.btn3 = Button(win, text='Set Tamb')
        # self.btn4 = Button(win, text='Read V DC')
        # self.btn5 = Button(win, text='Read Tamb')
        # self.btn6 = Button(win, text='Read Setpoint Tamb')
        self.b7 = Button(win, text='START TEST', command=self.start, width=width, bg='lightgreen')
        # self.b8 = Button(win, text='UTA OFF', command=self.spengi_UTA, width=width, bg='red')
        # variac
        # self.lbl1.place(x=100, y=30)
        # self.t1.place(x=100, y=70)
        # self.lbl7.place(x=250, y=30)
        # self.t7.place(x=200, y=70)
        # interruttori
        # self.lbl5.place(x=100, y=100)
        # self.t5.place(x=100, y=140)
        # self.lbl6.place(x=250, y=100)
        # self.t6.place(x=200, y=140)
        # uta
        # self.lbl2.place(x=100, y=170)
        # self.t2.place(x=100, y=210)
        # self.lbl8.place(x=250, y=170)
        # self.t8.place(x=200, y=210)
        # self.b7.place(x=120, y=250)
        # self.b8.place(x=230, y=250)
        # results
        # self.lbl9.place(x=820, y=40)

        self.b1 = Button(win, text='Config. Wattmeter', command=self.wt1, width=width, bg='#00FFFF')
        self.b2 = Button(win, text='Config. Agilent', command=self.agilent, width=width, bg='#00FFFF')
        self.b3 = Button(win, text='Config. Bridge', command=self.bridge, width=width, bg='#00FFFF')
        self.b4 = Button(win, text='Config. Inverter', command=self.inv, width=width, bg='#00FFFF')
        self.b5 = Button(win, text='Config. Wattmeter2', command=self.wt2, width=width, bg='#00FFFF')
        # self.b8 = Button(win, text='Chiusura Tutti\rInterruttori AC', command=self.cmdplc_all_AC, width=width, bg='#00FFFF')
        self.b6 = Button(win, text='Config. Colonnina', command=self.ACStation, width=width, bg='#00FFFF')
        # self.b7 = Button(win, text='Read setpoint\rTamb', command=self.read_settemp, width=width, bg='#00FFFF')
        # self.b2.bind('<Button-1>', self.set_vdc)
        # interruttori
        self.b1.place(x=100, y=65)
        self.b5.place(x=270, y=65)
        # self.b8.place(x=650, y=125)
        # self.t9.place(x=820, y=140)
        # variac
        self.b2.place(x=100, y=25)
        self.b3.place(x=270, y=25)
        # self.t10.place(x=820, y=70)
        # uta
        self.b4.place(x=100, y=105)
        self.b6.place(x=270, y=105)
        self.b7.place(x=190, y=145)
        # self.t11.place(x=820, y=210)

    def ACStation(self):
        newWindow = Toplevel(window)
        newWindow.title("AC Station Config")
        newWindow.geometry("800x750")

        self.lbl1 = Label(newWindow, text='Ip o COM', bg=bg)
        self.lbl2 = Label(newWindow, text='Porta', bg=bg)
        self.lbl3 = Label(newWindow, text='Modalità', bg=bg)
        self.lbl4 = Label(newWindow, text='Address', bg=bg)
        self.lbl5 = Label(newWindow, text='Telemetrie', bg=bg)
        self.lbl6 = Label(newWindow, text='Registri', bg=bg)

        self.t1 = Entry(newWindow, bd=3)
        self.t2 = Entry(newWindow, bd=3)
        self.t3 = Entry(newWindow, bd=3)
        self.t4 = Entry(newWindow, bd=3)
        self.t5 = Text(newWindow, height=20, width=30, bd=3)
        self.t6 = Text(newWindow, height=20, width=10, bd=3)

        self.lbl1.place(x=100, y=30)
        self.t1.place(x=100, y=70)
        self.lbl2.place(x=200, y=30)
        self.t2.place(x=200, y=70)
        self.lbl3.place(x=300, y=30)
        self.t3.place(x=300, y=70)
        self.lbl4.place(x=400, y=30)
        self.t4.place(x=400, y=70)

        self.lbl5.place(x=100, y=130)
        self.t5.place(x=100, y=170, height=500)
        self.lbl6.place(x=400, y=130)
        self.t6.place(x=400, y=170, height=500)

        self.b1 = Button(newWindow, text='Config', command=self.send_colonna, width=width, bg='lightgreen')
        self.b1.place(x=190, y=700)

        TEXT1 = 'READ ME:\r\rA meno di daisy chain:\r\r - Address:\r\tWallbox\t  -> address = 1\r\tAC station -> address = 2'
        self.lbl10 = Label(newWindow, text=TEXT1, justify='left', bg=bg)
        self.lbl10.place(x=530, y=170)

        # TEXT2 = '\r\r \rresistenza\rtermocoppia T\rtermocoppia K\rTensione dc\rTensione ' \
        #         'ac\rFrequenza\rCorrente dc (SOLO 34970)\rCorrente ac (SOLO 34970)'
        # self.lbl10 = Label(newWindow, text=TEXT2, justify='left', bg=bg)
        # self.lbl10.place(x=630, y=30)

    def bridge(self):
        newWindow = Toplevel(window)
        newWindow.title("Bridge Config")
        newWindow.geometry("700x1000")

        self.lbl1 = Label(newWindow, text='Tipo di misura', bg=bg)
        self.lbl2 = Label(newWindow, text='Freq min', bg=bg)
        self.lbl3 = Label(newWindow, text='Freq max', bg=bg)
        self.lbl4 = Label(newWindow, text='Sampling Freq', bg=bg)
        self.lbl5 = Label(newWindow, text='Level', bg=bg)
        self.lbl6 = Label(newWindow, text='Meas Time', bg=bg)
        self.lbl9 = Label(newWindow, text='Correction', bg=bg)
        self.lbl8 = Label(newWindow, text='Correction Lenght', bg=bg)
        self.lbl7 = Label(newWindow, text='Porta', bg=bg)

        self.t1 = Entry(newWindow, bd=3)
        self.t2 = Entry(newWindow, bd=3)
        self.t3 = Entry(newWindow, bd=3)
        self.t4 = Entry(newWindow, bd=3)
        self.t5 = Entry(newWindow, bd=3)
        self.t6 = Entry(newWindow, bd=3)
        self.t9 = Text(newWindow, bd=3, height=20, width=30)
        self.t8 = Entry(newWindow, bd=3)
        self.t7 = Entry(newWindow, bd=3)

        self.lbl1.place(x=100, y=30)
        self.t1.place(x=100, y=70)
        self.lbl2.place(x=200, y=30)
        self.t2.place(x=200, y=70)
        self.lbl3.place(x=300, y=30)
        self.t3.place(x=300, y=70)
        self.lbl4.place(x=400, y=30)
        self.t4.place(x=400, y=70)

        self.lbl5.place(x=100, y=130)
        self.t5.place(x=100, y=170)
        self.lbl6.place(x=200, y=130)
        self.t6.place(x=200, y=170)
        self.lbl7.place(x=300, y=130)
        self.t7.place(x=300, y=170)

        self.lbl8.place(x=200, y=230)
        self.t8.place(x=200, y=270)
        self.lbl9.place(x=100, y=230)
        self.t9.place(x=100, y=270, height=100, width=100)

        TEXT1 = 'READ ME:\r\r - Tipo di misura:\r\tCPD\r\tCPQ \r\tCPG \r\tCPRP \r\tCSD \r\tCSQ \r\tCSRS \r\tLPD ' \
                '\r\tLPQ \r\tLPG \r\tLPRP \r\tLPRD\r\tLSD \r\tLSQ \r\tLSRS \r\tLSRD \r\tRX \r\tZTD \r\tZTR \r\tGB ' \
                '\r\tYTD \r\tYTR \r\tVDID'
        self.lbl10 = Label(newWindow, text=TEXT1, justify='left', bg=bg)
        self.lbl10.place(x=100, y=360)

        TEXT2 = '\r\r - Level:\r\t[numero] + V/A se il livello è \r\t in tensione o in corrente\r\r - Meas ' \
                'Time:\r\tSHOR\r\tMED\r\tLONG\r\r - Correction:\r\tOPEN\r\tSHOR\r\tLOAD\r\tOFF\r\r - Correction ' \
                'Lenght:\r\t0\r\t1\r\t2\r\t4 '
        self.lbl10 = Label(newWindow, text=TEXT2, justify='left', bg=bg)
        self.lbl10.place(x=230, y=360)

        self.b1 = Button(newWindow, text='Config', command=self.send_bridge, width=width, bg='lightgreen')
        self.b1.place(x=190, y=900)

    def wt1(self):
        newWindow = Toplevel(window)
        newWindow.title("Wattmeter Config")
        newWindow.geometry("900x800")

        self.lbl1 = Label(newWindow, text='Modello', bg=bg)
        self.lbl2 = Label(newWindow, text='Porta', bg=bg)
        self.lbl3 = Label(newWindow, text='Canale', bg=bg)
        self.lbl4 = Label(newWindow, text='Misura', bg=bg)
        self.lbl5 = Label(newWindow, text='Label', bg=bg)

        OPTIONS = [
            "WT230",
            "WT500",
            "WT3000"
        ]
        self.WT = StringVar(newWindow)
        self.WT.set(OPTIONS[0])  # default value
        self.t1 = OptionMenu(newWindow, self.WT, *OPTIONS)
        self.t1.pack()

        self.t2 = Entry(newWindow, bd=3)
        self.t3 = Text(newWindow, bd=3, height=20, width=30)
        self.t4 = Text(newWindow, bd=3, height=20, width=30)
        self.t5 = Text(newWindow, bd=3, height=20, width=30)

        self.lbl1.place(x=100, y=30)
        self.t1.place(x=100, y=60)
        self.lbl2.place(x=250, y=30)
        self.t2.place(x=250, y=70)

        self.lbl3.place(x=100, y=100)
        self.t3.place(x=100, y=130, height=600, width=150)
        self.lbl4.place(x=250, y=100)
        self.t4.place(x=250, y=130, height=600, width=150)
        self.lbl5.place(x=400, y=100)
        self.t5.place(x=400, y=130, height=600, width=150)

        TEXT1 = 'READ ME:\r\r - Misura WT500:\r\tURMS\r\tUMN\r\tUDC\r\tUAC\r\tIRMS\r\tIMN\r\tIDC\r\tIRMN\r\tIAC\r\tP' \
                '\r\tS\r\tQ\r\tLAMB\r\tPHI\r\tFU\r\tFI\r\tUPP\r\tUMP\r\tIPP\r\tIMP\r\tWH\r\tWHP\r\tWHM\r\tAH\r\tAHP\r' \
                '\tAHM\r\tTIME '
        self.lbl10 = Label(newWindow, text=TEXT1, justify='left', bg=bg)
        self.lbl10.place(x=560, y=30)

        TEXT2 = '\r\r - Misura WT230:\r\tV\r\tA\r\tW\r\tVA\r\tVAR\r\tPF\r\tDEGR\r\tVHZ\r\tAHZ\r\tWH\r\tWHP\r\tWHM\r' \
                '\tAH\r\tAHP\r\tAHM\r\tTIME\r\r - Misura ' \
                'WT3000:\r\tU\r\tI\r\tP\r\tS\r\tQ\r\tLAMB\r\tIPP\r\tIMP\r\tWH\r\tWHP\r\tWHM\r\tAH\r\tAHP\r\tAHM\r' \
                '\tTIME '
        self.lbl10 = Label(newWindow, text=TEXT2, justify='left', bg=bg)
        self.lbl10.place(x=700, y=30)

        self.b1 = Button(newWindow, text='Config', command=self.send_wt1, width=width, bg='lightgreen')
        self.b1.place(x=190, y=750)

    def wt2(self):
        newWindow = Toplevel(window)
        newWindow.title("Wattmeter Config")
        newWindow.geometry("900x800")

        self.lbl1 = Label(newWindow, text='Modello', bg=bg)
        self.lbl2 = Label(newWindow, text='Porta', bg=bg)
        self.lbl3 = Label(newWindow, text='Canale', bg=bg)
        self.lbl4 = Label(newWindow, text='Misura', bg=bg)
        self.lbl5 = Label(newWindow, text='Label', bg=bg)

        OPTIONS = [
            "WT230",
            "WT500",
            "WT3000"
        ]
        self.WT = StringVar(newWindow)
        self.WT.set(OPTIONS[0])  # default value
        self.t1 = OptionMenu(newWindow, self.WT, *OPTIONS)
        self.t1.pack()

        self.t2 = Entry(newWindow, bd=3)
        self.t3 = Text(newWindow, bd=3, height=20, width=30)
        self.t4 = Text(newWindow, bd=3, height=20, width=30)
        self.t5 = Text(newWindow, bd=3, height=20, width=30)

        self.lbl1.place(x=100, y=30)
        self.t1.place(x=100, y=60)
        self.lbl2.place(x=250, y=30)
        self.t2.place(x=250, y=70)

        self.lbl3.place(x=100, y=100)
        self.t3.place(x=100, y=130, height=600, width=150)
        self.lbl4.place(x=250, y=100)
        self.t4.place(x=250, y=130, height=600, width=150)
        self.lbl5.place(x=400, y=100)
        self.t5.place(x=400, y=130, height=600, width=150)

        TEXT1 = 'READ ME:\r\r - Misura WT500:\r\tURMS\r\tUMN\r\tUDC\r\tUAC\r\tIRMS\r\tIMN\r\tIDC\r\tIRMN\r\tIAC\r\tP' \
                '\r\tS\r\tQ\r\tLAMB\r\tPHI\r\tFU\r\tFI\r\tUPP\r\tUMP\r\tIPP\r\tIMP\r\tWH\r\tWHP\r\tWHM\r\tAH\r\tAHP\r' \
                '\tAHM\r\tTIME '
        self.lbl10 = Label(newWindow, text=TEXT1, justify='left', bg=bg)
        self.lbl10.place(x=560, y=30)

        TEXT2 = '\r\r - Misura WT230:\r\tV\r\tA\r\tW\r\tVA\r\tVAR\r\tPF\r\tDEGR\r\tVHZ\r\tAHZ\r\tWH\r\tWHP\r\tWHM\r' \
                '\tAH\r\tAHP\r\tAHM\r\tTIME\r\r - Misura ' \
                'WT3000:\r\tU\r\tI\r\tP\r\tS\r\tQ\r\tLAMB\r\tIPP\r\tIMP\r\tWH\r\tWHP\r\tWHM\r\tAH\r\tAHP\r\tAHM\r' \
                '\tTIME '
        self.lbl10 = Label(newWindow, text=TEXT2, justify='left', bg=bg)
        self.lbl10.place(x=700, y=30)

        self.b1 = Button(newWindow, text='Config', command=self.send_wt2, width=width, bg='lightgreen')
        self.b1.place(x=190, y=750)

    def inv(self):
        newWindow = Toplevel(window)
        newWindow.title("Orion Inverter Config")
        newWindow.geometry("600x750")

        self.lbl1 = Label(newWindow, text='Ip o COM', bg=bg)
        self.lbl2 = Label(newWindow, text='Seriale', bg=bg)
        self.lbl5 = Label(newWindow, text='Telemetrie', bg=bg)

        self.t1 = Entry(newWindow, bd=3)
        self.t2 = Entry(newWindow, bd=3)
        self.t5 = Text(newWindow, height=20, width=30, bd=3)

        self.lbl1.place(x=100, y=30)
        self.t1.place(x=100, y=70)
        self.lbl2.place(x=200, y=30)
        self.t2.place(x=200, y=70)

        self.lbl5.place(x=100, y=130)
        self.t5.place(x=100, y=170, height=500)

        self.b1 = Button(newWindow, text='Config', command=self.send_inv, width=width, bg='lightgreen')
        self.b1.place(x=190, y=700)

    def agilent(self):
        newWindow = Toplevel(window)
        newWindow.title("Datalogger Config")
        newWindow.geometry("850x800")

        self.lbl1 = Label(newWindow, text='Modello', bg=bg)
        self.lbl2 = Label(newWindow, text='Porta', bg=bg)
        self.lbl3 = Label(newWindow, text='Misura', bg=bg)
        self.lbl4 = Label(newWindow, text='Label', bg=bg)

        OPTIONS = ["34970", "34980"]
        WT = StringVar(newWindow)
        WT.set(OPTIONS[0])  # default value
        self.t1 = OptionMenu(newWindow, WT, *OPTIONS)
        self.t1.pack()

        self.t2 = Entry(newWindow, bd=3)
        self.t3 = Text(newWindow, bd=3, height=20, width=30)
        self.t4 = Text(newWindow, bd=3, height=20, width=30)

        self.lbl1.place(x=100, y=30)
        self.t1.place(x=100, y=60)
        self.lbl2.place(x=250, y=30)
        self.t2.place(x=250, y=70)

        self.lbl3.place(x=100, y=100)
        self.t3.place(x=100, y=130, height=600, width=200)
        self.lbl4.place(x=300, y=100)
        self.t4.place(x=300, y=130, height=600, width=200)

        TEXT1 = 'READ ME:\r\r - Misura Agilent:\r\tOHM\r\tT\r\tK\r\tVDC\r\tVAC\r\tHZ\r\tIDC\r\tIAC'
        self.lbl10 = Label(newWindow, text=TEXT1, justify='left', bg=bg)
        self.lbl10.place(x=500, y=30)

        TEXT2 = '\r\r \rresistenza\rtermocoppia T\rtermocoppia K\rTensione dc\rTensione ' \
                'ac\rFrequenza\rCorrente dc (SOLO 34970)\rCorrente ac (SOLO 34970)'
        self.lbl10 = Label(newWindow, text=TEXT2, justify='left', bg=bg)
        self.lbl10.place(x=630, y=30)

        self.b1 = Button(newWindow, text='Config', command=self.send_agilent, width=width, bg='lightgreen')
        self.b1.place(x=500, y=300)

    def send_colonna(self):
        excel_config = pd.DataFrame()  # .read_excel(path_config + 'Config.xlsx', sheet_name='Colonnina')
        # nome = ['Ip o COM', 'Porta', 'Modalità', 'Address', 'Telemetrie', 'Registri']
        value = [self.t1.get(), self.t2.get(), self.t3.get(), self.t4.get(),
                 list(self.t5.get("1.0", 'end-1c').split('\n')),
                 list(self.t6.get("1.0", 'end-1c').split('\n'))]
        list_app = list()
        list_app.append(value[0])
        excel_config['IP'] = list_app
        df1 = pd.DataFrame()
        list_app = list()
        list_app.append(value[1])
        df1['PORTA'] = list_app
        excel_config = pd.concat([excel_config, df1], axis=1)
        df1 = pd.DataFrame()
        list_app = list()
        list_app.append(value[2])
        df1['MODALITA\''] = list_app
        excel_config = pd.concat([excel_config, df1], axis=1)
        df1 = pd.DataFrame()
        list_app = list()
        list_app.append(value[3])
        df1['ADDRESS'] = list_app
        excel_config = pd.concat([excel_config, df1], axis=1)
        df1 = pd.DataFrame()
        df1['TELEMETRIE'] = value[4]
        excel_config = pd.concat([excel_config, df1], axis=1)
        df1 = pd.DataFrame()
        df1['REGISTRO'] = value[5]
        excel_config = pd.concat([excel_config, df1], axis=1)
        del df1
        excel_config.to_excel(writer, sheet_name='Colonnina', index=False)
        writer.save()

    def send_inv(self):
        excel_config = pd.DataFrame()  # .read_excel(path_config + 'Config.xlsx', sheet_name='Colonnina')
        # nome = ['Ip o COM', 'Porta', 'Modalità', 'Address', 'Telemetrie', 'Registri']
        value = [self.t1.get(), self.t2.get(), list(self.t5.get("1.0", 'end-1c').split('\n'))]
        excel_config['IP DEVICE'] = list(value[0])
        df1 = pd.DataFrame()
        df1['ID'] = list(value[1])
        excel_config = pd.concat([excel_config, df1], axis=1)
        df1 = pd.DataFrame()
        df1['TELEMETRIE'] = list(value[2])
        excel_config = pd.concat([excel_config, df1], axis=1)
        del df1
        excel_config.to_excel(writer, sheet_name='Inverter', index=False)
        writer.save()

    def send_bridge(self):
        excel_config = pd.DataFrame()  # .read_excel(path_config + 'Config.xlsx', sheet_name='Colonnina')
        # nome = ['Ip o COM', 'Porta', 'Modalità', 'Address', 'Telemetrie', 'Registri']
        value = [str(self.t1.get()), self.t2.get(), self.t3.get(), self.t4.get(), str(self.t5.get()),
                 str(self.t6.get()), str(self.t7.get()), self.t8.get(), list(self.t9.get("1.0", 'end-1c').split('\n'))]
        list_freq = list()
        list_freq.append(value[1])
        list_freq.append(value[2])
        list_freq.append(value[3])
        list_corr = list()
        for i in range(0, len(value[8])):
            list_corr.append(value[8][i])
        list_app = list()
        list_app.append(value[0])
        excel_config['TIPOLOGIA'] = list_app
        df1 = pd.DataFrame()
        list_app = list()
        list_app.append(value[4])
        df1['LEVEL'] = list_app
        excel_config = pd.concat([excel_config, df1], axis=1)
        df1 = pd.DataFrame()
        list_app = list()
        list_app.append(value[5])
        df1['MEAS TIME'] = list_app
        excel_config = pd.concat([excel_config, df1], axis=1)
        df1 = pd.DataFrame()
        list_app = list()
        list_app.append(value[7])
        df1['CORRECTION LENGTH'] = list_app
        excel_config = pd.concat([excel_config, df1], axis=1)
        df1 = pd.DataFrame()
        df1['FREQUENZA'] = list_freq
        excel_config = pd.concat([excel_config, df1], axis=1)
        df1 = pd.DataFrame()
        df1['CORRECTION'] = list_corr
        excel_config = pd.concat([excel_config, df1], axis=1)
        df1 = pd.DataFrame()
        list_app = list()
        list_app.append(value[6])
        df1['PORTA BRIDGE'] = list_app
        excel_config = pd.concat([excel_config, df1], axis=1)
        del df1, list_app
        excel_config.to_excel(writer, sheet_name='Bridge', index=False)
        writer.save()

    def send_wt1(self):
        excel_config = pd.DataFrame()  # .read_excel(path_config + 'Config.xlsx', sheet_name='Colonnina')
        # nome = ['Ip o COM', 'Porta', 'Modalità', 'Address', 'Telemetrie', 'Registri']
        value = [self.WT.get(), self.t2.get(), self.t3.get("1.0", 'end-1c').split('\n'),
                 self.t4.get("1.0", 'end-1c').split('\n')]
        excel_config['PORTA WT'] = list(value[1])
        df1 = pd.DataFrame()
        df1['MODELLO'] = list(value[0])
        excel_config = pd.concat([excel_config, df1], axis=1)
        df1 = pd.DataFrame()
        df1['CHANNEL'] = value[2]
        excel_config = pd.concat([excel_config, df1], axis=1)
        df1 = pd.DataFrame()
        df1['TIPOLOGIA'] = value[3]
        excel_config = pd.concat([excel_config, df1], axis=1)
        df1 = pd.DataFrame()
        df1['LABEL'] = value[4]
        excel_config = pd.concat([excel_config, df1], axis=1)
        del df1
        excel_config.to_excel(writer, sheet_name='Wattmeter', index=False)
        writer.save()

    def send_wt2(self):
        excel_config = pd.DataFrame()
        value = [self.WT.get(), self.t2.get(), self.t3.get("1.0", 'end-1c').split('\n'),
                 self.t4.get("1.0", 'end-1c').split('\n')]
        excel_config['PORTA WT'] = list(value[1])
        df1 = pd.DataFrame()
        df1['MODELLO'] = list(value[0])
        excel_config = pd.concat([excel_config, df1], axis=1)
        df1 = pd.DataFrame()
        df1['CHANNEL'] = value[2]
        excel_config = pd.concat([excel_config, df1], axis=1)
        df1 = pd.DataFrame()
        df1['TIPOLOGIA'] = value[3]
        excel_config = pd.concat([excel_config, df1], axis=1)
        df1 = pd.DataFrame()
        df1['LABEL'] = value[4]
        excel_config = pd.concat([excel_config, df1], axis=1)
        del df1
        excel_config.to_excel(writer, sheet_name='Wattmeter2', index=False)
        writer.save()

    def send_agilent(self):
        excel_config = pd.DataFrame()
        value = ['', self.t2.get(), self.t3.get("1.0", 'end-1c').split('\n'), self.t4.get("1.0", 'end-1c').split('\n')]
        list_34970 = list()
        for i in [101, 102, 103, 104, 105, 106, 107, 108, 109, 110, 112, 111, 113, 114, 115,
                  116, 117, 118, 119, 120, 121, 122, 201, 202, 203, 204, 205, 206, 207, 208,
                  209, 210, 211, 212, 213, 214, 215, 216, 217, 218, 219, 220, 221, 222, 301,
                  302, 303, 304, 305, 306, 307, 308, 309, 310, 311, 312, 313, 314, 315, 316,
                  317, 318, 319, 320, 321, 322]:
            list_34970.append(i)
        list_34980 = list()
        for i in [1001, 1002, 1003, 1004, 1005, 1006, 1007, 1008, 1009, 1010, 1011, 1012, 1013,
                                 1014, 1015, 1016, 1017, 1018, 1019, 1020, 1021, 1022, 1023, 1024, 1025, 1026,
                                 1027, 1028, 1029, 1030, 1031, 1032, 1033, 1034, 1035, 1036, 1037, 1038, 1039,
                                 1040, 2001, 2002, 2003, 2004, 2005, 2006, 2007, 2008, 2009, 2010, 2011, 2012,
                                 2013, 2014, 2015, 2016, 2017, 2018, 2019, 2020, 2021, 2022, 2023, 2024, 2025,
                                 2026, 2027, 2028, 2029, 2030, 2031, 2032, 2033, 2034, 2035, 2036, 2037, 2038,
                                 2039, 2040, 3001, 3002, 3003, 3004, 3005, 3006, 3007, 3008, 3009, 3010, 3011,
                                 3012, 3013, 3014, 3015, 3016, 3017, 3018, 3019, 3020, 3021, 3022, 3023, 3024,
                                 3025, 3026, 3027, 3028, 3029, 3030, 3031, 3032, 3033, 3034, 3035, 3036, 3037,
                                 3038, 3039, 3040, 4001, 4002, 4003, 4004, 4005, 4006, 4007, 4008, 4009, 4010,
                                 4011, 4012, 4013, 4014, 4015, 4016, 4017, 4018, 4019, 4020, 4021, 4022, 4023,
                                 4024, 4025, 4026, 4027, 4028, 4029, 4030, 4031, 4032, 4033, 4034, 4035, 4036,
                                 4037, 4038, 4039, 4040, 5001, 5002, 5003, 5004, 5005, 5006, 5007, 5008, 5009,
                                 5010, 5011, 5012, 5013, 5014, 5015, 5016, 5017, 5018, 5019, 5020, 5021, 5022,
                                 5023, 5024, 5025, 5026, 5027, 5028, 5029, 5030, 5031, 5032, 5033, 5034, 5035,
                                 5036, 5037, 5038, 5039, 5040, 6001, 6002, 6003, 6004, 6005, 6006, 6007, 6008,
                                 6009, 6010, 6011, 6012, 6013, 6014, 6015, 6016, 6017, 6018, 6019, 6020, 6021,
                                 6022, 6023, 6024, 6025, 6026, 6027, 6028, 6029, 6030, 6031, 6032, 6033, 6034,
                                 6035, 6036, 6037, 6038, 6039, 6040, 7001, 7002, 7003, 7004, 7005, 7006, 7007,
                                 7008, 7009, 7010, 7011, 7012, 7013, 7014, 7015, 7016, 7017, 7018, 7019, 7020,
                                 7021, 7022, 7023, 7024, 7025, 7026, 7027, 7028, 7029, 7030, 7031, 7032, 7033,
                                 7034, 7035, 7036, 7037, 7038, 7039, 7040, 8001, 8002, 8003, 8004, 8005, 8006,
                                 8007, 8008, 8009, 8010, 8011, 8012, 8013, 8014, 8015, 8016, 8017, 8018, 8019,
                                 8020, 8021, 8022, 8023, 8024, 8025, 8026, 8027, 8028, 8029, 8030, 8031, 8032,
                                 8033, 8034, 8035, 8036, 8037, 8038, 8039, 8040]:
            list_34980.append(i)
        excel_config['CASSETTI 34970'] = list_34970
        df1 = pd.DataFrame()
        df1['CASSETTI 34980'] = list_34980
        excel_config = pd.concat([excel_config, df1], axis=1)
        df1 = pd.DataFrame()
        df1['PORTA DAQ'] = list(value[1])
        excel_config = pd.concat([excel_config, df1], axis=1)
        df1 = pd.DataFrame()
        df1['TIPOLOGIA'] = value[2]
        excel_config = pd.concat([excel_config, df1], axis=1)
        df1 = pd.DataFrame()
        df1['LABEL'] = value[3]
        excel_config = pd.concat([excel_config, df1], axis=1)
        del df1
        excel_config.to_excel(writer, sheet_name='Agilent', index=False)
        writer.save()

    def start(self):
        newWindow = Toplevel(window)
        newWindow.title("Test Config")
        newWindow.geometry("1100x350")

        self.t1 = IntVar()
        c1 = Checkbutton(newWindow, text='Bridge', variable=self.t1, onvalue=1, offvalue=0)
        c1.pack()
        self.t2 = IntVar()
        c2 = Checkbutton(newWindow, text='DataLogger', variable=self.t2, onvalue=1, offvalue=0)
        c2.pack()
        self.t3 = IntVar()
        c3 = Checkbutton(newWindow, text='First Wattmeter', variable=self.t3, onvalue=1, offvalue=0)
        c3.pack()
        self.t4 = IntVar()
        c4 = Checkbutton(newWindow, text='Second Wattmeter', variable=self.t4, onvalue=1, offvalue=0)
        c4.pack()
        self.t5 = IntVar()
        c5 = Checkbutton(newWindow, text='Orion Inverter', variable=self.t5, onvalue=1, offvalue=0)
        c5.pack()
        self.t6 = IntVar()
        c6 = Checkbutton(newWindow, text='Colonnina', variable=self.t6, onvalue=1, offvalue=0)
        c6.pack()
        self.tg = IntVar()
        c7 = Checkbutton(newWindow, text='Grafico Real Time', variable=self.tg, onvalue=1, offvalue=0)
        c7.pack()
        self.lbl1 = Label(newWindow, text='Strumenti', bg=bg)
        self.lbl1.place(x=50, y=20)
        c1.place(x=50, y=50)
        c2.place(x=50, y=90)
        c3.place(x=50, y=130)
        c4.place(x=50, y=170)
        c5.place(x=50, y=210)
        c6.place(x=50, y=250)
        c7.place(x=50, y=290)

        self.lbl2 = Label(newWindow, text='Percorso Output', bg=bg)
        self.t7 = Text(newWindow, bd=3, height=1, width=16)
        self.lbl2.place(x=250, y=20)
        self.t7.place(x=250, y=50)

        self.lbl3 = Label(newWindow, text='Nome Output', bg=bg)
        self.t8 = Text(newWindow, bd=3, height=1, width=16)
        self.lbl3.place(x=450, y=20)
        self.t8.place(x=450, y=50)

        self.lbl4 = Label(newWindow, text='Tempo Campionamento', bg=bg)
        self.t9 = Entry(newWindow, bd=3)
        self.lbl4.place(x=250, y=120)
        self.t9.place(x=250, y=150)

        self.lbl5 = Label(newWindow, text='Tempo di test', bg=bg)
        self.t10 = Entry(newWindow, bd=3)
        self.lbl5.place(x=450, y=120)
        self.t10.place(x=450, y=150)

        self.lbl6 = Label(newWindow, text='Grandezze da graficare', bg=bg)
        self.t11 = Text(newWindow, bd=3, height=10, width=20)
        self.lbl6.place(x=650, y=20)
        self.t11.place(x=650, y=50)

        self.lbl7 = Label(newWindow, text='Asse di riferimento', bg=bg)
        self.t12 = Text(newWindow, bd=3, height=10, width=20)
        self.lbl7.place(x=870, y=20)
        self.t12.place(x=870, y=50)

        self.lbl8 = Label(newWindow, text='Refresh del grafico', bg=bg)
        self.t13 = Entry(newWindow, bd=3)
        self.lbl8.place(x=650, y=270)
        self.t13.place(x=650, y=300)

        self.b1 = Button(newWindow, text='Start Test', command=self.start_test, width=width, bg='lightgreen')
        self.b1.place(x=250, y=240)
        # self.b2 = Button(newWindow, text='Visual Data', command=self.end_test, width=width, bg='lightblue')
        # self.b2.place(x=450, y=240)

    def start_test(self):
        strumenti = list()
        if self.t1.get() != 0:
            strumenti.append('Bridge')
        if self.t2.get() != 0:
            strumenti.append('Datalogger')
        if self.t3.get() != 0:
            strumenti.append('Wattmeter')
        if self.t4.get() != 0:
            strumenti.append('Wattmeter2')
        if self.t5.get() != 0:
            strumenti.append('Inverter')
        if self.t6.get() != 0:
            strumenti.append('Colonnina')
        if self.tg.get() != 0:
            strumenti.append('Grafico')

        value = [self.t11.get("1.0", 'end-1c').split('\n'), self.t12.get("1.0", 'end-1c').split('\n'), self.t13.get()]
        excel_config = pd.DataFrame()
        list_app = list()
        for i in value[0]:
            list_app.append(i)
        excel_config['PLOT'] = list_app
        df1 = pd.DataFrame()
        list_app = list()
        for i in value[1]:
            list_app.append(i)
        df1['ASSE'] = list_app
        excel_config = pd.concat([excel_config, df1], axis=1)
        df1 = pd.DataFrame()
        list_app = list()
        list_app.append(value[2])
        df1['REFRESH TIME'] = list_app
        excel_config = pd.concat([excel_config, df1], axis=1)
        del df1, list_app
        excel_config.to_excel(writer, sheet_name='Grafico', index=False)
        writer.save()

        value = [self.t7.get("1.0", 'end-1c').split('\n'), self.t8.get("1.0", 'end-1c').split('\n'),
                 self.t9.get(), self.t10.get()]
        excel_config = pd.DataFrame()
        excel_config['ELENCO STRUMENTI'] = strumenti
        df1 = pd.DataFrame()
        list_app = list()
        list_app.append(value[1][0])
        df1['NOME OUTPUT'] = list_app
        excel_config = pd.concat([excel_config, df1], axis=1)
        df1 = pd.DataFrame()
        list_app = list()
        list_app.append(value[2])
        df1['TEMPO CAMPIONAMENTO'] = list_app
        excel_config = pd.concat([excel_config, df1], axis=1)
        df1 = pd.DataFrame()
        list_app = list()
        list_app.append(value[3])
        df1['TEST TIME'] = list_app
        excel_config = pd.concat([excel_config, df1], axis=1)
        df1 = pd.DataFrame()
        list_app = list()
        list_app.append(value[0][0])
        df1['PERCORSO OUTPUT'] = list_app
        excel_config = pd.concat([excel_config, df1], axis=1)
        del df1, list_app
        excel_config.to_excel(writer, sheet_name='Strumenti', index=False)
        writer.save()

        os.system('python LogAITA.py')
        # writer.close()

    def end_test(self):
        os.system('python VisualDati.py')
        #writer.close()
        quit()


# plc = 'ciao'
bg = "#f5f6f7"
title_window = 'Comandi per PLC'

path_config = 'Config/'
if os.path.exists(path_config + 'Config.xlsx'):
    writer = pd.ExcelWriter(path_config + 'Config.xlsx', engine='openpyxl', mode='a', if_sheet_exists='replace')
else:
    writer = pd.ExcelWriter(path_config + 'Config.xlsx', engine='openpyxl')

# height=1
width = 20

window = Tk()
mywin = MyWindow(window)
window.title('AITA')
window.geometry("550x200")
window.mainloop()

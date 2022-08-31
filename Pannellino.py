import os
import threading
from tkinter import *

import pandas as pd

# import GraphDati
import LogAITA
import VisualDati
import TestSeq

exitFlag = 0

class myThread (threading.Thread):
    def __init__(self, threadID, name, counter, run_script):
       threading.Thread.__init__(self)
       self.threadID = threadID
       self.name = name
       self.counter = counter
       self.run_script = run_script

    def run(self):
       #print ("Starting " + self.name)
       if self.run_script == 0:
          LogAITA.main_log()
       if self.run_script == 1:
          TestSeq.main_test()
       if self.run_script == 2:
          VisualDati.main_grid()


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
        self.lbl3 = Label(newWindow, text='Telemetrie', bg=bg)
        self.lbl4 = Label(newWindow, text='Label', bg=bg)

        self.t1 = Entry(newWindow, bd=3)
        self.t2 = Entry(newWindow, bd=3)
        self.t3 = Text(newWindow, height=20, width=30, bd=3)
        self.t4 = Text(newWindow, height=20, width=20, bd=3)

        self.lbl1.place(x=100, y=30)
        self.t1.place(x=100, y=70)
        self.lbl2.place(x=200, y=30)
        self.t2.place(x=200, y=70)

        self.lbl3.place(x=100, y=130)
        self.t3.place(x=100, y=170, height=500)
        self.lbl4.place(x=300, y=130)
        self.t4.place(x=300, y=170, height=500)

        self.b1 = Button(newWindow, text='Config', command=self.send_inv, width=width, bg='lightgreen')
        self.b1.place(x=365, y=65)

    def agilent(self):
        newWindow = Toplevel(window)
        newWindow.title("Datalogger Config")
        newWindow.geometry("1050x800")

        self.lbl1 = Label(newWindow, text='Modello', bg=bg)
        self.lbl2 = Label(newWindow, text='Porta', bg=bg)
        self.lbl5 = Label(newWindow, text='Canale', bg=bg)
        self.lbl3 = Label(newWindow, text='Misura', bg=bg)
        self.lbl4 = Label(newWindow, text='Label', bg=bg)

        OPTIONS = ["34970", "34980"]
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

        self.lbl5.place(x=100, y=100)
        self.t5.place(x=100, y=130, height=600, width=200)
        self.lbl3.place(x=300, y=100)
        self.t3.place(x=300, y=130, height=600, width=200)
        self.lbl4.place(x=500, y=100)
        self.t4.place(x=500, y=130, height=600, width=200)

        TEXT1 = 'READ ME:\r\r - Misura Agilent:\r\tOHM\r\tT\r\tK\r\tVDC\r\tVAC\r\tHZ\r\tIDC\r\tIAC'
        self.lbl10 = Label(newWindow, text=TEXT1, justify='left', bg=bg)
        self.lbl10.place(x=710, y=130)

        TEXT2 = '\r\r \rresistenza\rtermocoppia T\rtermocoppia K\rTensione dc\rTensione ' \
                'ac\rFrequenza\rCorrente dc (SOLO 34970)\rCorrente ac (SOLO 34970)'
        self.lbl10 = Label(newWindow, text=TEXT2, justify='left', bg=bg)
        self.lbl10.place(x=840, y=130)

        self.b1 = Button(newWindow, text='Config', command=self.send_agilent, width=width, bg='lightgreen')
        self.b1.place(x=710, y=400)

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
        value = [self.t1.get(), self.t2.get(), list(self.t3.get("1.0", 'end-1c').split('\n')),
                 list(self.t4.get("1.0", 'end-1c').split('\n'))]
        list_app = list()
        list_app.append(value[0])
        excel_config['IP DEVICE'] = list_app
        df1 = pd.DataFrame()
        list_app = list()
        list_app.append(value[1])
        df1['ID'] = list_app
        excel_config = pd.concat([excel_config, df1], axis=1)
        df1 = pd.DataFrame()
        df1['TELEMETRIE'] = list(value[2])
        excel_config = pd.concat([excel_config, df1], axis=1)
        df1 = pd.DataFrame()
        df1['LABEL'] = list(value[3])
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
                 self.t4.get("1.0", 'end-1c').split('\n'), self.t5.get("1.0", 'end-1c').split('\n')]
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
        value = [self.t5.get("1.0", 'end-1c').split('\n'), self.t2.get(), self.t3.get("1.0", 'end-1c').split('\n'),
                 self.t4.get("1.0", 'end-1c').split('\n'), self.WT.get()]
        list_cassetto = list()
        for i in value[0]:
            list_cassetto.append(i)
        excel_config['CASSETTO '+str(value[4])] = list_cassetto
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
        newWindow.geometry("800x350")

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
        # self.tg = IntVar()
        # c7 = Checkbutton(newWindow, text='Grafico Real Time', variable=self.tg, onvalue=1, offvalue=0)
        # c7.pack()
        self.lbl1 = Label(newWindow, text='Strumenti', bg=bg)
        self.lbl1.place(x=50, y=20)
        c1.place(x=50, y=50)
        c2.place(x=50, y=90)
        c3.place(x=50, y=130)
        c4.place(x=50, y=170)
        c5.place(x=50, y=210)
        c6.place(x=50, y=250)
        # c7.place(x=50, y=290)

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

        # self.lbl6 = Label(newWindow, text='Grandezze da graficare', bg=bg)
        # self.t11 = Text(newWindow, bd=3, height=10, width=20)
        # self.lbl6.place(x=650, y=20)
        # self.t11.place(x=650, y=50)
        #
        # self.lbl7 = Label(newWindow, text='Asse di riferimento', bg=bg)
        # self.t12 = Text(newWindow, bd=3, height=10, width=20)
        # self.lbl7.place(x=870, y=20)
        # self.t12.place(x=870, y=50)
        #
        # self.lbl8 = Label(newWindow, text='Refresh del grafico', bg=bg)
        # self.t13 = Entry(newWindow, bd=3)
        # self.lbl8.place(x=650, y=270)
        # self.t13.place(x=650, y=300)

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
        # if self.tg.get() != 0:
        #     strumenti.append('Grafico')
        #
        # value = [self.t11.get("1.0", 'end-1c').split('\n'), self.t12.get("1.0", 'end-1c').split('\n'), self.t13.get()]
        # excel_config = pd.DataFrame()
        # list_app = list()
        # for i in value[0]:
        #     list_app.append(i)
        # excel_config['PLOT'] = list_app
        # df1 = pd.DataFrame()
        # list_app = list()
        # for i in value[1]:
        #     list_app.append(i)
        # df1['ASSE'] = list_app
        # excel_config = pd.concat([excel_config, df1], axis=1)
        # df1 = pd.DataFrame()
        # list_app = list()
        # list_app.append(value[2])
        # df1['REFRESH TIME'] = list_app
        # excel_config = pd.concat([excel_config, df1], axis=1)
        # del df1, list_app
        # excel_config.to_excel(writer, sheet_name='Grafico', index=False)
        # writer.save()

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
        window.after(100, lambda: window.destroy())

        # Create new threads
        thread1 = myThread(1, "Log", 1, 0)
        thread2 = myThread(2, "RealTime", 2, 2)
        thread3 = myThread(3, "TestSeq", 3, 1)

        # Start new Threads
        thread1.start()
        thread2.start()
        thread3.start()

    def end_test(self):
        # os.system('python VisualDati.py')
        # writer.close()
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
window.title('AITA - Config. Pannel')
window.geometry("550x200")
window.mainloop()

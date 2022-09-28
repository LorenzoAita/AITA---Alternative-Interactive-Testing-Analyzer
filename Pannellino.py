import os
import threading
from tkinter import *

import pandas as pd

# import GraphDati
import LogAITA
import VisualDati
import TestSeq
import Watchdog

exitFlag = 0


class myThread (threading.Thread):
    def __init__(self, threadID, name, counter, run_script):
       threading.Thread.__init__(self)
       self.threadID = threadID
       self.name = name
       self.counter = counter
       self.run_script = run_script

    def run(self):
       if self.run_script == 0:
          LogAITA.main_log()
       if self.run_script == 1:
          TestSeq.main_test()
       if self.run_script == 2:
          VisualDati.main_grid()
       if self.run_script == 3:
          Watchdog.main_wd()


class MyWindow:
    def __init__(self, win):
        newWindow = win
        self.t1 = IntVar()
        c1 = Checkbutton(newWindow, text='Bridge', variable=self.t1, onvalue=1, offvalue=0)
        c1.pack()
        self.t2 = IntVar()
        c2 = Checkbutton(newWindow, text='DataLoggers', variable=self.t2, onvalue=1, offvalue=0)
        c2.pack()
        self.t3 = IntVar()
        c3 = Checkbutton(newWindow, text='Wattmeters', variable=self.t3, onvalue=1, offvalue=0)
        c3.pack()
        self.t5 = IntVar()
        c5 = Checkbutton(newWindow, text='Orion Inverters', variable=self.t5, onvalue=1, offvalue=0)
        c5.pack()
        self.t6 = IntVar()
        c6 = Checkbutton(newWindow, text='AC Stations', variable=self.t6, onvalue=1, offvalue=0)
        c6.pack()
        self.lbl1 = Label(newWindow, text='Strumenti', bg=bg)
        self.lbl1.place(x=50, y=20)
        c1.place(x=50, y=50)
        c2.place(x=50, y=90)
        c3.place(x=50, y=130)
        c6.place(x=50, y=170)
        c5.place(x=50, y=210)

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

        self.b1 = Button(newWindow, text='Start Test', command=self.start_test, width=width, bg='lightgreen')
        self.b1.place(x=250, y=240)

    def start_test(self):
        strumenti = list()
        if self.t1.get() != 0:
            strumenti.append('Bridge')
        if self.t2.get() != 0:
            strumenti.append('Datalogger')
        if self.t3.get() != 0:
            strumenti.append('Wattmeter')
        if self.t5.get() != 0:
            strumenti.append('Inverter')
        if self.t6.get() != 0:
            strumenti.append('Colonnina')

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
        thread4 = myThread(4, "Watchdog", 4, 3)

        # Start new Threads
        #thread4.start()
        thread1.start()
        thread2.start()
        thread3.start()

    def end_test(self):
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
#window.title('AITA - Config. Pannel')
window.title('AITA - Test Config')
#window.geometry("550x200")
window.geometry("700x320")
window.iconbitmap('img/logo_hEC_icon.ico')
window.mainloop()

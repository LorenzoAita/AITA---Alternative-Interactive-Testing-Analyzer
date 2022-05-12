from tkinter import *
import serial
import time
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
        newWindow.geometry("700x750")

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

    def bridge(self):
        newWindow = Toplevel(window)
        newWindow.title("Bridge Agilent E4980A")
        newWindow.geometry("200x200")
        Label(newWindow,
              text="This is a new window").pack()

    def wt1(self):
        # Toplevel object which will
        # be treated as a new window
        newWindow = Toplevel(window)

        # sets the title of the
        # Toplevel widget
        newWindow.title("First Wattmeter Config")

        # sets the geometry of toplevel
        newWindow.geometry("200x200")

        # A Label widget to show in toplevel
        Label(newWindow,
              text="This is a new window").pack()

    def wt2(self):
        # Toplevel object which will
        # be treated as a new window
        newWindow = Toplevel(window)

        # sets the title of the
        # Toplevel widget
        newWindow.title("Second Wattmeter Config")

        # sets the geometry of toplevel
        newWindow.geometry("200x200")

        # A Label widget to show in toplevel
        Label(newWindow,
              text="This is a new window").pack()

    def inv(self):
        # Toplevel object which will
        # be treated as a new window
        newWindow = Toplevel(window)

        # sets the title of the
        # Toplevel widget
        newWindow.title("New Window")

        # sets the geometry of toplevel
        newWindow.geometry("Inverter Config")

        # A Label widget to show in toplevel
        Label(newWindow,
              text="This is a new window").pack()

    def agilent(self):
        # Toplevel object which will
        # be treated as a new window
        newWindow = Toplevel(window)

        # sets the title of the
        # Toplevel widget
        newWindow.title("Agilent Config")

        # sets the geometry of toplevel
        newWindow.geometry("200x200")

        # A Label widget to show in toplevel
        Label(newWindow,
              text="This is a new window").pack()

    def start(self):
        print('gogogo!')

    def send_colonna(self):
        excel_config = pd.DataFrame()#.read_excel(path_config + 'Config.xlsx', sheet_name='Colonnina')
        #nome = ['Ip o COM', 'Porta', 'Modalità', 'Address', 'Telemetrie', 'Registri']
        value = [self.t1.get(), self.t2.get(), self.t3.get(), self.t4.get(), list(self.t5.get("1.0", 'end-1c').split('\n')),
                 list(self.t6.get("1.0", 'end-1c').split('\n'))]
        excel_config['IP'] = value[0]
        excel_config['PORTA'] = value[1]
        excel_config['MODALITA\''] = value[2]
        excel_config['ADDRESS'] = value[3]
        excel_config['READ ME'] = 'wallbox -> address = 1'
        excel_config['TELEMETRIE'] = value[4]
        excel_config['REGISTRO'] = value[5]
        excel_config['IP'][0] = value[0]
        excel_config['PORTA'][0] = value[1]
        excel_config['MODALITA\''][0] = value[2]
        excel_config['ADDRESS'][0] = value[3]
        excel_config['READ ME'][0] = 'wallbox -> address = 1'
        excel_config['READ ME'][1] = 'ac station -> address = 2'

        excel_config.to_excel(path_config + 'Prova.xlsx', sheet_name='Colonnina')

# plc = 'ciao'
bg = "#f5f6f7"
title_window = 'Comandi per PLC'

path_config = 'Config/'
excel_config = pd.read_excel(path_config + 'Config.xlsx')

# height=1
width = 20

window = Tk()
mywin = MyWindow(window)
window.title('AITA')
window.geometry("550x200")
window.mainloop()

from tkinter import *
from pandastable import Table#, TableModel
import pandas as pd
import time
import plotly.graph_objects as go

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
            path_save = ''
        else:
            path_save = pd.read_excel(r'.\Config\Config.xlsx', sheet_name='Strumenti')['PERCORSO OUTPUT'][0]
        self.table = pt = Table(f,
                                dataframe=pd.read_csv(path_save + namefile + '.csv'),
                                showtoolbar=True, showstatusbar=True)
        pt.show()
        #self.main.after(20000, lambda: self.main.destroy())
        return


class MyWindow:
    def __init__(self, win, width):
        self.win = win
        self.b1 = Button(win, text='Raw Data', command=self.data, width=width, bg='lightgreen')
        self.b2 = Button(win, text='Grafico', command=self.graph, width=width, bg='lightblue')
        self.b1.place(x=80, y=25)
        self.b2.place(x=270, y=25)

        self.b3 = Button(win, text='Camera Climatica', command=self.cc, width=width, bg='lightblue')
        self.b3.place(x=80, y=75)

    def data(self):
        TestApp()
    
    def graph(self):
        newWindow = Toplevel(self.win)
        newWindow.title("AITA - Graph Config")
        newWindow.geometry("1300x750")
        yscrollbar = Scrollbar(newWindow)
        yscrollbar.pack(side=RIGHT, fill=Y)

        # label = Label(newWindow,
        #               text="Seleziona cosa vuoi graficare e come:",
        #               font=("Times New Roman", 10),
        #               padx=10, pady=10)
        # label.pack()
        self.listbox = Listbox(newWindow, selectmode="multiple",
                       yscrollcommand=yscrollbar.set)

        # # Widget expands horizontally and
        # # vertically by assigning both to
        # # fill option
        self.listbox.pack(#padx=10, pady=10,
                   expand=YES, fill="y"
                   )
        self.listbox.place(x=10, y=80)
        namefile = pd.read_excel(r'.\Config\Config.xlsx', sheet_name='Strumenti')['NOME OUTPUT'][0]
        if str(namefile) == 'nan':
            namefile = 'Data'
        if str(pd.read_excel(r'.\Config\Config.xlsx', sheet_name='Strumenti')['PERCORSO OUTPUT'][0]) == 'nan':
            path_save = ''
        else:
            path_save = pd.read_excel(r'.\Config\Config.xlsx', sheet_name='Strumenti')['PERCORSO OUTPUT'][0]

        x = pd.read_csv(path_save + namefile + '.csv').columns
        x = x.drop('Date')

        for each_item in range(len(x)):
            self.listbox.insert(END, x[each_item])
            self.listbox.itemconfig(each_item, bg="#f5f6f7")

        # Attach self.listbox to vertical scrollbar
        yscrollbar.config(command=self.listbox.yview)
        self.b1 = Button(newWindow, text='Grafico', command=self.print_d, width=20, bg='lightgreen')
        self.b1.place(x=10, y=10)
        self.lbl7 = Label(newWindow, text='Asse di riferimento', bg="#f5f6f7")
        self.t12 = Text(newWindow, bd=3, height=10, width=20)
        self.lbl7.place(x=200, y=10)
        self.t12.place(x=200, y=80)

    def print_d(self):
        value = self.t12.get("1.0", 'end-1c').split('\n')
        namefile = pd.read_excel(r'.\Config\Config.xlsx', sheet_name='Strumenti')['NOME OUTPUT'][0]
        if str(namefile) == 'nan':
            namefile = 'Data'
        if str(pd.read_excel(r'.\Config\Config.xlsx', sheet_name='Strumenti')['PERCORSO OUTPUT'][0]) == 'nan':
            path_save = ''
        else:
            path_save = pd.read_excel(r'.\Config\Config.xlsx', sheet_name='Strumenti')['PERCORSO OUTPUT'][0]
        plot_df = pd.read_csv(path_save + namefile + '.csv')
        # time_refresh = pd.read_excel(path_config + 'Config.xlsx', sheet_name='Grafico')['REFRESH TIME'][0]
        plot = list()
        for i in range(0, len(pd.read_excel(path_config + 'Config.xlsx', sheet_name='Grafico')['PLOT'])):
            if str(pd.read_excel(path_config + 'Config.xlsx', sheet_name='Grafico')['PLOT'][i]) != 'nan' and str(
                    pd.read_excel(path_config + 'Config.xlsx', sheet_name='Grafico')['ASSE'][i]) != 'nan':
                plot.append([pd.read_excel(path_config + 'Config.xlsx', sheet_name='Grafico')['PLOT'][i],
                             pd.read_excel(path_config + 'Config.xlsx', sheet_name='Grafico')['ASSE'][i]])
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
        list_plot=list()
        for i in self.listbox.curselection():
            list_plot.append(self.listbox.get(i))
        for i in range(len(list_plot)):
            if value[i] == 'V':
                fig.add_scatter(x=plot_df['Date'], y=plot_df[list_plot[i]], mode='lines', yaxis='y2',
                                name=list_plot[i])
            elif value[i] == 'P':
                fig.add_scatter(x=plot_df['Date'], y=plot_df[list_plot[i]], mode='lines', yaxis='y3',
                                name=list_plot[i])
            elif value[i] == 'T':
                fig.add_scatter(x=plot_df['Date'], y=plot_df[list_plot[i]], mode='lines', yaxis='y',
                                name=list_plot[i])
            else:
                fig.add_scatter(x=plot_df['Date'], y=plot_df[list_plot[i]], mode='lines', yaxis='y4',
                                name=list_plot[i])
        fig.show()

    def cc(self):       #____________PANNELLINO CAMERA CLIMATICA!_______________
        newWindow = Toplevel(self.win)
        newWindow.title("AITA - Command Climatic Chamber")
        newWindow.geometry("750x200")
        self.lbl1 = Label(newWindow, text='Modello', bg=bg)
        self.lbl2 = Label(newWindow, text='Porta', bg=bg)
        self.lbl5 = Label(newWindow, text='Temperature', bg=bg)
        self.lbl3 = Label(newWindow, text='Humidity', bg=bg)
        self.lbl4 = Label(newWindow, text='ON/OFF', bg=bg)

        OPTIONS = ["Weiss", "Angelantoni", "Endurance"]
        self.WT = StringVar(newWindow)
        self.WT.set(OPTIONS[0])  # default value
        self.t1 = OptionMenu(newWindow, self.WT, *OPTIONS)
        self.t1.pack() #tipologia camera

        self.t2 = Entry(newWindow, bd=3) #com
        self.t3 = Entry(newWindow, bd=3) #Umidità
        self.t4 = Entry(newWindow, bd=3) #On/Off
        self.t5 = Entry(newWindow, bd=3) #Temperature

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


def main_grid():
    try:
        width = 20
        window2 = Tk()
        window2.title('AITA - Run Time Log')
        window2.geometry("500x200")
        MyWindow(window2, width)
        # b1 = Button(window2, text='Raw Data', command=data(window2), width=width, bg='lightgreen')
        # b2 = Button(window2, text='Grafico', command=graph(window2), width=width, bg='lightblue')
        # b1.place(x=100, y=25)
        # b2.place(x=270, y=25)
        window2.mainloop()
        # while True:
        #     app = TestApp()
        #     app.mainloop()
        #     if 'Grafico' in device:
        #          app_w = TestGraph()
        #          app_w.mainloop()

    except:
        path_watchdog = r'S:\@Solar\Reliability Laboratory\0_Stazioni di Test\10_AITA\0_misc/watchdog.txt'
        file_object = open(path_watchdog, "w")
        file_object.write('3')
        file_object.close()

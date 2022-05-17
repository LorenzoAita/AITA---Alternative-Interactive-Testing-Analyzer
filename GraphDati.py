import tkinter
from tkinter import *
from PIL import Image, ImageTk
from pandastable import Table, TableModel
import pandas as pd
import time
import plotly.graph_objects as go

path_config = r'./Config/'
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
        self.main.after(20000, lambda: self.main.destroy())
        return


class TestGraph(Frame):
    def __init__(self, parent=None):
        self.parent = parent
        Frame.__init__(self)
        self.main = self.master
        self.main.geometry("1800x950+0+0")
        self.main.title('Graph Data')
        f = Frame(self.main)
        f.pack(fill=BOTH, expand=1)
        namefile = pd.read_excel(r'.\Config\Config.xlsx', sheet_name='Strumenti')['NOME OUTPUT'][0]
        if str(namefile) == 'nan':
            namefile = 'Data'
        if str(pd.read_excel(r'.\Config\Config.xlsx', sheet_name='Strumenti')['PERCORSO OUTPUT'][0]) == 'nan':
            path_save = ''
        else:
            path_save = pd.read_excel(r'.\Config\Config.xlsx', sheet_name='Strumenti')['PERCORSO OUTPUT'][0]
        plot_df = pd.read_csv(path_save + namefile + '.csv')
        time_refresh = pd.read_excel(path_config + 'Config.xlsx', sheet_name='Grafico')['REFRESH TIME'][0]
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
        for i in range(0, len(plot)):
            if plot[i][1] == 'W':
                fig.add_scatter(x=plot_df['Date'], y=plot_df[plot[i][0]], mode='lines', yaxis='y4',
                                name=plot[i][0])
            elif plot[i][1] == 'V':
                fig.add_scatter(x=plot_df['Date'], y=plot_df[plot[i][0]], mode='lines', yaxis='y2',
                                name=plot[i][0])
            elif plot[i][1] == 'P':
                fig.add_scatter(x=plot_df['Date'], y=plot_df[plot[i][0]], mode='lines', yaxis='y3',
                                name=plot[i][0])
            elif plot[i][1] == 'T':
                fig.add_scatter(x=plot_df['Date'], y=plot_df[plot[i][0]], mode='lines', yaxis='y',
                                name=plot[i][0])
        fig.write_image(path_save + "\RunTimeGraph.png", width=1750, height=900)
        image1 = Image.open(path_save + "\RunTimeGraph.png")
        test = ImageTk.PhotoImage(image1)
        label1 = tkinter.Label(image=test)
        label1.image = test
        label1.place(x=10, y=10)

        self.main.after(int(time_refresh * 1000), lambda: self.main.destroy())
        return


time.sleep(pd.read_excel(r'.\Config\Config.xlsx', sheet_name='Strumenti')['TEMPO CAMPIONAMENTO'][0] + 10)
while True:
    # app = TestApp()
    # app.mainloop()
    if 'Grafico' in device:
        app_w = TestGraph()
        app_w.mainloop()

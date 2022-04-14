import requests
import plotly.graph_objects as go
from subprocess import Popen


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

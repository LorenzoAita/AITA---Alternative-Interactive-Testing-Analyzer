import time
import os

def main():
    path_watchdog = r'S:\@Solar\Reliability Laboratory\0_Stazioni di Test\10_AITA\0_misc/watchdog.txt'
    file_object = open(path_watchdog, "r")
    while file_object.readline() == '0':
        time.sleep(5)
        print('sono dentro!')
        file_object.close()
        file_object = open(path_watchdog, "r")

    if file_object.readline() == '1':
        print('l\'errore è nel log!')
        os.system('wmic process where "commandline like \'%%LogAITA.py%%\'" delete')
    if file_object.readline() == '2':
        print('l\'errore è nella sequenza di test!')
        os.system('wmic process where "commandline like \'%%TestSeq.py%%\'" delete')
    if file_object.readline() == '3':
        print('l\'errore è nella visualizzazione dei dati!')
        os.system('wmic process where "commandline like \'%%VisualDati.py%%\'" delete')

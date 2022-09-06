import time
import os

def main():
    path_watchdog = r'S:\@Solar\Reliability Laboratory\0_Stazioni di Test\10_AITA\0_misc/watchdog.txt'
    file_object = open(path_watchdog, "r")
    while file_object.readline() in ['0', '4', 'test']:
        time.sleep(5)
        file_object.close()
        file_object = open(path_watchdog, "r")
        if file_object.readline() == '4':
            print('sono in pausa!')

    # stampo l'errore nel thread
    if file_object.readline() == '1':
        print('l\'errore è nel log!')
        # provo a farli ripartire
        os.system('python Watchdog.py LogAita.py')
    if file_object.readline() == '2':
        print('l\'errore è nella sequenza di test!')
        # killo tutti i processi
        os.system('wmic process where "commandline like \'%%LogAITA.py%%\'" delete')
        os.system('wmic process where "commandline like \'%%VisualDati.py%%\'" delete')
        os.system('wmic process where "commandline like \'%%TestSeq.py%%\'" delete')
    if file_object.readline() == '3':
        print('l\'errore è nella visualizzazione dei dati!')
        # provo a farli ripartire
        os.system('python Watchdog.py VisulaDati.py')

main()
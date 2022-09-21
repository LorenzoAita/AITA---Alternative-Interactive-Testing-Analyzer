import time
import os
from Pannellino import myThread

def main():
    path_watchdog = 'misc/watchdog.txt'
    file_object = open(path_watchdog, "w")
    file_object.write('0')
    file_object.close()

    file_object = open(path_watchdog, "r")
    lines = file_object.read()
    while lines.split('\n', 1)[0] in ['0', '4', 'test', '']:
        time.sleep(5)
        print('aspetto')
        file_object.close()
        file_object = open(path_watchdog, "r")
        lines = file_object.read()
        if lines.split('\n', 1)[0] == '4':
            print('sono in pausa!')
        file_object.close()
        file_object = open(path_watchdog, "r")
        lines = file_object.read()

    # os.system('taskkill /F /IM python.exe')

    # stampo l'errore nel thread
    if lines.split('\n', 1)[0] == '1':
        print('l\'errore è nel log!')
        # provo a farli ripartire
        # Create new threads
        thread1 = myThread(1, "Log", 1, 0)
        thread4 = myThread(4, "Watchdog", 4, 3)
        # Start new Threads
        thread4.start()
        thread1.start()

    if lines.split('\n', 1)[0] == '2':
        print('l\'errore è nella sequenza di test!')
        # killo tutti i processi
        os.system('taskkill /F /IM python.exe')
        # os.system('wmic process where "commandline like \'%%LogAITA.py%%\'" delete')
        # os.system('wmic process where "commandline like \'%%VisualDati.py%%\'" delete')
        # os.system('wmic process where "commandline like \'%%TestSeq.py%%\'" delete')

    if lines.split('\n', 1)[0] == '3':
        print('l\'errore è nella visualizzazione dei dati!')
        # provo a farli ripartire
        # Create new threads
        thread2 = myThread(2, "RealTime", 2, 2)
        thread4 = myThread(4, "Watchdog", 4, 3)

        # Start new Threads
        thread4.start()
        thread2.start()


# main()
import time

path_watchdog = r'S:\@Solar\Reliability Laboratory\0_Stazioni di Test\10_AITA\0_misc/watchdog.txt'
file_object = open(path_watchdog, "r")
while file_object.readline() == '0':
    time.sleep(5)
    print('sono dentro!')
    file_object.close()
    file_object = open(path_watchdog, "r")

print('sono uscito!')


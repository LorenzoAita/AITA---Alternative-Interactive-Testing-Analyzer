import serial
import sys
import struct

def get_aurorapsw(eut):
    """ Generazione della PW di secondo livello da protocollo aurora """
    SERVICE_PSW_SEED = [9, 1, 9, 5, 1, 0]  # seed for service-mode psw calculation
    ServicePSW = [0, 0, 0, 0, 0, 0]
    for i in range(0, 6):
        Temp1 = int(eut[i]) - 48
        if Temp1 > 9 or Temp1 < 0:
            Temp1 = 0
        if (i % 2) == 0:
            tmp = Temp1
            tmp += SERVICE_PSW_SEED[i]
            tmp &= 0xFF
            tmp %= 10
            tmp += 48
        else:
            tmp = Temp1
            tmp -= SERVICE_PSW_SEED[i]
            tmp &= 0xFF
            tmp %= 10
            tmp += 48

        ServicePSW[i] = tmp
    return ServicePSW

def calcolaCRC(msg, SM):
    # print(type(msg))
    srL = 255
    srH = 255
    if SM == 0:  # CRC16
        for i in range(0, 8):
            a = (srL ^ int(msg[i])) & 255
            b = ((a << 4) ^ a) & 255
            c = ((b << 3) ^ srH) & 255
            srL = ((b >> 4) ^ c) & 255
            srH = ((b >> 5) ^ b) & 255
    if SM == 1:
        for i in range(0, 8):  # CRC125
            a = (srL ^ int(msg[i])) & 255
            b = ((a << 5) ^ a) & 255
            c = ((b << 4) ^ srH) & 255
            srL = ((b >> 3) ^ c) & 255
            srH = ((b >> 4) ^ b) & 255
    hB = 0xFF & ~(srH & 255)
    lB = 0xFF & ~(srL & 255)

    return hB, lB

def mess():
    num_array = list()
    print('Enter numbers in array: ')
    for i in range(8):
        n = input('num :')
        num_array.append(int(n))
    print('Messaggio: ', num_array)
    return num_array

def cmd485(msg, cH, cL):
    msg.append(cH)
    msg.append(cL)
    print('Messaggio: ', msg)
    return msg

def mantissa(var):
    return int(var/256),var-256*int(var/256)

def four_numbers(val):
    return list(reversed(list(struct.pack('f', val))))

def cmd13X(addr,var,val):

    a, b = mantissa(var)

    msg1=[]
    msg2=[]

    msg1.extend([int(addr), int(135), int(1), int(a), int(b), int(6), int(0), int(67)])
    #print(msg1)

    listing = four_numbers(val)
    msg2.extend([int(addr), int(136), int(200), int(listing[0]), int(listing[1]), int(listing[2]), int(listing[3]), int(0)])
    #print(msg2)

    c1 = calcolaCRC(msg1, 0)
    msg1 = cmd485(msg1, c1[1], c1[0])

    c2 = calcolaCRC(msg2, 0)
    msg2 = cmd485(msg2, c2[1], c2[0])

    return msg1, msg2

def inviacomando(mess):
    conn.write(bytearray(mess))
    read_data = conn.readline()
    risp = list(read_data)
    print('Risposta:', risp)
    return ''

COM = sys.argv[1]  #NON inserire gli apici
addr = int(sys.argv[2]) #indirizzo macchina
state = int(sys.argv[3]) #stato prova: discesa = 0, salita = 1

#################RICHIESTA SERIALE################

conn = serial.Serial(COM, 19200, timeout=1, parity='N')

msgSN = [addr, 63, 0, 0, 0, 0, 0, 0] #mess()
cSN = calcolaCRC(msgSN, 0)
messSN = cmd485(msgSN, cSN[1], cSN[0])


conn.write(bytearray(messSN))
read_data = conn.readline()
eut = list(read_data)                    # eut=serial number array
print('Risposta:', eut)

#eut=[0, 0, 0, 0, 0, 52]

##################################################

#SERVICE MODE CON PASSWORD POWER1

msg125 = [addr, 125, 80, 79, 87, 69, 82, 49] #mess()
c125 = calcolaCRC(msg125, 1)
mess1 = cmd485(msg125, c125[1], c125[0])

#PASSWORD SECONDO LIVELLO CON SERIALE MACCHINA

msgSM = [addr, 84]
psw = get_aurorapsw(eut)
for i in range(0,6):
    msgSM.append(psw[i])
cSM = calcolaCRC(msgSM, 0)
mess2 = cmd485(msgSM, cSM[1], cSM[0])

if state == 0:

    #SELEZIONO IL PARAMETRO DA MODIFICARE E LO MODIFICO

    msg1, msg2 = cmd13X(addr, 4397, -40)
    inviacomando(msg1)
    inviacomando(msg2)


    msg1, msg2 = cmd13X(addr, 4398, -40)
    inviacomando(msg1)
    inviacomando(msg2)

    msg1, msg2 = cmd13X(addr, 4404, -40)
    inviacomando(msg1)
    inviacomando(msg2)

    msg1, msg2 = cmd13X(addr, 4405, -40)
    inviacomando(msg1)
    inviacomando(msg2)

    msg1, msg2 = cmd13X(addr, 4411, -40)
    inviacomando(msg1)
    inviacomando(msg2)

    msg1, msg2 = cmd13X(addr, 4412, -40)
    inviacomando(msg1)
    inviacomando(msg2)

    msg1, msg2 = cmd13X(addr, 4418, -40)
    inviacomando(msg1)
    inviacomando(msg2)

    msg1, msg2 = cmd13X(addr, 4419, -40)
    inviacomando(msg1)
    inviacomando(msg2)

    msg1, msg2 = cmd13X(addr, 4425, -40)
    inviacomando(msg1)
    inviacomando(msg2)

    msg1, msg2 = cmd13X(addr, 4426, -40)
    inviacomando(msg1)
    inviacomando(msg2)

    conn.close()

if state == 1:
    # SELEZIONO IL PARAMETRO DA MODIFICARE E LO MODIFICO

    msg1, msg2 = cmd13X(addr, 4397, 40)
    inviacomando(msg1)
    inviacomando(msg2)

    msg1, msg2 = cmd13X(addr, 4398, 65)
    inviacomando(msg1)
    inviacomando(msg2)

    msg1, msg2 = cmd13X(addr, 4404, 20)
    inviacomando(msg1)
    inviacomando(msg2)

    msg1, msg2 = cmd13X(addr, 4405, 70)
    inviacomando(msg1)
    inviacomando(msg2)

    msg1, msg2 = cmd13X(addr, 4411, 50)
    inviacomando(msg1)
    inviacomando(msg2)

    msg1, msg2 = cmd13X(addr, 4412, 90)
    inviacomando(msg1)
    inviacomando(msg2)

    msg1, msg2 = cmd13X(addr, 4418, 50)
    inviacomando(msg1)
    inviacomando(msg2)

    msg1, msg2 = cmd13X(addr, 4419, 90)
    inviacomando(msg1)
    inviacomando(msg2)

    msg1, msg2 = cmd13X(addr, 4425, 40)
    inviacomando(msg1)
    inviacomando(msg2)

    msg1, msg2 = cmd13X(addr, 4426, 65)
    inviacomando(msg1)
    inviacomando(msg2)

    conn.close()
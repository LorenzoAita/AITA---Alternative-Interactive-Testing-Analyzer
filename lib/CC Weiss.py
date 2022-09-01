import pyvisa
import requests
import time
import serial

cc = serial.Serial(port='COM3', baudrate=9600, parity="N", stopbits=1, timeout=.3, bytesize=8)
cc.write(bytes('$01I\r', 'utf-8'))
time.sleep(6)
cc.write(bytes('$01E 0020.0 0000.0 0100.0 0005.0 0030.0 01000000000000000000000000000000\r', 'utf-8'))
cc.close()
time.sleep(5)

cc = serial.Serial(port='COM3', baudrate=9600, parity="N", stopbits=1, timeout=.3, bytesize=8)
cc.write(bytes('$01I\r', 'utf-8'))
time.sleep(6)
cc.write(bytes('$01E 0020.0 0000.0 0100.0 0005.0 0030.0 00000000000000000000000000000000\r', 'utf-8'))
cc.close()
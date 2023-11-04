import serial
import time 

scaleSer = serial.Serial(port='COM3',baudrate=2400,timeout=None,parity=serial.PARITY_EVEN,stopbits=serial.STOPBITS_ONE,bytesize=serial.SEVENBITS)

if scaleSer.isOpen():
        scaleSer.readline()
        time.sleep(1)
        value = scaleSer.readline()
        qr = value.decode('utf-8')
        print(qr)

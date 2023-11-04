import time 
import serial
import serial.tools.list_ports


class arduino():
        

        def __init__(self):
                self.arduinoState = 0
                self.arduinoConnect()



        def arduinoConnect(self):
                self.ports = list(serial.tools.list_ports.comports())
                for p in self.ports:
                        if "Arduino" in p.description:
                                self.portName = p.name
                                self.ser = serial.Serial(self.portName,9600, timeout=1)
                                self.arduinoState =1
                        else:
                                self.arduinoState = 0
        
        def arduinoQuery(self, query):
                data=query.split("%")
                print(query)
                return data
        
        def isArduinoConnected(self):
                try:
                        self.ser.readline()
                        self.ser.write("Q0".encode())
                        time.sleep(0.1)
                        value = self.ser.readline()
                        qr = self.arduinoQuery(value.decode('utf-8'))
                        if(qr[0]=="Q0" and int(qr[1]) ==1):
                                return True
                        else:
                                return False
                except:
                        return False
                
                
        def isRFIDTagged(self):
                self.ser.write("Q1".encode())
                time.sleep(0.1)
                value = self.ser.readline()
                qr = self.arduinoQuery(value.decode('utf-8'))
                if(qr[0]=="Q1" and int(qr[1]) ==1):
                        return True
                else:
                        return False
        
        def readRFID(self):
                self.ser.write("Q2".encode())
                time.sleep(0.1)
                self.ser.readline()
                time.sleep(0.1)
                value = "NA"
                print("readRFID internal: Reading second line...")
                value = self.ser.readline()
                time.sleep(0.1)
                qr = self.arduinoQuery(value.decode('utf-8'))
                if qr[0]=="Q2":
                        return qr[1].strip()
                        
                return "NA"                

                
                
        def RFID_writeData(self,data):
                self.ser.write(("Q3%"+data).encode())
                time.sleep(0.1)
                value = self.ser.readline()
                qr = self.arduinoQuery(value.decode('utf-8'))
                if(qr[0]=="Q3" and int(qr[1]) ==1):
                        return True
                else:
                        return False
                
        def RFID_readData(self):
                self.ser.write("Q4".encode())
                time.sleep(0.1)
                value = self.ser.readline()
                qr = self.arduinoQuery(value.decode('utf-8'))
                if(qr[0]=="Q4" and int(qr[1]) ==1):
                        return qr[2]
                else:
                        return "ERROR!"
        
        def RFID_connectionStatus(self):
                if self.ser.isOpen():
                        return True
                return False


        
        def arduinoClose(self):
                try:
                        self.ser.close()
                except:
                        pass
        


   

        
                













              











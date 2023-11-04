import time 
import serial
import serial.tools.list_ports


class weighScale():
        

        def __init__(self):
                self.weighState = False
                self.weighList()



        def weighList(self):
                self.ports = list(serial.tools.list_ports.comports())
                self.portList = []
                for p in self.ports:
                        if "Arduino" in p.description:
                                pass
                        else:
                                self.portList.append(p.name)

        
        def weighConnect(self, Iport):
                
                self.scaleSer = serial.Serial(port=Iport,baudrate=2400,timeout=None,parity=serial.PARITY_EVEN,stopbits=serial.STOPBITS_ONE,bytesize=serial.SEVENBITS)
                
                if self.scaleSer.isOpen():
                        self.weighState = True
                        self.scaleSer.readline()
                        print("Weigh Connected")
                        return True
                print("Weigh Not Connected")
                return False
        
        def weighIsConnected(self):
                if self.scaleSer.isOpen():
                        self.weighState = True
                        print("Weigh Connected")
                        return True
                else: 
                        self.weighState = False
                        print("Weigh Not Connected")
                        return False
        
        def weighRead(self):
                self.scaleSer.readline()
                time.sleep(1)
                value = self.ser.readline()
                qr = value.decode('utf-8')
                return qr
        
        def checkConnection(self):
                if not(self.scaleSer.isOpen()):
                        self.weighState = 0
        
        def weighClose(self):
                self.ser.close()
        


   

        
                













              











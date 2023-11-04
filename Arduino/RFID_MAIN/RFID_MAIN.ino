#include "EasyMFRC522.h"

#include <SPI.h>
#include <MFRC522.h>

#define MAX_STRING_SIZE 100  // size of the char array that will be written to the tag
#define BLOCK 1              // initial block, from where the data will be stored in the tag

EasyMFRC522 rfidReader(10, 9);
MFRC522 rfid(10, 9);
MFRC522::MIFARE_Key key; 



byte nuidPICC[4];

String a;

void setup() {
Serial.begin(9600); 

rfidReader.init(); 

//SPI.begin(); // Init SPI bus
//rfid.PCD_Init();


}

void loop() {

while(Serial.available()) {
a= Serial.readString();// read the incoming data as string
String query = getValue(a,0);
resolveQuery(query);
}

}

String getValue(String data, int index)
{
  char separator = '%';
  int found = 0;
  int strIndex[] = {0, -1};
  int maxIndex = data.length()-1;

  for(int i=0; i<=maxIndex && found<=index; i++){
    if(data.charAt(i)==separator || i==maxIndex){
        found++;
        strIndex[0] = strIndex[1]+1;
        strIndex[1] = (i == maxIndex) ? i+1 : i;
    }
  }

  return found>index ? data.substring(strIndex[0], strIndex[1]) : "";
}

void resolveQuery(String query){
  if(query=="Q0"){
  //For Device RFID Sesnor Connectivity Status
  if(Serial){
    Serial.println("Q0%1");
    }
    else{
      Serial.println("Q0%0");
      }
  }
  
else if(query=="Q1"){
    // returns true if a Mifare tag is detected
    if(rfidReader.detectTag()){
        Serial.println("Q1%1");
      }   
      else{
        Serial.println("Q1%0");
        } 
  }

  if(query=="Q2"){
  //Get RFID ID
 rfidReader.init();
 if ( ! rfid.PICC_IsNewCardPresent())
    return;
 
  // Verify if the NUID has been readed
  if ( ! rfid.PICC_ReadCardSerial())
    return;
 
  MFRC522::PICC_Type piccType = rfid.PICC_GetType(rfid.uid.sak);
 
  Serial.print("Q2%");
  printHex(rfid.uid.uidByte, rfid.uid.size);
  Serial.println("");
  }

  else if(query=="Q3"){
    //Write data into tag
  String data = getValue(a,1);
 int result;
 char stringBuffer[MAX_STRING_SIZE];

 strcpy(stringBuffer, data.c_str());  
 int stringSize = strlen(stringBuffer);
    
 result = rfidReader.writeFile(BLOCK, "danaData", (byte*)stringBuffer, stringSize+1);// starting from tag's block #1, writes a data chunk labeled "mylabel", with its content given by stringBuffer, of stringSize+1 bytes (because of the trailing 0 in strings) 

 if (result >= 0) {
     Serial.println("Q3%1");
    } else {
     Serial.println("Q3%0%"+result);
    }
while (Serial.available() > 0) {  // clear "garbage" input from serial
    Serial.read();
  }
    
rfidReader.unselectMifareTag();
delay(1000);
    }

else if(query=="Q4"){//read rfid
    int result;
    char stringBuffer[MAX_STRING_SIZE];
    result = rfidReader.readFile(BLOCK, "danaData", (byte*)stringBuffer, MAX_STRING_SIZE);
    
    stringBuffer[MAX_STRING_SIZE-1] = 0;   // for safety; in case the string was stored without a \0 in the end 
                                           // (would not happen in this example, but it is a good practice when reading strings) 
    if (result >= 0) { // non-negative values indicate success, while negative ones indicate error

      Serial.print("Q4%1%");
      printfSerial("%s\n", stringBuffer);   
    } else { 
      Serial.println("Q4%0%");
      
    }
    }

    else if(query=="Q5"){;}
    else{;}
  }

  void printfSerial(const char *fmt, ...) {
  char buf[128];
  va_list args;
  va_start(args, fmt);
  vsnprintf(buf, sizeof(buf), fmt, args);
  va_end(args);
  Serial.print(buf);
}

void printHex(byte *buffer, byte bufferSize) {
  for (byte i = 0; i < bufferSize; i++) {
    Serial.print(buffer[i] < 0x10 ? "0" :"");
    Serial.print(buffer[i], HEX);
  }
}

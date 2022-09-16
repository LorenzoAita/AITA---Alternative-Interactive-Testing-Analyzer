#include <SoftwareSerial.h>
#include<stdio.h>
#include<stdlib.h>

const int leng = 15; 
char RFin_bytes[leng];
int scheda, cmd, relay, var1, var2, crc;
int pin;
int F;
float ONCycle; //oncycle variable
float OFFCycle; // offcycle variable got microsecond
float T; // tota l time to one cycle ONCycle + OFFcycle
float DutyCycle; // D = (TON/(TON+TOFF))*100 %

void setup() {
  // Open serial communications and wait for port to open:
  Serial.begin(38400);
  while (!Serial) {
    ; // wait for serial port to connect. Needed for native USB port only
  }
  //pinMode(PulseIN, INPUT);
}

void loop() { // run over and over
  while (Serial.available()<leng) {}
  for(int n=0; n<leng; n++)   { 
    RFin_bytes[n] = Serial.read();
    Serial.print(RFin_bytes[n]);
  }
  //Serial.print("Ciao");
  scheda = (int)(RFin_bytes[0])-48;
  cmd = ((int)(RFin_bytes[2])-48)*100 +  ((int)(RFin_bytes[3])-48)*10 + ((int)( RFin_bytes[4])-48);
  relay = (int)(RFin_bytes[6])-48;
  var1 = (int)(RFin_bytes[8])-48;
  var2 = (int)(RFin_bytes[10])-48;
  crc = ((int)(RFin_bytes[12])-48)*100 +  ((int)(RFin_bytes[13])-48)*10 +  ((int)(RFin_bytes[14])-48);
  /*Serial.println(" ");
  Serial.print(scheda);
  Serial.print(" ");
  Serial.print(cmd);
  Serial.print(" ");
  Serial.print(relay);
  Serial.print(" ");
  Serial.print(var1);
  Serial.print(" ");
  Serial.print(var2);
  Serial.print(" ");
  Serial.print(crc);*/

  if (relay == 1)
    pin = 10;
  else if (relay == 2)
    pin = 9;
  else if (relay == 3)
    pin = 12;
  else if (relay == 4)
    pin = 11;
  else if (relay == 5)
    pin = 8;
  else if (relay == 6)
    pin = 7;
  else if (relay == 7)
    pin = 6;
  else if (relay == 8)
    pin = 5;
    
  if (crc == scheda + cmd + relay + var1 + var2) {
    if (var1 == 1){
      Serial.print(" ");
      ONCycle = pulseIn(var2, HIGH);
      OFFCycle = pulseIn(var2, LOW);
      T = ONCycle + OFFCycle;
      DutyCycle = (ONCycle / T) * 100;
      F = 1000000 / T; // è in microsecondi
      Serial.print(DutyCycle);
      Serial.print(" ");
      Serial.print(F);
    }
    
    pinMode(pin, OUTPUT);
    if (cmd == 101)
      digitalWrite(pin, HIGH);
    else if (cmd == 100)
      digitalWrite(pin, LOW);
  }
  else{
    Serial.print(" ");
    Serial.print("CRC errato - Dati corrotti - STRONZO!");
  }
}

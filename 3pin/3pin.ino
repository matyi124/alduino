void setup() {
  Serial.begin(9600);

  // 3 bemenet beállítása
  pinMode(2, INPUT);
  pinMode(3, INPUT);
  pinMode(4, INPUT);
}

void loop() {
  int d2 = digitalRead(2);
  int d3 = digitalRead(3);
  int d4 = digitalRead(4);

  // soros kimenet
  Serial.print("D2:");
  Serial.print(d2);
  Serial.print(" D3:");
  Serial.print(d3);
  Serial.print(" D4:");
  Serial.println(d4);

  delay(200);
}

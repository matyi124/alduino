import serial
import pyautogui
import time

arduino = serial.Serial('COM5', 9600)

print("Program elindult. Most várunk 10 másodpercet, hogy legyen időd megnyitni a PowerPointot...")
time.sleep(10)
print("Most már figyeljük az Arduino jeleit!")

while True:
    if arduino.in_waiting > 0:
        sor = arduino.readline().decode().strip()
        print(f"Üzenet: {sor}")
        if "1" in sor:
            pyautogui.typewrite('6')
            pyautogui.press('enter')
        elif "0" in sor:
            pyautogui.typewrite('1')
            pyautogui.press('enter')
    time.sleep(0.1)

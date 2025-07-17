import serial
import pyautogui
import time
import os
import time
import win32com.client
import win32console
import win32gui

konzol_hwnd = win32console.GetConsoleWindow()
win32gui.ShowWindow(konzol_hwnd, 6)  # vagy 0 ha teljesen elrejtenéd

# aktuális mappa
mappa = os.path.dirname(os.path.abspath(__file__))
pptx_utvonal = os.path.join(mappa, "prezi.pptx")

# PowerPoint elindítása
powerpoint = win32com.client.Dispatch("PowerPoint.Application")
powerpoint.Visible = True

# fájl megnyitása
prezi = powerpoint.Presentations.Open(pptx_utvonal)

# diavetítés indítása
show = prezi.SlideShowSettings
show.Run()

print("PowerPoint bemutató elindult.")

# várhatsz pár másodpercet, vagy itt folytathatod a többi kódod
time.sleep(5)


arduino = serial.Serial('COM5', 9600)

print("Várunk 10 másodpercet, hogy indíthasd a PowerPoint-ot…")
print("Figyelem a jeleket…")

prev_action = None

while True:
    if arduino.in_waiting > 0:
        sor = arduino.readline().decode().strip()
        print(f"Üzenet: {sor}")

        try:
            parts = sor.split()
            d2 = int(parts[0].split(":")[1])
            d3 = int(parts[1].split(":")[1])

            action = None

            if d2 == 1 and d3 == 1:
                action = 1
            elif d2 == 0 and d3 == 1:
                action = 2
            elif d2 == 1 and d3 == 0:
                action = 3
            elif d2 == 0 and d3 == 0:
                action = None  # nem csinál semmit

            if action is not None and action != prev_action:
                print(f"Ugrunk a {action}. diára")
                pyautogui.typewrite(str(action))
                pyautogui.press('enter')
                prev_action = action

        except Exception as e:
            print(f"Hiba: {e}")

    time.sleep(0.1)

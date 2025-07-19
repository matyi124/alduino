import os
import sys
import time
import serial
import pyautogui
import win32com.client
import win32console
import win32gui

# konzol háttérbe
konzol_hwnd = win32console.GetConsoleWindow()
win32gui.ShowWindow(konzol_hwnd, 6)

# megtaláljuk a mappát akkor is, ha .exe-ből fut
if getattr(sys, 'frozen', False):
    # ha PyInstaller .exe
    mappa = sys._MEIPASS
else:
    # ha sima Python
    mappa = os.path.dirname(os.path.abspath(__file__))

pptx_utvonal = os.path.join(mappa, "prezi.pptx")

# PowerPoint indítása
powerpoint = win32com.client.Dispatch("PowerPoint.Application")
powerpoint.Visible = True

# fájl megnyitása
print(f"Próbáljuk megnyitni: {pptx_utvonal}")
prezi = powerpoint.Presentations.Open(pptx_utvonal)

# diavetítés indítása
show = prezi.SlideShowSettings
show.Run()

# előtérbe hozzuk a PowerPoint-ot
shell = win32com.client.Dispatch("WScript.Shell")
shell.AppActivate("PowerPoint")

print("PowerPoint bemutató elindult.")
time.sleep(5)

# Arduino port megnyitása
arduino = serial.Serial('COM5', 9600)

prev_action = None

while True:
    if arduino.in_waiting > 0:
        sor = arduino.readline().decode().strip()
        print(f"Üzenet: {sor}")

        try:
            parts = sor.split()
            d2 = int(parts[0].split(":")[1])
            d3 = int(parts[1].split(":")[1])
            d3 = int(parts[2].split(":")[1])

            action = None

            if d2 == 1 and d3 == 1 and d3 == 1:
                action = 1
            elif d2 == 0 and d3 == 1 and d3 == 1:
                action = 2
            elif d2 == 1 and d3 == 0 and d3 == 1:
                action = 3
            elif d2 == 1 and d3 == 1 and d3 == 0:
                action = 4
            elif d2 == 0 and d3 == 0 and d3 == 0:
                action = None

            if action is not None and action != prev_action:
                pyautogui.typewrite(str(action))
                pyautogui.press('enter')
                prev_action = action

        except Exception as e:
            print(f"Hiba: {e}")

    time.sleep(0.1)

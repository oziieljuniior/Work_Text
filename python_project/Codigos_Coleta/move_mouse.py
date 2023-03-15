import pyautogui 
from datetime import datetime
import time

while True:
    dt = int(datetime.now().strftime("%M"))
    print(dt)
    
    if dt % 5 == 0:
        pyautogui.click((348,292), interval=4)
        time.sleep(30)
        pyautogui.click((477,306), interval=4)
    #(348,292) (477,306)
        while dt % 5 == 0:
            print('Clicou: ', dt)
            dt = int(datetime.now().strftime("%M"))
    
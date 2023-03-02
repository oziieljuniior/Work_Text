import datetime as dt
import pyautogui


while True:
    time = dt.datetime.now().strftime('%M')
    print(time)
    if int(time) % 15 == 0:
        pyautogui.click((97,14), interval = 2)
        pyautogui.click((110,14), interval = 2)

    
import pandas as pd
import pyautogui

#C:/Users/M&A SOLUTIONS BRASIL/Documents/Att/Cias não financeiras - 12-12-2022.xlsx
data = pd.read_excel('C:/Users/M&A SOLUTIONS BRASIL/Documents/Att/Cias não financeiras - 12-12-2022.xlsx')
consulta = data['ticker'].to_list()


pyautogui.screenshot().save('aloha.png')

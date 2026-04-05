import pyautogui
from PIL import ImageGrab
import numpy as np
import time
time.sleep(5)
# Tirar captura de tela
img = ImageGrab.grab()

area_domain = (744, 640,745,650)
img = img.crop(area_domain)
img.save('order.jpeg','jpeg')
convert_im3 = img.convert('RGB').getcolors()
print(convert_im3)                                                    

# Definir cor específica (ex: verde)
cor_especifica = (0, 255, 0)  # RGB

# Encontrar coordenadas da cor específica
#for x in range(img_array.shape[1]):
#    for y in range(img_array.shape[0]):
#        if tuple(img_array[y, x]) == cor_especifica:
#            # Clicar na coordenada encontrada
#            pyautogui.click(x, y)
#            print(f"Clicou em ({x}, {y})")
#            break
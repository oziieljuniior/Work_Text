import pyscreenshot as ImageGrab
from PIL import Image
import pyautogui
import time
i = 0
i2 = 0
while i <= 20:
    im8 = Image.open("im6_referencia.png")
    #(432, 344, 665, 357)
    area3 = (432, 346, 664, 357)
    im9_corte1 = im8.crop(area3)
    im9_corte1.save('foto03.png','png')
    im11 = Image.open('foto03.png')
    
    i += 1
#    time.sleep(5)
    break
    while i2 <= 5:
        guide = pyautogui.locateOnScreen(im11, confidence=0.9)
        print(guide)
        # Verificando se a imagem foi encontrada
        if guide is not None:
            x = guide.left  # Posição X
            y = guide.top   # Posição Y
            click1 = (x,y)
            print(f"Posição X: {x}, Posição Y: {y}")
        else:
            print("Imagem não encontrada na tela.")
        i2 += 1
 
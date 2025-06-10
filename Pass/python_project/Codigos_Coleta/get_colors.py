import pyautogui
from PIL import Image

i = 0
while i <= 100:
    im7 = pyautogui.screenshot().save("im2.jpeg")
    im8 = Image.open("im2.jpeg")
    area3 = (725,310 ,726,311)
    area4 = (175,570,176,571)
    im9_corte1 = im8.crop(area3)
    im10_corte2 = im8.crop(area4)
    im9_corte1.save('foto03.jpeg','jpeg')
    im10_corte2.save('foto04.jpeg','jpeg')
    im11 = Image.open('foto03.jpeg')
    im12 = Image.open('foto04.jpeg')
    convert_im3 = im11.convert('RGB').getcolors()
    convert_im4 = im12.convert('RGB').getcolors()
    print(convert_im3,convert_im4)
    #print(convert_im1)
    i += 1
    #[(1, (255, 255, 255))] [(1, (237, 237, 239))]
    #[(1, (37, 37, 37))] [(1, (39, 39, 39))]
    if convert_im3 == [(1, (37, 37, 37))] and convert_im4 ==  [(1, (39, 39, 39))]:
        print("hello world")
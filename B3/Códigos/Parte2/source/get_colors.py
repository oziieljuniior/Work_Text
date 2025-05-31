import pyscreenshot as ImageGrab
from PIL import Image

i = 0
while i <= 20:
    im7 = ImageGrab.grab().save("im2.jpeg")
    im8 = Image.open("im2.jpeg")
    #[(1, (1, 44, 99))] [(1, (247, 255, 255))]
    area3 = (416, 411, 417, 412)
    area4 = (416, 411, 417, 412)
    #[(1, (1, 44, 99))] [(1, (247, 255, 255))]
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
    #[(1, (293, 227, 253))] [(1, (238, 241, 250))]
    #[(1, (88, 204, 241))] [(1, (255, 255, 255))]
    i += 1
    #[(1, (254, 150, 51))] [(1, (250, 244, 0))]
    if convert_im3 == [(1, (254, 150, 51))] and convert_im4 == [(1, (250, 244, 0))]:
        print("hello world")
 
#See documentation PyPDF2: https://pypdf2.readthedocs.io/en/latest/user/extract-text.html
from PyPDF2 import PdfReader

#Example to the future code
#first, set a path
reader = PdfReader('example.pdf')
print(len(reader.pages))
t = len(reader.pages)
text = open('gettext.txt','x')


for i in range(0,t):
    print(i)
    print(reader.pages[i].extract_text())
    sring = '\n'+reader.pages[i].extract_text()
    text.write(sring)

text.close()

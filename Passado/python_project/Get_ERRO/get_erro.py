#example to get text and view grammal error
#see doc extract_text.py
#Documentation language-tool: https://gihub.com/jxmorrir12/language_tool_python
from PyPDF2 import PdfReader
import language_tool_python as lt

reader = PdfReader('example.pdf')
print(reader.pages[1].extract_text())
tool = lt.LanguageTool('pt-BR', 'en')
text = reader.pages[1].extract_text()
erro = ['DFP - Demonstrações Financeiras Padronizadas', '/']
for s in erro:
    text.replace(s,"")
print(text)
matches = tool.check(text)
print(matches)
print(len(matches))

#option1


#tool.close()

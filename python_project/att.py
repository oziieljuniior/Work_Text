import language_tool_python as lt

tool = lt.LanguageTool("pt-BR", config = {'requestLimit' : 10})
print(tool)
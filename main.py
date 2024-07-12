import win32com.client as win32
import os, log, json, sys
log = log.Logger()

class AutoUpdate:
    __jsonData = None
    __excel = None

    def __init__(self) -> None:
        try:
            log.w("Inicializando objeto.")
            self.setJsonData(self.consumeJson())
            self.__excel = win32.gencache.EnsureDispatch('Excel.Application') 
        except Exception as e:
            log.w("ERROR", str(e))

    def setJsonData(self, content) -> None:
        try:
            log.w("Conteúdo do Json sendo alocado na memória.")
            self.__jsonData = content
        except Exception as e:
            log.w("ERROR", str(e))

    def getJsonData(self):
        try:
            log.w("Consumindo conteúdo do Json alocado em memória.")
            return self.__jsonData
        except Exception as e:
            log.w("ERROR", str(e))

    def consumeJson(self):
        try:
            log.w("Iniciando leitura do Json.")
            with open('spreadsheets.json', 'r', encoding='utf-8') as jsonFile:
                return json.load(jsonFile)
        except Exception as e:
            log.w("ERROR", str(e))

    def main(self) -> None:
        try:
            for dirs, info in self.getJsonData()['dirs'].items():
                log.w(f"Iterando sobre o diretório: {info['spreadsheet_route']}.")
                for file in info["spreadsheets_files"]:
                    self.update(os.path.join(info['spreadsheet_route'], file))
            else:
                log.w('        Fechando Excel.')
                self.__excel.Quit() # Fechar o Excel
                log.w("Fim da aplicação")
                sys.exit(0)

        except Exception as e:
            log.w("ERROR", f"{e}")

    def update(self, file) -> None:
        log.w(f"    Atuando sobre o arquivo {file}")
        try:

            log.w('        Executando Excel em background.')
            self.__excel.Visible = True # Executar em background
            #excel.Visible = False # Executar em background

            log.w('        Abrindo arquivo em workbook.')
            #workbook = excel.Workbooks.Open(file) # Abrir o arquivo
            workbook = self.__excel.Workbooks.Open(os.path.join(os.getcwd(), file)) # Abrir o arquivo

            log.w('        Atualizando conexões.')
            self.__excel.ActiveWorkbook.RefreshAll() # Atualizar todas as conexões

            log.w('        Salvando workbook.')
            workbook.Save() # Salvar

            log.w('        Fechando workbook')
            workbook.Close() # Fechar

        except Exception as e:
            log.w("ERROR", f"        {e}")

if __name__ == "__main__":
    try:
        log.w("Início da aplicação.")
        app = AutoUpdate()
        app.main()

    except Exception as e:
        log.w("ERROR", str(e))
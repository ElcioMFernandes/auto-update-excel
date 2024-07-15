import win32com.client as win32
import os, json, logging, sys, datetime

logging.basicConfig(level=logging.DEBUG,
                    format='%(asctime)s - %(levelname)s - %(message)s',
                    filename=f"{datetime.datetime.today().strftime('%d%m%Y')}.log",
                    filemode='a')
class AutoUpdate:
    __excelApp = None
    __JsonFile = None
    __jsonData = None
    
    def setJsonFile(self, jsonFile) -> None:
        self.__JsonFile = jsonFile

    def getJsonFile(self) -> str:
        return self.__JsonFile
    
    def setJsonData(self, jsonData) -> None:
        self.__jsonData = jsonData

    def getJsonData(self) -> dict:
        return self.__jsonData

    def setExcelApp(self) -> None:
        self.__excelApp = win32.gencache.EnsureDispatch('Excel.Application')
        self.__excelApp.Visible = True

    def getExcelApp(self) -> object:
        return self.__excelApp

    def __init__(self) -> None:
        logging.info("Construtor inicializado.")
        
        try:
            self.setExcelApp()
            logging.info("ExcelApp inicializado.")
        
            if os.path.isfile(os.path.join(os.getcwd(),'spreadsheets.json')):
        
                try:
                    self.setJsonFile('spreadsheets.json')
                    logging.info("Arquivo Json setado.")
        
                    try:
        
                        with open(self.getJsonFile(), 'r', encoding='utf-8') as jsonFileRead:
                            self.setJsonData(json.load(jsonFileRead))                        
                            logging.info("Arquivo Json lido.")
        
                    except Exception as e:
                        logging.error(f"Falha ao ler arquivo de configuração: {e}")
                        self.getExcelApp().Quit()
                        sys.exit(1)

                except Exception as e:
                    logging.error(f"Falha ao setar o arquivo de configuração: {e}")
                    self.getExcelApp().Quit()
                    sys.exit(1)

        except Exception as e:
            logging.error(f"Falha ao iniciar aplicativo: {e}")
            sys.exit(1)

    def main(self) -> None:
        for directory, information in self.getJsonData().items():
            logging.info(f"Chave: {directory}.")

            if os.path.isdir(information['route']):
                logging.info(f"Rota de {directory} existente.")

                for file in information['files']:
                    logging.info(f"Arquivo: {file}")

                    if os.path.isfile(os.path.join(information['route'], f"{file}.{information['type']}")):
                        logging.info(f"Arquivo: {file} existente.")
                        self.update(os.path.join(information['route'], f"{file}.{information['type']}"))
                    
                    else:
                        logging.error(f"O arquivo {file}, indicado em {directory} não existe")
                        pass

            else:
                logging.error(f"A rota indicada em {directory} não existe")
                pass

        else:
            logging.info("Fechando ExcelApp e encerrando aplicação.")
            self.getExcelApp().Quit()
            sys.exit(0)

    def update(self, fileToUpdate) -> None:
        logging.info(f"Iniciando atualização do arquivo {fileToUpdate}")
        
        try:
            workbook = self.getExcelApp().Workbooks.Open(fileToUpdate)
            logging.info("Workbook iniciado.")
        
            try:
                self.getExcelApp().ActiveWorkbook.RefreshAll()
                logging.info("Arquivo atualizado.")
        
                try:
                    workbook.Save()
                    logging.info("Arquivo salvo.")
                    workbook.Close()
                    logging.info("Workbook fechado.")
        
                except Exception as e:
                    logging.error(f"Erro ao salvar ou fechar: {e}")
        
            except Exception as e:
                logging.error(f"Erro ao atualizar tudo: {e}")
        
        except Exception as e:
            logging.error(f"Erro ao abrir o arquivo: {e}")

if __name__ == '__main__':
    app = AutoUpdate()
    app.main()
    
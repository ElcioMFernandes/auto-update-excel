import win32com.client as win32
import os, json, logging, sys, datetime
from tkinter import messagebox
#import schedule

class AutoUpdate:
    __excelApp = None
    __jsonFile = None
    __jsonData = None
    
    def setJsonFile(self, jsonFile:str) -> None:
        """
        Set a value to the private jsonFile attribute.
        """
        self.__jsonFile = jsonFile

    def getJsonFile(self) -> str:
        """
        Get on jsonFile private attribute value.
        """
        return self.__jsonFile
    
    def setJsonData(self, jsonData:dict) -> None:
        """
        Set a value to the private jsonData attribute.
        """
        self.__jsonData = jsonData

    def getJsonData(self) -> dict:
        """
        Get on jsonData private attribute value.
        """
        return self.__jsonData

    def setExcelApp(self) -> None:
        """
        Set in the excelApp private attribute that stores the win32com library object.
        """
        self.__excelApp = win32.gencache.EnsureDispatch('Excel.Application')
        self.__excelApp.Visible = True

    def getExcelApp(self) -> object:
        """
        Returns the ExcelApp object from the win32com library.
        """
        return self.__excelApp

    def __init__(self) -> None:
        """
        Class constructor.
        """
        logging.info("Iniciando construtor")
        try:
            self.setExcelApp()
            if os.path.isfile(os.path.join(os.getcwd(),'spreadsheets.json')):
                try:
                    self.setJsonFile('spreadsheets.json')
                    try:        
                        with open(self.getJsonFile(), 'r', encoding='utf-8') as jsonFileRead:
                            self.setJsonData(json.load(jsonFileRead))
                            logging.info("Construtor finalizado.")                      
                    except Exception as e:
                        logging.error(f"Falha ao ler arquivo de configuração: {e}")
                        self.getExcelApp().Quit()
                        sys.exit(1)

                except Exception as e:
                    logging.error(f"Falha ao setar o arquivo de configuração: {e}")
                    self.getExcelApp().Quit()
                    sys.exit(1)
            else:
                self.getExcelApp().Quit()
                self.createJson()
                messagebox.showerror(f"Erro no init","Arquivo Json vazio ou inexistente.")
                logging.error("Arquivo Json vazio ou inexistente.")
                sys.exit(1)

        except Exception as e:
            logging.error(f"Falha ao iniciar aplicativo: {e}")
            sys.exit(1)

    def main(self) -> None:
        """
        Main application method.
        """
        logging.info("Iniciando bloco principal.\n")
        try: 
            for directory, information in self.getJsonData().items():
                if os.path.isdir(information['route']):
                    logging.info(f"+{directory.replace('_', ' ').capitalize()}")
                    logging.info(f"|__{information['route']}")                    
                
                    for i, file in enumerate(information['files']):
                        is_last_file = i == len(information['files']) - 1
                        logging.info(f"   |__Arquivo: {file}")
                
                        if os.path.isfile(os.path.join(information['route'], f"{file}.{information['type']}")):
                            self.update(os.path.join(information['route'], f"{file}.{information['type']}"), is_last_file)
                
                        else:
                            logging.error(f"O arquivo {file}, indicado em {directory} não existe")
                            pass

                else:
                    logging.error(f"A rota indicada em {directory} não existe")
                    pass

            else:
                self.getExcelApp().Quit()
                logging.info("Finalizado todos os arquivos.")
                sys.exit(0)

        except AttributeError:
            self.getExcelApp().Quit()
            self.createJson()
            messagebox.showerror(f"Erro na main","Arquivo Json vazio ou inexistente.")
            logging.error("Arquivo Json vazio ou inexistente.")
            sys.exit(1)

        except Exception as e:
            self.getExcelApp().Quit()
            logging.error(e)
            sys.exit(1)

    def update(self, fileToUpdate:str, last_file) -> None:
        """
        Method that updates data in the spreadsheet.
        """
        try:
            workbook = self.getExcelApp().Workbooks.Open(fileToUpdate)
        
            try:
                self.getExcelApp().ActiveWorkbook.RefreshAll()
                try:
                    workbook.Save()
                    workbook.Close()
                    if last_file:
                        logging.info("      |__Arquivo atualizado.\n")
                    else:
                        logging.info("   |  |__Arquivo atualizado.")
        
                except Exception as e:
                    logging.error(f"Erro ao salvar ou fechar arquivo: {e}")
        
            except Exception as e:
                logging.error(f"Erro ao tentar atualizar: {e}")
        
        except Exception as e:
            logging.error(f"Erro ao abrir o arquivo: {e}")

    def createJson(self):
        """
        Method for creating a json template if it does not exist.
        """
        data = {
            "nome_identificacao_chave": {
                "owner": "proprietario_pasta",
                "route": "caminho_pasta",
                "type": "xlsx",
                "files": [
                    "nome_arquivo_sem_extensao1",
                    "nome_arquivo_sem_extensao2"
                ]
            },
            "nome_de_identificacao_chave": {
                "owner": "proprietario_pasta",
                "route": "caminho_pasta",
                "type": "xlsx",
                "files": [
                    "nome_arquivo_sem_extensao3",
                    "nome_arquivo_sem_extensao4"
                ]
            }
        }

        with open('spreadsheets.json', 'w') as json_file:
            json.dump(data, json_file, indent=4)

if __name__ == '__main__':
    if os.path.isdir(os.path.join(os.getcwd(),'log')) == False:
        os.makedirs('log')

    logging.basicConfig(level=logging.DEBUG,
                    format='%(asctime)s - %(levelname)s - %(message)s',
                    filename=os.path.join(os.getcwd(),'log',f"{datetime.datetime.now().strftime('%d-%m-%Y-%H-%M-%S')}.log"),
                    filemode='a',
                    encoding='utf-8')

    try:
        app = AutoUpdate()
        logging.info("Instância do objeto AutoUpdate inicializada.")
        app.main()
    except Exception as e:
        logging.error(e)
    
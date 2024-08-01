import os, sys, json, time, logging, datetime, tkinter.messagebox

try:
    import win32com.client as win32
except Exception as e:
    tkinter.messagebox.showerror('Erro','Erro ao importar a biblioteca win32com.client')
    sys.exit(1)

#import schedule

class AutoUpdate:
    __excelApp = None
    __jsonFile = None
    __jsonData = None
    __interval = None
    __visible  = None
    
    def __setVisible(self, visibility: bool) -> None:
        """ Método que atribui um valor passado por parâmetro para o atributo privado __visible. """
        self.__visible = visibility

    def __getVisible(self) -> bool:
        """ Método para pegar o valor do atributo privado __visible. """
        return self.__visible

    def __setInterval(self, time: int) -> None:
        """ Método para setar o tempo de intervalo entre a execução das atualizações de planilha. """
        self.__interval = time

    def __getInterval(self) -> int:
        """ Método para pegar o valor do atributo privado __interval, responsável pelo tempo de intervalo entre a execução das atualizações de planilha. """
        return self.__interval

    def __setJsonFile(self, jsonFile:str) -> None:
        """ Set a value to the private jsonFile attribute. """
        self.__jsonFile = jsonFile

    def __getJsonFile(self) -> str:
        """ Get on jsonFile private attribute value. """
        return self.__jsonFile
    
    def __setJsonData(self, jsonData:dict) -> None:
        """ Set a value to the private jsonData attribute. """
        self.__jsonData = jsonData

    def __getJsonData(self) -> dict:
        """ Get on jsonData private attribute value. """
        return self.__jsonData

    def __setExcelApp(self) -> None:
        """ Set in the excelApp private attribute that stores the win32com library object. """
        self.__excelApp = win32.gencache.EnsureDispatch('Excel.Application')
        self.__excelApp.Visible = self.__getVisible()

    def __getExcelApp(self) -> object:
        """ Returns the ExcelApp object from the win32com library. """
        return self.__excelApp

    def __init__(self) -> None:
        """ Classe construtora. """
        if os.path.isfile(os.path.join('C:\\Users', os.getlogin(), 'auto update','spreadsheets.json')): # Se o arquivo existir

            self.__setJsonFile(os.path.join('C:\\Users', os.getlogin(), 'auto update','spreadsheets.json')) # Atribui o __jsonFile como 'spreadsheets.json'
            logging.info('Arquivo Json setado.')

            try:        
                with open(self.__getJsonFile(), 'r', encoding='utf-8') as jsonFileRead:
                        try:
                            self.__setJsonData(json.load(jsonFileRead)) # Atribiu o conteúdo do Json ao __jsonData
                            logging.info('Conteúdo do arquivo json carregado em variável.')

                        except Exception as e:
                            logging.error('Conteúdo do json corrompido ou vazio.')
                        
                        try:
                            self.__setVisible(self.__getJsonData()['settings']['visible']) # Atribui o valor do dicionário settings, chave visible, a variável __visible
                            logging.info(f'Visibilidade da execução definida como {self.__getVisible()}.')

                        except Exception as e:
                            logging.error(logging.error('Conteúdo do json corrompido ou vazio no trecho settings, visible.'))

                        try:
                            self.__setInterval(self.__getJsonData()['settings']['interval']) # Atribui o valor do dicionário settings, chave interval, a variável __interval
                            logging.info(f'Intevalo entre a atualização definido para {self.__getInterval()} segundos.')
                        except Exception as e:
                            logging.error(logging.error('Conteúdo do json corrompido ou vazio no trecho settings, interval.'))

                        try:
                            self.__setExcelApp() # Chama o método ExcelApp
                            logging.info('Aplicativo Excel inicializado.')
                            
                        except Exception as e:
                            logging.critical(f"Falha no método __init__: {e}")
                            self.__getExcelApp().Quit()
                            sys.exit(1)

            except Exception as e:
                logging.critical(f"Falha no método __init__: {e}")
                sys.exit(1)

        else:
            self.__createJson()
            logging.error("Arquivo Json vazio ou inexistente.")
            sys.exit(1)

    def main(self) -> None:
        """ Main application method. """
        logging.info("Iniciando método main.\n")

        try: 
            for directory, information in self.__getJsonData().items():
                if directory != 'settings' and os.path.isdir(information['route']):
                    logging.info(f"+{directory.replace('_', ' ').capitalize()}")
                    logging.info(f"|__{information['route']}")                    
                
                    for i, file in enumerate(information['files']):
                        is_last_file = i == len(information['files']) - 1
                        logging.info(f"   |__Arquivo: {file}")
                
                        if os.path.isfile(os.path.join(information['route'], f"{file}.{information['type']}")):
                            self.__update(os.path.join(information['route'], f"{file}.{information['type']}"), is_last_file)
                            
                            time.sleep(self.__getInterval())
                
                        else:
                            logging.error(f"O arquivo {file}, indicado em {directory} não existe")
                            pass

                else:
                    if directory != 'settings':
                        logging.error(f"A rota indicada em {directory} não existe")
                    pass

            else:
                self.__getExcelApp().Quit()
                logging.info("Finalizado todos os arquivos.")
                sys.exit(0)

        except AttributeError:
            self.__getExcelApp().Quit()
            self.__createJson()
            logging.error("Arquivo Json vazio ou inexistente.")
            sys.exit(1)

        except Exception as e:
            self.__getExcelApp().Quit()
            logging.error(e)
            sys.exit(1)

    def __update(self, fileToUpdate:str, last_file) -> None:
        """ Method that updates data in the spreadsheet. """
        try:
            workbook = self.__getExcelApp().Workbooks.Open(fileToUpdate)
        
            try:
                self.__getExcelApp().ActiveWorkbook.RefreshAll()
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

    def __createJson(self):
        """ Method for creating a json template if it does not exist. """
        data = {
            "settings":{
                "visible": True,
                "interval": 10
            },
            "nome_de_identificacao_chave1": {
                "owner": "proprietario_pasta1",
                "route": "caminho_pasta1",
                "type": "xlsx",
                "files": [
                    "nome_arquivo_sem_extensao1",
                    "nome_arquivo_sem_extensao2"
                ]
            },
            "nome_de_identificacao_chave2": {
                "owner": "proprietario_pasta2",
                "route": "caminho_pasta2",
                "type": "xlsx",
                "files": [
                    "nome_arquivo_sem_extensao3",
                    "nome_arquivo_sem_extensao4"
                ]
            }
        }

        with open(os.path.join('C:\\Users', os.getlogin(), 'auto update','spreadsheets.json'), 'w') as json_file:
            json.dump(data, json_file, indent=4)

if __name__ == '__main__':
    if os.path.isdir(os.path.join('C:\\Users', os.getlogin(), 'auto update','log')) == False:
        os.makedirs('log')

    logging.basicConfig(level=logging.DEBUG,
                    format='%(asctime)s - %(levelname)s - %(message)s',
                    filename=os.path.join('C:\\Users', os.getlogin(), 'auto update','log',f"{datetime.datetime.now().strftime('%d-%m-%Y-%H-%M-%S')}.log"),
                    filemode='a',
                    encoding='utf-8')

    try:
        app = AutoUpdate()
        app.main()

    except Exception as e:
        logging.error(e)

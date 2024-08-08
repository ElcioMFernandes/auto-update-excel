# Importação de bibliotecas nativas do Python

import os, sys, json, time, logging, datetime, tkinter.messagebox

# Importação de bibliotecas externas

try: # Bloco try para captura de erros.
    import win32com.client as win32

except Exception as e:# Em caso de erro na importação.
    tkinter.messagebox.showerror('Erro','Erro ao importar a biblioteca win32com.client') # Apresenta uma mensagem de erro.
    sys.exit(1) # Encerra o programa.

# Classe principal
class AutoUpdate:
    __excelApp = None # Armazena a instância do objeto excel.
    __jsonFile = None # Armazena uma str com o caminho e nome do arquivo json.
    __jsonData = None # Armazena uma str com o conteúdo do arquivo json.
    __interval = None # Armazena um int, referente ao intervalo entre a atualização de uma planilha e outra em segundos.
    __visible  = None # Aemazena um valor bool, que defini se será em segundo plano ou não.
    
    def __setVisible(self, visibility: bool) -> None:
        """ Atribui um valor booleano para __visible. """
        self.__visible = visibility

    def __getVisible(self) -> bool:
        """ Retorna o valor de __visible. """
        return self.__visible

    def __setInterval(self, time: int) -> None:
        """ Atribui um valor int para __interval. """
        self.__interval = time

    def __getInterval(self) -> int:
        """ Retorna o valor de __interval. """
        return self.__interval

    def __setJsonFile(self, jsonFile:str) -> None:
        """ Atribui um valor str para __jsonFile. """
        self.__jsonFile = jsonFile

    def __getJsonFile(self) -> str:
        """ Retorna o valor de __jsonFile. """
        return self.__jsonFile
    
    def __setJsonData(self, jsonData:dict) -> None:
        """ Atribui um valor dict para __jsonData. """
        self.__jsonData = jsonData

    def __getJsonData(self) -> dict:
        """ Retorna o valor de __jsonData. """
        return self.__jsonData

    def __setExcelApp(self) -> None:
        """ Atribui a a __excelApp a instância de Excel.Application. """
        self.__excelApp = win32.gencache.EnsureDispatch('Excel.Application')
        self.__excelApp.Visible = self.__getVisible()
        self.__excelApp.DisplayAlerts = False
        self.__excelApp.AskToUpdateLinks = False

    def __getExcelApp(self) -> object:
        """ Retorna o objeto de __excelApp. """
        return self.__excelApp

    def __init__(self) -> None:
        """ Classe construtora. """
        # Se o arquivo spreadsheets.json existir no caminho C:\users\usuário atual\auto update\.
        if os.path.isfile(os.path.join('C:\\Users', os.getlogin(), 'auto update','spreadsheets.json')):
            # Atribui a variável jsonFile o caminho até o arquivo + spreadsheets.json.
            self.__setJsonFile(os.path.join('C:\\Users', os.getlogin(), 'auto update','spreadsheets.json'))
            # Informações de log.
            logging.info('Arquivo Json setado.')
            # Bloco try para captura de erros.
            try:
                # Abre o arquivo indicado em __jsonFile
                with open(self.__getJsonFile(), 'r', encoding='utf-8') as jsonFileRead:
                        # Bloco try para captura de erros.
                        try:
                            # Atribui em __jsonData o conteúdo do __jsonFile
                            self.__setJsonData(json.load(jsonFileRead))
                            # Informações de log.
                            logging.info('Conteúdo do arquivo json carregado em variável.')
                        # Em caso de erro.
                        except Exception as e:
                            # Informações de log.
                            logging.error('Conteúdo do json corrompido ou vazio.')
                        # Bloco try para captura de erros.
                        try:
                            # Atribui em __visible o conteúdo do dicionário settings, chave visible.
                            self.__setVisible(self.__getJsonData()['settings']['visible'])
                            # Informações de log.
                            logging.info(f'Visibilidade da execução definida como {self.__getVisible()}.')
                        # Em caso de erro.
                        except Exception as e:
                            # Informações de log.
                            logging.error(logging.error('Conteúdo do json corrompido ou vazio no trecho settings, visible.'))
                        # Bloco try para captura de erros.
                        try:
                            # Atribui em __interval o conteúdo do dicionário settings, chave interval.
                            self.__setInterval(self.__getJsonData()['settings']['interval'])
                            # Informações de log.
                            logging.info(f'Intevalo entre a atualização definido para {self.__getInterval()} segundos.')
                        # Em caso de erro.
                        except Exception as e:
                            # Informações de log.
                            logging.error(logging.error('Conteúdo do json corrompido ou vazio no trecho settings, interval.'))
                        # Bloco try para captura de erros.
                        try:
                            # Atribui a __excelApp o aplicativo Excel.
                            self.__setExcelApp()
                            # Informações de log.
                            logging.info('Aplicativo Excel inicializado.')
                        # Em caso de erro.
                        except Exception as e:
                            # Informações de log.
                            logging.critical(f"Falha no método __init__: {e}")
                            # Fecha o excel.
                            self.__getExcelApp().Quit()
                            # Encerra o programa.
                            sys.exit(1)

            # Em caso de erro.
            except Exception as e:
                # Informações de log.
                logging.critical(f"Falha no método __init__: {e}")
                # Encerra o programa.
                sys.exit(1)
        # Se o arquivo spreadsheets.json não existir, cria um exemplo e encerra o programa.

        else:
            # Chama a função __createJson para criar um template do json.
            self.__createJson()
            # Informações de log.
            logging.error("Arquivo Json vazio ou inexistente.")
            # Encerra o programa.
            sys.exit(1)

    def main(self) -> None:
        """ Bloco principal """
        # Informações de log.
        logging.info("Iniciando método main.\n")
        # Bloco try para captura de erros.
        try:
            json_data = self.__getJsonData()
            directories = [d for d in json_data.items() if d[0] != 'settings' and os.path.isdir(d[1]['route'])]
            total_directories = len(directories)
            processed_directories = 0

            # Para cada dicionário, chave no dict __jsonData.
            for directory, information in directories:
                processed_directories += 1
                is_last_directory = processed_directories == total_directories
                # Informações de log.
                logging.info(f"+{directory.replace('_', ' ').capitalize()}")
                # Informações de log.
                logging.info(f"|__{information['route']}")
                # Para cada arquivo do array da chave files.
                for i, file in enumerate(information['files']):
                    # Verifica se é o último arquivo.
                    is_last_file = i == len(information['files']) - 1
                    # Informações de log
                    logging.info(f"   |__Arquivo: {file}")
                    # Se o arquivo existir.
                    if os.path.isfile(os.path.join(information['route'], f"{file}.{information['type']}")):
                        # Chama o método __update para o arquivo.
                        self.__update(os.path.join(information['route'], f"{file}.{information['type']}"), is_last_file)
                        # Coloca o programa para dormir pelo período indicado em __interval (em segundos).
                        if not (is_last_file and is_last_directory):
                            time.sleep(self.__getInterval())
                    # Se o arquivo não existir
                    else:
                        # Informações de log.
                        logging.error(f"O arquivo {file}, indicado em {directory} não existe")
                        # Passa para o próximo
                        pass

            # Bloco else para quando o loop terminar.
            else:
                # Fecha o excel.
                self.__getExcelApp().Quit()
                # Informações de log.
                logging.info("Finalizado todos os arquivos.")
                # Encerra o programa.
                sys.exit(0)

        # Em caso de erro.
        except AttributeError:
            # Fecha o excel.
            self.__getExcelApp().Quit()
            # Cria um template do json.
            self.__createJson()
            # Informações de log.
            logging.error("Arquivo Json vazio ou inexistente.")
            # Encerra o programa.
            sys.exit(1)

        # Em caso de erro.
        except Exception as e:
            # Fecha o excel.
            self.__getExcelApp().Quit()
            # Informações de log.
            logging.error(e)
            # Fecha o programa.
            sys.exit(1)

    def __update(self, fileToUpdate: str, lastFile: bool) -> None:
        """ Método utilizado para atualizar planilhas.
        fileToUpdate: String com o caminho até a planilha.
        lastFile: Bool indicando se é o último arquivo.
        """
        connection_types = {
            1: "OLEDBConnection",
            2: "ODBCConnection",
            3: "WebConnection",
            4: "TextConnection",
            5: "WorksheetConnection",
            6: "ModelConnection",
        }

        try:
            # Cria uma instância de um workbook com o arquivo passado por parâmetro.
            workbook = self.__getExcelApp().Workbooks.Open(fileToUpdate, UpdateLinks=False)
            try:
                # Atualiza cada conexão individualmente.
                for i, conn in enumerate(workbook.Connections):
                    is_last_conn = i == len(workbook.Connections) - 1
                    logging.info(f"   |  |__Conexão {conn.Name}.")
                    connType = connection_types.get(conn.Type)
                    if connType:
                        try:
                            getattr(conn, connType).Refresh()
                            if is_last_conn:
                                logging.info(f"   |__|  |__Conexão atualizada.")
                            else:
                                logging.info(f"   |  |  |__Conexão atualizada.")
                        except Exception as e:
                            logging.error(f"   |  |  |__Erro ao atualizar conexão: {e}")
                    else:
                        pass
                else:
                    try:
                        # Fecha o workbook.
                        workbook.Close(SaveChanges=True)
                        if lastFile:
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
        # Template do json.
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
        # Cria um arquivo json.
        with open(os.path.join('C:\\Users', os.getlogin(), 'auto update','spreadsheets.json'), 'w') as json_file:
            # Escre o template no arquivo json.
            json.dump(data, json_file, indent=4)

# Se o nome do arquivo for main.
if __name__ == '__main__':
    # Verifica se existe uma pasta de log.
    if os.path.isdir(os.path.join('C:\\Users', os.getlogin(), 'auto update','log')) == False:
        # Cria uma pasta de log.
        os.makedirs('log')
    # Configuração do logger.
    logging.basicConfig(level=logging.DEBUG,
                    format='%(asctime)s - %(levelname)s - %(message)s',
                    filename=os.path.join('C:\\Users', os.getlogin(), 'auto update','log',f"{datetime.datetime.now().strftime('%d-%m-%Y-%H-%M-%S')}.log"),
                    filemode='a',
                    encoding='utf-8')
    # Bloco try para captura de erro.
    try:
        # Instância do objeto.
        app = AutoUpdate()
        # Executa método principal.
        app.main()
    # Em caso de erro.
    except Exception as e:
        # Informações de log.
        logging.error(e)
        # Encerra o programa.
        sys.exit(1)
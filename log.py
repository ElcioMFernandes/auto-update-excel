import os, datetime

class Logger:
    __logMessage = None
    __logType = None
    __logFile = None
    __logDir = None
    
    def setLogMessage(self, message):
        try:
            self.__logMessage = message
            return
        except Exception as e:
            return e

    def getLogMessage(self):
        return self.__logMessage

    def setLogType(self, type):
        choices = ['DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL']
        try:
            if type in choices:
                self.__logType = type
                return
            else:
                self.__logType = 'INFO'
                return
            
        except Exception as e:
            return e
        
    def getLogType(self):
        return self.__logType

    def setLogDir(self, dir='log'):
        try:
            self.__logDir = dir
            return
        except Exception as e:
            return e
        
    def getLogDir(self):
        return self.__logDir

    def setLogFile(self, file):
        try:
            self.__logFile = file
        except Exception as e:
            return e
        
    def getLogFile(self):
        return self.__logFile

    def __init__(self) -> None:
        self.setLogDir()
        if self.getLogDir() not in os.listdir(os.getcwd()):
            os.mkdir(self.getLogDir())
            
        self.setLogFile(f"{datetime.datetime.now().strftime('%d%m%Y-%H%M%S')}.log")

        if self.getLogFile() not in os.listdir(self.getLogDir()):
            with open(os.path.join(self.getLogDir(), self.getLogFile()), 'a', encoding='utf-8') as created_file:
                pass

        self.__ROOT = os.path.join(self.getLogDir(), self.getLogFile())

    def w(self, message, type='INFO'):
        self.setLogType(type)
        if message != None:
            self.setLogMessage(message)
            with open(self.__ROOT, 'a', encoding='utf-8') as lf:
                lf.write(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] - {self.getLogType()}: {self.getLogMessage()}\n")

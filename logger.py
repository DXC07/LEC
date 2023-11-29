import logging
from datetime import datetime
import os

loggers = {}

class Logging:

    
    def get_logger(self, name, log_path):

        global loggers

        if loggers.get(name):
            return loggers.get(name)
        else:
            today = datetime.now().strftime('%Y%m%d')
            user = '[' + os.getlogin() + ']'

            # Configurar nombre y nivel del logger
            logger = logging.getLogger(name)
            logger.setLevel(logging.DEBUG)

            # Dar formato al logger
            formatter = logging.Formatter(user + ' [%(name)s] [%(asctime)s] [%(levelname)s] %(message)s', datefmt = '%Y-%m-%d %H:%M:%S')

            # Definir ruta y formato del log
            path = os.path.abspath(log_path)

            if not os.path.exists(path):
                os.makedirs(path)

            # Crear file handler y definir el nivel del log
            file_handler = logging.FileHandler(os.path.join(path, 'log_' + today + '.log'), encoding = 'utf-8')
            file_handler.setLevel(logging.DEBUG)
            file_handler.setFormatter(formatter)

            # Crear file handler de consola y definir el nivel del log
            console_handler = logging.StreamHandler()
            console_handler.setLevel(logging.INFO)
            console_handler.setFormatter(formatter)

            # Adicionar handlers al logger
            logger.addHandler(file_handler)
            logger.addHandler(console_handler)
            loggers[name] = logger

            return logger
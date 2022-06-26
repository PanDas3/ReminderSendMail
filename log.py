from logging import getLogger
from logging import basicConfig
from logging import error
from logging import info
from logging import ERROR
from logging import INFO

class Log():

    def __init__(self):
        pass

    def log_error(self, msg):
        logger = getLogger("Logger")
        logger.setLevel(ERROR)
        basicConfig(filename="WAP_SendMail.log", level=ERROR, format="%(asctime)s \t %(levelname)s: \t %(message)s")
        error(msg)

    def log_info(self, msg):
        logger = getLogger("Logger")
        logger.setLevel(INFO)
        basicConfig(filename="WAP_SendMail.log", level=INFO, format="%(asctime)s \t %(levelname)s: \t\t %(message)s")
        info(msg)

    def __del__(self):
        pass
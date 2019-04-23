import logging
from functools import reduce


class SimpleLog(object):
    
    def __init__(self, file_name, console=False):
        self.logger = logging.getLogger(file_name)
        self.logger.setLevel(20)
        fmt = '[%(asctime)s] - %(levelname)s : %(message)s'
        formatter = logging.Formatter(fmt)
        if console:
            stream_handler = logging.StreamHandler()
            stream_handler.setFormatter(formatter)
            self.__msg = ''
            self.logger.addHandler(stream_handler)
        log_file = './' + file_name + '.log'
        file_handler = logging.FileHandler(log_file, encoding='utf8')
        file_handler.setFormatter(formatter)
        self.logger.addHandler(file_handler)

    def debug(self, msg):
        self.logger.debug(msg)

    def info(self, *msg):
        message = reduce((lambda a, b: a + b), [str(i) for i in msg])
        self.logger.info(message)

    def warning(self, *msg):
        message = reduce((lambda a, b: a + b), [str(i) for i in msg])
        self.logger.warning(message)
        
    def error(self, *msg):
        message = reduce((lambda a, b: a + b), [str(i) for i in msg])
        self.__msg = message
        self.logger.error(message)
        
        
    def critical(self, msg):
        self.logger.critical(msg)

    def log(self, level, msg):
        self.logger.log(level, msg)

    def set_level(self, level):
        self.logger.setLevel(level)
    
    @property
    def msg(self):
        return self.__msg
    
    @staticmethod
    def set_msg(*args):
        SimpleLog.msg_l.extend(list(args))

    @staticmethod
    def disable():
        logging.disable(50)


log = SimpleLog('info')
err_log = SimpleLog('warning', console=True)
import logging
import os
from logging.handlers import TimedRotatingFileHandler

class LogManager:
    def __init__(self, logfile='modbus.log'):
        os.makedirs(os.path.dirname(logfile) or '.', exist_ok=True)
        handler = TimedRotatingFileHandler(
            logfile, 
            when='midnight', 
            backupCount=30,
            encoding='utf-8'
        )
        fmt = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        handler.setFormatter(logging.Formatter(fmt))
        self.logger = logging.getLogger('ModbusLogger')
        self.logger.setLevel(logging.DEBUG)
        self.logger.addHandler(handler)

    def log(self, level, msg):
        self.logger.log(level, msg)

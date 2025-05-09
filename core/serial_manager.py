import serial
import serial.tools.list_ports
import logging
from core.protocol import Protocol

class SerialManager:
    def __init__(self, port=None, baudrate=9600, bytesize=8, parity='N', stopbits=1, timeout=1):
        self.ser = None
        self.port = port
        self.baudrate = baudrate
        self.bytesize = bytesize
        self.parity = parity
        self.stopbits = stopbits
        self.timeout = timeout
        self.logger = logging.getLogger(__name__)

    def open(self):
        try:
            if self.ser is not None and self.ser.is_open:
                self.ser.close()
            self.ser = serial.Serial(
                port=self.port,
                baudrate=self.baudrate,
                bytesize=self.bytesize,
                parity=self.parity,
                stopbits=self.stopbits,
                timeout=self.timeout
            )
            self.logger.info(f"串口 {self.port} 已打开")
            return True
        except Exception as e:
            self.logger.error(f"打开串口失败: {e}")
            return False

    def close(self):
        if self.ser is not None and self.ser.is_open:
            self.ser.close()
            self.logger.info(f"串口 {self.port} 已关闭")

    def is_open(self):
        return self.ser is not None and self.ser.is_open

    def write(self, data):
        if self.ser is not None and self.ser.is_open:
            try:
                self.ser.write(data)
                return True
            except Exception as e:
                self.logger.error(f"写入串口失败: {e}")
                return False
        return False

    def read(self, size):
        if self.ser is not None and self.ser.is_open:
            try:
                return self.ser.read(size)
            except Exception as e:
                self.logger.error(f"读取串口失败: {e}")
                return None
        return None

    def reset_input_buffer(self):
        if self.ser is not None and self.ser.is_open:
            try:
                self.ser.reset_input_buffer()
                return True
            except Exception as e:
                self.logger.error(f"清空输入缓冲区失败: {e}")
                return False
        return False

    @staticmethod
    def list_ports():
        ports = []
        for port in serial.tools.list_ports.comports():
            ports.append(port.device)
        return ports

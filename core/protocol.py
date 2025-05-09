import struct

class Protocol:
    """Modbus协议实现类"""
    @staticmethod
    def calc_crc(data: bytes) -> bytes:
        crc = 0xFFFF
        for a in data:
            crc ^= a
            for _ in range(8):
                if (crc & 0x0001):
                    crc >>= 1
                    crc ^= 0xA001
                else:
                    crc >>= 1
        return struct.pack('<H', crc)

    @staticmethod
    def build_rtu_request(slave_addr, func_code, start_addr, qty, data=b''):
        msg = bytes([slave_addr, func_code]) + start_addr.to_bytes(2, 'big') + qty.to_bytes(2, 'big') + data
        return msg + Protocol.calc_crc(msg)

    @staticmethod
    def parse_rtu_response(resp):
        """解析Modbus RTU响应，返回有效载荷（去除CRC）"""
        if len(resp) < 3:
            raise Exception("响应长度不足")
        # 计算CRC
        crc = Protocol.calc_crc(resp[:-2])
        # 比较接收的CRC（小端）
        if resp[-2] != crc[0] or resp[-1] != crc[1]:
            # 打印字节调试信息
            print(f"DEBUG: CRC检验失败 - 接收CRC={resp[-2:].hex()}, 计算CRC={crc.hex()}")
            raise Exception(f"CRC校验错误: 接收={resp[-2:].hex()}, 计算={crc.hex()}")
        # 打印解析调试信息
        print(f"DEBUG: RTU解析成功 - 载荷={resp[:-2].hex()}")
        return resp[:-2]

    @staticmethod
    def calc_lrc(data: bytes) -> bytes:
        lrc = sum(data) & 0xFF
        lrc = ((-lrc) & 0xFF)
        return bytes([lrc])

    @staticmethod
    def build_ascii_request(slave_addr, func_code, start_addr, qty, data=b''):
        body = bytes([slave_addr, func_code]) + start_addr.to_bytes(2, 'big') + qty.to_bytes(2, 'big') + data
        lrc = Protocol.calc_lrc(body)
        ascii_frame = b':' + body.hex().upper().encode() + lrc.hex().upper().encode() + b'\r\n'
        return ascii_frame

    @staticmethod
    def parse_ascii_response(resp):
        """解析Modbus ASCII响应，返回二进制载荷"""
        try:
            # 检查起始符和结束符
            if resp[0] != ord(':') or resp[-2:] != b'\r\n':
                raise Exception("起始符或结束符错误")
            # 转换ASCII为二进制
            hex_str = resp[1:-2].decode('ascii')
            # 计算LRC
            payload_hex = hex_str[:-2]
            recv_lrc = int(hex_str[-2:], 16)
            calc_lrc = Protocol.calc_lrc(bytes.fromhex(payload_hex))
            if recv_lrc != calc_lrc:
                # 打印字节调试信息
                print(f"DEBUG: LRC检验失败 - 接收LRC={hex_str[-2:]}, 计算LRC={calc_lrc:02x}")
                raise Exception(f"LRC校验错误: 接收={hex_str[-2:]}, 计算={calc_lrc:02x}")
            # 转换为二进制
            result = bytes.fromhex(payload_hex)
            # 打印解析调试信息
            print(f"DEBUG: ASCII解析成功 - 载荷={result.hex()}")
            return result
        except Exception as e:
            raise Exception(f"ASCII响应解析错误: {e}")

# 为了保持向后兼容性，保留原有的函数
def calc_crc(data: bytes) -> bytes:
    return Protocol.calc_crc(data)

def build_rtu_request(slave_addr, func_code, start_addr, qty, data=b''):
    return Protocol.build_rtu_request(slave_addr, func_code, start_addr, qty, data)

def parse_rtu_response(frame: bytes):
    return Protocol.parse_rtu_response(frame)

def calc_lrc(data: bytes) -> bytes:
    return Protocol.calc_lrc(data)

def build_ascii_request(slave_addr, func_code, start_addr, qty, data=b''):
    return Protocol.build_ascii_request(slave_addr, func_code, start_addr, qty, data)

def parse_ascii_response(frame: bytes):
    return Protocol.parse_ascii_response(frame)

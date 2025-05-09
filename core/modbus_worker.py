from PyQt5 import QtCore
import logging
from core.serial_manager import SerialManager
from core.protocol import Protocol
import pandas as pd
import time
import struct
from core.data_processor import decode_modbus_value, DISPLAY_SIGNED, DISPLAY_UNSIGNED, DISPLAY_HEX

class ModbusWorker(QtCore.QThread):
    comm_signal = QtCore.pyqtSignal(str, str)  # (类型, 内容)
    msg_signal = QtCore.pyqtSignal(str)
    data_signal = QtCore.pyqtSignal(int, str)  # (地址, 值)

    def __init__(self, ser, params_df, slave, mode, parent=None):
        super().__init__(parent)
        self.ser = ser
        # 强制addr列为无小数点字符串
        params_df['addr'] = params_df['addr'].apply(lambda x: str(int(float(x))) if pd.notna(x) and str(x).replace('.0','').isdigit() else str(x))
        
        # 查找数据类型列名（搜索多种可能的列名）
        data_type_col = None
        for col in params_df.columns:
            if str(col).strip().lower() in ['datatype', 'data_type', '数据类型', 'type', '类型']:
                data_type_col = col
                break
        
        if data_type_col is not None:
            # 确保dataType列处理正确
            params_df[data_type_col] = params_df[data_type_col].astype(str)
            # 修复NaN值和空字符串
            params_df[data_type_col] = params_df[data_type_col].apply(
                lambda x: 'UNSIGNED' if pd.isna(x) or x.upper() == 'NAN' or x.strip() == '' else x
            )
            
            # 特殊处理：处理所有SIGNED类型地址
            signed_addrs = []
            for idx, row in params_df.iterrows():
                if str(row[data_type_col]).strip().upper() == 'SIGNED':
                    signed_addrs.append(row['addr'])
            
            # 特殊处理：已知地址10000应该是SIGNED类型
            if '10000' in params_df['addr'].values:
                idx = params_df[params_df['addr'] == '10000'].index
                if len(idx) > 0:
                    params_df.loc[idx, data_type_col] = 'SIGNED'
                    if '10000' not in signed_addrs:
                        signed_addrs.append('10000')
                    print(f"DEBUG: 修复地址10000的数据类型为SIGNED")
            
            # 打印所有SIGNED地址，便于调试
            if signed_addrs:
                print(f"DEBUG: SIGNED类型的地址: {', '.join(signed_addrs)}")
            
            # 打印数据类型列的值分布，便于调试
            value_counts = params_df[data_type_col].value_counts()
            print(f"DEBUG: 数据类型分布: {value_counts.to_dict()}")
        
        self.data_type_col = data_type_col
        self.params_df = params_df
        self.slave = slave
        self.mode = mode
        self._running = True
        self.logger = logging.getLogger(__name__)
        self.poll_index = 0

    def stop(self):
        self._running = False

    def run(self):
        self.logger.info("开始轮询")
        data_type_col = self.data_type_col
        
        # 检查数据类型列是否存在，不存在则创建一个默认的
        if data_type_col is None or data_type_col not in self.params_df.columns:
            data_type_col = 'dataType'
            if data_type_col not in self.params_df.columns:
                self.params_df[data_type_col] = 'UNSIGNED'  # 默认为无符号
                # 特殊处理地址10000为SIGNED
                if '10000' in self.params_df['addr'].values:
                    idx = self.params_df[self.params_df['addr'] == '10000'].index
                    if len(idx) > 0:
                        self.params_df.loc[idx, data_type_col] = 'SIGNED'
        
        while self._running:
            try:
                # 检查串口和参数表
                if self.ser is None:
                    self.logger.error("串口未打开")
                    self.msg_signal.emit("串口未打开")
                    time.sleep(1)
                    continue

                if self.params_df is None or self.params_df.empty:
                    self.logger.error("参数表为空")
                    self.msg_signal.emit("参数表为空")
                    time.sleep(1)
                    continue

                # 获取所有需要轮询的地址并排序
                addrs = sorted(set(int(row['addr']) for _, row in self.params_df.iterrows() if str(row['addr']).isdigit()))
                if not addrs:
                    self.logger.error("没有有效的地址，无法轮询")
                    self.msg_signal.emit("参数表无有效地址，无法轮询")
                    time.sleep(1)
                    continue

                # 合并连续区间
                ranges = []
                start = prev = addrs[0]
                for addr in addrs[1:]:
                    if addr == prev + 1:
                        prev = addr
                    else:
                        ranges.append((start, prev))
                        start = prev = addr
                ranges.append((start, prev))

                # 对每个区间轮询，每个区间1秒发一个
                for start_addr, end_addr in ranges:
                    qty = end_addr - start_addr + 1
                    try:
                        self.ser.reset_input_buffer()
                        if self.mode == 'RTU':
                            req = Protocol.build_rtu_request(self.slave, 3, start_addr, qty)
                        else:
                            req = Protocol.build_ascii_request(self.slave, 3, start_addr, qty)
                        self.logger.info(f"发送请求: {req.hex(' ')}")
                        self.comm_signal.emit('send', req.hex(' '))
                        self.ser.write(req)

                        if self.mode == 'RTU':
                            min_len = 1 + 1 + 1 + 2 * qty + 2
                            resp = self.ser.read(min_len)
                            if resp:
                                self.logger.info(f"接收响应: {resp.hex(' ')} (len={len(resp)})")
                                self.comm_signal.emit('recv', f'{resp.hex(" ")} (len={len(resp)})')
                                if len(resp) >= min_len:
                                    try:
                                        payload = Protocol.parse_rtu_response(resp)
                                        self.logger.info(f"解析响应: {payload.hex(' ')}")
                                        byte_count = payload[2]
                                        data_bytes = payload[3:3+byte_count]
                                        self.logger.info(f"数据字节: {data_bytes.hex(' ')}")
                                        for i in range(qty):
                                            addr = start_addr + i
                                            addr_str = str(addr)
                                            
                                            # 使用DataFrame筛选更安全的方式，先确保addr是字符串类型
                                            mask = self.params_df['addr'].astype(str) == addr_str
                                            if not mask.any():
                                                self.logger.warning(f"地址 {addr} 未在参数表中找到")
                                                continue
                                                
                                            param_idx = self.params_df[mask].index[0]
                                            row = self.params_df.iloc[param_idx]
                                            
                                            # 获取数据类型，确保大写并去除空格
                                            data_type = 'UNSIGNED'  # 默认为无符号类型
                                            if data_type_col in row:
                                                dt_val = str(row[data_type_col]).strip().upper()
                                                if dt_val and dt_val != 'NAN':
                                                    data_type = dt_val
                                            
                                            # 调试信息：打印所有地址的数据类型
                                            print(f"DEBUG: 地址 {addr} 的数据类型是 {data_type}")
                                            
                                            # 特定地址强制使用SIGNED类型（临时调试措施）
                                            known_signed_addrs = [10000, 10001, 10002, 10003, 10004, 10005, 10006, 10007, 10008, 10009, 10010, 10011, 10012]  # 已知需要SIGNED类型的地址列表
                                            if addr in known_signed_addrs:
                                                data_type = 'SIGNED'
                                                print(f"DEBUG: 强制将地址 {addr} 的数据类型设置为SIGNED")
                                            
                                            # 确保所有标记为SIGNED的地址都能正确处理
                                            # 不再仅仅特殊处理地址10000
                                                
                                            self.logger.info(f"处理地址 {addr}: 数据类型={data_type}, 寄存器索引={i}")
                                            reg_bytes = data_bytes[2*i:2*(i+1)]
                                            # 额外的二进制数据调试
                                            unsigned_val = int.from_bytes(reg_bytes, byteorder='big', signed=False)
                                            signed_val = int.from_bytes(reg_bytes, byteorder='big', signed=True)
                                            print(f"DEBUG: 地址 {addr} 的原始字节={reg_bytes.hex()}, 无符号值={unsigned_val}, 带符号值={signed_val}")
                                            
                                            self.logger.info(f"寄存器原始字节: {reg_bytes.hex(' ')}")
                                            
                                            value = decode_modbus_value(reg_bytes, data_type, data_bytes, i, qty, param_idx, self.params_df)
                                            self.logger.info(f"解码结果: {value}")
                                            
                                            # 确保使用正确的数值格式，特别是负值
                                            if value and value.strip() and value.strip().replace('-', '').isdigit():
                                                # 可能是整数（包括负整数）
                                                self.logger.info(f"解码结果为整数: {value}")
                                                # 确保是整数字符串，不带小数点
                                                try:
                                                    value = str(int(float(value)))
                                                except:
                                                    pass
                                            
                                            # 增强调试输出
                                            print(f"DEBUG: 最终发送给UI的数据: 地址={addr}, 值={value}, 类型={type(value)}")
                                            
                                            self.data_signal.emit(addr, value)
                                    except Exception as e:
                                        self.logger.error(f"CRC校验失败: {e}")
                                        self.msg_signal.emit(f'地址 {start_addr} CRC校验失败: {e}')
                                else:
                                    self.logger.warning(f"地址 {start_addr} 响应长度不足: {len(resp)}/{min_len}")
                                    self.msg_signal.emit(f'地址 {start_addr} 响应长度不足: {len(resp)}/{min_len}')
                            else:
                                self.logger.warning(f"地址 {start_addr} 无响应")
                                self.msg_signal.emit(f'地址 {start_addr} 无响应')
                        else:
                            # ASCII模式处理（类似于RTU模式的修改）
                            resp = self.ser.read(64)
                            if resp:
                                self.logger.info(f"接收响应: {resp.hex(' ')} (len={len(resp)})")
                                self.comm_signal.emit('recv', f'{resp.hex(" ")} (len={len(resp)})')
                                try:
                                    payload = Protocol.parse_ascii_response(resp)
                                    self.logger.info(f"解析响应: {payload.hex(' ')}")
                                    byte_count = payload[2]
                                    data_bytes = payload[3:3+byte_count]
                                    self.logger.info(f"数据字节: {data_bytes.hex(' ')}")
                                    for i in range(qty):
                                        addr = start_addr + i
                                        addr_str = str(addr)
                                        
                                        # 使用更安全的DataFrame筛选方式
                                        mask = self.params_df['addr'].astype(str) == addr_str
                                        if not mask.any():
                                            self.logger.warning(f"地址 {addr} 未在参数表中找到")
                                            continue
                                            
                                        param_idx = self.params_df[mask].index[0]
                                        row = self.params_df.iloc[param_idx]
                                        
                                        # 获取数据类型，确保大写并去除空格
                                        data_type = 'UNSIGNED'  # 默认为无符号类型
                                        if data_type_col in row:
                                            dt_val = str(row[data_type_col]).strip().upper()
                                            if dt_val and dt_val != 'NAN':
                                                data_type = dt_val
                                        
                                        # 调试信息：打印所有地址的数据类型
                                        print(f"DEBUG: 地址 {addr} 的数据类型是 {data_type}")
                                        
                                        # 特定地址强制使用SIGNED类型（临时调试措施）
                                        known_signed_addrs = [10000, 10001, 10002, 10003, 10004, 10005, 10006, 10007, 10008, 10009, 10010, 10011, 10012]  # 已知需要SIGNED类型的地址列表
                                        if addr in known_signed_addrs:
                                            data_type = 'SIGNED'
                                            print(f"DEBUG: 强制将地址 {addr} 的数据类型设置为SIGNED")
                                        
                                        # 确保所有标记为SIGNED的地址都能正确处理
                                        # 不再仅仅特殊处理地址10000
                                                
                                        reg_bytes = data_bytes[2*i:2*(i+1)]
                                        # 额外的二进制数据调试
                                        unsigned_val = int.from_bytes(reg_bytes, byteorder='big', signed=False)
                                        signed_val = int.from_bytes(reg_bytes, byteorder='big', signed=True)
                                        print(f"DEBUG: 地址 {addr} 的原始字节={reg_bytes.hex()}, 无符号值={unsigned_val}, 带符号值={signed_val}")
                                        
                                        self.logger.info(f"寄存器原始字节: {reg_bytes.hex(' ')}")
                                        
                                        value = decode_modbus_value(reg_bytes, data_type, data_bytes, i, qty, param_idx, self.params_df)
                                        self.logger.info(f"解码结果: {value}")
                                        
                                        # 确保使用正确的数值格式，特别是负值
                                        if value and value.strip() and value.strip().replace('-', '').isdigit():
                                            # 可能是整数（包括负整数）
                                            self.logger.info(f"解码结果为整数: {value}")
                                            # 确保是整数字符串，不带小数点
                                            try:
                                                value = str(int(float(value)))
                                            except:
                                                pass
                                        
                                        # 增强调试输出
                                        print(f"DEBUG: 最终发送给UI的数据: 地址={addr}, 值={value}, 类型={type(value)}")
                                        
                                        self.data_signal.emit(addr, value)
                                except Exception as e:
                                    self.logger.error(f"ASCII校验失败: {e}")
                                    self.msg_signal.emit(f'地址 {start_addr} ASCII校验失败: {e}')
                            else:
                                self.logger.warning(f"地址 {start_addr} 无响应")
                                self.msg_signal.emit(f'地址 {start_addr} 无响应')
                        time.sleep(1)  # 每个区间1秒发一个
                    except Exception as e:
                        self.logger.error(f"通信错误: {e}")
                        self.msg_signal.emit(f'地址 {start_addr} 通信错误: {e}')
            except Exception as e:
                self.logger.error(f"轮询主循环异常: {e}", exc_info=True)
                self.msg_signal.emit(f"轮询主循环异常: {e}")
                time.sleep(1)

    def decode_modbus_value(self, reg_bytes, data_type, data_bytes, i, qty, param_idx):
        """解码Modbus寄存器值"""
        try:
            if data_type == 'UNSIGNED':
                return str(int.from_bytes(reg_bytes, byteorder='big', signed=False))
            elif data_type == 'SIGNED':
                # 确保signed=True以正确处理负数
                value = int.from_bytes(reg_bytes, byteorder='big', signed=True)
                return str(value)
            elif data_type == 'FLOAT32':
                # FLOAT32需要2个寄存器（4字节）
                if i + 1 < qty:
                    reg_bytes2 = data_bytes[2*i:2*(i+2)]
                    if len(reg_bytes2) == 4:
                        val = struct.unpack('>f', reg_bytes2)[0]
                        return str(val)
                    else:
                        return '数据不足'
                else:
                    return '数据不足'
            elif data_type == 'HEX':
                return reg_bytes.hex()
            else:
                return reg_bytes.hex()
        except Exception as e:
            self.logger.error(f"解码错误: {e}")
            return f'解码错: {e}' 
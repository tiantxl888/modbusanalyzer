import pandas as pd
from PyQt5 import QtWidgets, QtCore
import logging
from core.protocol import Protocol
import struct

# 显示策略常量
DISPLAY_SIGNED = 'SIGNED'      # 显示带符号十进制
DISPLAY_HEX = 'HEX'            # 显示16进制
DISPLAY_UNSIGNED = 'UNSIGNED'  # 默认无符号十进制

# 模块级解码函数，便于外部import
def decode_modbus_value(reg_bytes, data_type, payload, i, qty, param_idx, df):
    """解码Modbus寄存器值（模块级）"""
    try:
        # 获取当前处理的地址（方便调试）
        current_addr = None
        start_addr = None
        if df is not None and param_idx is not None and 'addr' in df.columns:
            try:
                current_addr = int(df.iloc[param_idx]['addr'])
                # 尝试从payload或其他参数推断起始地址
                if len(payload) >= 3 and i == 0:  # Modbus功能码3读取响应
                    start_addr = current_addr
            except:
                pass

        # 首先检查data_type是否为NaN或"NAN"
        if data_type is None or pd.isna(data_type) or str(data_type).strip().upper() == 'NAN' or str(data_type).strip() == '':
            # 尝试从UI配置或表中获取真实类型
            if df is not None and param_idx is not None:
                data_type_col = None
                # 查找可能的数据类型列
                for col in df.columns:
                    if str(col).strip().lower() in ['datatype', 'data_type', '数据类型', 'type', '类型']:
                        data_type_col = col
                        break
                
                if data_type_col and data_type_col in df.columns:
                    data_type = str(df.iloc[param_idx][data_type_col]).strip().upper()
                    if pd.isna(data_type) or data_type == 'NAN' or data_type == '':
                        # 如果这是一个需要SIGNED处理的地址，强制使用SIGNED
                        if current_addr == 10000:
                            data_type = 'SIGNED'
                            print(f"DEBUG: 为地址10000强制使用SIGNED类型")
                        else:
                            data_type = DISPLAY_UNSIGNED  # 默认无符号
                else:
                    data_type = DISPLAY_UNSIGNED
            else:
                data_type = DISPLAY_UNSIGNED
        
        # 统一将data_type强制转换为大写并去空格，严格匹配常量
        data_type = str(data_type).strip().upper()
        
        # 添加详细日志帮助调试
        if current_addr:
            print(f"DEBUG: 解码地址 {current_addr} 使用数据类型={data_type}, 原始字节={reg_bytes.hex()}")
        
        # 特殊处理：特定的地址（如第二行）强制使用SIGNED类型
        # 这是临时调试措施，帮助识别哪些地址需要SIGNED类型
        known_signed_addrs = [10000, 10001, 10002, 10003, 10004, 10005, 10006, 10007, 10008, 10009, 10010, 10011, 10012]
        if current_addr in known_signed_addrs:
            data_type = DISPLAY_SIGNED
            print(f"DEBUG: 强制地址 {current_addr} 使用SIGNED类型")
        
        if data_type == DISPLAY_UNSIGNED:
            # 解析为无符号整数
            value = int.from_bytes(reg_bytes, byteorder='big', signed=False)
            print(f"DEBUG: UNSIGNED解析: 原始字节={reg_bytes.hex()}, 解析值={value}")
            return str(value)
        elif data_type == DISPLAY_SIGNED:
            # 确保SIGNED类型使用signed=True
            # 读取两个字节并将其解释为带符号整数
            value = int.from_bytes(reg_bytes, byteorder='big', signed=True)
            print(f"DEBUG: SIGNED解析: 原始字节={reg_bytes.hex()}, 解析值={value}")
            # 返回数值字符串
            return str(value)
        elif data_type == 'FLOAT32':
            # 调试信息
            print(f"DEBUG: 尝试解析FLOAT32: 原始字节={reg_bytes.hex()}")
            # FLOAT32需要2个寄存器（4字节）
            if i + 1 < qty:
                reg_bytes2 = payload[3+2*i:3+2*(i+2)]
                if len(reg_bytes2) == 4:
                    val = struct.unpack('>f', reg_bytes2)[0]
                    print(f"DEBUG: FLOAT32解析: 原始字节={reg_bytes2.hex()}, 解析值={val}")
                    return str(val)
                else:
                    return '数据不足'
            else:
                return '数据不足'
        elif data_type == DISPLAY_HEX:
            print(f"DEBUG: HEX解析: 原始字节={reg_bytes.hex()}")
            return f"0x{reg_bytes.hex()}H"
        else:
            # 如果类型不匹配任何已知类型，记录并默认为UNSIGNED
            logging.warning(f"未知的数据类型: {data_type}，默认为无符号")
            # 如果数据类型包含"SIGNED"字样，尝试按带符号处理
            if "SIGNED" in data_type:
                value = int.from_bytes(reg_bytes, byteorder='big', signed=True)
                print(f"DEBUG: 检测到可能的SIGNED类型({data_type}): 原始字节={reg_bytes.hex()}, 带符号解析值={value}")
                return str(value)
            return str(int.from_bytes(reg_bytes, byteorder='big', signed=False))
    except Exception as e:
        logging.error(f"解码错误: {e}")
        return f'解码错: {e}'

class DataProcessor:
    @staticmethod
    def decode_modbus_value(reg_bytes, data_type, payload, i, qty, param_idx, df):
        """解码Modbus寄存器值（兼容旧用法）"""
        return decode_modbus_value(reg_bytes, data_type, payload, i, qty, param_idx, df)

    @staticmethod
    def update_param_value(addr, value, param_tables, current_sheet):
        """只更新参数值到表格，不写回Excel"""
        if current_sheet not in param_tables:
            return
            
        print(f"DEBUG: update_param_value - 更新 sheet={current_sheet}, 地址={addr}, 值={value}")
        
        # 检查值是否为负数或其他特殊处理
        try:
            # 尝试转换为数值，以便确认是否为负值
            value_num = float(value)
            if value_num < 0:
                print(f"DEBUG: 检测到负值 {value_num} 用于地址 {addr}")
        except:
            # 忽略无法转换的值
            pass
        
        table = param_tables[current_sheet]
        
        # All Parameters 表格和分组表格的处理不同
        if current_sheet == 'All Parameters':
            # 特殊处理 All Parameters 表
            column_headers = [table.horizontalHeaderItem(i).text() if table.horizontalHeaderItem(i) else "" for i in range(table.columnCount())]
            addr_col_idx = column_headers.index('Address') if 'Address' in column_headers else -1
            value_col_idx = column_headers.index('Current Value') if 'Current Value' in column_headers else -1
            
            # 如果找不到列，直接返回
            if addr_col_idx == -1 or value_col_idx == -1:
                print(f"WARNING: update_param_value - All Parameters表中找不到Address或Current Value列")
                return
                
            # 查找匹配地址的行
            for r in range(table.rowCount()):
                addr_item = table.item(r, addr_col_idx)
                if addr_item and addr_item.text() == str(addr):
                    # 只更新Current Value列
                    if table.item(r, value_col_idx) is None:
                        new_item = QtWidgets.QTableWidgetItem(value)
                        new_item.setTextAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
                        table.setItem(r, value_col_idx, new_item)
                    else:
                        table.item(r, value_col_idx).setText(value)
                    print(f"DEBUG: update_param_value - 更新All Parameters表, 行={r}, 值={value}")
        else:
            # 处理分组表格
            group_size = 3  # 每组3列
            
            # 计算表格有多少组
            groups = table.columnCount() // group_size
            print(f"DEBUG: update_param_value - 表格 {current_sheet} 有 {groups} 组")
            
            updated = False
            for r in range(table.rowCount()):
                for g in range(groups):
                    base_col = g * group_size
                    addr_col = base_col + 1  # 地址在每组的第2列
                    value_col = base_col + 2  # 当前值在每组的第3列
                    
                    addr_item = table.item(r, addr_col)
                    if addr_item and addr_item.text() == str(addr):
                        # 找到了匹配的地址项
                        print(f"DEBUG: update_param_value - 找到地址 {addr} 在 row={r}, group={g}")
                        
                        # 更新值
                        if table.item(r, value_col) is None:
                            new_item = QtWidgets.QTableWidgetItem(value)
                            new_item.setTextAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
                            table.setItem(r, value_col, new_item)
                        else:
                            table.item(r, value_col).setText(value)
                        
                        updated = True
                        break
                if updated:
                    break
                    
            if not updated:
                print(f"WARNING: update_param_value - 未找到地址 {addr} 在 sheet={current_sheet}")

    @staticmethod
    def build_param_tables(sheets, excel_file):
        """构建参数表格"""
        param_tables = {}
        param_dfs = {}
        for sheet in sheets:
            try:
                df = pd.read_excel(excel_file, sheet_name=sheet, header=1)
            except Exception:
                continue
            if 'addr' not in df.columns:
                continue
            df = df[df['addr'].apply(lambda x: pd.notna(x) and str(x).replace('.0','').isdigit())].copy()
            df['addr'] = df['addr'].apply(lambda x: str(int(float(x))) if pd.notna(x) and str(x).replace('.0','').isdigit() else '')
            show_cols = ['name', 'addr', '当前值']
            if '当前值' not in df.columns:
                df['当前值'] = ''
            if sheet in param_dfs:
                old_df = param_dfs[sheet]
                for idx, row in df.iterrows():
                    addr = row['addr']
                    if addr in old_df['addr'].values:
                        old_val = old_df.loc[old_df['addr'] == addr, '当前值'].values
                        if len(old_val) > 0:
                            df.at[idx, '当前值'] = old_val[0]
            param_dfs[sheet] = df.reset_index(drop=True)
        return param_tables, param_dfs 

    @staticmethod
    def process_data(df, ser, slave, mode):
        """处理数据"""
        try:
            # 检查串口和参数表
            if ser is None:
                logging.error("串口未打开")
                return False

            if df is None or df.empty:
                logging.error("参数表为空")
                return False

            # 获取地址列表
            addrs = []
            try:
                addrs = [int(row['addr']) for _, row in df.iterrows() if str(row['addr']).isdigit()]
                logging.info(f"找到地址列表: {addrs}")
            except Exception as e:
                logging.error(f"获取地址列表失败: {e}")
                return False

            if not addrs:
                logging.error("没有有效的地址，无法轮询")
                return False

            # 按地址排序
            addrs.sort()
            
            # 合并连续地址
            ranges = []
            start = addrs[0]
            prev = addrs[0]
            for addr in addrs[1:]:
                if addr == prev + 1:
                    prev = addr
                else:
                    ranges.append((start, prev))
                    start = addr
                    prev = addr
            ranges.append((start, prev))

            # 对每个区间进行轮询
            for start_addr, end_addr in ranges:
                qty = end_addr - start_addr + 1
                try:
                    ser.reset_input_buffer()
                    if mode == 'RTU':
                        req = Protocol.build_rtu_request(slave, 3, start_addr, qty)
                    else:
                        req = Protocol.build_ascii_request(slave, 3, start_addr, qty)
                    logging.info(f"发送请求: {req.hex(' ')}")
                    ser.write(req)

                    if mode == 'RTU':
                        # 标准Modbus RTU响应长度: 1(地址)+1(功能码)+1(字节数)+2*qty(数据)+2(CRC)
                        min_len = 1 + 1 + 1 + 2 * qty + 2
                        resp = ser.read(min_len)
                        logging.info(f"接收响应: {resp.hex(' ')} (len={len(resp)})")
                        if len(resp) >= min_len:
                            try:
                                payload = Protocol.parse_rtu_response(resp)
                                logging.info(f"解析响应: {payload.hex(' ')}")
                                # payload[2]为字节数，实际数据长度=payload[2]
                                byte_count = payload[2]
                                data_bytes = payload[3:3+byte_count]
                                logging.info(f"数据字节: {data_bytes.hex(' ')}")
                                for i in range(qty):
                                    addr = start_addr + i
                                    param_idx = df[df['addr'] == addr].index[0]
                                    data_type = str(df.iloc[param_idx].get('dataType', 'UNSIGNED')).upper()
                                    reg_bytes = data_bytes[2*i:2*(i+1)]
                                    logging.info(f"处理地址 {addr}: 数据类型={data_type}, 寄存器值={reg_bytes.hex(' ')}")
                                    value = DataProcessor.decode_modbus_value(reg_bytes, data_type, payload, i, qty, param_idx, df)
                                    logging.info(f"解码结果: {value}")
                                    df.at[param_idx, 'value'] = value
                            except Exception as e:
                                logging.error(f"CRC校验失败: {e}")
                                return False
                        elif len(resp) == 0:
                            logging.warning(f"地址 {start_addr} 无响应")
                            return False
                        else:
                            logging.warning(f"地址 {start_addr} 响应长度不足: {len(resp)}/{min_len}")
                            return False
                    else:
                        resp = ser.read(64)
                        logging.info(f"接收响应: {resp.hex(' ')} (len={len(resp)})")
                        try:
                            payload = Protocol.parse_ascii_response(resp)
                            logging.info(f"解析响应: {payload.hex(' ')}")
                            byte_count = payload[2]
                            data_bytes = payload[3:3+byte_count]
                            logging.info(f"数据字节: {data_bytes.hex(' ')}")
                            for i in range(qty):
                                addr = start_addr + i
                                param_idx = df[df['addr'] == addr].index[0]
                                data_type = str(df.iloc[param_idx].get('dataType', 'UNSIGNED')).upper()
                                reg_bytes = data_bytes[2*i:2*(i+1)]
                                logging.info(f"处理地址 {addr}: 数据类型={data_type}, 寄存器值={reg_bytes.hex(' ')}")
                                value = DataProcessor.decode_modbus_value(reg_bytes, data_type, payload, i, qty, param_idx, df)
                                logging.info(f"解码结果: {value}")
                                df.at[param_idx, 'value'] = value
                        except Exception as e:
                            logging.error(f"ASCII校验失败: {e}")
                            return False
                except Exception as e:
                    logging.error(f"通信错误: {e}")
                    return False

            return True
        except Exception as e:
            logging.error(f"处理数据异常: {e}", exc_info=True)
            return False 
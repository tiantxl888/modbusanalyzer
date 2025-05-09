#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from PyQt5 import QtWidgets, QtCore, QtGui
import logging
import os
import pandas as pd
from core.serial_manager import SerialManager
from core.modbus_worker import ModbusWorker
from core.data_processor import DataProcessor
from core.protocol import Protocol
from core.project_manager import ProjectManager
from ui.components import SerialConfigWidget, ParamTableWidget, CommLogWidget
from utils.excel_manager import ExcelManager
from utils.log_manager import LogManager
from PyQt5.QtCore import QThread, pyqtSignal, QTimer
import re

# 导入绘图相关库
import pyqtgraph as pg
import numpy as np
import time
import requests
import webbrowser
import tempfile
import shutil
import sys
from PyQt5.QtWidgets import QProgressDialog
from core.plugin_manager import PluginManager

# 配置日志
log_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'modbus.log')
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file, encoding='gb18030'),
        logging.StreamHandler()
    ]
)

LOCAL_VERSION = "0.0.1"

class TimeAxisItem(pg.AxisItem):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

    def tickStrings(self, values, scale, spacing):
        return [time.strftime('%H:%M:%S', time.gmtime(v)) for v in values]

class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setStyleSheet('''
            QMainWindow { background: #f0f4fa; }
            QTableWidget { background: #ffffff; alternate-background-color: #f5faff; gridline-color: #90caf9; }
            QHeaderView::section { background: #1976d2; color: white; font-weight: bold; border-radius: 4px; }
            QPushButton { background: #1976d2; color: white; border-radius: 6px; padding: 6px 16px; font-weight: bold; }
            QPushButton:hover { background: #1565c0; }
            QComboBox, QLineEdit { border-radius: 4px; padding: 2px 8px; }
        ''')
        self.setWindowTitle("Modbus Analyzer")
        self.setGeometry(100, 100, 1200, 800)
        
        # 创建插件管理器
        self.plugin_manager = PluginManager(self)
        
        # 初始化UI
        self._init_menu()
        self._init_main_layout()
        self.statusBar().showMessage('Ready')

        # 状态变量
        self.ser = None
        self.polling = False
        self.current_sheet = None
        self.param_tables = {}
        self.param_dfs = {}

        # 检查Excel文件
        self._check_excel_file()

        # 加载插件
        self.plugin_manager.load_plugins()

    def _init_menu(self):
        menubar = self.menuBar()
        file_menu = menubar.addMenu('File')
        file_menu.addAction('New Project')
        file_menu.addAction('Open Project')
        file_menu.addAction('Save Project')
        file_menu.addSeparator()
        import_action = file_menu.addAction('Import Excel')
        import_action.triggered.connect(self.import_excel)
        file_menu.addAction('Export Excel')
        file_menu.addSeparator()
        file_menu.addAction('Exit', self.close)
        tool_menu = menubar.addMenu('Tools')
        # 删除串口配置菜单项
        # tool_menu.addAction('Serial Config')
        
        # 增加曲线显示菜单项
        chart_action = tool_menu.addAction('Chart View')
        chart_action.triggered.connect(self.show_chart_view)
        
        tool_menu.addAction('Language')
        open_cfg_action = tool_menu.addAction('Open Config File')
        open_cfg_action.triggered.connect(self.open_config_file)
        restart_action = tool_menu.addAction('Restart')
        restart_action.triggered.connect(self.restart_app)
        help_menu = menubar.addMenu('Help')
        help_menu.addAction('Manual')
        help_menu.addAction('About')
        # 新增：检查更新
        check_update_action = help_menu.addAction('检查更新')
        check_update_action.triggered.connect(self.check_update)

    def _init_main_layout(self):
        central = QtWidgets.QWidget()
        vbox = QtWidgets.QVBoxLayout(central)
        
        # 串口设置区
        self.serial_config = SerialConfigWidget()
        self.open_btn = QtWidgets.QPushButton('Open Port')
        self.poll_btn = QtWidgets.QPushButton('Start Polling')
        self.open_btn.clicked.connect(self.toggle_port)
        self.poll_btn.clicked.connect(self.toggle_polling)
        self.poll_btn.setEnabled(False)
        
        # 添加图表按钮
        self.chart_btn = QtWidgets.QPushButton('Chart View')
        self.chart_btn.clicked.connect(self.show_chart_view)
        self.chart_btn.setToolTip('显示选中数据的实时曲线图')
        
        hbox = QtWidgets.QHBoxLayout()
        hbox.addWidget(self.serial_config)
        hbox.addWidget(self.open_btn)
        hbox.addWidget(self.poll_btn)
        hbox.addWidget(self.chart_btn)
        vbox.addLayout(hbox)

        # 自动加载串口设置
        self.load_serial_config_from_excel()

        # 中部：TabWidget
        self.tab_widget = QtWidgets.QTabWidget()
        self._build_all_tables()
        vbox.addWidget(self.tab_widget, stretch=1)

        # 底部通讯日志区
        self.comm_log = CommLogWidget()
        self.comm_log.customContextMenuRequested.connect(self.show_comm_log_menu)
        vbox.addWidget(self.comm_log)

        self.setCentralWidget(central)

    def _build_all_tables(self):
        self.tab_widget.clear()
        # 调试打印
        print("DEBUG: build_all_tables - 开始构建表格")
        
        self.param_tables = {}
        self.param_dfs = {}
        valid_group_count = 0
        all_valid_dfs = []
        try:
            xls = pd.ExcelFile('config_and_params.xlsx')
            sheets = xls.sheet_names
        except Exception as e:
            logging.error(f"打开Excel文件失败: {e}")
            sheets = []
        finally:
            try:
                xls.close()
            except Exception:
                pass
        for sheet in sheets:
            try:
                df = pd.read_excel('config_and_params.xlsx', sheet_name=sheet, header=1)
                if 'name' not in df.columns or 'addr' not in df.columns:
                    continue
                group_indices = []
                group_names = []
                for idx, row in df.iterrows():
                    name_val = str(row['name']).strip() if pd.notna(row['name']) else ''
                    addr_val = str(row['addr']).strip() if pd.notna(row['addr']) else ''
                    is_addr_invalid = (not addr_val or addr_val.lower() in ['nan', '']) or not str(addr_val).replace('.0','').isdigit()
                    if name_val and is_addr_invalid:
                        group_indices.append(idx)
                        group_names.append(name_val)
                group_indices.append(len(df))
                for i in range(len(group_names)):
                    start = group_indices[i] + 1
                    end = group_indices[i+1] if i+1 < len(group_indices) else len(df)
                    group_df = df.iloc[start:end].copy()
                    group_df = group_df[group_df['addr'].apply(lambda x: pd.notna(x) and str(x).replace('.0','').isdigit())]
                    group_df['addr'] = group_df['addr'].apply(lambda x: str(int(float(x))) if pd.notna(x) and str(x).replace('.0','').isdigit() else '')
                    group_df = group_df[group_df['addr'] != '']
                    if group_df.empty:
                        continue
                    
                    # 修复dataType列，确保正确处理数据类型
                    if 'dataType' in group_df.columns:
                        # 修复NaN值和空字符串
                        group_df['dataType'] = group_df['dataType'].apply(
                            lambda x: 'UNSIGNED' if pd.isna(x) or str(x).upper() == 'NAN' or str(x).strip() == '' else str(x)
                        )
                        # 特殊处理：标记所有需要SIGNED类型的地址
                        known_signed_addrs = ['10000', '10001', '10002', '10003', '10004', '10005', '10006', '10007', '10008', '10009', '10010', '10011', '10012']
                        for addr in known_signed_addrs:
                            if addr in group_df['addr'].values:
                                idx = group_df[group_df['addr'] == addr].index
                                if len(idx) > 0:
                                    group_df.loc[idx, 'dataType'] = 'SIGNED'
                                    print(f"DEBUG: 表格构建时修复地址{addr}的数据类型为SIGNED")
                        
                        # 打印所有SIGNED类型的地址，便于调试
                        signed_addrs = group_df[group_df['dataType'].str.upper() == 'SIGNED']['addr'].tolist()
                        if signed_addrs:
                            print(f"DEBUG: 检测到SIGNED类型的地址: {', '.join(signed_addrs)}")
                        
                        # 确保所有SIGNED的dataType大写，以便后续处理
                        group_df.loc[group_df['dataType'].str.upper() == 'SIGNED', 'dataType'] = 'SIGNED'
                    
                    if '当前值' not in group_df.columns:
                        group_df['当前值'] = ''
                    all_valid_dfs.append(group_df)
                    show_cols = ['name', 'addr', '当前值']
                    valid_df = group_df[show_cols].copy()
                    group_size = 3
                    group_count = self._get_group_count()
                    n = len(valid_df)
                    row_count = (n + group_count - 1) // group_count
                    table = ParamTableWidget()
                    table.setRowCount(row_count)
                    table.setColumnCount(group_count * group_size)
                    headers = []
                    for j in range(group_count):
                        headers.extend(['Data', 'Address', 'Value'])
                    table.setHorizontalHeaderLabels(headers)
                    for idx2 in range(n):
                        r = idx2 % row_count
                        g = idx2 // row_count
                        base_col = g * group_size
                        for c, col in enumerate(show_cols):
                            item = QtWidgets.QTableWidgetItem(str(valid_df.iloc[idx2].get(col, '')))
                            item.setTextAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
                            table.setItem(r, base_col + c, item)
                    table.resizeRowsToContents()
                    table.resizeColumnsToContents()
                    self.tab_widget.addTab(table, group_names[i])
                    self.param_tables[group_names[i]] = table
                    self.param_dfs[group_names[i]] = valid_df
                    valid_group_count += 1
            except Exception as e:
                logging.error(f"处理Sheet {sheet}失败: {e}")
                continue
        # 添加全部通讯数据Tab（显示所有字段）
        if valid_group_count > 0 and all_valid_dfs:
            all_params = pd.concat(all_valid_dfs, ignore_index=True)
            print(f"DEBUG: 合并所有参数表前的列: {all_params.columns.tolist()}")
            
            # 重命名列
            all_cols = list(all_params.columns)
            rename_map = {}
            
            if '当前值' in all_cols:
                rename_map['当前值'] = 'Current Value'
                
            if 'name' in all_cols:
                rename_map['name'] = 'Name'
                
            if 'addr' in all_cols:
                rename_map['addr'] = 'Address'
            
            # 应用重命名
            if rename_map:
                all_params = all_params.rename(columns=rename_map)
                
            all_cols = list(all_params.columns)
            print(f"DEBUG: 重命名后的列: {all_cols}")
            
            # 确保Current Value列为空，不继承attribute列的值
            if 'Current Value' in all_params.columns:
                all_params['Current Value'] = ''
                
            # 打印一行数据示例，便于调试
            if not all_params.empty:
                print(f"DEBUG: 第一行数据示例: {all_params.iloc[0].to_dict()}")
            
            n = len(all_params)
            table = ParamTableWidget()
            table.setRowCount(n)
            table.setColumnCount(len(all_cols))
            table.setHorizontalHeaderLabels(all_cols)
            for idx in range(n):
                for col_idx, col in enumerate(all_cols):
                    item = QtWidgets.QTableWidgetItem(str(all_params.iloc[idx].get(col, '')))
                    item.setTextAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
                    table.setItem(idx, col_idx, item)
            table.resizeRowsToContents()
            table.resizeColumnsToContents()
            self.tab_widget.insertTab(0, table, 'All Parameters')
            self.param_tables['All Parameters'] = table
            self.param_dfs['All Parameters'] = all_params
            self.current_sheet = self.tab_widget.tabText(0)
            # 设置"All Parameters"为当前活动标签页
            self.tab_widget.setCurrentIndex(0)
        if valid_group_count == 0:
            self.statusBar().showMessage('No valid parameter group found')
        else:
            self.statusBar().showMessage('Parameter groups loaded')

    def _get_group_count(self):
        # 根据当前tabwidget宽度动态计算每行能显示多少组
        table_width = self.tab_widget.width() if self.tab_widget.width() > 0 else 900
        group_width = 240  # 每组大约240像素宽
        return max(1, table_width // group_width)

    def on_tab_changed(self, idx):
        sheet = self.tab_widget.tabText(idx)
        if sheet not in self.param_tables:
            try:
                df = pd.read_excel('config_and_params.xlsx', sheet_name=sheet, header=1)
                if 'addr' not in df.columns:
                    logging.warning(f"Sheet {sheet} has no 'addr' column")
                    return
                df = df[df['addr'].apply(lambda x: pd.notna(x) and str(x).replace('.0','').isdigit())].copy()
                df['addr'] = df['addr'].apply(lambda x: str(int(float(x))) if pd.notna(x) and str(x).replace('.0','').isdigit() else '')
                show_cols = ['name', 'addr', '当前值']
                if '当前值' not in df.columns:
                    df['当前值'] = ''
                valid_df = df[show_cols].copy()
                if valid_df.empty:
                    return
                group_size = 3
                group_count = self._get_group_count()
                n = len(valid_df)
                row_count = (n + group_count - 1) // group_count
                table = ParamTableWidget()
                table.setRowCount(row_count)
                table.setColumnCount(group_count * group_size)
                headers = []
                for i in range(group_count):
                    headers.extend(['Name', 'Address', 'Current Value'])
                table.setHorizontalHeaderLabels(headers)
                for idx in range(n):
                    r = idx % row_count
                    g = idx // row_count
                    base_col = g * group_size
                    for c, col in enumerate(show_cols):
                        col_en = 'Name' if col == 'name' else ('Address' if col == 'addr' else ('Current Value' if col == '当前值' else col))
                        item = QtWidgets.QTableWidgetItem(str(valid_df.iloc[idx].get(col, '')))
                        item.setTextAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
                        table.setItem(r, base_col + c, item)
                table.resizeRowsToContents()
                table.resizeColumnsToContents()
                self.tab_widget.removeTab(idx)
                self.tab_widget.insertTab(idx, table, sheet)
                self.param_tables[sheet] = table
                self.param_dfs[sheet] = valid_df
            except Exception as e:
                logging.error(f"Failed to lazy load sheet {sheet}: {e}")
        self.current_sheet = sheet

    def toggle_port(self):
        if self.ser is None:
            try:
                config = self.serial_config.get_config()
                logging.info(f"准备打开串口，参数: {config}")
                self.serial_manager = SerialManager(
                    port=config['port'],
                    baudrate=config['baudrate'],
                    bytesize=config['bytesize'],
                    parity=config['parity'],
                    stopbits=config['stopbits'],
                    timeout=1
                )
                if self.serial_manager.open():
                    self.ser = self.serial_manager.ser
                    logging.info(f"串口打开成功: {config['port']}")
                    self.save_serial_config_to_excel(config)
                    self.open_btn.setText('Close Port')
                    self.poll_btn.setEnabled(True)
                    self.statusBar().showMessage('串口已打开')
                    self.serial_config.set_locked(True)
                else:
                    raise Exception("串口打开失败")
            except Exception as e:
                logging.error(f"打开串口失败: {e}")
                QtWidgets.QMessageBox.critical(self, '错误', f'打开串口失败: {e}')
        else:
            if self.polling:
                self.toggle_polling()
            try:
                logging.info(f"准备关闭串口")
                self.serial_manager.close()
                logging.info(f"串口关闭成功")
            except Exception as e:
                logging.error(f"关闭串口失败: {e}")
            self.ser = None
            self.open_btn.setText('Open Port')
            self.poll_btn.setEnabled(False)
            self.statusBar().showMessage('串口已关闭')
            self.serial_config.set_locked(False)

    def toggle_polling(self):
        if not self.polling:
            try:
                if not self.param_dfs:
                    QtWidgets.QMessageBox.warning(self, '警告', '请先导入Excel数据')
                    return
                if self.ser is None:
                    QtWidgets.QMessageBox.warning(self, '警告', '请先打开串口')
                    return
                import pandas as pd
                all_params = pd.concat(self.param_dfs.values(), ignore_index=True)
                all_params = all_params[all_params['addr'].apply(lambda x: str(x).replace('.0','').isdigit())].copy()
                all_params['addr'] = all_params['addr'].apply(lambda x: int(float(x)))
                all_params.sort_values('addr', inplace=True)
                self.poll_worker = ModbusWorker(
                    self.ser,
                    all_params,
                    1,
                    self.serial_config.mode_cb.currentText()
                )
                self.poll_worker.comm_signal.connect(self.on_comm_signal)
                self.poll_worker.msg_signal.connect(self.on_msg_signal)
                self.poll_worker.data_signal.connect(self.on_data_signal)
                self.poll_worker.start()
                self.polling = True
                self.poll_btn.setText('Stop Polling')
                self.poll_btn.setEnabled(True)
                self.open_btn.setEnabled(False)
                self.statusBar().showMessage('开始轮询')
            except Exception as e:
                QtWidgets.QMessageBox.critical(self, '错误', f'开始轮询失败: {e}')
        else:
            if hasattr(self, 'poll_worker'):
                self.poll_worker.stop()
                self.poll_worker.wait(2000)
                self.poll_worker = None
            self.polling = False
            self.poll_btn.setText('Start Polling')
            # 只有串口未关闭时可重新启用轮询按钮
            self.poll_btn.setEnabled(self.ser is not None)
            self.open_btn.setEnabled(True)
            self.statusBar().showMessage('停止轮询')

    def on_comm_signal(self, typ, content):
        # 移除(len=...)内容
        content = re.sub(r'\s*\(len=\d+\)', '', content)
        if typ == 'send':
            self.comm_log.append(f'<span style="color: #00ff99;">[Send] {content}</span>')
            try:
                print(content)
            except Exception:
                print('发送内容包含无法显示的字符，已跳过')
        else:
            self.comm_log.append(f'<span style="color: #ffff66;">[Receive] {content}</span>')
            try:
                print(content)
            except Exception:
                print('接收内容包含无法显示的字符，已跳过')

    def on_msg_signal(self, msg):
        self.comm_log.append(f'<span style="color: #ff9800;">{msg}</span>')

    def on_data_signal(self, addr, value):
        print(f"on_data_signal called: addr={addr}, value={value}")
        # 增加日志，查看接收的数据
        print(f"DEBUG: on_data_signal - 接收到地址 {addr} 的数据，解码值为 {value}")
        
        # 更新所有分组（所有tab）中的参数值
        for sheet, table in self.param_tables.items():
            DataProcessor.update_param_value(addr, value, self.param_tables, sheet)
        
        # 更新内存中的DataFrame值（不更新UI，UI由update_param_value处理）
        if 'All Parameters' in self.param_dfs:
            df = self.param_dfs['All Parameters']
            idxs = df.index[df['Address'] == str(addr)].tolist()
            print(f"DEBUG: on_data_signal - 找到地址 {addr} 在 All Parameters DataFrame中的行索引: {idxs}")
            
            if 'Current Value' in df.columns:
                for idx in idxs:
                    df.at[idx, 'Current Value'] = value

    def show_comm_log_menu(self, pos):
        menu = QtWidgets.QMenu(self)
        copy_action = menu.addAction('Copy')
        clear_action = menu.addAction('Clear Log')
        action = menu.exec_(self.comm_log.mapToGlobal(pos))
        if action == copy_action:
            self.comm_log.copy()
        elif action == clear_action:
            self.comm_log.clear()

    def import_excel(self):
        file_name, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Select Excel File",
            "",
            "Excel Files (*.xlsx *.xls)"
        )
        if file_name:
            try:
                # 复制文件到当前目录
                import shutil
                shutil.copy2(file_name, 'config_and_params.xlsx')
                # 重新加载数据
                self._build_all_tables()
                QtWidgets.QMessageBox.information(self, 'Success', 'Excel file imported successfully')
            except Exception as e:
                QtWidgets.QMessageBox.critical(self, 'Error', f'Failed to import Excel file: {e}')

    def closeEvent(self, event):
        """Handle window close event"""
        # 卸载插件
        self.plugin_manager.unload_plugins()
        if self.polling:
            self.toggle_polling()
        if self.ser is not None:
            self.ser.close()
        event.accept()

    def _check_excel_file(self):
        import os
        if not os.path.exists('config_and_params.xlsx'):
            reply = QtWidgets.QMessageBox.question(
                self,
                'Notice',
                'Excel config file not found. Import now?',
                QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                QtWidgets.QMessageBox.Yes
            )
            if reply == QtWidgets.QMessageBox.Yes:
                self.import_excel()
            else:
                QtWidgets.QMessageBox.warning(
                    self,
                    'Warning',
                    'Please import Excel file before using polling.'
                )
        else:
            self._build_all_tables()

    def resizeEvent(self, event):
        # 重新分组排列所有表格（不重新读Excel，只用内存数据重排）
        for name, df in self.param_dfs.items():
            table = self.param_tables[name]
            if name == 'All Parameters':
                # 全部参数页：每行一个参数，所有字段
                n = len(df)
                all_cols = list(df.columns)
                table.setRowCount(n)
                table.setColumnCount(len(all_cols))
                table.setHorizontalHeaderLabels(all_cols)
                for idx in range(n):
                    for col_idx, col in enumerate(all_cols):
                        item = QtWidgets.QTableWidgetItem(str(df.iloc[idx].get(col, '')))
                        item.setTextAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
                        table.setItem(idx, col_idx, item)
                table.resizeRowsToContents()
                table.resizeColumnsToContents()
                table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
            else:
                group_size = 3
                group_count = self._get_group_count()
                n = len(df)
                row_count = (n + group_count - 1) // group_count
                table.setRowCount(row_count)
                table.setColumnCount(group_count * group_size)
                headers = []
                for j in range(group_count):
                    headers.extend(['Data', 'Address', 'Value'])
                table.setHorizontalHeaderLabels(headers)
                for idx in range(n):
                    r = idx % row_count
                    g = idx // row_count
                    base_col = g * group_size
                    for c, col in enumerate(['name', 'addr', 'Current Value']):
                        item = QtWidgets.QTableWidgetItem(str(df.iloc[idx].get(col, '')))
                        item.setTextAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
                        table.setItem(r, base_col + c, item)
                table.resizeRowsToContents()
                table.resizeColumnsToContents()
                table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        return super().resizeEvent(event)

    def save_serial_config_to_excel(self, config):
        try:
            logging.info("准备写入串口设置到LocalSettings: %s", config)
            xls = pd.ExcelFile('config_and_params.xlsx')
            try:
                df = pd.read_excel(xls, sheet_name='LocalSettings')
                logging.info("读取LocalSettings页成功，当前内容: %s", df.to_dict())
            except Exception as e:
                logging.warning(f"读取LocalSettings页失败: {e}，新建DataFrame")
                df = pd.DataFrame(columns=['port', 'baudrate', 'bytesize', 'parity', 'stopbits', 'mode'])
                df.loc[0] = [None]*6
            for key in ['port', 'baudrate', 'bytesize', 'parity', 'stopbits', 'mode']:
                val = config.get(key, '')
                # 强制bytesize、stopbits、baudrate为字符串整数
                if key in ['bytesize', 'stopbits', 'baudrate']:
                    try:
                        val = str(int(float(val)))
                    except Exception:
                        val = str(val)
                if key in df.columns:
                    df.at[0, key] = val
                else:
                    df[key] = ''
                    df.at[0, key] = val
            logging.info("写入前LocalSettings内容: %s", df.to_dict())
            with pd.ExcelWriter('config_and_params.xlsx', engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name='LocalSettings', index=False)
            logging.info("成功写入串口设置到LocalSettings")
        except Exception as e:
            logging.error(f"保存串口设置到LocalSettings失败: {e}")

    def load_serial_config_from_excel(self):
        try:
            logging.info("准备读取LocalSettings页串口设置")
            df = pd.read_excel('config_and_params.xlsx', sheet_name='LocalSettings')
            logging.info("读取LocalSettings内容: %s", df.to_dict())
            if not df.empty:
                config = df.iloc[0]
                # 恢复port
                port = str(config.get('port', ''))
                if port and port not in [self.serial_config.port_cb.itemText(i) for i in range(self.serial_config.port_cb.count())]:
                    self.serial_config.port_cb.addItem(port)
                self.serial_config.port_cb.setCurrentText(port)
                # 恢复baudrate
                baudrate = str(config.get('baudrate', ''))
                try:
                    baudrate = str(int(float(baudrate)))
                except Exception:
                    pass
                baud_list = ['9600', '19200', '38400', '57600', '115200']
                if baudrate not in baud_list:
                    baudrate = '9600'
                self.serial_config.baud_cb.setCurrentText(baudrate)
                # 恢复bytesize
                bytesize = str(config.get('bytesize', ''))
                try:
                    bytesize = str(int(float(bytesize)))
                except Exception:
                    pass
                if bytesize not in ['8', '7']:
                    bytesize = '8'
                self.serial_config.data_cb.setCurrentText(bytesize)
                # 恢复parity
                parity = str(config.get('parity', ''))
                if parity and parity not in [self.serial_config.parity_cb.itemText(i) for i in range(self.serial_config.parity_cb.count())]:
                    self.serial_config.parity_cb.addItem(parity)
                self.serial_config.parity_cb.setCurrentText(parity)
                # 恢复stopbits
                stopbits = str(config.get('stopbits', ''))
                try:
                    stopbits = str(int(float(stopbits)))
                except Exception:
                    pass
                if stopbits not in ['1', '2']:
                    stopbits = '1'
                self.serial_config.stop_cb.setCurrentText(stopbits)
                # 恢复mode
                mode = str(config.get('mode', ''))
                if mode and mode not in [self.serial_config.mode_cb.itemText(i) for i in range(self.serial_config.mode_cb.count())]:
                    self.serial_config.mode_cb.addItem(mode)
                self.serial_config.mode_cb.setCurrentText(mode)
            logging.info("串口设置已恢复到UI")
        except Exception as e:
            logging.warning(f"读取串口设置失败: {e}")

    def open_config_file(self):
        import os
        import subprocess
        cfg_path = os.path.abspath('config_and_params.xlsx')
        if os.path.exists(cfg_path):
            try:
                os.startfile(cfg_path)
            except AttributeError:
                subprocess.Popen(['open', cfg_path])
        else:
            QtWidgets.QMessageBox.warning(self, '提示', '配置文件不存在')

    def restart_app(self):
        import sys
        import os
        python = sys.executable
        os.execl(python, python, *sys.argv)

    def show_chart_view(self):
        # 检查是否已有图表窗口打开
        if hasattr(self, 'chart_window') and self.chart_window.isVisible():
            self.chart_window.activateWindow()
            return
            
        # 获取当前选中的数据
        selected_addrs = []
        selected_names = {}  # 存储地址到名称的映射
        if self.current_sheet and self.current_sheet in self.param_tables:
            table = self.param_tables[self.current_sheet]
            # 获取选中的行
            selected_rows = set()
            for item in table.selectedItems():
                selected_rows.add(item.row())
                
            # 根据选中的行获取地址
            for row in selected_rows:
                # 查找此行中的地址项和名称项
                addr_col = -1
                name_col = -1
                for col in range(table.columnCount()):
                    header = table.horizontalHeaderItem(col).text() if table.horizontalHeaderItem(col) else ""
                    if header.lower() in ['address', 'addr']:
                        addr_col = col
                    elif header.lower() in ['name', 'data']:
                        name_col = col
                        
                if addr_col >= 0:
                    addr_item = table.item(row, addr_col)
                    # 同时获取对应的名称
                    name = ""
                    if name_col >= 0:
                        name_item = table.item(row, name_col)
                        if name_item:
                            name = name_item.text().strip()
                    
                    if addr_item and addr_item.text().strip() and addr_item.text().strip().isdigit():
                        addr = int(addr_item.text().strip())
                        selected_addrs.append(addr)
                        selected_names[addr] = name or f"地址{addr}"
        
        # 如果没有选中地址，显示提示
        if not selected_addrs:
            QtWidgets.QMessageBox.information(
                self, 
                '提示', 
                '请先在数据表中选择要显示的数据行'
            )
            return
            
        # 创建并显示图表窗口
        self.chart_window = ChartWindow(selected_addrs, selected_names, self)
        self.chart_window.show()
        
        # 如果已在轮询，连接信号以更新图表
        if self.polling and hasattr(self, 'poll_worker'):
            # 断开之前可能的连接，避免重复连接
            try:
                self.poll_worker.data_signal.disconnect(self.chart_window.update_data)
            except:
                pass  # 如果没有连接，忽略错误
                
            # 重新连接信号
            self.poll_worker.data_signal.connect(self.chart_window.update_data)
            print("已连接 data_signal 到 chart_window.update_data")

    def check_update(self):
        try:
            url = "https://tiantxl888.github.io/modbusanalyzer/update.json"
            resp = requests.get(url, timeout=5)
            if resp.status_code == 200:
                data = resp.json()
                download_url = data.get("url", "")
                changelog = data.get("changelog", "")
                msg = f"实验版本升级：\n\n{changelog}\n\n是否下载并自动升级？"
                if QtWidgets.QMessageBox.question(self, "升级", msg, QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No) == QtWidgets.QMessageBox.Yes:
                    self.download_and_replace(download_url)
            else:
                QtWidgets.QMessageBox.warning(self, "升级", "无法获取升级信息。")
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "升级", f"升级失败：{e}")

    def download_and_replace(self, download_url):
        try:
            # 获取当前exe路径
            if getattr(sys, 'frozen', False):
                exe_path = sys.executable
            else:
                QtWidgets.QMessageBox.warning(self, "升级", "请在打包后的exe中使用升级功能。")
                return
            # 下载到临时文件
            temp_dir = tempfile.gettempdir()
            temp_exe = os.path.join(temp_dir, "modbus_analyzer_new.exe")
            r = requests.get(download_url, stream=True)
            total = int(r.headers.get('content-length', 0))
            progress = QProgressDialog("正在下载新版本...", "取消", 0, total, self)
            progress.setWindowTitle("升级")
            progress.setWindowModality(QtCore.Qt.WindowModal)
            progress.show()
            with open(temp_exe, 'wb') as f:
                downloaded = 0
                for chunk in r.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
                        downloaded += len(chunk)
                        progress.setValue(downloaded)
                        QtWidgets.QApplication.processEvents()
                        if progress.wasCanceled():
                            QtWidgets.QMessageBox.information(self, "升级", "已取消下载。")
                            return
            progress.close()
            # 下载完成，准备替换
            QtWidgets.QMessageBox.information(self, "升级", "下载完成，程序将自动升级并重启。")
            # 生成升级批处理脚本
            updater_bat = os.path.join(temp_dir, "modbus_update.bat")
            with open(updater_bat, 'w', encoding='gbk') as bat:
                bat.write(f"""
@echo off
setlocal enableextensions
echo 正在升级...
:loop
TASKLIST | findstr /I /C:"{os.path.basename(exe_path)}" >nul 2>nul
if not errorlevel 1 (
    timeout /t 1 >nul
    goto loop
)
copy /Y "{temp_exe}" "{exe_path}"
start "" "{exe_path}"
del "%~f0"
                """)
            # 关闭主程序并运行升级脚本
            os.startfile(updater_bat)
            QtWidgets.QApplication.quit()
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "升级", f"升级失败：{e}")

# 图表窗口类
class ChartWindow(QtWidgets.QMainWindow):
    def __init__(self, addresses, names, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Data Chart')
        self.resize(1024, 768)
        
        # 保存名称映射，确保key为int
        self.names = {int(addr): names.get(addr, f"地址{addr}") for addr in addresses}
        
        # 创建主布局
        self.central_widget = QtWidgets.QWidget()
        self.main_layout = QtWidgets.QHBoxLayout(self.central_widget)
        
        # 创建左侧(图表)面板
        self.left_panel = QtWidgets.QWidget()
        self.chart_layout = QtWidgets.QVBoxLayout(self.left_panel)
        self.left_panel.setStyleSheet("background-color: black;")
        
        # 创建右侧(图例)面板
        self.right_panel = QtWidgets.QWidget()
        self.right_layout = QtWidgets.QVBoxLayout(self.right_panel)
        self.right_panel.setFixedWidth(200)
        self.right_panel.setStyleSheet("background-color: black; color: white;")
        
        # 创建图表
        self.plot_widget = pg.PlotWidget(axisItems={'bottom': TimeAxisItem(orientation='bottom')})
        self.plot_widget.setBackground('k')  # 黑色背景
        self.plot_widget.showGrid(x=True, y=True)  # 显示网格
        
        # 设置坐标轴颜色为白色
        for axis in ['left', 'bottom', 'right', 'top']:
            self.plot_widget.getAxis(axis).setPen(pg.mkPen(color='w'))
            if axis in ['left', 'bottom', 'right']:
                self.plot_widget.getAxis(axis).setTextPen('w')
        
        # 显示右侧Y轴
        self.plot_widget.showAxis('right')
        
        # 默认设置Y轴范围，确保负值正确显示
        self.plot_widget.setYRange(-5200, -3800)
        print("DEBUG: 初始Y轴范围设置为 [-5200, -3800]")
        
        # 不再固定X轴范围，自动扩展
        # self.plot_widget.setXRange(0, 60, padding=0)
        self.plot_widget.setMouseEnabled(x=False, y=False)
        
        # 添加图表到左侧面板
        self.chart_layout.addWidget(self.plot_widget)
        
        # 底部控制区域
        self.controls_widget = QtWidgets.QWidget()
        self.controls_layout = QtWidgets.QHBoxLayout(self.controls_widget)
        self.controls_widget.setStyleSheet("background-color: black; color: white;")
        
        # 添加地址显示
        self.address_label = QtWidgets.QLabel('监测地址: ' + ', '.join(map(str, addresses)))
        self.address_label.setStyleSheet("color: white;")
        self.controls_layout.addWidget(self.address_label)
        self.controls_layout.addStretch()
        
        # 暂停按钮
        self.pause_btn = QtWidgets.QPushButton('暂停')
        self.pause_btn.clicked.connect(self.toggle_pause)
        self.controls_layout.addWidget(self.pause_btn)
        
        # 点数控制
        self.controls_layout.addWidget(QtWidgets.QLabel('最大点数:'))
        self.max_points_combo = QtWidgets.QComboBox()
        self.max_points_combo.addItems(['100', '500', '1000', '5000', '10000'])
        self.max_points_combo.setCurrentText('1000')
        self.max_points_combo.currentTextChanged.connect(self.change_max_points)
        self.controls_layout.addWidget(self.max_points_combo)
        
        # 更新速率
        self.controls_layout.addWidget(QtWidgets.QLabel('更新速率:'))
        self.update_rate_combo = QtWidgets.QComboBox()
        self.update_rate_combo.addItems(['快速 (100ms)', '正常 (500ms)', '慢速 (1000ms)'])
        self.update_rate_combo.setCurrentText('正常 (500ms)')
        self.update_rate_combo.currentTextChanged.connect(self.change_update_rate)
        self.controls_layout.addWidget(self.update_rate_combo)
        
        # 清除和导出按钮
        self.clear_btn = QtWidgets.QPushButton('清除')
        self.clear_btn.clicked.connect(self.clear_data)
        self.controls_layout.addWidget(self.clear_btn)
        
        self.export_btn = QtWidgets.QPushButton('导出')
        self.export_btn.clicked.connect(self.export_data)
        self.controls_layout.addWidget(self.export_btn)
        
        # 添加控制区域到左侧面板
        self.chart_layout.addWidget(self.controls_widget)
        
        # 为每个地址创建图例项
        self.legend_items = {}
        colors = [(0, 0, 255), (255, 0, 0), (0, 255, 0), (255, 0, 255), 
                  (0, 255, 255), (255, 255, 0), (255, 255, 255)]  # 蓝,红,绿,洋红,青,黄,白
        
        # 添加右侧图例标题
        title_layout = QtWidgets.QHBoxLayout()
        title_label = QtWidgets.QLabel("数据图例")
        title_label.setStyleSheet("font-weight: bold; color: white; font-size: 14px;")
        title_layout.addWidget(title_label)
        
        # 添加标题旁边的值显示标签
        self.title_value_label = QtWidgets.QLabel("-")
        self.title_value_label.setStyleSheet("color: white; font-size: 14px;")
        self.title_value_label.setAlignment(QtCore.Qt.AlignRight)
        title_layout.addWidget(self.title_value_label, 1)  # 使用1的拉伸因子
        
        # 强制初始设置为一个默认值
        self.title_value_label.setText("最新: -9999")
        print("DEBUG: 初始化标题值设置为 '最新: -9999'")
        
        self.right_layout.addLayout(title_layout)
        self.right_layout.addSpacing(10)
        
        # 创建曲线和图例
        self.data = {}  # {addr: {'x': [], 'y': []}}
        self.curves = {}  # {addr: PlotCurveItem}
        self.paused = False
        self.paused_data = {}
        self.base_time = None  # 新增：记录第一个点的时间戳
        for i, addr in enumerate(addresses):
            addr_int = int(addr)
            color = colors[i % len(colors)]
            self.data[addr_int] = {'x': [], 'y': []}
            self.paused_data[addr_int] = {'x': [], 'y': []}
            name = self.names.get(addr_int, f"地址{addr_int}")
            legend_item = QtWidgets.QWidget()
            legend_layout = QtWidgets.QHBoxLayout(legend_item)
            legend_layout.setContentsMargins(2, 2, 2, 2)
            color_box = QtWidgets.QLabel()
            color_box.setFixedSize(16, 16)
            color_box.setStyleSheet(f"background-color: rgb({color[0]}, {color[1]}, {color[2]}); border: 1px solid white;")
            legend_layout.addWidget(color_box)
            name_label = QtWidgets.QLabel(name)
            name_label.setStyleSheet("color: white;")
            legend_layout.addWidget(name_label)
            value_label = QtWidgets.QLabel("-")
            value_label.setStyleSheet("color: white;")
            value_label.setAlignment(QtCore.Qt.AlignRight)
            legend_layout.addWidget(value_label)
            self.legend_items[addr_int] = {'label': name_label, 'value': value_label}
            self.right_layout.addWidget(legend_item)
            pen = pg.mkPen(color=color, width=2)
            # 只显示线条，不显示点
            curve = self.plot_widget.plot([], [], pen=pen, name=name)
            self.curves[addr_int] = curve
        
        # 添加弹性空间到右侧图例底部
        self.right_layout.addStretch(1)
        
        # 组装主布局
        self.main_layout.addWidget(self.left_panel, 1)  # 图表占更多空间
        self.main_layout.addWidget(self.right_panel)
        
        self.setCentralWidget(self.central_widget)
        
        # 设置数据
        self.start_time = time.time()
        self.max_points = 1000
        
        # 定时器
        self.update_timer = QtCore.QTimer()
        self.update_timer.timeout.connect(self.update_chart)
        self.update_timer.start(50)  # 刷新频率提升到50ms

    def toggle_pause(self):
        """暂停/继续图表更新"""
        self.paused = not self.paused
        if self.paused:
            self.pause_btn.setText('继续')
            # 启用鼠标交互（允许缩放和平移）
            self.plot_widget.setMouseEnabled(x=True, y=True)
            # 初始化暂停期间的数据存储
            for addr in self.data:
                self.paused_data[addr] = {'x': [], 'y': []}
        else:
            self.pause_btn.setText('暂停')
            # 禁用鼠标交互
            self.plot_widget.setMouseEnabled(x=False, y=False)
            # 恢复时，将暂停期间收集的数据添加到主数据中
            for addr in self.paused_data:
                if addr in self.data and self.paused_data[addr]['x']:
                    self.data[addr]['x'].extend(self.paused_data[addr]['x'])
                    self.data[addr]['y'].extend(self.paused_data[addr]['y'])
                    # 限制数据点数量
                    if len(self.data[addr]['y']) > self.max_points:
                        self.data[addr]['y'] = self.data[addr]['y'][-self.max_points:]
                        self.data[addr]['x'] = self.data[addr]['x'][-self.max_points:]
            # 清空暂停数据
            self.paused_data = {addr: {'x': [], 'y': []} for addr in self.data}
            # 立即更新图表以显示所有新数据
            self.update_chart()
            # 重置X轴范围到60秒
            self.plot_widget.setXRange(0, 60)

    def update_chart(self):
        if self.paused:
            return
        for addr in self.data:
            self._update_curve(addr)
        # 自动缩放X轴，显示所有历史数据，右侧留白30%
        all_x = []
        for d in self.data.values():
            all_x.extend(d['x'])
        if all_x:
            min_x = min(all_x)
            max_x = max(all_x)
            if min_x == max_x:
                min_x -= 1
                max_x += 1
            x_range = max_x - min_x
            # 右侧留白30%
            self.plot_widget.setXRange(min_x, max_x + x_range * 0.3, padding=0)
        # Y轴自适应，数据占90%高度，上下各留5%
        all_y = []
        for d in self.data.values():
            all_y.extend(d['y'])
        if all_y:
            min_y = min(all_y)
            max_y = max(all_y)
            if min_y == max_y:
                min_y -= 1
                max_y += 1
            y_range = max_y - min_y
            pad = y_range * 0.05
            self.plot_widget.setYRange(min_y - pad, max_y + pad, padding=0)

    def _update_curve(self, addr):
        """更新单个曲线的显示"""
        if addr in self.data and addr in self.curves:
            times = self.data[addr]['x']
            values = self.data[addr]['y']
            if times and values:
                try:
                    self.curves[addr].setData(times, values)
                    print(f"DEBUG: 更新曲线 addr={addr}, 时间点数={len(times)}, 值范围=[{min(values)}, {max(values)}]")
                    return True
                except Exception as e:
                    print(f"DEBUG: 更新曲线错误: {str(e)}")
        return False
    
    def clear_data(self):
        """清除所有数据"""
        for addr in self.data:
            self.data[addr]['x'] = []
            self.data[addr]['y'] = []
            # 清空显示的值
            if addr in self.legend_items:
                self.legend_items[addr]['value'].setText('-')
    
    def change_max_points(self, value_text):
        """改变最大显示点数"""
        try:
            self.max_points = int(value_text)
            # 裁剪现有数据
            for addr in self.data:
                if len(self.data[addr]['y']) > self.max_points:
                    self.data[addr]['y'] = self.data[addr]['y'][-self.max_points:]
                    self.data[addr]['x'] = self.data[addr]['x'][-self.max_points:]
        except ValueError:
            pass
            
    def change_update_rate(self, rate_text):
        """改变更新频率"""
        try:
            if "100ms" in rate_text:
                rate = 100
            elif "500ms" in rate_text:
                rate = 500
            else:
                rate = 1000
                
            self.update_timer.stop()
            self.update_timer.start(rate)
        except Exception:
            pass
            
    def export_data(self):
        """导出数据到CSV文件"""
        if not any(len(self.data[addr]['y']) for addr in self.data):
            QtWidgets.QMessageBox.information(self, '提示', '没有数据可导出')
            return
            
        file_name, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "保存数据",
            "",
            "CSV Files (*.csv)"
        )
        
        if file_name:
            try:
                import pandas as pd
                # 创建数据框
                df = pd.DataFrame()
                # 创建共同的时间轴
                all_times = []
                for addr in self.data:
                    all_times.extend(self.data[addr]['x'])
                all_times = sorted(list(set(all_times)))
                df['Time(s)'] = all_times
                
                # 为每个地址添加列
                for addr in self.data:
                    addr_values = []
                    addr_times = self.data[addr]['x']
                    addr_data = self.data[addr]['y']
                    
                    time_value_map = {t: v for t, v in zip(addr_times, addr_data)}
                    
                    for t in all_times:
                        if t in time_value_map:
                            addr_values.append(time_value_map[t])
                        else:
                            addr_values.append(float('nan'))
                    
                    df[f'Address_{addr}'] = addr_values
                
                # 保存到CSV
                df.to_csv(file_name, index=False)
                QtWidgets.QMessageBox.information(self, '成功', '数据已成功导出')
            except Exception as e:
                QtWidgets.QMessageBox.critical(self, '错误', f'导出数据失败: {e}')
                
    def closeEvent(self, event):
        """窗口关闭时停止定时器"""
        self.update_timer.stop()
        event.accept()

    def update_data(self, addr, value):
        print(f"update_data called: addr={addr}, value={value}")
        try:
            addr_int = int(addr)
            if addr_int in self.data:
                # 转换为数值
                try:
                    value_float = float(value)
                    value_int = int(value_float)
                except (ValueError, TypeError):
                    value_int = 0

                # 更新图例
                if addr_int in self.legend_items:
                    self.legend_items[addr_int]['value'].setText(str(value_int))
                self.title_value_label.setText(f"最新: {value_int}")
                QtWidgets.QApplication.processEvents()

                # 保存数据，x为时间偏移量
                t = time.time()
                if self.base_time is None:
                    self.base_time = t
                t_offset = t - self.base_time
                if self.paused:
                    self.paused_data[addr_int]['x'].append(t_offset)
                    self.paused_data[addr_int]['y'].append(value_int)
                else:
                    self.data[addr_int]['x'].append(t_offset)
                    self.data[addr_int]['y'].append(value_int)
                    # 限制点数
                    if len(self.data[addr_int]['y']) > self.max_points:
                        self.data[addr_int]['y'] = self.data[addr_int]['y'][-self.max_points:]
                        self.data[addr_int]['x'] = self.data[addr_int]['x'][-self.max_points:]
                    # 实时刷新（只setData，不plot_widget.update）
                    self._update_curve(addr_int)
        except Exception as e:
            print(f"update_data error: {e}")

if __name__ == '__main__':
    app = QtWidgets.QApplication([])
    win = MainWindow()
    win.show()
    app.exec_()

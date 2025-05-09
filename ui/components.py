from PyQt5 import QtWidgets, QtCore, QtGui
import serial.tools.list_ports

class SerialConfigWidget(QtWidgets.QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._init_ui()

    def _init_ui(self):
        layout = QtWidgets.QHBoxLayout(self)
        self.port_cb = QtWidgets.QComboBox()
        self.baud_cb = QtWidgets.QComboBox()
        self.data_cb = QtWidgets.QComboBox()
        self.parity_cb = QtWidgets.QComboBox()
        self.stop_cb = QtWidgets.QComboBox()
        self.mode_cb = QtWidgets.QComboBox()
        
        ports = [p.device for p in serial.tools.list_ports.comports()]
        self.port_cb.clear()
        self.port_cb.addItems(ports)
        self.port_cb.setEditable(False)
        baud_list = ['9600', '19200', '38400', '57600', '115200']
        self.baud_cb.clear()
        self.baud_cb.addItems(baud_list)
        self.baud_cb.setEditable(False)
        self.data_cb.clear()
        self.data_cb.addItems(['8', '7'])
        self.data_cb.setEditable(False)
        self.data_cb.setCurrentText('8')
        self.parity_cb.clear()
        self.parity_cb.addItems(['N', 'E', 'O'])
        self.parity_cb.setEditable(False)
        self.stop_cb.clear()
        self.stop_cb.addItems(['1', '2'])
        self.stop_cb.setEditable(False)
        self.stop_cb.setCurrentText('1')
        self.mode_cb.clear()
        self.mode_cb.addItems(['RTU', 'ASCII'])
        self.mode_cb.setEditable(False)

        layout.addWidget(QtWidgets.QLabel('Port'))
        layout.addWidget(self.port_cb)
        layout.addWidget(QtWidgets.QLabel('Baudrate'))
        layout.addWidget(self.baud_cb)
        layout.addWidget(QtWidgets.QLabel('Data Bits'))
        layout.addWidget(self.data_cb)
        layout.addWidget(QtWidgets.QLabel('Parity'))
        layout.addWidget(self.parity_cb)
        layout.addWidget(QtWidgets.QLabel('Stop Bits'))
        layout.addWidget(self.stop_cb)
        layout.addWidget(QtWidgets.QLabel('Mode'))
        layout.addWidget(self.mode_cb)
        layout.addStretch()

    def get_config(self):
        return {
            'port': self.port_cb.currentText(),
            'baudrate': int(self.baud_cb.currentText()),
            'bytesize': int(self.data_cb.currentText()),
            'parity': self.parity_cb.currentText(),
            'stopbits': int(self.stop_cb.currentText()),
            'mode': self.mode_cb.currentText()
        }

    def set_locked(self, locked: bool):
        self.port_cb.setEnabled(not locked)
        self.baud_cb.setEnabled(not locked)
        self.data_cb.setEnabled(not locked)
        self.parity_cb.setEnabled(not locked)
        self.stop_cb.setEnabled(not locked)
        self.mode_cb.setEnabled(not locked)

class ParamTableWidget(QtWidgets.QTableWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAlternatingRowColors(True)
        self.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)

class CommLogWidget(QtWidgets.QTextEdit):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setReadOnly(True)
        self.setStyleSheet('''
            background-color: #222b36;
            color: white;
            border-radius: 6px;
            font-family: Consolas, monospace;
            font-size: 13px;
        ''')
        self.setFixedHeight(6*22)  # 约6行高度（每行约22像素）
        self.setContextMenuPolicy(QtCore.Qt.CustomContextMenu) 
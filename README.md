# Modbus 分析软件

这是一个基于 PyQt5 的 Modbus 协议分析软件，支持 RTU 和 ASCII 模式。

## 功能特点

- 支持 Modbus RTU 和 ASCII 模式
- 自动识别串口设备
- 可配置的通信参数（波特率、数据位、校验位等）
- 参数表格化显示和实时更新
- 通信日志记录和查看
- Excel 导入导出功能
- 多语言支持

## 目录结构

```
.
├── ui/                  # 用户界面相关代码
│   ├── main_window.py  # 主窗口
│   └── components.py   # UI组件
├── core/               # 核心功能代码
│   ├── serial_manager.py   # 串口管理
│   ├── modbus_worker.py   # Modbus通信
│   ├── data_processor.py  # 数据处理
│   ├── protocol.py        # 协议实现
│   └── project_manager.py # 工程管理
└── utils/              # 工具函数
    ├── excel_manager.py   # Excel处理
    └── log_manager.py     # 日志管理
```

## 安装

1. 克隆仓库：
   ```bash
   git clone https://github.com/yourusername/modbusanalyzer.git
   cd modbusanalyzer
   ```

2. 安装依赖：
   ```bash
   pip install -r requirements.txt
   ```

## 使用

运行主程序：
```bash
python modbus_analyzer.py
```

## 配置

软件使用 Excel 文件 (`config_and_params.xlsx`) 存储配置信息和参数表：

- ConnectionConfig：通信配置
- LocalSettings：本地设置
- Language：界面语言
- Parameters：寄存器参数表

## 开发

1. 安装开发依赖：
   ```bash
   pip install -r requirements.txt
   ```

2. 运行测试：
   ```bash
   python -m pytest tests/
   ```

## 许可证

MIT License

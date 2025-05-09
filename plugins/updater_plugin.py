import os
import sys
import json
import requests
from PyQt5.QtWidgets import QAction, QMessageBox, QProgressDialog
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from core.plugin_base import PluginBase

class DownloadThread(QThread):
    """Thread for downloading updates"""
    progress = pyqtSignal(int)
    finished = pyqtSignal(bool, str)
    
    def __init__(self, url, save_path):
        super().__init__()
        self.url = url
        self.save_path = save_path
        
    def run(self):
        try:
            response = requests.get(self.url, stream=True)
            total_size = int(response.headers.get('content-length', 0))
            block_size = 1024
            downloaded = 0
            
            with open(self.save_path, 'wb') as f:
                for data in response.iter_content(block_size):
                    downloaded += len(data)
                    f.write(data)
                    if total_size > 0:
                        progress = int(downloaded * 100 / total_size)
                        self.progress.emit(progress)
            
            self.finished.emit(True, "下载完成")
        except Exception as e:
            self.finished.emit(False, str(e))

class UpdaterPlugin(PluginBase):
    """Plugin for handling application updates"""
    
    def __init__(self):
        self.main_window = None
        self.check_action = None
        self.download_thread = None
        self.progress_dialog = None
        
    def initialize(self):
        """Initialize the plugin"""
        pass
        
    def get_name(self):
        return "Updater"
        
    def get_version(self):
        return "1.0.0"
        
    def get_description(self):
        return "应用程序更新插件"
        
    def get_menu_items(self):
        """Get menu items for the main window"""
        self.check_action = QAction("检查更新", None)
        self.check_action.triggered.connect(self.check_for_updates)
        return [self.check_action]
        
    def cleanup(self):
        """Cleanup when plugin is unloaded"""
        if self.download_thread and self.download_thread.isRunning():
            self.download_thread.terminate()
            self.download_thread.wait()
            
    def set_main_window(self, main_window):
        """Set reference to main window"""
        self.main_window = main_window
        
    def check_for_updates(self):
        """Check for available updates"""
        try:
            # 获取当前版本
            current_version = self.main_window.get_version()
            
            # 从GitHub获取最新版本信息
            response = requests.get("https://api.github.com/repos/yourusername/modbusanalyzer/releases/latest")
            if response.status_code == 200:
                latest_release = response.json()
                latest_version = latest_release["tag_name"]
                
                if latest_version > current_version:
                    reply = QMessageBox.question(
                        self.main_window,
                        "发现新版本",
                        f"发现新版本 {latest_version}，是否更新？",
                        QMessageBox.Yes | QMessageBox.No
                    )
                    
                    if reply == QMessageBox.Yes:
                        self.download_update(latest_release["assets"][0]["browser_download_url"])
                else:
                    QMessageBox.information(
                        self.main_window,
                        "检查更新",
                        "当前已是最新版本"
                    )
            else:
                QMessageBox.warning(
                    self.main_window,
                    "检查更新",
                    "无法获取更新信息"
                )
        except Exception as e:
            QMessageBox.critical(
                self.main_window,
                "检查更新",
                f"检查更新时发生错误：{str(e)}"
            )
            
    def download_update(self, url):
        """Download the update"""
        try:
            # 创建下载目录
            download_dir = os.path.join(os.path.expanduser("~"), "Downloads", "ModbusAnalyzer")
            os.makedirs(download_dir, exist_ok=True)
            
            # 设置下载路径
            save_path = os.path.join(download_dir, "ModbusAnalyzer_new.exe")
            
            # 创建进度对话框
            self.progress_dialog = QProgressDialog("正在下载更新...", "取消", 0, 100, self.main_window)
            self.progress_dialog.setWindowModality(Qt.WindowModal)
            self.progress_dialog.setWindowTitle("下载更新")
            self.progress_dialog.setAutoClose(True)
            self.progress_dialog.setAutoReset(True)
            
            # 创建并启动下载线程
            self.download_thread = DownloadThread(url, save_path)
            self.download_thread.progress.connect(self.progress_dialog.setValue)
            self.download_thread.finished.connect(self.on_download_finished)
            self.download_thread.start()
            
        except Exception as e:
            QMessageBox.critical(
                self.main_window,
                "下载更新",
                f"下载更新时发生错误：{str(e)}"
            )
            
    def on_download_finished(self, success, message):
        """Handle download completion"""
        if success:
            reply = QMessageBox.question(
                self.main_window,
                "下载完成",
                "更新已下载完成，是否立即安装？",
                QMessageBox.Yes | QMessageBox.No
            )
            
            if reply == QMessageBox.Yes:
                self.install_update()
        else:
            QMessageBox.critical(
                self.main_window,
                "下载更新",
                f"下载更新失败：{message}"
            )
            
    def install_update(self):
        """Install the update"""
        try:
            # 获取当前程序路径
            current_exe = sys.executable
            download_dir = os.path.join(os.path.expanduser("~"), "Downloads", "ModbusAnalyzer")
            new_exe = os.path.join(download_dir, "ModbusAnalyzer_new.exe")
            
            # 创建批处理文件来执行更新
            batch_path = os.path.join(download_dir, "update.bat")
            with open(batch_path, "w") as f:
                f.write(f'@echo off\n')
                f.write(f'timeout /t 2 /nobreak\n')
                f.write(f'del "{current_exe}"\n')
                f.write(f'copy "{new_exe}" "{current_exe}"\n')
                f.write(f'del "{new_exe}"\n')
                f.write(f'start "" "{current_exe}"\n')
                f.write(f'del "%~f0"\n')
            
            # 执行批处理文件
            os.startfile(batch_path)
            
            # 关闭当前程序
            self.main_window.close()
            
        except Exception as e:
            QMessageBox.critical(
                self.main_window,
                "安装更新",
                f"安装更新时发生错误：{str(e)}"
            ) 
import os
import importlib.util
from PyQt5.QtWidgets import QMenu

class PluginManager:
    """Manager for handling plugins"""
    
    def __init__(self, main_window):
        self.main_window = main_window
        self.plugins = {}
        self.plugins_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), "plugins")
        
    def load_plugins(self):
        """Load all available plugins"""
        if not os.path.exists(self.plugins_dir):
            os.makedirs(self.plugins_dir)
            return
            
        # 遍历插件目录
        for filename in os.listdir(self.plugins_dir):
            if filename.endswith(".py") and not filename.startswith("__"):
                plugin_path = os.path.join(self.plugins_dir, filename)
                self.load_plugin(plugin_path)
                
    def load_plugin(self, plugin_path):
        """Load a single plugin"""
        try:
            # 动态导入插件模块
            spec = importlib.util.spec_from_file_location("plugin", plugin_path)
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)
            
            # 查找插件类
            for item in dir(module):
                obj = getattr(module, item)
                if isinstance(obj, type) and hasattr(obj, "__bases__"):
                    for base in obj.__bases__:
                        if base.__name__ == "PluginBase":
                            # 实例化插件
                            plugin = obj()
                            plugin.initialize()
                            plugin.set_main_window(self.main_window)
                            
                            # 添加菜单项
                            menu_items = plugin.get_menu_items()
                            if menu_items:
                                # 创建插件菜单
                                plugin_menu = QMenu(plugin.get_name(), self.main_window.menuBar())
                                self.main_window.menuBar().addMenu(plugin_menu)
                                
                                # 添加菜单项
                                for item in menu_items:
                                    plugin_menu.addAction(item)
                            
                            # 保存插件实例
                            self.plugins[plugin.get_name()] = plugin
                            break
                            
        except Exception as e:
            print(f"加载插件 {plugin_path} 失败: {str(e)}")
            
    def unload_plugins(self):
        """Unload all plugins"""
        for plugin in self.plugins.values():
            try:
                plugin.cleanup()
            except Exception as e:
                print(f"卸载插件 {plugin.get_name()} 失败: {str(e)}")
                
        self.plugins.clear() 
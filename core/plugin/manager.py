import os
import importlib.util
import ctypes
from typing import Dict, List, Any, Optional
from .base import PluginBase
from .config import PluginConfig

class PluginManager:
    def __init__(self, config_dir: str):
        self.config = PluginConfig(config_dir)
        self.plugins: Dict[str, PluginBase] = {}
        self.plugin_dirs = ["plugins"]
        
    def load_plugins(self):
        """Load all plugins"""
        for plugin_dir in self.plugin_dirs:
            if not os.path.exists(plugin_dir):
                continue
                
            for item in os.listdir(plugin_dir):
                plugin_path = os.path.join(plugin_dir, item)
                if os.path.isdir(plugin_path):
                    self._load_plugin_from_dir(plugin_path)
                    
    def _load_plugin_from_dir(self, plugin_dir: str):
        """Load plugin from directory"""
        # 首先尝试加载DLL插件
        dll_path = os.path.join(plugin_dir, f"{os.path.basename(plugin_dir)}.pyd")
        if os.path.exists(dll_path):
            try:
                # 加载DLL
                plugin_module = ctypes.CDLL(dll_path)
                plugin_class = getattr(plugin_module, f"{os.path.basename(plugin_dir)}Plugin")
                plugin = plugin_class()
                self._register_plugin(plugin)
                return
            except Exception as e:
                print(f"Error loading DLL plugin {plugin_dir}: {e}")
        
        # 如果DLL加载失败，尝试加载Python插件
        plugin_file = os.path.join(plugin_dir, "__init__.py")
        if os.path.exists(plugin_file):
            try:
                spec = importlib.util.spec_from_file_location(
                    os.path.basename(plugin_dir), plugin_file)
                if spec and spec.loader:
                    module = importlib.util.module_from_spec(spec)
                    spec.loader.exec_module(module)
                    if hasattr(module, "Plugin"):
                        plugin = module.Plugin()
                        self._register_plugin(plugin)
            except Exception as e:
                print(f"Error loading Python plugin {plugin_dir}: {e}")
                
    def _register_plugin(self, plugin: PluginBase):
        """Register a plugin"""
        plugin_name = plugin.get_name()
        if plugin_name in self.plugins:
            print(f"Plugin {plugin_name} already loaded")
            return
            
        if self.config.is_plugin_enabled(plugin_name):
            plugin.load_config(self.config.get_plugin_settings(plugin_name))
            plugin.initialize()
            plugin.enabled = True
            plugin.on_enable()
            
        self.plugins[plugin_name] = plugin
        
    def get_plugin(self, name: str) -> Optional[PluginBase]:
        """Get plugin by name"""
        return self.plugins.get(name)
        
    def get_all_plugins(self) -> List[PluginBase]:
        """Get all plugins"""
        return list(self.plugins.values())
        
    def enable_plugin(self, name: str):
        """Enable a plugin"""
        plugin = self.get_plugin(name)
        if plugin and not plugin.enabled:
            plugin.enabled = True
            plugin.on_enable()
            self.config.enable_plugin(name)
            
    def disable_plugin(self, name: str):
        """Disable a plugin"""
        plugin = self.get_plugin(name)
        if plugin and plugin.enabled:
            plugin.on_disable()
            plugin.enabled = False
            self.config.disable_plugin(name)
            
    def cleanup(self):
        """Cleanup all plugins"""
        for plugin in self.plugins.values():
            if plugin.enabled:
                plugin.cleanup() 
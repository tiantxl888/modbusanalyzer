import os
import json
from typing import Dict, List, Any

class PluginConfig:
    """Plugin configuration management"""
    
    def __init__(self, config_dir: str):
        self.config_dir = config_dir
        self.config_file = os.path.join(config_dir, "plugins.json")
        self.enabled_plugins: List[str] = []
        self.plugin_settings: Dict[str, Dict[str, Any]] = {}
        self.load_config()
        
    def load_config(self):
        """Load plugin configuration from file"""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.enabled_plugins = data.get('enabled_plugins', [])
                    self.plugin_settings = data.get('plugin_settings', {})
            except Exception as e:
                print(f"Error loading plugin config: {e}")
                self.enabled_plugins = []
                self.plugin_settings = {}
                
    def save_config(self):
        """Save plugin configuration to file"""
        try:
            os.makedirs(self.config_dir, exist_ok=True)
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump({
                    'enabled_plugins': self.enabled_plugins,
                    'plugin_settings': self.plugin_settings
                }, f, indent=4, ensure_ascii=False)
        except Exception as e:
            print(f"Error saving plugin config: {e}")
            
    def is_plugin_enabled(self, plugin_name: str) -> bool:
        """Check if a plugin is enabled"""
        return plugin_name in self.enabled_plugins
        
    def enable_plugin(self, plugin_name: str):
        """Enable a plugin"""
        if plugin_name not in self.enabled_plugins:
            self.enabled_plugins.append(plugin_name)
            self.save_config()
            
    def disable_plugin(self, plugin_name: str):
        """Disable a plugin"""
        if plugin_name in self.enabled_plugins:
            self.enabled_plugins.remove(plugin_name)
            self.save_config()
            
    def get_plugin_settings(self, plugin_name: str) -> Dict[str, Any]:
        """Get settings for a specific plugin"""
        return self.plugin_settings.get(plugin_name, {})
        
    def set_plugin_settings(self, plugin_name: str, settings: Dict[str, Any]):
        """Set settings for a specific plugin"""
        self.plugin_settings[plugin_name] = settings
        self.save_config() 
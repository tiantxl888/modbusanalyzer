from abc import ABC, abstractmethod
from typing import List, Dict, Any, Optional
from PyQt5.QtWidgets import QWidget

class PluginBase(ABC):
    """Base class for all plugins"""
    
    def __init__(self):
        self.enabled = False
        self.config = None
        self.main_window = None
        
    @abstractmethod
    def initialize(self):
        """Initialize the plugin"""
        pass
        
    @abstractmethod
    def get_name(self) -> str:
        """Get plugin name"""
        pass
        
    @abstractmethod
    def get_version(self) -> str:
        """Get plugin version"""
        pass
        
    @abstractmethod
    def get_description(self) -> str:
        """Get plugin description"""
        pass
        
    @abstractmethod
    def get_menu_items(self) -> List[Any]:
        """Get menu items to be added to the main window"""
        return []
        
    def get_dependencies(self) -> List[str]:
        """Get plugin dependencies"""
        return []
        
    def check_compatibility(self, app_version: str) -> bool:
        """Check if plugin is compatible with the application version"""
        return True
        
    def on_enable(self):
        """Called when plugin is enabled"""
        pass
        
    def on_disable(self):
        """Called when plugin is disabled"""
        pass
        
    def get_config_widget(self) -> Optional[QWidget]:
        """Get plugin configuration widget"""
        return None
        
    def load_config(self, config: Dict[str, Any]):
        """Load plugin configuration"""
        self.config = config
        
    def save_config(self) -> Dict[str, Any]:
        """Save plugin configuration"""
        return self.config or {}
        
    def set_main_window(self, main_window):
        """Set reference to main window"""
        self.main_window = main_window
        
    def cleanup(self):
        """Cleanup when plugin is unloaded"""
        if self.enabled:
            self.on_disable()
            self.enabled = False 
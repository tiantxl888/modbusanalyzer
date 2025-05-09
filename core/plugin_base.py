from abc import ABC, abstractmethod

class PluginBase(ABC):
    """Base class for all plugins"""
    
    @abstractmethod
    def initialize(self):
        """Initialize the plugin"""
        pass
        
    @abstractmethod
    def get_name(self):
        """Get plugin name"""
        pass
        
    @abstractmethod
    def get_version(self):
        """Get plugin version"""
        pass
        
    @abstractmethod
    def get_description(self):
        """Get plugin description"""
        pass
        
    @abstractmethod
    def get_menu_items(self):
        """Get menu items to be added to the main window"""
        return []
        
    @abstractmethod
    def cleanup(self):
        """Cleanup when plugin is unloaded"""
        pass 
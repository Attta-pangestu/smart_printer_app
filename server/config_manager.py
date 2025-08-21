import yaml
import os
from typing import Dict, Any, List, Optional
from dataclasses import dataclass
import logging

logger = logging.getLogger(__name__)

@dataclass
class PrinterConfig:
    """Konfigurasi printer default"""
    default_name: str
    default_id: str
    auto_discovery: bool
    refresh_interval: int
    allowed_printers: List[str]
    fallback_enabled: bool
    fallback_keywords: List[str]

@dataclass
class MonitoringConfig:
    """Konfigurasi monitoring printer"""
    status_check_interval: int
    retry_connection: bool
    max_retries: int
    retry_delay: int

@dataclass
class UIConfig:
    """Konfigurasi UI"""
    show_connection_status: bool
    show_visual_indicator: bool
    auto_refresh_status: bool
    refresh_interval_ui: int

@dataclass
class LoggingConfig:
    """Konfigurasi logging"""
    level: str
    log_discovery: bool
    log_status_changes: bool

class ConfigManager:
    """Manager untuk membaca dan mengelola konfigurasi printer"""
    
    def __init__(self, config_path: str = "printer_config.yaml"):
        self.config_path = config_path
        self._config_data: Optional[Dict[str, Any]] = None
        self._printer_config: Optional[PrinterConfig] = None
        self._monitoring_config: Optional[MonitoringConfig] = None
        self._ui_config: Optional[UIConfig] = None
        self._logging_config: Optional[LoggingConfig] = None
        
        self.load_config()
    
    def load_config(self) -> None:
        """Memuat konfigurasi dari file YAML"""
        try:
            if not os.path.exists(self.config_path):
                logger.warning(f"Config file {self.config_path} not found, using defaults")
                self._create_default_config()
                return
            
            with open(self.config_path, 'r', encoding='utf-8') as file:
                self._config_data = yaml.safe_load(file)
            
            self._parse_config()
            logger.info(f"Configuration loaded from {self.config_path}")
            
        except Exception as e:
            logger.error(f"Error loading config: {e}")
            self._create_default_config()
    
    def _create_default_config(self) -> None:
        """Membuat konfigurasi default jika file tidak ditemukan"""
        self._config_data = {
            'printer': {
                'default_name': 'EPSON L120 Series (Copy 1)',
                'default_id': 'epson_l120_series_copy_1',
                'auto_discovery': False,
                'refresh_interval': 300,
                'allowed_printers': ['EPSON L120 Series (Copy 1)', 'EPSON L120 Series'],
                'fallback': {
                    'enabled': True,
                    'search_keywords': ['EPSON L120', 'L120']
                }
            },
            'monitoring': {
                'status_check_interval': 30,
                'retry_connection': True,
                'max_retries': 3,
                'retry_delay': 5
            },
            'ui': {
                'show_connection_status': True,
                'show_visual_indicator': True,
                'auto_refresh_status': True,
                'refresh_interval_ui': 10
            },
            'logging': {
                'level': 'INFO',
                'log_discovery': True,
                'log_status_changes': True
            }
        }
        self._parse_config()
    
    def _parse_config(self) -> None:
        """Parse konfigurasi ke dalam dataclass"""
        if not self._config_data:
            return
        
        # Parse printer config
        printer_data = self._config_data.get('printer', {})
        fallback_data = printer_data.get('fallback', {})
        
        self._printer_config = PrinterConfig(
            default_name=printer_data.get('default_name', 'EPSON L120 Series (Copy 1)'),
            default_id=printer_data.get('default_id', 'epson_l120_series_copy_1'),
            auto_discovery=printer_data.get('auto_discovery', False),
            refresh_interval=printer_data.get('refresh_interval', 300),
            allowed_printers=printer_data.get('allowed_printers', []),
            fallback_enabled=fallback_data.get('enabled', True),
            fallback_keywords=fallback_data.get('search_keywords', [])
        )
        
        # Parse monitoring config
        monitoring_data = self._config_data.get('monitoring', {})
        self._monitoring_config = MonitoringConfig(
            status_check_interval=monitoring_data.get('status_check_interval', 30),
            retry_connection=monitoring_data.get('retry_connection', True),
            max_retries=monitoring_data.get('max_retries', 3),
            retry_delay=monitoring_data.get('retry_delay', 5)
        )
        
        # Parse UI config
        ui_data = self._config_data.get('ui', {})
        self._ui_config = UIConfig(
            show_connection_status=ui_data.get('show_connection_status', True),
            show_visual_indicator=ui_data.get('show_visual_indicator', True),
            auto_refresh_status=ui_data.get('auto_refresh_status', True),
            refresh_interval_ui=ui_data.get('refresh_interval_ui', 10)
        )
        
        # Parse logging config
        logging_data = self._config_data.get('logging', {})
        self._logging_config = LoggingConfig(
            level=logging_data.get('level', 'INFO'),
            log_discovery=logging_data.get('log_discovery', True),
            log_status_changes=logging_data.get('log_status_changes', True)
        )
    
    @property
    def printer(self) -> PrinterConfig:
        """Mendapatkan konfigurasi printer"""
        return self._printer_config
    
    @property
    def monitoring(self) -> MonitoringConfig:
        """Mendapatkan konfigurasi monitoring"""
        return self._monitoring_config
    
    @property
    def ui(self) -> UIConfig:
        """Mendapatkan konfigurasi UI"""
        return self._ui_config
    
    @property
    def logging_config(self) -> LoggingConfig:
        """Mendapatkan konfigurasi logging"""
        return self._logging_config
    
    def save_config(self) -> None:
        """Menyimpan konfigurasi ke file"""
        try:
            with open(self.config_path, 'w', encoding='utf-8') as file:
                yaml.dump(self._config_data, file, default_flow_style=False, allow_unicode=True)
            logger.info(f"Configuration saved to {self.config_path}")
        except Exception as e:
            logger.error(f"Error saving config: {e}")
    
    def update_printer_config(self, **kwargs) -> None:
        """Update konfigurasi printer"""
        if not self._config_data:
            self._create_default_config()
        
        printer_data = self._config_data.setdefault('printer', {})
        
        for key, value in kwargs.items():
            if key in ['default_name', 'default_id', 'auto_discovery', 'refresh_interval', 'allowed_printers']:
                printer_data[key] = value
            elif key.startswith('fallback_'):
                fallback_data = printer_data.setdefault('fallback', {})
                fallback_key = key.replace('fallback_', '')
                if fallback_key == 'keywords':
                    fallback_data['search_keywords'] = value
                else:
                    fallback_data[fallback_key] = value
        
        self._parse_config()
        self.save_config()
    
    def reload_config(self) -> None:
        """Memuat ulang konfigurasi dari file"""
        try:
            self.load_config()
            logger.info("Configuration reloaded successfully")
        except Exception as e:
            logger.error(f"Error reloading config: {e}")
            raise

# Global config manager instance
config_manager = ConfigManager()
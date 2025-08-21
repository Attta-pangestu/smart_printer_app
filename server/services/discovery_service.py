import socket
import threading
import time
import json
from typing import Dict, List, Optional, Callable
from datetime import datetime
import logging
from zeroconf import ServiceInfo, Zeroconf, ServiceBrowser, ServiceListener
import psutil

from models.printer import PrinterDiscovery

logger = logging.getLogger(__name__)


class PrinterDiscoveryListener(ServiceListener):
    """Listener untuk discovery printer di network"""
    
    def __init__(self, callback: Callable[[PrinterDiscovery], None]):
        self.callback = callback
        self.discovered_services = {}
    
    def add_service(self, zc: Zeroconf, type_: str, name: str) -> None:
        """Service baru ditemukan"""
        try:
            info = zc.get_service_info(type_, name)
            if info:
                discovery = PrinterDiscovery(
                    service_name=name,
                    ip_address=socket.inet_ntoa(info.addresses[0]),
                    port=info.port,
                    printer_name=info.properties.get(b'printer_name', b'Unknown').decode('utf-8'),
                    capabilities={
                        key.decode('utf-8'): value.decode('utf-8') 
                        for key, value in info.properties.items()
                    }
                )
                
                self.discovered_services[name] = discovery
                self.callback(discovery)
                logger.info(f"Discovered printer service: {name} at {discovery.ip_address}:{discovery.port}")
                
        except Exception as e:
            logger.error(f"Error processing discovered service {name}: {e}")
    
    def remove_service(self, zc: Zeroconf, type_: str, name: str) -> None:
        """Service hilang dari network"""
        if name in self.discovered_services:
            del self.discovered_services[name]
            logger.info(f"Printer service removed: {name}")
    
    def update_service(self, zc: Zeroconf, type_: str, name: str) -> None:
        """Service diupdate"""
        # Treat as new service
        self.add_service(zc, type_, name)


class DiscoveryService:
    """Service untuk network discovery dan broadcasting"""
    
    def __init__(self, service_name: str = "_printer-share._tcp.local.", port: int = 8080):
        self.service_name = service_name
        self.port = port
        
        # Zeroconf untuk mDNS
        self.zeroconf = None
        self.service_info = None
        self.browser = None
        
        # Discovery state
        self.discovered_printers: Dict[str, PrinterDiscovery] = {}
        self.discovery_callbacks: List[Callable[[PrinterDiscovery], None]] = []
        
        # Broadcasting state
        self.is_broadcasting = False
        self.broadcast_thread = None
        
        # Network info
        self.local_ip = self._get_local_ip()
        self.hostname = socket.gethostname()
    
    def start_broadcasting(self, printer_info: Dict[str, str] = None) -> bool:
        """Mulai broadcast service di network"""
        try:
            if self.is_broadcasting:
                return True
            
            self.zeroconf = Zeroconf()
            
            # Prepare service properties
            properties = {
                'version': '1.0.0',
                'hostname': self.hostname,
                'server_type': 'printer_share',
                'timestamp': str(int(time.time()))
            }
            
            if printer_info:
                properties.update(printer_info)
            
            # Convert properties to bytes
            properties_bytes = {
                key.encode('utf-8'): str(value).encode('utf-8')
                for key, value in properties.items()
            }
            
            # Create service info
            self.service_info = ServiceInfo(
                self.service_name,
                f"{self.hostname}.{self.service_name}",
                addresses=[socket.inet_aton(self.local_ip)],
                port=self.port,
                properties=properties_bytes,
                server=f"{self.hostname}.local."
            )
            
            # Register service
            self.zeroconf.register_service(self.service_info)
            self.is_broadcasting = True
            
            logger.info(f"Started broadcasting printer share service on {self.local_ip}:{self.port}")
            return True
            
        except Exception as e:
            logger.error(f"Error starting broadcast: {e}")
            return False
    
    def stop_broadcasting(self) -> bool:
        """Stop broadcast service"""
        try:
            if not self.is_broadcasting:
                return True
            
            if self.service_info and self.zeroconf:
                self.zeroconf.unregister_service(self.service_info)
            
            if self.zeroconf:
                self.zeroconf.close()
                self.zeroconf = None
            
            self.service_info = None
            self.is_broadcasting = False
            
            logger.info("Stopped broadcasting printer share service")
            return True
            
        except Exception as e:
            logger.error(f"Error stopping broadcast: {e}")
            return False
    
    def start_discovery(self) -> bool:
        """Mulai discovery printer di network"""
        try:
            if self.browser:
                return True
            
            if not self.zeroconf:
                self.zeroconf = Zeroconf()
            
            # Create listener
            listener = PrinterDiscoveryListener(self._on_printer_discovered)
            
            # Start browser
            self.browser = ServiceBrowser(self.zeroconf, self.service_name, listener)
            
            logger.info(f"Started printer discovery for {self.service_name}")
            return True
            
        except Exception as e:
            logger.error(f"Error starting discovery: {e}")
            return False
    
    def stop_discovery(self) -> bool:
        """Stop discovery"""
        try:
            if self.browser:
                self.browser.cancel()
                self.browser = None
            
            logger.info("Stopped printer discovery")
            return True
            
        except Exception as e:
            logger.error(f"Error stopping discovery: {e}")
            return False
    
    def get_discovered_printers(self) -> List[PrinterDiscovery]:
        """Dapatkan list printer yang ditemukan"""
        return list(self.discovered_printers.values())
    
    def add_discovery_callback(self, callback: Callable[[PrinterDiscovery], None]):
        """Tambah callback untuk discovery events"""
        self.discovery_callbacks.append(callback)
    
    def remove_discovery_callback(self, callback: Callable[[PrinterDiscovery], None]):
        """Hapus callback"""
        if callback in self.discovery_callbacks:
            self.discovery_callbacks.remove(callback)
    
    def scan_network_printers(self) -> List[Dict[str, str]]:
        """Scan network untuk printer (non-mDNS)"""
        printers = []
        
        try:
            # Get network interfaces
            for interface, addrs in psutil.net_if_addrs().items():
                for addr in addrs:
                    if addr.family == socket.AF_INET and not addr.address.startswith('127.'):
                        network = self._get_network_range(addr.address, addr.netmask)
                        printers.extend(self._scan_network_range(network))
            
        except Exception as e:
            logger.error(f"Error scanning network printers: {e}")
        
        return printers
    
    def test_connection(self, ip_address: str, port: int = 8080, timeout: int = 5) -> bool:
        """Test koneksi ke printer share server"""
        try:
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(timeout)
            result = sock.connect_ex((ip_address, port))
            sock.close()
            return result == 0
            
        except Exception as e:
            logger.error(f"Error testing connection to {ip_address}:{port}: {e}")
            return False
    
    def get_network_info(self) -> Dict[str, str]:
        """Dapatkan informasi network"""
        return {
            'local_ip': self.local_ip,
            'hostname': self.hostname,
            'service_name': self.service_name,
            'port': str(self.port),
            'is_broadcasting': str(self.is_broadcasting),
            'discovered_count': str(len(self.discovered_printers))
        }
    
    def _on_printer_discovered(self, discovery: PrinterDiscovery):
        """Callback saat printer ditemukan"""
        self.discovered_printers[discovery.service_name] = discovery
        
        # Notify callbacks
        for callback in self.discovery_callbacks:
            try:
                callback(discovery)
            except Exception as e:
                logger.error(f"Error in discovery callback: {e}")
    
    def _get_local_ip(self) -> str:
        """Dapatkan local IP address"""
        try:
            # Connect ke external address untuk mendapatkan local IP
            with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as s:
                s.connect(("8.8.8.8", 80))
                return s.getsockname()[0]
        except Exception:
            return "127.0.0.1"
    
    def _get_network_range(self, ip: str, netmask: str) -> str:
        """Dapatkan network range dari IP dan netmask"""
        try:
            import ipaddress
            network = ipaddress.IPv4Network(f"{ip}/{netmask}", strict=False)
            return str(network)
        except Exception:
            # Fallback ke /24 network
            ip_parts = ip.split('.')
            return f"{'.'.join(ip_parts[:3])}.0/24"
    
    def _scan_network_range(self, network_range: str) -> List[Dict[str, str]]:
        """Scan range network untuk printer services"""
        printers = []
        
        try:
            import ipaddress
            network = ipaddress.IPv4Network(network_range)
            
            # Scan subset IP addresses (untuk performa)
            for ip in list(network.hosts())[:50]:  # Limit ke 50 IP
                if self.test_connection(str(ip), self.port, timeout=1):
                    printers.append({
                        'ip_address': str(ip),
                        'port': str(self.port),
                        'discovered_method': 'network_scan',
                        'discovered_at': datetime.now().isoformat()
                    })
                    
        except Exception as e:
            logger.error(f"Error scanning network range {network_range}: {e}")
        
        return printers
    
    def __del__(self):
        """Cleanup saat object dihapus"""
        self.stop_discovery()
        self.stop_broadcasting()
import win32print
import win32api
import psutil
from typing import List, Optional, Dict, Any
from datetime import datetime
import logging
import re
from models.printer import Printer, PrinterStatus, PrinterCapability, PrinterInfo
from models.job import PrintJob, JobStatus
from config_manager import config_manager

logger = logging.getLogger(__name__)


class PrinterService:
    """Service untuk mengelola printer"""
    
    def __init__(self):
        self._printers_cache: Dict[str, Printer] = {}
        self._last_refresh = datetime.now()
        self._refresh_interval = config_manager.monitoring.status_check_interval
        self._config = config_manager
        
        # Initialize dengan printer default jika auto_discovery dimatikan
        if not self._config.printer.auto_discovery:
            self._initialize_default_printer()
    
    def get_all_printers(self, force_refresh: bool = False) -> List[Printer]:
        """Mendapatkan semua printer yang tersedia"""
        if self._config.printer.auto_discovery:
            if force_refresh or self._should_refresh():
                self._refresh_printers()
        elif not self._printers_cache or force_refresh:
            self._initialize_default_printer()
        
        return list(self._printers_cache.values())
    
    def get_printer(self, printer_id: str) -> Optional[Printer]:
        """Mendapatkan printer berdasarkan ID dengan fallback pencarian yang lebih baik"""
        if not printer_id:
            logger.debug("get_printer called with empty printer_id")
            return None
            
        logger.debug(f"Searching for printer with ID: '{printer_id}'")
        
        # Coba cari dengan ID yang diberikan
        if printer_id in self._printers_cache:
            logger.debug(f"Found printer in cache: '{self._printers_cache[printer_id].name}'")
            return self._printers_cache[printer_id]
        
        # Jika tidak ditemukan, coba refresh
        logger.debug("Printer not found in cache, checking if refresh needed...")
        if self._config.printer.auto_discovery:
            if self._should_refresh():
                logger.debug("Refreshing printers...")
                self._refresh_printers()
        else:
            logger.debug("Initializing default printer...")
            self._initialize_default_printer()
        
        # Coba lagi setelah refresh
        if printer_id in self._printers_cache:
            logger.debug(f"Found printer after refresh: '{self._printers_cache[printer_id].name}'")
            return self._printers_cache[printer_id]
        
        # Log semua printer yang tersedia untuk debugging
        logger.debug(f"Available printers: {list(self._printers_cache.keys())}")
        
        # Fallback: cari berdasarkan nama printer
        # 1. Coba dengan mengembalikan underscore ke spasi
        printer_name_from_id = printer_id.replace('_', ' ')
        logger.debug(f"Trying name conversion: '{printer_name_from_id}'")
        for printer in self._printers_cache.values():
            if printer.name.lower() == printer_name_from_id.lower():
                logger.debug(f"Found printer by name conversion: {printer.name} for ID: {printer_id}")
                return printer
        
        # 2. Coba dengan URL decode
        import urllib.parse
        try:
            decoded_name = urllib.parse.unquote(printer_id.replace('_', ' '))
            logger.debug(f"Trying URL decode: '{decoded_name}'")
            for printer in self._printers_cache.values():
                if printer.name.lower() == decoded_name.lower():
                    logger.debug(f"Found printer by URL decode: {printer.name} for ID: {printer_id}")
                    return printer
        except Exception as e:
            logger.debug(f"URL decode failed: {e}")
        
        # 3. Coba cari berdasarkan ID yang dinormalisasi dari nama printer
        logger.debug("Trying normalized ID matching...")
        for printer in self._printers_cache.values():
            normalized_name = self._normalize_printer_id(printer.name)
            logger.debug(f"Comparing normalized '{normalized_name}' with '{printer_id.lower()}'")
            if normalized_name == printer_id.lower():
                logger.debug(f"Found printer by normalized name: {printer.name} for ID: {printer_id}")
                return printer
        
        # 4. Fallback terakhir: cari berdasarkan substring
        logger.debug("Trying substring matching...")
        printer_id_lower = printer_id.lower()
        for printer in self._printers_cache.values():
            if printer_id_lower in printer.name.lower() or printer.name.lower() in printer_id_lower:
                logger.debug(f"Found printer by substring match: {printer.name} for ID: {printer_id}")
                return printer
        
        logger.warning(f"Printer not found for ID: {printer_id}. Available printers: {list(self._printers_cache.keys())}")
        return None
    
    def get_default_printer(self) -> Optional[Printer]:
        """Mendapatkan printer default"""
        try:
            default_printer_name = win32print.GetDefaultPrinter()
            for printer in self.get_all_printers():
                if printer.name == default_printer_name:
                    return printer
        except Exception as e:
            logger.error(f"Error getting default printer: {e}")
        
        return None
    
    def get_printer_status(self, printer_id: str) -> PrinterStatus:
        """Mendapatkan status printer dengan monitoring yang lebih akurat"""
        printer = self.get_printer(printer_id)
        if not printer:
            return PrinterStatus.OFFLINE
        
        try:
            # Buka printer handle
            printer_handle = win32print.OpenPrinter(printer.name)
            
            # Dapatkan status printer
            printer_info = win32print.GetPrinter(printer_handle, 2)
            status = printer_info['Status']
            jobs_count = printer_info.get('cJobs', 0)
            
            win32print.ClosePrinter(printer_handle)
            
            # Konversi status Windows ke enum kita dengan prioritas yang benar
            if status & 0x00000002:  # PRINTER_STATUS_ERROR
                return PrinterStatus.ERROR
            elif status & 0x00000400:  # PRINTER_STATUS_OFFLINE
                return PrinterStatus.OFFLINE
            elif status & 0x00000001:  # PRINTER_STATUS_PAUSED
                return PrinterStatus.PAUSED
            elif status & 0x00000200 or jobs_count > 0:  # PRINTER_STATUS_BUSY or has jobs
                return PrinterStatus.BUSY
            elif status == 0:  # PRINTER_STATUS_READY
                return PrinterStatus.ONLINE
            else:
                # Default untuk status yang tidak dikenal
                return PrinterStatus.ONLINE
                
        except Exception as e:
            logger.error(f"Error getting printer status for {printer_id}: {e}")
            return PrinterStatus.ERROR
    
    def get_detailed_printer_status(self, printer_id: str) -> Dict[str, Any]:
        """Mendapatkan status printer yang detail untuk monitoring real-time"""
        printer = self.get_printer(printer_id)
        if not printer:
            return {
                'status': PrinterStatus.OFFLINE,
                'jobs_count': 0,
                'error_message': 'Printer not found',
                'last_updated': datetime.now().isoformat()
            }
        
        try:
            printer_handle = win32print.OpenPrinter(printer.name)
            printer_info = win32print.GetPrinter(printer_handle, 2)
            
            status = printer_info['Status']
            jobs_count = printer_info.get('cJobs', 0)
            
            # Get job details if any
            jobs_info = []
            if jobs_count > 0:
                try:
                    jobs = win32print.EnumJobs(printer_handle, 0, -1, 1)
                    for job in jobs:
                        jobs_info.append({
                            'job_id': job.get('JobId', 0),
                            'document': job.get('pDocument', 'Unknown'),
                            'status': job.get('Status', 0),
                            'pages_printed': job.get('PagesPrinted', 0),
                            'total_pages': job.get('TotalPages', 0)
                        })
                except Exception as job_error:
                    logger.warning(f"Error getting job details: {job_error}")
            
            win32print.ClosePrinter(printer_handle)
            
            # Determine status
            printer_status = PrinterStatus.ONLINE
            error_message = None
            
            if status & 0x00000002:  # PRINTER_STATUS_ERROR
                printer_status = PrinterStatus.ERROR
                error_message = "Printer error detected"
            elif status & 0x00000400:  # PRINTER_STATUS_OFFLINE
                printer_status = PrinterStatus.OFFLINE
                error_message = "Printer is offline"
            elif status & 0x00000001:  # PRINTER_STATUS_PAUSED
                printer_status = PrinterStatus.PAUSED
                error_message = "Printer is paused"
            elif status & 0x00000200 or jobs_count > 0:  # PRINTER_STATUS_BUSY
                printer_status = PrinterStatus.BUSY
            
            return {
                'status': printer_status,
                'jobs_count': jobs_count,
                'jobs_info': jobs_info,
                'raw_status': status,
                'error_message': error_message,
                'last_updated': datetime.now().isoformat()
            }
                
        except Exception as e:
            logger.error(f"Error getting detailed printer status for {printer_id}: {e}")
            return {
                'status': PrinterStatus.ERROR,
                'jobs_count': 0,
                'error_message': str(e),
                'last_updated': datetime.now().isoformat()
            }
    
    def get_printer_jobs(self, printer_id: str) -> List[Dict[str, Any]]:
        """Mendapatkan job yang sedang berjalan di printer"""
        printer = self.get_printer(printer_id)
        if not printer:
            return []
        
        try:
            printer_handle = win32print.OpenPrinter(printer.name)
            jobs = win32print.EnumJobs(printer_handle, 0, -1, 1)
            win32print.ClosePrinter(printer_handle)
            
            return jobs
            
        except Exception as e:
            logger.error(f"Error getting printer jobs for {printer_id}: {e}")
            return []
    
    def print_test_page(self, printer_id: str) -> bool:
        """Print test page"""
        logger.info(f"Attempting to print test page to printer: {printer_id}")
        
        printer = self.get_printer(printer_id)
        if not printer:
            logger.error(f"Printer not found: {printer_id}")
            return False
        
        logger.info(f"Found printer: {printer.name}")
        
        try:
            # Buat test page content
            test_content = self._create_test_page_content(printer.name)
            logger.info(f"Created test content, length: {len(test_content)}")
            
            # Print menggunakan Windows API
            logger.info(f"Opening printer: {printer.name}")
            printer_handle = win32print.OpenPrinter(printer.name)
            
            job_info = ('Test Page', None, 'RAW')
            
            logger.info("Starting print job...")
            job_id = win32print.StartDocPrinter(printer_handle, 1, job_info)
            logger.info(f"Started doc printer, job ID: {job_id}")
            
            win32print.StartPagePrinter(printer_handle)
            logger.info("Started page printer")
            
            win32print.WritePrinter(printer_handle, test_content.encode('utf-8'))
            logger.info("Wrote content to printer")
            
            win32print.EndPagePrinter(printer_handle)
            logger.info("Ended page printer")
            
            win32print.EndDocPrinter(printer_handle)
            logger.info("Ended doc printer")
            
            win32print.ClosePrinter(printer_handle)
            logger.info("Closed printer")
            
            logger.info(f"Test page sent to printer {printer_id}, job ID: {job_id}")
            return True
            
        except Exception as e:
            logger.error(f"Error printing test page to {printer_id}: {e}")
            logger.error(f"Exception type: {type(e).__name__}")
            import traceback
            logger.error(f"Traceback: {traceback.format_exc()}")
            return False
    
    def _initialize_default_printer(self):
        """Inisialisasi printer default dari konfigurasi atau auto-detect USB printer"""
        try:
            self._printers_cache.clear()
            
            # Enumerate hanya printer lokal (USB)
            printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)
            
            found_printer = None
            usb_printers = []
            
            # Filter printer USB yang benar-benar terhubung dan ready
            for printer_info in printers:
                _, _, name, _ = printer_info
                # Cek apakah printer benar-benar tersedia dan bukan virtual/network printer
                if self._is_physical_usb_printer(name) and self._is_printer_ready(name):
                    usb_printers.append(printer_info)

            # Urutkan printer: prioritaskan yang tanpa kata 'Copy' 
            usb_printers.sort(key=lambda x: ('copy' in x[2].lower(), x[2]))

            # Cari printer default berdasarkan konfigurasi
            default_name = self._config.printer.default_name

            # Cari printer dengan nama exact match
            for printer_info in usb_printers:
                _, _, name, _ = printer_info
                if name == default_name:
                    found_printer = printer_info
                    break
            
            # Jika tidak ditemukan, coba fallback dengan keyword
            if not found_printer and self._config.printer.fallback_enabled:
                for keyword in self._config.printer.fallback_keywords:
                    for printer_info in usb_printers:
                        _, _, name, _ = printer_info
                        if keyword.lower() in name.lower():
                            found_printer = printer_info
                            logger.info(f"Found fallback printer: {name} using keyword: {keyword}")
                            break
                    if found_printer:
                        break
            
            # Jika masih tidak ditemukan, gunakan printer USB pertama yang tersedia
            if not found_printer and usb_printers:
                found_printer = usb_printers[0]
                _, _, name, _ = found_printer
                logger.info(f"Auto-selected first available USB printer: {name}")
            
            if found_printer:
                printer = self._create_printer_from_info(found_printer)
                self._printers_cache[printer.id] = printer
                
                # Juga tambahkan dengan ID konfigurasi untuk mapping
                config_id = self._config.printer.default_id
                if config_id != printer.id:
                    self._printers_cache[config_id] = printer
                
                # Set sebagai default printer sistem jika belum
                try:
                    current_default = win32print.GetDefaultPrinter()
                    if current_default != printer.name:
                        win32print.SetDefaultPrinter(printer.name)
                        logger.info(f"Set {printer.name} as system default printer")
                except Exception as e:
                    logger.warning(f"Could not set default printer: {e}")
                
                logger.info(f"Initialized default printer: {printer.name} with ID: {printer.id}")
            else:
                logger.warning(f"No USB printers found")
            
            self._last_refresh = datetime.now()
            
        except Exception as e:
            logger.error(f"Error initializing default printer: {e}")
    
    def _is_physical_usb_printer(self, printer_name: str) -> bool:
        """Cek apakah printer adalah printer fisik USB yang terhubung"""
        try:
            # Buka printer handle untuk cek status
            printer_handle = win32print.OpenPrinter(printer_name)
            printer_info = win32print.GetPrinter(printer_handle, 2)
            win32print.ClosePrinter(printer_handle)
            
            port_name = printer_info.get('pPortName', '').upper()
            
            # Filter out virtual/network printers
            virtual_ports = ['FILE:', 'XPS:', 'MICROSOFT', 'ONENOTE', 'FAX']
            network_indicators = ['\\\\', 'HTTP:', 'HTTPS:', 'IPP:', 'WSD']
            
            # Cek apakah port mengandung indikator virtual atau network
            for virtual_port in virtual_ports:
                if virtual_port in port_name:
                    return False
            
            for network_indicator in network_indicators:
                if network_indicator in port_name:
                    return False
            
            # Cek apakah printer name mengandung indikator virtual
            printer_upper = printer_name.upper()
            virtual_names = ['MICROSOFT', 'ONENOTE', 'XPS', 'FAX', 'PDF']
            for virtual_name in virtual_names:
                if virtual_name in printer_upper:
                    return False
            
            # Coba test koneksi printer untuk memastikan benar-benar tersedia
            try:
                test_handle = win32print.OpenPrinter(printer_name)
                win32print.ClosePrinter(test_handle)
                return True
            except:
                return False
                
        except Exception as e:
            logger.debug(f"Error checking printer {printer_name}: {e}")
            return False
    
    def _is_printer_ready(self, printer_name: str) -> bool:
        """Cek apakah printer dalam kondisi ready dan siap digunakan"""
        try:
            printer_handle = win32print.OpenPrinter(printer_name)
            printer_info = win32print.GetPrinter(printer_handle, 2)
            win32print.ClosePrinter(printer_handle)
            
            # Cek status printer
            status = printer_info.get('Status', 0)
            attributes = printer_info.get('Attributes', 0)
            
            # Status flags yang menunjukkan printer tidak ready
            not_ready_flags = [
                win32print.PRINTER_STATUS_OFFLINE,
                win32print.PRINTER_STATUS_ERROR,
                win32print.PRINTER_STATUS_PAPER_JAM,
                win32print.PRINTER_STATUS_PAPER_OUT,
                win32print.PRINTER_STATUS_OUTPUT_BIN_FULL,
                win32print.PRINTER_STATUS_NOT_AVAILABLE,
                win32print.PRINTER_STATUS_NO_TONER,
                win32print.PRINTER_STATUS_OUT_OF_MEMORY,
                win32print.PRINTER_STATUS_DOOR_OPEN,
                win32print.PRINTER_STATUS_USER_INTERVENTION
            ]
            
            # Cek apakah ada status error
            for flag in not_ready_flags:
                if status & flag:
                    logger.debug(f"Printer {printer_name} not ready: status flag {flag}")
                    return False
            
            # Cek apakah printer dalam mode work offline
            if attributes & win32print.PRINTER_ATTRIBUTE_WORK_OFFLINE:
                logger.debug(f"Printer {printer_name} is set to work offline")
                return False
            
            return True
            
        except Exception as e:
            logger.debug(f"Error checking printer readiness {printer_name}: {e}")
            return False
    
    def _normalize_printer_id(self, printer_name: str) -> str:
        """Normalisasi nama printer menjadi ID yang konsisten"""
        # Hapus karakter khusus dan ganti dengan underscore
        normalized = re.sub(r'[^a-zA-Z0-9\s]', '', printer_name)
        # Ganti spasi dengan underscore
        normalized = normalized.replace(' ', '_')
        # Hapus underscore berturut-turut
        normalized = re.sub(r'_+', '_', normalized)
        # Hapus underscore di awal dan akhir
        normalized = normalized.strip('_')
        # Lowercase
        return normalized.lower()
    
    def _should_refresh(self) -> bool:
        """Cek apakah perlu refresh cache"""
        return (datetime.now() - self._last_refresh).total_seconds() > self._refresh_interval
    
    def _refresh_printers(self):
        """Refresh cache printer"""
        try:
            self._printers_cache.clear()
            
            # Enumerate hanya printer lokal (USB) - tidak termasuk printer remote/network
            printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)
            
            for printer_info in printers:
                printer = self._create_printer_from_info(printer_info)
                self._printers_cache[printer.id] = printer
            
            self._last_refresh = datetime.now()
            logger.info(f"Refreshed {len(self._printers_cache)} printers")
            
        except Exception as e:
            logger.error(f"Error refreshing printers: {e}")
    
    def _create_printer_from_info(self, printer_info: tuple) -> Printer:
        """Buat object Printer dari info Windows"""
        flags, description, name, comment = printer_info
        
        # Generate ID unik untuk printer dengan normalisasi yang lebih baik
        printer_id = self._normalize_printer_id(name)
        
        # Log mapping untuk debugging
        logger.debug(f"Creating printer: '{name}' -> ID: '{printer_id}'")
        
        # Dapatkan detail printer
        capabilities = self._get_printer_capabilities(name)
        status = self._get_printer_status_from_name(name)
        
        # Cek apakah default printer
        is_default = False
        try:
            default_printer = win32print.GetDefaultPrinter()
            is_default = (name == default_printer)
        except:
            pass
        
        return Printer(
            id=printer_id,
            name=name,
            driver_name=description or "Unknown",
            port_name="",  # Akan diisi nanti jika diperlukan
            comment=comment,
            status=status,
            is_default=is_default,
            is_shared=(flags & win32print.PRINTER_ENUM_SHARED) != 0,
            capabilities=capabilities,
            jobs_count=0,  # Will be updated separately to avoid recursion
            last_seen=datetime.now()
        )
    
    def _get_printer_capabilities(self, printer_name: str) -> PrinterCapability:
        """Dapatkan capabilities printer"""
        try:
            printer_handle = win32print.OpenPrinter(printer_name)
            
            # Dapatkan device capabilities
            # Ini adalah implementasi sederhana, bisa diperluas
            capabilities = PrinterCapability(
                supports_color=True,  # Asumsi support color
                supports_duplex=True,  # Asumsi support duplex
                max_paper_size="A3",
                supported_paper_sizes=["A4", "A3", "Letter", "Legal"],
                supported_qualities=["draft", "normal", "high"],
                max_copies=999,
                supported_orientations=["portrait", "landscape"]
            )
            
            win32print.ClosePrinter(printer_handle)
            return capabilities
            
        except Exception as e:
            logger.error(f"Error getting capabilities for {printer_name}: {e}")
            return PrinterCapability()
    
    def _get_printer_status_from_name(self, printer_name: str) -> PrinterStatus:
        """Dapatkan status printer dari nama"""
        try:
            printer_handle = win32print.OpenPrinter(printer_name)
            printer_info = win32print.GetPrinter(printer_handle, 2)
            status = printer_info['Status']
            win32print.ClosePrinter(printer_handle)
            
            if status == 0:
                return PrinterStatus.ONLINE
            elif status & 0x00000001:
                return PrinterStatus.PAUSED
            elif status & 0x00000002:
                return PrinterStatus.ERROR
            elif status & 0x00000200:
                return PrinterStatus.BUSY
            else:
                return PrinterStatus.OFFLINE
                
        except Exception as e:
            logger.error(f"Error getting status for {printer_name}: {e}")
            return PrinterStatus.ERROR
    
    def _create_test_page_content(self, printer_name: str) -> str:
        """Buat konten test page"""
        from datetime import datetime
        now = datetime.now()
        date_str = now.strftime('%Y-%m-%d')
        time_str = now.strftime('%H:%M:%S')
        
        content = f"""PRINTER SHARING SYSTEM - TEST PAGE
==================================

Printer: {printer_name}
Date: {date_str}
Time: {time_str}

This is a test page to verify that your printer is working correctly.

Test patterns:
- Text printing: OK
- Character encoding: áéíóú àèìòù âêîôû ñç
- Numbers: 0123456789
- Symbols: !@#$%^&*()_+-=[]{{}}|;':",./<>?

If you can read this text clearly, your printer is working properly.

==================================
End of test page"""
        return content
    
    def get_printer_config(self) -> Dict[str, Any]:
        """Mendapatkan konfigurasi printer saat ini"""
        return {
            "auto_discovery": self._config.printer.auto_discovery,
            "default_name": self._config.printer.default_name,
            "default_id": self._config.printer.default_id,
            "refresh_interval": self._config.printer.refresh_interval,
            "allowed_printers": self._config.printer.allowed_printers,
            "fallback_enabled": self._config.printer.fallback_enabled,
            "fallback_keywords": self._config.printer.fallback_keywords,
            "status_check_interval": self._config.monitoring.status_check_interval,
            "retry_connection": self._config.monitoring.retry_connection,
            "max_retries": self._config.monitoring.max_retries,
            "retry_delay": self._config.monitoring.retry_delay,
            "show_connection_status": self._config.ui.show_connection_status,
            "show_visual_indicator": self._config.ui.show_visual_indicator,
            "auto_refresh_status": self._config.ui.auto_refresh_status,
            "refresh_interval_ui": self._config.ui.refresh_interval_ui
        }
    
    def update_printer_config(self, config_data: Dict[str, Any]) -> None:
        """Update konfigurasi printer"""
        try:
            # Update konfigurasi printer
            if "auto_discovery" in config_data:
                self._config.printer.auto_discovery = config_data["auto_discovery"]
            if "default_name" in config_data:
                self._config.printer.default_name = config_data["default_name"]
            if "default_id" in config_data:
                self._config.printer.default_id = config_data["default_id"]
            if "refresh_interval" in config_data:
                self._config.printer.refresh_interval = config_data["refresh_interval"]
            if "allowed_printers" in config_data:
                self._config.printer.allowed_printers = config_data["allowed_printers"]
            if "fallback_enabled" in config_data:
                self._config.printer.fallback_enabled = config_data["fallback_enabled"]
            if "fallback_keywords" in config_data:
                self._config.printer.fallback_keywords = config_data["fallback_keywords"]
            
            # Update konfigurasi monitoring
            if "status_check_interval" in config_data:
                self._config.monitoring.status_check_interval = config_data["status_check_interval"]
                self._refresh_interval = config_data["status_check_interval"]
            if "retry_connection" in config_data:
                self._config.monitoring.retry_connection = config_data["retry_connection"]
            if "max_retries" in config_data:
                self._config.monitoring.max_retries = config_data["max_retries"]
            if "retry_delay" in config_data:
                self._config.monitoring.retry_delay = config_data["retry_delay"]
            
            # Update konfigurasi UI
            if "show_connection_status" in config_data:
                self._config.ui.show_connection_status = config_data["show_connection_status"]
            if "show_visual_indicator" in config_data:
                self._config.ui.show_visual_indicator = config_data["show_visual_indicator"]
            if "auto_refresh_status" in config_data:
                self._config.ui.auto_refresh_status = config_data["auto_refresh_status"]
            if "refresh_interval_ui" in config_data:
                self._config.ui.refresh_interval_ui = config_data["refresh_interval_ui"]
            
            # Simpan konfigurasi ke file
            self._config.save_config()
            
            # Refresh printer cache jika diperlukan
            if "auto_discovery" in config_data or "default_name" in config_data or "default_id" in config_data:
                self._printers_cache.clear()
                if not self._config.printer.auto_discovery:
                    self._initialize_default_printer()
                    
            logger.info("Printer configuration updated successfully")
            
        except Exception as e:
            logger.error(f"Failed to update printer configuration: {e}")
            raise
    
    def reload_config(self) -> None:
        """Reload konfigurasi dari file"""
        try:
            self._config.reload_config()
            self._refresh_interval = self._config.monitoring.status_check_interval
            
            # Clear cache dan reinitialize
            self._printers_cache.clear()
            if not self._config.printer.auto_discovery:
                self._initialize_default_printer()
                
            logger.info("Printer configuration reloaded successfully")
            
        except Exception as e:
            logger.error(f"Failed to reload printer configuration: {e}")
            raise

#!/usr/bin/env python3
"""
Enhanced Print Service
Layanan cetak dengan multiple fallback methods
"""

import os
import sys
import subprocess
import win32print
import win32api
import win32con
import time
from pathlib import Path
from datetime import datetime

class EnhancedPrintService:
    def __init__(self):
        self.base_dir = Path("D:/Gawean Rebinmas/Driver_Epson_L120")
        self.tools_dir = self.base_dir / "print_tools"
        self.sumatra_exe = self.tools_dir / "SumatraPDF.exe"
        self.pdftoprinter_exe = self.tools_dir / "PDFtoPrinter.exe"
        
    def find_printer(self, pattern="EPSON L120"):
        """Cari printer berdasarkan pattern nama"""
        try:
            printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)
            for printer in printers:
                if pattern.lower() in printer[2].lower():
                    return printer[2]
            return None
        except Exception as e:
            print(f"Error finding printer: {e}")
            return None
    
    def get_queue_count(self, printer_name):
        """Dapatkan jumlah job dalam antrean printer"""
        try:
            handle = win32print.OpenPrinter(printer_name)
            jobs = win32print.EnumJobs(handle, 0, -1, 1)
            win32print.ClosePrinter(handle)
            return len(jobs)
        except Exception as e:
            print(f"Error getting queue count: {e}")
            return 0
    
    def print_with_sumatrapdf(self, file_path, printer_name, settings=None):
        """Cetak menggunakan SumatraPDF"""
        if not self.sumatra_exe.exists():
            return False, "SumatraPDF not found"
        
        try:
            cmd = [
                str(self.sumatra_exe),
                "-print-to", printer_name,
                "-silent",
                file_path
            ]
            
            # Add settings if provided
            if settings:
                if settings.get("copies", 1) > 1:
                    cmd.extend(["-print-settings", f"copies={settings['copies']}"])
            
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
            
            if result.returncode == 0:
                return True, "SumatraPDF print command executed"
            else:
                return False, f"SumatraPDF error: {result.stderr}"
                
        except Exception as e:
            return False, f"SumatraPDF exception: {e}"
    
    def print_with_pdftoprinter(self, file_path, printer_name, settings=None):
        """Cetak menggunakan PDFtoPrinter"""
        if not self.pdftoprinter_exe.exists():
            return False, "PDFtoPrinter not found"
        
        try:
            cmd = [str(self.pdftoprinter_exe), file_path, printer_name]
            
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
            
            if result.returncode == 0:
                return True, "PDFtoPrinter executed successfully"
            else:
                return False, f"PDFtoPrinter error: {result.stderr}"
                
        except Exception as e:
            return False, f"PDFtoPrinter exception: {e}"
    
    def print_with_powershell(self, file_path, printer_name, settings=None):
        """Cetak menggunakan PowerShell dengan printer spesifik"""
        try:
            # Create PowerShell script for printing
            ps_script = f"""
$printer = Get-Printer -Name "{printer_name}"
if ($printer) {{
    Start-Process -FilePath "{file_path}" -Verb Print -WindowStyle Hidden
    Write-Output "Print job sent to $($printer.Name)"
}} else {{
    Write-Error "Printer {printer_name} not found"
}}
"""
            
            cmd = ["powershell", "-Command", ps_script]
            
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
            
            if result.returncode == 0:
                return True, f"PowerShell print executed: {result.stdout}"
            else:
                return False, f"PowerShell error: {result.stderr}"
                
        except Exception as e:
            return False, f"PowerShell exception: {e}"
    
    def print_with_win32api(self, file_path, printer_name, settings=None):
        """Cetak menggunakan win32api langsung"""
        try:
            result = win32api.ShellExecute(
                0,
                "print",
                file_path,
                f'/d:"{printer_name}"',
                ".",
                win32con.SW_HIDE
            )
            
            if result > 32:  # Success
                return True, f"win32api print executed, result: {result}"
            else:
                return False, f"win32api print failed, result: {result}"
                
        except Exception as e:
            return False, f"win32api exception: {e}"
    
    def print_pdf_with_fallbacks(self, file_path, printer_name=None, settings=None):
        """Cetak PDF dengan multiple fallback methods"""
        if not printer_name:
            printer_name = self.find_printer()
            if not printer_name:
                return False, "No suitable printer found"
        
        print(f"Attempting to print {file_path} to {printer_name}")
        
        # Get initial queue count
        initial_queue = self.get_queue_count(printer_name)
        print(f"Initial queue count: {initial_queue}")
        
        # Try different print methods in order of preference
        methods = [
            ("SumatraPDF", self.print_with_sumatrapdf),
            ("PDFtoPrinter", self.print_with_pdftoprinter),
            ("PowerShell", self.print_with_powershell),
            ("Win32API", self.print_with_win32api)
        ]
        
        for method_name, method_func in methods:
            print(f"\nTrying {method_name}...")
            
            try:
                success, message = method_func(file_path, printer_name, settings)
                print(f"  {method_name}: {message}")
                
                if success:
                    # Wait a moment and check if job was added to queue
                    time.sleep(2)
                    new_queue = self.get_queue_count(printer_name)
                    print(f"  Queue count after {method_name}: {new_queue}")
                    
                    if new_queue > initial_queue:
                        print(f"✓ {method_name} successfully added job to queue")
                        return True, f"Successfully printed using {method_name}"
                    else:
                        print(f"⚠️  {method_name} executed but no job in queue")
                        # Continue to next method
                else:
                    print(f"❌ {method_name} failed: {message}")
                    
            except Exception as e:
                print(f"❌ {method_name} exception: {e}")
        
        return False, "All print methods failed"
    
    def monitor_print_job(self, printer_name, timeout=60):
        """Monitor print job progress"""
        print(f"\nMonitoring print job on {printer_name} for {timeout} seconds...")
        
        start_time = time.time()
        last_count = self.get_queue_count(printer_name)
        
        while time.time() - start_time < timeout:
            current_count = self.get_queue_count(printer_name)
            
            if current_count != last_count:
                timestamp = datetime.now().strftime("%H:%M:%S")
                print(f"[{timestamp}] Queue count changed: {last_count} -> {current_count}")
                last_count = current_count
                
                if current_count == 0:
                    print("✓ Print job completed (queue empty)")
                    return True
            
            time.sleep(1)
        
        print(f"⚠️  Monitoring timeout after {timeout} seconds")
        return False

# Test function
def test_enhanced_print_service():
    """Test enhanced print service"""
    service = EnhancedPrintService()
    
    test_file = "D:/Gawean Rebinmas/Driver_Epson_L120/temp/Test_print.pdf"
    
    if not os.path.exists(test_file):
        print(f"❌ Test file not found: {test_file}")
        return False
    
    # Test settings from user request
    settings = {
        "color_mode": "color",
        "copies": 1,
        "paper_size": "A4",
        "orientation": "portrait",
        "quality": "normal",
        "duplex": "none",
        "scale": 100
    }
    
    print("=== TESTING ENHANCED PRINT SERVICE ===")
    print(f"Test file: {test_file}")
    print(f"Settings: {settings}")
    
    success, message = service.print_pdf_with_fallbacks(test_file, settings=settings)
    
    if success:
        print(f"\n✓ Print successful: {message}")
        
        # Monitor the job
        printer_name = service.find_printer()
        if printer_name:
            service.monitor_print_job(printer_name, timeout=30)
        
        return True
    else:
        print(f"\n❌ Print failed: {message}")
        return False

if __name__ == "__main__":
    test_enhanced_print_service()

#!/usr/bin/env python3
"""
Implement Print Fixes
Mengimplementasikan perbaikan untuk masalah print service
"""

import os
import sys
import json
import subprocess
import urllib.request
import zipfile
from pathlib import Path

class PrintServiceFixer:
    def __init__(self):
        self.base_dir = Path("D:/Gawean Rebinmas/Driver_Epson_L120")
        self.tools_dir = self.base_dir / "print_tools"
        self.tools_dir.mkdir(exist_ok=True)
        
    def download_sumatrapdf(self):
        """Download SumatraPDF portable"""
        print("\n=== DOWNLOADING SUMATRAPDF ===")
        
        sumatra_url = "https://www.sumatrapdfreader.org/dl/rel/3.4.6/SumatraPDF-3.4.6-64.zip"
        sumatra_zip = self.tools_dir / "SumatraPDF.zip"
        sumatra_exe = self.tools_dir / "SumatraPDF.exe"
        
        if sumatra_exe.exists():
            print("✓ SumatraPDF already exists")
            return True
            
        try:
            print(f"Downloading SumatraPDF from {sumatra_url}")
            urllib.request.urlretrieve(sumatra_url, sumatra_zip)
            
            print("Extracting SumatraPDF...")
            with zipfile.ZipFile(sumatra_zip, 'r') as zip_ref:
                zip_ref.extractall(self.tools_dir)
            
            # Find the extracted exe
            for file in self.tools_dir.rglob("SumatraPDF.exe"):
                file.rename(sumatra_exe)
                break
            
            # Cleanup
            sumatra_zip.unlink()
            
            if sumatra_exe.exists():
                print("✓ SumatraPDF downloaded and extracted successfully")
                return True
            else:
                print("❌ Failed to extract SumatraPDF")
                return False
                
        except Exception as e:
            print(f"❌ Error downloading SumatraPDF: {e}")
            return False
    
    def download_pdftoprinter(self):
        """Download PDFtoPrinter utility"""
        print("\n=== DOWNLOADING PDFTOPRINTER ===")
        
        # Note: PDFtoPrinter requires manual download from official site
        # We'll create a placeholder and instructions
        
        pdftoprinter_exe = self.tools_dir / "PDFtoPrinter.exe"
        
        if pdftoprinter_exe.exists():
            print("✓ PDFtoPrinter already exists")
            return True
        
        instructions = """
PDFtoPrinter Download Instructions:
1. Go to: http://www.columbia.edu/~em36/pdftoprinter.html
2. Download PDFtoPrinter.exe
3. Place it in: {}
4. Re-run this script
""".format(self.tools_dir)
        
        print(instructions)
        
        # Create a batch file to help with download
        batch_content = f"""
@echo off
echo Opening PDFtoPrinter download page...
start http://www.columbia.edu/~em36/pdftoprinter.html
echo.
echo Please download PDFtoPrinter.exe and place it in:
echo {self.tools_dir}
echo.
echo Then re-run the Python script.
pause
"""
        
        batch_file = self.tools_dir / "download_pdftoprinter.bat"
        with open(batch_file, 'w') as f:
            f.write(batch_content)
        
        print(f"Created helper batch file: {batch_file}")
        return False
    
    def create_enhanced_print_service(self):
        """Buat enhanced print service dengan multiple fallback methods"""
        print("\n=== CREATING ENHANCED PRINT SERVICE ===")
        
        service_content = '''
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
'''
        
        service_file = self.base_dir / "enhanced_print_service.py"
        with open(service_file, 'w', encoding='utf-8') as f:
            f.write(service_content)
        
        print(f"✓ Enhanced print service created: {service_file}")
        return True
    
    def create_server_integration(self):
        """Buat integrasi dengan server yang ada"""
        print("\n=== CREATING SERVER INTEGRATION ===")
        
        integration_content = '''
#!/usr/bin/env python3
"""
Server Integration for Enhanced Print Service
Integrasi enhanced print service dengan server yang ada
"""

import sys
import os
from pathlib import Path

# Add current directory to path
sys.path.insert(0, str(Path(__file__).parent))

from enhanced_print_service import EnhancedPrintService

def patch_main_server():
    """Patch main server untuk menggunakan enhanced print service"""
    
    main_file = Path("server/main.py")
    
    if not main_file.exists():
        print(f"❌ Main server file not found: {main_file}")
        return False
    
    # Backup original file
    backup_file = main_file.with_suffix(".py.backup")
    if not backup_file.exists():
        import shutil
        shutil.copy2(main_file, backup_file)
        print(f"✓ Backup created: {backup_file}")
    
    # Read current content
    with open(main_file, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Add enhanced print service import
    if "from enhanced_print_service import EnhancedPrintService" not in content:
        import_line = "from enhanced_print_service import EnhancedPrintService\n"
        
        # Find a good place to insert the import
        lines = content.split('\n')
        insert_index = 0
        
        for i, line in enumerate(lines):
            if line.startswith('import ') or line.startswith('from '):
                insert_index = i + 1
        
        lines.insert(insert_index, import_line.strip())
        content = '\n'.join(lines)
    
    # Replace print service usage
    replacements = [
        # Replace DirectPrintService with EnhancedPrintService
        ("DirectPrintService()", "EnhancedPrintService()"),
        ("print_service = DirectPrintService()", "print_service = EnhancedPrintService()"),
        
        # Replace print method calls
        (".print_file_direct(", ".print_pdf_with_fallbacks("),
        (".print_pdf(", ".print_pdf_with_fallbacks("),
    ]
    
    for old, new in replacements:
        if old in content:
            content = content.replace(old, new)
            print(f"✓ Replaced: {old} -> {new}")
    
    # Write updated content
    with open(main_file, 'w', encoding='utf-8') as f:
        f.write(content)
    
    print(f"✓ Server integration completed: {main_file}")
    return True

def test_server_integration():
    """Test server integration"""
    print("\n=== TESTING SERVER INTEGRATION ===")
    
    try:
        # Test import
        from enhanced_print_service import EnhancedPrintService
        print("✓ Enhanced print service import successful")
        
        # Test service creation
        service = EnhancedPrintService()
        print("✓ Enhanced print service creation successful")
        
        # Test printer finding
        printer = service.find_printer()
        if printer:
            print(f"✓ Printer found: {printer}")
        else:
            print("⚠️  No printer found")
        
        return True
        
    except Exception as e:
        print(f"❌ Integration test failed: {e}")
        return False

if __name__ == "__main__":
    if patch_main_server():
        test_server_integration()
'''
        
        integration_file = self.base_dir / "integrate_enhanced_service.py"
        with open(integration_file, 'w', encoding='utf-8') as f:
            f.write(integration_content)
        
        print(f"✓ Server integration script created: {integration_file}")
        return True
    
    def run_complete_fix(self):
        """Jalankan perbaikan lengkap"""
        print("=== IMPLEMENTING COMPLETE PRINT FIXES ===")
        
        steps = [
            ("Downloading SumatraPDF", self.download_sumatrapdf),
            ("Setting up PDFtoPrinter", self.download_pdftoprinter),
            ("Creating Enhanced Print Service", self.create_enhanced_print_service),
            ("Creating Server Integration", self.create_server_integration)
        ]
        
        results = []
        
        for step_name, step_func in steps:
            print(f"\n--- {step_name} ---")
            try:
                result = step_func()
                results.append((step_name, result))
                
                if result:
                    print(f"✓ {step_name} completed successfully")
                else:
                    print(f"⚠️  {step_name} completed with warnings")
                    
            except Exception as e:
                print(f"❌ {step_name} failed: {e}")
                results.append((step_name, False))
        
        # Summary
        print("\n=== IMPLEMENTATION SUMMARY ===")
        for step_name, result in results:
            status = "✓" if result else "❌"
            print(f"  {status} {step_name}")
        
        print("\n=== NEXT STEPS ===")
        print("1. If PDFtoPrinter download is needed, run the batch file in print_tools/")
        print("2. Run 'python enhanced_print_service.py' to test the new service")
        print("3. Run 'python integrate_enhanced_service.py' to integrate with main server")
        print("4. Restart the main server to use the enhanced print service")
        
        return True

def main():
    fixer = PrintServiceFixer()
    fixer.run_complete_fix()
    return 0

if __name__ == "__main__":
    sys.exit(main())
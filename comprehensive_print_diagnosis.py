#!/usr/bin/env python3
"""
Comprehensive Print Diagnosis
Mendiagnosis masalah print service secara mendalam
"""

import os
import sys
import json
import time
import subprocess
import win32print
import win32api
from datetime import datetime
from pathlib import Path

class ComprehensivePrintDiagnosis:
    def __init__(self):
        self.test_file = "D:\\Gawean Rebinmas\\Driver_Epson_L120\\temp\\Test_print.pdf"
        self.results = {}
        
    def check_system_environment(self):
        """Periksa environment sistem"""
        print("\n=== SYSTEM ENVIRONMENT CHECK ===")
        
        checks = {
            "Python version": sys.version,
            "Operating System": os.name,
            "Current working directory": os.getcwd(),
            "Test file exists": os.path.exists(self.test_file),
            "Test file size": os.path.getsize(self.test_file) if os.path.exists(self.test_file) else "N/A"
        }
        
        for check, result in checks.items():
            print(f"  {check}: {result}")
            
        return checks
    
    def check_printer_drivers(self):
        """Periksa driver printer yang terinstall"""
        print("\n=== PRINTER DRIVERS CHECK ===")
        
        try:
            # Enumerate all printers
            printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)
            
            print(f"Found {len(printers)} local printers:")
            
            printer_info = []
            for printer in printers:
                name = printer[2]
                print(f"  - {name}")
                
                try:
                    # Get detailed printer info
                    handle = win32print.OpenPrinter(name)
                    info = win32print.GetPrinter(handle, 2)
                    
                    printer_details = {
                        "name": name,
                        "driver": info.get('pDriverName', 'Unknown'),
                        "port": info.get('pPortName', 'Unknown'),
                        "status": info.get('Status', 0),
                        "attributes": info.get('Attributes', 0)
                    }
                    
                    print(f"    Driver: {printer_details['driver']}")
                    print(f"    Port: {printer_details['port']}")
                    print(f"    Status: {printer_details['status']}")
                    
                    # Check if printer is ready
                    if printer_details['status'] == 0:
                        print(f"    ✓ Printer appears ready")
                    else:
                        print(f"    ⚠️  Printer status indicates issues")
                    
                    printer_info.append(printer_details)
                    win32print.ClosePrinter(handle)
                    
                except Exception as e:
                    print(f"    ❌ Error getting printer details: {e}")
                    
            return printer_info
            
        except Exception as e:
            print(f"❌ Error enumerating printers: {e}")
            return []
    
    def check_pdf_associations(self):
        """Periksa asosiasi file PDF"""
        print("\n=== PDF FILE ASSOCIATIONS CHECK ===")
        
        try:
            # Check default PDF viewer
            result = subprocess.run([
                "powershell",
                "-Command",
                "Get-ItemProperty HKLM:\\SOFTWARE\\Classes\\.pdf"
            ], capture_output=True, text=True, timeout=10)
            
            if result.returncode == 0:
                print("✓ PDF association found")
                print(f"  Output: {result.stdout.strip()}")
            else:
                print("❌ No PDF association found")
                
            # Check if Adobe Reader or similar is installed
            common_pdf_viewers = [
                "AcroRd32.exe",
                "Acrobat.exe",
                "msedge.exe",
                "chrome.exe",
                "firefox.exe"
            ]
            
            print("\nChecking for PDF viewers:")
            for viewer in common_pdf_viewers:
                try:
                    result = subprocess.run([
                        "where", viewer
                    ], capture_output=True, text=True, timeout=5)
                    
                    if result.returncode == 0:
                        print(f"  ✓ Found: {viewer} at {result.stdout.strip()}")
                    else:
                        print(f"  - Not found: {viewer}")
                except:
                    print(f"  - Error checking: {viewer}")
                    
        except Exception as e:
            print(f"❌ Error checking PDF associations: {e}")
    
    def test_direct_print_methods(self):
        """Test berbagai metode print langsung"""
        print("\n=== DIRECT PRINT METHODS TEST ===")
        
        methods = [
            self.test_shell_execute_print,
            self.test_powershell_print,
            self.test_win32api_print,
            self.test_gsprint_method
        ]
        
        results = {}
        
        for method in methods:
            try:
                method_name = method.__name__
                print(f"\n--- Testing {method_name} ---")
                
                success, details = method()
                results[method_name] = {
                    "success": success,
                    "details": details
                }
                
                if success:
                    print(f"✓ {method_name} executed successfully")
                else:
                    print(f"❌ {method_name} failed")
                    
            except Exception as e:
                print(f"❌ Error in {method.__name__}: {e}")
                results[method.__name__] = {
                    "success": False,
                    "details": str(e)
                }
        
        return results
    
    def test_shell_execute_print(self):
        """Test menggunakan shell execute"""
        try:
            import win32api
            
            result = win32api.ShellExecute(
                0,
                "print",
                self.test_file,
                None,
                ".",
                0
            )
            
            return True, f"ShellExecute returned: {result}"
            
        except Exception as e:
            return False, str(e)
    
    def test_powershell_print(self):
        """Test menggunakan PowerShell"""
        try:
            cmd = [
                "powershell",
                "-Command",
                f"Start-Process -FilePath '{self.test_file}' -Verb Print -WindowStyle Hidden -PassThru"
            ]
            
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
            
            return result.returncode == 0, f"Return code: {result.returncode}, Output: {result.stdout}, Error: {result.stderr}"
            
        except Exception as e:
            return False, str(e)
    
    def test_win32api_print(self):
        """Test menggunakan win32api langsung"""
        try:
            import win32api
            import win32con
            
            # Try to print using win32api
            result = win32api.ShellExecute(
                0,
                "print",
                self.test_file,
                '/d:"EPSON L120 Series"',
                ".",
                win32con.SW_HIDE
            )
            
            return True, f"win32api print returned: {result}"
            
        except Exception as e:
            return False, str(e)
    
    def test_gsprint_method(self):
        """Test menggunakan gsprint jika tersedia"""
        try:
            # Check if gsprint is available
            result = subprocess.run(["where", "gsprint"], capture_output=True, text=True, timeout=5)
            
            if result.returncode != 0:
                return False, "gsprint not found in PATH"
            
            # Try to use gsprint
            cmd = [
                "gsprint",
                "-printer", "EPSON L120 Series",
                self.test_file
            ]
            
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
            
            return result.returncode == 0, f"gsprint result: {result.returncode}, Output: {result.stdout}, Error: {result.stderr}"
            
        except Exception as e:
            return False, str(e)
    
    def check_printer_queue_realtime(self):
        """Monitor printer queue secara real-time"""
        print("\n=== REAL-TIME PRINTER QUEUE MONITORING ===")
        
        try:
            printer_name = "EPSON L120 Series"
            handle = win32print.OpenPrinter(printer_name)
            
            print(f"Monitoring printer queue for: {printer_name}")
            print("Checking queue every 2 seconds for 30 seconds...")
            
            for i in range(15):  # 30 seconds total
                try:
                    jobs = win32print.EnumJobs(handle, 0, -1, 1)
                    timestamp = datetime.now().strftime("%H:%M:%S")
                    
                    print(f"[{timestamp}] Queue count: {len(jobs)}")
                    
                    if jobs:
                        for job in jobs:
                            print(f"  Job ID: {job['JobId']}, Status: {job['Status']}, Document: {job['pDocument']}")
                    
                    time.sleep(2)
                    
                except Exception as e:
                    print(f"  Error reading queue: {e}")
            
            win32print.ClosePrinter(handle)
            
        except Exception as e:
            print(f"❌ Error monitoring printer queue: {e}")
    
    def generate_recommendations(self):
        """Generate recommendations based on diagnosis"""
        print("\n=== RECOMMENDATIONS ===")
        
        recommendations = [
            "1. DRIVER ISSUES:",
            "   - Reinstall EPSON L120 printer driver from official website",
            "   - Ensure printer is set as default printer",
            "   - Check printer cable connections",
            "",
            "2. PDF VIEWER ISSUES:",
            "   - Install Adobe Acrobat Reader DC",
            "   - Set Adobe Reader as default PDF viewer",
            "   - Test printing from Adobe Reader manually",
            "",
            "3. WINDOWS PRINT SPOOLER:",
            "   - Restart Windows Print Spooler service",
            "   - Clear print spooler cache",
            "   - Check Windows Event Viewer for print errors",
            "",
            "4. ALTERNATIVE SOLUTIONS:",
            "   - Install GSPrint for command-line PDF printing",
            "   - Use SumatraPDF for lightweight PDF printing",
            "   - Consider using PDFtoPrinter utility",
            "",
            "5. SERVER IMPLEMENTATION:",
            "   - Implement proper job validation",
            "   - Add real-time queue monitoring",
            "   - Use multiple fallback print methods",
            "   - Add comprehensive error logging"
        ]
        
        for rec in recommendations:
            print(rec)
    
    def run_comprehensive_diagnosis(self):
        """Jalankan diagnosis lengkap"""
        print("=== COMPREHENSIVE PRINT DIAGNOSIS ===")
        print(f"Timestamp: {datetime.now()}")
        print(f"Test file: {self.test_file}")
        
        # Run all diagnostic checks
        self.results['system_env'] = self.check_system_environment()
        self.results['printer_drivers'] = self.check_printer_drivers()
        self.check_pdf_associations()
        self.results['print_methods'] = self.test_direct_print_methods()
        
        # Ask user if they want real-time monitoring
        print("\n=== REAL-TIME MONITORING ===")
        print("Do you want to run real-time printer queue monitoring?")
        print("This will monitor the printer queue for 30 seconds.")
        
        try:
            choice = input("Run monitoring? (y/n): ").strip().lower()
            if choice == 'y':
                self.check_printer_queue_realtime()
        except KeyboardInterrupt:
            print("\nMonitoring cancelled by user")
        
        # Generate recommendations
        self.generate_recommendations()
        
        # Save results to file
        self.save_diagnosis_report()
        
        print("\n=== DIAGNOSIS COMPLETE ===")
        print("Report saved to: diagnosis_report.json")
    
    def save_diagnosis_report(self):
        """Simpan hasil diagnosis ke file"""
        try:
            report = {
                "timestamp": datetime.now().isoformat(),
                "test_file": self.test_file,
                "results": self.results
            }
            
            with open("diagnosis_report.json", "w") as f:
                json.dump(report, f, indent=2, default=str)
                
        except Exception as e:
            print(f"❌ Error saving report: {e}")

def main():
    diagnosis = ComprehensivePrintDiagnosis()
    diagnosis.run_comprehensive_diagnosis()
    return 0

if __name__ == "__main__":
    sys.exit(main())
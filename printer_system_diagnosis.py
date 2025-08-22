#!/usr/bin/env python3
"""
Printer System Diagnosis
Diagnosis mendalam sistem printer untuk mengidentifikasi masalah
"""

import os
import sys
import subprocess
import time
from pathlib import Path

def check_printer_service():
    """Check Windows Print Spooler service"""
    print("=== CHECKING WINDOWS PRINT SPOOLER SERVICE ===")
    
    try:
        # Check spooler service status
        result = subprocess.run(
            ['sc', 'query', 'spooler'], 
            capture_output=True, 
            text=True, 
            timeout=10
        )
        
        if result.returncode == 0:
            print("✅ Print Spooler service found")
            print(f"Service status:\n{result.stdout}")
            
            if "RUNNING" in result.stdout:
                print("✅ Print Spooler is RUNNING")
                return True
            else:
                print("❌ Print Spooler is NOT RUNNING")
                
                # Try to start the service
                print("Attempting to start Print Spooler...")
                start_result = subprocess.run(
                    ['sc', 'start', 'spooler'], 
                    capture_output=True, 
                    text=True, 
                    timeout=15
                )
                
                if start_result.returncode == 0:
                    print("✅ Print Spooler started successfully")
                    return True
                else:
                    print(f"❌ Failed to start Print Spooler: {start_result.stderr}")
                    return False
        else:
            print(f"❌ Failed to query Print Spooler: {result.stderr}")
            return False
            
    except Exception as e:
        print(f"❌ Error checking Print Spooler: {e}")
        return False

def check_printer_drivers():
    """Check installed printer drivers"""
    print("\n=== CHECKING PRINTER DRIVERS ===")
    
    try:
        # List installed printers
        result = subprocess.run(
            ['wmic', 'printer', 'get', 'name,drivername,portname,status'], 
            capture_output=True, 
            text=True, 
            timeout=15
        )
        
        if result.returncode == 0:
            print("Installed Printers:")
            print(result.stdout)
            
            # Check specifically for EPSON L120
            if "EPSON L120" in result.stdout or "L120" in result.stdout:
                print("✅ EPSON L120 printer found in system")
                return True
            else:
                print("❌ EPSON L120 printer NOT found in system")
                return False
        else:
            print(f"❌ Failed to list printers: {result.stderr}")
            return False
            
    except Exception as e:
        print(f"❌ Error checking printer drivers: {e}")
        return False

def check_printer_connection():
    """Check printer physical connection"""
    print("\n=== CHECKING PRINTER CONNECTION ===")
    
    try:
        # Check USB devices
        result = subprocess.run(
            ['wmic', 'path', 'Win32_USBHub', 'get', 'name,description'], 
            capture_output=True, 
            text=True, 
            timeout=15
        )
        
        if result.returncode == 0:
            print("USB Devices:")
            print(result.stdout)
            
            if "EPSON" in result.stdout.upper() or "L120" in result.stdout.upper():
                print("✅ EPSON device found in USB devices")
                return True
            else:
                print("❌ EPSON device NOT found in USB devices")
                
                # Check if any printer-related USB devices
                if "PRINTER" in result.stdout.upper() or "PRINT" in result.stdout.upper():
                    print("⚠️  Other printer devices found, but not EPSON L120")
                else:
                    print("❌ No printer devices found in USB")
                return False
        else:
            print(f"❌ Failed to check USB devices: {result.stderr}")
            return False
            
    except Exception as e:
        print(f"❌ Error checking USB connection: {e}")
        return False

def test_basic_print():
    """Test basic printing with a simple text file"""
    print("\n=== TESTING BASIC PRINT FUNCTIONALITY ===")
    
    try:
        # Create a simple test file
        test_file = "test_print.txt"
        with open(test_file, 'w') as f:
            f.write("PRINTER TEST\n")
            f.write("=============\n")
            f.write("This is a test print from the diagnosis script.\n")
            f.write(f"Date: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("If you can read this, basic printing works.\n")
        
        print(f"✅ Test file created: {test_file}")
        
        # Try to print using notepad
        print("Attempting to print test file using notepad...")
        result = subprocess.run(
            ['notepad', '/p', test_file], 
            capture_output=True, 
            text=True, 
            timeout=30
        )
        
        if result.returncode == 0:
            print("✅ Notepad print command executed")
            
            user_input = input("\nDid the test page print successfully? (y/n): ").strip().lower()
            
            if user_input == 'y':
                print("✅ BASIC PRINTING WORKS - The issue is with PDF printing specifically")
                os.remove(test_file)
                return True
            else:
                print("❌ BASIC PRINTING FAILED - Fundamental printer issue")
                os.remove(test_file)
                return False
        else:
            print(f"❌ Notepad print failed: {result.stderr}")
            os.remove(test_file)
            return False
            
    except Exception as e:
        print(f"❌ Error testing basic print: {e}")
        if os.path.exists(test_file):
            os.remove(test_file)
        return False

def check_pdf_associations():
    """Check PDF file associations"""
    print("\n=== CHECKING PDF FILE ASSOCIATIONS ===")
    
    try:
        # Check default PDF application
        result = subprocess.run(
            ['assoc', '.pdf'], 
            capture_output=True, 
            text=True, 
            timeout=10
        )
        
        if result.returncode == 0:
            print(f"PDF association: {result.stdout.strip()}")
            
            # Get the actual application
            pdf_type = result.stdout.strip().split('=')[1] if '=' in result.stdout else None
            
            if pdf_type:
                app_result = subprocess.run(
                    ['ftype', pdf_type], 
                    capture_output=True, 
                    text=True, 
                    timeout=10
                )
                
                if app_result.returncode == 0:
                    print(f"PDF application: {app_result.stdout.strip()}")
                    
                    if "acrobat" in app_result.stdout.lower() or "reader" in app_result.stdout.lower():
                        print("✅ Adobe Acrobat/Reader detected")
                        return True
                    elif "edge" in app_result.stdout.lower():
                        print("⚠️  Microsoft Edge detected (limited print support)")
                        return False
                    elif "chrome" in app_result.stdout.lower():
                        print("⚠️  Chrome detected (limited print support)")
                        return False
                    else:
                        print("⚠️  Unknown PDF viewer detected")
                        return False
                else:
                    print("❌ Could not determine PDF application")
                    return False
            else:
                print("❌ Could not parse PDF association")
                return False
        else:
            print("❌ No PDF association found")
            return False
            
    except Exception as e:
        print(f"❌ Error checking PDF associations: {e}")
        return False

def generate_diagnosis_report(results):
    """Generate comprehensive diagnosis report"""
    print("\n" + "="*60)
    print("=== COMPREHENSIVE DIAGNOSIS REPORT ===")
    print("="*60)
    
    spooler_ok, drivers_ok, connection_ok, basic_print_ok, pdf_ok = results
    
    print("\n📊 DIAGNOSIS RESULTS:")
    print(f"  Print Spooler Service: {'✅ OK' if spooler_ok else '❌ FAILED'}")
    print(f"  Printer Drivers: {'✅ OK' if drivers_ok else '❌ FAILED'}")
    print(f"  Physical Connection: {'✅ OK' if connection_ok else '❌ FAILED'}")
    print(f"  Basic Text Printing: {'✅ OK' if basic_print_ok else '❌ FAILED'}")
    print(f"  PDF File Association: {'✅ OK' if pdf_ok else '❌ FAILED'}")
    
    # Determine primary issue
    print("\n🔍 PRIMARY ISSUE ANALYSIS:")
    
    if not spooler_ok:
        print("  🚨 CRITICAL: Print Spooler service is not running")
        print("     This is the most likely cause of all printing failures")
        print("     SOLUTION: Restart the Print Spooler service")
        
    elif not drivers_ok:
        print("  🚨 CRITICAL: EPSON L120 printer driver not properly installed")
        print("     The printer is not recognized by Windows")
        print("     SOLUTION: Reinstall EPSON L120 drivers")
        
    elif not connection_ok:
        print("  🚨 CRITICAL: Printer not physically connected or powered off")
        print("     The printer hardware is not detected")
        print("     SOLUTION: Check USB cable and power connection")
        
    elif not basic_print_ok:
        print("  🚨 CRITICAL: Basic printing functionality is broken")
        print("     Even simple text printing fails")
        print("     SOLUTION: Check printer status, paper, ink levels")
        
    elif not pdf_ok:
        print("  ⚠️  WARNING: PDF printing issue detected")
        print("     Basic printing works but PDF handling is problematic")
        print("     SOLUTION: Install proper PDF viewer with print support")
        
    else:
        print("  ✅ All basic components appear functional")
        print("     The issue may be with the specific print service implementation")
        print("     SOLUTION: Review print service code and error handling")
    
    # Provide specific recommendations
    print("\n🔧 RECOMMENDED ACTIONS:")
    
    if not spooler_ok:
        print("  1. Run as Administrator: 'net start spooler'")
        print("  2. Check Windows Services and ensure Print Spooler is set to Automatic")
        print("  3. Restart the computer if service won't start")
        
    if not drivers_ok:
        print("  1. Download latest EPSON L120 drivers from EPSON website")
        print("  2. Uninstall existing printer from Control Panel")
        print("  3. Reinstall with new drivers")
        print("  4. Set EPSON L120 as default printer")
        
    if not connection_ok:
        print("  1. Check USB cable connection (try different USB port)")
        print("  2. Ensure printer is powered on")
        print("  3. Check printer display for error messages")
        print("  4. Try different USB cable if available")
        
    if not basic_print_ok:
        print("  1. Check printer paper tray (load paper if empty)")
        print("  2. Check ink levels (replace cartridges if low)")
        print("  3. Run printer self-test (usually button combination on printer)")
        print("  4. Clear any paper jams")
        
    if not pdf_ok:
        print("  1. Install Adobe Acrobat Reader DC")
        print("  2. Set Adobe Reader as default PDF application")
        print("  3. Test PDF printing manually from Adobe Reader")
        print("  4. Consider using SumatraPDF as alternative")
    
    # Overall assessment
    issues_count = sum(1 for result in results if not result)
    
    print(f"\n📈 OVERALL ASSESSMENT:")
    print(f"  Issues found: {issues_count}/5")
    
    if issues_count == 0:
        print("  🎉 All systems appear functional - investigate print service code")
    elif issues_count <= 2:
        print("  ⚠️  Minor issues detected - should be easily fixable")
    else:
        print("  🚨 Major issues detected - significant troubleshooting required")
    
    return issues_count

def main():
    """Main diagnosis function"""
    print("EPSON L120 PRINTER SYSTEM DIAGNOSIS")
    print("===================================\n")
    
    print("This script will perform a comprehensive diagnosis of your printer system.")
    print("Please ensure the EPSON L120 printer is connected and powered on.\n")
    
    input("Press Enter to begin diagnosis...")
    
    # Run all diagnostic tests
    results = [
        check_printer_service(),
        check_printer_drivers(),
        check_printer_connection(),
        test_basic_print(),
        check_pdf_associations()
    ]
    
    # Generate comprehensive report
    issues_count = generate_diagnosis_report(results)
    
    return 0 if issues_count == 0 else 1

if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\n\n⚠️  Diagnosis interrupted by user")
        sys.exit(1)
    except Exception as e:
        print(f"\n\n❌ Unexpected error during diagnosis: {e}")
        sys.exit(1)
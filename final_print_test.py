#!/usr/bin/env python3
"""
Final Print Test with User Validation
Test akhir dengan validasi manual dari pengguna
"""

import os
import sys
import time
from pathlib import Path
from enhanced_print_service import EnhancedPrintService

def test_print_with_user_validation():
    """Test print dengan validasi manual dari pengguna"""
    
    print("=== FINAL PRINT TEST WITH USER VALIDATION ===")
    print("Testing file: D:/Gawean Rebinmas/Driver_Epson_L120/temp/Test_print.pdf")
    
    # Settings sesuai permintaan pengguna
    settings = {
        "id": "72f4f36a-fed4-46f1-8da9-0ffb5e0922da",
        "printer_id": "epson_l120_series",
        "file_name": "Test_print.pdf",
        "file_path": "temp\\Test_print.pdf",
        "file_size": 607,
        "file_type": "pdf",
        "title": "Test_print.pdf",
        "user": "anonymous",
        "client_ip": None,
        "settings": {
            "color_mode": "color",
            "copies": 1,
            "paper_size": "A4",
            "orientation": "portrait",
            "quality": "normal",
            "duplex": "none",
            "scale": 100,
            "margins": {
                "top": 0.5,
                "bottom": 0.5,
                "left": 0.5,
                "right": 0.5
            },
            "collate": True,
            "reverse_order": False,
            "page_range": "",
            "pages_per_sheet": 1,
            "custom_paper": None,
            "header_footer": {
                "enabled": False,
                "header_left": "",
                "header_center": "",
                "header_right": "",
                "footer_left": "",
                "footer_center": "",
                "footer_right": ""
            },
            "page_breaks": {
                "avoid_page_breaks": False,
                "insert_manual_breaks": False,
                "break_positions": ""
            }
        },
        "status": "pending",
        "total_pages": 100,
        "pages_printed": 0,
        "progress_percentage": 0.0,
        "error_message": None,
        "retry_count": 0,
        "max_retries": 3,
        "metadata": {}
    }
    
    print("\nPrint Settings:")
    print(f"  Color Mode: {settings['settings']['color_mode']}")
    print(f"  Copies: {settings['settings']['copies']}")
    print(f"  Paper Size: {settings['settings']['paper_size']}")
    print(f"  Orientation: {settings['settings']['orientation']}")
    print(f"  Quality: {settings['settings']['quality']}")
    print(f"  Duplex: {settings['settings']['duplex']}")
    print(f"  Scale: {settings['settings']['scale']}%")
    
    # Test dengan Enhanced Print Service
    service = EnhancedPrintService()
    test_file = "D:/Gawean Rebinmas/Driver_Epson_L120/temp/Test_print.pdf"
    
    if not os.path.exists(test_file):
        print(f"❌ Test file not found: {test_file}")
        return False
    
    print("\n=== METHOD 1: Enhanced Print Service (PowerShell) ===")
    
    try:
        # Test PowerShell method specifically
        printer_name = service.find_printer()
        if not printer_name:
            print("❌ No printer found")
            return False
        
        print(f"Using printer: {printer_name}")
        
        # Get initial queue count
        initial_queue = service.get_queue_count(printer_name)
        print(f"Initial queue count: {initial_queue}")
        
        # Try PowerShell print
        success, message = service.print_with_powershell(test_file, printer_name, settings['settings'])
        print(f"PowerShell result: {message}")
        
        if success:
            print("\n⏳ Waiting 3 seconds for print job to process...")
            time.sleep(3)
            
            new_queue = service.get_queue_count(printer_name)
            print(f"Queue count after print: {new_queue}")
            
            print("\n📋 Please check your printer now:")
            print("1. Is the printer LED blinking or showing activity?")
            print("2. Did any paper come out of the printer?")
            print("3. Can you hear the printer making any sounds?")
            
            user_input = input("\nDid the printer show any activity or print anything? (y/n): ").strip().lower()
            
            if user_input == 'y':
                print("✅ METHOD 1 (PowerShell) - SUCCESS: Printer showed activity")
                method1_success = True
            else:
                print("❌ METHOD 1 (PowerShell) - FAILED: No printer activity")
                method1_success = False
        else:
            print("❌ METHOD 1 (PowerShell) - FAILED: Command execution failed")
            method1_success = False
            
    except Exception as e:
        print(f"❌ METHOD 1 (PowerShell) - ERROR: {e}")
        method1_success = False
    
    print("\n" + "="*60)
    
    print("\n=== METHOD 2: Enhanced Print Service (Win32API) ===")
    
    try:
        # Test Win32API method specifically
        success, message = service.print_with_win32api(test_file, printer_name, settings['settings'])
        print(f"Win32API result: {message}")
        
        if success:
            print("\n⏳ Waiting 3 seconds for print job to process...")
            time.sleep(3)
            
            new_queue = service.get_queue_count(printer_name)
            print(f"Queue count after print: {new_queue}")
            
            print("\n📋 Please check your printer again:")
            print("1. Is the printer LED blinking or showing activity?")
            print("2. Did any paper come out of the printer?")
            print("3. Can you hear the printer making any sounds?")
            
            user_input = input("\nDid the printer show any activity or print anything? (y/n): ").strip().lower()
            
            if user_input == 'y':
                print("✅ METHOD 2 (Win32API) - SUCCESS: Printer showed activity")
                method2_success = True
            else:
                print("❌ METHOD 2 (Win32API) - FAILED: No printer activity")
                method2_success = False
        else:
            print("❌ METHOD 2 (Win32API) - FAILED: Command execution failed")
            method2_success = False
            
    except Exception as e:
        print(f"❌ METHOD 2 (Win32API) - ERROR: {e}")
        method2_success = False
    
    print("\n" + "="*60)
    
    print("\n=== METHOD 3: Direct File Association ===")
    
    try:
        import subprocess
        
        # Try opening the PDF with default application
        print("Opening PDF with default application...")
        result = subprocess.run(["start", "", test_file], shell=True, capture_output=True, text=True, timeout=10)
        
        if result.returncode == 0:
            print("✅ PDF opened successfully")
            print("\n📋 Manual Print Test:")
            print("1. The PDF should have opened in your default PDF viewer")
            print("2. Please manually press Ctrl+P or use File > Print")
            print("3. Select the EPSON L120 printer")
            print("4. Click Print")
            
            user_input = input("\nWere you able to print manually from the PDF viewer? (y/n): ").strip().lower()
            
            if user_input == 'y':
                print("✅ METHOD 3 (Manual Print) - SUCCESS: Manual printing works")
                method3_success = True
            else:
                print("❌ METHOD 3 (Manual Print) - FAILED: Manual printing failed")
                method3_success = False
        else:
            print(f"❌ METHOD 3 - FAILED: Could not open PDF: {result.stderr}")
            method3_success = False
            
    except Exception as e:
        print(f"❌ METHOD 3 - ERROR: {e}")
        method3_success = False
    
    # Final Summary
    print("\n" + "="*60)
    print("\n=== FINAL TEST SUMMARY ===")
    
    methods = [
        ("METHOD 1 (PowerShell)", method1_success),
        ("METHOD 2 (Win32API)", method2_success),
        ("METHOD 3 (Manual Print)", method3_success)
    ]
    
    success_count = 0
    for method_name, success in methods:
        status = "✅ SUCCESS" if success else "❌ FAILED"
        print(f"  {method_name}: {status}")
        if success:
            success_count += 1
    
    print(f"\nOverall Result: {success_count}/3 methods successful")
    
    if success_count > 0:
        print("\n🎉 GOOD NEWS: At least one print method works!")
        print("\n📝 RECOMMENDATIONS:")
        
        if method3_success:
            print("  • Manual printing works - the issue is with automated printing")
            print("  • Consider implementing a hybrid approach:")
            print("    - Open PDF in viewer automatically")
            print("    - Provide user instructions to print manually")
            print("    - Or investigate PDF viewer command-line printing options")
        
        if method1_success or method2_success:
            print("  • Automated printing partially works")
            print("  • The print commands execute but may not reach the printer")
            print("  • Check Windows Print Spooler service")
            print("  • Verify printer driver installation")
        
        print("\n🔧 NEXT STEPS:")
        print("  1. Integrate working method(s) into the main server")
        print("  2. Add proper error handling and user feedback")
        print("  3. Consider implementing a print queue monitoring system")
        print("  4. Test with different PDF files and settings")
        
    else:
        print("\n❌ BAD NEWS: No print methods are working")
        print("\n🔍 TROUBLESHOOTING NEEDED:")
        print("  1. Check if printer is properly connected and powered on")
        print("  2. Verify printer driver installation")
        print("  3. Test printing from other applications (Notepad, Word, etc.)")
        print("  4. Check Windows Print Spooler service status")
        print("  5. Review Windows Event Viewer for print-related errors")
        print("  6. Consider reinstalling printer drivers")
    
    return success_count > 0

def main():
    """Main function"""
    try:
        success = test_print_with_user_validation()
        return 0 if success else 1
    except KeyboardInterrupt:
        print("\n\n⚠️  Test interrupted by user")
        return 1
    except Exception as e:
        print(f"\n\n❌ Unexpected error: {e}")
        return 1

if __name__ == "__main__":
    sys.exit(main())
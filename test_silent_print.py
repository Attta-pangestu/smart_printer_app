#!/usr/bin/env python3
"""
Test script untuk menguji silent print service secara langsung
"""

import sys
import os
import logging
from pathlib import Path

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def test_silent_print():
    """Test silent print service dengan file PDF"""
    try:
        # Add server directory to path
        server_dir = os.path.join(os.path.dirname(__file__), 'server')
        if server_dir not in sys.path:
            sys.path.insert(0, server_dir)
        
        from silent_print_service import SilentPrintService
        
        # Initialize service
        service = SilentPrintService()
        service.printer_name = "EPSON L120 Series"
        
        # Test file path
        test_file = "test_files/sample.pdf"
        
        if not os.path.exists(test_file):
            logger.error(f"Test file not found: {test_file}")
            return False
        
        logger.info(f"Testing silent print with file: {test_file}")
        logger.info(f"Target printer: {service.printer_name}")
        
        # Test silent printing
        success, message = service.print_pdf_silent(test_file)
        
        if success:
            logger.info(f"✅ Silent print SUCCESS: {message}")
        else:
            logger.error(f"❌ Silent print FAILED: {message}")
        
        # Cleanup
        service.cleanup()
        
        return success
        
    except ImportError as e:
        logger.error(f"Failed to import silent print service: {e}")
        return False
    except Exception as e:
        logger.error(f"Test failed with error: {e}")
        return False

def test_printer_availability():
    """Test apakah printer tersedia"""
    try:
        import win32print
        
        # List all printers
        printers = [printer[2] for printer in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)]
        logger.info(f"Available printers: {printers}")
        
        # Check if EPSON L120 is available
        target_printer = "EPSON L120 Series"
        if target_printer in printers:
            logger.info(f"✅ Target printer '{target_printer}' is available")
            return True
        else:
            logger.error(f"❌ Target printer '{target_printer}' not found")
            return False
            
    except Exception as e:
        logger.error(f"Failed to check printer availability: {e}")
        return False

if __name__ == "__main__":
    logger.info("=== Testing Silent Print Service ===")
    
    # Test 1: Check printer availability
    logger.info("\n1. Testing printer availability...")
    printer_ok = test_printer_availability()
    
    if not printer_ok:
        logger.error("Printer not available, stopping tests")
        sys.exit(1)
    
    # Test 2: Test silent printing
    logger.info("\n2. Testing silent print service...")
    print_ok = test_silent_print()
    
    if print_ok:
        logger.info("\n🎉 All tests PASSED! Silent printing is working.")
    else:
        logger.error("\n💥 Tests FAILED! Silent printing needs fixing.")
        sys.exit(1)
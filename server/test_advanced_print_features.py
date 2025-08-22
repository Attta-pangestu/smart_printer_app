#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Test script untuk menguji fitur-fitur cetak lanjutan:
- Fit to page modes
- Split PDF
- Page range
- Orientasi dan color modes
"""

import os
import sys
from pathlib import Path

# Add the server directory to Python path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from silent_print_service import SilentPrintService
from models.job import PrintSettings, ColorMode, Orientation, FitToPageMode

def create_test_pdf():
    """Buat PDF test sederhana menggunakan reportlab"""
    try:
        from reportlab.pdfgen import canvas
        from reportlab.lib.pagesizes import letter, A4
        
        test_pdf_path = "test_advanced_features.pdf"
        c = canvas.Canvas(test_pdf_path, pagesize=A4)
        
        # Page 1
        c.drawString(100, 750, "Test PDF - Page 1")
        c.drawString(100, 700, "This is a test document for advanced print features")
        c.drawString(100, 650, "Testing fit to page, split PDF, and page range")
        c.showPage()
        
        # Page 2
        c.drawString(100, 750, "Test PDF - Page 2")
        c.drawString(100, 700, "Second page content")
        c.drawString(100, 650, "Color mode and orientation testing")
        c.showPage()
        
        # Page 3
        c.drawString(100, 750, "Test PDF - Page 3")
        c.drawString(100, 700, "Third page content")
        c.drawString(100, 650, "Final page for range testing")
        c.showPage()
        
        c.save()
        return test_pdf_path
    except ImportError:
        print("reportlab not installed, using existing PDF if available")
        return None

def test_fit_to_page_modes():
    """Test berbagai mode fit to page"""
    print("\n=== Testing Fit to Page Modes ===")
    
    pdf_path = create_test_pdf()
    if not pdf_path or not os.path.exists(pdf_path):
        print("No test PDF available, skipping fit to page tests")
        return
    
    try:
        service = SilentPrintService()
        
        # Test 1: Actual Size
        print("\n1. Testing Actual Size mode...")
        settings = PrintSettings(
            color_mode=ColorMode.COLOR,
            orientation=Orientation.PORTRAIT,
            copies=1,
            fit_to_page=FitToPageMode.ACTUAL_SIZE
        )
        success, message = service.print_pdf_silent(pdf_path, settings)
        print(f"   Result: {'✓' if success else '✗'} {message}")
        
        # Test 2: Fit to Paper
        print("\n2. Testing Fit to Paper mode...")
        settings = PrintSettings(
            color_mode=ColorMode.COLOR,
            orientation=Orientation.PORTRAIT,
            copies=1,
            fit_to_page=FitToPageMode.FIT_TO_PAPER
        )
        success, message = service.print_pdf_silent(pdf_path, settings)
        print(f"   Result: {'✓' if success else '✗'} {message}")
        
        # Test 3: Shrink to Fit
        print("\n3. Testing Shrink to Fit mode...")
        settings = PrintSettings(
            color_mode=ColorMode.COLOR,
            orientation=Orientation.PORTRAIT,
            copies=1,
            fit_to_page=FitToPageMode.SHRINK_TO_FIT
        )
        success, message = service.print_pdf_silent(pdf_path, settings)
        print(f"   Result: {'✓' if success else '✗'} {message}")
        
        # Test 4: Default Fit to Page
        print("\n4. Testing Default Fit to Page mode...")
        settings = PrintSettings(
            color_mode=ColorMode.COLOR,
            orientation=Orientation.PORTRAIT,
            copies=1,
            fit_to_page=FitToPageMode.FIT_TO_PAGE
        )
        success, message = service.print_pdf_silent(pdf_path, settings)
        print(f"   Result: {'✓' if success else '✗'} {message}")
        
    except Exception as e:
        print(f"Error in fit to page tests: {e}")
    finally:
        if pdf_path and os.path.exists(pdf_path):
            try:
                os.remove(pdf_path)
            except:
                pass

def test_page_range():
    """Test fitur page range"""
    print("\n=== Testing Page Range ===")
    
    pdf_path = create_test_pdf()
    if not pdf_path or not os.path.exists(pdf_path):
        print("No test PDF available, skipping page range tests")
        return
    
    try:
        service = SilentPrintService()
        
        # Test 1: Single page
        print("\n1. Testing single page (page 2)...")
        settings = PrintSettings(
            color_mode=ColorMode.COLOR,
            orientation=Orientation.PORTRAIT,
            copies=1,
            page_range="2",
            fit_to_page=FitToPageMode.FIT_TO_PAGE
        )
        success, message = service.print_pdf_silent(pdf_path, settings)
        print(f"   Result: {'✓' if success else '✗'} {message}")
        
        # Test 2: Page range
        print("\n2. Testing page range (pages 1-2)...")
        settings = PrintSettings(
            color_mode=ColorMode.COLOR,
            orientation=Orientation.PORTRAIT,
            copies=1,
            page_range="1-2",
            fit_to_page=FitToPageMode.FIT_TO_PAGE
        )
        success, message = service.print_pdf_silent(pdf_path, settings)
        print(f"   Result: {'✓' if success else '✗'} {message}")
        
        # Test 3: Mixed range
        print("\n3. Testing mixed range (pages 1,3)...")
        settings = PrintSettings(
            color_mode=ColorMode.COLOR,
            orientation=Orientation.PORTRAIT,
            copies=1,
            page_range="1,3",
            fit_to_page=FitToPageMode.FIT_TO_PAGE
        )
        success, message = service.print_pdf_silent(pdf_path, settings)
        print(f"   Result: {'✓' if success else '✗'} {message}")
        
    except Exception as e:
        print(f"Error in page range tests: {e}")
    finally:
        if pdf_path and os.path.exists(pdf_path):
            try:
                os.remove(pdf_path)
            except:
                pass

def test_split_pdf():
    """Test fitur split PDF"""
    print("\n=== Testing Split PDF ===")
    
    pdf_path = create_test_pdf()
    if not pdf_path or not os.path.exists(pdf_path):
        print("No test PDF available, skipping split PDF tests")
        return
    
    try:
        service = SilentPrintService()
        
        # Test 1: Split all pages
        print("\n1. Testing split all pages...")
        settings = PrintSettings(
            color_mode=ColorMode.COLOR,
            orientation=Orientation.PORTRAIT,
            copies=1,
            split_pdf=True,
            split_output_prefix="test_page_",
            fit_to_page=FitToPageMode.FIT_TO_PAGE
        )
        success, message = service.print_pdf_silent(pdf_path, settings)
        print(f"   Result: {'✓' if success else '✗'} {message}")
        
        # Test 2: Split specific range
        print("\n2. Testing split specific range (pages 1-2)...")
        settings = PrintSettings(
            color_mode=ColorMode.COLOR,
            orientation=Orientation.PORTRAIT,
            copies=1,
            split_pdf=True,
            split_page_range="1-2",
            split_output_prefix="range_page_",
            fit_to_page=FitToPageMode.FIT_TO_PAGE
        )
        success, message = service.print_pdf_silent(pdf_path, settings)
        print(f"   Result: {'✓' if success else '✗'} {message}")
        
    except Exception as e:
        print(f"Error in split PDF tests: {e}")
    finally:
        if pdf_path and os.path.exists(pdf_path):
            try:
                os.remove(pdf_path)
            except:
                pass

def test_orientation_and_color():
    """Test orientasi dan mode warna"""
    print("\n=== Testing Orientation and Color Modes ===")
    
    pdf_path = create_test_pdf()
    if not pdf_path or not os.path.exists(pdf_path):
        print("No test PDF available, skipping orientation and color tests")
        return
    
    try:
        service = SilentPrintService()
        
        # Test 1: Landscape + Color
        print("\n1. Testing Landscape + Color...")
        settings = PrintSettings(
            color_mode=ColorMode.COLOR,
            orientation=Orientation.LANDSCAPE,
            copies=1,
            page_range="1",
            fit_to_page=FitToPageMode.FIT_TO_PAGE
        )
        success, message = service.print_pdf_silent(pdf_path, settings)
        print(f"   Result: {'✓' if success else '✗'} {message}")
        
        # Test 2: Portrait + Grayscale
        print("\n2. Testing Portrait + Grayscale...")
        settings = PrintSettings(
            color_mode=ColorMode.GRAYSCALE,
            orientation=Orientation.PORTRAIT,
            copies=1,
            page_range="2",
            fit_to_page=FitToPageMode.FIT_TO_PAGE
        )
        success, message = service.print_pdf_silent(pdf_path, settings)
        print(f"   Result: {'✓' if success else '✗'} {message}")
        
        # Test 3: Multiple copies
        print("\n3. Testing Multiple copies (2 copies)...")
        settings = PrintSettings(
            color_mode=ColorMode.COLOR,
            orientation=Orientation.PORTRAIT,
            copies=2,
            page_range="3",
            fit_to_page=FitToPageMode.FIT_TO_PAGE
        )
        success, message = service.print_pdf_silent(pdf_path, settings)
        print(f"   Result: {'✓' if success else '✗'} {message}")
        
    except Exception as e:
        print(f"Error in orientation and color tests: {e}")
    finally:
        if pdf_path and os.path.exists(pdf_path):
            try:
                os.remove(pdf_path)
            except:
                pass

def main():
    """Main test function"""
    print("Advanced Print Features Test Suite")
    print("=" * 50)
    
    try:
        # Test all advanced features
        test_fit_to_page_modes()
        test_page_range()
        test_split_pdf()
        test_orientation_and_color()
        
        print("\n" + "=" * 50)
        print("All tests completed!")
        print("Check your printer output to verify the results.")
        
    except Exception as e:
        print(f"\nTest suite failed with error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
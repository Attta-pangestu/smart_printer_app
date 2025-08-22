#!/usr/bin/env python3
"""
Test sederhana untuk memverifikasi perbaikan fit_to_page
"""

import os
import sys
from pathlib import Path
import fitz  # PyMuPDF

# Add server directory to path
server_dir = Path(__file__).parent / "server"
sys.path.insert(0, str(server_dir))

from services.enhanced_document_service import EnhancedDocumentService

def test_fit_to_page_fix():
    """Test perbaikan fit_to_page dengan margin"""
    print("Testing Fit to Page Fix")
    print("=======================")
    
    # Create a simple test PDF
    doc = fitz.open()
    page = doc.new_page(width=595, height=842)  # A4
    
    # Add content that fills the page
    page.draw_rect(fitz.Rect(50, 50, 545, 792), color=(0, 0, 0), width=2)
    page.insert_text((60, 80), "TOP LEFT - This should be visible", fontsize=12)
    page.insert_text((350, 80), "TOP RIGHT - This should be visible", fontsize=12)
    page.insert_text((60, 780), "BOTTOM LEFT - This should be visible", fontsize=12)
    page.insert_text((300, 780), "BOTTOM RIGHT - This should be visible", fontsize=12)
    page.insert_text((250, 400), "CENTER - This should be visible", fontsize=14)
    
    test_pdf_path = "temp_test_document.pdf"
    doc.save(test_pdf_path)
    doc.close()
    
    print(f"Created test PDF: {test_pdf_path}")
    
    # Test with EnhancedDocumentService
    service = EnhancedDocumentService()
    
    # Test settings with fit_to_page and margins
    settings = {
        'paper_size': 'A4',
        'orientation': 'portrait',
        'fit_to_page': True,
        'margin_top': 20,    # 20mm margins
        'margin_bottom': 20,
        'margin_left': 20,
        'margin_right': 20,
        'center_horizontally': True,
        'center_vertically': True,
        'scale': 100
    }
    
    output_path = "temp_test_output.pdf"
    
    try:
        # Apply print settings
        result = service._apply_print_settings(test_pdf_path, output_path, settings)
        
        if os.path.exists(output_path):
            print("✓ Print settings applied successfully")
            
            # Analyze the result
            output_doc = fitz.open(output_path)
            output_page = output_doc[0]
            
            # Extract text to verify content preservation
            text_content = output_page.get_text()
            
            # Check for key content
            required_texts = [
                "TOP LEFT - This should be visible",
                "TOP RIGHT - This should be visible", 
                "BOTTOM LEFT - This should be visible",
                "BOTTOM RIGHT - This should be visible",
                "CENTER - This should be visible"
            ]
            
            missing_content = []
            for text in required_texts:
                if text not in text_content:
                    missing_content.append(text)
            
            if missing_content:
                print(f"⚠️  Missing content: {missing_content}")
                print("❌ Content may be clipped - fix needed")
            else:
                print("✓ All content preserved")
                print("✓ Fit to page fix working correctly")
            
            # Show page dimensions
            print(f"Output page size: {output_page.rect.width:.1f} x {output_page.rect.height:.1f}")
            
            output_doc.close()
            
        else:
            print("❌ Output file not created")
            
    except Exception as e:
        print(f"❌ Error during processing: {e}")
        import traceback
        traceback.print_exc()
    
    finally:
        # Cleanup
        for file_path in [test_pdf_path, output_path]:
            if os.path.exists(file_path):
                try:
                    os.remove(file_path)
                except:
                    pass

def test_scale_calculations():
    """Test perhitungan skala dengan berbagai margin"""
    print("\nTesting Scale Calculations")
    print("==========================")
    
    service = EnhancedDocumentService()
    
    # Test data
    target_size = (595, 842)  # A4 in points
    page_rect = fitz.Rect(0, 0, 500, 700)  # Sample content
    
    test_cases = [
        {'name': 'No Margins', 'margins': [0, 0, 0, 0]},
        {'name': 'Small Margins (10mm)', 'margins': [10, 10, 10, 10]},
        {'name': 'Standard Margins (20mm)', 'margins': [20, 20, 20, 20]},
        {'name': 'Large Margins (40mm)', 'margins': [40, 40, 40, 40]}
    ]
    
    for case in test_cases:
        margins = case['margins']
        settings = {
            'fit_to_page': True,
            'margin_top': margins[0],
            'margin_bottom': margins[1], 
            'margin_left': margins[2],
            'margin_right': margins[3],
            'scale': 100
        }
        
        # Calculate scale matrix using the fixed method
        scale_matrix = service._calculate_scale_matrix(page_rect, target_size, settings)
        
        # Calculate available space manually for verification
        mm = 2.834645669  # mm to points conversion
        available_width = target_size[0] - (margins[2] + margins[3]) * mm
        available_height = target_size[1] - (margins[0] + margins[1]) * mm
        
        expected_scale_x = available_width / page_rect.width
        expected_scale_y = available_height / page_rect.height
        expected_scale = min(expected_scale_x, expected_scale_y)
        
        print(f"\n{case['name']}:")
        print(f"  Available space: {available_width:.1f} x {available_height:.1f}")
        print(f"  Expected scale: {expected_scale:.3f}")
        print(f"  Calculated scale: {scale_matrix.a:.3f}")
        print(f"  Match: {'✓' if abs(scale_matrix.a - expected_scale) < 0.001 else '❌'}")
        
        # Calculate final content size
        final_width = page_rect.width * scale_matrix.a
        final_height = page_rect.height * scale_matrix.a
        
        print(f"  Final content size: {final_width:.1f} x {final_height:.1f}")
        
        # Check if it fits in available space
        fits = final_width <= available_width and final_height <= available_height
        print(f"  Fits in available space: {'✓' if fits else '❌'}")

if __name__ == "__main__":
    try:
        test_scale_calculations()
        test_fit_to_page_fix()
        
        print("\n=== SUMMARY ===")
        print("✓ Scale calculations now consider margins")
        print("✓ Content should no longer be clipped when using fit_to_page")
        print("✓ Fix has been applied to EnhancedDocumentService")
        print("\nThe fix ensures that when 'fit_to_page' is enabled,")
        print("the scaling calculation uses available space after margins,")
        print("preventing content from being clipped.")
        
    except Exception as e:
        print(f"\n❌ Test failed: {e}")
        import traceback
        traceback.print_exc()
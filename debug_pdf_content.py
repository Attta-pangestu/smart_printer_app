#!/usr/bin/env python3
"""
Debug script untuk melihat konten PDF sebelum dan sesudah processing
"""

import os
import sys
from pathlib import Path
import fitz  # PyMuPDF

# Add server directory to path
server_dir = Path(__file__).parent / "server"
sys.path.insert(0, str(server_dir))

from services.enhanced_document_service import EnhancedDocumentService

def debug_pdf_processing():
    """Debug PDF processing step by step"""
    print("PDF Processing Debug")
    print("===================")
    
    # Create a simple test PDF
    doc = fitz.open()
    page = doc.new_page(width=595, height=842)  # A4
    
    # Add simple content
    page.insert_text((100, 100), "TOP TEXT", fontsize=20)
    page.insert_text((100, 400), "MIDDLE TEXT", fontsize=20)
    page.insert_text((100, 700), "BOTTOM TEXT", fontsize=20)
    
    test_pdf_path = "debug_input.pdf"
    doc.save(test_pdf_path)
    doc.close()
    
    print(f"Created input PDF: {test_pdf_path}")
    
    # Analyze input PDF
    input_doc = fitz.open(test_pdf_path)
    input_page = input_doc[0]
    input_text = input_page.get_text()
    print(f"Input PDF text: {repr(input_text)}")
    print(f"Input page size: {input_page.rect.width} x {input_page.rect.height}")
    input_doc.close()
    
    # Process with EnhancedDocumentService
    service = EnhancedDocumentService()
    
    settings = {
        'paper_size': 'A4',
        'orientation': 'portrait',
        'fit_to_page': True,
        'margin_top': 20,
        'margin_bottom': 20,
        'margin_left': 20,
        'margin_right': 20,
        'center_horizontally': True,
        'center_vertically': True,
        'scale': 100
    }
    
    output_path = "debug_output.pdf"
    
    try:
        # Apply print settings
        result = service._apply_print_settings(test_pdf_path, output_path, settings)
        
        if os.path.exists(output_path):
            print("\n✓ Processing completed")
            
            # Analyze output PDF
            output_doc = fitz.open(output_path)
            output_page = output_doc[0]
            output_text = output_page.get_text()
            
            print(f"Output PDF text: {repr(output_text)}")
            print(f"Output page size: {output_page.rect.width} x {output_page.rect.height}")
            
            # Check if text is preserved
            input_texts = ["TOP TEXT", "MIDDLE TEXT", "BOTTOM TEXT"]
            for text in input_texts:
                if text in output_text:
                    print(f"✓ Found: {text}")
                else:
                    print(f"❌ Missing: {text}")
            
            # Get image info
            images = output_page.get_images()
            print(f"\nImages in output: {len(images)}")
            
            if images:
                for i, img in enumerate(images):
                    print(f"  Image {i}: {img}")
                    
                # Get image rect
                img_rects = output_page.get_image_rects(images[0][0])
                if img_rects:
                    print(f"  Image rect: {img_rects[0]}")
            
            output_doc.close()
            
        else:
            print("❌ Output file not created")
            
    except Exception as e:
        print(f"❌ Error: {e}")
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

def test_manual_scaling():
    """Test manual scaling approach"""
    print("\nManual Scaling Test")
    print("==================")
    
    # Create test PDF
    doc = fitz.open()
    page = doc.new_page(width=500, height=700)
    page.insert_text((50, 50), "TEST CONTENT", fontsize=16)
    
    test_pdf = "manual_test.pdf"
    doc.save(test_pdf)
    doc.close()
    
    # Manual processing
    input_doc = fitz.open(test_pdf)
    input_page = input_doc[0]
    
    # Target A4 with margins
    target_width = 595
    target_height = 842
    margin = 20 * 2.834645669  # 20mm to points
    
    available_width = target_width - 2 * margin
    available_height = target_height - 2 * margin
    
    # Calculate scale
    scale_x = available_width / input_page.rect.width
    scale_y = available_height / input_page.rect.height
    scale = min(scale_x, scale_y)
    
    print(f"Input size: {input_page.rect.width} x {input_page.rect.height}")
    print(f"Available space: {available_width:.1f} x {available_height:.1f}")
    print(f"Scale: {scale:.3f}")
    
    # Create output
    output_doc = fitz.open()
    output_page = output_doc.new_page(width=target_width, height=target_height)
    
    # Calculate final size and position
    final_width = input_page.rect.width * scale
    final_height = input_page.rect.height * scale
    
    x = margin + (available_width - final_width) / 2
    y = margin + (available_height - final_height) / 2
    
    insert_rect = fitz.Rect(x, y, x + final_width, y + final_height)
    
    print(f"Final size: {final_width:.1f} x {final_height:.1f}")
    print(f"Insert rect: {insert_rect}")
    
    # Insert image
    pix = input_page.get_pixmap()
    output_page.insert_image(insert_rect, pixmap=pix)
    
    # Check result
    output_text = output_page.get_text()
    print(f"Output text: {repr(output_text)}")
    
    if "TEST CONTENT" in output_text:
        print("✓ Manual scaling preserves content")
    else:
        print("❌ Manual scaling loses content")
    
    # Cleanup
    input_doc.close()
    output_doc.close()
    
    for file_path in [test_pdf]:
        if os.path.exists(file_path):
            try:
                os.remove(file_path)
            except:
                pass

if __name__ == "__main__":
    debug_pdf_processing()
    test_manual_scaling()
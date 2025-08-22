#!/usr/bin/env python3
"""
Debug scaling logic step by step
"""

import os
import sys
from pathlib import Path
import fitz  # PyMuPDF

# Add server directory to path
server_dir = Path(__file__).parent / "server"
sys.path.insert(0, str(server_dir))

from services.enhanced_document_service import EnhancedDocumentService

def debug_scaling_calculation():
    """Debug perhitungan scaling step by step"""
    print("Scaling Logic Debug")
    print("==================")
    
    service = EnhancedDocumentService()
    
    # Test case: A4 document with 20mm margins
    page_rect = fitz.Rect(0, 0, 595, 842)  # A4 input
    target_size = (595, 842)  # A4 output
    
    settings = {
        'fit_to_page': True,
        'margin_top': 20,
        'margin_bottom': 20,
        'margin_left': 20,
        'margin_right': 20,
        'center_horizontally': True,
        'center_vertically': True,
        'scale': 100
    }
    
    mm = 2.834645669
    
    print(f"Input page size: {page_rect.width} x {page_rect.height}")
    print(f"Target page size: {target_size[0]} x {target_size[1]}")
    print(f"Margins: {settings['margin_left']}mm = {settings['margin_left'] * mm:.1f} points")
    
    # Step 1: Calculate scale matrix
    print("\n--- Step 1: Calculate Scale Matrix ---")
    
    margin_left = settings.get('margin_left', 20) * mm
    margin_top = settings.get('margin_top', 20) * mm
    margin_right = settings.get('margin_right', 20) * mm
    margin_bottom = settings.get('margin_bottom', 20) * mm
    
    available_width = target_size[0] - margin_left - margin_right
    available_height = target_size[1] - margin_top - margin_bottom
    
    print(f"Available width: {target_size[0]} - {margin_left:.1f} - {margin_right:.1f} = {available_width:.1f}")
    print(f"Available height: {target_size[1]} - {margin_top:.1f} - {margin_bottom:.1f} = {available_height:.1f}")
    
    scale_x = available_width / page_rect.width
    scale_y = available_height / page_rect.height
    scale = min(scale_x, scale_y)
    
    print(f"Scale X: {available_width:.1f} / {page_rect.width} = {scale_x:.3f}")
    print(f"Scale Y: {available_height:.1f} / {page_rect.height} = {scale_y:.3f}")
    print(f"Final scale: min({scale_x:.3f}, {scale_y:.3f}) = {scale:.3f}")
    
    scale_matrix = service._calculate_scale_matrix(page_rect, target_size, settings)
    print(f"Service calculated scale: {scale_matrix.a:.3f}")
    
    # Step 2: Calculate scaled dimensions
    print("\n--- Step 2: Calculate Scaled Dimensions ---")
    
    scaled_width = page_rect.width * scale
    scaled_height = page_rect.height * scale
    
    print(f"Scaled width: {page_rect.width} * {scale:.3f} = {scaled_width:.1f}")
    print(f"Scaled height: {page_rect.height} * {scale:.3f} = {scaled_height:.1f}")
    
    # Step 3: Calculate insert position
    print("\n--- Step 3: Calculate Insert Position ---")
    
    # Horizontal centering
    x = margin_left + (available_width - scaled_width) / 2
    print(f"X position: {margin_left:.1f} + ({available_width:.1f} - {scaled_width:.1f}) / 2 = {x:.1f}")
    
    # Vertical centering
    y = margin_top + (available_height - scaled_height) / 2
    print(f"Y position: {margin_top:.1f} + ({available_height:.1f} - {scaled_height:.1f}) / 2 = {y:.1f}")
    
    insert_rect = service._calculate_insert_rect(page_rect, target_size, settings, scale_matrix)
    print(f"Service calculated rect: {insert_rect}")
    
    # Step 4: Verify margins
    print("\n--- Step 4: Verify Final Margins ---")
    
    final_margin_left = insert_rect.x0
    final_margin_top = insert_rect.y0
    final_margin_right = target_size[0] - insert_rect.x1
    final_margin_bottom = target_size[1] - insert_rect.y1
    
    print(f"Final margins:")
    print(f"  Left: {final_margin_left:.1f} (expected: {margin_left:.1f})")
    print(f"  Top: {final_margin_top:.1f} (expected: {margin_top:.1f})")
    print(f"  Right: {final_margin_right:.1f} (expected: {margin_right:.1f})")
    print(f"  Bottom: {final_margin_bottom:.1f} (expected: {margin_bottom:.1f})")
    
    # Check if content fits
    content_width = insert_rect.width
    content_height = insert_rect.height
    
    print(f"\nContent size: {content_width:.1f} x {content_height:.1f}")
    print(f"Available space: {available_width:.1f} x {available_height:.1f}")
    
    fits_width = content_width <= available_width
    fits_height = content_height <= available_height
    
    print(f"Fits width: {'✓' if fits_width else '❌'} ({content_width:.1f} <= {available_width:.1f})")
    print(f"Fits height: {'✓' if fits_height else '❌'} ({content_height:.1f} <= {available_height:.1f})")
    
    return {
        'scale': scale,
        'scaled_width': scaled_width,
        'scaled_height': scaled_height,
        'insert_rect': insert_rect,
        'fits': fits_width and fits_height
    }

def debug_different_input_sizes():
    """Debug dengan berbagai ukuran input"""
    print("\n" + "="*60)
    print("Different Input Sizes Debug")
    print("="*60)
    
    service = EnhancedDocumentService()
    
    test_cases = [
        {'size': (595, 842), 'name': 'A4 Same Size'},
        {'size': (400, 600), 'name': 'Smaller Portrait'},
        {'size': (700, 500), 'name': 'Landscape'},
        {'size': (300, 400), 'name': 'Very Small'}
    ]
    
    target_size = (595, 842)  # A4
    
    settings = {
        'fit_to_page': True,
        'margin_top': 20,
        'margin_bottom': 20,
        'margin_left': 20,
        'margin_right': 20,
        'center_horizontally': True,
        'center_vertically': True,
        'scale': 100
    }
    
    mm = 2.834645669
    expected_margin = 20 * mm
    
    for case in test_cases:
        print(f"\n--- {case['name']} ({case['size'][0]} x {case['size'][1]}) ---")
        
        page_rect = fitz.Rect(0, 0, case['size'][0], case['size'][1])
        
        # Calculate expected scale
        available_width = target_size[0] - 2 * expected_margin
        available_height = target_size[1] - 2 * expected_margin
        
        scale_x = available_width / page_rect.width
        scale_y = available_height / page_rect.height
        expected_scale = min(scale_x, scale_y)
        
        print(f"Expected scale: {expected_scale:.3f}")
        
        # Calculate using service
        scale_matrix = service._calculate_scale_matrix(page_rect, target_size, settings)
        insert_rect = service._calculate_insert_rect(page_rect, target_size, settings, scale_matrix)
        
        print(f"Service scale: {scale_matrix.a:.3f}")
        print(f"Insert rect: {insert_rect}")
        
        # Check margins
        margin_left = insert_rect.x0
        margin_top = insert_rect.y0
        margin_right = target_size[0] - insert_rect.x1
        margin_bottom = target_size[1] - insert_rect.y1
        
        print(f"Margins: L={margin_left:.1f}, T={margin_top:.1f}, R={margin_right:.1f}, B={margin_bottom:.1f}")
        print(f"Expected: {expected_margin:.1f}")
        
        # Check accuracy
        margin_accuracy = (
            abs(margin_left - expected_margin) < 1 and
            abs(margin_right - expected_margin) < 1 and
            abs(margin_top - expected_margin) < 1 and
            abs(margin_bottom - expected_margin) < 1
        )
        
        print(f"Margin accuracy: {'✓' if margin_accuracy else '❌'}")

if __name__ == "__main__":
    result = debug_scaling_calculation()
    debug_different_input_sizes()
    
    print("\n" + "="*60)
    print("DEBUG SUMMARY")
    print("="*60)
    
    if result['fits']:
        print("✓ Content fits within available space")
    else:
        print("❌ Content exceeds available space")
    
    print("\nThis debug helps identify where the margin calculation goes wrong.")
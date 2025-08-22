#!/usr/bin/env python3
"""
Debug script to analyze aspect ratio issues with fit_to_page
"""

import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), 'server'))

from services.enhanced_document_service import EnhancedDocumentService
from reportlab.lib.units import mm
from reportlab.lib.pagesizes import A4
import fitz

def create_test_pdf(filename, width_mm, height_mm):
    """Create a test PDF with specific dimensions"""
    from reportlab.pdfgen import canvas
    from reportlab.lib.units import mm
    
    width_points = width_mm * mm
    height_points = height_mm * mm
    
    c = canvas.Canvas(filename, pagesize=(width_points, height_points))
    c.setFillColorRGB(0.8, 0.8, 0.8)  # Light gray background
    c.rect(0, 0, width_points, height_points, fill=1)
    
    # Add some text to make it visible
    c.setFillColorRGB(0, 0, 0)  # Black text
    c.setFont("Helvetica", 12)
    c.drawString(10, height_points - 20, f"Test content: {width_mm}x{height_mm}mm")
    c.drawString(10, height_points - 40, f"Aspect ratio: {width_mm/height_mm:.2f}")
    
    c.save()
    return filename

def analyze_aspect_ratio_issue():
    print("Aspect Ratio Analysis for fit_to_page")
    print("=====================================")
    
    # A4 dimensions
    a4_width_mm = 210
    a4_height_mm = 297
    a4_aspect = a4_width_mm / a4_height_mm
    print(f"A4 page: {a4_width_mm}x{a4_height_mm}mm, aspect ratio: {a4_aspect:.3f}")
    
    # Test margin
    margin_mm = 20
    available_width_mm = a4_width_mm - 2 * margin_mm
    available_height_mm = a4_height_mm - 2 * margin_mm
    available_aspect = available_width_mm / available_height_mm
    print(f"Available space: {available_width_mm}x{available_height_mm}mm, aspect ratio: {available_aspect:.3f}")
    print()
    
    # Test different content aspect ratios
    test_cases = [
        (100, 100, "Square content"),
        (150, 100, "Wide content"),
        (100, 150, "Tall content"),
        (170, 100, "Very wide content"),
        (100, 200, "Very tall content")
    ]
    
    service = EnhancedDocumentService()
    
    for width_mm, height_mm, description in test_cases:
        print(f"--- {description} ({width_mm}x{height_mm}mm) ---")
        content_aspect = width_mm / height_mm
        print(f"Content aspect ratio: {content_aspect:.3f}")
        
        # Calculate scaling
        scale_x = available_width_mm / width_mm
        scale_y = available_height_mm / height_mm
        scale = min(scale_x, scale_y)
        
        print(f"Scale X: {scale_x:.3f}, Scale Y: {scale_y:.3f}")
        print(f"Used scale: {scale:.3f} ({'limited by width' if scale == scale_x else 'limited by height'})")
        
        # Calculate actual dimensions after scaling
        scaled_width_mm = width_mm * scale
        scaled_height_mm = height_mm * scale
        
        print(f"Scaled dimensions: {scaled_width_mm:.1f}x{scaled_height_mm:.1f}mm")
        
        # Calculate unused space
        unused_width = available_width_mm - scaled_width_mm
        unused_height = available_height_mm - scaled_height_mm
        
        print(f"Unused space: width={unused_width:.1f}mm, height={unused_height:.1f}mm")
        
        # This unused space becomes extra margin when content is positioned at top-left
        extra_margin_right = unused_width
        extra_margin_bottom = unused_height
        
        actual_margin_right = margin_mm + extra_margin_right
        actual_margin_bottom = margin_mm + extra_margin_bottom
        
        print(f"Actual margins: left={margin_mm}mm, top={margin_mm}mm, right={actual_margin_right:.1f}mm, bottom={actual_margin_bottom:.1f}mm")
        print()

if __name__ == "__main__":
    analyze_aspect_ratio_issue()
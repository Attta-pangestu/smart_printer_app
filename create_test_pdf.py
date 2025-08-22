#!/usr/bin/env python3
"""
Script untuk membuat file PDF test sederhana
"""

import os
from pathlib import Path

def create_test_pdf():
    """Buat file PDF test sederhana"""
    try:
        from reportlab.pdfgen import canvas
        from reportlab.lib.pagesizes import letter
        
        # Buat direktori test_files jika belum ada
        test_dir = Path("test_files")
        test_dir.mkdir(exist_ok=True)
        
        # Path file PDF
        pdf_path = test_dir / "sample.pdf"
        
        # Buat PDF sederhana
        c = canvas.Canvas(str(pdf_path), pagesize=letter)
        width, height = letter
        
        # Tulis teks
        c.setFont("Helvetica", 12)
        c.drawString(100, height - 100, "Test PDF untuk Silent Print Service")
        c.drawString(100, height - 130, "Ini adalah file PDF test sederhana")
        c.drawString(100, height - 160, "Dibuat untuk menguji pencetakan otomatis")
        c.drawString(100, height - 190, "Tanggal: 2025-08-21")
        
        # Tambah beberapa baris lagi
        y_pos = height - 250
        for i in range(5):
            c.drawString(100, y_pos - (i * 20), f"Baris test {i + 1}: Lorem ipsum dolor sit amet")
        
        # Simpan PDF
        c.save()
        
        print(f"✅ PDF test berhasil dibuat: {pdf_path}")
        print(f"📁 Ukuran file: {pdf_path.stat().st_size} bytes")
        
        return str(pdf_path)
        
    except ImportError:
        print("❌ reportlab tidak tersedia, mencoba dengan alternatif...")
        return create_simple_pdf_alternative()
    except Exception as e:
        print(f"❌ Error membuat PDF: {e}")
        return None

def create_simple_pdf_alternative():
    """Alternatif membuat PDF tanpa reportlab"""
    try:
        # Buat direktori test_files jika belum ada
        test_dir = Path("test_files")
        test_dir.mkdir(exist_ok=True)
        
        # Path file PDF
        pdf_path = test_dir / "sample.pdf"
        
        # PDF minimal yang valid
        pdf_content = """%PDF-1.4
1 0 obj
<<
/Type /Catalog
/Pages 2 0 R
>>
endobj

2 0 obj
<<
/Type /Pages
/Kids [3 0 R]
/Count 1
>>
endobj

3 0 obj
<<
/Type /Page
/Parent 2 0 R
/MediaBox [0 0 612 792]
/Contents 4 0 R
/Resources <<
/Font <<
/F1 5 0 R
>>
>>
>>
endobj

4 0 obj
<<
/Length 100
>>
stream
BT
/F1 12 Tf
100 700 Td
(Test PDF untuk Silent Print Service) Tj
0 -20 Td
(Ini adalah file PDF test sederhana) Tj
ET
endstream
endobj

5 0 obj
<<
/Type /Font
/Subtype /Type1
/BaseFont /Helvetica
>>
endobj

xref
0 6
0000000000 65535 f 
0000000009 00000 n 
0000000058 00000 n 
0000000115 00000 n 
0000000274 00000 n 
0000000424 00000 n 
trailer
<<
/Size 6
/Root 1 0 R
>>
startxref
492
%%EOF"""
        
        # Tulis file PDF
        with open(pdf_path, 'w', encoding='latin-1') as f:
            f.write(pdf_content)
        
        print(f"✅ PDF test sederhana berhasil dibuat: {pdf_path}")
        print(f"📁 Ukuran file: {pdf_path.stat().st_size} bytes")
        
        return str(pdf_path)
        
    except Exception as e:
        print(f"❌ Error membuat PDF alternatif: {e}")
        return None

if __name__ == "__main__":
    print("=== Membuat File PDF Test ===")
    pdf_file = create_test_pdf()
    
    if pdf_file:
        print(f"\n🎉 File PDF siap untuk testing: {pdf_file}")
    else:
        print("\n💥 Gagal membuat file PDF test")
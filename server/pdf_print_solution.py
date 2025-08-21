#!/usr/bin/env python3
"""
PDF Print Solution
Solusi untuk mengatasi masalah pencetakan PDF dengan metode alternatif
"""

import os
import sys
import subprocess
import tempfile
import time
from pathlib import Path
import win32print
import win32api
from PIL import Image, ImageDraw, ImageFont
import fitz  # PyMuPDF

class PDFPrintSolution:
    """Solusi pencetakan PDF dengan berbagai metode fallback"""
    
    def __init__(self):
        self.printer_name = None
        self.temp_dir = tempfile.mkdtemp(prefix="pdf_print_")
        
    def find_printer(self, pattern="EPSON L120"):
        """Cari printer berdasarkan pola nama"""
        try:
            printers = [printer[2] for printer in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)]
            
            for printer in printers:
                if pattern.lower() in printer.lower():
                    self.printer_name = printer
                    return printer
            
            # Jika tidak ditemukan, gunakan default printer
            default_printer = win32print.GetDefaultPrinter()
            self.printer_name = default_printer
            return default_printer
            
        except Exception as e:
            print(f"Error finding printer: {e}")
            return None
    
    def pdf_to_images(self, pdf_path):
        """Konversi PDF ke gambar menggunakan PyMuPDF"""
        try:
            doc = fitz.open(pdf_path)
            image_paths = []
            
            for page_num in range(len(doc)):
                page = doc.load_page(page_num)
                
                # Render halaman sebagai gambar dengan resolusi tinggi
                mat = fitz.Matrix(2.0, 2.0)  # 2x zoom untuk kualitas lebih baik
                pix = page.get_pixmap(matrix=mat)
                
                # Simpan sebagai PNG
                image_path = os.path.join(self.temp_dir, f"page_{page_num + 1}.png")
                pix.save(image_path)
                image_paths.append(image_path)
            
            doc.close()
            return image_paths
            
        except Exception as e:
            print(f"Error converting PDF to images: {e}")
            return []
    
    def print_image_with_pil(self, image_path, printer_name):
        """Cetak gambar menggunakan PIL dan win32api"""
        try:
            # Buka gambar
            img = Image.open(image_path)
            
            # Konversi ke mode RGB jika perlu
            if img.mode != 'RGB':
                img = img.convert('RGB')
            
            # Simpan sebagai BMP sementara (format yang didukung win32api)
            temp_bmp = os.path.join(self.temp_dir, "temp_print.bmp")
            img.save(temp_bmp, "BMP")
            
            # Cetak menggunakan win32api
            result = win32api.ShellExecute(
                0,
                "print",
                temp_bmp,
                f'/d:"{printer_name}"',
                ".",
                0
            )
            
            # Hapus file sementara
            time.sleep(2)  # Tunggu sebentar sebelum menghapus
            if os.path.exists(temp_bmp):
                os.remove(temp_bmp)
            
            return result > 32  # ShellExecute mengembalikan > 32 jika berhasil
            
        except Exception as e:
            print(f"Error printing image: {e}")
            return False
    
    def print_with_notepad_conversion(self, pdf_path, printer_name):
        """Konversi PDF ke teks dan cetak dengan notepad"""
        try:
            # Ekstrak teks dari PDF
            doc = fitz.open(pdf_path)
            text_content = ""
            
            for page_num in range(len(doc)):
                page = doc.load_page(page_num)
                text_content += f"\n--- Page {page_num + 1} ---\n"
                text_content += page.get_text()
                text_content += "\n\n"
            
            doc.close()
            
            # Simpan sebagai file teks
            text_file = os.path.join(self.temp_dir, "pdf_content.txt")
            with open(text_file, 'w', encoding='utf-8') as f:
                f.write(text_content)
            
            # Cetak dengan notepad
            result = subprocess.run(
                ['notepad', '/p', text_file],
                capture_output=True,
                text=True,
                timeout=30
            )
            
            # Hapus file sementara
            time.sleep(2)
            if os.path.exists(text_file):
                os.remove(text_file)
            
            return result.returncode == 0
            
        except Exception as e:
            print(f"Error printing with notepad conversion: {e}")
            return False
    
    def print_with_powershell_image(self, image_path, printer_name):
        """Cetak gambar menggunakan PowerShell"""
        try:
            ps_script = f'''
            Add-Type -AssemblyName System.Drawing
            Add-Type -AssemblyName System.Windows.Forms
            
            $image = [System.Drawing.Image]::FromFile("{image_path}")
            $printDocument = New-Object System.Drawing.Printing.PrintDocument
            $printDocument.PrinterSettings.PrinterName = "{printer_name}"
            
            $printDocument.add_PrintPage({{
                param($sender, $e)
                $e.Graphics.DrawImage($image, $e.MarginBounds)
            }})
            
            $printDocument.Print()
            $image.Dispose()
            '''
            
            result = subprocess.run(
                ['powershell', '-Command', ps_script],
                capture_output=True,
                text=True,
                timeout=30
            )
            
            return result.returncode == 0
            
        except Exception as e:
            print(f"Error printing with PowerShell: {e}")
            return False
    
    def print_pdf(self, pdf_path, settings=None):
        """Cetak PDF dengan berbagai metode fallback"""
        if not self.printer_name:
            if not self.find_printer():
                return False, "No printer found"
        
        if not os.path.exists(pdf_path):
            return False, f"PDF file not found: {pdf_path}"
        
        print(f"Attempting to print {pdf_path} to {self.printer_name}")
        
        # Method 1: Konversi PDF ke gambar dan cetak
        print("\nMethod 1: PDF to Images...")
        try:
            image_paths = self.pdf_to_images(pdf_path)
            if image_paths:
                print(f"  Converted to {len(image_paths)} images")
                
                success_count = 0
                for i, image_path in enumerate(image_paths):
                    print(f"  Printing page {i + 1}...")
                    
                    # Try PIL method
                    if self.print_image_with_pil(image_path, self.printer_name):
                        print(f"    Page {i + 1} printed successfully (PIL)")
                        success_count += 1
                    else:
                        # Try PowerShell method
                        if self.print_with_powershell_image(image_path, self.printer_name):
                            print(f"    Page {i + 1} printed successfully (PowerShell)")
                            success_count += 1
                        else:
                            print(f"    Page {i + 1} failed to print")
                    
                    # Cleanup
                    if os.path.exists(image_path):
                        os.remove(image_path)
                
                if success_count > 0:
                    return True, f"Printed {success_count}/{len(image_paths)} pages successfully"
                else:
                    print("  All pages failed to print")
            else:
                print("  Failed to convert PDF to images")
        except Exception as e:
            print(f"  Method 1 error: {e}")
        
        # Method 2: Konversi ke teks dan cetak dengan notepad
        print("\nMethod 2: Text Conversion...")
        try:
            if self.print_with_notepad_conversion(pdf_path, self.printer_name):
                return True, "PDF printed as text successfully"
            else:
                print("  Text conversion method failed")
        except Exception as e:
            print(f"  Method 2 error: {e}")
        
        # Method 3: Buka PDF dan minta user cetak manual
        print("\nMethod 3: Manual Print...")
        try:
            subprocess.run(['start', '', pdf_path], shell=True, timeout=5)
            
            print("\n📋 MANUAL PRINT INSTRUCTIONS:")
            print("1. PDF has been opened in your default viewer")
            print("2. Press Ctrl+P or go to File > Print")
            print(f"3. Select printer: {self.printer_name}")
            print("4. Click Print")
            
            user_input = input("\nDid you successfully print manually? (y/n): ").strip().lower()
            
            if user_input == 'y':
                return True, "PDF printed manually by user"
            else:
                return False, "Manual printing failed"
                
        except Exception as e:
            print(f"  Method 3 error: {e}")
        
        return False, "All print methods failed"
    
    def cleanup(self):
        """Bersihkan file sementara"""
        try:
            import shutil
            if os.path.exists(self.temp_dir):
                shutil.rmtree(self.temp_dir)
        except Exception as e:
            print(f"Cleanup error: {e}")
    
    def __del__(self):
        """Destructor untuk cleanup otomatis"""
        self.cleanup()

def test_pdf_print_solution():
    """Test PDF print solution"""
    print("=== TESTING PDF PRINT SOLUTION ===")
    
    # File test
    test_file = "D:/Gawean Rebinmas/Driver_Epson_L120/temp/Test_print.pdf"
    
    if not os.path.exists(test_file):
        print(f"❌ Test file not found: {test_file}")
        return False
    
    # Inisialisasi solution
    solution = PDFPrintSolution()
    
    # Cari printer
    printer = solution.find_printer()
    if not printer:
        print("❌ No printer found")
        return False
    
    print(f"✅ Using printer: {printer}")
    
    # Test print
    success, message = solution.print_pdf(test_file)
    
    print(f"\nResult: {'✅ SUCCESS' if success else '❌ FAILED'}")
    print(f"Message: {message}")
    
    # Cleanup
    solution.cleanup()
    
    return success

if __name__ == "__main__":
    try:
        success = test_pdf_print_solution()
        sys.exit(0 if success else 1)
    except KeyboardInterrupt:
        print("\n\n⚠️  Test interrupted by user")
        sys.exit(1)
    except Exception as e:
        print(f"\n\n❌ Unexpected error: {e}")
        sys.exit(1)
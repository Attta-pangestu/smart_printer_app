#!/usr/bin/env python3
"""
Silent Print Service - True Silent Printing Without Dialogs
Implementasi pencetakan otomatis tanpa dialog menggunakan Win32 API langsung
"""

import os
import sys
import time
import tempfile
import subprocess
from pathlib import Path
import win32print
import win32api
import win32con
import win32ui
import win32gui
from PIL import Image, ImageWin
import fitz  # PyMuPDF
from typing import Optional

# Import PrintSettings model
try:
    from models.job import PrintSettings, ColorMode, Orientation, PaperSize, FitToPageMode
except ImportError:
    # Fallback if models not available
    PrintSettings = None
    ColorMode = None
    Orientation = None
    PaperSize = None
    FitToPageMode = None

class SilentPrintService:
    """Service untuk pencetakan silent tanpa dialog"""
    
    def __init__(self):
        self.printer_name = None
        self.printer_handle = None
        self.temp_dir = tempfile.mkdtemp(prefix="silent_print_")
        # Print tools directory
        current_dir = os.path.dirname(os.path.dirname(__file__))
        self.print_tools_dir = os.path.join(current_dir, "print_tools")
        # Print settings
        self.print_settings = None
        self.original_printer_settings = None
        
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
    
    def open_printer(self):
        """Buka koneksi ke printer"""
        try:
            if not self.printer_name:
                if not self.find_printer():
                    return False
                    
            self.printer_handle = win32print.OpenPrinter(self.printer_name)
            return True
            
        except Exception as e:
            print(f"Error opening printer: {e}")
            return False
    
    def close_printer(self):
        """Tutup koneksi printer"""
        try:
            if self.printer_handle:
                win32print.ClosePrinter(self.printer_handle)
                self.printer_handle = None
        except Exception as e:
            print(f"Error closing printer: {e}")
    
    def set_print_settings(self, settings: Optional['PrintSettings'] = None):
        """Set print settings untuk printer"""
        if not settings or not PrintSettings:
            return True
            
        self.print_settings = settings
        
        try:
            if not self.printer_handle:
                if not self.open_printer():
                    return False
            
            # Get current printer settings
            printer_info = win32print.GetPrinter(self.printer_handle, 2)
            
            # Get printer device mode using DocumentProperties
            try:
                # Get current devmode - first get the size needed
                size_needed = win32print.DocumentProperties(0, self.printer_handle, self.printer_name, None, None, 0)
                if size_needed <= 0:
                    print("Warning: Could not determine devmode size")
                    return False
                
                # Now get the actual devmode
                devmode = win32print.DocumentProperties(0, self.printer_handle, self.printer_name, None, None, win32con.DM_OUT_BUFFER)
                if not devmode:
                    print("Warning: Could not get printer device mode")
                    return False
            except Exception as e:
                print(f"Error getting device mode: {e}")
                # Try alternative method
                try:
                    printer_info = win32print.GetPrinter(self.printer_handle, 2)
                    devmode = printer_info.get('pDevMode')
                    if not devmode:
                        print("Warning: Could not get devmode from printer info")
                        return False
                except Exception as e2:
                    print(f"Alternative method also failed: {e2}")
                    return False
            
            # Store original settings for restoration
            if not self.original_printer_settings:
                try:
                    self.original_printer_settings = {
                        'Color': getattr(devmode, 'Color', 1),
                        'Orientation': getattr(devmode, 'Orientation', 1),
                        'Copies': getattr(devmode, 'Copies', 1),
                        'PaperSize': getattr(devmode, 'PaperSize', 9),  # A4
                        'PrintQuality': getattr(devmode, 'PrintQuality', -3),  # Normal
                        'Duplex': getattr(devmode, 'Duplex', 1)  # Simplex
                    }
                except Exception as e:
                    print(f"Warning: Could not store original settings: {e}")
                    self.original_printer_settings = {}
            
            # Apply color mode
            if hasattr(settings, 'color_mode') and ColorMode:
                try:
                    if settings.color_mode == ColorMode.COLOR:
                        devmode.Color = 2  # Color
                    elif settings.color_mode == ColorMode.GRAYSCALE:
                        devmode.Color = 1  # Monochrome/Grayscale
                    elif settings.color_mode == ColorMode.BLACK_WHITE:
                        devmode.Color = 1  # Monochrome
                        devmode.PrintQuality = -4  # Draft quality for pure B&W
                    print(f"Applied color mode: {settings.color_mode} -> {devmode.Color}")
                except Exception as e:
                    print(f"Warning: Could not set color mode: {e}")
            
            # Apply orientation
            if hasattr(settings, 'orientation') and Orientation:
                try:
                    if settings.orientation == Orientation.PORTRAIT:
                        devmode.Orientation = 1  # Portrait
                    elif settings.orientation == Orientation.LANDSCAPE:
                        devmode.Orientation = 2  # Landscape
                    print(f"Applied orientation: {settings.orientation} -> {devmode.Orientation}")
                except Exception as e:
                    print(f"Warning: Could not set orientation: {e}")
            
            # Apply copies
            if hasattr(settings, 'copies'):
                try:
                    devmode.Copies = max(1, min(999, settings.copies))
                    print(f"Applied copies: {settings.copies} -> {devmode.Copies}")
                except Exception as e:
                    print(f"Warning: Could not set copies: {e}")
            
            # Apply paper size
            if hasattr(settings, 'paper_size') and PaperSize:
                try:
                    paper_size_map = {
                        PaperSize.A4: 9,
                        PaperSize.A3: 8,
                        PaperSize.A5: 11,
                        PaperSize.LETTER: 1,
                        PaperSize.LEGAL: 5,
                        PaperSize.TABLOID: 3
                    }
                    if settings.paper_size in paper_size_map:
                        devmode.PaperSize = paper_size_map[settings.paper_size]
                        print(f"Applied paper size: {settings.paper_size} -> {devmode.PaperSize}")
                except Exception as e:
                    print(f"Warning: Could not set paper size: {e}")
            
            # Apply print quality
            if hasattr(settings, 'quality'):
                try:
                    quality_map = {
                        'draft': -4,
                        'normal': -3,
                        'high': -2,
                        'photo': 600
                    }
                    if settings.quality in quality_map:
                        devmode.PrintQuality = quality_map[settings.quality]
                        print(f"Applied quality: {settings.quality} -> {devmode.PrintQuality}")
                except Exception as e:
                    print(f"Warning: Could not set print quality: {e}")
            
            # Apply duplex mode
            if hasattr(settings, 'duplex'):
                try:
                    duplex_map = {
                        'none': 1,      # Simplex
                        'horizontal': 2, # Horizontal duplex
                        'vertical': 3    # Vertical duplex
                    }
                    if settings.duplex in duplex_map:
                        devmode.Duplex = duplex_map[settings.duplex]
                        print(f"Applied duplex: {settings.duplex} -> {devmode.Duplex}")
                except Exception as e:
                    print(f"Warning: Could not set duplex mode: {e}")
            
            # Apply scale
            if hasattr(settings, 'scale'):
                try:
                    devmode.Scale = max(25, min(400, settings.scale))
                    print(f"Applied scale: {settings.scale} -> {devmode.Scale}")
                except Exception as e:
                    print(f"Warning: Could not set scale: {e}")
            
            # Set the modified device mode
            try:
                # Apply settings using DocumentProperties
                result = win32print.DocumentProperties(0, self.printer_handle, self.printer_name, devmode, devmode, win32con.DM_IN_BUFFER | win32con.DM_OUT_BUFFER)
                
                # Update printer info with new devmode
                printer_info['pDevMode'] = devmode
                win32print.SetPrinter(self.printer_handle, 2, printer_info, 0)
                
                print(f"✅ Print settings applied successfully:")
                print(f"   Color Mode: {getattr(devmode, 'Color', 'N/A')} (1=Mono, 2=Color)")
                print(f"   Orientation: {getattr(devmode, 'Orientation', 'N/A')} (1=Portrait, 2=Landscape)")
                print(f"   Copies: {getattr(devmode, 'Copies', 'N/A')}")
                print(f"   Paper Size: {getattr(devmode, 'PaperSize', 'N/A')}")
                return True
            except Exception as e:
                print(f"Warning: Could not apply all printer settings: {e}")
                return False
                
        except Exception as e:
            print(f"Error setting print settings: {e}")
            return False
    
    def restore_printer_settings(self):
        """Restore original printer settings"""
        if not self.original_printer_settings:
            return True
            
        try:
            if not self.printer_handle:
                if not self.open_printer():
                    return False
            
            # Get current devmode
            try:
                devmode = win32print.DocumentProperties(0, self.printer_handle, self.printer_name, None, None, 0)
            except:
                return False
            
            if not devmode:
                return False
            
            # Restore original settings
            for key, value in self.original_printer_settings.items():
                if hasattr(devmode, key):
                    setattr(devmode, key, value)
            
            # Apply restored settings
            try:
                win32print.DocumentProperties(0, self.printer_handle, self.printer_name, devmode, devmode, win32con.DM_IN_BUFFER | win32con.DM_OUT_BUFFER)
                print("✅ Original printer settings restored")
                return True
            except Exception as e:
                print(f"Warning: Could not restore printer settings: {e}")
                return False
                
        except Exception as e:
            print(f"Error restoring printer settings: {e}")
            return False
    
    def print_raw_data(self, data, job_name="Silent Print Job"):
        """Cetak data raw langsung ke printer"""
        try:
            if not self.printer_handle:
                if not self.open_printer():
                    return False, "Cannot open printer"
            
            # Start document
            job_info = (job_name, None, "RAW")
            job_id = win32print.StartDocPrinter(self.printer_handle, 1, job_info)
            
            if job_id <= 0:
                return False, "Failed to start print job"
            
            # Start page
            win32print.StartPagePrinter(self.printer_handle)
            
            # Write data
            if isinstance(data, str):
                data = data.encode('utf-8')
            
            bytes_written = win32print.WritePrinter(self.printer_handle, data)
            
            # End page and document
            win32print.EndPagePrinter(self.printer_handle)
            win32print.EndDocPrinter(self.printer_handle)
            
            return True, f"Printed {bytes_written} bytes, Job ID: {job_id}"
            
        except Exception as e:
            return False, f"Print error: {e}"
    
    def print_text_file(self, file_path):
        """Cetak file teks langsung"""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # Add printer control codes
            printer_data = "\x1B@"  # ESC @ - Initialize printer
            printer_data += content
            printer_data += "\x0C"  # Form feed
            
            return self.print_raw_data(printer_data, f"Text File: {os.path.basename(file_path)}")
            
        except Exception as e:
            return False, f"Text file print error: {e}"
    
    def split_pdf_pages(self, pdf_path, page_range=None, output_prefix="page_"):
        """Split PDF into separate files for each page"""
        try:
            doc = fitz.open(pdf_path)
            split_files = []
            
            # Parse page range if provided
            pages_to_split = []
            if page_range:
                # Parse range like "1-5,8,11-13"
                for part in page_range.split(','):
                    if '-' in part:
                        start, end = map(int, part.split('-'))
                        pages_to_split.extend(range(start-1, end))  # Convert to 0-based
                    else:
                        pages_to_split.append(int(part)-1)  # Convert to 0-based
            else:
                # Split all pages
                pages_to_split = list(range(len(doc)))
            
            for page_num in pages_to_split:
                if page_num < len(doc):
                    # Create new document with single page
                    new_doc = fitz.open()
                    new_doc.insert_pdf(doc, from_page=page_num, to_page=page_num)
                    
                    # Save split file
                    split_filename = f"{output_prefix}{page_num + 1}.pdf"
                    split_path = os.path.join(self.temp_dir, split_filename)
                    new_doc.save(split_path)
                    new_doc.close()
                    
                    split_files.append(split_path)
            
            doc.close()
            return split_files
            
        except Exception as e:
            print(f"Error splitting PDF: {e}")
            return []
    
    def pdf_to_images(self, pdf_path, dpi=150):
        """Konversi PDF ke gambar dengan resolusi tinggi"""
        try:
            doc = fitz.open(pdf_path)
            image_paths = []
            
            # Determine which pages to convert based on page_range setting
            pages_to_convert = []
            if hasattr(self, 'print_settings') and self.print_settings and hasattr(self.print_settings, 'page_range') and self.print_settings.page_range:
                # Parse page range like "1-5,8,11-13"
                for part in self.print_settings.page_range.split(','):
                    part = part.strip()
                    if '-' in part:
                        start, end = map(int, part.split('-'))
                        pages_to_convert.extend(range(start-1, end))  # Convert to 0-based
                    else:
                        pages_to_convert.append(int(part)-1)  # Convert to 0-based
                # Remove duplicates and sort
                pages_to_convert = sorted(list(set(pages_to_convert)))
                # Filter out invalid page numbers
                pages_to_convert = [p for p in pages_to_convert if 0 <= p < len(doc)]
            else:
                # Convert all pages
                pages_to_convert = list(range(len(doc)))
            
            for page_num in pages_to_convert:
                page = doc.load_page(page_num)
                
                # Render dengan DPI tinggi
                mat = fitz.Matrix(dpi/72.0, dpi/72.0)
                
                # Check if grayscale mode is needed
                if hasattr(self, 'print_settings') and self.print_settings and hasattr(self.print_settings, 'color_mode') and ColorMode:
                    if self.print_settings.color_mode == ColorMode.GRAYSCALE or self.print_settings.color_mode == ColorMode.BLACK_WHITE:
                        # Render in grayscale
                        pix = page.get_pixmap(matrix=mat, colorspace=fitz.csGRAY)
                    else:
                        # Render in color
                        pix = page.get_pixmap(matrix=mat)
                else:
                    # Default color rendering
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
    
    def print_image_direct(self, image_path):
        """Cetak gambar langsung menggunakan Win32 GDI"""
        try:
            if not self.printer_handle:
                if not self.open_printer():
                    return False, "Cannot open printer"
            
            # Buka gambar
            img = Image.open(image_path)
            
            # Konversi ke RGB jika perlu
            if img.mode != 'RGB':
                img = img.convert('RGB')
            
            # Check for landscape orientation and rotate image if needed
            is_landscape = False
            if hasattr(self, 'print_settings') and self.print_settings and hasattr(self.print_settings, 'orientation') and Orientation:
                is_landscape = self.print_settings.orientation == Orientation.LANDSCAPE
                if is_landscape:
                    # Rotate image 90 degrees for landscape
                    img = img.rotate(-90, expand=True)
            
            # Buat device context untuk printer
            printer_dc = win32ui.CreateDC()
            printer_dc.CreatePrinterDC(self.printer_name)
            
            # Start document dengan string sederhana untuk menghindari Unicode error
            job_name = "Silent_Print_Job"
            printer_dc.StartDoc(job_name)
            printer_dc.StartPage()
            
            # Get printer capabilities
            printer_width = printer_dc.GetDeviceCaps(win32con.HORZRES)
            printer_height = printer_dc.GetDeviceCaps(win32con.VERTRES)
            
            # Calculate scaling based on fit to page mode
            img_width, img_height = img.size
            
            # Determine fit to page mode
            fit_mode = None
            if hasattr(self, 'print_settings') and self.print_settings and hasattr(self.print_settings, 'fit_to_page') and FitToPageMode:
                fit_mode = self.print_settings.fit_to_page
            
            if fit_mode == FitToPageMode.ACTUAL_SIZE if FitToPageMode else None:
                # Print at actual size (no scaling)
                scale = 1.0
                scaled_width = img_width
                scaled_height = img_height
            elif fit_mode == FitToPageMode.FIT_TO_PAPER if FitToPageMode else None:
                # Fit to entire paper (100% of printable area)
                scale_x = printer_width / img_width
                scale_y = printer_height / img_height
                scale = min(scale_x, scale_y)
                scaled_width = int(img_width * scale)
                scaled_height = int(img_height * scale)
            elif fit_mode == FitToPageMode.SHRINK_TO_FIT if FitToPageMode else None:
                # Only shrink if image is larger than page
                scale_x = printer_width / img_width
                scale_y = printer_height / img_height
                scale = min(scale_x, scale_y, 1.0)  # Don't enlarge
                scaled_width = int(img_width * scale)
                scaled_height = int(img_height * scale)
            else:
                # Default: FIT_TO_PAGE or NONE (90% untuk margin)
                scale_x = (printer_width * 0.9) / img_width
                scale_y = (printer_height * 0.9) / img_height
                scale = min(scale_x, scale_y)
                scaled_width = int(img_width * scale)
                scaled_height = int(img_height * scale)
            
            # Calculate centered position
            x = (printer_width - scaled_width) // 2
            y = (printer_height - scaled_height) // 2
            
            # Print image dengan error handling yang lebih baik
            try:
                dib = ImageWin.Dib(img)
                # Pastikan semua koordinat adalah integer untuk menghindari Unicode error
                dest_rect = (int(x), int(y), int(x + scaled_width), int(y + scaled_height))
                # Konversi handle ke integer juga untuk keamanan
                printer_handle = int(printer_dc.GetHandleOutput())
                dib.draw(printer_handle, dest_rect)
            except Exception as draw_error:
                # Fallback: coba dengan bitmap sederhana
                print(f"Direct draw failed, trying bitmap method: {draw_error}")
                # Simpan sebagai bitmap temporary
                import tempfile
                with tempfile.NamedTemporaryFile(suffix='.bmp', delete=False) as tmp_file:
                    temp_bmp = tmp_file.name
                    img.save(temp_bmp, 'BMP')
                
                try:
                    # Load bitmap dan print
                    bmp = win32ui.CreateBitmap()
                    bmp.LoadBitmapFile(temp_bmp)
                    
                    # Create compatible DC
                    mem_dc = printer_dc.CreateCompatibleDC()
                    mem_dc.SelectObject(bmp)
                    
                    # Stretch blit dengan koordinat integer yang aman
                    printer_dc.StretchBlt((int(x), int(y)), (int(scaled_width), int(scaled_height)), 
                                        mem_dc, (0, 0), (int(img_width), int(img_height)), win32con.SRCCOPY)
                    
                    mem_dc.DeleteDC()
                    bmp.DeleteObject()
                finally:
                    try:
                        os.unlink(temp_bmp)
                    except:
                        pass
            
            # End document
            printer_dc.EndPage()
            printer_dc.EndDoc()
            printer_dc.DeleteDC()
            
            return True, f"Image printed successfully: {os.path.basename(image_path)}"
            
        except Exception as e:
            return False, f"Image print error: {str(e)}"
    
    def print_pdf_silent(self, pdf_path, print_settings: Optional['PrintSettings'] = None):
        """Cetak PDF secara silent tanpa dialog dengan print settings"""
        if not os.path.exists(pdf_path):
            return False, f"PDF file not found: {pdf_path}"
        
        print(f"Starting silent PDF print: {pdf_path}")
        print(f"Target printer: {self.printer_name}")
        
        # Store print settings as instance attribute for access by other methods
        self.print_settings = print_settings
        
        # Check if split PDF is requested
        if print_settings and hasattr(print_settings, 'split_pdf') and print_settings.split_pdf:
            print("Split PDF mode enabled")
            split_range = getattr(print_settings, 'split_page_range', None)
            split_prefix = getattr(print_settings, 'split_output_prefix', 'page_')
            
            # Split PDF into separate files
            split_files = self.split_pdf_pages(pdf_path, split_range, split_prefix)
            if split_files:
                print(f"PDF split into {len(split_files)} files")
                
                # Print each split file separately
                total_success = 0
                for i, split_file in enumerate(split_files):
                    print(f"\nPrinting split file {i+1}/{len(split_files)}: {os.path.basename(split_file)}")
                    
                    # Create a copy of settings without split_pdf to avoid recursion
                    split_settings = None
                    if print_settings:
                        split_settings = print_settings.copy()
                        split_settings.split_pdf = False
                    
                    success, message = self.print_pdf_silent(split_file, split_settings)
                    if success:
                        total_success += 1
                        print(f"  ✓ Split file {i+1} printed successfully")
                    else:
                        print(f"  ✗ Split file {i+1} failed: {message}")
                    
                    # Cleanup split file
                    try:
                        os.remove(split_file)
                    except:
                        pass
                
                return total_success > 0, f"Printed {total_success}/{len(split_files)} split files successfully"
            else:
                print("Failed to split PDF")
                return False, "Failed to split PDF"
        
        # Apply print settings if provided
        settings_applied = False
        if print_settings:
            print(f"Applying print settings...")
            if hasattr(print_settings, 'color_mode'):
                print(f"  Color Mode: {print_settings.color_mode}")
            if hasattr(print_settings, 'orientation'):
                print(f"  Orientation: {print_settings.orientation}")
            if hasattr(print_settings, 'copies'):
                print(f"  Copies: {print_settings.copies}")
            if hasattr(print_settings, 'fit_to_page'):
                print(f"  Fit to Page: {print_settings.fit_to_page}")
            if hasattr(print_settings, 'page_range'):
                print(f"  Page Range: {print_settings.page_range}")
            
            settings_applied = self.set_print_settings(print_settings)
            if settings_applied:
                print("✅ Print settings applied successfully")
            else:
                print("⚠️ Warning: Could not apply all print settings")
        
        # Method 1: Konversi PDF ke gambar dan cetak langsung
        print("\nMethod 1: PDF to Images (Direct GDI)...")
        try:
            image_paths = self.pdf_to_images(pdf_path)
            if image_paths:
                print(f"  Converted to {len(image_paths)} images")
                
                success_count = 0
                for i, image_path in enumerate(image_paths):
                    print(f"  Printing page {i + 1}...")
                    
                    success, message = self.print_image_direct(image_path)
                    if success:
                        print(f"    ✓ Page {i + 1} printed successfully")
                        success_count += 1
                    else:
                        print(f"    ✗ Page {i + 1} failed: {message}")
                    
                    # Cleanup image
                    try:
                        os.remove(image_path)
                    except:
                        pass
                
                if success_count > 0:
                    # Restore printer settings before returning
                    if settings_applied:
                        self.restore_printer_settings()
                    return True, f"Printed {success_count}/{len(image_paths)} pages successfully"
                else:
                    print("  All pages failed to print")
            else:
                print("  Failed to convert PDF to images")
        except Exception as e:
            print(f"  Method 1 error: {e}")
        
        # Method 2: Gunakan SumatraPDF dengan parameter silent
        print("\nMethod 2: SumatraPDF Silent...")
        try:
            sumatra_paths = [
                os.path.join(self.print_tools_dir, "SumatraPDF-3.4.6-64.exe"),
                os.path.join(self.print_tools_dir, "SumatraPDF.exe"),
                "C:\\Program Files\\SumatraPDF\\SumatraPDF.exe",
                "C:\\Program Files (x86)\\SumatraPDF\\SumatraPDF.exe"
            ]
            
            print("  Checking SumatraPDF paths:")
            sumatra_exe = None
            for path in sumatra_paths:
                exists = os.path.exists(path)
                print(f"    {path} - exists: {exists}")
                if exists:
                    sumatra_exe = path
                    break
            
            if sumatra_exe:
                print(f"  Using SumatraPDF: {sumatra_exe}")
                print(f"  File exists: {os.path.exists(pdf_path)}")
                
                cmd = [sumatra_exe, "-print-to", self.printer_name, "-silent", pdf_path]
                print(f"  Running: {' '.join(cmd[:5])}")
                print(f"  Running: {' '.join(cmd)}")
                
                result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
                
                if result.returncode == 0:
                    # Restore printer settings before returning
                    if settings_applied:
                        self.restore_printer_settings()
                    return True, "PDF printed successfully with SumatraPDF"
                else:
                    print(f"  SumatraPDF failed: {result.stderr}")
            else:
                print("  SumatraPDF not found")
        except Exception as e:
            print(f"  Method 2 error: {e}")
        
        # Method 3: Gunakan PDFtoPrinter_m.exe
        print("\nMethod 3: PDFtoPrinter_m.exe...")
        try:
            # Gunakan PDFtoPrinter_m.exe yang sudah diunduh di print_tools
            current_dir = os.path.dirname(os.path.dirname(__file__))
            local_pdftoprinter = os.path.join(current_dir, "print_tools", "PDFtoPrinter_m.exe")
            
            pdftoprinter_paths = [
                local_pdftoprinter,  # Local PDFtoPrinter_m.exe di print_tools
                r"C:\Program Files\PDFtoPrinter\PDFtoPrinter.exe",
                r"C:\Program Files (x86)\PDFtoPrinter\PDFtoPrinter.exe",
                "PDFtoPrinter.exe"  # If in PATH
            ]
            
            pdftoprinter_exe = None
            print(f"  Checking PDFtoPrinter paths:")
            for path in pdftoprinter_paths:
                exists = os.path.exists(path)
                print(f"    {path} - exists: {exists}")
                if exists or path == "PDFtoPrinter.exe":
                    pdftoprinter_exe = path
                    break
            
            if pdftoprinter_exe:
                cmd = [
                    pdftoprinter_exe,
                    pdf_path,
                    self.printer_name
                ]
                
                print(f"  Using PDFtoPrinter: {pdftoprinter_exe}")
                print(f"  File exists: {os.path.exists(pdftoprinter_exe)}")
                print(f"  Running: {' '.join(cmd)}")
                result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
                
                if result.returncode == 0:
                    # Restore printer settings before returning
                    if settings_applied:
                        self.restore_printer_settings()
                    return True, "PDF printed successfully with PDFtoPrinter_m.exe"
                else:
                    print(f"  PDFtoPrinter_m.exe failed: {result.stderr}")
            else:
                print("  PDFtoPrinter_m.exe not found")
        except Exception as e:
            print(f"  Method 3 error: {e}")
        
        # Method 4: Direct Win32 API with hidden window
        print("\nMethod 4: Win32 API Direct...")
        success, message = self._print_with_win32_direct(pdf_path)
        if success:
            # Restore printer settings before returning
            if settings_applied:
                self.restore_printer_settings()
            return True, f"Method 4 success: {message}"
        print(f"  Method 4 error: {message}")
        
        # Method 5: PowerShell silent print
        print("\nMethod 5: PowerShell Silent...")
        success, message = self._print_with_powershell_silent(pdf_path)
        if success:
            # Restore printer settings before returning
            if settings_applied:
                self.restore_printer_settings()
            return True, f"Method 5 success: {message}"
        print(f"  Method 5 error: {message}")
        
        # Method 6: Advanced Win32 Raw Printing
        print("\nMethod 6: Win32 Raw Printing...")
        success, message = self._print_with_win32_raw(pdf_path)
        if success:
            # Restore printer settings before returning
            if settings_applied:
                self.restore_printer_settings()
            return True, f"Method 6 success: {message}"
        print(f"  Method 6 error: {message}")
        
        # Method 7: Network/TCP Printing (if printer supports it)
        print("\nMethod 7: Network TCP Printing...")
        try:
            success, message = self._print_with_tcp(pdf_path)
            if success:
                # Restore printer settings before returning
                if settings_applied:
                    self.restore_printer_settings()
                return True, f"Method 7 success: {message}"
            print(f"  Method 7 error: {message}")
        except AttributeError:
            print("  Method 7: TCP printing not implemented yet")
        
        # Restore printer settings before final return
        if settings_applied:
            self.restore_printer_settings()
        
        return False, "All silent print methods failed"
    
    def _print_with_win32_direct(self, pdf_path):
        """Print PDF using Win32 API with hidden window"""
        printer_handle = None
        original_devmode = None
        original_printer = None
        
        try:
            import win32api
            import win32print
            import win32con
            
            # Set printer as default temporarily
            original_printer = win32print.GetDefaultPrinter()
            win32print.SetDefaultPrinter(self.printer_name)
            
            # Apply print settings if available
            if hasattr(self, 'current_settings') and self.current_settings:
                try:
                    printer_handle = win32print.OpenPrinter(self.printer_name)
                    
                    # Get current device mode
                    devmode = win32print.GetPrinter(printer_handle, 2)['pDevMode']
                    if devmode:
                        # Store original settings
                        original_devmode = {
                            'dmColor': getattr(devmode, 'dmColor', None),
                            'dmOrientation': getattr(devmode, 'dmOrientation', None),
                            'dmCopies': getattr(devmode, 'dmCopies', None)
                        }
                        
                        # Apply color mode
                        if hasattr(self.current_settings, 'color_mode'):
                            if self.current_settings.color_mode == ColorMode.GRAYSCALE:
                                devmode.dmColor = win32con.DMCOLOR_MONOCHROME
                            elif self.current_settings.color_mode == ColorMode.COLOR:
                                devmode.dmColor = win32con.DMCOLOR_COLOR
                        
                        # Apply orientation
                        if hasattr(self.current_settings, 'orientation'):
                            if self.current_settings.orientation == Orientation.PORTRAIT:
                                devmode.dmOrientation = win32con.DMORIENT_PORTRAIT
                            elif self.current_settings.orientation == Orientation.LANDSCAPE:
                                devmode.dmOrientation = win32con.DMORIENT_LANDSCAPE
                        
                        # Apply copies
                        if hasattr(self.current_settings, 'copies') and self.current_settings.copies:
                            devmode.dmCopies = self.current_settings.copies
                        
                        # Apply the modified device mode
                        win32print.DocumentProperties(0, printer_handle, self.printer_name, devmode, devmode, win32con.DM_IN_BUFFER | win32con.DM_OUT_BUFFER)
                        win32print.SetPrinter(printer_handle, 2, {'pDevMode': devmode}, 0)
                        
                except Exception as settings_error:
                    print(f"Warning: Could not apply print settings: {settings_error}")
            
            try:
                # Use ShellExecute with hidden window
                result = win32api.ShellExecute(
                    0,                    # hwnd
                    "print",              # verb
                    pdf_path,             # file
                    None,                 # parameters
                    None,                 # directory
                    win32con.SW_HIDE     # show command (hidden)
                )
                
                if result > 32:  # Success
                    return True, "PDF sent to printer via Win32 API (hidden)"
                else:
                    return False, f"Win32 API error code: {result}"
                    
            finally:
                # Restore original printer
                if original_printer:
                    win32print.SetDefaultPrinter(original_printer)
                    
        except Exception as e:
            return False, f"Win32 direct print error: {str(e)}"
        finally:
            # Cleanup: restore original settings and close handles
            if printer_handle and original_devmode:
                try:
                    # Get current device mode to restore
                    devmode = win32print.GetPrinter(printer_handle, 2)['pDevMode']
                    if devmode:
                        # Restore original settings
                        if original_devmode['dmColor'] is not None:
                            devmode.dmColor = original_devmode['dmColor']
                        if original_devmode['dmOrientation'] is not None:
                            devmode.dmOrientation = original_devmode['dmOrientation']
                        if original_devmode['dmCopies'] is not None:
                            devmode.dmCopies = original_devmode['dmCopies']
                        
                        # Apply restored settings
                        win32print.DocumentProperties(0, printer_handle, self.printer_name, devmode, devmode, win32con.DM_IN_BUFFER | win32con.DM_OUT_BUFFER)
                        win32print.SetPrinter(printer_handle, 2, {'pDevMode': devmode}, 0)
                except Exception as cleanup_error:
                    print(f"Warning: Could not restore original printer settings: {cleanup_error}")
            
            if printer_handle:
                try:
                    win32print.ClosePrinter(printer_handle)
                except Exception as close_error:
                    print(f"Warning: Could not close printer handle: {close_error}")
    
    def _print_with_powershell_silent(self, pdf_path):
        """Print PDF using PowerShell with silent execution"""
        try:
            import subprocess
            import os
            
            # PowerShell script untuk print silent
            ps_script = f'''
            $printer = "{self.printer_name}"
            $file = "{pdf_path}"
            
            # Coba print dengan Start-Process
            try {{
                Start-Process -FilePath $file -Verb Print -WindowStyle Hidden -Wait
                Write-Output "SUCCESS: PDF printed silently"
            }} catch {{
                Write-Output "ERROR: $($_.Exception.Message)"
            }}
            '''
            
            # Jalankan PowerShell dengan window tersembunyi
            result = subprocess.run([
                "powershell.exe",
                "-WindowStyle", "Hidden",
                "-ExecutionPolicy", "Bypass",
                "-Command", ps_script
            ], capture_output=True, text=True, timeout=30)
            
            if result.returncode == 0 and "SUCCESS" in result.stdout:
                return True, "PDF printed via PowerShell (silent)"
            else:
                return False, f"PowerShell error: {result.stderr or result.stdout}"
                
        except subprocess.TimeoutExpired:
            return False, "PowerShell print timeout"
        except Exception as e:
            return False, f"PowerShell print error: {str(e)}"
    
    def _print_with_win32_raw(self, pdf_path):
        """Advanced Win32 Raw Printing with ESC/POS commands and direct driver communication"""
        try:
            # First convert PDF to images for raw printing
            image_paths = self.pdf_to_images(pdf_path, dpi=200)  # Higher DPI for better quality
            if not image_paths:
                return False, "Failed to convert PDF to images for raw printing"
            
            print(f"  Converting {len(image_paths)} pages for raw printing...")
            
            # Open printer for raw data access
            if not self.open_printer():
                return False, "Cannot open printer for raw access"
            
            try:
                # Initialize printer with ESC/POS commands
                success, msg = self._send_raw_printer_commands()
                if not success:
                    return False, f"Failed to initialize printer: {msg}"
                
                # Process each page
                for i, image_path in enumerate(image_paths):
                    print(f"    Processing page {i + 1} for raw printing...")
                    
                    # Convert image to printer-compatible format
                    success, msg = self._print_image_raw(image_path, i + 1)
                    if not success:
                        print(f"    Warning: Page {i + 1} failed: {msg}")
                    else:
                        print(f"    Page {i + 1} sent to printer successfully")
                    
                    # Cleanup
                    try:
                        os.remove(image_path)
                    except:
                        pass
                
                # Send final commands
                self._send_raw_finish_commands()
                
                return True, f"PDF printed via Win32 Raw Printing ({len(image_paths)} pages)"
                
            finally:
                self.close_printer()
                
        except Exception as e:
            return False, f"Win32 Raw Printing error: {str(e)}"
    
    def _send_raw_printer_commands(self):
        """Send ESC/POS initialization commands to printer"""
        try:
            # ESC/POS commands for EPSON printers
            init_commands = (
                b'\x1B@'      # ESC @ - Initialize printer
                b'\x1B\x33\x00'  # ESC 3 - Set line spacing to 0
                b'\x1Bt\x00'     # ESC t - Select character code table
                b'\x1Ba\x00'     # ESC a - Left align
            )
            
            # Start raw document
            job_info = ("Win32_Raw_Print_Job", None, "RAW")
            job_id = win32print.StartDocPrinter(self.printer_handle, 1, job_info)
            
            if job_id <= 0:
                return False, "Failed to start raw print job"
            
            # Start page
            win32print.StartPagePrinter(self.printer_handle)
            
            # Send initialization commands
            bytes_written = win32print.WritePrinter(self.printer_handle, init_commands)
            
            if bytes_written > 0:
                return True, f"Printer initialized, Job ID: {job_id}"
            else:
                return False, "Failed to send initialization commands"
                
        except Exception as e:
            return False, f"Raw command error: {str(e)}"
    
    def _print_image_raw(self, image_path, page_num):
        """Convert image to raw printer data and send"""
        try:
            from PIL import Image
            import struct
            
            # Load and process image
            img = Image.open(image_path)
            
            # Convert to grayscale for better printer compatibility
            if img.mode != 'L':
                img = img.convert('L')
            
            # Resize to fit printer width (assuming 8 inches at 180 DPI)
            max_width = 576  # 8 inches * 72 DPI for EPSON L120
            if img.width > max_width:
                ratio = max_width / img.width
                new_height = int(img.height * ratio)
                img = img.resize((max_width, new_height), Image.Resampling.LANCZOS)
            
            # Convert to monochrome bitmap for raw printing
            img = img.convert('1')  # 1-bit monochrome
            
            # Get image data
            width, height = img.size
            img_data = list(img.getdata())
            
            # Convert to ESC/POS bitmap format
            # ESC * command for bit image printing
            raw_data = b''
            
            # Process image line by line
            bytes_per_line = (width + 7) // 8  # Round up to nearest byte
            
            for y in range(height):
                # Start line with ESC * command (single density)
                line_header = b'\x1B*\x00' + struct.pack('<H', bytes_per_line)
                raw_data += line_header
                
                # Convert pixels to bytes
                line_bytes = []
                for x in range(0, width, 8):
                    byte_val = 0
                    for bit in range(8):
                        if x + bit < width:
                            pixel_idx = y * width + x + bit
                            if pixel_idx < len(img_data) and img_data[pixel_idx] == 0:  # Black pixel
                                byte_val |= (1 << (7 - bit))
                    line_bytes.append(byte_val)
                
                # Pad to required length
                while len(line_bytes) < bytes_per_line:
                    line_bytes.append(0)
                
                raw_data += bytes(line_bytes)
                raw_data += b'\r\n'  # Carriage return + line feed
            
            # Send the raw image data
            bytes_written = win32print.WritePrinter(self.printer_handle, raw_data)
            
            if bytes_written > 0:
                return True, f"Page {page_num} raw data sent ({bytes_written} bytes)"
            else:
                return False, f"Failed to send raw data for page {page_num}"
                
        except Exception as e:
            return False, f"Raw image processing error: {str(e)}"
    
    def _send_raw_finish_commands(self):
        """Send finishing commands to printer"""
        try:
            # ESC/POS finishing commands
            finish_commands = (
                b'\x0C'      # Form feed
                b'\x1B@'     # ESC @ - Reset printer
            )
            
            # Send finish commands
            win32print.WritePrinter(self.printer_handle, finish_commands)
            
            # End page and document
            win32print.EndPagePrinter(self.printer_handle)
            win32print.EndDocPrinter(self.printer_handle)
            
        except Exception as e:
            print(f"Warning: Failed to send finish commands: {e}")
    
    def cleanup(self):
        """Bersihkan resources"""
        try:
            self.close_printer()
            
            # Hapus temp directory
            import shutil
            if os.path.exists(self.temp_dir):
                shutil.rmtree(self.temp_dir)
        except Exception as e:
            print(f"Cleanup error: {e}")
    
    def __del__(self):
        """Destructor untuk cleanup otomatis"""
        self.cleanup()

def test_silent_print():
    """Test silent print service"""
    print("=== TESTING SILENT PRINT SERVICE ===")
    
    # Test file
    test_pdf = "D:/Gawean Rebinmas/Driver_Epson_L120/temp/Test_print.pdf"
    
    if not os.path.exists(test_pdf):
        print(f"❌ Test file not found: {test_pdf}")
        return False
    
    # Initialize service
    service = SilentPrintService()
    
    # Find printer
    printer = service.find_printer()
    if not printer:
        print("❌ No printer found")
        return False
    
    print(f"✅ Using printer: {printer}")
    
    # Test silent print
    success, message = service.print_pdf_silent(test_pdf)
    
    print(f"\nResult: {'✅ SUCCESS' if success else '❌ FAILED'}")
    print(f"Message: {message}")
    
    # Cleanup
    service.cleanup()
    
    return success

if __name__ == "__main__":
    try:
        success = test_silent_print()
        sys.exit(0 if success else 1)
    except KeyboardInterrupt:
        print("\n\n⚠️  Test interrupted by user")
        sys.exit(1)
    except Exception as e:
        print(f"\n\n❌ Unexpected error: {e}")
        sys.exit(1)
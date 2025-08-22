import os
import logging
from typing import Dict, Any, Optional, Tuple, List
from pathlib import Path
from datetime import datetime
import tempfile
import shutil

# PDF processing libraries
import PyPDF2
import fitz  # PyMuPDF
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, A4, legal, A3, A5
from reportlab.lib.units import inch, cm, mm
from reportlab.lib.utils import ImageReader
from reportlab.platypus import SimpleDocTemplate, Spacer
from reportlab.lib import colors
from PIL import Image
import io

logger = logging.getLogger(__name__)


class DocumentService:
    """Service untuk memproses dan memodifikasi dokumen PDF berdasarkan pengaturan cetak"""
    
    def __init__(self, temp_dir: str = "temp"):
        self.temp_dir = Path(temp_dir)
        self.temp_dir.mkdir(exist_ok=True)
        
        # Supported paper sizes
        self.paper_sizes = {
            'A4': A4,
            'A3': A3, 
            'A5': A5,
            'Letter': letter,
            'Legal': legal,
            'Custom': None  # Will be calculated based on width/height
        }
        
        # Default settings
        self.default_settings = {
            'paper_size': 'A4',
            'orientation': 'portrait',  # portrait, landscape
            'margin_top': 20,  # mm
            'margin_bottom': 20,  # mm
            'margin_left': 20,  # mm
            'margin_right': 20,  # mm
            'scale': 100,  # percentage
            'fit_to_page': False,
            'center_horizontally': True,
            'center_vertically': True,
            'print_quality': 'high',  # low, medium, high
            'color_mode': 'color',  # color, grayscale, black_white
            'duplex': False,  # single-sided, double-sided
            'copies': 1
        }
    
    def process_document(self, input_pdf_path: str, settings: Dict[str, Any]) -> Dict[str, Any]:
        """Proses dokumen PDF dengan pengaturan yang diberikan"""
        try:
            # Validasi input file
            if not os.path.exists(input_pdf_path):
                raise FileNotFoundError(f"Input PDF not found: {input_pdf_path}")
            
            # Merge settings dengan default
            final_settings = {**self.default_settings, **settings}
            
            # Generate output filename
            input_path = Path(input_pdf_path)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f"{input_path.stem}_processed_{timestamp}.pdf"
            output_path = self.temp_dir / output_filename
            
            logger.info(f"Processing document: {input_pdf_path} -> {output_path}")
            logger.info(f"Settings: {final_settings}")
            
            # Proses PDF berdasarkan pengaturan
            result = self._process_pdf_with_settings(input_pdf_path, str(output_path), final_settings)
            
            # Return hasil processing
            return {
                'success': True,
                'input_path': input_pdf_path,
                'output_path': str(output_path),
                'output_filename': output_filename,
                'settings_applied': final_settings,
                'file_size': os.path.getsize(output_path),
                'pages_processed': result.get('pages_processed', 0),
                'processing_time': result.get('processing_time', 0),
                'created_at': datetime.now().isoformat()
            }
            
        except Exception as e:
            logger.error(f"Error processing document {input_pdf_path}: {e}")
            return {
                'success': False,
                'error': str(e),
                'input_path': input_pdf_path
            }
    
    def _process_pdf_with_settings(self, input_path: str, output_path: str, settings: Dict[str, Any]) -> Dict[str, Any]:
        """Proses PDF dengan pengaturan spesifik"""
        start_time = datetime.now()
        
        try:
            # Buka PDF input
            pdf_document = fitz.open(input_path)
            pages_processed = 0
            
            # Dapatkan ukuran kertas target
            target_page_size = self._get_target_page_size(settings)
            
            # Tentukan halaman yang akan diproses berdasarkan page_range
            pages_to_process = self._get_pages_to_process(pdf_document, settings)
            
            # Buat PDF output baru
            output_pdf = fitz.open()  # Buat dokumen kosong
            
            # Proses halaman yang dipilih
            for page_num in pages_to_process:
                if page_num < len(pdf_document):
                    page = pdf_document[page_num]
                    
                    # Proses halaman berdasarkan pengaturan
                    processed_page = self._process_page(page, settings, target_page_size)
                    
                    # Tambahkan halaman yang sudah diproses ke output
                    output_pdf.insert_pdf(processed_page)
                    pages_processed += 1
                    
                    processed_page.close()
            
            # Simpan PDF output
            output_pdf.save(output_path)
            output_pdf.close()
            pdf_document.close()
            
            processing_time = (datetime.now() - start_time).total_seconds()
            
            logger.info(f"Document processed successfully: {pages_processed} pages in {processing_time:.2f}s")
            
            return {
                'pages_processed': pages_processed,
                'processing_time': processing_time
            }
            
        except Exception as e:
            logger.error(f"Error in PDF processing: {e}")
            raise
    
    def _process_page(self, page: fitz.Page, settings: Dict[str, Any], target_size: Tuple[float, float]) -> fitz.Document:
        """Proses satu halaman PDF"""
        try:
            # Buat dokumen baru untuk halaman ini
            new_doc = fitz.open()
            
            # Dapatkan ukuran halaman asli
            original_rect = page.rect
            original_width = original_rect.width
            original_height = original_rect.height
            
            # Hitung transformasi berdasarkan pengaturan
            transform_matrix = self._calculate_transform_matrix(
                original_width, original_height, target_size, settings
            )
            
            # Buat halaman baru dengan ukuran target
            new_page = new_doc.new_page(width=target_size[0], height=target_size[1])
            
            # Render halaman asli sebagai gambar dengan transformasi
            if settings.get('print_quality') == 'high':
                zoom = 2.0  # High quality
            elif settings.get('print_quality') == 'medium':
                zoom = 1.5  # Medium quality
            else:
                zoom = 1.0  # Low quality
            
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat)
            
            # Konversi ke format yang sesuai berdasarkan color_mode
            color_mode = settings.get('color_mode')
            if color_mode == 'grayscale' or (hasattr(color_mode, 'value') and color_mode.value == 'GRAYSCALE'):
                pix = fitz.Pixmap(fitz.csGRAY, pix)
            elif color_mode == 'black_white' or (hasattr(color_mode, 'value') and color_mode.value == 'BLACK_WHITE'):
                pix = fitz.Pixmap(fitz.csGRAY, pix)
                # Tambahan processing untuk black & white bisa ditambahkan di sini
            
            # Hitung posisi untuk centering jika diperlukan
            insert_rect = self._calculate_insert_position(
                pix.width / zoom, pix.height / zoom, target_size, settings
            )
            
            # Insert gambar ke halaman baru
            new_page.insert_image(insert_rect, pixmap=pix)
            
            return new_doc
            
        except Exception as e:
            logger.error(f"Error processing page: {e}")
            raise
    
    def _get_target_page_size(self, settings: Dict[str, Any]) -> Tuple[float, float]:
        """Dapatkan ukuran halaman target berdasarkan pengaturan"""
        paper_size = settings.get('paper_size', 'A4')
        orientation = settings.get('orientation', 'portrait')
        
        if paper_size == 'Custom':
            width = settings.get('custom_width', 210) * mm  # Default A4 width
            height = settings.get('custom_height', 297) * mm  # Default A4 height
        else:
            size = self.paper_sizes.get(paper_size, A4)
            width, height = size
        
        # Tukar width dan height untuk landscape
        if orientation == 'landscape':
            width, height = height, width
        
        return (width, height)
    
    def _calculate_transform_matrix(self, orig_width: float, orig_height: float, 
                                  target_size: Tuple[float, float], settings: Dict[str, Any]) -> fitz.Matrix:
        """Hitung matrix transformasi untuk scaling dan positioning"""
        target_width, target_height = target_size
        
        # Hitung margin dalam points
        margin_left = settings.get('margin_left', 20) * mm / 72 * 72  # Convert mm to points
        margin_right = settings.get('margin_right', 20) * mm / 72 * 72
        margin_top = settings.get('margin_top', 20) * mm / 72 * 72
        margin_bottom = settings.get('margin_bottom', 20) * mm / 72 * 72
        
        # Area yang tersedia untuk konten
        available_width = target_width - margin_left - margin_right
        available_height = target_height - margin_top - margin_bottom
        
        # Hitung scale factor
        scale_percent = settings.get('scale', 100) / 100.0
        
        if settings.get('fit_to_page', False):
            # Fit to page: hitung scale untuk fit dalam area yang tersedia
            scale_x = available_width / orig_width
            scale_y = available_height / orig_height
            scale_factor = min(scale_x, scale_y) * scale_percent
        else:
            # Manual scale
            scale_factor = scale_percent
        
        return fitz.Matrix(scale_factor, scale_factor)
    
    def _calculate_insert_position(self, content_width: float, content_height: float,
                                 target_size: Tuple[float, float], settings: Dict[str, Any]) -> fitz.Rect:
        """Hitung posisi insert untuk centering"""
        target_width, target_height = target_size
        
        # Margin dalam points
        margin_left = settings.get('margin_left', 20) * mm / 72 * 72
        margin_top = settings.get('margin_top', 20) * mm / 72 * 72
        
        # Posisi default (top-left dengan margin)
        x = margin_left
        y = margin_top
        
        # Center horizontally jika diatur
        if settings.get('center_horizontally', True):
            available_width = target_width - margin_left - settings.get('margin_right', 20) * mm / 72 * 72
            x = margin_left + (available_width - content_width) / 2
        
        # Center vertically jika diatur
        if settings.get('center_vertically', True):
            available_height = target_height - margin_top - settings.get('margin_bottom', 20) * mm / 72 * 72
            y = margin_top + (available_height - content_height) / 2
        
        return fitz.Rect(x, y, x + content_width, y + content_height)
    
    def get_document_info(self, pdf_path: str) -> Dict[str, Any]:
        """Dapatkan informasi dokumen PDF"""
        try:
            pdf_document = fitz.open(pdf_path)
            
            info = {
                'pages': len(pdf_document),
                'title': pdf_document.metadata.get('title', ''),
                'author': pdf_document.metadata.get('author', ''),
                'subject': pdf_document.metadata.get('subject', ''),
                'creator': pdf_document.metadata.get('creator', ''),
                'producer': pdf_document.metadata.get('producer', ''),
                'creation_date': pdf_document.metadata.get('creationDate', ''),
                'modification_date': pdf_document.metadata.get('modDate', ''),
                'file_size': os.path.getsize(pdf_path)
            }
            
            # Informasi halaman pertama
            if len(pdf_document) > 0:
                first_page = pdf_document[0]
                rect = first_page.rect
                info['page_width'] = rect.width
                info['page_height'] = rect.height
                info['page_size_mm'] = {
                    'width': rect.width * 25.4 / 72,  # Convert points to mm
                    'height': rect.height * 25.4 / 72
                }
            
            pdf_document.close()
            return info
            
        except Exception as e:
            logger.error(f"Error getting document info for {pdf_path}: {e}")
            raise
    
    def cleanup_temp_files(self, older_than_hours: int = 24) -> int:
        """Bersihkan file temporary yang lama"""
        try:
            count = 0
            cutoff_time = datetime.now().timestamp() - (older_than_hours * 3600)
            
            for file_path in self.temp_dir.rglob('*_processed_*.pdf'):
                if file_path.is_file() and file_path.stat().st_mtime < cutoff_time:
                    file_path.unlink()
                    count += 1
                    logger.debug(f"Deleted temp file: {file_path}")
            
            logger.info(f"Cleaned up {count} temporary processed files")
            return count
            
        except Exception as e:
            logger.error(f"Error cleaning up temp files: {e}")
            return 0
    
    def validate_settings(self, settings: Dict[str, Any]) -> Dict[str, Any]:
        """Validasi dan normalisasi pengaturan"""
        validated = {}
        errors = []
        
        # Validasi paper_size
        paper_size = settings.get('paper_size', 'A4')
        if paper_size not in self.paper_sizes:
            errors.append(f"Invalid paper size: {paper_size}")
        else:
            validated['paper_size'] = paper_size
        
        # Validasi orientation
        orientation = settings.get('orientation', 'portrait')
        if orientation not in ['portrait', 'landscape']:
            errors.append(f"Invalid orientation: {orientation}")
        else:
            validated['orientation'] = orientation
        
        # Validasi margins (harus positif)
        for margin in ['margin_top', 'margin_bottom', 'margin_left', 'margin_right']:
            value = settings.get(margin, 20)
            try:
                value = float(value)
                if value < 0:
                    errors.append(f"Margin {margin} must be positive")
                else:
                    validated[margin] = value
            except (ValueError, TypeError):
                errors.append(f"Invalid margin value for {margin}: {value}")
        
        # Validasi scale (1-500%)
        scale = settings.get('scale', 100)
        try:
            scale = float(scale)
            if not (1 <= scale <= 500):
                errors.append("Scale must be between 1% and 500%")
            else:
                validated['scale'] = scale
        except (ValueError, TypeError):
            errors.append(f"Invalid scale value: {scale}")
        
        # Validasi boolean settings
        for bool_setting in ['fit_to_page', 'center_horizontally', 'center_vertically', 'duplex']:
            value = settings.get(bool_setting, False)
            validated[bool_setting] = bool(value)
        
        # Validasi color_mode
        color_mode = settings.get('color_mode', 'color')
        valid_color_modes = ['color', 'grayscale', 'black_white']
        
        # Support for enum ColorMode
        if hasattr(color_mode, 'value'):
            color_mode_value = color_mode.value.lower()
            if color_mode_value in valid_color_modes:
                validated['color_mode'] = color_mode
            else:
                errors.append(f"Invalid color mode enum: {color_mode}")
        elif color_mode in valid_color_modes:
            validated['color_mode'] = color_mode
        else:
            errors.append(f"Invalid color mode: {color_mode}")
        
        # Validasi print_quality
        print_quality = settings.get('print_quality', 'high')
        if print_quality not in ['low', 'medium', 'high']:
            errors.append(f"Invalid print quality: {print_quality}")
        else:
            validated['print_quality'] = print_quality
        
        # Validasi copies
        copies = settings.get('copies', 1)
        try:
            copies = int(copies)
            if copies < 1:
                errors.append("Copies must be at least 1")
            else:
                validated['copies'] = copies
        except (ValueError, TypeError):
            errors.append(f"Invalid copies value: {copies}")
        
        return {
            'valid': len(errors) == 0,
            'errors': errors,
            'validated_settings': validated
        }
    
    def _get_pages_to_process(self, pdf_document: fitz.Document, settings: Dict[str, Any]) -> List[int]:
        """Tentukan halaman yang akan diproses berdasarkan page_range_type dan page_range"""
        total_pages = len(pdf_document)
        page_range_type = settings.get('page_range_type', 'all')
        page_range = settings.get('page_range', '')
        
        if page_range_type == 'all':
            return list(range(total_pages))
        elif page_range_type == 'odd':
            return [i for i in range(total_pages) if (i + 1) % 2 == 1]  # 1, 3, 5, ...
        elif page_range_type == 'even':
            return [i for i in range(total_pages) if (i + 1) % 2 == 0]  # 2, 4, 6, ...
        elif page_range_type == 'range' and page_range:
            return self._parse_page_range(page_range, total_pages)
        else:
            # Default ke semua halaman jika tidak valid
            return list(range(total_pages))
    
    def _parse_page_range(self, page_range: str, total_pages: int) -> List[int]:
        """Parse page range string seperti '1-5,8,11-13' menjadi list page numbers (0-indexed)"""
        pages = []
        
        try:
            # Split by comma untuk mendapatkan range atau single pages
            parts = [part.strip() for part in page_range.split(',')]
            
            for part in parts:
                if '-' in part:
                    # Range seperti '1-5'
                    start, end = part.split('-', 1)
                    start_page = max(1, min(int(start.strip()), total_pages))
                    end_page = max(1, min(int(end.strip()), total_pages))
                    
                    # Tambahkan range (convert to 0-indexed)
                    for page_num in range(start_page - 1, end_page):
                        if page_num not in pages:
                            pages.append(page_num)
                else:
                    # Single page seperti '8'
                    page_num = max(1, min(int(part.strip()), total_pages))
                    if (page_num - 1) not in pages:  # Convert to 0-indexed
                        pages.append(page_num - 1)
            
            # Sort pages
            pages.sort()
            return pages
            
        except (ValueError, IndexError) as e:
            logger.warning(f"Invalid page range format '{page_range}': {e}. Using all pages.")
            return list(range(total_pages))

    def __del__(self):
        """Cleanup temporary files on destruction"""
        try:
            self.cleanup_temp_files(max_age_hours=0)
        except Exception:
            pass  # Ignore cleanup errors during destruction
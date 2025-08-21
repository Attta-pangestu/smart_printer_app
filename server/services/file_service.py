import os
import shutil
import mimetypes
from typing import Dict, Any, Optional, List
from pathlib import Path
import hashlib
from datetime import datetime
import logging
from PIL import Image
import PyPDF2

logger = logging.getLogger(__name__)


class FileService:
    """Service untuk mengelola file uploads dan processing"""
    
    def __init__(self, upload_dir: str = "uploads", temp_dir: str = "temp"):
        self.upload_dir = Path(upload_dir)
        self.temp_dir = Path(temp_dir)
        
        # Buat direktori jika belum ada
        self.upload_dir.mkdir(exist_ok=True)
        self.temp_dir.mkdir(exist_ok=True)
        
        # Supported file types
        self.supported_types = {
            'pdf': ['.pdf'],
            'text': ['.txt', '.text', '.log'],
            'document': ['.doc', '.docx', '.rtf', '.odt'],
            'image': ['.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.tif'],
            'spreadsheet': ['.xls', '.xlsx', '.ods', '.csv'],
            'presentation': ['.ppt', '.pptx', '.odp']
        }
        
        # Max file size (50MB default)
        self.max_file_size = 50 * 1024 * 1024
    
    def save_uploaded_file(self, file_content: bytes, filename: str, user: str = "anonymous") -> Dict[str, Any]:
        """Simpan file yang diupload"""
        try:
            # Validasi ukuran file
            if len(file_content) > self.max_file_size:
                raise ValueError(f"File size exceeds maximum allowed size ({self.max_file_size} bytes)")
            
            # Validasi tipe file
            file_type = self.get_file_type(filename)
            if not self.is_supported_file(filename):
                raise ValueError(f"File type not supported: {file_type}")
            
            # Generate unique filename
            safe_filename = self._generate_safe_filename(filename)
            file_path = self.upload_dir / safe_filename
            
            # Simpan file
            with open(file_path, 'wb') as f:
                f.write(file_content)
            
            # Generate file info
            file_info = self.get_file_info(str(file_path))
            file_info.update({
                'upload_path': str(file_path),
                'uploaded_by': user,
                'uploaded_at': datetime.now().isoformat()
            })
            
            logger.info(f"File saved: {safe_filename} ({len(file_content)} bytes)")
            return file_info
            
        except Exception as e:
            logger.error(f"Error saving file {filename}: {e}")
            raise
    
    def get_file_info(self, file_path: str) -> Dict[str, Any]:
        """Dapatkan informasi file"""
        try:
            path = Path(file_path)
            
            if not path.exists():
                raise FileNotFoundError(f"File not found: {file_path}")
            
            stat = path.stat()
            file_type = self.get_file_type(path.name)
            
            info = {
                'name': path.name,
                'path': str(path.absolute()),
                'size': stat.st_size,
                'type': file_type,
                'mime_type': mimetypes.guess_type(str(path))[0],
                'created_at': datetime.fromtimestamp(stat.st_ctime).isoformat(),
                'modified_at': datetime.fromtimestamp(stat.st_mtime).isoformat(),
                'hash': self._calculate_file_hash(file_path)
            }
            
            # Tambahan info berdasarkan tipe file
            if file_type == 'pdf':
                info['pages'] = self._get_pdf_pages(file_path)
            elif file_type == 'image':
                info.update(self._get_image_info(file_path))
            elif file_type == 'text':
                info['lines'] = self._get_text_lines(file_path)
            
            return info
            
        except Exception as e:
            logger.error(f"Error getting file info for {file_path}: {e}")
            raise
    
    def get_file_type(self, filename: str) -> str:
        """Dapatkan tipe file berdasarkan extension"""
        ext = Path(filename).suffix.lower()
        
        for file_type, extensions in self.supported_types.items():
            if ext in extensions:
                return file_type
        
        return 'unknown'
    
    def is_supported_file(self, filename: str) -> bool:
        """Cek apakah file type didukung"""
        return self.get_file_type(filename) != 'unknown'
    
    def delete_file(self, file_path: str) -> bool:
        """Hapus file"""
        try:
            path = Path(file_path)
            if path.exists():
                path.unlink()
                logger.info(f"File deleted: {file_path}")
                return True
            return False
            
        except Exception as e:
            logger.error(f"Error deleting file {file_path}: {e}")
            return False
    
    def cleanup_old_files(self, days: int = 7) -> int:
        """Hapus file lama"""
        try:
            count = 0
            cutoff_time = datetime.now().timestamp() - (days * 24 * 60 * 60)
            
            for file_path in self.upload_dir.rglob('*'):
                if file_path.is_file() and file_path.stat().st_mtime < cutoff_time:
                    file_path.unlink()
                    count += 1
            
            logger.info(f"Cleaned up {count} old files")
            return count
            
        except Exception as e:
            logger.error(f"Error cleaning up old files: {e}")
            return 0
    
    def get_upload_stats(self) -> Dict[str, Any]:
        """Dapatkan statistik upload"""
        try:
            total_files = 0
            total_size = 0
            file_types = {}
            
            for file_path in self.upload_dir.rglob('*'):
                if file_path.is_file():
                    total_files += 1
                    total_size += file_path.stat().st_size
                    
                    file_type = self.get_file_type(file_path.name)
                    file_types[file_type] = file_types.get(file_type, 0) + 1
            
            return {
                'total_files': total_files,
                'total_size': total_size,
                'total_size_mb': round(total_size / (1024 * 1024), 2),
                'file_types': file_types,
                'upload_dir': str(self.upload_dir.absolute())
            }
            
        except Exception as e:
            logger.error(f"Error getting upload stats: {e}")
            return {}
    
    def _generate_safe_filename(self, filename: str) -> str:
        """Generate filename yang aman"""
        # Ambil nama dan extension
        path = Path(filename)
        name = path.stem
        ext = path.suffix
        
        # Bersihkan nama file
        safe_name = "".join(c for c in name if c.isalnum() or c in (' ', '-', '_')).rstrip()
        safe_name = safe_name.replace(' ', '_')
        
        # Tambahkan timestamp untuk uniqueness
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        return f"{timestamp}_{safe_name}{ext}"
    
    def _calculate_file_hash(self, file_path: str) -> str:
        """Hitung hash file untuk integrity check"""
        try:
            hash_md5 = hashlib.md5()
            with open(file_path, "rb") as f:
                for chunk in iter(lambda: f.read(4096), b""):
                    hash_md5.update(chunk)
            return hash_md5.hexdigest()
            
        except Exception as e:
            logger.error(f"Error calculating hash for {file_path}: {e}")
            return ""
    
    def _get_pdf_pages(self, file_path: str) -> int:
        """Dapatkan jumlah halaman PDF"""
        try:
            with open(file_path, 'rb') as f:
                pdf_reader = PyPDF2.PdfReader(f)
                return len(pdf_reader.pages)
                
        except Exception as e:
            logger.error(f"Error getting PDF pages for {file_path}: {e}")
            return 0
    
    def _get_image_info(self, file_path: str) -> Dict[str, Any]:
        """Dapatkan info image"""
        try:
            with Image.open(file_path) as img:
                return {
                    'width': img.width,
                    'height': img.height,
                    'format': img.format,
                    'mode': img.mode
                }
                
        except Exception as e:
            logger.error(f"Error getting image info for {file_path}: {e}")
            return {}
    
    def _get_text_lines(self, file_path: str) -> int:
        """Dapatkan jumlah baris text file"""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                return sum(1 for line in f)
                
        except Exception as e:
            logger.error(f"Error getting text lines for {file_path}: {e}")
            return 0
    
    def create_temp_file(self, content: bytes, extension: str = ".tmp") -> str:
        """Buat temporary file"""
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
            temp_filename = f"temp_{timestamp}{extension}"
            temp_path = self.temp_dir / temp_filename
            
            with open(temp_path, 'wb') as f:
                f.write(content)
            
            return str(temp_path)
            
        except Exception as e:
            logger.error(f"Error creating temp file: {e}")
            raise
    
    def cleanup_temp_files(self) -> int:
        """Hapus semua temporary files"""
        try:
            count = 0
            for file_path in self.temp_dir.rglob('*'):
                if file_path.is_file():
                    file_path.unlink()
                    count += 1
            
            logger.info(f"Cleaned up {count} temp files")
            return count
            
        except Exception as e:
            logger.error(f"Error cleaning up temp files: {e}")
            return 0
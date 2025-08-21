import os
import uuid
import asyncio
import win32print
import win32api
from typing import List, Optional, Dict, Any
from datetime import datetime, timedelta
import logging
from pathlib import Path
import threading
from queue import Queue, Empty

from models.job import PrintJob, JobStatus, PrintSettings
from models.printer import Printer, PrinterStatus
from services.printer_service import PrinterService
from services.file_service import FileService

logger = logging.getLogger(__name__)


class JobService:
    """Service untuk mengelola print jobs"""
    
    def __init__(self, printer_service: PrinterService, file_service: FileService):
        self.printer_service = printer_service
        self.file_service = file_service
        
        # Job storage
        self._jobs: Dict[str, PrintJob] = {}
        self._job_queue: Queue = Queue()
        
        # Worker thread untuk processing jobs
        self._worker_thread = None
        self._stop_worker = False
        
        # Statistics
        self._total_jobs = 0
        self._completed_jobs = 0
        self._failed_jobs = 0
        
        self._start_worker()
    
    def submit_job(self, 
                   printer_id: str, 
                   file_path: str, 
                   settings: PrintSettings,
                   title: str = None,
                   user: str = "anonymous",
                   client_ip: str = None) -> PrintJob:
        """Submit print job baru"""
        
        logger.info(f"JobService: Submitting job for printer_id: '{printer_id}'")
        
        # Validasi printer
        printer = self.printer_service.get_printer(printer_id)
        if not printer:
            logger.error(f"JobService: Printer '{printer_id}' not found")
            raise ValueError(f"Printer {printer_id} not found")
        
        logger.info(f"JobService: Found printer: '{printer.name}' (ID: '{printer.id}')")
        
        # Validasi file
        if not os.path.exists(file_path):
            raise ValueError(f"File {file_path} not found")
        
        file_info = self.file_service.get_file_info(file_path)
        
        # Buat job baru
        job = PrintJob(
            printer_id=printer_id,
            file_name=file_info['name'],
            file_path=file_path,
            file_size=file_info['size'],
            file_type=file_info['type'],
            title=title or file_info['name'],
            user=user,
            client_ip=client_ip,
            settings=settings,
            total_pages=file_info.get('pages', 0)
        )
        
        # Simpan job
        self._jobs[job.id] = job
        self._total_jobs += 1
        
        # Masukkan ke queue
        self._job_queue.put(job.id)
        
        logger.info(f"Job {job.id} submitted for printer {printer_id}")
        return job
    
    def get_job(self, job_id: str) -> Optional[PrintJob]:
        """Dapatkan job berdasarkan ID"""
        return self._jobs.get(job_id)
    
    def get_all_jobs(self, 
                     status_filter: Optional[JobStatus] = None,
                     printer_filter: Optional[str] = None,
                     user_filter: Optional[str] = None,
                     limit: Optional[int] = None) -> List[PrintJob]:
        """Dapatkan semua jobs dengan filter"""
        jobs = list(self._jobs.values())
        
        # Apply filters
        if status_filter:
            jobs = [job for job in jobs if job.status == status_filter]
        
        if printer_filter:
            jobs = [job for job in jobs if job.printer_id == printer_filter]
        
        if user_filter:
            jobs = [job for job in jobs if job.user == user_filter]
        
        # Sort by created_at descending
        jobs.sort(key=lambda x: x.created_at, reverse=True)
        
        # Apply limit
        if limit:
            jobs = jobs[:limit]
        
        return jobs
    
    def get_jobs(self, status_filter: Optional[JobStatus] = None, 
                 printer_filter: Optional[str] = None,
                 user_filter: Optional[str] = None,
                 limit: Optional[int] = None) -> List[PrintJob]:
        """Alias untuk get_all_jobs untuk kompatibilitas"""
        return self.get_all_jobs(status_filter, printer_filter, user_filter, limit)
    
    def submit_test_job(self, 
                       printer_id: Optional[str] = None,
                       user: str = "test_user",
                       client_ip: str = None) -> PrintJob:
        """Submit test print job langsung ke printer default/aktif
        
        Args:
            printer_id: ID printer tujuan (optional, akan menggunakan default jika None)
            user: Nama user yang mengirim test print
            client_ip: IP address client
            
        Returns:
            PrintJob: Job test print yang telah dibuat
            
        Raises:
            ValueError: Jika tidak ada printer yang tersedia atau printer tidak online
        """
        # Tentukan printer yang akan digunakan
        target_printer = None
        
        if printer_id:
            # Gunakan printer yang ditentukan
            target_printer = self.printer_service.get_printer(printer_id)
            if not target_printer:
                raise ValueError(f"Printer {printer_id} tidak ditemukan")
        else:
            # Cari printer default atau yang online
            printers = self.printer_service.get_printers()
            
            # Coba cari printer default terlebih dahulu
            for printer in printers:
                if printer.is_default and printer.status == "online":
                    target_printer = printer
                    break
            
            # Jika tidak ada printer default yang online, ambil printer online pertama
            if not target_printer:
                for printer in printers:
                    if printer.status == "online":
                        target_printer = printer
                        break
        
        if not target_printer:
            raise ValueError("Tidak ada printer yang tersedia atau online")
        
        if target_printer.status != "online":
            raise ValueError(f"Printer {target_printer.name} tidak online (status: {target_printer.status})")
        
        # Buat konten test page
        test_content = self._create_test_page_content(target_printer)
        
        # Simpan test page ke file temporary
        import tempfile
        with tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False, encoding='utf-8') as f:
            f.write(test_content)
            test_file_path = f.name
        
        try:
            # Buat job test print
            job = PrintJob(
                printer_id=target_printer.id,
                file_name="Test Page",
                file_path=test_file_path,
                file_size=len(test_content.encode('utf-8')),
                file_type="text/plain",
                title="Test Print Page",
                user=user,
                client_ip=client_ip,
                settings=PrintSettings(),  # Gunakan setting default
                total_pages=1
            )
            
            # Simpan job
            self._jobs[job.id] = job
            self._total_jobs += 1
            
            # Masukkan ke queue untuk diproses
            self._job_queue.put(job.id)
            
            logger.info(f"Test print job {job.id} submitted untuk printer {target_printer.name}")
            return job
            
        except Exception as e:
            # Hapus file temporary jika terjadi error
            try:
                os.unlink(test_file_path)
            except:
                pass
            raise ValueError(f"Gagal membuat test print job: {str(e)}")
    
    def _create_test_page_content(self, printer: Printer) -> str:
        """Buat konten halaman test"""
        now = datetime.now()
        
        content = f"""========================================
           HALAMAN TEST PRINT
========================================

Tanggal/Waktu: {now.strftime('%d/%m/%Y %H:%M:%S')}
Printer: {printer.name}
Status: {printer.status}
Driver: {printer.driver_name or 'Unknown'}

========================================
           TEST PATTERN
========================================

ABCDEFGHIJKLMNOPQRSTUVWXYZ
abcdefghijklmnopqrstuvwxyz
0123456789
!@#$%^&*()_+-=[]{{}}|;':",./<>?

========================================
           ALIGNMENT TEST
========================================

Left aligned text
    Center aligned text
                Right aligned text

========================================
           PRINT QUALITY TEST
========================================

████████████████████████████████████████
▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓
▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒
░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░

========================================

Test print berhasil!
Jika Anda dapat membaca teks ini dengan jelas,
printer berfungsi dengan baik.

========================================
"""
        return content
    
    def get_active_jobs(self) -> List[PrintJob]:
        """Dapatkan jobs yang sedang aktif"""
        return [job for job in self._jobs.values() if job.is_active]
    
    def get_printer_jobs(self, printer_id: str) -> List[PrintJob]:
        """Dapatkan jobs untuk printer tertentu"""
        return [job for job in self._jobs.values() if job.printer_id == printer_id]
    
    def cancel_job(self, job_id: str) -> bool:
        """Cancel job"""
        job = self.get_job(job_id)
        if not job:
            return False
        
        if job.status in [JobStatus.COMPLETED, JobStatus.FAILED, JobStatus.CANCELLED]:
            return False
        
        job.status = JobStatus.CANCELLED
        job.completed_at = datetime.now()
        
        logger.info(f"Job {job_id} cancelled")
        return True
    
    def pause_job(self, job_id: str) -> bool:
        """Pause job"""
        job = self.get_job(job_id)
        if not job:
            return False
        
        if job.status not in [JobStatus.PENDING, JobStatus.PROCESSING, JobStatus.PRINTING]:
            return False
        
        job.status = JobStatus.PAUSED
        
        logger.info(f"Job {job_id} paused")
        return True
    
    def resume_job(self, job_id: str) -> bool:
        """Resume paused job"""
        job = self.get_job(job_id)
        if not job or job.status != JobStatus.PAUSED:
            return False
        
        job.status = JobStatus.PENDING
        
        # Masukkan kembali ke queue
        self._job_queue.put(job_id)
        
        logger.info(f"Job {job_id} resumed")
        return True
    
    def retry_job(self, job_id: str) -> bool:
        """Retry failed job"""
        job = self.get_job(job_id)
        if not job or job.status != JobStatus.FAILED:
            return False
        
        if job.retry_count >= job.max_retries:
            return False
        
        job.status = JobStatus.PENDING
        job.retry_count += 1
        job.error_message = None
        job.started_at = None
        job.completed_at = None
        job.progress_percentage = 0.0
        
        # Masukkan kembali ke queue
        self._job_queue.put(job_id)
        
        logger.info(f"Job {job_id} retried (attempt {job.retry_count})")
        return True
    
    def get_queue_status(self) -> Dict[str, Any]:
        """Dapatkan status queue"""
        active_jobs = self.get_active_jobs()
        
        return {
            "queue_size": self._job_queue.qsize(),
            "active_jobs": len(active_jobs),
            "total_jobs": self._total_jobs,
            "completed_jobs": self._completed_jobs,
            "failed_jobs": self._failed_jobs,
            "worker_running": self._worker_thread and self._worker_thread.is_alive()
        }
    
    def update_job_progress(self, job_id: str, pages_printed: int, total_pages: int = None) -> bool:
        """Update job progress"""
        job = self.get_job(job_id)
        if not job:
            return False
        
        job.pages_printed = pages_printed
        if total_pages is not None:
            job.total_pages = total_pages
        
        # Calculate progress percentage
        if job.total_pages > 0:
            job.progress_percentage = (job.pages_printed / job.total_pages) * 100
        else:
            job.progress_percentage = 0.0
        
        logger.debug(f"Job {job_id} progress: {job.pages_printed}/{job.total_pages} ({job.progress_percentage:.1f}%)")
        return True
    
    def get_job_progress(self, job_id: str) -> Optional[Dict[str, Any]]:
        """Get real-time job progress"""
        job = self.get_job(job_id)
        if not job:
            return None
        
        return {
            "job_id": job.id,
            "status": job.status,
            "progress_percentage": job.progress_percentage,
            "pages_printed": job.pages_printed,
            "total_pages": job.total_pages,
            "created_at": job.created_at,
            "started_at": job.started_at,
            "estimated_completion": self._estimate_completion_time(job),
            "can_pause": job.status in [JobStatus.PENDING, JobStatus.PROCESSING, JobStatus.PRINTING],
            "can_resume": job.status == JobStatus.PAUSED,
            "can_cancel": job.status in [JobStatus.PENDING, JobStatus.PROCESSING, JobStatus.PRINTING, JobStatus.PAUSED]
        }
    
    def _estimate_completion_time(self, job: PrintJob) -> Optional[str]:
        """Estimate job completion time"""
        if not job.started_at or job.total_pages == 0 or job.pages_printed == 0:
            return None
        
        elapsed = (datetime.now() - job.started_at).total_seconds()
        pages_per_second = job.pages_printed / elapsed
        remaining_pages = job.total_pages - job.pages_printed
        
        if pages_per_second > 0:
            estimated_seconds = remaining_pages / pages_per_second
            estimated_completion = datetime.now() + timedelta(seconds=estimated_seconds)
            return estimated_completion.isoformat()
        
        return None
    
    def _start_worker(self):
        """Start worker thread"""
        if self._worker_thread and self._worker_thread.is_alive():
            return
        
        self._stop_worker = False
        self._worker_thread = threading.Thread(target=self._worker_loop, daemon=True)
        self._worker_thread.start()
        logger.info("Job worker thread started")
    
    def stop(self):
        """Stop the job service"""
        logger.info("Stopping job service...")
        self._stop_worker_thread()
        logger.info("Job service stopped")
    
    def _stop_worker_thread(self):
        """Stop worker thread"""
        self._stop_worker = True
        if self._worker_thread:
            self._worker_thread.join(timeout=5)
    
    def _worker_loop(self):
        """Main worker loop"""
        logger.info("Job worker loop started")
        
        while not self._stop_worker:
            try:
                # Ambil job dari queue dengan timeout
                job_id = self._job_queue.get(timeout=1)
                
                job = self.get_job(job_id)
                if not job:
                    continue
                
                # Process job
                self._process_job(job)
                
            except Empty:
                # Timeout, continue loop
                continue
            except Exception as e:
                logger.error(f"Error in worker loop: {e}")
        
        logger.info("Job worker loop stopped")
    
    def _process_job(self, job: PrintJob):
        """Process single job with fallback printer support"""
        try:
            logger.info(f"Processing job {job.id}")
            
            # Check if job is paused
            if job.status == JobStatus.PAUSED:
                logger.info(f"Job {job.id} is paused, skipping")
                return
            
            # Update status
            job.status = JobStatus.PROCESSING
            job.started_at = datetime.now()
            
            # Validasi file masih ada
            if not os.path.exists(job.file_path):
                raise Exception(f"File {job.file_path} not found")
            
            # Check if job is paused before printing
            if job.status == JobStatus.PAUSED:
                logger.info(f"Job {job.id} was paused before printing")
                return
            
            # Try to find available printer with fallback mechanism
            printer, actual_printer_id = self._find_available_printer(job.printer_id)
            if not printer:
                raise Exception(f"No available printer found (original: {job.printer_id})")
            
            # Update job with actual printer used if different
            if actual_printer_id != job.printer_id:
                logger.info(f"Using fallback printer {actual_printer_id} instead of {job.printer_id}")
                job.printer_id = actual_printer_id
            
            # Update status ke printing
            job.status = JobStatus.PRINTING
            
            # Lakukan printing berdasarkan tipe file
            success = self._print_file(job, printer)
            
            if success:
                job.status = JobStatus.COMPLETED
                job.progress_percentage = 100.0
                job.completed_at = datetime.now()
                self._completed_jobs += 1
                logger.info(f"Job {job.id} completed successfully on printer {actual_printer_id}")
            else:
                raise Exception("Printing failed")
                
        except Exception as e:
            job.status = JobStatus.FAILED
            job.error_message = str(e)
            job.completed_at = datetime.now()
            self._failed_jobs += 1
            logger.error(f"Job {job.id} failed: {e}")
    
    def _print_file(self, job: PrintJob, printer: Printer) -> bool:
        """Print file ke printer"""
        try:
            file_type = job.file_type.lower()
            
            if file_type == 'pdf':
                return self._print_pdf(job, printer)
            elif file_type in ['txt', 'text']:
                return self._print_text(job, printer)
            elif file_type in ['doc', 'docx']:
                return self._print_document(job, printer)
            elif file_type in ['jpg', 'jpeg', 'png', 'bmp', 'gif']:
                return self._print_image(job, printer)
            else:
                # Try generic printing
                return self._print_generic(job, printer)
                
        except Exception as e:
            logger.error(f"Error printing file {job.file_path}: {e}")
            return False
    
    def _print_pdf(self, job: PrintJob, printer: Printer) -> bool:
        """Print PDF file"""
        # Untuk PDF, kita bisa menggunakan Adobe Reader atau SumatraPDF
        # Atau convert ke image dulu
        try:
            # Implementasi sederhana: gunakan default PDF viewer
            import subprocess
            
            cmd = [
                'powershell', '-Command',
                f'Start-Process -FilePath "{job.file_path}" -Verb Print -WindowStyle Hidden'
            ]
            
            result = subprocess.run(cmd, capture_output=True, text=True)
            return result.returncode == 0
            
        except Exception as e:
            logger.error(f"Error printing PDF: {e}")
            return False
    
    def _print_text(self, job: PrintJob, printer: Printer) -> bool:
        """Print text file"""
        try:
            # Baca file text
            with open(job.file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # Estimate pages (assuming ~3000 chars per page)
            estimated_pages = max(1, len(content) // 3000)
            job.total_pages = estimated_pages
            
            # Print menggunakan Windows API
            printer_handle = win32print.OpenPrinter(printer.name)
            
            job_info = ('Test Page', None, 'RAW')
            
            job_id = win32print.StartDocPrinter(printer_handle, 1, job_info)
            
            # Simulate progress during printing
            import time
            content_bytes = content.encode('utf-8')
            chunk_size = len(content_bytes) // estimated_pages if estimated_pages > 1 else len(content_bytes)
            
            win32print.StartPagePrinter(printer_handle)
            
            for i in range(estimated_pages):
                # Check if job is paused or cancelled
                current_job = self.get_job(job.id)
                if current_job.status in [JobStatus.PAUSED, JobStatus.CANCELLED]:
                    win32print.EndPagePrinter(printer_handle)
                    win32print.EndDocPrinter(printer_handle)
                    win32print.ClosePrinter(printer_handle)
                    return False
                
                # Update progress
                self.update_job_progress(job.id, i + 1, estimated_pages)
                
                # Simulate printing time
                time.sleep(0.5)  # Simulate printing delay
            
            # Write all content at once (actual printing)
            win32print.WritePrinter(printer_handle, content_bytes)
            win32print.EndPagePrinter(printer_handle)
            win32print.EndDocPrinter(printer_handle)
            win32print.ClosePrinter(printer_handle)
            
            return True
            
        except Exception as e:
            logger.error(f"Error printing text: {e}")
            return False
    
    def _print_document(self, job: PrintJob, printer: Printer) -> bool:
        """Print document (Word, etc.)"""
        try:
            # Gunakan aplikasi default untuk print
            import subprocess
            
            cmd = [
                'powershell', '-Command',
                f'Start-Process -FilePath "{job.file_path}" -Verb Print -WindowStyle Hidden'
            ]
            
            result = subprocess.run(cmd, capture_output=True, text=True)
            return result.returncode == 0
            
        except Exception as e:
            logger.error(f"Error printing document: {e}")
            return False
    
    def _print_image(self, job: PrintJob, printer: Printer) -> bool:
        """Print image file"""
        try:
            # Gunakan aplikasi default untuk print
            import subprocess
            
            cmd = [
                'powershell', '-Command',
                f'Start-Process -FilePath "{job.file_path}" -Verb Print -WindowStyle Hidden'
            ]
            
            result = subprocess.run(cmd, capture_output=True, text=True)
            return result.returncode == 0
            
        except Exception as e:
            logger.error(f"Error printing image: {e}")
            return False
    
    def _find_available_printer(self, preferred_printer_id: str) -> tuple:
        """Find available printer with fallback mechanism
        
        Returns:
            tuple: (Printer object, actual_printer_id) or (None, None) if no printer available
        """
        # First try the preferred printer
        printer = self.printer_service.get_printer(preferred_printer_id)
        if printer:
            status = self.printer_service.get_printer_status(preferred_printer_id)
            if status == PrinterStatus.ONLINE:
                logger.info(f"Using preferred printer {preferred_printer_id} (status: {status})")
                return printer, preferred_printer_id
            else:
                logger.warning(f"Preferred printer {preferred_printer_id} not ready (status: {status}), trying fallback")
        else:
            logger.warning(f"Preferred printer {preferred_printer_id} not found, trying fallback")
        
        # Try fallback printers - get all available printers
        all_printers = self.printer_service.get_all_printers(force_refresh=True)
        
        # Sort printers by priority: default first, then by name
        fallback_printers = sorted(all_printers, key=lambda p: (not p.is_default, p.name))
        
        for fallback_printer in fallback_printers:
            if fallback_printer.name == preferred_printer_id:
                continue  # Skip the original printer we already tried
            
            status = self.printer_service.get_printer_status(fallback_printer.name)
            if status == PrinterStatus.ONLINE:
                logger.info(f"Found fallback printer {fallback_printer.name} (status: {status})")
                return fallback_printer, fallback_printer.name
            else:
                logger.debug(f"Fallback printer {fallback_printer.name} not ready (status: {status})")
        
        logger.error("No available printers found for fallback")
        return None, None
    
    def _print_generic(self, job: PrintJob, printer: Printer) -> bool:
        """Generic printing untuk file types lain"""
        try:
            # Coba print dengan aplikasi default
            import subprocess
            
            cmd = [
                'powershell', '-Command',
                f'Start-Process -FilePath "{job.file_path}" -Verb Print -WindowStyle Hidden'
            ]
            
            result = subprocess.run(cmd, capture_output=True, text=True)
            return result.returncode == 0
            
        except Exception as e:
            logger.error(f"Error generic printing: {e}")
            return False
    
    def __del__(self):
        """Cleanup saat object dihapus"""
        self._stop_worker_thread()
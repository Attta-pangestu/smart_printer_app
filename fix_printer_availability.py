#!/usr/bin/env python3
"""
Fix Printer Availability Issue
Script untuk memperbaiki masalah "No available printer found" dengan implementasi queue system
"""

import os
import shutil
from datetime import datetime

def backup_original_file():
    """Backup file asli sebelum modifikasi"""
    original_file = "D:\\Gawean Rebinmas\\Driver_Epson_L120\\server\\services\\job_service.py"
    backup_file = f"{original_file}.backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    
    if os.path.exists(original_file):
        shutil.copy2(original_file, backup_file)
        print(f"Backup created: {backup_file}")
        return True
    else:
        print(f"Original file not found: {original_file}")
        return False

def create_improved_find_available_printer():
    """Buat implementasi yang diperbaiki untuk _find_available_printer"""
    
    improved_code = '''
    def _find_available_printer(self, preferred_printer_id: str, allow_busy: bool = True) -> tuple:
        """Find available printer with improved fallback mechanism
        
        Args:
            preferred_printer_id: ID printer yang diinginkan
            allow_busy: Apakah boleh menggunakan printer yang sedang busy
        
        Returns:
            tuple: (Printer object, actual_printer_id) or (None, None) if no printer available
        """
        # First try the preferred printer
        printer = self.printer_service.get_printer(preferred_printer_id)
        if printer:
            status = self.printer_service.get_printer_status(preferred_printer_id)
            
            # Accept ONLINE or BUSY (if allowed) status
            if status == PrinterStatus.ONLINE:
                logger.info(f"Using preferred printer {preferred_printer_id} (status: {status})")
                return printer, preferred_printer_id
            elif status == PrinterStatus.BUSY and allow_busy:
                logger.info(f"Using preferred printer {preferred_printer_id} (status: {status}, will queue)")
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
            
            # Accept ONLINE or BUSY (if allowed) status for fallback too
            if status == PrinterStatus.ONLINE:
                logger.info(f"Found fallback printer {fallback_printer.name} (status: {status})")
                return fallback_printer, fallback_printer.name
            elif status == PrinterStatus.BUSY and allow_busy:
                logger.info(f"Found fallback printer {fallback_printer.name} (status: {status}, will queue)")
                return fallback_printer, fallback_printer.name
            else:
                logger.debug(f"Fallback printer {fallback_printer.name} not ready (status: {status})")
        
        # If no printer found with allow_busy=True, try again with stricter criteria
        if allow_busy:
            logger.warning("No printer found with relaxed criteria, trying strict mode")
            return self._find_available_printer(preferred_printer_id, allow_busy=False)
        
        logger.error("No available printers found for fallback")
        return None, None
'''
    
    return improved_code

def create_improved_process_job():
    """Buat implementasi yang diperbaiki untuk _process_job dengan retry mechanism"""
    
    improved_code = '''
    def _process_job(self, job: PrintJob):
        """Process a single print job with improved error handling and retry mechanism"""
        import time
        
        max_printer_wait_attempts = 3
        printer_wait_delay = 2  # seconds
        
        try:
            logger.info(f"Processing job {job.id}: {job.file_name}")
            
            # Validasi file exists
            if not os.path.exists(job.file_path):
                raise FileNotFoundError(f"File not found: {job.file_path}")
            
            # Check if job was paused
            if job.status == JobStatus.PAUSED:
                logger.info(f"Job {job.id} was paused before printing")
                return
            
            # Try to find available printer with retry mechanism
            printer = None
            actual_printer_id = None
            
            for attempt in range(max_printer_wait_attempts):
                printer, actual_printer_id = self._find_available_printer(job.printer_id, allow_busy=True)
                
                if printer:
                    break
                    
                if attempt < max_printer_wait_attempts - 1:
                    logger.warning(f"No printer available (attempt {attempt + 1}/{max_printer_wait_attempts}), waiting {printer_wait_delay}s...")
                    time.sleep(printer_wait_delay)
                else:
                    raise Exception(f"No available printer found after {max_printer_wait_attempts} attempts (original: {job.printer_id})")
            
            # Update job with actual printer used if different
            if actual_printer_id != job.printer_id:
                logger.info(f"Using fallback printer {actual_printer_id} instead of {job.printer_id}")
                job.printer_id = actual_printer_id
            
            # Update status ke printing
            job.status = JobStatus.PRINTING
            
            # Check if printer is busy and implement simple queuing
            printer_status = self.printer_service.get_printer_status(actual_printer_id)
            if printer_status == PrinterStatus.BUSY:
                logger.info(f"Printer {actual_printer_id} is busy, waiting for it to become available...")
                
                # Wait for printer to become available (with timeout)
                wait_timeout = 30  # seconds
                wait_start = time.time()
                
                while time.time() - wait_start < wait_timeout:
                    current_status = self.printer_service.get_printer_status(actual_printer_id)
                    if current_status == PrinterStatus.ONLINE:
                        logger.info(f"Printer {actual_printer_id} is now available")
                        break
                    elif current_status in [PrinterStatus.ERROR, PrinterStatus.OFFLINE]:
                        raise Exception(f"Printer {actual_printer_id} became unavailable (status: {current_status})")
                    
                    time.sleep(1)  # Check every second
                else:
                    logger.warning(f"Printer {actual_printer_id} still busy after {wait_timeout}s, proceeding anyway")
            
            # Lakukan printing berdasarkan tipe file
            success = self._print_file(job, printer)
            
            if success:
                job.status = JobStatus.COMPLETED
                job.completed_at = datetime.now()
                job.progress_percentage = 100.0
                logger.info(f"Job {job.id} completed successfully")
            else:
                raise Exception("Printing failed")
                
        except Exception as e:
            error_msg = str(e)
            logger.error(f"Job {job.id} failed: {error_msg}")
            
            job.status = JobStatus.FAILED
            job.error_message = error_msg
            job.completed_at = datetime.now()
            
            # Increment retry count
            job.retry_count += 1
            
            # Save temp file for debugging
            if hasattr(job, 'file_path') and os.path.exists(job.file_path):
                temp_debug_path = f"temp/failed_{job.id}_{job.file_name}"
                try:
                    os.makedirs(os.path.dirname(temp_debug_path), exist_ok=True)
                    shutil.copy2(job.file_path, temp_debug_path)
                    logger.info(f"Failed job file saved to: {temp_debug_path}")
                except Exception as copy_error:
                    logger.warning(f"Could not save failed job file: {copy_error}")
'''
    
    return improved_code

def main():
    print("=== PRINTER AVAILABILITY FIX ===")
    print(f"Time: {datetime.now()}")
    
    # Backup original file
    print("\n1. Creating backup...")
    if not backup_original_file():
        print("Failed to create backup. Exiting.")
        return
    
    print("\n2. Generated improved code:")
    print("\n--- Improved _find_available_printer ---")
    improved_find_printer = create_improved_find_available_printer()
    print(improved_find_printer[:500] + "...")
    
    print("\n--- Improved _process_job ---")
    improved_process_job = create_improved_process_job()
    print(improved_process_job[:500] + "...")
    
    print("\n3. Key improvements:")
    print("   - Allow printer to be used even when BUSY (with queuing)")
    print("   - Retry mechanism for finding available printer")
    print("   - Wait for busy printer to become available")
    print("   - Better error handling and logging")
    print("   - Fallback to strict mode if relaxed mode fails")
    
    print("\n4. Next steps:")
    print("   - Review the generated code")
    print("   - Apply the changes to job_service.py")
    print("   - Test with concurrent print jobs")
    print("   - Monitor for reduced failure rate")
    
    print("\nNote: The backup file has been created. You can restore it if needed.")

if __name__ == "__main__":
    main()
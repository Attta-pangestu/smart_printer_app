#!/usr/bin/env python3
"""
Script untuk debug masalah printer availability
Analisis mengapa beberapa pekerjaan gagal dengan 'No available printer found'
"""

import requests
import time
import json
from datetime import datetime

SERVER_URL = "http://localhost:8081"

def check_printer_status():
    """Cek status printer"""
    try:
        response = requests.get(f"{SERVER_URL}/api/printers")
        if response.status_code == 200:
            data = response.json()
            print(f"\n=== Printer Status at {datetime.now().strftime('%H:%M:%S')} ===")
            
            # Handle different response formats
            if isinstance(data, dict) and 'data' in data:
                printers = data['data']
            elif isinstance(data, list):
                printers = data
            else:
                print(f"Unexpected response format: {data}")
                return
                
            for printer in printers:
                if isinstance(printer, dict):
                    print(f"Printer: {printer.get('name', 'unknown')} (ID: {printer.get('id', 'unknown')})")
                    print(f"  Status: {printer.get('status', 'unknown')}")
                    print(f"  Is Default: {printer.get('is_default', False)}")
                    print(f"  Is Online: {printer.get('is_online', False)}")
                    
                    # Cek detailed status
                    printer_id = printer.get('id')
                    if printer_id:
                        detail_response = requests.get(f"{SERVER_URL}/api/printers/{printer_id}/status")
                        if detail_response.status_code == 200:
                            detail = detail_response.json()
                            if isinstance(detail, dict):
                                print(f"  Detailed Status: {detail.get('status', 'unknown')}")
                                print(f"  Jobs in Queue: {detail.get('jobs_count', 0)}")
                    print()
                else:
                    print(f"Invalid printer data: {printer}")
        else:
            print(f"Error getting printers: {response.status_code} - {response.text}")
    except Exception as e:
        print(f"Error checking printer status: {e}")

def check_jobs_status():
    """Cek status pekerjaan"""
    try:
        response = requests.get(f"{SERVER_URL}/api/jobs")
        if response.status_code == 200:
            data = response.json()
            print(f"\n=== Jobs Status at {datetime.now().strftime('%H:%M:%S')} ===")
            
            # Handle different response formats
            if isinstance(data, dict) and 'data' in data:
                jobs = data['data']
            elif isinstance(data, list):
                jobs = data
            else:
                print(f"Unexpected response format: {data}")
                return
            
            # Group by status
            status_counts = {}
            failed_jobs = []
            
            for job in jobs:
                if isinstance(job, dict):
                    status = job.get('status', 'unknown')
                    status_counts[status] = status_counts.get(status, 0) + 1
                    
                    if status == 'failed':
                        failed_jobs.append(job)
                else:
                    print(f"Invalid job data: {job}")
            
            print("Job Status Summary:")
            for status, count in status_counts.items():
                print(f"  {status}: {count}")
            
            if failed_jobs:
                print("\nFailed Jobs Details:")
                for job in failed_jobs[-5:]:  # Show last 5 failed jobs
                    print(f"  Job ID: {job.get('id', 'unknown')}")
                    print(f"    Printer: {job.get('printer_id', 'unknown')}")
                    print(f"    Error: {job.get('error_message', 'no error message')}")
                    print(f"    Created: {job.get('created_at', 'unknown')}")
                    print(f"    Settings: {job.get('print_settings', {})}")
                    print()
        else:
            print(f"Error getting jobs: {response.status_code} - {response.text}")
    except Exception as e:
        print(f"Error checking jobs status: {e}")

def simulate_concurrent_jobs():
    """Simulasi pekerjaan bersamaan untuk test race condition"""
    print(f"\n=== Simulating Concurrent Jobs at {datetime.now().strftime('%H:%M:%S')} ===")
    
    # Cek status printer sebelum test
    check_printer_status()
    
    # Buat beberapa pekerjaan bersamaan
    jobs_created = []
    
    for i in range(3):
        try:
            # Buat test document
            files = {
                'file': ('test_concurrent.pdf', open('test_document.pdf', 'rb'), 'application/pdf')
            }
            
            data = {
                'printer_id': 'epson_l120_series',
                'copies': 1,
                'color_mode': 'color',
                'paper_size': 'A4',
                'orientation': 'portrait',
                'page_range_type': 'all' if i == 0 else ('odd' if i == 1 else 'range'),
                'page_range': '' if i != 2 else '1-2'
            }
            
            print(f"Creating job {i+1} with settings: {data}")
            
            response = requests.post(
                f"{SERVER_URL}/api/print/with-processing",
                files=files,
                data=data
            )
            
            files['file'][1].close()  # Close file
            
            if response.status_code == 200:
                job_data = response.json()
                if isinstance(job_data, dict) and 'job_id' in job_data:
                    jobs_created.append(job_data['job_id'])
                    print(f"  ✓ Job {i+1} created: {job_data['job_id']}")
                else:
                    print(f"  ✗ Job {i+1} unexpected response: {job_data}")
            else:
                print(f"  ✗ Job {i+1} failed: {response.status_code} - {response.text}")
                
        except Exception as e:
            print(f"  ✗ Job {i+1} error: {e}")
        
        # Small delay between jobs
        time.sleep(0.5)
    
    # Monitor jobs for a while
    print(f"\nMonitoring {len(jobs_created)} jobs...")
    
    for _ in range(10):  # Monitor for 10 seconds
        time.sleep(1)
        
        # Check job status
        completed = 0
        failed = 0
        processing = 0
        
        for job_id in jobs_created:
            try:
                response = requests.get(f"{SERVER_URL}/api/jobs/{job_id}")
                if response.status_code == 200:
                    job = response.json()
                    if isinstance(job, dict):
                        status = job.get('status', 'unknown')
                        
                        if status == 'completed':
                            completed += 1
                        elif status == 'failed':
                            failed += 1
                            print(f"  Job {job_id} failed: {job.get('error_message', 'no error')}")
                        elif status in ['processing', 'printing']:
                            processing += 1
            except Exception as e:
                print(f"  Error checking job {job_id}: {e}")
        
        print(f"  Status: {completed} completed, {failed} failed, {processing} processing")
        
        if completed + failed == len(jobs_created):
            break
    
    print("\nFinal job status check:")
    check_jobs_status()

def main():
    """Main function"""
    print("=== Printer Availability Debug Tool ===")
    print(f"Server: {SERVER_URL}")
    print(f"Time: {datetime.now()}")
    
    # Initial status check
    check_printer_status()
    check_jobs_status()
    
    # Test concurrent jobs
    simulate_concurrent_jobs()
    
    # Final status check
    print("\n=== Final Status ===")
    check_printer_status()
    check_jobs_status()

if __name__ == "__main__":
    main()
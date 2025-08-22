#!/usr/bin/env python3
"""
Debug Job Creation - Monitor real-time job creation and processing
Analyzes the gap between server job status and actual printer output
"""

import requests
import time
import win32print
import json
from datetime import datetime

def get_printer_queue_status():
    """Get actual printer queue status"""
    try:
        printer_name = "EPSON L120 Series"
        handle = win32print.OpenPrinter(printer_name)
        printer_info = win32print.GetPrinter(handle, 2)
        win32print.ClosePrinter(handle)
        
        return {
            'jobs_in_queue': printer_info['cJobs'],
            'status': printer_info['Status'],
            'attributes': printer_info['Attributes']
        }
    except Exception as e:
        return {'error': str(e)}

def get_server_jobs():
    """Get jobs from server API"""
    try:
        response = requests.get('http://127.0.0.1:8081/api/jobs')
        if response.status_code == 200:
            return response.json()['jobs']
        return []
    except Exception as e:
        print(f"Error getting server jobs: {e}")
        return []

def monitor_job_creation():
    """Monitor job creation and processing in real-time"""
    print("=== REAL-TIME JOB CREATION MONITOR ===")
    print("Monitoring server jobs vs actual printer queue...\n")
    
    last_job_count = 0
    monitoring = True
    
    while monitoring:
        try:
            # Get current status
            server_jobs = get_server_jobs()
            printer_status = get_printer_queue_status()
            
            current_time = datetime.now().strftime("%H:%M:%S")
            
            # Check for new jobs
            if len(server_jobs) != last_job_count:
                print(f"\n[{current_time}] JOB COUNT CHANGED: {last_job_count} -> {len(server_jobs)}")
                
                # Analyze latest job
                if server_jobs:
                    latest_job = server_jobs[-1]
                    print(f"Latest Job Analysis:")
                    print(f"  ID: {latest_job['id']}")
                    print(f"  Status: {latest_job['status']}")
                    print(f"  Progress: {latest_job['progress_percentage']}%")
                    print(f"  Created: {latest_job['created_at']}")
                    print(f"  Started: {latest_job.get('started_at', 'Not started')}")
                    print(f"  Completed: {latest_job.get('completed_at', 'Not completed')}")
                    print(f"  Error: {latest_job.get('error_message', 'None')}")
                    
                    # Compare with printer queue
                    print(f"\nPrinter Queue Status:")
                    print(f"  Jobs in Queue: {printer_status.get('jobs_in_queue', 'Unknown')}")
                    print(f"  Printer Status: {printer_status.get('status', 'Unknown')}")
                    
                    # Detect discrepancy
                    if (latest_job['status'] == 'completed' and 
                        latest_job['progress_percentage'] == 100.0 and
                        printer_status.get('jobs_in_queue', 0) == 0):
                        print("\n⚠️  CRITICAL ISSUE DETECTED:")
                        print("   Server reports job completed (100%)")
                        print("   But printer queue shows 0 jobs (no actual printing)")
                        print("   This confirms FALSE POSITIVE completion!")
                
                last_job_count = len(server_jobs)
            
            # Show current status every 5 seconds
            print(f"[{current_time}] Server Jobs: {len(server_jobs)}, Printer Queue: {printer_status.get('jobs_in_queue', 0)}", end='\r')
            
            time.sleep(1)
            
        except KeyboardInterrupt:
            print("\n\nMonitoring stopped by user.")
            monitoring = False
        except Exception as e:
            print(f"\nError during monitoring: {e}")
            time.sleep(2)

def analyze_existing_jobs():
    """Analyze existing jobs for patterns"""
    print("=== ANALYZING EXISTING JOBS ===")
    
    server_jobs = get_server_jobs()
    printer_status = get_printer_queue_status()
    
    print(f"Total jobs in server: {len(server_jobs)}")
    print(f"Jobs in printer queue: {printer_status.get('jobs_in_queue', 0)}")
    
    if server_jobs:
        print("\nJob Status Summary:")
        status_counts = {}
        for job in server_jobs:
            status = job['status']
            status_counts[status] = status_counts.get(status, 0) + 1
        
        for status, count in status_counts.items():
            print(f"  {status}: {count} jobs")
        
        # Analyze completed jobs
        completed_jobs = [j for j in server_jobs if j['status'] == 'completed']
        if completed_jobs:
            print(f"\nCompleted Jobs Analysis ({len(completed_jobs)} jobs):")
            for job in completed_jobs[-3:]:  # Show last 3
                duration = 'Unknown'
                if job.get('started_at') and job.get('completed_at'):
                    start = datetime.fromisoformat(job['started_at'].replace('Z', '+00:00'))
                    end = datetime.fromisoformat(job['completed_at'].replace('Z', '+00:00'))
                    duration = f"{(end - start).total_seconds():.1f}s"
                
                print(f"  Job {job['id'][:8]}...")
                print(f"    Duration: {duration}")
                print(f"    Progress: {job['progress_percentage']}%")
                print(f"    Pages: {job.get('pages_printed', 0)}/{job.get('total_pages', 0)}")
                print(f"    Error: {job.get('error_message', 'None')}")
    
    # Critical analysis
    if (len(server_jobs) > 0 and 
        printer_status.get('jobs_in_queue', 0) == 0 and
        any(j['status'] == 'completed' for j in server_jobs)):
        print("\n🚨 CRITICAL FINDING:")
        print("   Server has completed jobs but printer queue is empty")
        print("   This indicates FALSE POSITIVE job completion")
        print("   Jobs are marked complete without actual printing")

if __name__ == "__main__":
    print("Job Creation Debug Tool")
    print("1. Analyze existing jobs")
    print("2. Monitor real-time job creation")
    print("3. Both")
    
    choice = input("\nSelect option (1-3): ").strip()
    
    if choice in ['1', '3']:
        analyze_existing_jobs()
        print("\n" + "="*50)
    
    if choice in ['2', '3']:
        if choice == '3':
            input("\nPress Enter to start real-time monitoring...")
        monitor_job_creation()
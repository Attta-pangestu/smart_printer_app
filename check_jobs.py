#!/usr/bin/env python3
"""
Script to check job details and verify document processing settings
"""

try:
    import requests
except ImportError:
    print("❌ requests library not found. Installing...")
    import subprocess
    subprocess.check_call(["pip", "install", "requests"])
    import requests

import json
from datetime import datetime

def check_jobs():
    """Check all jobs and their processing details"""
    base_url = "http://localhost:8081"
    
    print("=== Checking Print Jobs ===\n")
    
    try:
        # Get all jobs
        response = requests.get(f"{base_url}/api/jobs")
        if response.status_code == 200:
            data = response.json()
            jobs = data.get('jobs', [])
            
            if not jobs:
                print("No jobs found.")
                return
                
            print(f"Found {len(jobs)} jobs:\n")
            
            for i, job in enumerate(jobs, 1):
                print(f"Job {i}: {job['id']}")
                print(f"  File: {job['file_name']}")
                print(f"  Printer: {job['printer_id']}")
                print(f"  Status: {job['status']}")
                print(f"  Created: {job['created_at']}")
                
                # Check print settings
                if 'settings' in job:
                    print(f"  Print Settings:")
                    settings = job['settings']
                    for key, value in settings.items():
                        print(f"    {key}: {value}")
                
                # Check metadata for document settings
                if 'metadata' in job and job['metadata']:
                    metadata = job['metadata']
                    print(f"  Metadata:")
                    
                    if 'document_settings' in metadata:
                        print(f"    Document Settings:")
                        doc_settings = metadata['document_settings']
                        for key, value in doc_settings.items():
                            print(f"      {key}: {value}")
                    
                    if 'original_file_path' in metadata:
                        print(f"    Original File: {metadata['original_file_path']}")
                    
                    if 'processed_file_path' in metadata:
                        print(f"    Processed File: {metadata['processed_file_path']}")
                
                # Check error message if failed
                if job['status'] == 'failed' and 'error_message' in job and job['error_message']:
                    print(f"  Error: {job['error_message']}")
                
                print("  ---")
                
        else:
            print(f"❌ Failed to get jobs: {response.status_code}")
            print(f"Response: {response.text}")
            
    except Exception as e:
        print(f"❌ Error checking jobs: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    check_jobs()
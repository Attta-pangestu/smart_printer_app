#!/usr/bin/env python3
"""
Test script untuk menguji integrasi pencetakan PDF melalui server API
"""

import requests
import json
import time
import os

def test_pdf_printing():
    """Test PDF printing through server API"""
    base_url = "http://localhost:8081"
    
    # Test file path
    test_file = os.path.abspath("test_integration.pdf")
    
    if not os.path.exists(test_file):
        print(f"❌ Test file not found: {test_file}")
        return False
    
    print(f"🔍 Testing PDF printing integration...")
    print(f"📄 Test file: {test_file}")
    
    try:
        # 1. Get available printers
        print("\n1️⃣ Getting available printers...")
        response = requests.get(f"{base_url}/api/printers")
        if response.status_code != 200:
            print(f"❌ Failed to get printers: {response.status_code}")
            return False
        
        printers_data = response.json()
        print(f"📋 Response: {printers_data}")
        
        # Handle different response formats
        if isinstance(printers_data, dict):
            if 'data' in printers_data:
                printers = printers_data['data']
            elif 'printers' in printers_data:
                printers = printers_data['printers']
            else:
                printers = [printers_data]  # Single printer response
        else:
            printers = printers_data  # Direct list
        
        print(f"✅ Found {len(printers)} printers")
        
        if not printers:
            print("❌ No printers available")
            return False
        
        # Use first available printer
        printer = printers[0]
        printer_id = printer.get('id') or printer.get('printer_id')
        printer_name = printer.get('name') or printer.get('printer_name')
        print(f"🖨️ Using printer: {printer_name} (ID: {printer_id})")
        
        if not printer_id:
            print("❌ Could not get printer ID")
            return False
        
        # 2. Create print job
        print("\n2️⃣ Creating print job...")
        job_data = {
            "file_path": test_file,
            "printer_id": printer_id,
            "copies": 1,
            "settings": {
                "color_mode": "color",
                "paper_size": "A4",
                "quality": "normal"
            }
        }
        
        response = requests.post(f"{base_url}/api/jobs/submit", json=job_data)
        if response.status_code != 201:
            print(f"❌ Failed to create job: {response.status_code}")
            print(f"Response: {response.text}")
            return False
        
        job = response.json()
        job_id = job['id']
        print(f"✅ Job created successfully: {job_id}")
        print(f"📋 Job status: {job['status']}")
        
        # 3. Monitor job progress
        print("\n3️⃣ Monitoring job progress...")
        max_wait = 30  # seconds
        start_time = time.time()
        
        while time.time() - start_time < max_wait:
            response = requests.get(f"{base_url}/api/jobs/{job_id}")
            if response.status_code != 200:
                print(f"❌ Failed to get job status: {response.status_code}")
                return False
            
            job = response.json()
            status = job['status']
            print(f"📊 Job status: {status}")
            
            if status == 'completed':
                print(f"✅ Job completed successfully!")
                print(f"📄 Pages printed: {job.get('pages_printed', 'N/A')}")
                return True
            elif status == 'failed':
                print(f"❌ Job failed: {job.get('error_message', 'Unknown error')}")
                return False
            elif status in ['pending', 'processing']:
                print(f"⏳ Job still {status}, waiting...")
                time.sleep(2)
            else:
                print(f"❓ Unknown status: {status}")
                time.sleep(2)
        
        print(f"⏰ Timeout waiting for job completion")
        return False
        
    except requests.exceptions.ConnectionError:
        print("❌ Cannot connect to server. Make sure server is running on http://localhost:8081")
        return False
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        return False

def main():
    """Main test function"""
    print("🧪 PDF Printing Integration Test")
    print("=" * 40)
    
    success = test_pdf_printing()
    
    print("\n" + "=" * 40)
    if success:
        print("🎉 Integration test PASSED!")
        print("✅ PDF printing is working correctly through the server API")
    else:
        print("💥 Integration test FAILED!")
        print("❌ There are issues with PDF printing integration")
    
    return success

if __name__ == "__main__":
    main()
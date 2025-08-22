#!/usr/bin/env python3
"""
Test script untuk memverifikasi integrasi pengaturan UI dengan document processing
"""

try:
    import requests
except ImportError:
    print("❌ requests library not found. Installing...")
    import subprocess
    subprocess.check_call(["pip", "install", "requests"])
    import requests

import json
import os
from pathlib import Path
from datetime import datetime

def test_ui_integration():
    """Test integrasi pengaturan UI dengan backend"""
    
    # Server URL
    base_url = "http://localhost:8081"
    
    print("=== Testing UI Settings Integration ===")
    
    # 1. Test upload file
    print("\n1. Testing file upload...")
    test_file = Path("test_document.pdf")
    if not test_file.exists():
        print(f"❌ Test file {test_file} not found")
        return False
    
    try:
        with open(test_file, 'rb') as f:
            files = {'file': f}
            response = requests.post(f"{base_url}/api/files/upload", files=files)
        
        if response.status_code == 200:
            file_data = response.json()
            file_id = file_data['file_id']
            print(f"✅ File uploaded successfully: {file_id}")
        else:
            print(f"❌ File upload failed: {response.status_code}")
            return False
    except Exception as e:
        print(f"❌ File upload error: {e}")
        return False
    
    # 2. Test get printers
    print("\n2. Testing printer discovery...")
    try:
        print(f"Making request to: {base_url}/api/printers")
        response = requests.get(f"{base_url}/api/printers")
        print(f"Response status: {response.status_code}")
        print(f"Response content: {response.text[:200]}...")
        
        if response.status_code == 200:
            printers_data = response.json()
            print(f"Parsed JSON: {printers_data}")
            
            # Check if response has 'printers' key
            if 'printers' in printers_data:
                printers = printers_data['printers']
            else:
                printers = printers_data
                
            if printers and len(printers) > 0:
                printer_id = printers[0]['id']
                print(f"✅ Found printer: {printer_id}")
            else:
                print("⚠️ No printers found, using default printer ID")
                printer_id = "epson_l120_series"  # Default printer ID
        else:
            print(f"❌ Printer discovery failed: {response.status_code}")
            print(f"Response: {response.text}")
            return False
    except Exception as e:
        print(f"❌ Printer discovery error: {e}")
        import traceback
        traceback.print_exc()
        return False
    
    # 3. Test print with processing - All pages
    print("\n3. Testing print with processing (all pages)...")
    test_cases = [
        {
            "name": "All pages, A4 portrait",
            "print_settings": {
                "copies": 1,
                "color_mode": "color",
                "paper_size": "A4",
                "orientation": "portrait",
                "quality": "normal",
                "page_range": ""
            },
            "document_settings": {
                "page_range_type": "all",
                "page_range": "",
                "orientation": "portrait",
                "paper_size": "A4"
            }
        },
        {
            "name": "Odd pages only",
            "print_settings": {
                "copies": 1,
                "color_mode": "grayscale",
                "paper_size": "A4",
                "orientation": "portrait",
                "quality": "high",
                "page_range": "odd"
            },
            "document_settings": {
                "page_range_type": "odd",
                "page_range": "odd",
                "orientation": "portrait",
                "paper_size": "A4"
            }
        },
        {
            "name": "Page range 1-3",
            "print_settings": {
                "copies": 2,
                "color_mode": "color",
                "paper_size": "A4",
                "orientation": "landscape",
                "quality": "normal",
                "page_range": "1-3"
            },
            "document_settings": {
                "page_range_type": "range",
                "page_range": "1-3",
                "orientation": "landscape",
                "paper_size": "A4"
            }
        }
    ]
    
    for i, test_case in enumerate(test_cases, 1):
        print(f"\n3.{i} Testing: {test_case['name']}")
        
        request_data = {
            "printer_id": printer_id,
            "file_id": file_id,
            "print_settings": test_case["print_settings"],
            "document_settings": test_case["document_settings"]
        }
        
        try:
            response = requests.post(
                f"{base_url}/api/print/with-processing",
                headers={'Content-Type': 'application/json'},
                data=json.dumps(request_data)
            )
            
            if response.status_code == 200:
                result = response.json()
                job_id = result.get('job_id')
                print(f"✅ Print job submitted: {job_id}")
                print(f"   Settings applied: {result.get('document_settings_applied')}")
            else:
                print(f"❌ Print job failed: {response.status_code}")
                print(f"   Error: {response.text}")
                
        except Exception as e:
            print(f"❌ Print job error: {e}")
    
    print("\n=== Integration Test Complete ===")
    return True

if __name__ == "__main__":
    test_ui_integration()
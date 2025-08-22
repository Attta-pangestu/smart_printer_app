import requests
import json
import time
import os

def test_upload_and_print():
    base_url = "http://localhost:8081"
    test_file = "test_document.pdf"
    
    print("=== Testing Upload and Print Process ===")
    
    # 1. Check if file exists
    if not os.path.exists(test_file):
        print(f"❌ Test file {test_file} not found!")
        return
    
    print(f"✓ Test file {test_file} found")
    
    try:
        # 2. Get printer list
        print("\n--- Step 1: Getting printer list ---")
        response = requests.get(f"{base_url}/api/printers")
        print(f"Status: {response.status_code}")
        print(f"Response: {response.text}")
        
        if response.status_code != 200:
            print("❌ Failed to get printer list")
            return
        
        result = response.json()
        printers = result.get('printers', [])
        if not printers:
            print("❌ No printers found")
            return
        
        printer_id = printers[0]['id']
        print(f"✓ Using printer: {printers[0]['name']} (ID: {printer_id})")
        
        # 3. Upload file
        print("\n--- Step 2: Uploading file ---")
        with open(test_file, 'rb') as f:
            files = {'file': (test_file, f, 'application/pdf')}
            response = requests.post(f"{base_url}/api/files/upload", files=files)
        
        print(f"Status: {response.status_code}")
        print(f"Response: {response.text}")
        
        if response.status_code != 200:
            print("❌ Failed to upload file")
            return
        
        upload_result = response.json()
        file_path = upload_result.get('file_path')
        print(f"✓ File uploaded: {file_path}")
        
        # 3. Create print job
        print("\n--- Step 3: Creating print job ---")
        file_id = upload_result.get('file_id')
        job_data = {
            "printer_id": printer_id,
            "file_id": file_id,
            "settings": {
                "copies": 1,
                "color_mode": "color",
                "paper_size": "A4",
                "orientation": "portrait",
                "quality": "normal",
                "duplex": "none"
            }
        }
        
        response = requests.post(f"{base_url}/api/jobs/submit", json=job_data)
        print(f"Status: {response.status_code}")
        print(f"Response: {response.text}")
        
        if response.status_code != 200:
            print("❌ Failed to create print job")
            return
        
        job_result = response.json()
        job_id = job_result.get('job_id')
        print(f"✓ Print job created: {job_id}")
        
        # 5. Monitor job progress
        print("\n--- Step 4: Monitoring job progress ---")
        max_wait = 60  # 60 seconds
        start_time = time.time()
        
        while time.time() - start_time < max_wait:
            response = requests.get(f"{base_url}/api/jobs/{job_id}")
            if response.status_code == 200:
                job_status = response.json()
                status = job_status.get('status')
                progress = job_status.get('progress', 0)
                total_pages = job_status.get('total_pages', 0)
                
                print(f"Job Status: {status}, Progress: {progress}/{total_pages}")
                
                if status in ['completed', 'failed', 'cancelled']:
                    print(f"✓ Job finished with status: {status}")
                    break
            else:
                print(f"❌ Failed to get job status: {response.status_code}")
                break
            
            time.sleep(2)
        else:
            print("⚠ Job monitoring timeout reached")
        
        # 6. Get final job details
        print("\n--- Step 5: Final job details ---")
        response = requests.get(f"{base_url}/api/jobs/{job_id}")
        if response.status_code == 200:
            final_status = response.json()
            print(f"Final Status: {json.dumps(final_status, indent=2)}")
        
        # 7. Check printer status
        print("\n--- Step 6: Checking printer status ---")
        response = requests.get(f"{base_url}/api/printers/{printer_id}/status")
        if response.status_code == 200:
            printer_status = response.json()
            print(f"Printer Status: {json.dumps(printer_status, indent=2)}")
        
        # 8. Get detailed printer status
        print("\n--- Step 7: Detailed printer status ---")
        response = requests.get(f"{base_url}/api/printers/{printer_id}/status/detailed")
        if response.status_code == 200:
            detailed_status = response.json()
            print(f"Detailed Status: {json.dumps(detailed_status, indent=2)}")
        
    except requests.exceptions.ConnectionError:
        print("❌ Cannot connect to server. Make sure server is running on http://localhost:8081")
    except Exception as e:
        print(f"❌ Error during testing: {e}")
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    test_upload_and_print()
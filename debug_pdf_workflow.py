import requests
import json
import time
import os

def test_pdf_workflow():
    base_url = "http://localhost:8081"
    
    print("=== DEBUG PDF UPLOAD & PRINT WORKFLOW ===")
    
    # 1. Get printers
    print("\n1. Getting available printers...")
    try:
        response = requests.get(f"{base_url}/api/printers")
        print(f"Status: {response.status_code}")
        printers_data = response.json()
        print(f"Response: {json.dumps(printers_data, indent=2)}")
        
        # Extract printer info
        if isinstance(printers_data, list):
            printers = printers_data
        elif 'data' in printers_data:
            printers = printers_data['data']
        elif 'printers' in printers_data:
            printers = printers_data['printers']
        else:
            printers = [printers_data]
            
        if not printers:
            print("No printers found!")
            return
            
        printer = printers[0]
        printer_id = printer.get('id') or printer.get('printer_id') or printer.get('name', '').lower().replace(' ', '_')
        printer_name = printer.get('name') or printer.get('printer_name', 'Unknown')
        
        print(f"Using printer: {printer_name} (ID: {printer_id})")
        
    except Exception as e:
        print(f"Error getting printers: {e}")
        return
    
    # 2. Upload PDF file
    print("\n2. Uploading PDF file...")
    pdf_path = "D:\\Gawean Rebinmas\\Driver_Epson_L120\\test_integration.pdf"
    
    if not os.path.exists(pdf_path):
        print(f"PDF file not found: {pdf_path}")
        return
        
    try:
        with open(pdf_path, 'rb') as f:
            files = {'file': ('test_integration.pdf', f, 'application/pdf')}
            response = requests.post(f"{base_url}/api/files/upload", files=files)
            
        print(f"Upload Status: {response.status_code}")
        upload_data = response.json()
        print(f"Upload Response: {json.dumps(upload_data, indent=2)}")
        
        if response.status_code != 200:
            print("Upload failed!")
            return
            
        file_id = upload_data.get('file_id') or upload_data.get('file_name')
        print(f"File uploaded successfully. File ID: {file_id}")
        
    except Exception as e:
        print(f"Error uploading file: {e}")
        return
    
    # 3. Submit print job
    print("\n3. Submitting print job...")
    try:
        job_data = {
            "printer_id": printer_id,
            "file_id": file_id,
            "settings": {
                "copies": 1,
                "orientation": "portrait"
            }
        }
        
        print(f"Job data: {json.dumps(job_data, indent=2)}")
        
        response = requests.post(f"{base_url}/api/jobs/submit", json=job_data)
        print(f"Job Submit Status: {response.status_code}")
        
        if response.status_code == 200:
            job_response = response.json()
            print(f"Job Response: {json.dumps(job_response, indent=2)}")
            
            job_id = job_response.get('job_id')
            if job_id:
                print(f"Job created successfully. Job ID: {job_id}")
                
                # 4. Monitor job status
                print("\n4. Monitoring job status...")
                for i in range(30):  # Monitor for 30 seconds
                    try:
                        status_response = requests.get(f"{base_url}/api/jobs/{job_id}")
                        if status_response.status_code == 200:
                            status_data = status_response.json()
                            status = status_data.get('status', 'unknown')
                            progress = status_data.get('progress', 0)
                            
                            print(f"[{i+1:2d}s] Status: {status}, Progress: {progress}%")
                            
                            if status in ['completed', 'failed', 'cancelled']:
                                print(f"\nFinal status: {status}")
                                print(f"Full status data: {json.dumps(status_data, indent=2)}")
                                break
                        else:
                            print(f"[{i+1:2d}s] Error getting status: {status_response.status_code}")
                            
                    except Exception as e:
                        print(f"[{i+1:2d}s] Error: {e}")
                        
                    time.sleep(1)
                else:
                    print("\nJob monitoring timeout after 30 seconds")
            else:
                print("No job ID returned!")
        else:
            print(f"Job submission failed: {response.text}")
            
    except Exception as e:
        print(f"Error submitting job: {e}")
        return
    
    print("\n=== WORKFLOW COMPLETE ===")

if __name__ == "__main__":
    test_pdf_workflow()
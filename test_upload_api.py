import requests
import os

# Test file upload API
url = "http://localhost:8081/api/files/upload"
test_file = "test_upload.txt"

if os.path.exists(test_file):
    with open(test_file, 'rb') as f:
        files = {'file': (test_file, f, 'text/plain')}
        try:
            response = requests.post(url, files=files)
            print(f"Status Code: {response.status_code}")
            print(f"Response: {response.text}")
            if response.status_code == 200:
                result = response.json()
                print(f"Upload successful!")
                print(f"File ID: {result.get('file_id')}")
                print(f"File Name: {result.get('file_name')}")
                print(f"Upload Path: {result.get('upload_path')}")
            else:
                print(f"Upload failed: {response.text}")
        except Exception as e:
            print(f"Error: {e}")
else:
    print(f"Test file {test_file} not found")
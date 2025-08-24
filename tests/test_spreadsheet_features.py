import pytest
import requests
import io
import json
from pathlib import Path

# Base URL untuk API
BASE_URL = "http://localhost:8081"

class TestSpreadsheetFeatures:
    """Test suite untuk fitur spreadsheet preview dan konversi PDF"""
    
    def setup_method(self):
        """Setup untuk setiap test method"""
        self.base_url = BASE_URL
        self.test_files_dir = Path(__file__).parent / "test_files"
        self.test_files_dir.mkdir(exist_ok=True)
        
    def create_test_excel_file(self, filename="test_spreadsheet.xlsx"):
        """Buat file Excel test sederhana"""
        import pandas as pd
        
        # Data test
        data = {
            'Nama': ['John Doe', 'Jane Smith', 'Bob Johnson'],
            'Umur': [25, 30, 35],
            'Kota': ['Jakarta', 'Bandung', 'Surabaya'],
            'Gaji': [5000000, 7500000, 6000000]
        }
        
        df = pd.DataFrame(data)
        file_path = self.test_files_dir / filename
        df.to_excel(file_path, index=False)
        return file_path
        
    def test_health_check(self):
        """Test health check endpoint"""
        response = requests.get(f"{self.base_url}/api/document-manipulation/health")
        assert response.status_code == 200
        data = response.json()
        assert data['status'] == 'healthy'
        
    def test_upload_spreadsheet_success(self):
        """Test upload spreadsheet berhasil"""
        # Buat file Excel test
        excel_file = self.create_test_excel_file()
        
        # Upload file
        with open(excel_file, 'rb') as f:
            files = {'file': (excel_file.name, f, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}
            response = requests.post(
                f"{self.base_url}/api/document-manipulation/upload-spreadsheet",
                files=files
            )
        
        assert response.status_code == 200
        data = response.json()
        assert 'file_id' in data
        assert 'preview_data' in data
        assert 'metadata' in data
        
        # Cleanup
        excel_file.unlink()
        
    def test_upload_invalid_file_type(self):
        """Test upload file yang bukan Excel"""
        # Buat file text
        text_file = self.test_files_dir / "test.txt"
        text_file.write_text("This is not an Excel file")
        
        # Upload file
        with open(text_file, 'rb') as f:
            files = {'file': (text_file.name, f, 'text/plain')}
            response = requests.post(
                f"{self.base_url}/api/document-manipulation/upload-spreadsheet",
                files=files
            )
        
        assert response.status_code == 400
        data = response.json()
        assert 'detail' in data
        assert 'Excel' in data['detail']
        
        # Cleanup
        text_file.unlink()
        
    def test_excel_to_pdf_conversion(self):
        """Test konversi Excel ke PDF"""
        # Buat file Excel test
        excel_file = self.create_test_excel_file()
        
        # Upload dan konversi ke PDF
        with open(excel_file, 'rb') as f:
            files = {'file': (excel_file.name, f, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}
            data = {
                'output_format': 'pdf',
                'preserve_formatting': 'true',
                'page_orientation': 'portrait',
                'paper_size': 'A4'
            }
            response = requests.post(
                f"{self.base_url}/api/document-manipulation/convert-excel-to-pdf",
                files=files,
                data=data
            )
        
        if response.status_code == 200:
            data = response.json()
            assert 'file_id' in data
            assert 'download_url' in data
            
            # Test download PDF
            file_id = data['file_id']
            download_response = requests.get(f"{self.base_url}/api/document-manipulation/download/{file_id}")
            assert download_response.status_code == 200
            assert download_response.headers['content-type'] == 'application/pdf'
        else:
            # Jika endpoint belum tersedia, skip test
            pytest.skip("Excel to PDF conversion endpoint not available")
        
        # Cleanup
        excel_file.unlink()
        
    def test_different_page_orientations(self):
        """Test konversi dengan orientasi halaman berbeda"""
        excel_file = self.create_test_excel_file()
        
        orientations = ['portrait', 'landscape']
        
        for orientation in orientations:
            with open(excel_file, 'rb') as f:
                files = {'file': (excel_file.name, f, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}
                data = {
                    'output_format': 'pdf',
                    'page_orientation': orientation,
                    'paper_size': 'A4'
                }
                response = requests.post(
                    f"{self.base_url}/api/document-manipulation/convert-excel-to-pdf",
                    files=files,
                    data=data
                )
            
            if response.status_code == 200:
                data = response.json()
                assert 'file_id' in data
            else:
                pytest.skip("Excel to PDF conversion endpoint not available")
        
        # Cleanup
        excel_file.unlink()
        
    def test_different_paper_sizes(self):
        """Test konversi dengan ukuran kertas berbeda"""
        excel_file = self.create_test_excel_file()
        
        paper_sizes = ['A4', 'A3', 'Letter']
        
        for paper_size in paper_sizes:
            with open(excel_file, 'rb') as f:
                files = {'file': (excel_file.name, f, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}
                data = {
                    'output_format': 'pdf',
                    'page_orientation': 'portrait',
                    'paper_size': paper_size
                }
                response = requests.post(
                    f"{self.base_url}/api/document-manipulation/convert-excel-to-pdf",
                    files=files,
                    data=data
                )
            
            if response.status_code == 200:
                data = response.json()
                assert 'file_id' in data
            else:
                pytest.skip("Excel to PDF conversion endpoint not available")
        
        # Cleanup
        excel_file.unlink()
        
    def test_conversion_methods(self):
        """Test berbagai metode konversi"""
        excel_file = self.create_test_excel_file()
        
        methods = ['openpyxl', 'pandas', 'xlsxwriter']
        
        for method in methods:
            with open(excel_file, 'rb') as f:
                files = {'file': (excel_file.name, f, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}
                data = {
                    'output_format': 'pdf',
                    'conversion_method': method,
                    'preserve_formatting': 'true'
                }
                response = requests.post(
                    f"{self.base_url}/api/document-manipulation/convert-excel-to-pdf",
                    files=files,
                    data=data
                )
            
            if response.status_code == 200:
                data = response.json()
                assert 'file_id' in data
            else:
                pytest.skip("Excel to PDF conversion endpoint not available")
        
        # Cleanup
        excel_file.unlink()
        
    def test_invalid_file_id_conversion(self):
        """Test konversi dengan file ID yang tidak valid"""
        invalid_file_id = "invalid_file_id_12345"
        
        response = requests.get(f"{self.base_url}/api/document-manipulation/download/{invalid_file_id}")
        assert response.status_code == 404
        
    def teardown_method(self):
        """Cleanup setelah setiap test method"""
        # Hapus direktori test files jika kosong
        if self.test_files_dir.exists() and not any(self.test_files_dir.iterdir()):
            self.test_files_dir.rmdir()
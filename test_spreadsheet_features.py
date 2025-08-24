import pytest
import requests
import os
import tempfile
from pathlib import Path
import time

# Base URL for the API
BASE_URL = "http://localhost:8081"

class TestSpreadsheetFeatures:
    """Test class untuk fitur spreadsheet preview dan PDF conversion"""
    
    def setup_method(self):
        """Setup untuk setiap test method"""
        self.test_file_path = Path("test_files/Hektaris_progress_undevelop.xlsx")
        self.uploaded_file_id = None
    
    def test_health_check(self):
        """Test health check endpoint"""
        response = requests.get(f"{BASE_URL}/api/document-manipulation/health")
        assert response.status_code == 200
        data = response.json()
        assert data["status"] == "healthy"
        assert data["service"] == "document_manipulation"
    
    def test_upload_excel_file(self):
        """Test upload Excel file"""
        if not self.test_file_path.exists():
            pytest.skip(f"Test file not found: {self.test_file_path}")
        
        with open(self.test_file_path, "rb") as f:
            files = {"file": ("test.xlsx", f, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
            response = requests.post(f"{BASE_URL}/api/document-manipulation/upload", files=files)
        
        assert response.status_code == 200
        data = response.json()
        assert "file_id" in data
        assert data["filename"] == "test.xlsx"
        
        # Store file_id for other tests
        self.uploaded_file_id = data["file_id"]
        return data["file_id"]
    
    def test_spreadsheet_preview(self):
        """Test spreadsheet preview endpoint"""
        file_id = self.test_upload_excel_file()
        
        response = requests.post(
            f"{BASE_URL}/api/document-manipulation/upload-spreadsheet",
            json={
                "preserve_formatting": True,
                "max_rows": 50,
                "max_columns": 10
            },
            headers={"X-File-ID": file_id}
        )
        
        assert response.status_code == 200
        data = response.json()
        assert "sheets" in data
        assert "file_id" in data
        assert len(data["sheets"]) > 0
        
        # Check first sheet structure
        first_sheet = data["sheets"][0]
        assert "name" in first_sheet
        assert "data" in first_sheet
        assert "formatting" in first_sheet
    
    def test_excel_to_pdf_conversion(self):
        """Test Excel to PDF conversion"""
        file_id = self.test_upload_excel_file()
        
        response = requests.post(
            f"{BASE_URL}/api/document-manipulation/convert-excel-to-pdf",
            json={
                "file_id": file_id,
                "preserve_formatting": True,
                "preserve_charts": True,
                "preserve_images": True,
                "preserve_formulas": True,
                "preserve_colors": True,
                "preserve_fonts": True,
                "preserve_borders": True,
                "preserve_alignment": True,
                "fit_to_page": True,
                "high_quality": True,
                "include_all_sheets": True,
                "page_orientation": "landscape",
                "paper_size": "A4",
                "margins": "normal",
                "conversion_method": "auto"
            }
        )
        
        assert response.status_code == 200
        data = response.json()
        assert data["status"] == "success"
        assert "pdf_file_id" in data
        assert "pdf_filename" in data
        assert "conversion_result" in data
        assert "download_url" in data
        
        # Check conversion result structure
        conversion_result = data["conversion_result"]
        assert "method" in conversion_result
        assert "quality_score" in conversion_result
        assert "features_preserved" in conversion_result
    
    def test_download_converted_pdf(self):
        """Test download converted PDF file"""
        file_id = self.test_upload_excel_file()
        
        # Convert to PDF first
        response = requests.post(
            f"{BASE_URL}/api/document-manipulation/convert-excel-to-pdf",
            json={
                "file_id": file_id,
                "preserve_formatting": True,
                "page_orientation": "portrait",
                "conversion_method": "auto"
            }
        )
        
        assert response.status_code == 200
        pdf_data = response.json()
        pdf_file_id = pdf_data["pdf_file_id"]
        
        # Download the PDF
        download_response = requests.get(f"{BASE_URL}/api/document-manipulation/download/{pdf_file_id}")
        assert download_response.status_code == 200
        assert download_response.headers["content-type"] == "application/pdf"
        assert len(download_response.content) > 0
    
    def test_invalid_file_id_conversion(self):
        """Test conversion with invalid file_id"""
        response = requests.post(
            f"{BASE_URL}/api/document-manipulation/convert-excel-to-pdf",
            json={
                "file_id": "invalid-file-id",
                "preserve_formatting": True
            }
        )
        
        assert response.status_code == 404
        data = response.json()
        assert "detail" in data
        assert "not found" in data["detail"].lower()
    
    def test_conversion_options_validation(self):
        """Test various conversion options"""
        file_id = self.test_upload_excel_file()
        
        # Test different page orientations
        for orientation in ["portrait", "landscape", "auto"]:
            response = requests.post(
                f"{BASE_URL}/api/document-manipulation/convert-excel-to-pdf",
                json={
                    "file_id": file_id,
                    "page_orientation": orientation,
                    "preserve_formatting": True
                }
            )
            assert response.status_code == 200
        
        # Test different paper sizes
        for paper_size in ["A4", "A3", "Letter"]:
            response = requests.post(
                f"{BASE_URL}/api/document-manipulation/convert-excel-to-pdf",
                json={
                    "file_id": file_id,
                    "paper_size": paper_size,
                    "preserve_formatting": True
                }
            )
            assert response.status_code == 200
        
        # Test different conversion methods
        for method in ["auto", "openpyxl", "basic"]:
            response = requests.post(
                f"{BASE_URL}/api/document-manipulation/convert-excel-to-pdf",
                json={
                    "file_id": file_id,
                    "conversion_method": method,
                    "preserve_formatting": True
                }
            )
            assert response.status_code == 200

if __name__ == "__main__":
    pytest.main(["-v", __file__])
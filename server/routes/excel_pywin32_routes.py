from fastapi import APIRouter, HTTPException, UploadFile, File, Form, Depends
from fastapi.responses import FileResponse, JSONResponse
from typing import List, Optional, Dict, Any
import json
import os
import logging
from pathlib import Path
import tempfile
from datetime import datetime

from services.excel_pywin32_service import ExcelPyWin32Service
from services.file_service import FileService

logger = logging.getLogger(__name__)

router = APIRouter()

# Initialize services
excel_service = None
file_service = None

def init_services(excel_svc: ExcelPyWin32Service, file_svc: FileService):
    """Initialize services for the router"""
    global excel_service, file_service
    excel_service = excel_svc
    file_service = file_svc

@router.post("/upload-excel")
async def upload_excel_file(file: UploadFile = File(...)):
    """Upload Excel file and get basic information"""
    try:
        # Validate file type
        if not file.filename.lower().endswith(('.xlsx', '.xls')):
            raise HTTPException(status_code=400, detail="File harus berformat Excel (.xlsx atau .xls)")
        
        # Read and save file
        file_content = await file.read()
        file_info = file_service.save_uploaded_file(file_content, file.filename)
        
        # Get Excel information
        file_path = str(file_service.upload_dir / file_info["name"])
        excel_info = excel_service.get_excel_info(file_path)
        
        if not excel_info['success']:
            raise HTTPException(status_code=500, detail=f"Gagal membaca file Excel: {excel_info.get('error', 'Unknown error')}")
        
        return {
            "success": True,
            "file_id": file_info["name"],
            "original_name": file.filename,
            "file_path": file_path,
            "excel_info": excel_info,
            "message": "File Excel berhasil diupload"
        }
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error uploading Excel file: {e}")
        raise HTTPException(status_code=500, detail=f"Gagal mengupload file: {str(e)}")

@router.get("/excel-info/{file_id}")
async def get_excel_info(file_id: str):
    """Get detailed information about uploaded Excel file"""
    try:
        file_path = str(file_service.upload_dir / file_id)
        
        if not os.path.exists(file_path):
            raise HTTPException(status_code=404, detail="File tidak ditemukan")
        
        excel_info = excel_service.get_excel_info(file_path)
        
        if not excel_info['success']:
            raise HTTPException(status_code=500, detail=f"Gagal membaca file Excel: {excel_info.get('error', 'Unknown error')}")
        
        return excel_info
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error getting Excel info: {e}")
        raise HTTPException(status_code=500, detail=f"Gagal mendapatkan informasi Excel: {str(e)}")

@router.post("/preview-pdf-settings")
async def preview_pdf_settings(
    file_id: str = Form(...),
    sheet_names: Optional[str] = Form(None),
    cell_range: Optional[str] = Form(None),
    print_settings: Optional[str] = Form(None)
):
    """Preview PDF conversion settings without converting"""
    try:
        file_path = str(file_service.upload_dir / file_id)
        
        if not os.path.exists(file_path):
            raise HTTPException(status_code=404, detail="File tidak ditemukan")
        
        # Parse parameters
        selected_sheets = None
        if sheet_names:
            try:
                selected_sheets = json.loads(sheet_names)
            except json.JSONDecodeError:
                selected_sheets = [name.strip() for name in sheet_names.split(',') if name.strip()]
        
        settings = None
        if print_settings:
            try:
                settings = json.loads(print_settings)
            except json.JSONDecodeError:
                logger.warning(f"Invalid print settings JSON: {print_settings}")
                settings = excel_service.get_default_print_settings()
        else:
            settings = excel_service.get_default_print_settings()
        
        # Get preview information
        preview_info = excel_service.preview_pdf_settings(
            input_path=file_path,
            sheet_names=selected_sheets,
            cell_range=cell_range,
            print_settings=settings
        )
        
        if not preview_info['success']:
            raise HTTPException(status_code=500, detail=preview_info.get('error', 'Unknown error'))
        
        return preview_info
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error previewing PDF settings: {e}")
        raise HTTPException(status_code=500, detail=f"Gagal preview pengaturan PDF: {str(e)}")

@router.post("/convert-to-pdf")
async def convert_excel_to_pdf(
    file_id: str = Form(...),
    sheet_names: Optional[str] = Form(None),
    cell_range: Optional[str] = Form(None),
    print_settings: Optional[str] = Form(None),
    output_filename: Optional[str] = Form(None)
):
    """Convert Excel file to PDF with specified settings"""
    try:
        file_path = str(file_service.upload_dir / file_id)
        
        if not os.path.exists(file_path):
            raise HTTPException(status_code=404, detail="File tidak ditemukan")
        
        # Parse parameters
        selected_sheets = None
        if sheet_names:
            try:
                selected_sheets = json.loads(sheet_names)
            except json.JSONDecodeError:
                selected_sheets = [name.strip() for name in sheet_names.split(',') if name.strip()]
        
        settings = None
        if print_settings:
            try:
                settings = json.loads(print_settings)
            except json.JSONDecodeError:
                logger.warning(f"Invalid print settings JSON: {print_settings}")
                settings = excel_service.get_default_print_settings()
        else:
            settings = excel_service.get_default_print_settings()
        
        # Generate output path
        if output_filename:
            output_path = str(file_service.temp_dir / output_filename)
        else:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            original_name = Path(file_path).stem
            output_path = str(file_service.temp_dir / f"{timestamp}_{original_name}.pdf")
        
        # Convert to PDF
        result = excel_service.convert_to_pdf(
            input_path=file_path,
            output_path=output_path,
            sheet_names=selected_sheets,
            cell_range=cell_range,
            print_settings=settings
        )
        
        if not result['success']:
            raise HTTPException(status_code=500, detail=result.get('error', 'Conversion failed'))
        
        # Get PDF file info
        pdf_filename = Path(result['output_path']).name
        
        return {
            "success": True,
            "pdf_file_id": pdf_filename,
            "pdf_path": result['output_path'],
            "file_size": result['file_size'],
            "download_url": f"/api/excel-pywin32/download-pdf/{pdf_filename}",
            "message": result['message']
        }
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error converting Excel to PDF: {e}")
        raise HTTPException(status_code=500, detail=f"Gagal mengkonversi ke PDF: {str(e)}")

@router.get("/download-pdf/{pdf_filename}")
async def download_pdf(pdf_filename: str):
    """Download generated PDF file"""
    try:
        pdf_path = file_service.temp_dir / pdf_filename
        
        if not pdf_path.exists():
            raise HTTPException(status_code=404, detail="File PDF tidak ditemukan")
        
        return FileResponse(
            path=str(pdf_path),
            filename=pdf_filename,
            media_type="application/pdf"
        )
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error downloading PDF: {e}")
        raise HTTPException(status_code=500, detail=f"Gagal mendownload PDF: {str(e)}")

@router.get("/default-print-settings")
async def get_default_print_settings():
    """Get default print settings"""
    try:
        return {
            "success": True,
            "settings": excel_service.get_default_print_settings()
        }
    except Exception as e:
        logger.error(f"Error getting default settings: {e}")
        raise HTTPException(status_code=500, detail=f"Gagal mendapatkan pengaturan default: {str(e)}")

@router.delete("/cleanup/{file_id}")
async def cleanup_files(file_id: str):
    """Clean up uploaded and generated files"""
    try:
        # Remove uploaded file
        upload_path = file_service.upload_dir / file_id
        if upload_path.exists():
            upload_path.unlink()
        
        # Remove associated PDF files
        file_stem = Path(file_id).stem
        for pdf_file in file_service.temp_dir.glob(f"*{file_stem}*.pdf"):
            try:
                pdf_file.unlink()
            except Exception as e:
                logger.warning(f"Could not remove PDF file {pdf_file}: {e}")
        
        return {
            "success": True,
            "message": "File berhasil dihapus"
        }
        
    except Exception as e:
        logger.error(f"Error cleaning up files: {e}")
        raise HTTPException(status_code=500, detail=f"Gagal menghapus file: {str(e)}")

@router.get("/health")
async def health_check():
    """Health check endpoint"""
    try:
        # Test Excel COM availability
        import win32com.client
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Quit()
        
        return {
            "success": True,
            "status": "healthy",
            "excel_available": True,
            "message": "Excel PyWin32 service is ready"
        }
    except Exception as e:
        logger.error(f"Health check failed: {e}")
        return {
            "success": False,
            "status": "unhealthy",
            "excel_available": False,
            "error": str(e),
            "message": "Excel PyWin32 service is not available"
        }
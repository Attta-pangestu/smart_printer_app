from fastapi import APIRouter, HTTPException, UploadFile, File, Form, Depends
from fastapi.responses import FileResponse, JSONResponse
from typing import List, Optional, Dict, Any
import json
import os
import logging
from pathlib import Path
import tempfile
from datetime import datetime

from services.excel_visual_service import ExcelVisualService
from services.excel_pywin32_service import ExcelPyWin32Service
from services.file_service import FileService

logger = logging.getLogger(__name__)

router = APIRouter()

# Initialize services
excel_visual_service = None
excel_pywin32_service = None
file_service = None

def init_visual_services(visual_svc: ExcelVisualService, pywin32_svc: ExcelPyWin32Service, file_svc: FileService):
    """Initialize services for the visual Excel router"""
    global excel_visual_service, excel_pywin32_service, file_service
    excel_visual_service = visual_svc
    excel_pywin32_service = pywin32_svc
    file_service = file_svc

@router.post("/upload-excel-visual")
async def upload_excel_for_visual_selection(file: UploadFile = File(...)):
    """Upload Excel file untuk visual selection"""
    try:
        # Validate file type
        if not file.filename.lower().endswith(('.xlsx', '.xls')):
            raise HTTPException(status_code=400, detail="File harus berformat Excel (.xlsx atau .xls)")
        
        # Read and save file
        file_content = await file.read()
        file_info = file_service.save_uploaded_file(file_content, file.filename)
        
        # Get Excel sheets information
        file_path = str(file_service.upload_dir / file_info["name"])
        sheets_info = excel_visual_service.get_excel_sheets_info(file_path)
        
        if not sheets_info['success']:
            raise HTTPException(status_code=500, detail=f"Gagal membaca file Excel: {sheets_info.get('error', 'Unknown error')}")
        
        return {
            "success": True,
            "file_id": file_info["name"],
            "original_name": file.filename,
            "file_size": len(file_content),
            "sheets_info": sheets_info
        }
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error uploading Excel file: {e}")
        raise HTTPException(status_code=500, detail=f"Error uploading file: {str(e)}")

@router.get("/excel-data/{file_id}")
async def get_excel_data_for_preview(
    file_id: str,
    sheet_name: Optional[str] = None,
    max_rows: int = 100,
    max_cols: int = 50
):
    """Mendapatkan data Excel untuk preview visual"""
    try:
        file_path = str(file_service.upload_dir / file_id)
        
        if not os.path.exists(file_path):
            raise HTTPException(status_code=404, detail="File tidak ditemukan")
        
        excel_data = excel_visual_service.read_excel_data(
            file_path, 
            sheet_name=sheet_name,
            max_rows=max_rows,
            max_cols=max_cols
        )
        
        if not excel_data['success']:
            raise HTTPException(status_code=500, detail=f"Gagal membaca data Excel: {excel_data.get('error', 'Unknown error')}")
        
        return excel_data
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error getting Excel data: {e}")
        raise HTTPException(status_code=500, detail=f"Error getting Excel data: {str(e)}")

@router.post("/validate-cell-range")
async def validate_selected_cell_range(
    file_id: str = Form(...),
    sheet_name: str = Form(...),
    cell_range: str = Form(...)
):
    """Validasi range sel yang dipilih"""
    try:
        file_path = str(file_service.upload_dir / file_id)
        
        if not os.path.exists(file_path):
            raise HTTPException(status_code=404, detail="File tidak ditemukan")
        
        validation_result = excel_visual_service.validate_cell_range(
            file_path, sheet_name, cell_range
        )
        
        return validation_result
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error validating cell range: {e}")
        raise HTTPException(status_code=500, detail=f"Error validating cell range: {str(e)}")

@router.post("/preview-selected-area")
async def preview_selected_area(
    file_id: str = Form(...),
    sheet_name: str = Form(...),
    cell_range: str = Form(...),
    print_settings: Optional[str] = Form(None)
):
    """Preview area yang dipilih dengan pengaturan print"""
    try:
        file_path = str(file_service.upload_dir / file_id)
        
        if not os.path.exists(file_path):
            raise HTTPException(status_code=404, detail="File tidak ditemukan")
        
        # Validasi range sel
        validation = excel_visual_service.validate_cell_range(file_path, sheet_name, cell_range)
        if not validation['success'] or not validation['valid']:
            raise HTTPException(status_code=400, detail=validation.get('error', 'Range sel tidak valid'))
        
        # Mendapatkan data dari range yang dipilih
        range_data = excel_visual_service.get_cell_range_data(file_path, sheet_name, cell_range)
        if not range_data['success']:
            raise HTTPException(status_code=500, detail=range_data.get('error', 'Gagal membaca data range'))
        
        # Parse print settings jika ada
        settings = {}
        if print_settings:
            try:
                settings = json.loads(print_settings)
            except json.JSONDecodeError:
                logger.warning("Invalid print settings JSON, using defaults")
        
        # Preview dengan pywin32 service untuk estimasi halaman
        preview_result = excel_pywin32_service.preview_pdf_settings(
            file_path,
            sheet_names=[sheet_name],
            cell_range=cell_range,
            print_settings=settings
        )
        
        return {
            "success": True,
            "range_data": range_data['range_data'],
            "range_info": range_data['range_info'],
            "validation": validation,
            "preview_info": preview_result,
            "print_settings": settings
        }
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error previewing selected area: {e}")
        raise HTTPException(status_code=500, detail=f"Error previewing selected area: {str(e)}")

@router.post("/convert-selected-area-to-pdf")
async def convert_selected_area_to_pdf(
    file_id: str = Form(...),
    sheet_name: str = Form(...),
    cell_range: str = Form(...),
    print_settings: Optional[str] = Form(None),
    output_filename: Optional[str] = Form(None)
):
    """Konversi area yang dipilih ke PDF"""
    try:
        file_path = str(file_service.upload_dir / file_id)
        
        if not os.path.exists(file_path):
            raise HTTPException(status_code=404, detail="File tidak ditemukan")
        
        # Validasi range sel
        validation = excel_visual_service.validate_cell_range(file_path, sheet_name, cell_range)
        if not validation['success'] or not validation['valid']:
            raise HTTPException(status_code=400, detail=validation.get('error', 'Range sel tidak valid'))
        
        # Parse print settings
        settings = {}
        if print_settings:
            try:
                settings = json.loads(print_settings)
            except json.JSONDecodeError:
                logger.warning("Invalid print settings JSON, using defaults")
                settings = excel_pywin32_service.get_default_print_settings()
        else:
            settings = excel_pywin32_service.get_default_print_settings()
        
        # Generate output filename jika tidak diberikan
        if not output_filename:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            range_safe = cell_range.replace(':', '_')
            output_filename = f"{timestamp}_{sheet_name}_{range_safe}.pdf"
        
        # Konversi menggunakan pywin32 service
        conversion_result = excel_pywin32_service.convert_to_pdf(
            input_path=file_path,
            output_path=None,  # Will be generated automatically
            sheet_names=[sheet_name],
            cell_range=cell_range,
            print_settings=settings
        )
        
        if not conversion_result['success']:
            raise HTTPException(status_code=500, detail=f"Konversi gagal: {conversion_result.get('error', 'Unknown error')}")
        
        return {
            "success": True,
            "pdf_filename": os.path.basename(conversion_result['output_path']),
            "pdf_path": conversion_result['output_path'],
            "range_info": {
                "sheet_name": sheet_name,
                "cell_range": cell_range,
                "validation": validation
            },
            "conversion_info": conversion_result,
            "download_url": f"/api/excel-visual/download-pdf/{os.path.basename(conversion_result['output_path'])}"
        }
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error converting selected area to PDF: {e}")
        raise HTTPException(status_code=500, detail=f"Error converting to PDF: {str(e)}")

@router.get("/download-pdf/{pdf_filename}")
async def download_generated_pdf(pdf_filename: str):
    """Download PDF yang telah dikonversi"""
    try:
        # Cari file PDF di temp directory
        pdf_path = None
        
        # Cek di temp directory dari excel_pywin32_service
        temp_dir = excel_pywin32_service.temp_dir
        potential_path = temp_dir / pdf_filename
        
        if potential_path.exists():
            pdf_path = str(potential_path)
        else:
            # Cek di output directory jika ada
            output_dir = Path("server/output")
            if output_dir.exists():
                potential_path = output_dir / pdf_filename
                if potential_path.exists():
                    pdf_path = str(potential_path)
        
        if not pdf_path or not os.path.exists(pdf_path):
            raise HTTPException(status_code=404, detail="File PDF tidak ditemukan")
        
        return FileResponse(
            path=pdf_path,
            filename=pdf_filename,
            media_type='application/pdf'
        )
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error downloading PDF: {e}")
        raise HTTPException(status_code=500, detail=f"Error downloading PDF: {str(e)}")

@router.get("/cell-range-data/{file_id}")
async def get_cell_range_data(
    file_id: str,
    sheet_name: str,
    cell_range: str
):
    """Mendapatkan data dari range sel tertentu"""
    try:
        file_path = str(file_service.upload_dir / file_id)
        
        if not os.path.exists(file_path):
            raise HTTPException(status_code=404, detail="File tidak ditemukan")
        
        range_data = excel_visual_service.get_cell_range_data(file_path, sheet_name, cell_range)
        
        if not range_data['success']:
            raise HTTPException(status_code=500, detail=range_data.get('error', 'Gagal membaca data range'))
        
        return range_data
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error getting cell range data: {e}")
        raise HTTPException(status_code=500, detail=f"Error getting cell range data: {str(e)}")

@router.delete("/cleanup-visual/{file_id}")
async def cleanup_visual_files(file_id: str):
    """Cleanup file yang diupload untuk visual selection"""
    try:
        file_path = file_service.upload_dir / file_id
        
        if file_path.exists():
            file_path.unlink()
            return {
                "success": True,
                "message": "File berhasil dihapus"
            }
        else:
            return {
                "success": False,
                "message": "File tidak ditemukan"
            }
            
    except Exception as e:
        logger.error(f"Error cleaning up file: {e}")
        raise HTTPException(status_code=500, detail=f"Error cleaning up file: {str(e)}")

@router.get("/health-visual")
async def health_check_visual():
    """Health check untuk visual Excel service"""
    try:
        return {
            "success": True,
            "service": "Excel Visual Selection Service",
            "status": "healthy",
            "timestamp": datetime.now().isoformat(),
            "services_initialized": {
                "excel_visual_service": excel_visual_service is not None,
                "excel_pywin32_service": excel_pywin32_service is not None,
                "file_service": file_service is not None
            }
        }
    except Exception as e:
        logger.error(f"Health check failed: {e}")
        return {
            "success": False,
            "error": str(e)
        }
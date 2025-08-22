from fastapi import APIRouter, HTTPException, UploadFile, File, Form, Depends
from fastapi.responses import JSONResponse, FileResponse
from typing import Dict, Any, Optional, List
import json
import logging
from pathlib import Path
import tempfile
import os
from datetime import datetime

from services.enhanced_document_service import EnhancedDocumentService
from models.request_models import DocumentProcessingRequest
from pydantic import BaseModel

logger = logging.getLogger(__name__)

# Router untuk document manipulation
router = APIRouter(tags=["document-manipulation"])

# Global service instance - will be injected from main app
_enhanced_document_service: Optional[EnhancedDocumentService] = None

def get_enhanced_document_service() -> EnhancedDocumentService:
    """Dependency to get enhanced document service"""
    if _enhanced_document_service is None:
        raise HTTPException(status_code=500, detail="Enhanced document service not initialized")
    return _enhanced_document_service

def set_enhanced_document_service(service: EnhancedDocumentService):
    """Set the enhanced document service instance"""
    global _enhanced_document_service
    _enhanced_document_service = service

# Request models
class DocumentManipulationRequest(BaseModel):
    file_id: str
    settings: Dict[str, Any]
    preview_only: bool = False
    
    class Config:
        schema_extra = {
            "example": {
                "file_id": "doc_123456",
                "settings": {
                    "color_mode": "black_white",
                    "convert_to_bw": True,
                    "page_range": "1-5",
                    "brightness": 10,
                    "contrast": 5,
                    "paper_size": "A4",
                    "orientation": "portrait",
                    "margin_top": 20,
                    "margin_bottom": 20,
                    "margin_left": 20,
                    "margin_right": 20,
                    "fit_to_page": True,
                    "center_horizontally": True,
                    "center_vertically": True
                },
                "preview_only": False
            }
        }

class PDFSplitRequest(BaseModel):
    file_id: str
    split_type: str = "pages"  # 'pages', 'range'
    split_value: int = 1  # pages per file
    ranges: Optional[List[str]] = None  # for range split: ["1-3", "4-6"]
    
    class Config:
        schema_extra = {
            "example": {
                "file_id": "doc_123456",
                "split_type": "pages",
                "split_value": 2,
                "ranges": ["1-3", "4-6", "7-10"]
            }
        }

class DocumentConversionRequest(BaseModel):
    file_id: str
    target_format: str = "pdf"
    conversion_settings: Optional[Dict[str, Any]] = None
    
    class Config:
        schema_extra = {
            "example": {
                "file_id": "doc_123456",
                "target_format": "pdf",
                "conversion_settings": {
                    "quality": "high",
                    "preserve_formatting": True
                }
            }
        }

class PreviewRequest(BaseModel):
    file_id: str
    page_number: int = 1
    settings: Optional[Dict[str, Any]] = None
    
    class Config:
        schema_extra = {
            "example": {
                "file_id": "doc_123456",
                "page_number": 1,
                "settings": {
                    "color_mode": "grayscale",
                    "brightness": 0,
                    "contrast": 0
                }
            }
        }

# Helper function to get file path from file_id
def get_file_path_from_id(file_id: str) -> str:
    """Get actual file path from file_id"""
    # This should integrate with your existing file service
    # For now, assuming files are stored in uploads directory
    uploads_dir = Path("uploads")
    
    # Try to find file with this ID
    for file_path in uploads_dir.glob(f"*{file_id}*"):
        if file_path.is_file():
            return str(file_path)
    
    # If not found, try temp directory
    temp_dir = Path("temp")
    for file_path in temp_dir.glob(f"*{file_id}*"):
        if file_path.is_file():
            return str(file_path)
    
    raise HTTPException(status_code=404, detail=f"File with ID {file_id} not found")

@router.post("/manipulate")
async def manipulate_document(
    request: DocumentManipulationRequest,
    service: EnhancedDocumentService = Depends(get_enhanced_document_service)
):
    """Manipulasi dokumen berdasarkan pengaturan yang diberikan"""
    try:
        logger.info(f"Processing document manipulation request for file_id: {request.file_id}")
        
        # Get file path
        file_path = get_file_path_from_id(request.file_id)
        
        # Process document with enhanced service
        result = service.process_document_with_manipulation(
            input_file_path=file_path,
            settings=request.settings
        )
        
        if not result['success']:
            raise HTTPException(status_code=500, detail=result.get('error', 'Processing failed'))
        
        # If preview only, generate preview data
        if request.preview_only:
            preview_data = service.get_preview_data(result['output_path'])
            result['preview'] = preview_data
        
        return JSONResponse(content={
            "success": True,
            "message": "Document manipulated successfully",
            "data": result
        })
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error in document manipulation: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@router.post("/split")
async def split_pdf_document(
    request: PDFSplitRequest,
    service: EnhancedDocumentService = Depends(get_enhanced_document_service)
):
    """Split PDF dokumen menjadi beberapa file"""
    try:
        logger.info(f"Processing PDF split request for file_id: {request.file_id}")
        
        # Get file path
        file_path = get_file_path_from_id(request.file_id)
        
        # Prepare split settings
        split_settings = {
            'split_type': request.split_type,
            'split_value': request.split_value
        }
        
        if request.ranges:
            split_settings['ranges'] = request.ranges
        
        # Split PDF
        split_results = service.split_pdf(file_path, split_settings)
        
        if not split_results:
            raise HTTPException(status_code=500, detail="Failed to split PDF")
        
        return JSONResponse(content={
            "success": True,
            "message": f"PDF split into {len(split_results)} files",
            "data": {
                "original_file": request.file_id,
                "split_files": split_results,
                "total_files": len(split_results)
            }
        })
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error in PDF splitting: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@router.post("/convert")
async def convert_document(
    request: DocumentConversionRequest,
    service: EnhancedDocumentService = Depends(get_enhanced_document_service)
):
    """Konversi dokumen ke format yang diinginkan"""
    try:
        logger.info(f"Processing document conversion request for file_id: {request.file_id}")
        
        # Get file path
        file_path = get_file_path_from_id(request.file_id)
        
        # Prepare conversion settings
        conversion_settings = request.conversion_settings or {}
        conversion_settings['target_format'] = request.target_format
        
        # Process document (conversion is part of manipulation)
        result = service.process_document_with_manipulation(
            input_file_path=file_path,
            settings=conversion_settings
        )
        
        if not result['success']:
            raise HTTPException(status_code=500, detail=result.get('error', 'Conversion failed'))
        
        return JSONResponse(content={
            "success": True,
            "message": f"Document converted to {request.target_format} successfully",
            "data": result
        })
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error in document conversion: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@router.post("/preview")
async def get_document_preview(
    request: PreviewRequest,
    service: EnhancedDocumentService = Depends(get_enhanced_document_service)
):
    """Dapatkan preview dokumen dengan pengaturan tertentu"""
    try:
        logger.info(f"Generating preview for file_id: {request.file_id}, page: {request.page_number}")
        
        # Get file path
        file_path = get_file_path_from_id(request.file_id)
        
        # If settings provided, apply them first
        if request.settings:
            # Create temporary processed version
            temp_result = service.process_document_with_manipulation(
                input_file_path=file_path,
                settings=request.settings
            )
            
            if temp_result['success']:
                file_path = temp_result['output_path']
        
        # Generate preview
        preview_data = service.get_preview_data(
            pdf_path=file_path,
            page_num=request.page_number - 1  # Convert to 0-based
        )
        
        if not preview_data['success']:
            raise HTTPException(status_code=500, detail=preview_data.get('error', 'Preview generation failed'))
        
        return JSONResponse(content={
            "success": True,
            "message": "Preview generated successfully",
            "data": preview_data
        })
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error generating preview: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@router.post("/preview-processed")
async def get_processed_document_preview(
    request: DocumentManipulationRequest,
    service: EnhancedDocumentService = Depends(get_enhanced_document_service)
):
    """Generate preview of processed document without saving permanently"""
    try:
        logger.info(f"Generating processed preview for file_id: {request.file_id}")
        
        # Get file path
        file_path = get_file_path_from_id(request.file_id)
        
        # Process document temporarily for preview
        result = service.process_document_with_manipulation(
            input_file_path=file_path,
            settings=request.settings
        )
        
        if not result['success']:
            raise HTTPException(status_code=500, detail=result.get('error', 'Processing failed'))
        
        # Generate preview data
        preview_data = service.get_preview_data(result['output_path'])
        
        # Create response with processed file info
        response_data = {
            "success": True,
            "message": "Processed preview generated successfully",
            "processed_file_id": result['output_filename'].replace('.pdf', ''),
            "preview_url": f"/api/files/{result['output_filename'].replace('.pdf', '')}/preview",
            "preview_data": preview_data,
            "processing_info": {
                "original_format": result.get('original_format'),
                "settings_applied": result.get('settings_applied'),
                "pages_count": result.get('pages_count'),
                "file_size": result.get('file_size')
            }
        }
        
        return JSONResponse(content=response_data)
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error generating processed preview: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@router.get("/download/{file_id}")
async def download_processed_file(file_id: str):
    """Download file yang sudah diproses"""
    try:
        # Get file path
        file_path = get_file_path_from_id(file_id)
        
        if not os.path.exists(file_path):
            raise HTTPException(status_code=404, detail="File not found")
        
        # Get filename
        filename = Path(file_path).name
        
        return FileResponse(
            path=file_path,
            filename=filename,
            media_type='application/pdf'
        )
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error downloading file: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@router.post("/upload-and-process")
async def upload_and_process_document(
    file: UploadFile = File(...),
    settings: str = Form(...),
    service: EnhancedDocumentService = Depends(get_enhanced_document_service)
):
    """Upload dan langsung proses dokumen"""
    try:
        # Parse settings
        try:
            settings_dict = json.loads(settings)
        except json.JSONDecodeError:
            raise HTTPException(status_code=400, detail="Invalid settings JSON")
        
        # Save uploaded file temporarily
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        temp_filename = f"{timestamp}_{file.filename}"
        temp_path = Path("temp") / temp_filename
        temp_path.parent.mkdir(exist_ok=True)
        
        # Write file
        with open(temp_path, "wb") as buffer:
            content = await file.read()
            buffer.write(content)
        
        logger.info(f"Uploaded file saved to: {temp_path}")
        
        # Process document
        result = service.process_document_with_manipulation(
            input_file_path=str(temp_path),
            settings=settings_dict
        )
        
        if not result['success']:
            # Cleanup temp file
            try:
                os.remove(temp_path)
            except:
                pass
            raise HTTPException(status_code=500, detail=result.get('error', 'Processing failed'))
        
        # Generate preview
        preview_data = service.get_preview_data(result['output_path'])
        result['preview'] = preview_data
        
        # Cleanup original temp file
        try:
            os.remove(temp_path)
        except:
            pass
        
        return JSONResponse(content={
            "success": True,
            "message": "File uploaded and processed successfully",
            "data": result
        })
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error in upload and process: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@router.get("/supported-formats")
async def get_supported_formats():
    """Dapatkan daftar format yang didukung"""
    return JSONResponse(content={
        "success": True,
        "data": {
            "input_formats": {
                "pdf": [".pdf"],
                "word": [".docx", ".doc"],
                "excel": [".xlsx", ".xls", ".csv"],
                "image": [".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".gif"]
            },
            "output_formats": ["pdf"],
            "manipulation_features": [
                "color_conversion",
                "black_white_conversion",
                "grayscale_conversion",
                "brightness_adjustment",
                "contrast_adjustment",
                "page_splitting",
                "page_range_selection",
                "auto_rotation",
                "format_conversion"
            ],
            "paper_sizes": ["A4", "A3", "A5", "Letter", "Legal"],
            "orientations": ["portrait", "landscape"]
        }
    })

@router.delete("/cleanup")
async def cleanup_temp_files(
    older_than_hours: int = 24,
    service: EnhancedDocumentService = Depends(get_enhanced_document_service)
):
    """Bersihkan file temporary yang lama"""
    try:
        cleaned_count = service.cleanup_temp_files(older_than_hours)
        
        return JSONResponse(content={
            "success": True,
            "message": f"Cleaned up {cleaned_count} temporary files",
            "data": {
                "cleaned_files": cleaned_count,
                "older_than_hours": older_than_hours
            }
        })
        
    except Exception as e:
        logger.error(f"Error cleaning up files: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@router.get("/health")
async def health_check():
    """Health check untuk document manipulation service"""
    try:
        # Test basic functionality
        temp_dir = Path("temp")
        temp_dir.mkdir(exist_ok=True)
        
        return JSONResponse(content={
            "success": True,
            "message": "Document manipulation service is healthy",
            "data": {
                "service_status": "active",
                "temp_directory": str(temp_dir),
                "temp_directory_exists": temp_dir.exists(),
                "timestamp": datetime.now().isoformat()
            }
        })
        
    except Exception as e:
        logger.error(f"Health check failed: {e}")
        raise HTTPException(status_code=500, detail=str(e))
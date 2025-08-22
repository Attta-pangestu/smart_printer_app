from typing import Optional, Dict, Any
from pydantic import BaseModel, Field
from .job import PrintSettings

class DocumentProcessingRequest(BaseModel):
    """Request model untuk document processing dengan printing"""
    file_id: str = Field(..., description="ID file yang akan diproses dan dicetak")
    printer_id: str = Field(..., description="ID printer yang akan digunakan")
    print_settings: Optional[Dict[str, Any]] = Field(default_factory=dict, description="Pengaturan printing")
    document_settings: Optional[Dict[str, Any]] = Field(default_factory=dict, description="Pengaturan document processing")
    user: Optional[str] = Field(default="anonymous", description="User yang melakukan request")
    
    class Config:
        json_schema_extra = {
            "example": {
                "file_id": "test_document.pdf",
                "printer_id": "epson_l120_series",
                "print_settings": {
                    "copies": 1,
                    "color_mode": "color",
                    "paper_size": "A4",
                    "orientation": "portrait",
                    "quality": "normal",
                    "duplex": "none",
                    "scale": 100,
                    "margins": {
                        "top": 1.0,
                        "bottom": 1.0,
                        "left": 1.0,
                        "right": 1.0
                    }
                },
                "document_settings": {
                    "apply_watermark": False,
                    "adjust_brightness": 0,
                    "apply_filters": []
                },
                "user": "test_user"
            }
        }

class PrintJobRequest(BaseModel):
    """Request model untuk print job biasa"""
    printer_id: str = Field(..., description="ID printer yang akan digunakan")
    file_id: str = Field(..., description="ID file yang akan dicetak")
    settings: Optional[Dict[str, Any]] = Field(default_factory=dict, description="Pengaturan printing")
    user: Optional[str] = Field(default="anonymous", description="User yang melakukan request")
    
    class Config:
        json_schema_extra = {
            "example": {
                "printer_id": "epson_l120_series",
                "file_id": "test_document.pdf",
                "settings": {
                    "copies": 1,
                    "color_mode": "color",
                    "paper_size": "A4",
                    "orientation": "portrait",
                    "quality": "normal",
                    "duplex": "none"
                },
                "user": "test_user"
            }
        }
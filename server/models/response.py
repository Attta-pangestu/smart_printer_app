from typing import Optional, Any, List, Dict
from pydantic import BaseModel, Field
from datetime import datetime


class APIResponse(BaseModel):
    """Standard API response format"""
    success: bool = True
    message: str = "OK"
    data: Optional[Any] = None
    error: Optional[str] = None
    timestamp: datetime = Field(default_factory=datetime.now)
    
    class Config:
        json_encoders = {
            datetime: lambda v: v.isoformat()
        }
    
    @classmethod
    def success_response(cls, data: Any = None, message: str = "OK") -> "APIResponse":
        """Create success response"""
        return cls(success=True, message=message, data=data)
    
    @classmethod
    def error_response(cls, error: str, message: str = "Error") -> "APIResponse":
        """Create error response"""
        return cls(success=False, message=message, error=error)


class PaginatedResponse(BaseModel):
    """Paginated response format"""
    items: List[Any]
    total: int
    page: int = 1
    per_page: int = 10
    pages: int
    has_next: bool
    has_prev: bool
    
    @classmethod
    def create(cls, items: List[Any], total: int, page: int = 1, per_page: int = 10) -> "PaginatedResponse":
        """Create paginated response"""
        pages = (total + per_page - 1) // per_page
        return cls(
            items=items,
            total=total,
            page=page,
            per_page=per_page,
            pages=pages,
            has_next=page < pages,
            has_prev=page > 1
        )


class StatusResponse(BaseModel):
    """System status response"""
    server_status: str = "running"
    uptime: float  # seconds
    version: str = "1.0.0"
    printers_count: int = 0
    active_jobs: int = 0
    total_jobs: int = 0
    memory_usage: Dict[str, float] = Field(default_factory=dict)
    disk_usage: Dict[str, float] = Field(default_factory=dict)
    

class PrinterStatusResponse(BaseModel):
    """Printer status response"""
    printer_id: str
    status: str
    jobs_in_queue: int
    last_job_time: Optional[datetime] = None
    error_message: Optional[str] = None
    
    class Config:
        json_encoders = {
            datetime: lambda v: v.isoformat() if v else None
        }


class FileUploadResponse(BaseModel):
    """File upload response"""
    file_id: str
    file_name: str
    file_size: int
    file_type: str
    upload_path: str
    pages_detected: Optional[int] = None
    

class JobSubmissionResponse(BaseModel):
    """Job submission response"""
    job_id: str
    status: str
    estimated_pages: int
    estimated_duration: Optional[float] = None  # seconds
    queue_position: int
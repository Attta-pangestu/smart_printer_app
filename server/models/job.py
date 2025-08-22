from enum import Enum
from typing import Optional, Dict, Any
from pydantic import BaseModel, Field
from datetime import datetime
import uuid


class JobStatus(str, Enum):
    """Status print job"""
    PENDING = "pending"
    PROCESSING = "processing"
    PRINTING = "printing"
    COMPLETED = "completed"
    FAILED = "failed"
    CANCELLED = "cancelled"
    PAUSED = "paused"


class ColorMode(str, Enum):
    """Mode warna printing"""
    COLOR = "color"
    GRAYSCALE = "grayscale"
    BLACK_WHITE = "bw"


class PaperSize(str, Enum):
    """Ukuran kertas"""
    A4 = "A4"
    A3 = "A3"
    A5 = "A5"
    LETTER = "Letter"
    LEGAL = "Legal"
    TABLOID = "Tabloid"
    CUSTOM = "Custom"


class Orientation(str, Enum):
    """Orientasi kertas"""
    PORTRAIT = "portrait"
    LANDSCAPE = "landscape"


class PrintQuality(str, Enum):
    """Kualitas print"""
    DRAFT = "draft"
    NORMAL = "normal"
    HIGH = "high"
    PHOTO = "photo"


class DuplexMode(str, Enum):
    """Mode duplex"""
    NONE = "none"
    HORIZONTAL = "horizontal"
    VERTICAL = "vertical"


class FitToPageMode(str, Enum):
    """Mode fit to page"""
    NONE = "none"
    FIT_TO_PAGE = "fit_to_page"
    FIT_TO_PAPER = "fit_to_paper"
    SHRINK_TO_FIT = "shrink_to_fit"
    ACTUAL_SIZE = "actual_size"


class PrintSettings(BaseModel):
    """Pengaturan printing"""
    color_mode: ColorMode = ColorMode.COLOR
    copies: int = Field(1, ge=1, le=999, description="Jumlah copy")
    paper_size: PaperSize = PaperSize.A4
    orientation: Orientation = Orientation.PORTRAIT
    quality: PrintQuality = PrintQuality.NORMAL
    duplex: DuplexMode = DuplexMode.NONE
    
    # Advanced settings
    scale: int = Field(100, ge=25, le=400, description="Scale percentage")
    margins: Dict[str, float] = Field(
        default_factory=lambda: {"top": 0.5, "bottom": 0.5, "left": 0.5, "right": 0.5}
    )
    collate: bool = True
    reverse_order: bool = False
    
    # Page range
    page_range: Optional[str] = Field(None, description="e.g., '1-5,8,11-13'")
    pages_per_sheet: int = Field(1, ge=1, le=16, description="Pages per sheet")
    
    # Fit to page settings
    fit_to_page: FitToPageMode = FitToPageMode.NONE
    
    # Split PDF settings
    split_pdf: bool = Field(False, description="Split PDF into separate print jobs per page")
    split_page_range: Optional[str] = Field(None, description="Page range for splitting, e.g., '1-5,8,11-13'")
    split_output_prefix: str = Field("page_", description="Prefix for split page files")
    
    # Custom paper size (if paper_size is CUSTOM)
    custom_paper: Optional[Dict[str, float]] = Field(None, description="Custom paper size with width and height in inches")
    
    # Header/Footer settings
    header_footer: Optional[Dict[str, Any]] = Field(
        default_factory=lambda: {
            "enabled": False,
            "header_left": "",
            "header_center": "",
            "header_right": "",
            "footer_left": "",
            "footer_center": "",
            "footer_right": ""
        },
        description="Header and footer settings"
    )
    
    # Page break settings
    page_breaks: Optional[Dict[str, Any]] = Field(
        default_factory=lambda: {
            "avoid_page_breaks": False,
            "insert_manual_breaks": False,
            "break_positions": ""
        },
        description="Page break control settings"
    )
    
    class Config:
        use_enum_values = True


class PrintJob(BaseModel):
    """Model print job"""
    id: str = Field(default_factory=lambda: str(uuid.uuid4()))
    printer_id: str = Field(..., description="ID printer yang digunakan")
    file_name: str = Field(..., description="Nama file yang akan diprint")
    file_path: str = Field(..., description="Path file")
    file_size: int = Field(..., description="Ukuran file dalam bytes")
    file_type: str = Field(..., description="Tipe file (pdf, doc, txt, etc.)")
    
    # Job info
    title: str = Field(..., description="Judul job")
    user: str = Field("anonymous", description="User yang submit job")
    client_ip: Optional[str] = Field(None, description="IP address client")
    
    # Print settings
    settings: PrintSettings = Field(default_factory=PrintSettings)
    
    # Status and timing
    status: JobStatus = JobStatus.PENDING
    created_at: datetime = Field(default_factory=datetime.now)
    started_at: Optional[datetime] = None
    completed_at: Optional[datetime] = None
    
    # Progress
    total_pages: int = 0
    pages_printed: int = 0
    progress_percentage: float = 0.0
    
    # Error handling
    error_message: Optional[str] = None
    retry_count: int = 0
    max_retries: int = 3
    
    # Metadata
    metadata: Dict[str, Any] = Field(default_factory=dict)
    
    class Config:
        json_encoders = {
            datetime: lambda v: v.isoformat() if v else None
        }
    
    @property
    def duration(self) -> Optional[float]:
        """Durasi job dalam detik"""
        if self.started_at and self.completed_at:
            return (self.completed_at - self.started_at).total_seconds()
        return None
    
    @property
    def is_active(self) -> bool:
        """Apakah job sedang aktif"""
        return self.status in [JobStatus.PENDING, JobStatus.PROCESSING, JobStatus.PRINTING]
    
    @property
    def is_finished(self) -> bool:
        """Apakah job sudah selesai"""
        return self.status in [JobStatus.COMPLETED, JobStatus.FAILED, JobStatus.CANCELLED]


class JobSummary(BaseModel):
    """Ringkasan job untuk list view"""
    id: str
    title: str
    printer_id: str
    status: JobStatus
    progress_percentage: float
    created_at: datetime
    user: str
    total_pages: int
    
    class Config:
        json_encoders = {
            datetime: lambda v: v.isoformat()
        }
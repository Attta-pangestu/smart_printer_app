from enum import Enum
from typing import Optional, List, Dict, Any
from pydantic import BaseModel, Field
from datetime import datetime


class PrinterStatus(str, Enum):
    """Status printer"""
    ONLINE = "online"
    OFFLINE = "offline"
    BUSY = "busy"
    ERROR = "error"
    PAUSED = "paused"


class PrinterCapability(BaseModel):
    """Kemampuan printer"""
    supports_color: bool = False
    supports_duplex: bool = False
    max_paper_size: str = "A4"
    supported_paper_sizes: List[str] = Field(default_factory=lambda: ["A4", "A3", "Letter"])
    supported_qualities: List[str] = Field(default_factory=lambda: ["draft", "normal", "high"])
    max_copies: int = 999
    supported_orientations: List[str] = Field(default_factory=lambda: ["portrait", "landscape"])


class Printer(BaseModel):
    """Model printer"""
    id: str = Field(..., description="Unique printer ID")
    name: str = Field(..., description="Printer name")
    driver_name: str = Field(..., description="Driver name")
    port_name: str = Field(..., description="Port name")
    server_name: Optional[str] = Field(None, description="Server name")
    share_name: Optional[str] = Field(None, description="Share name")
    location: Optional[str] = Field(None, description="Printer location")
    comment: Optional[str] = Field(None, description="Printer comment")
    status: PrinterStatus = PrinterStatus.OFFLINE
    is_default: bool = False
    is_shared: bool = False
    capabilities: PrinterCapability = Field(default_factory=PrinterCapability)
    
    # Network info
    ip_address: Optional[str] = Field(None, description="IP address if network printer")
    mac_address: Optional[str] = Field(None, description="MAC address")
    
    # Status info
    jobs_count: int = 0
    last_seen: datetime = Field(default_factory=datetime.now)
    error_message: Optional[str] = None
    
    # Statistics
    total_pages_printed: int = 0
    total_jobs_completed: int = 0
    
    class Config:
        json_encoders = {
            datetime: lambda v: v.isoformat()
        }


class PrinterInfo(BaseModel):
    """Informasi printer yang lebih ringkas"""
    id: str
    name: str
    status: PrinterStatus
    is_default: bool
    jobs_count: int
    capabilities: PrinterCapability


class PrinterDiscovery(BaseModel):
    """Model untuk discovery printer di network"""
    service_name: str
    ip_address: str
    port: int
    printer_name: str
    capabilities: Dict[str, Any] = Field(default_factory=dict)
    discovered_at: datetime = Field(default_factory=datetime.now)
    
    class Config:
        json_encoders = {
            datetime: lambda v: v.isoformat()
        }
from fastapi import FastAPI, HTTPException, UploadFile, File, Form, Depends
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
import uvicorn
import yaml
import os
import logging
import tempfile
import threading
from datetime import datetime
from typing import List, Optional, Dict, Any
from pathlib import Path
from docx2pdf import convert

from services import PrinterService, JobService, DiscoveryService, FileService, DocumentService
from services.enhanced_document_service import EnhancedDocumentService
from api.document_manipulation import router as document_router
from models.printer import Printer, PrinterInfo, PrinterDiscovery
from models.job import PrintJob, PrintSettings, JobStatus
from models.response import (
    APIResponse, PaginatedResponse, StatusResponse,
    PrinterStatusResponse, FileUploadResponse, JobSubmissionResponse
)
from models.request_models import DocumentProcessingRequest, PrintJobRequest

# Configure logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

class PrintServerApp:
    """Main application class for Epson L120 Print Server"""
    
    def __init__(self):
        self.config = self._load_config()
        self.app = self._create_app()
        self._setup_middleware()
        self._setup_static_files()
        self._initialize_services()
        self._setup_routes()
    
    def _load_config(self) -> dict:
        """Load configuration from YAML file"""
        config_path = Path(__file__).parent.parent / "config.yaml"
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                return yaml.safe_load(f)
        except Exception as e:
            logger.error(f"Failed to load config: {e}")
            return self._get_default_config()
    
    def _get_default_config(self) -> dict:
        """Return default configuration"""
        return {
            'server': {
                'host': '0.0.0.0',
                'port': 8082,
                'debug': False,
                'cors_origins': ['*']
            },
            'security': {
                'enable_auth': False,
                'api_key': 'your-api-key'
            },
            'storage': {
                'temp_dir': './temp'
            },
            'printer': {
                'discovery': {
                    'enabled': True
                }
            }
        }
    
    def _create_app(self) -> FastAPI:
        """Create FastAPI application instance"""
        return FastAPI(
            title="Epson L120 Print Server",
            description="Rebinmas Remote Print Server for Epson L120 printer management",
            version="2.0.0",
            docs_url="/docs" if self.config['server']['debug'] else None,
            redoc_url="/redoc" if self.config['server']['debug'] else None
        )
    
    def _setup_middleware(self):
        """Configure middleware"""
        self.app.add_middleware(
            CORSMiddleware,
            allow_origins=self.config['server']['cors_origins'],
            allow_credentials=True,
            allow_methods=["*"],
            allow_headers=["*"],
        )
    
    def _setup_static_files(self):
        """Mount static file directories"""
        # Mount static files for web interface - use server/static directory
        static_dir = Path(__file__).parent / "static"
        if static_dir.exists():
            self.app.mount("/static", StaticFiles(directory=str(static_dir)), name="static")
        
        # Legacy static directory support
        legacy_static_dir = Path(__file__).parent.parent / "static"
        if legacy_static_dir.exists():
            self.app.mount("/legacy", StaticFiles(directory=str(legacy_static_dir)), name="legacy")
        
        # Legacy web directory support
        web_dir = Path(__file__).parent.parent / "web"
        if web_dir.exists():
            self.app.mount("/web", StaticFiles(directory=str(web_dir)), name="web")
    
    def _initialize_services(self):
        """Initialize all services"""
        self.security = HTTPBearer() if self.config['security']['enable_auth'] else None
        
        # Initialize core services
        self.printer_service = PrinterService()
        self.file_service = FileService(
            upload_dir=self.config['storage']['temp_dir'],
            temp_dir=self.config['storage']['temp_dir']
        )
        self.document_service = DocumentService()
        self.enhanced_document_service = EnhancedDocumentService()
        self.job_service = JobService(self.printer_service, self.file_service)
        self.discovery_service = DiscoveryService(
            port=self.config['server']['port']
        )
    
    def _setup_routes(self):
        """Setup all application routes"""
        self._setup_web_routes()
        self._setup_api_routes()
        self._setup_event_handlers()
    
    def _setup_web_routes(self):
        """Setup web interface routes"""
        
        @self.app.get("/", response_class=HTMLResponse)
        async def root():
            """Serve main dashboard"""
            # Try server/static first (correct location)
            index_file = Path(__file__).parent / "static" / "index.html"
            if index_file.exists():
                return FileResponse(index_file)
            
            # Fallback to legacy location
            legacy_index_file = Path(__file__).parent.parent / "static" / "index.html"
            if legacy_index_file.exists():
                return FileResponse(legacy_index_file)
            
            return HTMLResponse("""
            <html>
                <head><title>Epson L120 Print Server</title></head>
                <body>
                    <h1>Epson L120 Print Server</h1>
                    <p>Server is running. Dashboard not available.</p>
                    <p><a href="/docs">API Documentation</a></p>
                </body>
            </html>
            """)
        
        @self.app.get("/config", response_class=HTMLResponse)
        async def config_page():
            """Serve configuration page"""
            # Try server/static first
            config_file = Path(__file__).parent / "static" / "config.html"
            if config_file.exists():
                return FileResponse(config_file)
            
            # Fallback to legacy location
            legacy_config_file = Path(__file__).parent.parent / "static" / "config.html"
            if legacy_config_file.exists():
                return FileResponse(legacy_config_file)
            
            raise HTTPException(status_code=404, detail="Configuration page not found")
    
    def _setup_api_routes(self):
        """Setup all API endpoints"""
        pass  # API endpoints will be setup after class instantiation
    
    def _setup_event_handlers(self):
        """Setup application event handlers"""
        
        @self.app.on_event("startup")
        async def startup_event():
            """Handle application startup"""
            self.start_time = datetime.now()
            logger.info("Print server started successfully")
            
            # Start discovery service if enabled
            if self.config['printer']['discovery']['enabled']:
                self.discovery_service.start_broadcasting()
                self.discovery_service.start_discovery()
        
        @self.app.on_event("shutdown")
        async def shutdown_event():
            """Handle application shutdown"""
            logger.info("Shutting down print server...")
            
            # Stop services
            self.discovery_service.stop_broadcasting()
            self.discovery_service.stop_discovery()
            self.job_service.stop()
            
            # Cleanup temp files
            self.file_service.cleanup_temp_files()
            
            logger.info("Print server shutdown complete")

# Create application instance
server = PrintServerApp()
app = server.app
config = server.config

# Services (for backward compatibility)
printer_service = server.printer_service
file_service = server.file_service
document_service = server.document_service
job_service = server.job_service
discovery_service = server.discovery_service
security = server.security

def verify_auth():
    """Verify API authentication"""
    if not server.config['security']['enable_auth']:
        return True
    return True

def verify_auth_with_bearer(credentials: HTTPAuthorizationCredentials = Depends(security)):
    """Verify API authentication with bearer token"""
    if not credentials or credentials.credentials != server.config['security']['api_key']:
        raise HTTPException(status_code=401, detail="Invalid authentication")
    return True
    
def _setup_system_endpoints(server_instance):
    """Setup system-related API endpoints"""
    global server
    server = server_instance
    
    @server.app.get("/api/printers")
    async def get_printers():
        """Get available printers"""
        printers = server.printer_service.get_all_printers(force_refresh=True)
        return {"printers": [printer.__dict__ for printer in printers]}
    
    @server.app.get("/api/printers/{printer_id}/status")
    async def get_printer_status(printer_id: str):
        """Get printer status"""
        printer = server.printer_service.get_printer(printer_id)
        if not printer:
            raise HTTPException(status_code=404, detail="Printer not found")
        
        status = server.printer_service.get_printer_status(printer_id)
        
        # Convert enum to string and create response object
        status_map = {
            "ONLINE": {"status": "online", "message": "Printer is ready"},
            "OFFLINE": {"status": "offline", "message": "Printer is offline"},
            "BUSY": {"status": "busy", "message": "Printer is busy"},
            "PAUSED": {"status": "paused", "message": "Printer is paused"},
            "ERROR": {"status": "error", "message": "Printer has an error"}
        }
        
        status_str = status.name if hasattr(status, 'name') else str(status)
        return status_map.get(status_str, {"status": "unknown", "message": "Unknown status"})
    
    @server.app.get("/api/printers/{printer_id}/status/detailed")
    async def get_detailed_printer_status(printer_id: str):
        """Get detailed printer status for real-time monitoring"""
        printer = server.printer_service.get_printer(printer_id)
        if not printer:
            raise HTTPException(status_code=404, detail="Printer not found")
        
        detailed_status = server.printer_service.get_detailed_printer_status(printer_id)
        
        # Convert enum to string for JSON serialization
        if hasattr(detailed_status['status'], 'name'):
            detailed_status['status'] = detailed_status['status'].name
        
        return detailed_status
    
    @server.app.post("/api/files/upload")
    async def upload_file(file: UploadFile = File(...)):
        """Upload a file for printing"""
        try:
            # Validate file
            if not file.filename:
                raise HTTPException(status_code=422, detail="No filename provided")
            
            # Read file content
            file_content = await file.read()
            
            if not file_content:
                raise HTTPException(status_code=422, detail="Empty file uploaded")
            
            # Save file using FileService
            file_info = server.file_service.save_uploaded_file(file_content, file.filename)
            
            # Construct preview URL
            preview_url = f"/api/files/{file_info['name']}/preview"

            return {
                "file_id": file_info["name"],  # Use filename as file_id
                "file_name": file_info["name"],
                "filename": file_info["name"],
                "file_size": file_info["size"],
                "size": file_info["size"],
                "file_type": file_info["type"],
                "type": file_info["type"],
                "upload_path": file_info["path"],
                "path": file_info["path"],
                "pages_detected": file_info.get("pages", None),
                "preview_url": preview_url
            }
        except HTTPException:
            raise
        except ValueError as e:
            raise HTTPException(status_code=422, detail=str(e))
        except Exception as e:
            logger.error(f"Upload error: {e}")
            raise HTTPException(status_code=500, detail="Internal server error during file upload")
    
    @server.app.get("/api/files/{file_id}/preview")
    async def get_file_preview(file_id: str):
        """Get preview URL for uploaded file"""
        try:
            # Validate file exists
            file_path = server.file_service.upload_dir / file_id
            if not file_path.exists():
                raise HTTPException(status_code=404, detail="File not found")

            # Check file type and return appropriate response
            if file_id.lower().endswith('.pdf'):
                return FileResponse(
                    path=str(file_path),
                    media_type='application/pdf',
                    headers={"Content-Disposition": f"inline; filename={file_id}"}
                )
            elif file_id.lower().endswith('.docx'):
                pdf_path = file_path.with_suffix('.pdf')
                if not pdf_path.exists():
                    from docx2pdf import convert
                    convert(str(file_path), str(pdf_path))
                return FileResponse(
                    path=str(pdf_path),
                    media_type='application/pdf',
                    headers={"Content-Disposition": f"inline; filename={pdf_path.name}"}
                )
            else:
                # For other file types, return as octet-stream
                return FileResponse(
                    path=str(file_path),
                    media_type='application/octet-stream',
                    headers={"Content-Disposition": f"inline; filename={file_id}"}
                )
        except HTTPException:
            raise
        except Exception as e:
            logger.error(f"Preview error: {e}")
            raise HTTPException(status_code=500, detail="Internal server error during file preview")
    
    @server.app.post("/api/documents/process")
    async def process_document(request_data: dict):
        """Process document with layout settings"""
        try:
            file_id = request_data.get("file_id")
            settings = request_data.get("settings", {})
            
            if not file_id:
                raise HTTPException(status_code=400, detail="file_id is required")
            
            # Validate file exists
            file_path = server.file_service.upload_dir / file_id
            if not file_path.exists():
                raise HTTPException(status_code=404, detail="File not found")
            
            # Process document
            processed_file_info = server.document_service.process_document(
                str(file_path), settings
            )
            
            return {
                "success": True,
                "processed_file": processed_file_info,
                "original_file": file_id,
                "settings_applied": settings
            }
        except HTTPException:
            raise
        except Exception as e:
            logger.error(f"Document processing error: {e}")
            raise HTTPException(status_code=500, detail="Internal server error during document processing")
    
    @server.app.get("/api/documents/{file_id}/preview")
    async def get_processed_document_preview(file_id: str):
        """Get preview of processed document"""
        try:
            # Check if file exists in processed directory
            processed_path = server.document_service.output_dir / file_id
            if not processed_path.exists():
                raise HTTPException(status_code=404, detail="Processed file not found")
            
            return FileResponse(
                path=str(processed_path),
                media_type='application/pdf',
                headers={"Content-Disposition": f"inline; filename={file_id}"}
            )
        except HTTPException:
            raise
        except Exception as e:
            logger.error(f"Processed document preview error: {e}")
            raise HTTPException(status_code=500, detail="Internal server error during processed document preview")
    
    @server.app.get("/api/files/serve/{file_path:path}")
    async def serve_uploaded_file(file_path: str):
        """Serve an uploaded file for preview."""
        try:
            full_path = server.file_service.upload_dir / file_path
            if not full_path.exists() or not full_path.is_file():
                raise HTTPException(status_code=404, detail="File not found")
            
            # Determine media type for better browser handling
            media_type = 'application/octet-stream'
            if file_path.lower().endswith('.pdf'):
                media_type = 'application/pdf'
            elif file_path.lower().endswith(('.jpg', '.jpeg')):
                media_type = 'image/jpeg'
            elif file_path.lower().endswith('.png'):
                media_type = 'image/png'
            elif file_path.lower().endswith(('.txt', '.log')):
                media_type = 'text/plain'
            
            return FileResponse(
                path=str(full_path),
                media_type=media_type,
                headers={"Content-Disposition": f"inline; filename=\"{full_path.name}\""}
            )
        except HTTPException:
            raise
        except Exception as e:
            logger.error(f"File serving error for {file_path}: {e}")
            raise HTTPException(status_code=500, detail="Internal server error")
    
    @server.app.post("/api/files/convert-to-pdf")
    async def convert_to_pdf(request_data: dict):
        """Convert Excel or Word files to PDF for preview"""
        try:
            file_path = request_data.get("file_path")
            file_type = request_data.get("file_type")
            
            if not file_path or not file_type:
                raise HTTPException(status_code=400, detail="file_path and file_type are required")
            
            # Get full path
            full_path = server.file_service.upload_dir / file_path
            if not full_path.exists():
                raise HTTPException(status_code=404, detail="File not found")
            
            # Convert based on file type
            if file_type == "excel":
                # For Excel files, we'll use a simple conversion approach
                pdf_path = full_path.with_suffix('.pdf')
                
                try:
                    # Try to convert using pandas and matplotlib for basic Excel preview
                    import pandas as pd
                    import matplotlib.pyplot as plt
                    from matplotlib.backends.backend_pdf import PdfPages
                    
                    # Read Excel file
                    df = pd.read_excel(str(full_path))
                    
                    # Create PDF
                    with PdfPages(str(pdf_path)) as pdf:
                        fig, ax = plt.subplots(figsize=(11.69, 8.27))  # A4 size
                        ax.axis('tight')
                        ax.axis('off')
                        
                        # Create table
                        table = ax.table(cellText=df.values, colLabels=df.columns, 
                                       cellLoc='center', loc='center')
                        table.auto_set_font_size(False)
                        table.set_fontsize(8)
                        table.scale(1, 1.5)
                        
                        pdf.savefig(fig, bbox_inches='tight')
                        plt.close()
                    
                    return {"pdf_path": pdf_path.name, "success": True}
                    
                except Exception as e:
                    logger.warning(f"Excel conversion failed: {e}")
                    # Fallback: create a simple text-based PDF
                    from reportlab.pdfgen import canvas
                    from reportlab.lib.pagesizes import A4
                    
                    c = canvas.Canvas(str(pdf_path), pagesize=A4)
                    c.drawString(100, 750, f"Excel File: {file_path}")
                    c.drawString(100, 720, "Preview not available - file ready for printing")
                    c.save()
                    
                    return {"pdf_path": pdf_path.name, "success": True}
                    
            elif file_type == "word":
                # For Word files, try to convert using docx2pdf
                pdf_path = full_path.with_suffix('.pdf')
                
                try:
                    from docx2pdf import convert
                    convert(str(full_path), str(pdf_path))
                    return {"pdf_path": pdf_path.name, "success": True}
                    
                except Exception as e:
                    logger.warning(f"Word conversion failed: {e}")
                    # Fallback: create a simple text-based PDF
                    from reportlab.pdfgen import canvas
                    from reportlab.lib.pagesizes import A4
                    
                    c = canvas.Canvas(str(pdf_path), pagesize=A4)
                    c.drawString(100, 750, f"Word Document: {file_path}")
                    c.drawString(100, 720, "Preview not available - file ready for printing")
                    c.save()
                    
                    return {"pdf_path": pdf_path.name, "success": True}
            
            else:
                raise HTTPException(status_code=400, detail="Unsupported file type for conversion")
                
        except HTTPException:
            raise
        except Exception as e:
            logger.error(f"Conversion error: {e}")
            raise HTTPException(status_code=500, detail=f"Conversion failed: {str(e)}")
    
    @server.app.post("/api/jobs/submit")
    async def submit_print_job(request_data: dict):
        """Submit a print job with full print settings"""
        try:
            printer_id = request_data.get("printer_id")
            file_id = request_data.get("file_id")
            settings = request_data.get("settings", {})
            
            logger.info(f"Received print job request - printer_id: '{printer_id}', file_id: '{file_id}'")
            
            if not printer_id or not file_id:
                raise HTTPException(status_code=400, detail="printer_id and file_id are required")
            
            # Convert file_id to file_path
            # file_id is actually the filename from upload (already includes timestamp)
            file_path = str(server.file_service.upload_dir / file_id)
            
            logger.info(f"Looking for file at path: {file_path}")
            
            # Check if file exists
            if not os.path.exists(file_path):
                logger.error(f"File not found at path: {file_path}")
                # List files in upload directory for debugging
                upload_files = list(server.file_service.upload_dir.glob("*"))
                logger.info(f"Available files in upload directory: {[f.name for f in upload_files]}")
                raise HTTPException(status_code=400, detail=f"File {file_id} not found")
            
            # Create PrintSettings object from the settings data
            print_settings = PrintSettings(
                copies=settings.get("copies", 1),
                color_mode=settings.get("color_mode", "color"),
                paper_size=settings.get("paper_size", "A4"),
                orientation=settings.get("orientation", "portrait"),
                quality=settings.get("quality", "normal"),
                duplex=settings.get("duplex", "none"),
                scale=settings.get("scale", 100),
                pages_per_sheet=settings.get("pages_per_sheet", 1),
                page_range=settings.get("page_range", ""),
                collate=settings.get("collate", True),
                reverse_order=settings.get("reverse_order", False),
                margins=settings.get("margins", {
                    "top": 0.5,
                    "bottom": 0.5,
                    "left": 0.5,
                    "right": 0.5
                }),
                custom_paper=settings.get("custom_paper"),
                header_footer=settings.get("header_footer", {
                    "enabled": False,
                    "header_left": "",
                    "header_center": "",
                    "header_right": "",
                    "footer_left": "",
                    "footer_center": "",
                    "footer_right": ""
                }),
                page_breaks=settings.get("page_breaks", {
                    "avoid_page_breaks": False,
                    "insert_manual_breaks": False,
                    "break_positions": ""
                }),
                # New advanced features
                fit_to_page=settings.get("fit_to_page", "actual_size"),
                split_pdf=settings.get("split_pdf", False),
                split_page_range=settings.get("split_page_range", ""),
                split_output_prefix=settings.get("split_output_prefix", "page")
            )
            
            job = server.job_service.submit_job(printer_id, file_path, print_settings)
            
            logger.info(f"Print job submitted successfully - job_id: {job.id}")
            return {
                "job_id": job.id, 
                "status": job.status.value if hasattr(job.status, 'value') else str(job.status),
                "message": "Print job submitted successfully"
            }
        except Exception as e:
            raise HTTPException(status_code=400, detail=str(e))
    
    @server.app.post("/api/jobs/test-print")
    async def submit_test_print(printer_id: Optional[str] = Form(None), user: str = Form("test_user")):
        """Submit a test print job"""
        try:
            job = server.job_service.submit_test_job(
                printer_id=printer_id,
                user=user,
                client_ip=None  # Could be extracted from request if needed
            )
            return {
                "job_id": job.id, 
                "status": job.status.value if hasattr(job.status, 'value') else str(job.status),
                "printer_id": job.printer_id,
                "message": f"Test print job submitted to printer {job.printer_id}"
            }
        except ValueError as e:
            raise HTTPException(status_code=400, detail=str(e))
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"Internal server error: {str(e)}")
    
    @server.app.get("/api/jobs")
    async def get_jobs():
        """Get all print jobs"""
        jobs = server.job_service.get_jobs()
        return {"jobs": [job.__dict__ for job in jobs]}
    
    @server.app.get("/api/jobs/{job_id}")
    async def get_job(job_id: str):
        """Get specific job details"""
        job = server.job_service.get_job(job_id)
        if not job:
            raise HTTPException(status_code=404, detail="Job not found")
        return job.__dict__
    
    @server.app.delete("/api/jobs/{job_id}")
    async def cancel_job(job_id: str):
        """Cancel a print job"""
        success = server.job_service.cancel_job(job_id)
        if not success:
            raise HTTPException(status_code=404, detail="Job not found")
        return {"message": "Job cancelled successfully"}
    
    @server.app.post("/api/print")
    async def print_document(file: UploadFile = File(...), printer_id: str = Form(...), copies: int = Form(1)):
        """Print a document directly"""
        try:
            # Read file content
            file_content = await file.read()
            
            # Save the uploaded file
            file_info = server.file_service.save_uploaded_file(file_content, file.filename)
            
            # Convert file_id to file_path
            file_path = str(server.file_service.upload_dir / file_info["name"])
            
            # Create PrintSettings object
            print_settings = PrintSettings(
                copies=copies,
                color_mode="color",
                paper_size="A4",
                orientation="portrait",
                quality="normal",
                duplex="none",
                scale=100,
                pages_per_sheet=1,
                page_range="",
                collate=True,
                reverse_order=False,
                margins={"top": 1.0, "bottom": 1.0, "left": 1.0, "right": 1.0},
                custom_paper=None,
                # New advanced features with default values
                fit_to_page="actual_size",
                split_pdf=False,
                split_page_range="",
                split_output_prefix="page"
            )
            
            # Submit print job
            job = server.job_service.submit_job(printer_id, file_path, print_settings)
            
            return {"job_id": job.id, "status": job.status.value if hasattr(job.status, 'value') else str(job.status), "message": "Print job submitted successfully"}
        except Exception as e:
            raise HTTPException(status_code=400, detail=str(e))
    
    @server.app.post("/api/print/test")
    async def print_test_page(request_data: dict):
        """Print a test page"""
        try:
            printer_id = request_data.get("printer_id")
            if not printer_id:
                raise HTTPException(status_code=400, detail="printer_id is required")
            
            job = server.job_service.submit_test_job(printer_id)
            return {"job_id": job.id, "status": job.status, "message": "Test page job submitted successfully"}
        except Exception as e:
            raise HTTPException(status_code=400, detail=str(e))
    
    @server.app.post("/api/print/with-processing")
    async def print_document_with_processing(request: DocumentProcessingRequest):
        """Print document dengan document processing"""
        try:
            file_id = request.file_id
            printer_id = request.printer_id
            print_settings_data = request.print_settings
            document_settings = request.document_settings
            user = request.user
            
            if not file_id or not printer_id:
                raise HTTPException(status_code=400, detail="file_id and printer_id are required")
            
            # Validate file exists
            file_path = server.file_service.upload_dir / file_id
            if not file_path.exists():
                raise HTTPException(status_code=404, detail="File not found")
            
            # Create PrintSettings object
            print_settings = PrintSettings(
                copies=print_settings_data.get("copies", 1),
                color_mode=print_settings_data.get("color_mode", "color"),
                paper_size=print_settings_data.get("paper_size", "A4"),
                orientation=print_settings_data.get("orientation", "portrait"),
                quality=print_settings_data.get("quality", "normal"),
                duplex=print_settings_data.get("duplex", "none"),
                scale=print_settings_data.get("scale", 100),
                pages_per_sheet=print_settings_data.get("pages_per_sheet", 1),
                page_range=print_settings_data.get("page_range", ""),
                collate=print_settings_data.get("collate", True),
                reverse_order=print_settings_data.get("reverse_order", False),
                margins=print_settings_data.get("margins", {"top": 1.0, "bottom": 1.0, "left": 1.0, "right": 1.0}),
                custom_paper=print_settings_data.get("custom_paper"),
                # New advanced features
                fit_to_page=print_settings_data.get("fit_to_page", "actual_size"),
                split_pdf=print_settings_data.get("split_pdf", False),
                split_page_range=print_settings_data.get("split_page_range", ""),
                split_output_prefix=print_settings_data.get("split_output_prefix", "page")
            )
            
            # Submit print job dengan document processing
            job = server.job_service.submit_job_with_processing(
                printer_id=printer_id,
                file_path=str(file_path),
                settings=print_settings,
                document_settings=document_settings,
                user=user
            )
            
            return {
                "job_id": job.id,
                "status": job.status.value if hasattr(job.status, 'value') else str(job.status),
                "message": "Print job with document processing submitted successfully",
                "document_settings_applied": document_settings
            }
        except HTTPException:
            raise
        except Exception as e:
            logger.error(f"Print with processing error: {e}")
            raise HTTPException(status_code=500, detail="Internal server error during print with processing")
    
    @server.app.get("/api/config/printers")
    async def get_printer_config():
        """Get printer configuration"""
        printers = server.printer_service.get_all_printers()
        return {"printers": [printer.__dict__ for printer in printers]}
    
    @server.app.post("/api/config/printers")
    async def update_printer_config(config_data: dict):
        """Update printer configuration"""
        try:
            # Update printer configuration
            server.printer_service.update_printer_config(config_data)
            return {"message": "Printer configuration updated successfully"}
        except Exception as e:
            raise HTTPException(status_code=400, detail=str(e))
    
    @server.app.post("/api/config/printers/reload")
    async def reload_printer_config():
        """Reload printer configuration"""
        try:
            server.printer_service.reload_config()
            return {"message": "Printer configuration reloaded successfully"}
        except Exception as e:
            raise HTTPException(status_code=500, detail=str(e))
    
    @server.app.get("/api/status")
    async def get_status():
        """Get system status"""
        uptime = datetime.now() - server.start_time
        
        # Get disk usage
        disk_usage = {
            "total": 0,
            "used": 0,
            "free": 0
        }
        
        try:
            import shutil
            total, used, free = shutil.disk_usage("/")
            disk_usage = {
                "total": total,
                "used": used,
                "free": free
            }
        except Exception:
            pass
        
        # Get memory usage
        memory_usage = {
            "total": 0,
            "used": 0,
            "available": 0
        }
        
        try:
            import psutil
            memory = psutil.virtual_memory()
            memory_usage = {
                "total": memory.total,
                "used": memory.used,
                "available": memory.available
            }
        except Exception:
            pass
        
        return {
            "status": "running",
            "uptime": str(uptime),
            "disk_usage": disk_usage,
            "memory_usage": memory_usage,
            "printer_count": len(server.printer_service.get_all_printers()),
            "active_jobs": len(server.job_service.get_jobs())
        }
    
    @server.app.get("/api/network")
    async def get_network_info():
        """Get network information"""
        try:
            import socket
            hostname = socket.gethostname()
            local_ip = socket.gethostbyname(hostname)
            
            return {
                "hostname": hostname,
                "local_ip": local_ip,
                "port": server.config['server']['port']
            }
        except Exception as e:
            return {
                "hostname": "unknown",
                "local_ip": "unknown",
                "port": server.config['server']['port'],
                "error": str(e)
            }
    
    @server.app.post("/api/printers/scan")
    async def scan_network_printers():
        """Scan network for available printers"""
        try:
            # Scan menggunakan discovery service
            discovered_printers = server.discovery_service.scan_network_printers()
            
            # Juga dapatkan printer yang sudah ditemukan via mDNS
            mdns_printers = server.discovery_service.get_discovered_printers()
            
            return {
                "success": True,
                "message": f"Found {len(discovered_printers)} network printers and {len(mdns_printers)} mDNS printers",
                "network_printers": discovered_printers,
                "mdns_printers": [printer.__dict__ for printer in mdns_printers],
                "total_found": len(discovered_printers) + len(mdns_printers)
            }
        except Exception as e:
            logger.error(f"Error scanning network printers: {e}")
            raise HTTPException(status_code=500, detail=f"Failed to scan network printers: {str(e)}")
    
    @server.app.post("/api/discovery/start")
    async def start_discovery():
        """Start printer discovery service"""
        try:
            server.discovery_service.start_discovery()
            server.discovery_service.start_broadcasting()
            return {
                "success": True,
                "message": "Discovery service started successfully"
            }
        except Exception as e:
            logger.error(f"Error starting discovery service: {e}")
            raise HTTPException(status_code=500, detail=f"Failed to start discovery service: {str(e)}")
    
    @server.app.post("/api/discovery/stop")
    async def stop_discovery():
        """Stop printer discovery service"""
        try:
            server.discovery_service.stop_discovery()
            server.discovery_service.stop_broadcasting()
            return {
                "success": True,
                "message": "Discovery service stopped successfully"
            }
        except Exception as e:
            logger.error(f"Error stopping discovery service: {e}")
            raise HTTPException(status_code=500, detail=f"Failed to stop discovery service: {str(e)}")
    
    @server.app.get("/api/discovery/status")
    async def get_discovery_status():
        """Get discovery service status"""
        try:
            network_info = server.discovery_service.get_network_info()
            discovered_printers = server.discovery_service.get_discovered_printers()
            
            return {
                "success": True,
                "network_info": network_info,
                "discovered_printers": [printer.__dict__ for printer in discovered_printers],
                "discovery_enabled": server.config['printer']['discovery']['enabled']
            }
        except Exception as e:
            logger.error(f"Error getting discovery status: {e}")
            raise HTTPException(status_code=500, detail=f"Failed to get discovery status: {str(e)}")
    
    @server.app.get("/health")
    async def health_check():
        """Health check endpoint"""
        return {"status": "healthy", "timestamp": datetime.now().isoformat()}
    
    # Include document manipulation router
    server.app.include_router(document_router, prefix="/api/documents", tags=["documents"])
    
    # Set the enhanced document service for the router
    from api.document_manipulation import set_enhanced_document_service
    set_enhanced_document_service(server.enhanced_document_service)

# Main execution
if __name__ == "__main__":
    try:
        # Initialize the application
        print_server = PrintServerApp()
        app = print_server.app
        
        # Setup API endpoints after initialization
        _setup_system_endpoints(print_server)
        
        # Run the server
        uvicorn.run(
            app,
            host=print_server.config['server']['host'],
            port=print_server.config['server']['port'],
            reload=False,  # Disable reload to avoid import string requirement
            log_level="info"
        )
    except KeyboardInterrupt:
        logger.info("Server stopped by user")
    except Exception as e:
        logger.error(f"Failed to start server: {e}")
        raise
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
from services.excel_pywin32_service import ExcelPyWin32Service
from services.enhanced_document_service import EnhancedDocumentService
from routes.excel_pywin32_routes import router as excel_pywin32_router, init_services
from api.document_manipulation import router as document_manipulation_router, set_enhanced_document_service
from models.printer import Printer, PrinterInfo, PrinterDiscovery
from models.job import PrintJob, PrintSettings, JobStatus, PrintMethod
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

# Security setup
security = HTTPBearer()

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
                'port': 8000,
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
        self.job_service = JobService(self.printer_service, self.file_service)
        self.discovery_service = DiscoveryService(
            port=self.config['server']['port']
        )
        
        # Initialize Excel PyWin32 service
        self.excel_pywin32_service = ExcelPyWin32Service(
            temp_dir=self.config['storage']['temp_dir']
        )
        
        # Initialize Enhanced Document service
        self.enhanced_document_service = EnhancedDocumentService(
            temp_dir=self.config['storage']['temp_dir']
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
        
        @self.app.get("/excel-test", response_class=HTMLResponse)
        async def excel_test_page():
            """Serve Excel test page"""
            test_file = Path(__file__).parent / "static" / "excel-test.html"
            if test_file.exists():
                return FileResponse(test_file)
            
            raise HTTPException(status_code=404, detail="Excel test page not found")
        
        @self.app.get("/excel-test-simple", response_class=HTMLResponse)
        async def excel_test_simple_page():
            """Serve simple Excel test page"""
            test_file = Path(__file__).parent / "static" / "excel-test-simple.html"
            if test_file.exists():
                return FileResponse(test_file)
            
            raise HTTPException(status_code=404, detail="Simple Excel test page not found")
        
        @self.app.get("/excel-fix-test", response_class=HTMLResponse)
        async def excel_fix_test_page():
            """Serve Excel fix test page"""
            test_file = Path(__file__).parent / "static" / "excel-fix-test.html"
            if test_file.exists():
                return FileResponse(test_file)
            
            raise HTTPException(status_code=404, detail="Excel fix test page not found")
        
        @self.app.get("/spreadsheet-test", response_class=HTMLResponse)
        async def spreadsheet_test_page():
            """Serve spreadsheet test page"""
            test_file = Path(__file__).parent / "static" / "spreadsheet-test.html"
            if test_file.exists():
                return FileResponse(test_file)
            
            raise HTTPException(status_code=404, detail="Spreadsheet test page not found")
        
        @self.app.get("/pdf-conversion-test", response_class=HTMLResponse)
        async def pdf_conversion_test_page():
            """Serve PDF conversion test page"""
            test_file = Path(__file__).parent / "static" / "pdf-conversion-test.html"
            if test_file.exists():
                return FileResponse(test_file)
            
            raise HTTPException(status_code=404, detail="PDF conversion test page not found")
        
        @self.app.get("/groupdocs-test", response_class=HTMLResponse)
        async def groupdocs_test():
            """Serve GroupDocs Excel to PDF test page"""
            test_file = Path(__file__).parent / "static" / "groupdocs-test.html"
            if test_file.exists():
                return FileResponse(test_file)
            
            raise HTTPException(status_code=404, detail="GroupDocs test page not found")
        
        @self.app.get("/groupdocs-excel-demo")
        async def groupdocs_excel_demo():
            return FileResponse("web/groupdocs_excel_demo.html")
        
        @self.app.get("/professional-pdf-test", response_class=HTMLResponse)
        async def professional_pdf_test_page():
            """Serve Professional PDF conversion test page"""
            test_file = Path(__file__).parent / "static" / "professional-pdf-test.html"
            if test_file.exists():
                return FileResponse(test_file)
            
            raise HTTPException(status_code=404, detail="Professional PDF test page not found")
        
        @self.app.get("/perfect-pdf-conversion", response_class=HTMLResponse)
        async def perfect_pdf_conversion_page():
            """Serve Perfect PDF conversion page"""
            test_file = Path(__file__).parent / "static" / "perfect-pdf-conversion.html"
            if test_file.exists():
                return FileResponse(test_file)
            
            raise HTTPException(status_code=404, detail="Perfect PDF conversion page not found")
        
        @self.app.get("/excel-pywin32-converter", response_class=HTMLResponse)
        async def excel_pywin32_converter_page():
            """Serve Excel PyWin32 converter page"""
            test_file = Path(__file__).parent / "static" / "excel-pywin32-converter.html"
            if test_file.exists():
                return FileResponse(test_file)
            
            raise HTTPException(status_code=404, detail="Excel PyWin32 converter page not found")
    
    def _setup_api_routes(self):
        """Setup all API endpoints"""
        self._setup_system_endpoints()
        
        # Initialize and include Excel PyWin32 router (only PDF conversion method)
        init_services(self.excel_pywin32_service, self.file_service)
        self.app.include_router(excel_pywin32_router, prefix="/api/excel-pywin32", tags=["excel-pywin32"])
        
        # Initialize and include Document Manipulation router
        set_enhanced_document_service(self.enhanced_document_service)
        self.app.include_router(document_manipulation_router, prefix="/api/document-manipulation", tags=["document-manipulation"])
    
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
    
    def verify_auth_with_bearer(self, credentials: HTTPAuthorizationCredentials = Depends(security)):
        """Verify API authentication with bearer token"""
        if not credentials or credentials.credentials != self.config['security']['api_key']:
            raise HTTPException(status_code=401, detail="Invalid authentication")
        return True
    
    def _setup_system_endpoints(self):
        """Setup system-related API endpoints"""
    
        @self.app.get("/api/printers")
        async def get_printers():
            """Get available printers"""
            printers = self.printer_service.get_all_printers(force_refresh=True)
            return {"printers": [printer.__dict__ for printer in printers]}
    
        @self.app.get("/api/printers/{printer_id}/status")
        async def get_printer_status(printer_id: str):
            """Get printer status"""
            printer = self.printer_service.get_printer(printer_id)
            if not printer:
                raise HTTPException(status_code=404, detail="Printer not found")
            
            status = self.printer_service.get_printer_status(printer_id)
            
            # Convert enum to string and create response object
            status_map = {
                "ONLINE": {"status": "online", "message": "Printer is ready"},
                "OFFLINE": {"status": "offline", "message": "Printer is offline"},
                "BUSY": {"status": "busy", "message": "Printer is busy"},
                "PAUSED": {"status": "paused", "message": "Printer is paused"},
                "ERROR": {"status": "error", "message": "Printer has an error"}
            }
            
            return status_map.get(status.name, {"status": "unknown", "message": "Unknown status"})
        
        @self.app.get("/api/jobs")
        async def get_jobs():
            """Get all print jobs"""
            jobs = self.job_service.get_jobs()
            return {"jobs": [job.__dict__ for job in jobs]}
        
        @self.app.get("/api/jobs/{job_id}")
        async def get_job(job_id: str):
            """Get specific job details"""
            job = self.job_service.get_job(job_id)
            if not job:
                raise HTTPException(status_code=404, detail="Job not found")
            return job.__dict__
        
        @self.app.delete("/api/jobs/{job_id}")
        async def cancel_job(job_id: str):
            """Cancel a print job"""
            success = self.job_service.cancel_job(job_id)
            if not success:
                raise HTTPException(status_code=404, detail="Job not found")
            return {"message": "Job cancelled successfully"}
        
        @self.app.post("/api/files/upload")
        async def upload_file(file: UploadFile = File(...)):
            """Upload a file for processing with automatic PDF conversion for non-PDF files"""
            try:
                # Read file content
                file_content = await file.read()
                
                # Save the uploaded file
                file_info = self.file_service.save_uploaded_file(file_content, file.filename)
                
                # Check if file is not PDF and convert automatically
                file_extension = Path(file.filename).suffix.lower()
                pdf_path = None
                converted_to_pdf = False
                
                if file_extension != '.pdf':
                    # Check if file type is supported for conversion
                    supported_extensions = {'.xlsx', '.xls', '.docx', '.doc', '.pptx', '.ppt', 
                                          '.txt', '.csv', '.jpg', '.jpeg', '.png', '.bmp', 
                                          '.tiff', '.gif'}
                    
                    if file_extension in supported_extensions:
                        try:
                            # Convert to PDF using pywin32 service
                            input_path = str(self.file_service.upload_dir / file_info["name"])
                            
                            result = self.excel_pywin32_service.convert_to_pdf(
                                input_path=input_path
                            )
                            
                            if result["success"]:
                                pdf_path = result["output_path"]
                                converted_to_pdf = True
                                logger.info(f"File {file.filename} automatically converted to PDF: {pdf_path}")
                            else:
                                logger.warning(f"Failed to convert {file.filename} to PDF: {result.get('error', 'Unknown error')}")
                        except Exception as conv_error:
                            logger.error(f"Error converting {file.filename} to PDF: {conv_error}")
                
                response_data = {
                    "file_id": file_info["name"],
                    "original_name": file.filename,
                    "size": len(file_content),
                    "file_type": file_info.get("type", "unknown"),
                    "converted_to_pdf": converted_to_pdf,
                    "message": "File uploaded successfully"
                }
                
                if converted_to_pdf and pdf_path:
                    pdf_filename = Path(pdf_path).name
                    response_data.update({
                        "pdf_file_id": pdf_filename,
                        "pdf_path": pdf_path,
                        "pdf_url": f"/api/files/download/{pdf_filename}",
                        "message": "File uploaded and automatically converted to PDF"
                    })
                
                return response_data
                
            except Exception as e:
                logger.error(f"Error uploading file: {e}")
                raise HTTPException(status_code=500, detail=f"Upload failed: {str(e)}")
        
        @self.app.post("/api/print")
        async def print_document(
            file: UploadFile = File(...), 
            printer_id: str = Form(...), 
            copies: int = Form(1)
        ):
            """Print a document with automatic PDF conversion for non-PDF files"""
            try:
                # Read file content
                file_content = await file.read()
                
                # Save the uploaded file
                file_info = self.file_service.save_uploaded_file(file_content, file.filename)
                
                # Convert file_id to file_path
                file_path = str(self.file_service.upload_dir / file_info["name"])
                
                # Check if file is not PDF and convert automatically for printing
                file_extension = Path(file.filename).suffix.lower()
                final_file_path = file_path
                converted_for_print = False
                
                if file_extension != '.pdf':
                    # Check if file type is supported for conversion
                    supported_extensions = {'.xlsx', '.xls', '.docx', '.doc', '.pptx', '.ppt', 
                                          '.txt', '.csv', '.jpg', '.jpeg', '.png', '.bmp', 
                                          '.tiff', '.gif'}
                    
                    if file_extension in supported_extensions:
                        try:
                            # Convert to PDF using pywin32 service for printing
                            result = self.excel_pywin32_service.convert_to_pdf(
                                input_path=file_path,
                                quality="high",
                                preserve_formatting=True
                            )
                            
                            if result["success"]:
                                final_file_path = result["output_path"]
                                converted_for_print = True
                                logger.info(f"File {file.filename} automatically converted to PDF for printing: {final_file_path}")
                            else:
                                logger.warning(f"Failed to convert {file.filename} to PDF for printing: {result.get('error', 'Unknown error')}")
                                # Continue with original file if conversion fails
                        except Exception as conv_error:
                            logger.error(f"Error converting {file.filename} to PDF for printing: {conv_error}")
                            # Continue with original file if conversion fails
                
                # Create PrintSettings object
                print_settings = PrintSettings(
                    copies=copies,
                    color_mode="color",
                    paper_size="A4",
                    orientation="portrait",
                    quality="normal",
                    duplex="none",
                    fit_to_page="actual_size"
                )
                
                # Submit print job with the final file path (PDF if converted, original if not)
                job = self.job_service.submit_job(
                    file_path=final_file_path,
                    printer_id=printer_id,
                    settings=print_settings,
                    user="web_user",
                    client_ip=None
                )
                
                response_data = {
                    "job_id": job.id,
                    "status": job.status.value if hasattr(job.status, 'value') else str(job.status),
                    "converted_to_pdf_for_print": converted_for_print,
                    "message": "Print job submitted successfully"
                }
                
                if converted_for_print:
                    response_data["message"] = "File converted to PDF and print job submitted successfully"
                
                return response_data
                
            except Exception as e:
                logger.error(f"Error printing document: {e}")
                raise HTTPException(status_code=400, detail=str(e))
        
        @self.app.post("/api/excel-pywin32/convert")
        async def convert_to_pywin32_pdf(
            file: UploadFile = File(...),
            quality: str = Form("high"),
            output_filename: Optional[str] = Form(None)
        ):
            """Convert document to PDF using pywin32 service"""
            try:
                # Read file content
                file_content = await file.read()
                
                # Save the uploaded file
                file_info = self.file_service.save_uploaded_file(file_content, file.filename)
                input_path = str(self.file_service.upload_dir / file_info["name"])
                
                # Convert to PDF
                result = self.excel_pywin32_service.convert_to_pdf(
                    input_path=input_path,
                    output_path=output_filename,
                    quality=quality
                )
                
                if result["success"]:
                    return {
                        "success": True,
                        "input_file": file.filename,
                        "output_path": result["output_path"],
                        "file_size": result["file_size"],
                        "quality": result["quality"],
                        "conversion_method": result["conversion_method"],
                        "metadata": result.get("metadata", {}),
                        "message": "Document converted to PDF successfully"
                    }
                else:
                    raise HTTPException(status_code=400, detail=result["error"])
                    
            except Exception as e:
                logger.error(f"Error converting to PDF: {e}")
                raise HTTPException(status_code=500, detail=f"Conversion failed: {str(e)}")
        
        @self.app.post("/api/excel-pywin32/preview")
        async def get_pdf_preview(
            request: dict
        ):
            """Get PDF preview with pywin32 conversion"""
            try:
                file_id = request.get('file_id')
                if not file_id:
                    raise HTTPException(status_code=400, detail="file_id is required")
                
                # Get file path from file service
                input_path = str(self.file_service.upload_dir / file_id)
                
                # Get preview
                result = self.excel_pywin32_service.preview_pdf_settings(input_path)
                
                if result["success"]:
                    return {
                        "success": True,
                        "excel_info": result["excel_info"],
                        "selected_sheets": result["selected_sheets"],
                        "estimated_pages": result["estimated_pages"],
                        "message": "Preview generated successfully"
                    }
                else:
                    raise HTTPException(status_code=400, detail=result["error"])
                    
            except Exception as e:
                logger.error(f"Error generating preview: {e}")
                raise HTTPException(status_code=500, detail=f"Preview generation failed: {str(e)}")
        
        @self.app.post("/api/files/convert-to-pdf")
        async def convert_to_pdf(
            file: UploadFile = File(...),
            quality: str = Form("high"),
            preserve_formatting: bool = Form(True)
        ):
            """Convert file to PDF - endpoint for frontend compatibility"""
            try:
                # Read file content
                file_content = await file.read()
                
                # Save the uploaded file
                file_info = self.file_service.save_uploaded_file(file_content, file.filename)
                input_path = str(self.file_service.upload_dir / file_info["name"])
                
                # Convert to PDF using pywin32 service
                result = self.excel_pywin32_service.convert_to_pdf(
                    input_path=input_path
                )
                
                if result["success"]:
                    # Return PDF file path for frontend
                    pdf_filename = Path(result["output_path"]).name
                    return {
                        "success": True,
                        "pdf_url": f"/api/files/download/{pdf_filename}",
                        "pdf_path": result["output_path"],
                        "file_size": result["file_size"],
                        "quality": result["quality"],
                        "message": "File converted to PDF successfully"
                    }
                else:
                    raise HTTPException(status_code=400, detail=result["error"])
                    
            except Exception as e:
                logger.error(f"Error converting file to PDF: {e}")
                raise HTTPException(status_code=500, detail=f"Conversion failed: {str(e)}")
        
        @self.app.get("/api/files/download/{filename}")
        async def download_file(filename: str):
            """Download converted PDF file"""
            try:
                # Look for file in temp directory
                file_path = Path(self.file_service.temp_dir) / filename
                
                if not file_path.exists():
                    # Also check upload directory
                    file_path = Path(self.file_service.upload_dir) / filename
                
                if not file_path.exists():
                    raise HTTPException(status_code=404, detail="File not found")
                
                return FileResponse(
                    path=str(file_path),
                    filename=filename,
                    media_type="application/pdf"
                )
                
            except Exception as e:
                logger.error(f"Error downloading file: {e}")
                raise HTTPException(status_code=500, detail=f"Download failed: {str(e)}")
        
        @self.app.get("/api/files/serve/{file_path:path}")
        async def serve_file(file_path: str):
            """Serve files from temp or upload directory"""
            try:
                # Clean the file path to prevent directory traversal
                clean_path = Path(file_path).name
                
                # Look for file in temp directory first
                full_path = Path(self.file_service.temp_dir) / clean_path
                
                if not full_path.exists():
                    # Also check upload directory
                    full_path = Path(self.file_service.upload_dir) / clean_path
                
                if not full_path.exists():
                    logger.error(f"File not found: {file_path} (cleaned: {clean_path})")
                    raise HTTPException(status_code=404, detail=f"File not found: {file_path}")
                
                # Determine media type based on file extension
                media_type = "application/pdf" if full_path.suffix.lower() == ".pdf" else "application/octet-stream"
                
                return FileResponse(
                    path=str(full_path),
                    filename=clean_path,
                    media_type=media_type
                )
                
            except HTTPException:
                raise
            except Exception as e:
                logger.error(f"Error serving file {file_path}: {e}")
                raise HTTPException(status_code=500, detail=f"Serve failed: {str(e)}")

# Main execution
if __name__ == "__main__":
    try:
        # Initialize the application
        print_server = PrintServerApp()
        app = print_server.app
        
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
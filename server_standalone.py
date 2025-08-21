#!/usr/bin/env python3
"""
Standalone Server untuk Printer Sharing
File ini akan di-compile menjadi executable (.exe) untuk Windows
"""

import sys
import os
import logging
import traceback
import socket
import threading
import time
import json
from pathlib import Path
from typing import List, Dict, Any, Optional
import win32print
import win32api
import yaml
import uvicorn
from fastapi import FastAPI, HTTPException, UploadFile, File, Form
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from contextlib import asynccontextmanager

# Add current directory to Python path
current_dir = Path(__file__).parent
sys.path.insert(0, str(current_dir))

# Global variables for services
printer_service = None
job_service = None
file_service = None
error_forwarder = None

class PrinterAutoDiscovery:
    """Auto-discovery dan management printer Windows"""
    
    def __init__(self):
        self.printers = {}
        self.default_printer = None
        self.logger = logging.getLogger(__name__)
    
    def discover_printers(self) -> List[Dict[str, Any]]:
        """Discover semua printer yang tersedia di sistem"""
        try:
            printers = []
            
            # Get default printer
            try:
                self.default_printer = win32print.GetDefaultPrinter()
            except Exception as e:
                self.logger.warning(f"No default printer set: {e}")
                self.default_printer = None
            
            # Enumerate all printers
            printer_enum = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
            
            for printer_info in printer_enum:
                printer_name = printer_info[2]  # Printer name
                
                try:
                    # Get printer handle
                    printer_handle = win32print.OpenPrinter(printer_name)
                    
                    # Get printer info
                    printer_info_dict = win32print.GetPrinter(printer_handle, 2)
                    
                    printer_data = {
                        'id': printer_name.replace(' ', '_').lower(),
                        'name': printer_name,
                        'driver': printer_info_dict.get('pDriverName', 'Unknown'),
                        'port': printer_info_dict.get('pPortName', 'Unknown'),
                        'location': printer_info_dict.get('pLocation', ''),
                        'comment': printer_info_dict.get('pComment', ''),
                        'status': self._get_printer_status(printer_handle),
                        'is_default': printer_name == self.default_printer,
                        'is_online': True,  # Assume online if we can open it
                        'capabilities': self._get_printer_capabilities(printer_handle)
                    }
                    
                    printers.append(printer_data)
                    self.printers[printer_name] = printer_data
                    
                    win32print.ClosePrinter(printer_handle)
                    
                except Exception as e:
                    self.logger.error(f"Error getting info for printer {printer_name}: {e}")
                    # Add basic info even if detailed info fails
                    printers.append({
                        'id': printer_name.replace(' ', '_').lower(),
                        'name': printer_name,
                        'driver': 'Unknown',
                        'port': 'Unknown',
                        'location': '',
                        'comment': '',
                        'status': 'Unknown',
                        'is_default': printer_name == self.default_printer,
                        'is_online': False,
                        'capabilities': {}
                    })
            
            self.logger.info(f"Discovered {len(printers)} printers")
            return printers
            
        except Exception as e:
            self.logger.error(f"Error discovering printers: {e}")
            return []
    
    def _get_printer_status(self, printer_handle) -> str:
        """Get printer status"""
        try:
            printer_info = win32print.GetPrinter(printer_handle, 2)
            status = printer_info.get('Status', 0)
            
            if status == 0:
                return 'ready'
            elif status & win32print.PRINTER_STATUS_BUSY:
                return 'busy'
            elif status & win32print.PRINTER_STATUS_ERROR:
                return 'error'
            elif status & win32print.PRINTER_STATUS_OFFLINE:
                return 'offline'
            elif status & win32print.PRINTER_STATUS_OUT_OF_MEMORY:
                return 'error'
            elif status & win32print.PRINTER_STATUS_PAPER_OUT:
                return 'error'
            else:
                return 'unknown'
        except:
            return 'unknown'
    
    def _get_printer_capabilities(self, printer_handle) -> Dict[str, Any]:
        """Get printer capabilities"""
        try:
            # Basic capabilities - can be extended
            return {
                'color': True,  # Assume color support
                'duplex': True,  # Assume duplex support
                'paper_sizes': ['A4', 'Letter', 'Legal'],  # Common sizes
                'resolutions': ['300x300', '600x600', '1200x1200']
            }
        except:
            return {}
    
    def set_default_printer(self, printer_name: str) -> bool:
        """Set default printer"""
        try:
            win32print.SetDefaultPrinter(printer_name)
            self.default_printer = printer_name
            self.logger.info(f"Default printer set to: {printer_name}")
            return True
        except Exception as e:
            self.logger.error(f"Error setting default printer: {e}")
            return False
    
    def print_document(self, printer_name: str, file_path: str, options: Dict[str, Any] = None) -> bool:
        """Print document to specified printer"""
        try:
            # Simple print implementation
            printer_handle = win32print.OpenPrinter(printer_name)
            
            # Start print job
            job_info = ("Python Print Job", None, "RAW")
            job_id = win32print.StartDocPrinter(printer_handle, 1, job_info)
            
            try:
                win32print.StartPagePrinter(printer_handle)
                
                # Read file and send to printer
                with open(file_path, 'rb') as f:
                    data = f.read()
                    win32print.WritePrinter(printer_handle, data)
                
                win32print.EndPagePrinter(printer_handle)
                win32print.EndDocPrinter(printer_handle)
                
                self.logger.info(f"Print job {job_id} sent to {printer_name}")
                return True
                
            except Exception as e:
                win32print.AbortPrinter(printer_handle)
                raise e
            finally:
                win32print.ClosePrinter(printer_handle)
                
        except Exception as e:
            self.logger.error(f"Error printing to {printer_name}: {e}")
            return False

class ErrorForwarder:
    """Forward errors to remote server"""
    
    def __init__(self, remote_host: str = None, remote_port: int = 9999):
        self.remote_host = remote_host
        self.remote_port = remote_port
        self.logger = logging.getLogger(__name__)
        self.enabled = remote_host is not None
    
    def forward_error(self, error_msg: str, error_type: str = "ERROR"):
        """Forward error message to remote server"""
        if not self.enabled:
            return
        
        try:
            message = {
                'timestamp': time.strftime('%Y-%m-%d %H:%M:%S'),
                'type': error_type,
                'message': error_msg,
                'hostname': socket.gethostname()
            }
            
            # Send via UDP for simplicity
            sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            sock.settimeout(5)
            
            message_str = f"{message['timestamp']} [{message['type']}] {message['hostname']}: {message['message']}"
            sock.sendto(message_str.encode('utf-8'), (self.remote_host, self.remote_port))
            sock.close()
            
        except Exception as e:
            self.logger.error(f"Failed to forward error: {e}")

class CustomLogHandler(logging.Handler):
    """Custom log handler that forwards errors"""
    
    def __init__(self, error_forwarder: ErrorForwarder):
        super().__init__()
        self.error_forwarder = error_forwarder
    
    def emit(self, record):
        if record.levelno >= logging.ERROR:
            self.error_forwarder.forward_error(
                self.format(record),
                "ERROR" if record.levelno == logging.ERROR else "CRITICAL"
            )

# Global printer discovery instance
auto_discovery = None

# FastAPI app
app = FastAPI(
    title="Printer Sharing Server",
    description="Server untuk berbagi printer di jaringan",
    version="1.0.0"
)

# CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/")
async def root():
    """Root endpoint"""
    return HTMLResponse("""
    <!DOCTYPE html>
    <html>
    <head>
        <title>Printer Sharing Server</title>
        <style>
            body { font-family: Arial, sans-serif; margin: 40px; }
            .container { max-width: 800px; margin: 0 auto; }
            .printer { border: 1px solid #ddd; padding: 15px; margin: 10px 0; border-radius: 5px; }
            .default { background-color: #e8f5e8; }
            .offline { background-color: #ffe8e8; }
            .status { font-weight: bold; }
            .actions { margin-top: 10px; }
            button { padding: 8px 16px; margin: 5px; border: none; border-radius: 3px; cursor: pointer; }
            .btn-primary { background-color: #007bff; color: white; }
            .btn-success { background-color: #28a745; color: white; }
        </style>
    </head>
    <body>
        <div class="container">
            <h1>🖨️ Printer Sharing Server</h1>
            <p>Server berjalan dan siap digunakan!</p>
            
            <h2>Quick Links</h2>
            <ul>
                <li><a href="/printers">📋 Daftar Printer (JSON)</a></li>
                <li><a href="/docs">📖 API Documentation</a></li>
                <li><a href="/status">⚡ Server Status</a></li>
            </ul>
            
            <h2>Discovered Printers</h2>
            <div id="printers">Loading...</div>
        </div>
        
        <script>
            fetch('/printers')
                .then(response => response.json())
                .then(data => {
                    const container = document.getElementById('printers');
                    if (data.length === 0) {
                        container.innerHTML = '<p>No printers found</p>';
                        return;
                    }
                    
                    container.innerHTML = data.map(printer => `
                        <div class="printer ${printer.is_default ? 'default' : ''} ${!printer.is_online ? 'offline' : ''}">
                            <h3>${printer.name} ${printer.is_default ? '(Default)' : ''}</h3>
                            <p><strong>Driver:</strong> ${printer.driver}</p>
                            <p><strong>Port:</strong> ${printer.port}</p>
                            <p><strong>Status:</strong> <span class="status">${printer.status}</span></p>
                            ${printer.location ? `<p><strong>Location:</strong> ${printer.location}</p>` : ''}
                            <div class="actions">
                                ${!printer.is_default ? `<button class="btn-primary" onclick="setDefault('${printer.name}')">Set as Default</button>` : ''}
                                <button class="btn-success" onclick="testPrint('${printer.name}')">Test Print</button>
                            </div>
                        </div>
                    `).join('');
                })
                .catch(error => {
                    document.getElementById('printers').innerHTML = '<p>Error loading printers</p>';
                });
            
            function setDefault(printerName) {
                fetch(`/printers/${encodeURIComponent(printerName)}/set-default`, { method: 'POST' })
                    .then(response => response.json())
                    .then(data => {
                        if (data.success) {
                            location.reload();
                        } else {
                            alert('Failed to set default printer');
                        }
                    });
            }
            
            function testPrint(printerName) {
                alert('Test print feature will be implemented');
            }
        </script>
    </body>
    </html>
    """)

@app.get("/api/status")
async def get_status():
    """Get server status"""
    return {
        "status": "running",
        "hostname": socket.gethostname(),
        "printers_count": len(auto_discovery.printers) if auto_discovery else 0,
        "default_printer": auto_discovery.default_printer if auto_discovery else None
    }

@app.get("/api/printers")
async def get_printers():
    """Get list of available printers"""
    if not auto_discovery:
        return {"data": []}
    
    printers = auto_discovery.discover_printers()
    return {"data": printers}

@app.post("/api/printers/{printer_name}/set-default")
async def set_default_printer(printer_name: str):
    """Set default printer"""
    if not auto_discovery:
        raise HTTPException(status_code=500, detail="Printer service not available")
    
    success = auto_discovery.set_default_printer(printer_name)
    return {"success": success, "printer": printer_name}

@app.post("/api/files/upload")
async def upload_file(file: UploadFile = File(...)):
    """Upload file for printing"""
    try:
        # Save uploaded file temporarily
        temp_dir = Path("temp")
        temp_dir.mkdir(exist_ok=True)
        
        # Generate unique file ID
        import uuid
        file_id = str(uuid.uuid4())
        file_extension = Path(file.filename).suffix
        temp_file = temp_dir / f"{file_id}{file_extension}"
        
        # Save file
        with open(temp_file, "wb") as f:
            content = await file.read()
            f.write(content)
        
        return {
            "success": True,
            "file_id": file_id,
            "filename": file.filename,
            "size": len(content),
            "type": file.content_type or "application/octet-stream",
            "message": "File uploaded successfully"
        }
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

@app.post("/api/jobs/submit")
async def submit_print_job(
    file_id: str = Form(...),
    printer_name: str = Form(...),
    copies: int = Form(1),
    color_mode: str = Form("color"),
    paper_size: str = Form("A4"),
    orientation: str = Form("portrait"),
    quality: str = Form("normal"),
    duplex: str = Form("none")
):
    """Submit print job"""
    if not auto_discovery:
        raise HTTPException(status_code=500, detail="Printer service not available")
    
    # Find the uploaded file
    temp_dir = Path("temp")
    temp_files = list(temp_dir.glob(f"{file_id}.*"))
    
    if not temp_files:
        raise HTTPException(status_code=404, detail="File not found")
    
    temp_file = temp_files[0]
    
    try:
        # Print the file
        success = auto_discovery.print_document(printer_name, str(temp_file))
        
        if success:
            # Generate job ID
            import uuid
            job_id = str(uuid.uuid4())
            
            return {
                "success": True,
                "job_id": job_id,
                "message": f"Print job submitted to {printer_name}",
                "printer_name": printer_name,
                "copies": copies,
                "status": "submitted"
            }
        else:
            raise HTTPException(status_code=500, detail="Print job failed")
    
    finally:
        # Clean up temp file
        if temp_file.exists():
            temp_file.unlink()

@app.post("/api/print")
async def print_document(file: UploadFile = File(...), printer_name: str = None):
    """Print document"""
    if not auto_discovery:
        raise HTTPException(status_code=500, detail="Printer service not available")
    
    # Use default printer if none specified
    if not printer_name:
        printer_name = auto_discovery.default_printer
        if not printer_name:
            raise HTTPException(status_code=400, detail="No printer specified and no default printer set")
    
    # Save uploaded file temporarily
    temp_dir = Path("temp")
    temp_dir.mkdir(exist_ok=True)
    
    temp_file = temp_dir / file.filename
    with open(temp_file, "wb") as f:
        content = await file.read()
        f.write(content)
    
    try:
        # Print the file
        success = auto_discovery.print_document(printer_name, str(temp_file))
        
        if success:
            return {"success": True, "message": f"Document sent to {printer_name}"}
        else:
            raise HTTPException(status_code=500, detail="Print job failed")
    
    finally:
        # Clean up temp file
        if temp_file.exists():
            temp_file.unlink()

def setup_logging(error_forwarder: ErrorForwarder = None):
    """Setup logging configuration"""
    # Create logs directory
    logs_dir = Path("logs")
    logs_dir.mkdir(exist_ok=True)
    
    # Configure logging
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(logs_dir / "server.log"),
            logging.StreamHandler(sys.stdout)
        ]
    )
    
    # Add error forwarder if provided
    if error_forwarder:
        logger = logging.getLogger()
        logger.addHandler(CustomLogHandler(error_forwarder))

def load_config() -> Dict[str, Any]:
    """Load configuration"""
    config_path = Path("config.yaml")
    
    if config_path.exists():
        with open(config_path, 'r', encoding='utf-8') as f:
            return yaml.safe_load(f)
    else:
        # Default configuration
        return {
            'server': {
                'host': '0.0.0.0',
                'port': 8080,
                'debug': False,
                'cors_origins': ['*']
            },
            'storage': {
                'temp_dir': 'temp'
            },
            'error_forwarding': {
                'enabled': False,
                'remote_host': None,
                'remote_port': 9999
            }
        }

def main():
    """Main entry point"""
    global auto_discovery, error_forwarder
    
    try:
        print("=== Printer Sharing Server ===")
        print("Starting server...")
        
        # Load configuration
        config = load_config()
        
        # Setup error forwarding
        if config.get('error_forwarding', {}).get('enabled', False):
            error_forwarder = ErrorForwarder(
                remote_host=config['error_forwarding'].get('remote_host'),
                remote_port=config['error_forwarding'].get('remote_port', 9999)
            )
        
        # Setup logging
        setup_logging(error_forwarder)
        logger = logging.getLogger(__name__)
        
        # Initialize auto-discovery
        auto_discovery = PrinterAutoDiscovery()
        
        # Print server info
        host = config['server']['host']
        port = config['server']['port']
        
        print(f"\nServer started successfully!")
        print(f"Web Interface: http://localhost:{port}")
        print(f"API Documentation: http://localhost:{port}/docs")
        print(f"\nDiscovered Printers:")
        
        printers = auto_discovery.discover_printers()
        for printer in printers:
            status = "[DEFAULT]" if printer['is_default'] else ""
            print(f"  - {printer['name']} {status} ({printer['status']})")
        
        if auto_discovery.default_printer:
            print(f"\nDefault Printer: {auto_discovery.default_printer}")
        
        print(f"\nPress Ctrl+C to stop the server")
        
        # Start server
        uvicorn.run(
            app,
            host=host,
            port=port,
            log_level="info",
            access_log=True
        )
        
    except KeyboardInterrupt:
        print("\nShutting down server...")
    except Exception as e:
        error_msg = f"Server error: {str(e)}\n{traceback.format_exc()}"
        print(error_msg)
        
        if error_forwarder:
            error_forwarder.forward_error(error_msg, "CRITICAL")
        
        sys.exit(1)
    finally:
        print("Server stopped.")

if __name__ == "__main__":
    main()
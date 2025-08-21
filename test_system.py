#!/usr/bin/env python3
"""
Test script for Printer Sharing System
This script tests the basic functionality of the printer sharing system.
"""

import asyncio
import sys
import os
import time
from pathlib import Path

# Add project root to path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

try:
    from client.client import PrinterClient, create_client
    from client.discovery import PrinterDiscoveryClient
    from client.models import PrintSettings, ColorMode, PaperSize, Orientation, Quality
    from server.services.printer_service import PrinterService
    from server.services.discovery_service import DiscoveryService
except ImportError as e:
    print(f"Import error: {e}")
    print("Please make sure all dependencies are installed: pip install -r requirements.txt")
    sys.exit(1)

class SystemTester:
    """Test class for the printer sharing system."""
    
    def __init__(self):
        self.client = None
        self.discovery_client = None
        self.test_results = []
    
    def log_test(self, test_name: str, success: bool, message: str = ""):
        """Log test result."""
        status = "PASS" if success else "FAIL"
        result = f"[{status}] {test_name}"
        if message:
            result += f": {message}"
        print(result)
        self.test_results.append((test_name, success, message))
    
    def test_imports(self):
        """Test if all required modules can be imported."""
        print("\n=== Testing Imports ===")
        
        try:
            import fastapi
            self.log_test("FastAPI import", True)
        except ImportError:
            self.log_test("FastAPI import", False, "FastAPI not installed")
        
        try:
            import pydantic
            self.log_test("Pydantic import", True)
        except ImportError:
            self.log_test("Pydantic import", False, "Pydantic not installed")
        
        try:
            import zeroconf
            self.log_test("Zeroconf import", True)
        except ImportError:
            self.log_test("Zeroconf import", False, "Zeroconf not installed")
        
        try:
            import win32print
            self.log_test("Win32print import", True)
        except ImportError:
            self.log_test("Win32print import", False, "pywin32 not installed")
    
    def test_printer_service(self):
        """Test printer service functionality."""
        print("\n=== Testing Printer Service ===")
        
        try:
            printer_service = PrinterService()
            printers = printer_service.get_all_printers()
            
            if printers:
                self.log_test("Get printers", True, f"Found {len(printers)} printers")
                
                # Test getting specific printer
                first_printer = printers[0]
                printer_info = printer_service.get_printer(first_printer.name)
                if printer_info:
                    self.log_test("Get specific printer", True, f"Got info for {first_printer.name}")
                else:
                    self.log_test("Get specific printer", False, "Could not get printer info")
            else:
                self.log_test("Get printers", False, "No printers found")
                
        except Exception as e:
            self.log_test("Printer service", False, str(e))
    
    def test_discovery_service(self):
        """Test discovery service functionality."""
        print("\n=== Testing Discovery Service ===")
        
        try:
            discovery_service = DiscoveryService()
            
            # Test starting discovery
            discovery_service.start_discovery()
            self.log_test("Start discovery", True)
            
            # Wait a bit for discovery
            time.sleep(2)
            
            # Test getting discovered printers
            discovered = discovery_service.get_discovered_printers()
            self.log_test("Get discovered printers", True, f"Found {len(discovered)} services")
            
            # Stop discovery
            discovery_service.stop_discovery()
            self.log_test("Stop discovery", True)
            
        except Exception as e:
            self.log_test("Discovery service", False, str(e))
    
    async def test_client_discovery(self):
        """Test client discovery functionality."""
        print("\n=== Testing Client Discovery ===")
        
        try:
            self.discovery_client = PrinterDiscoveryClient()
            
            # Start discovery
            self.discovery_client.start_discovery()
            self.log_test("Start client discovery", True)
            
            # Wait for discovery
            await asyncio.sleep(3)
            
            # Get discovered servers
            servers = self.discovery_client.get_discovered_servers()
            self.log_test("Get discovered servers", True, f"Found {len(servers)} servers")
            
            # Stop discovery
            self.discovery_client.stop_discovery()
            self.log_test("Stop client discovery", True)
            
        except Exception as e:
            self.log_test("Client discovery", False, str(e))
    
    async def test_client_connection(self):
        """Test client connection to server."""
        print("\n=== Testing Client Connection ===")
        
        try:
            # Try to connect to local server
            self.client = create_client("http://localhost:8000")
            
            # Test connection (this will fail if server is not running)
            try:
                status = await self.client.get_server_status()
                self.log_test("Connect to server", True, "Server is running")
                
                # Test getting printers
                printers = await self.client.get_printers()
                self.log_test("Get printers from server", True, f"Found {len(printers)} printers")
                
            except Exception as e:
                self.log_test("Connect to server", False, "Server not running or not accessible")
                
        except Exception as e:
            self.log_test("Client connection", False, str(e))
    
    def test_file_structure(self):
        """Test if all required files exist."""
        print("\n=== Testing File Structure ===")
        
        required_files = [
            "server/main.py",
            "server/models/__init__.py",
            "server/models/printer.py",
            "server/models/job.py",
            "server/models/response.py",
            "server/services/__init__.py",
            "server/services/printer_service.py",
            "server/services/job_service.py",
            "server/services/file_service.py",
            "server/services/discovery_service.py",
            "client/__init__.py",
            "client/models.py",
            "client/client.py",
            "client/discovery.py",
            "client/cli.py",
            "client/gui.py",
            "web/index.html",
            "web/app.js",
            "web/style.css",
            "requirements.txt",
            "config.yaml",
            "README.md",
            "setup.py"
        ]
        
        for file_path in required_files:
            full_path = project_root / file_path
            exists = full_path.exists()
            self.log_test(f"File exists: {file_path}", exists)
    
    def test_configuration(self):
        """Test configuration file."""
        print("\n=== Testing Configuration ===")
        
        try:
            import yaml
            config_path = project_root / "config.yaml"
            
            if config_path.exists():
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = yaml.safe_load(f)
                
                # Check required sections
                required_sections = ['server', 'printer', 'network', 'logging']
                for section in required_sections:
                    if section in config:
                        self.log_test(f"Config section: {section}", True)
                    else:
                        self.log_test(f"Config section: {section}", False, "Missing section")
            else:
                self.log_test("Config file exists", False, "config.yaml not found")
                
        except ImportError:
            self.log_test("YAML import", False, "PyYAML not installed")
        except Exception as e:
            self.log_test("Configuration test", False, str(e))
    
    async def run_all_tests(self):
        """Run all tests."""
        print("Printer Sharing System - Test Suite")
        print("====================================")
        
        # Run synchronous tests
        self.test_imports()
        self.test_file_structure()
        self.test_configuration()
        self.test_printer_service()
        self.test_discovery_service()
        
        # Run asynchronous tests
        await self.test_client_discovery()
        await self.test_client_connection()
        
        # Cleanup
        if self.client:
            await self.client.close()
        if self.discovery_client:
            self.discovery_client.stop_discovery()
        
        # Print summary
        self.print_summary()
    
    def print_summary(self):
        """Print test summary."""
        print("\n=== Test Summary ===")
        
        total_tests = len(self.test_results)
        passed_tests = sum(1 for _, success, _ in self.test_results if success)
        failed_tests = total_tests - passed_tests
        
        print(f"Total tests: {total_tests}")
        print(f"Passed: {passed_tests}")
        print(f"Failed: {failed_tests}")
        print(f"Success rate: {(passed_tests/total_tests)*100:.1f}%")
        
        if failed_tests > 0:
            print("\nFailed tests:")
            for test_name, success, message in self.test_results:
                if not success:
                    print(f"  - {test_name}: {message}")
        
        print("\nNote: Some tests may fail if the server is not running or dependencies are missing.")
        print("To start the server: python -m server.main")
        print("To install dependencies: pip install -r requirements.txt")

def main():
    """Main test function."""
    tester = SystemTester()
    
    try:
        asyncio.run(tester.run_all_tests())
    except KeyboardInterrupt:
        print("\nTests interrupted by user.")
    except Exception as e:
        print(f"\nTest suite error: {e}")
        return 1
    
    return 0

if __name__ == "__main__":
    sys.exit(main())
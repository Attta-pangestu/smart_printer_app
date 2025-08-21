#!/usr/bin/env python3
"""
Script untuk build executable dari server printer sharing
"""

import os
import sys
import shutil
import subprocess
from pathlib import Path

def clean_build_dirs():
    """Clean build directories"""
    dirs_to_clean = ['build', 'dist', '__pycache__']
    
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            print(f"Cleaning {dir_name}...")
            shutil.rmtree(dir_name)
    
    # Clean .spec files
    for spec_file in Path('.').glob('*.spec'):
        print(f"Removing {spec_file}...")
        spec_file.unlink()

def create_pyinstaller_spec():
    """Create PyInstaller spec file"""
    spec_content = '''
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['server_standalone.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('static', 'static'),
        ('templates', 'templates'),
        ('config.yaml', '.'),
    ],
    hiddenimports=[
        'uvicorn.lifespan.on',
        'uvicorn.lifespan.off',
        'uvicorn.protocols.websockets.auto',
        'uvicorn.protocols.websockets.websockets_impl',
        'uvicorn.protocols.http.auto',
        'uvicorn.protocols.http.h11_impl',
        'uvicorn.protocols.http.httptools_impl',
        'uvicorn.loops.auto',
        'uvicorn.loops.asyncio',
        'uvicorn.loops.uvloop',
        'fastapi.applications',
        'fastapi.routing',
        'fastapi.middleware',
        'fastapi.middleware.cors',
        'fastapi.staticfiles',
        'win32print',
        'win32api',
        'win32con',
        'pywintypes',
        'yaml',
        'PIL',
        'PIL._tkinter_finder',
        'zeroconf',
        'zeroconf._utils.ipaddress',
        'zeroconf._handlers.answers',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='PrinterServer',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico' if os.path.exists('icon.ico') else None,
)
'''
    
    with open('printer_server.spec', 'w') as f:
        f.write(spec_content)
    
    print("Created printer_server.spec")

def install_dependencies():
    """Install required dependencies"""
    print("Installing dependencies...")
    
    try:
        subprocess.run([sys.executable, '-m', 'pip', 'install', '-r', 'requirements.txt'], 
                      check=True, capture_output=True, text=True)
        print("Dependencies installed successfully")
    except subprocess.CalledProcessError as e:
        print(f"Error installing dependencies: {e}")
        print(f"stdout: {e.stdout}")
        print(f"stderr: {e.stderr}")
        return False
    
    return True

def build_executable():
    """Build executable using PyInstaller"""
    print("Building executable...")
    
    try:
        # Use the spec file for more control
        cmd = [sys.executable, '-m', 'PyInstaller', 'printer_server.spec', '--clean']
        
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        print("Build completed successfully!")
        print(f"Executable created: dist/PrinterServer.exe")
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"Build failed: {e}")
        print(f"stdout: {e.stdout}")
        print(f"stderr: {e.stderr}")
        return False

def copy_additional_files():
    """Copy additional files to dist directory"""
    dist_dir = Path('dist')
    
    if not dist_dir.exists():
        print("Dist directory not found")
        return
    
    # Files to copy
    files_to_copy = [
        'config.yaml',
        'README.md',
        'INSTALL.md',
        'USER_GUIDE.md'
    ]
    
    # Directories to copy
    dirs_to_copy = [
        'static',
        'templates'
    ]
    
    print("Copying additional files...")
    
    for file_name in files_to_copy:
        if os.path.exists(file_name):
            shutil.copy2(file_name, dist_dir)
            print(f"Copied {file_name}")
    
    for dir_name in dirs_to_copy:
        if os.path.exists(dir_name):
            dest_dir = dist_dir / dir_name
            if dest_dir.exists():
                shutil.rmtree(dest_dir)
            shutil.copytree(dir_name, dest_dir)
            print(f"Copied {dir_name}/")
    
    # Create batch file for easy startup
    batch_content = '''@echo off
echo Starting Printer Server...
echo.
echo Web Interface will be available at: http://localhost:8080
echo Press Ctrl+C to stop the server
echo.
PrinterServer.exe
pause
'''
    
    with open(dist_dir / 'start_server.bat', 'w') as f:
        f.write(batch_content)
    
    print("Created start_server.bat")

def create_installer_info():
    """Create installer information"""
    dist_dir = Path('dist')
    
    installer_info = '''
# Printer Server Installer

## Files Included:
- PrinterServer.exe - Main executable
- config.yaml - Configuration file
- static/ - Web interface files
- templates/ - HTML templates
- start_server.bat - Easy startup script
- README.md, INSTALL.md, USER_GUIDE.md - Documentation

## Installation:
1. Copy all files to desired directory (e.g., C:\\PrinterServer\\)
2. Run start_server.bat or PrinterServer.exe directly
3. Open web browser to http://localhost:8080

## Configuration:
Edit config.yaml to customize:
- Server port and host
- Error forwarding settings
- Storage directories
- Logging settings

## Requirements:
- Windows 7 or later
- No additional software required (all dependencies included)

## Default Settings:
- Server runs on port 8080
- Auto-discovers all installed printers
- Web interface available at http://localhost:8080
- Logs saved to logs/ directory

## Troubleshooting:
- Check logs/ directory for error messages
- Ensure port 8080 is not in use by other applications
- Run as Administrator if printer access issues occur
'''
    
    with open(dist_dir / 'INSTALLER_INFO.md', 'w') as f:
        f.write(installer_info)
    
    print("Created INSTALLER_INFO.md")

def main():
    """Main build process"""
    print("=== Printer Server Build Script ===")
    print()
    
    # Check if we're in the right directory
    if not os.path.exists('server_standalone.py'):
        print("Error: server_standalone.py not found")
        print("Please run this script from the project root directory")
        sys.exit(1)
    
    # Clean previous builds
    clean_build_dirs()
    
    # Install dependencies
    if not install_dependencies():
        print("Failed to install dependencies")
        sys.exit(1)
    
    # Create spec file
    create_pyinstaller_spec()
    
    # Build executable
    if not build_executable():
        print("Failed to build executable")
        sys.exit(1)
    
    # Copy additional files
    copy_additional_files()
    
    # Create installer info
    create_installer_info()
    
    print()
    print("=== Build Complete! ===")
    print(f"Executable: dist/PrinterServer.exe")
    print(f"Startup script: dist/start_server.bat")
    print(f"Size: {os.path.getsize('dist/PrinterServer.exe') / 1024 / 1024:.1f} MB")
    print()
    print("To test the executable:")
    print("1. cd dist")
    print("2. start_server.bat")
    print("3. Open http://localhost:8080 in browser")
    print()
    print("To distribute:")
    print("- Copy entire 'dist' folder to target machine")
    print("- Run start_server.bat on target machine")

if __name__ == "__main__":
    main()
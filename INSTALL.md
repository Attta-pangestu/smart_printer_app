# Installation Guide - Printer Sharing System

This guide will help you install and set up the Printer Sharing System on Windows.

## System Requirements

### Operating System
- Windows 10 or Windows 11
- Windows Server 2016 or later

### Software Requirements
- Python 3.8 or higher
- pip (Python package installer)
- Administrative privileges (for printer access)

### Hardware Requirements
- Minimum 2GB RAM
- 500MB free disk space
- Network connection (for printer sharing)

## Installation Steps

### 1. Install Python

If Python is not already installed:

1. Download Python from [python.org](https://www.python.org/downloads/)
2. Run the installer
3. **Important**: Check "Add Python to PATH" during installation
4. Verify installation:
   ```cmd
   python --version
   pip --version
   ```

### 2. Download the Project

Option A: Download ZIP
1. Download the project ZIP file
2. Extract to desired location (e.g., `C:\PrinterSharing`)

Option B: Clone with Git
```cmd
git clone <repository-url>
cd printer-sharing-system
```

### 3. Install Dependencies

Open Command Prompt or PowerShell as Administrator and navigate to the project directory:

```cmd
cd "d:\Gawean Rebinmas\Driver_Epson_L120"
pip install -r requirements.txt
```

If you encounter permission issues:
```cmd
pip install --user -r requirements.txt
```

### 4. Verify Installation

Run the test script to verify everything is working:

```cmd
python test_system.py
```

This will check:
- All required dependencies
- File structure
- Basic functionality
- Available printers

## Configuration

### 1. Basic Configuration

The system uses `config.yaml` for configuration. Default settings should work for most users, but you can customize:

```yaml
server:
  host: "0.0.0.0"  # Listen on all interfaces
  port: 8080        # Server port
  debug: false      # Set to true for development

printer:
  default_settings:
    color_mode: "color"
    copies: 1
    paper_size: "A4"
    orientation: "portrait"
    quality: "normal"
```

### 2. Network Configuration

For network sharing, ensure:
- Windows Firewall allows the application
- Port 8080 is open (or your configured port)
- Network discovery is enabled

### 3. Printer Setup

Ensure printers are:
- Properly installed on the server machine
- Set as shared (if using Windows sharing)
- Accessible to the user running the server

## Running the System

### Method 1: Using Batch Files (Recommended)

**Start Server:**
```cmd
start_server.bat
```

**Start Client:**
```cmd
start_client.bat
```

### Method 2: Manual Start

**Start Server:**
```cmd
python -m server.main
```

**Start Client GUI:**
```cmd
python -m client.gui
```

**Start Client CLI:**
```cmd
python -m client.cli --help
```

### Method 3: Using Setup.py

```cmd
pip install -e .
printer-server
printer-client
printer-gui
```

## First Time Setup

### 1. Start the Server

1. Run `start_server.bat` or `python -m server.main`
2. Wait for "Server started" message
3. Note the server URL (usually http://localhost:8080)

### 2. Access Web Interface

1. Open web browser
2. Go to http://localhost:8080
3. You should see the Printer Control Panel
4. Verify that printers are detected

### 3. Test Client Connection

1. Run `start_client.bat` or `python -m client.gui`
2. Click "Discover Servers" or enter server IP manually
3. Connect to the server
4. Verify printer list appears

### 4. Test Printing

1. Select a printer
2. Choose a test document or create a simple text file
3. Configure print settings
4. Submit print job
5. Verify document prints successfully

## Troubleshooting

### Common Issues

**"No printers found"**
- Ensure printers are installed and accessible
- Run as Administrator
- Check printer status in Windows

**"Server connection failed"**
- Verify server is running
- Check firewall settings
- Ensure correct IP address and port

**"Permission denied"**
- Run as Administrator
- Check user permissions for printer access

**"Module not found" errors**
- Reinstall dependencies: `pip install -r requirements.txt`
- Check Python PATH

### Network Issues

**Server not discoverable**
- Check Windows Firewall
- Ensure network discovery is enabled
- Verify mDNS/Bonjour service is running

**Slow discovery**
- Check network configuration
- Reduce discovery timeout in config

### Performance Issues

**Slow printing**
- Check network bandwidth
- Reduce file size or quality
- Use local printer drivers when possible

**High memory usage**
- Reduce max concurrent jobs in config
- Clear old print jobs regularly

## Advanced Configuration

### Security Settings

```yaml
security:
  enable_auth: true
  api_key: "your-secret-key"
  allowed_ips:
    - "192.168.1.0/24"
```

### Logging Configuration

```yaml
logging:
  level: "DEBUG"  # For troubleshooting
  file: "logs/printer_share.log"
```

### Custom Printer Settings

```yaml
printer:
  overrides:
    "HP LaserJet":
      color_mode: "grayscale"
      quality: "high"
    "Canon Inkjet":
      color_mode: "color"
      quality: "normal"
```

## Uninstallation

1. Stop all running services
2. Remove the project directory
3. Uninstall Python packages (optional):
   ```cmd
   pip uninstall -r requirements.txt
   ```

## Getting Help

- Check the logs in `logs/` directory
- Run `python test_system.py` for diagnostics
- Review configuration in `config.yaml`
- Check Windows Event Viewer for system errors

## Next Steps

- Read [USER_GUIDE.md](USER_GUIDE.md) for usage instructions
- See [API_DOCUMENTATION.md](API_DOCUMENTATION.md) for API details
- Check [TROUBLESHOOTING.md](TROUBLESHOOTING.md) for common issues
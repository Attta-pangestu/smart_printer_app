# User Guide - Printer Sharing System

This guide explains how to use the Printer Sharing System for sharing printers across your network.

## Overview

The Printer Sharing System consists of:
- **Server**: Manages printers and print jobs
- **Web Interface**: Browser-based control panel
- **Client Applications**: Desktop and command-line tools

## Getting Started

### Starting the Server

1. **Using Batch File (Recommended)**:
   - Double-click `start_server.bat`
   - Wait for "Server started" message

2. **Manual Start**:
   ```cmd
   python -m server.main
   ```

3. **Verify Server is Running**:
   - Open browser to http://localhost:8080
   - You should see the Printer Control Panel

### Accessing the Web Interface

1. Open your web browser
2. Navigate to: `http://[server-ip]:8080`
   - Local access: `http://localhost:8080`
   - Network access: `http://192.168.1.100:8080` (replace with actual IP)

## Web Interface Guide

### Dashboard Overview

The main dashboard shows:
- **System Status**: Server health and statistics
- **Network Info**: Server IP and discovery status
- **Available Printers**: List of detected printers
- **Print Jobs**: Current and recent print jobs

### Managing Printers

#### Viewing Printer Information
1. Click on any printer card
2. View detailed information:
   - Printer status (Ready, Busy, Error)
   - Capabilities (Color, Duplex, Paper sizes)
   - Current jobs
   - Driver information

#### Testing Printers
1. Click "Test Print" on any printer
2. A test page will be sent to the printer
3. Check the job status in the Print Jobs section

#### Printer Discovery
1. Click "Scan Network" to discover new printers
2. Wait for scan to complete
3. New printers will appear automatically

### Printing Documents

#### Upload and Print
1. **Upload File**:
   - Drag and drop files to the upload area
   - Or click "Choose Files" to browse
   - Supported formats: PDF, DOC, DOCX, TXT, JPG, PNG

2. **Configure Print Settings**:
   - **Printer**: Select target printer
   - **Copies**: Number of copies (1-99)
   - **Color Mode**: Color, Grayscale, or Black & White
   - **Paper Size**: A4, Letter, Legal, etc.
   - **Orientation**: Portrait or Landscape
   - **Quality**: Draft, Normal, or High
   - **Duplex**: None, Long Edge, or Short Edge

3. **Submit Job**:
   - Click "Print Document"
   - Job will appear in the Print Jobs list
   - Monitor progress in real-time

#### Print Settings Explained

**Color Mode**:
- **Color**: Full color printing (if supported)
- **Grayscale**: Black and white with shades of gray
- **Black & White**: Pure black and white only

**Paper Sizes**:
- **A4**: 210 × 297 mm (standard in most countries)
- **Letter**: 8.5 × 11 inches (US standard)
- **Legal**: 8.5 × 14 inches
- **A3**: 297 × 420 mm (larger format)

**Quality Settings**:
- **Draft**: Fast, lower quality (saves ink/toner)
- **Normal**: Balanced speed and quality
- **High**: Best quality (slower, uses more ink/toner)

**Duplex (Double-sided) Printing**:
- **None**: Single-sided printing
- **Long Edge**: Flip along long edge (book style)
- **Short Edge**: Flip along short edge (calendar style)

### Managing Print Jobs

#### Job Status
- **Pending**: Waiting to be processed
- **Processing**: Currently being prepared
- **Printing**: Sent to printer
- **Completed**: Successfully printed
- **Failed**: Error occurred
- **Cancelled**: Manually cancelled

#### Job Actions
1. **Cancel Job**: Stop a pending or processing job
2. **Retry Job**: Retry a failed job
3. **View Details**: See job information and error messages

#### Job History
- View recent print jobs
- Filter by status or printer
- See detailed job information

## Client Applications

### Desktop Client (GUI)

#### Starting the Client
1. **Using Batch File**: Double-click `start_client.bat`
2. **Manual Start**: `python -m client.gui`

#### Connecting to Server
1. **Auto Discovery**:
   - Click "Discover Servers"
   - Select server from list
   - Click "Connect"

2. **Manual Connection**:
   - Enter server IP address
   - Enter port (default: 8080)
   - Click "Connect"

#### Using the Desktop Client
1. **Server Tab**: Connect and view server status
2. **Printers Tab**: View and test printers
3. **Print Tab**: Upload and print documents
4. **Jobs Tab**: Monitor print jobs
5. **Logs Tab**: View application logs

### Command Line Client (CLI)

#### Basic Commands

**Discover Servers**:
```cmd
python -m client.cli discover
```

**List Printers**:
```cmd
python -m client.cli printers --server 192.168.1.100
```

**Print File**:
```cmd
python -m client.cli print document.pdf --server 192.168.1.100 --printer "HP LaserJet"
```

**Print with Options**:
```cmd
python -m client.cli print document.pdf \
  --server 192.168.1.100 \
  --printer "Canon Inkjet" \
  --copies 2 \
  --color-mode color \
  --paper-size A4 \
  --quality high
```

**Monitor Jobs**:
```cmd
python -m client.cli jobs --server 192.168.1.100
```

**Cancel Job**:
```cmd
python -m client.cli cancel JOB_ID --server 192.168.1.100
```

#### CLI Options

**Print Options**:
- `--copies N`: Number of copies
- `--color-mode MODE`: color, grayscale, bw
- `--paper-size SIZE`: A4, Letter, Legal, A3
- `--orientation ORIENT`: portrait, landscape
- `--quality QUALITY`: draft, normal, high
- `--duplex MODE`: none, long-edge, short-edge

**Connection Options**:
- `--server IP`: Server IP address
- `--port PORT`: Server port (default: 8080)
- `--timeout SECONDS`: Connection timeout

## Network Setup

### Server Configuration

1. **Configure Network Access**:
   - Edit `config.yaml`
   - Set `server.host` to `"0.0.0.0"` for network access
   - Choose appropriate port (default: 8080)

2. **Firewall Configuration**:
   - Allow the application through Windows Firewall
   - Open the configured port (8080)

3. **Network Discovery**:
   - Ensure mDNS/Bonjour is enabled
   - Check network discovery settings

### Client Configuration

1. **Automatic Discovery**:
   - Clients will automatically find servers on the same network
   - Uses mDNS for service discovery

2. **Manual Configuration**:
   - Note the server's IP address
   - Use IP address in client applications

## Tips and Best Practices

### Performance Optimization

1. **File Sizes**:
   - Compress large images before printing
   - Use PDF format for documents when possible
   - Avoid very large files (>50MB)

2. **Network Performance**:
   - Use wired connections for better reliability
   - Ensure good WiFi signal strength
   - Consider printer location relative to server

3. **Print Settings**:
   - Use "Draft" quality for internal documents
   - Use "Normal" quality for most purposes
   - Reserve "High" quality for important documents

### Security Considerations

1. **Network Security**:
   - Use on trusted networks only
   - Consider enabling authentication in config
   - Restrict access by IP if needed

2. **File Security**:
   - Files are temporarily stored on server
   - Files are automatically cleaned up
   - Sensitive documents should be printed locally

### Troubleshooting Common Issues

**"No printers found"**:
- Check printer installation on server
- Ensure printers are powered on
- Restart the server application

**"Connection failed"**:
- Verify server is running
- Check IP address and port
- Check firewall settings

**"Print job failed"**:
- Check printer status
- Verify file format is supported
- Check printer paper and ink/toner

**"Slow printing"**:
- Check network connection
- Reduce file size or quality
- Check printer queue

## Advanced Features

### Batch Printing

Using CLI for multiple files:
```cmd
for %f in (*.pdf) do python -m client.cli print "%f" --server 192.168.1.100
```

### Custom Print Profiles

Edit `config.yaml` to create printer-specific defaults:
```yaml
printer:
  overrides:
    "Color Printer":
      color_mode: "color"
      quality: "high"
    "Draft Printer":
      color_mode: "grayscale"
      quality: "draft"
```

### API Integration

The system provides a REST API for custom integrations:
- API documentation available at `/docs` endpoint
- Use for custom applications or scripts
- Supports all printer and job management functions

## Support and Maintenance

### Log Files

- Server logs: `logs/printer_share.log`
- Check logs for error messages and debugging
- Adjust log level in `config.yaml` if needed

### Regular Maintenance

1. **Clean up old files**:
   - Server automatically cleans temporary files
   - Adjust cleanup settings in config if needed

2. **Monitor disk space**:
   - Check available space in upload directory
   - Clear old log files if needed

3. **Update software**:
   - Keep Python and dependencies updated
   - Check for system updates regularly

### Getting Help

- Check log files for error messages
- Run `python test_system.py` for diagnostics
- Review configuration settings
- Consult installation guide for setup issues
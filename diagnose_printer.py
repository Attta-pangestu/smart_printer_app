import win32print
import win32api
import time
import subprocess

def check_printer_detailed():
    printer_name = "EPSON L120 Series"
    print(f"=== DETAILED PRINTER DIAGNOSIS: {printer_name} ===")
    
    try:
        # Get printer handle
        printer_handle = win32print.OpenPrinter(printer_name)
        print(f"✓ Printer handle opened successfully")
        
        # Get printer info level 2 (detailed info)
        printer_info = win32print.GetPrinter(printer_handle, 2)
        print(f"\n=== PRINTER CONFIGURATION ===")
        print(f"Printer Name: {printer_info['pPrinterName']}")
        print(f"Share Name: {printer_info['pShareName']}")
        print(f"Port Name: {printer_info['pPortName']}")
        print(f"Driver Name: {printer_info['pDriverName']}")
        print(f"Comment: {printer_info['pComment']}")
        print(f"Location: {printer_info['pLocation']}")
        print(f"Status: {printer_info['Status']}")
        print(f"Jobs: {printer_info['cJobs']}")
        print(f"Attributes: {printer_info['Attributes']}")
        print(f"Priority: {printer_info['Priority']}")
        print(f"Default Priority: {printer_info['DefaultPriority']}")
        
        # Check printer capabilities
        print(f"\n=== PRINTER CAPABILITIES ===")
        try:
            caps = win32print.GetPrinter(printer_handle, 8)
            print(f"Device Capabilities available")
        except Exception as e:
            print(f"Could not get device capabilities: {e}")
        
        # Try to send a simple ESC/P command (Epson printer language)
        print(f"\n=== TESTING RAW PRINTER COMMUNICATION ===")
        
        # ESC/P commands for EPSON printers
        test_commands = [
            b'\x1B@',  # ESC @ - Initialize printer
            b'\x1B\x69\x01\x00',  # ESC i - Print information
            b'Test Print from Python\r\n',  # Simple text
            b'\x0C',  # Form feed
        ]
        
        for i, cmd in enumerate(test_commands):
            try:
                job_info = ('Raw Test Job', None, 'RAW')
                job_id = win32print.StartDocPrinter(printer_handle, 1, job_info)
                print(f"Started raw job {i+1} with ID: {job_id}")
                
                win32print.StartPagePrinter(printer_handle)
                bytes_written = win32print.WritePrinter(printer_handle, cmd)
                print(f"Wrote {bytes_written} bytes for command {i+1}")
                
                win32print.EndPagePrinter(printer_handle)
                win32print.EndDocPrinter(printer_handle)
                print(f"✓ Raw command {i+1} sent successfully")
                
                time.sleep(1)  # Wait between commands
                
            except Exception as e:
                print(f"✗ Failed to send raw command {i+1}: {e}")
        
        # Check if printer responds to status queries
        print(f"\n=== PRINTER STATUS MONITORING ===")
        for check in range(5):
            try:
                current_info = win32print.GetPrinter(printer_handle, 2)
                print(f"Check {check+1}: Jobs={current_info['cJobs']}, Status={current_info['Status']}")
                time.sleep(2)
            except Exception as e:
                print(f"Status check {check+1} failed: {e}")
        
        win32print.ClosePrinter(printer_handle)
        print(f"\n✓ Printer handle closed")
        
    except Exception as e:
        print(f"✗ Error accessing printer: {e}")
    
    # Try alternative method - direct port communication
    print(f"\n=== TESTING PORT COMMUNICATION ===")
    try:
        # Get port info
        ports = win32print.EnumPorts(None, 1)
        for port in ports:
            if 'USB' in port['pName'] and 'EPSON' in port['pName']:
                print(f"Found EPSON USB port: {port['pName']}")
                print(f"Port type: {port['fPortType']}")
                print(f"Port status: {port['Status']}")
    except Exception as e:
        print(f"Could not enumerate ports: {e}")
    
    # Check Windows printer troubleshooter
    print(f"\n=== SYSTEM PRINTER STATUS ===")
    try:
        result = subprocess.run(['powershell', '-Command', 
                               f'Get-Printer -Name \"{printer_name}\" | Select-Object Name, PrinterStatus, JobCount'], 
                              capture_output=True, text=True, timeout=10)
        if result.returncode == 0:
            print("PowerShell printer status:")
            print(result.stdout)
        else:
            print(f"PowerShell command failed: {result.stderr}")
    except Exception as e:
        print(f"Could not run PowerShell command: {e}")

if __name__ == "__main__":
    check_printer_detailed()
    print("\n=== DIAGNOSIS COMPLETE ===")
    print("If printer still doesn't respond physically:")
    print("1. Check printer power and USB connection")
    print("2. Try printing from another application (Notepad, etc.)")
    print("3. Check printer ink levels and paper")
    print("4. Restart printer and try again")
    print("5. Check if printer is in error state (blinking lights)")
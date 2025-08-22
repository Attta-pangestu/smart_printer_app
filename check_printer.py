import win32print
import win32con

def check_epson_printer():
    try:
        # Find EPSON printers
        printers = win32print.EnumPrinters(2)
        epson_printers = [p for p in printers if 'epson' in p[2].lower() or 'l120' in p[2].lower()]
        
        print('EPSON Printers found:')
        for p in epson_printers:
            print(f'  - {p[2]}')
        
        if not epson_printers:
            print('No EPSON printers found!')
            return
        
        # Test EPSON L120 connection
        printer_name = 'EPSON L120 Series'
        print(f'\nTesting {printer_name} connection:')
        
        try:
            handle = win32print.OpenPrinter(printer_name)
            info = win32print.GetPrinter(handle, 2)
            
            status = info['Status']
            jobs_count = info.get('cJobs', 0)
            port = info.get('pPortName', 'Unknown')
            
            print(f'Status Code: {status}')
            print(f'Jobs in Queue: {jobs_count}')
            print(f'Port: {port}')
            
            # Decode status flags using correct constants
            status_flags = []
            if status & 0x00000400:  # PRINTER_STATUS_OFFLINE
                status_flags.append('OFFLINE')
            if status & 0x00000002:  # PRINTER_STATUS_ERROR
                status_flags.append('ERROR')
            if status & 0x00000200:  # PRINTER_STATUS_BUSY
                status_flags.append('BUSY')
            if status & 0x00000001:  # PRINTER_STATUS_PAUSED
                status_flags.append('PAUSED')
            if status == 0:
                status_flags.append('READY')
            
            print(f'Status Flags: {", ".join(status_flags) if status_flags else "READY"}')
            
            # Check if printer is physically connected
            if 'USB' in port.upper():
                print('✓ Printer connected via USB')
            else:
                print(f'⚠ Printer port: {port} (may not be USB)')
            
            win32print.ClosePrinter(handle)
            
        except Exception as e:
            print(f'Error accessing printer: {e}')
            
    except Exception as e:
        print(f'Error checking printers: {e}')

if __name__ == '__main__':
    check_epson_printer()
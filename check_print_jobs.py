import win32print
import time

def check_and_manage_print_jobs():
    printer_name = "EPSON L120 Series"
    
    try:
        handle = win32print.OpenPrinter(printer_name)
        info = win32print.GetPrinter(handle, 2)
        
        print(f"Printer: {printer_name}")
        print(f"Status Code: {info['Status']}")
        print(f"Jobs in Queue: {info.get('cJobs', 0)}")
        print(f"Port: {info.get('pPortName', 'Unknown')}")
        
        # Get detailed printer attributes
        attributes = info.get('Attributes', 0)
        print(f"Attributes: {attributes}")
        
        # Check specific attributes
        if attributes & 0x00000001:  # PRINTER_ATTRIBUTE_QUEUED
            print("✓ Printer is queued")
        if attributes & 0x00000002:  # PRINTER_ATTRIBUTE_DIRECT
            print("✓ Printer uses direct printing")
        if attributes & 0x00000004:  # PRINTER_ATTRIBUTE_DEFAULT
            print("✓ This is the default printer")
        if attributes & 0x00000008:  # PRINTER_ATTRIBUTE_SHARED
            print("✓ Printer is shared")
        if attributes & 0x00000400:  # PRINTER_ATTRIBUTE_WORK_OFFLINE
            print("⚠ Printer is set to work offline!")
        
        jobs_count = info.get('cJobs', 0)
        if jobs_count > 0:
            print(f"\nDetailed job information:")
            try:
                jobs = win32print.EnumJobs(handle, 0, -1, 1)
                for i, job in enumerate(jobs, 1):
                    print(f"\nJob {i}:")
                    print(f"  Job ID: {job.get('JobId', 'Unknown')}")
                    print(f"  Document: {job.get('pDocument', 'Unknown')}")
                    print(f"  Status: {job.get('Status', 0)}")
                    print(f"  Pages Printed: {job.get('PagesPrinted', 0)}")
                    print(f"  Total Pages: {job.get('TotalPages', 0)}")
                    print(f"  User: {job.get('pUserName', 'Unknown')}")
                    print(f"  Machine: {job.get('pMachineName', 'Unknown')}")
                    print(f"  Submitted: {job.get('Submitted', 'Unknown')}")
                    
                    # Decode job status
                    job_status = job.get('Status', 0)
                    status_flags = []
                    if job_status & 0x00000001:  # JOB_STATUS_PAUSED
                        status_flags.append('PAUSED')
                    if job_status & 0x00000002:  # JOB_STATUS_ERROR
                        status_flags.append('ERROR')
                    if job_status & 0x00000004:  # JOB_STATUS_DELETING
                        status_flags.append('DELETING')
                    if job_status & 0x00000008:  # JOB_STATUS_SPOOLING
                        status_flags.append('SPOOLING')
                    if job_status & 0x00000010:  # JOB_STATUS_PRINTING
                        status_flags.append('PRINTING')
                    if job_status & 0x00000020:  # JOB_STATUS_OFFLINE
                        status_flags.append('OFFLINE')
                    if job_status & 0x00000040:  # JOB_STATUS_PAPEROUT
                        status_flags.append('PAPEROUT')
                    if job_status & 0x00000080:  # JOB_STATUS_PRINTED
                        status_flags.append('PRINTED')
                    if job_status & 0x00000100:  # JOB_STATUS_DELETED
                        status_flags.append('DELETED')
                    if job_status & 0x00000200:  # JOB_STATUS_BLOCKED_DEVQ
                        status_flags.append('BLOCKED')
                    if job_status & 0x00000400:  # JOB_STATUS_USER_INTERVENTION
                        status_flags.append('USER_INTERVENTION')
                    if job_status & 0x00000800:  # JOB_STATUS_RESTART
                        status_flags.append('RESTART')
                    
                    if status_flags:
                        print(f"  Status Flags: {', '.join(status_flags)}")
                    else:
                        print(f"  Status Flags: NORMAL (code: {job_status})")
                        
            except Exception as e:
                print(f"Error getting job details: {e}")
        
        # Option to cancel stuck jobs
        if jobs_count > 0:
            print(f"\nFound {jobs_count} jobs in queue.")
            response = input("Do you want to cancel all jobs? (y/n): ")
            if response.lower() == 'y':
                try:
                    jobs = win32print.EnumJobs(handle, 0, -1, 1)
                    for job in jobs:
                        job_id = job.get('JobId')
                        if job_id:
                            win32print.SetJob(handle, job_id, 0, None, win32print.JOB_CONTROL_CANCEL)
                            print(f"Cancelled job {job_id}")
                    print("All jobs cancelled.")
                except Exception as e:
                    print(f"Error cancelling jobs: {e}")
        
        win32print.ClosePrinter(handle)
        
    except Exception as e:
        print(f"Error: {e}")

if __name__ == '__main__':
    check_and_manage_print_jobs()
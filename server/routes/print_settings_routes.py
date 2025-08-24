from fastapi import APIRouter, HTTPException, UploadFile, File
from fastapi.responses import FileResponse
from pydantic import BaseModel
from typing import Optional, Dict, Any
import win32com.client
import os
import tempfile
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/api/print-settings", tags=["print-settings"])

class PrintSettings(BaseModel):
    orientation: str = "portrait"  # portrait or landscape
    paper_size: str = "A4"  # A4, Letter, Executive, Legal, A3
    margins: str = "normal"  # normal, narrow, wide, custom
    scaling: str = "fit_to_page"  # fit_to_page, actual_size, custom_scale
    custom_margins: Optional[Dict[str, float]] = None
    custom_scale: Optional[int] = None
    # New fields for sheet and range selection
    selected_sheets: Optional[list[str]] = None  # List of sheet names to export
    print_ranges: Optional[Dict[str, str]] = None  # Dict of sheet_name: range (e.g., "Sheet1": "A1:D20")
    export_mode: str = "all_sheets"  # all_sheets, selected_sheets, specific_ranges

class PrintJobRequest(BaseModel):
    file_id: str
    file_type: str
    print_settings: PrintSettings

# Paper size constants for Excel
PAPER_SIZES = {
    "A4": 9,
    "Letter": 1,
    "Executive": 7,
    "Legal": 5,
    "A3": 8
}

# Margin presets (in inches)
MARGIN_PRESETS = {
    "normal": {"top": 1.0, "bottom": 1.0, "left": 1.0, "right": 1.0},
    "narrow": {"top": 0.5, "bottom": 0.5, "left": 0.5, "right": 0.5},
    "wide": {"top": 1.5, "bottom": 1.5, "left": 1.5, "right": 1.5}
}

@router.post("/apply-excel-settings")
async def apply_excel_print_settings(request: PrintJobRequest):
    """Apply print settings to Excel file and convert to PDF"""
    try:
        # Get file path
        upload_dir = Path("uploads")
        temp_dir = Path("temp")
        
        file_path = upload_dir / request.file_id
        if not file_path.exists():
            file_path = temp_dir / request.file_id
            
        if not file_path.exists():
            raise HTTPException(status_code=404, detail="File not found")
        
        # Generate output PDF path
        output_filename = f"{file_path.stem}_formatted.pdf"
        output_path = temp_dir / output_filename
        
        # Apply Excel print settings using pywin32
        success = apply_excel_settings_pywin32(
            str(file_path), 
            str(output_path), 
            request.print_settings
        )
        
        if not success:
            raise HTTPException(status_code=500, detail="Failed to apply print settings")
        
        return {
            "success": True,
            "message": "Print settings applied successfully",
            "output_file": output_filename,
            "download_url": f"/api/files/download/{output_filename}"
        }
        
    except Exception as e:
        logger.error(f"Error applying Excel print settings: {e}")
        raise HTTPException(status_code=500, detail=f"Error: {str(e)}")

def apply_excel_settings_pywin32(excel_file: str, output_pdf: str, settings: PrintSettings) -> bool:
    """Apply print settings to Excel file using pywin32"""
    excel = None
    workbook = None
    
    try:
        # Open Excel application
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        # Open the Excel workbook
        workbook = excel.Workbooks.Open(excel_file)
        
        # Determine which sheets to process
        sheets_to_process = []
        
        if settings.export_mode == "selected_sheets" and settings.selected_sheets:
            # Process only selected sheets
            for sheet_name in settings.selected_sheets:
                try:
                    sheet = workbook.Sheets[sheet_name]
                    sheets_to_process.append(sheet)
                except Exception as e:
                    logger.warning(f"Sheet '{sheet_name}' not found: {e}")
        elif settings.export_mode == "specific_ranges" and settings.print_ranges:
            # Process sheets with specific ranges
            for sheet_name in settings.print_ranges.keys():
                try:
                    sheet = workbook.Sheets[sheet_name]
                    sheets_to_process.append(sheet)
                except Exception as e:
                    logger.warning(f"Sheet '{sheet_name}' not found: {e}")
        else:
            # Process all sheets (default behavior)
            for i in range(1, workbook.Sheets.Count + 1):
                sheets_to_process.append(workbook.Sheets[i])
        
        # Apply settings to each sheet
        for sheet in sheets_to_process:
            # Set page orientation
            if settings.orientation == "landscape":
                sheet.PageSetup.Orientation = 2  # xlLandscape
            else:
                sheet.PageSetup.Orientation = 1  # xlPortrait
            
            # Set paper size
            paper_size_code = PAPER_SIZES.get(settings.paper_size, 9)  # Default to A4
            sheet.PageSetup.PaperSize = paper_size_code
            
            # Set margins
            if settings.margins == "custom" and settings.custom_margins:
                margins = settings.custom_margins
            else:
                margins = MARGIN_PRESETS.get(settings.margins, MARGIN_PRESETS["normal"])
            
            # Convert inches to points (1 inch = 72 points)
            sheet.PageSetup.TopMargin = margins["top"] * 72
            sheet.PageSetup.BottomMargin = margins["bottom"] * 72
            sheet.PageSetup.LeftMargin = margins["left"] * 72
            sheet.PageSetup.RightMargin = margins["right"] * 72
            
            # Set scaling
            if settings.scaling == "fit_to_page":
                sheet.PageSetup.FitToPagesWide = 1
                sheet.PageSetup.FitToPagesTall = 1
                sheet.PageSetup.Zoom = False  # Disable zoom when using fit to page
            elif settings.scaling == "custom_scale" and settings.custom_scale:
                sheet.PageSetup.Zoom = settings.custom_scale
                sheet.PageSetup.FitToPagesWide = False
                sheet.PageSetup.FitToPagesTall = False
            else:  # actual_size
                sheet.PageSetup.Zoom = 100
                sheet.PageSetup.FitToPagesWide = False
                sheet.PageSetup.FitToPagesTall = False
            
            # Set print area if specific ranges are defined
            if settings.export_mode == "specific_ranges" and settings.print_ranges:
                sheet_name = sheet.Name
                if sheet_name in settings.print_ranges:
                    print_range = settings.print_ranges[sheet_name]
                    sheet.PageSetup.PrintArea = print_range
                    logger.info(f"Set print area for sheet '{sheet_name}': {print_range}")
                else:
                    # Clear print area if not specified
                    sheet.PageSetup.PrintArea = ""
            else:
                # Clear print area for full sheet export
                sheet.PageSetup.PrintArea = ""
        
        # Export as PDF
        workbook.ExportAsFixedFormat(0, output_pdf)  # 0 = xlTypePDF
        
        logger.info(f"Excel file successfully converted to PDF with custom settings: {output_pdf}")
        logger.info(f"Export mode: {settings.export_mode}, Processed {len(sheets_to_process)} sheets")
        return True
        
    except Exception as e:
        logger.error(f"Error in apply_excel_settings_pywin32: {e}")
        return False
        
    finally:
        # Clean up
        try:
            if workbook:
                workbook.Close(False)
            if excel:
                excel.Quit()
        except Exception as e:
            logger.error(f"Error closing Excel application: {e}")

@router.post("/apply-pdf-settings")
async def apply_pdf_print_settings(request: PrintJobRequest):
    """Apply print settings to PDF file (for future implementation)"""
    try:
        # For now, just return the settings that would be applied
        # In the future, this could use libraries like PyPDF2 or reportlab
        # to modify PDF print settings
        
        return {
            "success": True,
            "message": "PDF print settings received (implementation pending)",
            "settings_applied": request.print_settings.dict(),
            "note": "PDF print settings will be applied during actual printing"
        }
        
    except Exception as e:
        logger.error(f"Error applying PDF print settings: {e}")
        raise HTTPException(status_code=500, detail=f"Error: {str(e)}")

@router.get("/paper-sizes")
async def get_available_paper_sizes():
    """Get list of available paper sizes"""
    return {
        "paper_sizes": [
            {"value": "A4", "label": "A4 (210 x 297 mm)", "code": 9},
            {"value": "Letter", "label": "Letter (8.5 x 11 in)", "code": 1},
            {"value": "Executive", "label": "Executive (7.25 x 10.5 in)", "code": 7},
            {"value": "Legal", "label": "Legal (8.5 x 14 in)", "code": 5},
            {"value": "A3", "label": "A3 (297 x 420 mm)", "code": 8}
        ]
    }

@router.get("/margin-presets")
async def get_margin_presets():
    """Get available margin presets"""
    return {
        "margin_presets": MARGIN_PRESETS
    }

@router.get("/excel-sheets/{file_id}")
async def get_excel_sheets(file_id: str):
    """Get list of sheet names from Excel file"""
    try:
        # Get file path
        upload_dir = Path("uploads")
        temp_dir = Path("temp")
        
        file_path = upload_dir / file_id
        if not file_path.exists():
            file_path = temp_dir / file_id
            
        if not file_path.exists():
            raise HTTPException(status_code=404, detail="File not found")
        
        # Get sheet names using pywin32
        sheet_names = get_excel_sheet_names(str(file_path))
        
        return {
            "success": True,
            "sheets": sheet_names
        }
        
    except Exception as e:
        logger.error(f"Error getting Excel sheets: {e}")
        raise HTTPException(status_code=500, detail=f"Error: {str(e)}")

def get_excel_sheet_names(excel_file: str) -> list[str]:
    """Get list of sheet names from Excel file using pywin32"""
    excel = None
    workbook = None
    
    try:
        # Open Excel application
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        # Open the Excel workbook
        workbook = excel.Workbooks.Open(excel_file)
        
        # Get sheet names
        sheet_names = []
        for i in range(1, workbook.Sheets.Count + 1):
            sheet_names.append(workbook.Sheets[i].Name)
        
        return sheet_names
        
    except Exception as e:
        logger.error(f"Error in get_excel_sheet_names: {e}")
        return []
        
    finally:
        # Clean up
        try:
            if workbook:
                workbook.Close(False)
            if excel:
                excel.Quit()
        except Exception as e:
            logger.error(f"Error closing Excel application: {e}")
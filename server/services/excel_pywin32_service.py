import win32com.client
import os
import logging
from pathlib import Path
from typing import Dict, List, Optional, Any, Tuple
from datetime import datetime
import tempfile
import pythoncom
import threading

logger = logging.getLogger(__name__)

class ExcelPyWin32Service:
    """Service untuk konversi Excel ke PDF menggunakan pywin32 dengan fitur advanced"""
    
    def __init__(self, temp_dir: str = None):
        self.temp_dir = Path(temp_dir) if temp_dir else Path(tempfile.gettempdir())
        self.temp_dir.mkdir(exist_ok=True)
        
        # Excel constants
        self.xlTypePDF = 0
        self.xlQualityStandard = 0
        self.xlQualityMinimum = 1
        
        # Page orientation constants
        self.xlPortrait = 1
        self.xlLandscape = 2
        
        # Paper size constants
        self.xlPaperA4 = 9
        self.xlPaperLetter = 1
        
    def _initialize_excel(self) -> win32com.client.Dispatch:
        """Initialize Excel application with COM"""
        try:
            # Initialize COM for this thread
            pythoncom.CoInitialize()
            
            # Create Excel application
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            return excel
        except Exception as e:
            logger.error(f"Failed to initialize Excel: {e}")
            raise
    
    def _cleanup_excel(self, excel, workbook=None):
        """Clean up Excel resources"""
        try:
            if workbook:
                workbook.Close(False)
            if excel:
                excel.Quit()
            pythoncom.CoUninitialize()
        except Exception as e:
            logger.warning(f"Error during Excel cleanup: {e}")
    
    def get_excel_info(self, file_path: str) -> Dict[str, Any]:
        """Get information about Excel file (sheets, ranges, etc.)"""
        excel = None
        workbook = None
        
        try:
            # Convert to absolute path
            abs_file_path = os.path.abspath(file_path)
            if not os.path.exists(abs_file_path):
                raise FileNotFoundError(f"File not found: {abs_file_path}")
                
            excel = self._initialize_excel()
            workbook = excel.Workbooks.Open(abs_file_path)
            
            sheets_info = []
            for sheet in workbook.Worksheets:
                # Get used range info
                used_range = sheet.UsedRange
                if used_range:
                    last_row = used_range.Rows.Count
                    last_col = used_range.Columns.Count
                    range_address = f"A1:{self._get_column_letter(last_col)}{last_row}"
                else:
                    last_row = 0
                    last_col = 0
                    range_address = "A1:A1"
                
                sheets_info.append({
                    "name": sheet.Name,
                    "index": sheet.Index,
                    "used_range": range_address,
                    "row_count": last_row,
                    "column_count": last_col,
                    "visible": sheet.Visible
                })
            
            return {
                "success": True,
                "file_name": Path(file_path).name,
                "sheets": sheets_info,
                "total_sheets": len(sheets_info)
            }
            
        except Exception as e:
            logger.error(f"Error getting Excel info: {e}")
            return {
                "success": False,
                "error": str(e)
            }
        finally:
            self._cleanup_excel(excel, workbook)
    
    def _get_column_letter(self, col_num: int) -> str:
        """Convert column number to Excel column letter"""
        result = ""
        while col_num > 0:
            col_num -= 1
            result = chr(col_num % 26 + ord('A')) + result
            col_num //= 26
        return result
    
    def convert_to_pdf(
        self,
        input_path: str,
        output_path: str = None,
        sheet_names: List[str] = None,
        cell_range: str = None,
        print_settings: Dict[str, Any] = None
    ) -> Dict[str, Any]:
        """Convert Excel to PDF with advanced options"""
        
        excel = None
        workbook = None
        
        try:
            # Convert to absolute path
            abs_input_path = os.path.abspath(input_path)
            if not os.path.exists(abs_input_path):
                raise FileNotFoundError(f"Input file not found: {abs_input_path}")
            
            # Generate output path if not provided
            if not output_path:
                input_file = Path(abs_input_path)
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_path = str(self.temp_dir / f"{timestamp}_{input_file.stem}.pdf")
            
            # Ensure output directory exists
            output_dir = os.path.dirname(output_path)
            os.makedirs(output_dir, exist_ok=True)
            
            # Convert to absolute path
            abs_output_path = os.path.abspath(output_path)
            
            # Initialize Excel
            excel = self._initialize_excel()
            workbook = excel.Workbooks.Open(abs_input_path)
            
            # Apply print settings if provided
            if print_settings:
                self._apply_print_settings(workbook, print_settings)
            
            # Handle sheet selection
            if sheet_names:
                # Hide all sheets first
                for sheet in workbook.Worksheets:
                    if sheet.Name not in sheet_names:
                        sheet.Visible = False
                
                # Show only selected sheets
                for sheet_name in sheet_names:
                    try:
                        sheet = workbook.Worksheets(sheet_name)
                        sheet.Visible = True
                        
                        # Apply cell range if specified
                        if cell_range:
                            self._set_print_area(sheet, cell_range)
                            
                    except Exception as e:
                        logger.warning(f"Could not process sheet {sheet_name}: {e}")
            
            # Save workbook before export to ensure it's not in a dirty state
            workbook.Save()
            
            # Export to PDF
            workbook.ExportAsFixedFormat(
                Type=self.xlTypePDF,
                Filename=abs_output_path,
                Quality=self.xlQualityStandard,
                IgnorePrintAreas=False,
                OpenAfterPublish=False
            )
            
            # Verify the file was created
            if not os.path.exists(abs_output_path):
                raise Exception(f"PDF file was not created at {abs_output_path}")
            
            # Get file size
            file_size = os.path.getsize(abs_output_path)
            
            return {
                "success": True,
                "output_path": abs_output_path,
                "file_size": file_size,
                "message": "Excel berhasil dikonversi ke PDF"
            }
            
        except Exception as e:
            logger.error(f"Error converting Excel to PDF: {e}")
            return {
                "success": False,
                "error": str(e),
                "message": f"Gagal mengkonversi Excel ke PDF: {str(e)}"
            }
        finally:
            self._cleanup_excel(excel, workbook)
    
    def _apply_print_settings(self, workbook, settings: Dict[str, Any]):
        """Apply print settings to workbook"""
        try:
            for sheet in workbook.Worksheets:
                if not sheet.Visible:
                    continue
                    
                page_setup = sheet.PageSetup
                
                # Orientation
                if 'orientation' in settings:
                    if settings['orientation'].lower() == 'landscape':
                        page_setup.Orientation = self.xlLandscape
                    else:
                        page_setup.Orientation = self.xlPortrait
                
                # Paper size
                if 'paper_size' in settings:
                    if settings['paper_size'].upper() == 'A4':
                        page_setup.PaperSize = self.xlPaperA4
                    elif settings['paper_size'].upper() == 'LETTER':
                        page_setup.PaperSize = self.xlPaperLetter
                
                # Margins (in inches)
                if 'margins' in settings:
                    margins = settings['margins']
                    if 'top' in margins:
                        page_setup.TopMargin = excel.InchesToPoints(margins['top'])
                    if 'bottom' in margins:
                        page_setup.BottomMargin = excel.InchesToPoints(margins['bottom'])
                    if 'left' in margins:
                        page_setup.LeftMargin = excel.InchesToPoints(margins['left'])
                    if 'right' in margins:
                        page_setup.RightMargin = excel.InchesToPoints(margins['right'])
                
                # Scaling
                if 'scale' in settings:
                    scale = settings['scale']
                    if isinstance(scale, (int, float)) and 10 <= scale <= 400:
                        page_setup.Zoom = scale
                
                # Fit to pages
                if 'fit_to_pages' in settings:
                    fit_settings = settings['fit_to_pages']
                    if 'width' in fit_settings or 'height' in fit_settings:
                        page_setup.Zoom = False
                        page_setup.FitToPagesWide = fit_settings.get('width', 1)
                        page_setup.FitToPagesTall = fit_settings.get('height', 1)
                
                # Headers and footers
                if 'header' in settings:
                    header = settings['header']
                    if 'left' in header:
                        page_setup.LeftHeader = header['left']
                    if 'center' in header:
                        page_setup.CenterHeader = header['center']
                    if 'right' in header:
                        page_setup.RightHeader = header['right']
                
                if 'footer' in settings:
                    footer = settings['footer']
                    if 'left' in footer:
                        page_setup.LeftFooter = footer['left']
                    if 'center' in footer:
                        page_setup.CenterFooter = footer['center']
                    if 'right' in footer:
                        page_setup.RightFooter = footer['right']
                
                # Print quality
                if 'print_quality' in settings:
                    quality = settings['print_quality']
                    if quality in [300, 600, 1200]:
                        page_setup.PrintQuality = quality
                
                # Grid lines
                if 'print_gridlines' in settings:
                    page_setup.PrintGridlines = bool(settings['print_gridlines'])
                
                # Row and column headings
                if 'print_headings' in settings:
                    page_setup.PrintHeadings = bool(settings['print_headings'])
                    
        except Exception as e:
            logger.error(f"Error applying print settings: {e}")
    
    def _set_print_area(self, sheet, cell_range: str):
        """Set print area for a sheet"""
        try:
            sheet.PageSetup.PrintArea = cell_range
        except Exception as e:
            logger.error(f"Error setting print area {cell_range}: {e}")
    
    def preview_pdf_settings(
        self,
        input_path: str,
        sheet_names: List[str] = None,
        cell_range: str = None,
        print_settings: Dict[str, Any] = None
    ) -> Dict[str, Any]:
        """Preview PDF conversion settings without actually converting"""
        
        try:
            # Get Excel file info
            excel_info = self.get_excel_info(input_path)
            if not excel_info['success']:
                return excel_info
            
            # Validate sheet names
            available_sheets = [sheet['name'] for sheet in excel_info['sheets']]
            if sheet_names:
                invalid_sheets = [name for name in sheet_names if name not in available_sheets]
                if invalid_sheets:
                    return {
                        "success": False,
                        "error": f"Sheet tidak ditemukan: {', '.join(invalid_sheets)}"
                    }
            
            # Prepare preview info
            preview_info = {
                "success": True,
                "excel_info": excel_info,
                "selected_sheets": sheet_names or available_sheets,
                "cell_range": cell_range,
                "print_settings": print_settings or {},
                "estimated_pages": self._estimate_page_count(
                    excel_info, sheet_names, cell_range, print_settings
                )
            }
            
            return preview_info
            
        except Exception as e:
            logger.error(f"Error previewing PDF settings: {e}")
            return {
                "success": False,
                "error": str(e)
            }
    
    def _estimate_page_count(
        self,
        excel_info: Dict,
        sheet_names: List[str] = None,
        cell_range: str = None,
        print_settings: Dict[str, Any] = None
    ) -> int:
        """Estimate number of pages in PDF output"""
        try:
            total_pages = 0
            sheets_to_process = sheet_names or [sheet['name'] for sheet in excel_info['sheets']]
            
            for sheet_info in excel_info['sheets']:
                if sheet_info['name'] in sheets_to_process:
                    # Simple estimation based on row/column count
                    rows = sheet_info['row_count']
                    cols = sheet_info['column_count']
                    
                    # Estimate pages based on typical Excel page layout
                    # This is a rough estimation
                    rows_per_page = 50  # Typical for portrait A4
                    cols_per_page = 8   # Typical for portrait A4
                    
                    # Adjust for orientation
                    if print_settings and print_settings.get('orientation') == 'landscape':
                        rows_per_page = 35
                        cols_per_page = 12
                    
                    # Adjust for scaling
                    if print_settings and 'scale' in print_settings:
                        scale = print_settings['scale'] / 100
                        rows_per_page = int(rows_per_page * scale)
                        cols_per_page = int(cols_per_page * scale)
                    
                    page_rows = max(1, (rows + rows_per_page - 1) // rows_per_page)
                    page_cols = max(1, (cols + cols_per_page - 1) // cols_per_page)
                    
                    total_pages += page_rows * page_cols
            
            return max(1, total_pages)
            
        except Exception as e:
            logger.error(f"Error estimating page count: {e}")
            return 1
    
    def get_default_print_settings(self) -> Dict[str, Any]:
        """Get default print settings"""
        return {
            "orientation": "portrait",
            "paper_size": "A4",
            "margins": {
                "top": 0.75,
                "bottom": 0.75,
                "left": 0.7,
                "right": 0.7
            },
            "scale": 100,
            "print_gridlines": False,
            "print_headings": False,
            "print_quality": 600,
            "header": {
                "left": "",
                "center": "",
                "right": ""
            },
            "footer": {
                "left": "",
                "center": "",
                "right": ""
            }
        }
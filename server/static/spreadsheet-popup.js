/**
 * Spreadsheet Popup Module
 * Provides formatted spreadsheet display in popup with preserved Excel formatting
 */

class SpreadsheetPopupManager {
    constructor() {
        this.currentFile = null;
        this.currentSheet = 0;
        this.sheets = [];
        this.zoomLevel = 100;
        this.selectedCells = [];
        this.isModalOpen = false;
    }

    /**
     * Upload and display spreadsheet in popup
     */
    async uploadAndShowSpreadsheet(fileInput) {
        try {
            const file = fileInput.files[0];
            if (!file) {
                this.showError('Silakan pilih file Excel terlebih dahulu');
                return;
            }

            // Validate file type
            if (!file.name.toLowerCase().endsWith('.xlsx') && !file.name.toLowerCase().endsWith('.xls')) {
                this.showError('File harus berformat Excel (.xlsx atau .xls)');
                return;
            }

            this.showLoadingState('Mengupload dan memproses file Excel...');

            // Create FormData
            const formData = new FormData();
            formData.append('file', file);
            formData.append('preserve_formatting', 'true');
            formData.append('max_rows', '1000');
            formData.append('max_columns', '50');

            // Upload file
            const response = await fetch('/api/excel-pywin32/upload-excel', {
                method: 'POST',
                body: formData
            });

            const result = await response.json();

            if (!response.ok) {
                throw new Error(result.detail || `HTTP ${response.status}: Gagal mengupload file`);
            }

            // Store data
            this.currentFile = result;
            this.sheets = result.sheets;
            this.currentSheet = 0;

            this.hideLoadingState();
            this.showSpreadsheetModal(result.filename);
            this.updateSheetTabs();
            this.loadSheetContent(result.sheets[0]);
            this.updateFileInfo(result.file_info);

        } catch (error) {
            this.hideLoadingState();
            console.error('Error uploading spreadsheet:', error);
            this.showError(`Gagal mengupload spreadsheet: ${error.message}`);
        }
    }

    /**
     * Show the spreadsheet modal
     */
    showSpreadsheetModal(fileName) {
        let modal = document.getElementById('spreadsheetModal');
        if (!modal) {
            modal = this.createSpreadsheetModal();
            document.body.appendChild(modal);
        }

        document.getElementById('spreadsheetFileName').textContent = fileName;
        modal.style.display = 'block';
        this.isModalOpen = true;

        // Add escape key listener
        document.addEventListener('keydown', this.handleEscapeKey.bind(this));
    }

    /**
     * Create the spreadsheet modal HTML
     */
    createSpreadsheetModal() {
        const modal = document.createElement('div');
        modal.id = 'spreadsheetModal';
        modal.className = 'spreadsheet-modal';
        modal.innerHTML = `
            <div class="spreadsheet-modal-content">
                <div class="spreadsheet-header">
                    <div class="spreadsheet-title">
                        <h3 id="spreadsheetFileName">Spreadsheet Preview</h3>
                        <div class="spreadsheet-info" id="spreadsheetInfo"></div>
                    </div>
                    <div class="spreadsheet-controls">
                        <div class="zoom-controls">
                            <button onclick="spreadsheetPopup.zoomOut()" title="Zoom Out">-</button>
                            <span id="zoomLevel">100%</span>
                            <button onclick="spreadsheetPopup.zoomIn()" title="Zoom In">+</button>
                            <button onclick="spreadsheetPopup.resetZoom()" title="Reset Zoom">Reset</button>
                        </div>
                        <button class="close-btn" onclick="spreadsheetPopup.closeModal()">&times;</button>
                    </div>
                </div>
                
                <div class="sheet-tabs" id="sheetTabs"></div>
                
                <div class="spreadsheet-container">
                    <div class="spreadsheet-table-wrapper" id="spreadsheetTableWrapper">
                        <table class="spreadsheet-table" id="spreadsheetTable">
                            <thead id="spreadsheetHeader"></thead>
                            <tbody id="spreadsheetBody"></tbody>
                        </table>
                    </div>
                </div>
                
                <div class="spreadsheet-footer">
                    <div class="spreadsheet-actions">
                        <button onclick="spreadsheetPopup.downloadAsExcel()" class="action-btn">Download Excel</button>
                        <button onclick="spreadsheetPopup.showConvertToPdfOptions()" class="action-btn">Convert to PDF</button>
                        <button onclick="spreadsheetPopup.printSpreadsheet()" class="action-btn">Print</button>
                    </div>
                </div>
            </div>
        `;

        this.addSpreadsheetStyles();
        return modal;
    }

    /**
     * Add CSS styles for spreadsheet modal
     */
    addSpreadsheetStyles() {
        if (document.getElementById('spreadsheetStyles')) return;

        const style = document.createElement('style');
        style.id = 'spreadsheetStyles';
        style.textContent = `
            .spreadsheet-modal {
                display: none;
                position: fixed;
                z-index: 10000;
                left: 0;
                top: 0;
                width: 100%;
                height: 100%;
                background-color: rgba(0, 0, 0, 0.8);
                overflow: hidden;
            }

            .spreadsheet-modal-content {
                background-color: #ffffff;
                margin: 2% auto;
                width: 95%;
                height: 95%;
                border-radius: 8px;
                display: flex;
                flex-direction: column;
                box-shadow: 0 4px 20px rgba(0, 0, 0, 0.3);
            }

            .spreadsheet-header {
                display: flex;
                justify-content: space-between;
                align-items: center;
                padding: 15px 20px;
                border-bottom: 2px solid #e0e0e0;
                background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                color: white;
            }

            .spreadsheet-title h3 {
                margin: 0;
                font-size: 1.4em;
                font-weight: 600;
            }

            .spreadsheet-info {
                font-size: 0.9em;
                opacity: 0.9;
                margin-top: 5px;
            }

            .spreadsheet-controls {
                display: flex;
                align-items: center;
                gap: 15px;
            }

            .zoom-controls {
                display: flex;
                align-items: center;
                gap: 8px;
                background: rgba(255, 255, 255, 0.2);
                padding: 5px 10px;
                border-radius: 20px;
            }

            .zoom-controls button {
                background: rgba(255, 255, 255, 0.3);
                border: none;
                color: white;
                width: 30px;
                height: 30px;
                border-radius: 50%;
                cursor: pointer;
                font-weight: bold;
                transition: background 0.3s;
            }

            .zoom-controls button:hover {
                background: rgba(255, 255, 255, 0.5);
            }

            .zoom-controls span {
                color: white;
                font-weight: 500;
                min-width: 50px;
                text-align: center;
            }

            .close-btn {
                background: rgba(255, 255, 255, 0.2);
                border: none;
                color: white;
                font-size: 24px;
                width: 40px;
                height: 40px;
                border-radius: 50%;
                cursor: pointer;
                transition: background 0.3s;
            }

            .close-btn:hover {
                background: rgba(255, 255, 255, 0.3);
            }

            .sheet-tabs {
                display: flex;
                background: #f5f5f5;
                border-bottom: 1px solid #ddd;
                overflow-x: auto;
                padding: 0 20px;
            }

            .sheet-tab {
                padding: 10px 20px;
                background: #e0e0e0;
                border: none;
                cursor: pointer;
                border-top-left-radius: 8px;
                border-top-right-radius: 8px;
                margin-right: 2px;
                font-weight: 500;
                transition: all 0.3s;
                white-space: nowrap;
            }

            .sheet-tab.active {
                background: white;
                color: #667eea;
                border-bottom: 2px solid #667eea;
            }

            .sheet-tab:hover:not(.active) {
                background: #d0d0d0;
            }

            .spreadsheet-container {
                flex: 1;
                overflow: hidden;
                padding: 20px;
                background: #fafafa;
            }

            .spreadsheet-table-wrapper {
                width: 100%;
                height: 100%;
                overflow: auto;
                border: 2px solid #ddd;
                border-radius: 8px;
                background: white;
                box-shadow: 0 2px 10px rgba(0, 0, 0, 0.1);
            }

            .spreadsheet-table {
                width: 100%;
                border-collapse: separate;
                border-spacing: 0;
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                font-size: 13px;
            }

            .spreadsheet-table th {
                background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
                border: 1px solid #dee2e6;
                padding: 8px 12px;
                text-align: center;
                font-weight: 600;
                color: #495057;
                position: sticky;
                top: 0;
                z-index: 10;
            }

            .spreadsheet-table td {
                border: 1px solid #dee2e6;
                padding: 6px 10px;
                min-width: 80px;
                max-width: 200px;
                overflow: hidden;
                text-overflow: ellipsis;
                white-space: nowrap;
                position: relative;
            }

            .spreadsheet-table td:hover {
                background-color: #f8f9fa;
            }

            .spreadsheet-table td.selected {
                background-color: #e3f2fd;
                border: 2px solid #2196f3;
            }

            .row-header {
                background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
                font-weight: 600;
                text-align: center;
                color: #495057;
                position: sticky;
                left: 0;
                z-index: 5;
            }

            .spreadsheet-footer {
                padding: 15px 20px;
                border-top: 1px solid #e0e0e0;
                background: #f8f9fa;
            }

            .spreadsheet-actions {
                display: flex;
                gap: 10px;
                justify-content: center;
            }

            .action-btn {
                padding: 10px 20px;
                background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                color: white;
                border: none;
                border-radius: 25px;
                cursor: pointer;
                font-weight: 500;
                transition: all 0.3s;
                box-shadow: 0 2px 10px rgba(102, 126, 234, 0.3);
            }

            .action-btn:hover {
                transform: translateY(-2px);
                box-shadow: 0 4px 15px rgba(102, 126, 234, 0.4);
            }

            .loading-overlay {
                position: absolute;
                top: 0;
                left: 0;
                width: 100%;
                height: 100%;
                background: rgba(255, 255, 255, 0.9);
                display: flex;
                justify-content: center;
                align-items: center;
                z-index: 1000;
            }

            .loading-spinner {
                border: 4px solid #f3f3f3;
                border-top: 4px solid #667eea;
                border-radius: 50%;
                width: 50px;
                height: 50px;
                animation: spin 1s linear infinite;
            }

            @keyframes spin {
                0% { transform: rotate(0deg); }
                100% { transform: rotate(360deg); }
            }

            /* Responsive design */
            @media (max-width: 768px) {
                .spreadsheet-modal-content {
                    margin: 1% auto;
                    width: 98%;
                    height: 98%;
                }

                .spreadsheet-header {
                    padding: 10px 15px;
                }

                .spreadsheet-title h3 {
                    font-size: 1.2em;
                }

                .zoom-controls {
                    gap: 5px;
                }

                .spreadsheet-actions {
                    flex-wrap: wrap;
                }

                .action-btn {
                    padding: 8px 16px;
                    font-size: 0.9em;
                }
            }
        `;
        document.head.appendChild(style);
    }

    /**
     * Update sheet tabs
     */
    updateSheetTabs() {
        const tabsContainer = document.getElementById('sheetTabs');
        if (!tabsContainer || !this.sheets) return;

        tabsContainer.innerHTML = '';
        this.sheets.forEach((sheet, index) => {
            const tab = document.createElement('button');
            tab.className = `sheet-tab ${index === this.currentSheet ? 'active' : ''}`;
            tab.textContent = sheet.name;
            tab.onclick = () => this.switchSheet(index);
            tabsContainer.appendChild(tab);
        });
    }

    /**
     * Switch to different sheet
     */
    async switchSheet(sheetIndex) {
        if (sheetIndex === this.currentSheet || !this.sheets[sheetIndex]) return;

        this.currentSheet = sheetIndex;
        this.updateSheetTabs();
        this.loadSheetContent(this.sheets[sheetIndex]);
    }

    /**
     * Load sheet content with formatting
     */
    loadSheetContent(sheetData) {
        const table = document.getElementById('spreadsheetTable');
        const header = document.getElementById('spreadsheetHeader');
        const body = document.getElementById('spreadsheetBody');

        if (!table || !header || !body || !sheetData) return;

        // Clear existing content
        header.innerHTML = '';
        body.innerHTML = '';

        const data = sheetData.data;
        if (!data || data.length === 0) {
            body.innerHTML = '<tr><td colspan="100%" style="text-align: center; padding: 20px;">No data available</td></tr>';
            return;
        }

        // Create header row
        const headerRow = document.createElement('tr');
        
        // Add row number header
        const rowNumHeader = document.createElement('th');
        rowNumHeader.className = 'row-header';
        rowNumHeader.textContent = '';
        headerRow.appendChild(rowNumHeader);

        // Add column headers (A, B, C, ...)
        const maxCols = Math.max(...data.map(row => row.length));
        for (let i = 0; i < maxCols; i++) {
            const th = document.createElement('th');
            th.textContent = this.getColumnName(i);
            headerRow.appendChild(th);
        }
        header.appendChild(headerRow);

        // Create data rows
        data.forEach((rowData, rowIndex) => {
            const tr = document.createElement('tr');
            
            // Add row number
            const rowNumCell = document.createElement('td');
            rowNumCell.className = 'row-header';
            rowNumCell.textContent = rowIndex + 1;
            tr.appendChild(rowNumCell);

            // Add data cells
            for (let colIndex = 0; colIndex < maxCols; colIndex++) {
                const td = document.createElement('td');
                const cellData = rowData[colIndex];
                
                if (cellData) {
                    td.textContent = cellData.value || '';
                    
                    // Apply formatting if available
                    if (cellData.format) {
                        this.applyCellFormatting(td, cellData.format);
                    }
                    
                    // Add click handler for cell selection
                    td.onclick = () => this.selectCell(rowIndex, colIndex, td);
                }
                
                tr.appendChild(td);
            }
            
            body.appendChild(tr);
        });

        // Update zoom
        this.updateZoom();
    }

    /**
     * Apply cell formatting
     */
    applyCellFormatting(cell, format) {
        if (!format) return;

        // Font formatting
        if (format.font) {
            if (format.font.bold) cell.style.fontWeight = 'bold';
            if (format.font.italic) cell.style.fontStyle = 'italic';
            if (format.font.underline) {
                cell.style.textDecoration = cell.style.textDecoration ? 
                    cell.style.textDecoration + ' underline' : 'underline';
            }
            if (format.font.strikethrough) {
                cell.style.textDecoration = cell.style.textDecoration ? 
                    cell.style.textDecoration + ' line-through' : 'line-through';
            }
            if (format.font.color && format.font.color !== '00000000' && format.font.color !== 'FF000000') {
                // Handle both ARGB and RGB color formats
                const colorHex = format.font.color.length === 8 ? 
                    format.font.color.substring(2) : format.font.color;
                cell.style.color = `#${colorHex}`;
            }
            if (format.font.size) cell.style.fontSize = `${format.font.size}px`;
            if (format.font.name) cell.style.fontFamily = `"${format.font.name}", sans-serif`;
        }

        // Background color
        if (format.fill && format.fill.color && format.fill.color !== '00000000' && format.fill.color !== 'FF000000') {
            // Handle both ARGB and RGB color formats
            const colorHex = format.fill.color.length === 8 ? 
                format.fill.color.substring(2) : format.fill.color;
            cell.style.backgroundColor = `#${colorHex}`;
        }

        // Alignment
        if (format.alignment) {
            if (format.alignment.horizontal) {
                const alignMap = {
                    'left': 'left',
                    'center': 'center', 
                    'right': 'right',
                    'justify': 'justify',
                    'general': 'left'
                };
                cell.style.textAlign = alignMap[format.alignment.horizontal] || format.alignment.horizontal;
            }
            if (format.alignment.vertical) {
                const vAlignMap = {
                    'top': 'top',
                    'middle': 'middle',
                    'bottom': 'bottom',
                    'center': 'middle'
                };
                cell.style.verticalAlign = vAlignMap[format.alignment.vertical] || format.alignment.vertical;
            }
            if (format.alignment.wrap_text) {
                cell.style.whiteSpace = 'pre-wrap';
                cell.style.wordWrap = 'break-word';
                cell.style.overflowWrap = 'break-word';
            }
            if (format.alignment.indent && format.alignment.indent > 0) {
                cell.style.paddingLeft = `${format.alignment.indent * 8}px`;
            }
        }

        // Borders with enhanced styling
        if (format.border) {
            const getBorderStyle = (border) => {
                if (!border) return 'none';
                const style = border.style || 'thin';
                const color = border.color ? `#${border.color.length === 8 ? border.color.substring(2) : border.color}` : '#000000';
                const width = style === 'thick' ? '2px' : style === 'medium' ? '1.5px' : '1px';
                const borderType = style === 'dashed' ? 'dashed' : style === 'dotted' ? 'dotted' : 'solid';
                return `${width} ${borderType} ${color}`;
            };
            
            if (format.border.top) cell.style.borderTop = getBorderStyle(format.border.top);
            if (format.border.bottom) cell.style.borderBottom = getBorderStyle(format.border.bottom);
            if (format.border.left) cell.style.borderLeft = getBorderStyle(format.border.left);
            if (format.border.right) cell.style.borderRight = getBorderStyle(format.border.right);
        }

        // Number format preservation
        if (format.number_format) {
            cell.setAttribute('data-number-format', format.number_format);
        }

        // Cell protection
        if (format.protection) {
            if (format.protection.locked) {
                cell.style.pointerEvents = 'none';
                cell.style.opacity = '0.7';
            }
        }
    }

    /**
     * Get column name (A, B, C, ...)
     */
    getColumnName(index) {
        let result = '';
        while (index >= 0) {
            result = String.fromCharCode(65 + (index % 26)) + result;
            index = Math.floor(index / 26) - 1;
        }
        return result;
    }

    /**
     * Select cell
     */
    selectCell(row, col, cellElement) {
        // Clear previous selection
        document.querySelectorAll('.spreadsheet-table td.selected').forEach(cell => {
            cell.classList.remove('selected');
        });

        // Select new cell
        cellElement.classList.add('selected');
        this.selectedCells = [{ row, col, element: cellElement }];
    }

    /**
     * Update file info
     */
    updateFileInfo(fileInfo) {
        const infoElement = document.getElementById('spreadsheetInfo');
        if (!infoElement || !fileInfo) return;

        const sizeText = this.formatFileSize(fileInfo.size);
        infoElement.innerHTML = `
            ${fileInfo.sheets_count} sheet(s) • ${sizeText} • 
            ${fileInfo.preserve_formatting ? 'Formatted' : 'Plain'}
        `;
    }

    /**
     * Format file size
     */
    formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }

    /**
     * Zoom functions
     */
    zoomIn() {
        this.zoomLevel = Math.min(this.zoomLevel + 10, 200);
        this.updateZoom();
    }

    zoomOut() {
        this.zoomLevel = Math.max(this.zoomLevel - 10, 50);
        this.updateZoom();
    }

    resetZoom() {
        this.zoomLevel = 100;
        this.updateZoom();
    }

    updateZoom() {
        const table = document.getElementById('spreadsheetTable');
        const zoomDisplay = document.getElementById('zoomLevel');
        
        if (table) {
            table.style.transform = `scale(${this.zoomLevel / 100})`;
            table.style.transformOrigin = 'top left';
        }
        
        if (zoomDisplay) {
            zoomDisplay.textContent = `${this.zoomLevel}%`;
        }
    }

    /**
     * Action functions
     */
    async downloadAsExcel() {
        if (!this.currentFile) {
            this.showError('No file available for download');
            return;
        }
        
        this.showSuccess('Download feature will be implemented');
    }

    showConvertToPdfOptions() {
        if (!this.currentFile) {
            this.showError('No file available for conversion');
            return;
        }
        
        // Create options modal
        const optionsModal = document.createElement('div');
        optionsModal.className = 'pdf-options-modal';
        optionsModal.innerHTML = `
            <div class="pdf-options-content">
                <h3>Convert to PDF Options</h3>
                <div class="pdf-options-form">
                    <div class="option-group">
                        <label>
                            <input type="radio" name="pdfType" value="standard" checked>
                            Standard PDF
                        </label>
                        <p class="option-description">Basic PDF conversion with standard formatting</p>
                    </div>
                    <div class="option-group">
                        <label>
                            <input type="radio" name="pdfType" value="searchable">
                            Searchable PDF (OCR)
                        </label>
                        <p class="option-description">PDF with OCR text recognition for better searchability</p>
                    </div>
                    <div class="option-group">
                        <label>
                            <input type="checkbox" id="preserveFormatting" checked>
                            Preserve Excel formatting
                        </label>
                    </div>
                    <div class="option-group">
                        <label for="pageOrientation">Page Orientation:</label>
                        <select id="pageOrientation">
                            <option value="portrait">Portrait</option>
                            <option value="landscape" selected>Landscape</option>
                        </select>
                    </div>
                </div>
                <div class="pdf-options-actions">
                    <button onclick="this.remove()" class="cancel-btn">Cancel</button>
                    <button onclick="spreadsheetPopup.convertToPDF(this)" class="convert-btn">Convert to PDF</button>
                </div>
            </div>
        `;
        
        document.body.appendChild(optionsModal);
        
        // Add styles for options modal
        if (!document.getElementById('pdfOptionsStyles')) {
            const style = document.createElement('style');
            style.id = 'pdfOptionsStyles';
            style.textContent = `
                .pdf-options-modal {
                    position: fixed;
                    top: 0;
                    left: 0;
                    width: 100%;
                    height: 100%;
                    background: rgba(0, 0, 0, 0.5);
                    display: flex;
                    justify-content: center;
                    align-items: center;
                    z-index: 10001;
                }
                .pdf-options-content {
                    background: white;
                    padding: 30px;
                    border-radius: 10px;
                    max-width: 500px;
                    width: 90%;
                }
                .pdf-options-form {
                    margin: 20px 0;
                }
                .option-group {
                    margin-bottom: 15px;
                }
                .option-description {
                    font-size: 0.9em;
                    color: #666;
                    margin: 5px 0 0 25px;
                }
                .pdf-options-actions {
                    display: flex;
                    gap: 10px;
                    justify-content: flex-end;
                }
                .cancel-btn, .convert-btn {
                    padding: 10px 20px;
                    border: none;
                    border-radius: 5px;
                    cursor: pointer;
                }
                .cancel-btn {
                    background: #ccc;
                }
                .convert-btn {
                    background: #667eea;
                    color: white;
                }
            `;
            document.head.appendChild(style);
        }
    }

    async convertToPDF(optionsModal) {
        const pdfType = optionsModal.querySelector('input[name="pdfType"]:checked').value;
        const preserveFormatting = optionsModal.querySelector('#preserveFormatting').checked;
        const pageOrientation = optionsModal.querySelector('#pageOrientation').value;
        
        // Remove options modal
        optionsModal.remove();
        
        try {
            this.showLoadingState('Converting to PDF...');
            
            // Call PDF conversion API
            const response = await fetch('/api/excel-pywin32/convert-to-pdf', {
                 method: 'POST',
                 headers: {
                     'Content-Type': 'application/json'
                 },
                 body: JSON.stringify({ 
                     file_id: this.currentFile.file_id,
                     preserve_formatting: preserveFormatting,
                     preserve_charts: true,
                     preserve_images: true,
                     preserve_formulas: true,
                     preserve_colors: true,
                     preserve_fonts: true,
                     preserve_borders: true,
                     preserve_alignment: true,
                     fit_to_page: true,
                     high_quality: true,
                     include_all_sheets: true,
                     page_orientation: pageOrientation === 'portrait' ? 'portrait' : 'landscape',
                     paper_size: 'A4',
                     margins: 'normal',
                     conversion_method: 'auto'
                 })
             });
            
            const result = await response.json();
            
            if (!response.ok) {
                throw new Error(result.detail || 'Conversion failed');
            }
            
            this.hideLoadingState();
            
            if (result.status === 'success') {
                // Show conversion details
                const conversionInfo = result.conversion_result;
                const message = `PDF conversion completed successfully!\n` +
                    `Method: ${conversionInfo.method}\n` +
                    `Quality Score: ${conversionInfo.quality_score}/10\n` +
                    `File Size: ${(result.file_size / 1024 / 1024).toFixed(2)} MB`;
                
                this.showSuccess(message);
                
                // Download the converted PDF
                const downloadLink = document.createElement('a');
                downloadLink.href = result.download_url;
                downloadLink.download = result.pdf_filename;
                downloadLink.click();
            } else {
                throw new Error(result.detail || 'PDF conversion failed');
            }
            
        } catch (error) {
            this.hideLoadingState();
            this.showError(`PDF conversion failed: ${error.message}`);
        }
    }

    async printSpreadsheet() {
        if (!this.currentFile) {
            this.showError('No file available for printing');
            return;
        }
        
        // Open print dialog for the current sheet
        const printWindow = window.open('', '_blank');
        const table = document.getElementById('spreadsheetTable').cloneNode(true);
        
        printWindow.document.write(`
            <html>
                <head>
                    <title>Print - ${this.currentFile.filename}</title>
                    <style>
                        body { font-family: Arial, sans-serif; margin: 20px; }
                        table { border-collapse: collapse; width: 100%; }
                        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
                        th { background-color: #f2f2f2; }
                        @media print { body { margin: 0; } }
                    </style>
                </head>
                <body>
                    <h2>${this.currentFile.filename} - ${this.sheets[this.currentSheet].name}</h2>
                    ${table.outerHTML}
                </body>
            </html>
        `);
        
        printWindow.document.close();
        printWindow.print();
    }

    /**
     * Handle escape key
     */
    handleEscapeKey(event) {
        if (event.key === 'Escape' && this.isModalOpen) {
            this.closeModal();
        }
    }

    /**
     * Close modal
     */
    closeModal() {
        const modal = document.getElementById('spreadsheetModal');
        if (modal) {
            modal.style.display = 'none';
            this.isModalOpen = false;
        }
        
        // Remove escape key listener
        document.removeEventListener('keydown', this.handleEscapeKey.bind(this));
    }

    /**
     * Utility functions
     */
    showLoadingState(message = 'Loading...') {
        const modal = document.getElementById('spreadsheetModal');
        if (!modal) return;

        let overlay = modal.querySelector('.loading-overlay');
        if (!overlay) {
            overlay = document.createElement('div');
            overlay.className = 'loading-overlay';
            modal.appendChild(overlay);
        }

        overlay.innerHTML = `
            <div style="text-align: center;">
                <div class="loading-spinner"></div>
                <p style="margin-top: 15px; font-weight: 500;">${message}</p>
            </div>
        `;
        overlay.style.display = 'flex';
    }

    hideLoadingState() {
        const overlay = document.querySelector('.loading-overlay');
        if (overlay) {
            overlay.style.display = 'none';
        }
    }

    showError(message) {
        alert(`Error: ${message}`);
    }

    showSuccess(message) {
        alert(`Success: ${message}`);
    }
}

// Global instance
const spreadsheetPopup = new SpreadsheetPopupManager();

// Export for module usage
if (typeof module !== 'undefined' && module.exports) {
    module.exports = SpreadsheetPopupManager;
}
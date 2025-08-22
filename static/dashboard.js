// Dashboard JavaScript for Epson L120 Print Manager

class PrintDashboard {
    constructor() {
        this.currentFile = null;
        this.currentFileId = null;
        this.selectedPrinter = null;
        this.jobQueue = [];
        this.jobHistory = [];
        this.jobIdCounter = 1;
        
        this.init();
    }

    init() {
        this.setupEventListeners();
        this.loadPrinters();
        this.startJobMonitoring();
        this.loadJobHistory();
    }

    setupEventListeners() {
        // File upload events
        const uploadArea = document.getElementById('upload-area');
        const fileInput = document.getElementById('file-input');
        
        uploadArea.addEventListener('click', () => fileInput.click());
        uploadArea.addEventListener('dragover', this.handleDragOver.bind(this));
        uploadArea.addEventListener('dragleave', this.handleDragLeave.bind(this));
        uploadArea.addEventListener('drop', this.handleDrop.bind(this));
        fileInput.addEventListener('change', this.handleFileSelect.bind(this));

        // Printer selection
        document.getElementById('printer-select').addEventListener('change', this.handlePrinterSelect.bind(this));

        // Print actions
        document.getElementById('print-btn').addEventListener('click', this.printDocument.bind(this));
        
        // Update print button state initially
        this.updatePrintButtonState();
        document.getElementById('clear-preview').addEventListener('click', this.clearPreview.bind(this));
        document.getElementById('test-print-btn').addEventListener('click', this.printTestPage.bind(this));

        // Job management
        document.getElementById('clear-queue').addEventListener('click', this.clearQueue.bind(this));
        document.getElementById('clear-history').addEventListener('click', this.clearHistory.bind(this));
        document.getElementById('refresh-all').addEventListener('click', this.refreshAll.bind(this));
    }

    // File Upload Handlers
    handleDragOver(e) {
        e.preventDefault();
        e.currentTarget.classList.add('dragover');
    }

    handleDragLeave(e) {
        e.preventDefault();
        e.currentTarget.classList.remove('dragover');
    }

    handleDrop(e) {
        e.preventDefault();
        e.currentTarget.classList.remove('dragover');
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            this.processFile(files[0]);
        }
    }
    
    updatePrintButtonState() {
        const printBtn = document.getElementById('print-btn');
        if (printBtn) {
            const hasFile = !!this.currentFile;
            const hasPrinter = !!this.selectedPrinter;
            const hasFileId = !!this.currentFileId;
            const canPrint = hasFile && hasPrinter && hasFileId;
            
            printBtn.disabled = !canPrint;
            
            // Enhanced logging for debugging
            console.log('Print button state updated:', {
                currentFile: hasFile,
                currentFileName: this.currentFile ? this.currentFile.name : 'none',
                selectedPrinter: hasPrinter,
                selectedPrinterValue: this.selectedPrinter || 'none',
                currentFileId: hasFileId,
                currentFileIdValue: this.currentFileId || 'none',
                canPrint: canPrint,
                buttonDisabled: printBtn.disabled,
                buttonElement: !!printBtn
            });
            
            // Visual feedback for button state
            if (canPrint) {
                printBtn.classList.remove('btn-secondary');
                printBtn.classList.add('btn-primary');
                printBtn.title = 'Ready to print';
            } else {
                printBtn.classList.remove('btn-primary');
                printBtn.classList.add('btn-secondary');
                
                let missingItems = [];
                if (!hasFile) missingItems.push('file');
                if (!hasPrinter) missingItems.push('printer');
                if (!hasFileId) missingItems.push('file upload');
                
                printBtn.title = `Missing: ${missingItems.join(', ')}`;
            }
        } else {
            console.error('Print button element not found!');
        }
    }

    handleFileSelect(e) {
        const files = e.target.files;
        if (files.length > 0) {
            this.processFile(files[0]);
        }
    }

    async processFile(file) {
        this.currentFile = file;
        this.showMessage(`Processing file: ${file.name}`, 'info');
        
        try {
            // Upload file to server first
            const uploadResult = await this.uploadFile(file);
            this.currentFileId = uploadResult.file_id;
            
            // Generate preview using uploaded file
            await this.generatePreview(file, uploadResult);
            this.showPreviewSection();
            this.showPrintSettings();
            this.showMessage(`File uploaded successfully: ${file.name}`, 'success');
            
            // Update print button state after file is processed
            this.updatePrintButtonState();
        } catch (error) {
            this.showMessage(`Error processing file: ${error.message}`, 'error');
            console.error('File processing error:', error);
        }
    }

    async uploadFile(file) {
        const formData = new FormData();
        formData.append('file', file);
        
        const response = await fetch('/api/files/upload', {
            method: 'POST',
            body: formData
        });
        
        if (!response.ok) {
            const error = await response.json();
            throw new Error(error.detail || 'Upload failed');
        }
        
        return await response.json();
    }

    async generatePreview(file, uploadResult) {
        const previewContent = document.getElementById('preview-content');
        const fileType = file.type;
        const fileName = file.name.toLowerCase();
        const previewUrl = `/api/files/${uploadResult.file_id}/preview`;

        previewContent.innerHTML = '';

        if (fileType.startsWith('image/')) {
            const img = document.createElement('img');
            img.src = previewUrl;
            img.style.maxWidth = '100%';
            img.style.maxHeight = '300px';
            img.style.objectFit = 'contain';
            img.onerror = () => {
                this.showFileInfo(file, previewContent);
            };
            previewContent.appendChild(img);
        } else if (fileType === 'application/pdf') {
            const embed = document.createElement('embed');
            embed.src = previewUrl;
            embed.type = 'application/pdf';
            embed.style.width = '100%';
            embed.style.height = '300px';
            embed.onerror = () => {
                this.showFileInfo(file, previewContent);
            };
            previewContent.appendChild(embed);
        } else if (fileType === 'text/plain' || fileName.endsWith('.txt')) {
            try {
                const response = await fetch(previewUrl);
                if (response.ok) {
                    const text = await response.text();
                    const pre = document.createElement('pre');
                    pre.textContent = text.substring(0, 1000) + (text.length > 1000 ? '...' : '');
                    pre.style.whiteSpace = 'pre-wrap';
                    pre.style.fontSize = '0.9rem';
                    pre.style.lineHeight = '1.4';
                    previewContent.appendChild(pre);
                } else {
                    this.showFileInfo(file, previewContent);
                }
            } catch (error) {
                this.showFileInfo(file, previewContent);
            }
        } else {
            this.showFileInfo(file, previewContent);
        }
    }

    showFileInfo(file, container) {
        container.innerHTML = `
            <div style="text-align: center; padding: 2rem;">
                <i class="fas fa-file-alt" style="font-size: 3rem; color: #cbd5e0; margin-bottom: 1rem;"></i>
                <p><strong>${file.name}</strong></p>
                <p>Size: ${this.formatFileSize(file.size)}</p>
                <p>Type: ${file.type || 'Unknown'}</p>
                <p style="color: #64748b; font-size: 0.9rem;">Preview not available for this file type</p>
            </div>
        `;
    }

    showPreviewSection() {
        document.getElementById('preview-section').style.display = 'block';
    }

    showPrintSettings() {
        document.getElementById('print-settings').style.display = 'block';
    }

    clearPreview() {
        this.currentFile = null;
        this.currentFileId = null;
        document.getElementById('preview-section').style.display = 'none';
        document.getElementById('print-settings').style.display = 'none';
        document.getElementById('file-input').value = '';
        this.showMessage('Preview cleared', 'info');
    }

    // Printer Management
    async loadPrinters() {
        try {
            const response = await fetch('/api/printers');
            const data = await response.json();
            
            const select = document.getElementById('printer-select');
            select.innerHTML = '<option value="">Select a printer...</option>';
            
            if (data.printers && data.printers.length > 0) {
                let epsonL120Found = false;
                
                data.printers.forEach(printer => {
                    const option = document.createElement('option');
                    option.value = printer.id || printer.name;
                    option.textContent = printer.name;
                    select.appendChild(option);
                    
                    // Auto-select EPSON L120 if found
                    if (printer.name && printer.name.toLowerCase().includes('epson l120')) {
                        epsonL120Found = true;
                        select.value = printer.id || printer.name;
                        this.selectedPrinter = printer.id || printer.name;
                    }
                });
                
                // If EPSON L120 was found and selected, update status and button state
                if (epsonL120Found) {
                    console.log('EPSON L120 auto-selected:', this.selectedPrinter);
                    this.updatePrinterStatus();
                    this.updatePrintButtonState();
                }
            } else {
                select.innerHTML = '<option value="">No printers found</option>';
            }
        } catch (error) {
            console.error('Error loading printers:', error);
            this.showMessage('Failed to load printers', 'error');
        }
    }

    handlePrinterSelect(e) {
        this.selectedPrinter = e.target.value;
        console.log('Printer selected:', this.selectedPrinter);
        
        this.updatePrinterStatus();
        
        const testPrintBtn = document.getElementById('test-print-btn');
        if (testPrintBtn) {
            testPrintBtn.disabled = !this.selectedPrinter;
        }
        
        // Update print button state
        this.updatePrintButtonState();
    }

    async updatePrinterStatus() {
        const indicator = document.getElementById('printer-status-indicator');
        const statusDot = indicator.querySelector('.status-dot');
        const statusText = indicator.querySelector('.status-text');

        if (!this.selectedPrinter) {
            statusDot.className = 'status-dot';
            statusText.textContent = 'No printer selected';
            return;
        }

        try {
            const response = await fetch(`/api/printers/${this.selectedPrinter}/status`);
            const data = await response.json();
            
            statusDot.className = `status-dot ${data.status || 'offline'}`;
            statusText.textContent = data.message || 'Unknown status';
        } catch (error) {
            statusDot.className = 'status-dot offline';
            statusText.textContent = 'Status unavailable';
        }
    }

    // Print Operations
    async printDocument() {
        if (!this.currentFile || !this.selectedPrinter || !this.currentFileId) {
            this.showMessage('Please select a file and printer', 'error');
            return;
        }

        const settings = this.getPrintSettings();
        const docProcessingSettings = this.getDocumentProcessingSettings();
        
        // Validate page range if specified
        if (settings.page_range && settings.page_range !== 'odd' && settings.page_range !== 'even') {
            const pageRangePattern = /^\d+(-\d+)?(,\s*\d+(-\d+)?)*$/;
            if (!pageRangePattern.test(settings.page_range)) {
                this.showMessage('Invalid page range format. Use format like: 1-5, 8, 10-12', 'error');
                return;
            }
        }
        
        const jobId = this.jobIdCounter++;
        
        const job = {
            id: jobId,
            name: this.currentFile.name,
            fileName: this.currentFile.name, // Add fileName for job monitoring
            printer: this.selectedPrinter,
            status: 'pending',
            progress: 0,
            timestamp: new Date(),
            settings: settings,
            documentProcessing: docProcessingSettings
        };

        this.addJobToQueue(job);
        this.showLoadingOverlay('Sending print job...');

        try {
            // Use the print/with-processing endpoint with document processing
            const requestData = {
                printer_id: this.selectedPrinter,
                file_id: this.currentFileId,
                print_settings: {
                    copies: settings.copies,
                    color_mode: settings.color_mode,
                    paper_size: settings.paper_size,
                    orientation: settings.orientation,
                    quality: settings.quality,
                    page_range: settings.page_range
                },
                document_settings: {
                    // Include document manipulation settings
                    format_conversion: docProcessingSettings.format_conversion,
                    color_processing: docProcessingSettings.color_processing,
                    brightness: docProcessingSettings.brightness,
                    contrast: docProcessingSettings.contrast,
                    pdf_split_type: docProcessingSettings.pdf_split_type,
                    pdf_split_range: docProcessingSettings.pdf_split_range,
                    page_range_type: settings.page_range_type,
                    page_range: settings.page_range,
                    orientation: settings.orientation,
                    paper_size: settings.paper_size
                }
            };

            const response = await fetch('/api/print/with-processing', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(requestData)
            });

            const result = await response.json();
            
            if (response.ok) {
                job.status = 'printing';
                this.updateJobInQueue(job);
                this.showMessage(`Print job with document processing started: ${this.currentFile.name}`, 'success');
                this.monitorJobProgress(job);
            } else {
                job.status = 'error';
                this.updateJobInQueue(job);
                this.showMessage(`Print failed: ${result.detail || 'Unknown error'}`, 'error');
            }
        } catch (error) {
            job.status = 'error';
            this.updateJobInQueue(job);
            this.showMessage(`Print error: ${error.message}`, 'error');
        } finally {
            this.hideLoadingOverlay();
        }
    }

    async printTestPage() {
        if (!this.selectedPrinter) {
            this.showMessage('Please select a printer', 'error');
            return;
        }

        const jobId = this.jobIdCounter++;
        const job = {
            id: jobId,
            name: 'Test Page',
            printer: this.selectedPrinter,
            status: 'pending',
            progress: 0,
            timestamp: new Date()
        };

        this.addJobToQueue(job);
        this.showLoadingOverlay('Printing test page...');

        try {
            const response = await fetch('/api/print/test', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({ printer_id: this.selectedPrinter })
            });

            const result = await response.json();
            
            if (response.ok) {
                job.status = 'printing';
                this.updateJobInQueue(job);
                this.showMessage('Test page sent to printer', 'success');
                this.simulateJobProgress(job);
            } else {
                job.status = 'error';
                this.updateJobInQueue(job);
                this.showMessage(`Test print failed: ${result.detail || 'Unknown error'}`, 'error');
            }
        } catch (error) {
            job.status = 'error';
            this.updateJobInQueue(job);
            this.showMessage(`Test print error: ${error.message}`, 'error');
        } finally {
            this.hideLoadingOverlay();
        }
    }

    getPrintSettings() {
        const pageRangeType = document.getElementById('page-range-type').value;
        const pageRangeInput = document.getElementById('page-range').value;
        const fitToPage = document.getElementById('fit-to-page').value;
        const pdfSplit = document.getElementById('pdf-split').value;
        const splitRange = document.getElementById('split-range').value;
        
        let pageRange = null;
        if (pageRangeType === 'range' && pageRangeInput.trim()) {
            pageRange = pageRangeInput.trim();
        } else if (pageRangeType === 'odd') {
            pageRange = 'odd';
        } else if (pageRangeType === 'even') {
            pageRange = 'even';
        }
        
        // Handle split PDF settings
        let splitPdf = false;
        let splitPageRange = null;
        let splitOutputPrefix = 'split_page';
        
        if (pdfSplit === 'pages' && splitRange.trim()) {
            splitPdf = true;
            splitPageRange = splitRange.trim();
        } else if (pdfSplit === 'single') {
            splitPdf = true;
            splitPageRange = 'all';
        }
        
        return {
            copies: parseInt(document.getElementById('copies').value) || 1,
            color_mode: document.getElementById('color-mode').value || 'auto',
            paper_size: document.getElementById('paper-size').value || 'A4',
            quality: document.getElementById('quality').value || 'normal',
            orientation: document.getElementById('orientation').value || 'portrait',
            page_range: pageRange,
            fit_to_page: fitToPage,
            split_pdf: splitPdf,
            split_page_range: splitPageRange,
            split_output_prefix: splitOutputPrefix
        };
    }

    getDocumentProcessingSettings() {
        const formatConversion = document.getElementById('format-conversion').value;
        const colorProcessing = document.getElementById('color-processing').value;
        const brightness = parseFloat(document.getElementById('brightness').value) || 1.0;
        const contrast = parseFloat(document.getElementById('contrast').value) || 1.0;
        const pdfSplitType = document.getElementById('pdf-split-type').value;
        const pdfSplitRange = document.getElementById('pdf-split-range').value;
        
        return {
            format_conversion: formatConversion !== 'none' ? formatConversion : null,
            color_processing: colorProcessing !== 'none' ? colorProcessing : null,
            brightness: brightness !== 1.0 ? brightness : null,
            contrast: contrast !== 1.0 ? contrast : null,
            pdf_split_type: pdfSplitType !== 'none' ? pdfSplitType : null,
            pdf_split_range: pdfSplitType === 'range' && pdfSplitRange.trim() ? pdfSplitRange.trim() : null
        };
    }

    // Job Queue Management
    addJobToQueue(job) {
        this.jobQueue.push(job);
        this.updateJobQueueDisplay();
    }

    updateJobInQueue(job) {
        const index = this.jobQueue.findIndex(j => j.id === job.id);
        if (index !== -1) {
            this.jobQueue[index] = job;
            this.updateJobQueueDisplay();
        }
    }

    removeJobFromQueue(jobId) {
        const jobIndex = this.jobQueue.findIndex(j => j.id === jobId);
        if (jobIndex !== -1) {
            const job = this.jobQueue[jobIndex];
            this.jobQueue.splice(jobIndex, 1);
            this.addJobToHistory(job);
            this.updateJobQueueDisplay();
            this.updateJobHistoryDisplay();
        }
    }

    updateJobQueueDisplay() {
        const jobList = document.getElementById('job-list');
        const jobCount = document.getElementById('job-count');
        
        jobCount.textContent = this.jobQueue.length;

        if (this.jobQueue.length === 0) {
            jobList.innerHTML = `
                <div class="no-jobs">
                    <i class="fas fa-inbox"></i>
                    <p>No print jobs</p>
                </div>
            `;
            return;
        }

        jobList.innerHTML = this.jobQueue.map(job => `
            <div class="job-item" data-job-id="${job.id}">
                <div class="job-info">
                    <div class="job-name">${job.name}</div>
                    <div class="job-details">${job.printer} • ${this.formatTimestamp(job.timestamp)}</div>
                </div>
                <div class="job-progress">
                    <div class="progress-bar">
                        <div class="progress-fill" style="width: ${job.progress}%"></div>
                    </div>
                    <div class="progress-text">${job.status} (${job.progress}%)</div>
                </div>
                <div class="job-actions">
                    <button class="btn btn-outline btn-sm" onclick="dashboard.cancelJob(${job.id})">
                        <i class="fas fa-times"></i>
                    </button>
                </div>
            </div>
        `).join('');
    }

    async monitorJobProgress(job) {
        console.log('Starting real-time job monitoring for:', job.id);
        
        const interval = setInterval(async () => {
            if (job.status === 'error' || job.status === 'cancelled') {
                clearInterval(interval);
                this.removeJobFromQueue(job.id);
                return;
            }

            try {
                // Get detailed printer status for real-time monitoring
                const response = await fetch(`/api/printers/${this.selectedPrinter}/status/detailed`);
                if (response.ok) {
                    const printerStatus = await response.json();
                    
                    // Check if our job is in the printer queue
                    const currentJob = printerStatus.jobs?.find(pJob => 
                        pJob.document && pJob.document.includes(job.fileName)
                    );
                    
                    if (currentJob) {
                        // Calculate progress based on pages printed
                        if (currentJob.total_pages > 0) {
                            job.progress = Math.min(95, (currentJob.pages_printed / currentJob.total_pages) * 100);
                        } else {
                            // Fallback to incremental progress
                            job.progress = Math.min(90, job.progress + 10);
                        }
                        
                        console.log(`Job ${job.id} progress: ${job.progress}% (${currentJob.pages_printed}/${currentJob.total_pages} pages)`);
                    } else {
                        // Job not found in printer queue - might be completed or processing
                        if (job.progress < 95) {
                            job.progress = Math.min(95, job.progress + 15);
                        } else {
                            // Job likely completed
                            job.progress = 100;
                            job.status = 'completed';
                            console.log(`Job ${job.id} completed - no longer in printer queue`);
                            clearInterval(interval);
                            setTimeout(() => {
                                this.removeJobFromQueue(job.id);
                            }, 2000);
                        }
                    }
                    
                    // Check printer status for errors
                    if (printerStatus.status === 'error' || printerStatus.status === 'offline') {
                        job.status = 'error';
                        job.error = `Printer ${printerStatus.status}: ${printerStatus.message || 'Unknown error'}`;
                        clearInterval(interval);
                        this.removeJobFromQueue(job.id);
                        return;
                    }
                } else {
                    console.warn('Failed to get printer status, using fallback progress');
                    // Fallback to simulated progress if API fails
                    job.progress += Math.random() * 10 + 5;
                }
            } catch (error) {
                console.error('Error monitoring job progress:', error);
                // Fallback to simulated progress on error
                job.progress += Math.random() * 10 + 5;
            }
            
            // Ensure progress doesn't exceed 100%
            if (job.progress >= 100) {
                job.progress = 100;
                job.status = 'completed';
                clearInterval(interval);
                setTimeout(() => {
                    this.removeJobFromQueue(job.id);
                }, 2000);
            }

            this.updateJobInQueue(job);
        }, 2000); // Check every 2 seconds for more accurate monitoring
    }

    cancelJob(jobId) {
        const job = this.jobQueue.find(j => j.id === jobId);
        if (job) {
            job.status = 'cancelled';
            job.progress = 0;
            this.updateJobInQueue(job);
            setTimeout(() => {
                this.removeJobFromQueue(jobId);
            }, 1000);
        }
    }

    clearQueue() {
        this.jobQueue.forEach(job => {
            job.status = 'cancelled';
            this.addJobToHistory(job);
        });
        this.jobQueue = [];
        this.updateJobQueueDisplay();
        this.updateJobHistoryDisplay();
        this.showMessage('Print queue cleared', 'info');
    }

    // Job History Management
    addJobToHistory(job) {
        this.jobHistory.unshift({
            ...job,
            completedAt: new Date()
        });
        
        // Keep only last 20 jobs
        if (this.jobHistory.length > 20) {
            this.jobHistory = this.jobHistory.slice(0, 20);
        }
        
        this.saveJobHistory();
    }

    updateJobHistoryDisplay() {
        const historyList = document.getElementById('job-history-list');
        
        if (this.jobHistory.length === 0) {
            historyList.innerHTML = `
                <div class="no-history">
                    <p>No recent jobs</p>
                </div>
            `;
            return;
        }

        historyList.innerHTML = this.jobHistory.map(job => `
            <div class="history-item">
                <div class="history-status ${job.status}"></div>
                <div class="history-info">
                    <div class="history-name">${job.name}</div>
                    <div class="history-time">${this.formatTimestamp(job.completedAt || job.timestamp)}</div>
                </div>
            </div>
        `).join('');
    }

    clearHistory() {
        this.jobHistory = [];
        this.updateJobHistoryDisplay();
        this.saveJobHistory();
        this.showMessage('Job history cleared', 'info');
    }

    saveJobHistory() {
        try {
            localStorage.setItem('printJobHistory', JSON.stringify(this.jobHistory));
        } catch (error) {
            console.error('Failed to save job history:', error);
        }
    }

    loadJobHistory() {
        try {
            const saved = localStorage.getItem('printJobHistory');
            if (saved) {
                this.jobHistory = JSON.parse(saved);
                this.updateJobHistoryDisplay();
            }
        } catch (error) {
            console.error('Failed to load job history:', error);
        }
    }

    // Monitoring and Refresh
    startJobMonitoring() {
        // Update printer status every 30 seconds
        setInterval(() => {
            if (this.selectedPrinter) {
                this.updatePrinterStatus();
            }
        }, 30000);
    }

    async refreshAll() {
        this.showLoadingOverlay('Refreshing...');
        try {
            await this.loadPrinters();
            if (this.selectedPrinter) {
                await this.updatePrinterStatus();
            }
            this.showMessage('Dashboard refreshed', 'success');
        } catch (error) {
            this.showMessage('Refresh failed', 'error');
        } finally {
            this.hideLoadingOverlay();
        }
    }

    // Utility Functions
    formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }

    formatTimestamp(timestamp) {
        const date = new Date(timestamp);
        return date.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    }

    showMessage(message, type = 'info') {
        const messageEl = document.getElementById('message');
        messageEl.textContent = message;
        messageEl.className = `message ${type}`;
        messageEl.style.display = 'block';
        
        setTimeout(() => {
            messageEl.style.display = 'none';
        }, 5000);
    }

    showLoadingOverlay(message = 'Loading...') {
        const overlay = document.getElementById('loading-overlay');
        const text = overlay.querySelector('p');
        text.textContent = message;
        overlay.style.display = 'flex';
    }

    hideLoadingOverlay() {
        document.getElementById('loading-overlay').style.display = 'none';
    }

    // Document Manipulation Functions
    async applyDocumentProcessing() {
        if (!this.currentFileId) {
            this.showMessage('No file selected for processing', 'error');
            return;
        }

        const settings = this.getDocumentProcessingSettings();
        
        try {
            this.showLoadingOverlay('Processing document...');
            
            const response = await fetch('/api/files/manipulate', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    file_id: this.currentFileId,
                    settings: settings
                })
            });

            if (!response.ok) {
                const error = await response.json();
                throw new Error(error.detail || 'Processing failed');
            }

            const result = await response.json();
            this.currentFileId = result.processed_file_id;
            
            this.showMessage('Document processed successfully', 'success');
            
            // Update preview with processed document
            await this.updatePreviewWithProcessed(result);
            
        } catch (error) {
            this.showMessage(`Processing error: ${error.message}`, 'error');
            console.error('Document processing error:', error);
        } finally {
            this.hideLoadingOverlay();
        }
    }

    async previewProcessedDocument() {
        if (!this.currentFileId) {
            this.showMessage('No file selected for preview', 'error');
            return;
        }

        const settings = this.getDocumentProcessingSettings();
        
        try {
            this.showLoadingOverlay('Generating preview...');
            
            const response = await fetch('/api/files/preview-processed', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    file_id: this.currentFileId,
                    settings: settings
                })
            });

            if (!response.ok) {
                const error = await response.json();
                throw new Error(error.detail || 'Preview generation failed');
            }

            const result = await response.json();
            
            // Update preview with processed version
            await this.updatePreviewWithProcessed(result);
            
            this.showMessage('Preview updated with processed document', 'success');
            
        } catch (error) {
            this.showMessage(`Preview error: ${error.message}`, 'error');
            console.error('Preview generation error:', error);
        } finally {
            this.hideLoadingOverlay();
        }
    }

    getDocumentProcessingSettings() {
        return {
            target_format: document.getElementById('target-format').value,
            color_processing: document.getElementById('color-processing').value,
            brightness: parseInt(document.getElementById('brightness').value),
            contrast: parseInt(document.getElementById('contrast').value),
            pdf_split: document.getElementById('pdf-split').value,
            split_range: document.getElementById('split-range').value
        };
    }

    async updatePreviewWithProcessed(result) {
        const previewContent = document.getElementById('preview-content');
        
        if (result.preview_url) {
            // If server provides a preview URL, use it
            const embed = document.createElement('embed');
            embed.src = result.preview_url;
            embed.type = 'application/pdf';
            embed.style.width = '100%';
            embed.style.height = '300px';
            
            previewContent.innerHTML = '';
            previewContent.appendChild(embed);
        } else if (result.processed_file_id) {
            // Use the processed file ID to generate preview
            const previewUrl = `/api/files/${result.processed_file_id}/preview`;
            
            const embed = document.createElement('embed');
            embed.src = previewUrl;
            embed.type = 'application/pdf';
            embed.style.width = '100%';
            embed.style.height = '300px';
            
            previewContent.innerHTML = '';
            previewContent.appendChild(embed);
        }
    }
}

// Initialize dashboard when DOM is loaded
let dashboard;
document.addEventListener('DOMContentLoaded', () => {
    dashboard = new PrintDashboard();
});

// Make dashboard globally accessible for inline event handlers
window.dashboard = dashboard;
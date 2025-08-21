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
        document.getElementById('print-document').addEventListener('click', this.printDocument.bind(this));
        document.getElementById('clear-preview').addEventListener('click', this.clearPreview.bind(this));
        document.getElementById('test-print').addEventListener('click', this.printTestPage.bind(this));

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
                data.printers.forEach(printer => {
                    const option = document.createElement('option');
                    option.value = printer.id || printer.name;
                    option.textContent = printer.name;
                    select.appendChild(option);
                });
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
        this.updatePrinterStatus();
        
        const testPrintBtn = document.getElementById('test-print');
        testPrintBtn.disabled = !this.selectedPrinter;
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
        const jobId = this.jobIdCounter++;
        
        const job = {
            id: jobId,
            name: this.currentFile.name,
            printer: this.selectedPrinter,
            status: 'pending',
            progress: 0,
            timestamp: new Date(),
            settings: settings
        };

        this.addJobToQueue(job);
        this.showLoadingOverlay('Sending print job...');

        try {
            // Use the jobs/submit endpoint with uploaded file_id
            const requestData = {
                printer_id: this.selectedPrinter,
                file_id: this.currentFileId,
                settings: {
                    copies: settings.copies,
                    color_mode: settings.colorMode,
                    paper_size: settings.paperSize,
                    quality: settings.quality
                }
            };

            const response = await fetch('/api/jobs/submit', {
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
                this.showMessage(`Print job started: ${this.currentFile.name}`, 'success');
                this.simulateJobProgress(job);
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
        return {
            copies: parseInt(document.getElementById('copies').value) || 1,
            colorMode: document.getElementById('color-mode').value || 'auto',
            paperSize: document.getElementById('paper-size').value || 'A4',
            quality: document.getElementById('quality').value || 'normal'
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

    simulateJobProgress(job) {
        const interval = setInterval(() => {
            if (job.status === 'error' || job.status === 'cancelled') {
                clearInterval(interval);
                this.removeJobFromQueue(job.id);
                return;
            }

            job.progress += Math.random() * 15 + 5;
            
            if (job.progress >= 100) {
                job.progress = 100;
                job.status = 'completed';
                clearInterval(interval);
                setTimeout(() => {
                    this.removeJobFromQueue(job.id);
                }, 2000);
            }

            this.updateJobInQueue(job);
        }, 1000);
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
}

// Initialize dashboard when DOM is loaded
let dashboard;
document.addEventListener('DOMContentLoaded', () => {
    dashboard = new PrintDashboard();
});

// Make dashboard globally accessible for inline event handlers
window.dashboard = dashboard;
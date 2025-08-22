// Global variables
let currentFileId = null;
let currentPage = 1;
let refreshInterval = null;

// API base URL
const API_BASE = '/api';

// Initialize app
document.addEventListener('DOMContentLoaded', function() {
    initializeApp();
    
    // Initialize advanced print settings
    initializeAdvancedSettings();
});

// Initialize advanced print settings
function initializeAdvancedSettings() {
    // Handle fit to page change
    const fitToPageSelect = document.getElementById('fitToPage');
    if (fitToPageSelect) {
        fitToPageSelect.addEventListener('change', function() {
            toggleCustomScale();
        });
    }
    
    // Initialize custom scale visibility
    toggleCustomScale();
}

// Toggle custom scale input visibility
function toggleCustomScale() {
    const fitToPage = document.getElementById('fitToPage').value;
    const customScaleDiv = document.getElementById('customScaleDiv');
    
    if (fitToPage === 'custom') {
        customScaleDiv.style.display = 'block';
    } else {
        customScaleDiv.style.display = 'none';
    }
}

// Toggle page range input visibility
function togglePageRange() {
    const pageRangeType = document.getElementById('pageRangeType').value;
    const pageRangeDiv = document.getElementById('pageRangeDiv');
    
    if (pageRangeType === 'range') {
        pageRangeDiv.style.display = 'block';
    } else {
        pageRangeDiv.style.display = 'none';
    }
}

function initializeApp() {
    // Load initial data
    loadNetworkInfo();
    loadSystemStatus();
    loadPrinters();
    loadJobs();
    
    // Setup file drop zone
    setupFileDropZone();
    
    // Setup form handlers
    setupFormHandlers();
    
    // Start auto-refresh
    startAutoRefresh();
    
    // Setup tab change handlers
    document.getElementById('jobs-tab').addEventListener('click', () => {
        loadJobs();
    });
}

function setupFileDropZone() {
    const dropZone = document.getElementById('fileDropZone');
    
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('dragover');
    });
    
    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('dragover');
    });
    
    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('dragover');
        
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            handleFileSelect(files[0]);
        }
    });
}

function setupFormHandlers() {
    document.getElementById('printForm').addEventListener('submit', (e) => {
        e.preventDefault();
        submitPrintJob();
    });
    
    document.getElementById('jobStatusFilter').addEventListener('change', () => {
        currentPage = 1;
        loadJobs();
    });
}

function startAutoRefresh() {
    // Main refresh interval for general data
    refreshInterval = setInterval(() => {
        loadSystemStatus();
        
        // Refresh current tab content
        const activeTab = document.querySelector('.nav-link.active').id;
        if (activeTab === 'printers-tab') {
            loadPrinters();
        } else if (activeTab === 'jobs-tab') {
            loadJobs();
        }
    }, 30000); // Refresh every 30 seconds
    
    // More frequent refresh for printer status when on printers tab
    const printerStatusInterval = setInterval(() => {
        const activeTab = document.querySelector('.nav-link.active').id;
        if (activeTab === 'printers-tab') {
            refreshPrinterStatus();
        }
    }, 10000); // Refresh printer status every 10 seconds
    
    // Store both intervals for cleanup
    window.printerStatusInterval = printerStatusInterval;
}

// API functions
async function apiCall(endpoint, options = {}) {
    try {
        const response = await fetch(`${API_BASE}${endpoint}`, {
            headers: {
                'Content-Type': 'application/json',
                ...options.headers
            },
            ...options
        });
        
        if (!response.ok) {
            throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        }
        
        return await response.json();
    } catch (error) {
        console.error('API call failed:', error);
        showAlert('error', `API Error: ${error.message}`);
        throw error;
    }
}

// Load functions
async function loadNetworkInfo() {
    try {
        const response = await apiCall('/network');
        displayNetworkInfo(response.data);
    } catch (error) {
        document.getElementById('networkInfo').innerHTML = 
            '<div class="text-danger"><i class="fas fa-exclamation-triangle"></i> Failed to load</div>';
    }
}

async function loadSystemStatus() {
    try {
        const response = await apiCall('/status');
        displaySystemStatus(response);
    } catch (error) {
        document.getElementById('systemStatus').innerHTML = 
            '<div class="text-danger"><i class="fas fa-exclamation-triangle"></i> Failed to load</div>';
    }
}

async function loadPrinters() {
    try {
        const response = await apiCall('/printers/status');
        displayPrinters(response.data);
        updatePrinterSelect(response.data);
    } catch (error) {
        document.getElementById('printersContainer').innerHTML = 
            '<div class="alert alert-danger"><i class="fas fa-exclamation-triangle"></i> Failed to load printers</div>';
    }
}

async function refreshPrinterStatus() {
    try {
        const response = await apiCall('/printers/status');
        
        if (response.data) {
            displayPrinters(response.data);
            // Update status indicators without changing printer select
            updatePrinterStatusIndicators(response.data);
        }
    } catch (error) {
        console.error('Error refreshing printer status:', error);
    }
}

function updatePrinterStatusIndicators(printers) {
    // Update any status indicators in the UI
    const statusSummary = printers.reduce((acc, printer) => {
        if (printer.is_connected) {
            acc.connected++;
        } else {
            acc.disconnected++;
        }
        acc.total++;
        return acc;
    }, { connected: 0, disconnected: 0, total: 0 });
    
    // Update system status if element exists
    const systemStatusElement = document.getElementById('systemStatus');
    if (systemStatusElement) {
        const statusHtml = `
            <div class="d-flex justify-content-between">
                <span>Printers:</span>
                <span>
                    <span class="text-success">${statusSummary.connected} online</span>
                    ${statusSummary.disconnected > 0 ? `<span class="text-warning ms-2">${statusSummary.disconnected} offline</span>` : ''}
                </span>
            </div>
        `;
        systemStatusElement.innerHTML = statusHtml;
    }
}

async function loadJobs() {
    try {
        const status = document.getElementById('jobStatusFilter').value;
        const params = new URLSearchParams({
            page: currentPage,
            limit: 10
        });
        
        if (status) {
            params.append('status', status);
        }
        
        const response = await apiCall(`/jobs?${params}`);
        displayJobs(response.data, response.total, response.pages);
    } catch (error) {
        document.getElementById('jobsContainer').innerHTML = 
            '<div class="alert alert-danger"><i class="fas fa-exclamation-triangle"></i> Failed to load jobs</div>';
    }
}

// Display functions
function displayNetworkInfo(data) {
    const html = `
        <div class="small">
            <div class="mb-1"><strong>IP:</strong> ${data.local_ip}</div>
            <div class="mb-1"><strong>Host:</strong> ${data.hostname}</div>
            <div class="mb-1"><strong>Port:</strong> ${data.port}</div>
            <div class="mb-1">
                <strong>Broadcasting:</strong> 
                <span class="badge ${data.is_broadcasting === 'true' ? 'bg-success' : 'bg-secondary'}">
                    ${data.is_broadcasting === 'true' ? 'Active' : 'Inactive'}
                </span>
            </div>
            <div><strong>Discovered:</strong> ${data.discovered_count}</div>
        </div>
    `;
    document.getElementById('networkInfo').innerHTML = html;
}

function displaySystemStatus(data) {
    const html = `
        <div class="small">
            <div class="mb-1">
                <strong>Status:</strong> 
                <span class="badge bg-success">${data.status}</span>
            </div>
            <div class="mb-1"><strong>Version:</strong> ${data.version}</div>
            <div class="mb-1"><strong>Uptime:</strong> ${formatUptime(data.uptime)}</div>
            <div class="mb-1"><strong>Active Jobs:</strong> ${data.active_jobs}</div>
            <div><strong>Printers:</strong> ${data.total_printers}</div>
        </div>
    `;
    document.getElementById('systemStatus').innerHTML = html;
}

function displayPrinters(printers) {
    if (printers.length === 0) {
        document.getElementById('printersContainer').innerHTML = 
            '<div class="alert alert-info"><i class="fas fa-info-circle"></i> No printers found. Try refreshing or check your printer connections.</div>';
        return;
    }
    
    const html = printers.map(printer => {
        const connectionStatus = printer.is_connected ? 'connected' : 'disconnected';
        const connectionIcon = printer.is_connected ? 'fas fa-check-circle text-success' : 'fas fa-exclamation-triangle text-warning';
        const connectionText = printer.is_connected ? 'Connected' : 'Disconnected';
        
        return `
        <div class="col-md-6 col-lg-4 mb-3">
            <div class="card printer-card h-100 ${printer.is_connected ? '' : 'border-warning'}">
                <div class="card-body">
                    <div class="d-flex justify-content-between align-items-start mb-2">
                        <h6 class="card-title mb-0">${printer.name}</h6>
                        <span class="badge ${getStatusBadgeClass(printer.status)}">
                            ${printer.status}
                        </span>
                    </div>
                    
                    <div class="small text-muted mb-3">
                        <div><i class="fas fa-server me-1"></i> ${printer.driver || 'Unknown'}</div>
                        <div><i class="fas fa-map-marker-alt me-1"></i> ${printer.location || 'Local'}</div>
                        ${printer.is_default ? '<div><i class="fas fa-star me-1"></i> Default Printer</div>' : ''}
                        <div class="mt-2">
                            <i class="${connectionIcon} me-1"></i> 
                            <span class="${printer.is_connected ? 'text-success' : 'text-warning'}">${connectionText}</span>
                            ${printer.queue_length > 0 ? `<span class="badge bg-info ms-2">${printer.queue_length} jobs</span>` : ''}
                        </div>
                        ${!printer.is_connected && printer.connection_error ? 
                            `<div class="text-danger small mt-1"><i class="fas fa-exclamation-circle me-1"></i> ${printer.connection_error}</div>` : ''}
                    </div>
                    
                    <div class="d-flex gap-2">
                        <button class="btn btn-outline-primary btn-sm" onclick="showPrinterDetails('${printer.name}')">
                            <i class="fas fa-info-circle"></i> Details
                        </button>
                        <button class="btn btn-outline-success btn-sm" onclick="testPrinter('${printer.name}')" ${!printer.is_connected ? 'disabled' : ''}>
                            <i class="fas fa-print"></i> Test
                        </button>
                    </div>
                </div>
            </div>
        </div>
        `;
    }).join('');
    
    document.getElementById('printersContainer').innerHTML = `<div class="row">${html}</div>`;
}

function displayJobs(jobs, total, pages) {
    if (jobs.length === 0) {
        document.getElementById('jobsContainer').innerHTML = 
            '<div class="alert alert-info"><i class="fas fa-info-circle"></i> No jobs found</div>';
        document.getElementById('jobsPagination').style.display = 'none';
        return;
    }
    
    const html = `
        <div class="table-responsive">
            <table class="table table-hover">
                <thead>
                    <tr>
                        <th>Job ID</th>
                        <th>File</th>
                        <th>Printer</th>
                        <th>Status</th>
                        <th>Created</th>
                        <th>Actions</th>
                    </tr>
                </thead>
                <tbody>
                    ${jobs.map(job => `
                        <tr>
                            <td><code>${job.job_id.substring(0, 8)}</code></td>
                            <td>
                                <div class="fw-bold">${job.filename}</div>
                                <small class="text-muted">${job.copies} copies</small>
                            </td>
                            <td>${job.printer_name}</td>
                            <td>
                                <span class="job-status job-${job.status}">
                                    ${job.status}
                                </span>
                            </td>
                            <td>
                                <small>${formatDateTime(job.created_at)}</small>
                            </td>
                            <td>
                                <div class="btn-group btn-group-sm">
                                    ${job.status === 'pending' || job.status === 'printing' ? 
                                        `<button class="btn btn-outline-danger" onclick="cancelJob('${job.job_id}')">
                                            <i class="fas fa-times"></i>
                                        </button>` : ''}
                                    ${job.status === 'failed' ? 
                                        `<button class="btn btn-outline-primary" onclick="retryJob('${job.job_id}')">
                                            <i class="fas fa-redo"></i>
                                        </button>` : ''}
                                </div>
                            </td>
                        </tr>
                    `).join('')}
                </tbody>
            </table>
        </div>
    `;
    
    document.getElementById('jobsContainer').innerHTML = html;
    
    // Update pagination
    if (pages > 1) {
        updatePagination(pages, total);
        document.getElementById('jobsPagination').style.display = 'block';
    } else {
        document.getElementById('jobsPagination').style.display = 'none';
    }
}

function updatePrinterSelect(printers) {
    const select = document.getElementById('printerSelect');
    select.innerHTML = '<option value="">Select printer...</option>';
    
    printers.forEach(printer => {
        const option = document.createElement('option');
        option.value = printer.name;
        option.textContent = `${printer.name}${printer.is_default ? ' (Default)' : ''}`;
        if (printer.status !== 'ready') {
            option.disabled = true;
            option.textContent += ` - ${printer.status}`;
        }
        select.appendChild(option);
    });
}

// File handling
async function handleFileSelect(file) {
    if (!file) return;
    
    // Validate file type
    const allowedTypes = ['application/pdf', 'application/msword', 
                         'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                         'text/plain', 'image/jpeg', 'image/png'];
    
    if (!allowedTypes.includes(file.type)) {
        showAlert('error', 'File type not supported');
        return;
    }
    
    // Validate file size (10MB limit)
    if (file.size > 10 * 1024 * 1024) {
        showAlert('error', 'File size too large (max 10MB)');
        return;
    }
    
    try {
        // Show upload progress
        showFileInfo(file, 'uploading');
        
        // Upload file
        const formData = new FormData();
        formData.append('file', file);
        
        const response = await fetch(`${API_BASE}/files/upload`, {
            method: 'POST',
            body: formData
        });
        
        if (!response.ok) {
            throw new Error(`Upload failed: ${response.statusText}`);
        }
        
        const result = await response.json();
        
        if (result.success) {
            currentFileId = result.file_id;
            showFileInfo(file, 'ready', result);
            document.getElementById('printButton').disabled = false;
            showAlert('success', 'File uploaded successfully');
        } else {
            throw new Error(result.message || 'Upload failed');
        }
        
    } catch (error) {
        showFileInfo(file, 'error');
        showAlert('error', `Upload failed: ${error.message}`);
    }
}

function showFileInfo(file, status, uploadResult = null) {
    const fileInfo = document.getElementById('fileInfo');
    let html = '';
    
    switch (status) {
        case 'uploading':
            html = `
                <div class="alert alert-info">
                    <div class="d-flex align-items-center">
                        <div class="spinner-border spinner-border-sm me-2"></div>
                        <div>
                            <strong>Uploading:</strong> ${file.name}<br>
                            <small>Size: ${formatFileSize(file.size)}</small>
                        </div>
                    </div>
                </div>
            `;
            break;
            
        case 'ready':
            html = `
                <div class="alert alert-success">
                    <div class="d-flex justify-content-between align-items-start">
                        <div>
                            <strong><i class="fas fa-check-circle me-1"></i> Ready:</strong> ${file.name}<br>
                            <small>Size: ${formatFileSize(file.size)} | Type: ${file.type}</small>
                            ${uploadResult ? `<br><small>File ID: ${uploadResult.file_id}</small>` : ''}
                        </div>
                        <button class="btn btn-outline-danger btn-sm" onclick="clearFile()">
                            <i class="fas fa-times"></i>
                        </button>
                    </div>
                </div>
            `;
            break;
            
        case 'error':
            html = `
                <div class="alert alert-danger">
                    <strong><i class="fas fa-exclamation-triangle me-1"></i> Error:</strong> ${file.name}<br>
                    <small>Upload failed</small>
                </div>
            `;
            break;
    }
    
    fileInfo.innerHTML = html;
    fileInfo.style.display = 'block';
}

function clearFile() {
    currentFileId = null;
    document.getElementById('fileInfo').style.display = 'none';
    document.getElementById('fileInput').value = '';
    document.getElementById('printButton').disabled = true;
}

// Print job submission
async function submitPrintJob() {
    if (!currentFileId) {
        showAlert('error', 'Please upload a file first');
        return;
    }
    
    const printerName = document.getElementById('printerSelect').value;
    if (!printerName) {
        showAlert('error', 'Please select a printer');
        return;
    }
    
    // Advanced settings
    const fitToPage = document.getElementById('fitToPage').value;
    const customScale = document.getElementById('customScale').value;
    const pageRangeType = document.getElementById('pageRangeType').value;
    const pageRange = document.getElementById('pageRange').value;
    const marginTop = document.getElementById('marginTop').value;
    const marginBottom = document.getElementById('marginBottom').value;
    const marginLeft = document.getElementById('marginLeft').value;
    const marginRight = document.getElementById('marginRight').value;
    const centerHorizontally = document.getElementById('centerHorizontally').checked;
    const centerVertically = document.getElementById('centerVertically').checked;
    
    // Validate page range if specified
    if (pageRangeType === 'range' && pageRange) {
        if (!validatePageRange(pageRange)) {
            showAlert('error', 'Invalid page range format. Use examples: 1-5, 1,3,5, 1-3,5,7-9');
            return;
        }
    }
    
    const formData = new FormData();
    formData.append('file_id', currentFileId);
    formData.append('printer_name', printerName);
    formData.append('copies', document.getElementById('copies').value);
    formData.append('color_mode', document.getElementById('colorMode').value);
    formData.append('paper_size', document.getElementById('paperSize').value);
    formData.append('orientation', document.getElementById('orientation').value);
    formData.append('quality', document.getElementById('quality').value);
    formData.append('duplex', document.getElementById('duplex').value);
    
    // Advanced settings
    formData.append('fit_to_page', fitToPage);
    if (fitToPage === 'custom') {
        formData.append('custom_scale', customScale);
    }
    formData.append('page_range_type', pageRangeType);
    if (pageRangeType === 'range' && pageRange) {
        formData.append('page_range', pageRange);
    }
    formData.append('margin_top', marginTop);
    formData.append('margin_bottom', marginBottom);
    formData.append('margin_left', marginLeft);
    formData.append('margin_right', marginRight);
    formData.append('center_horizontally', centerHorizontally);
    formData.append('center_vertically', centerVertically);
    
    try {
        const response = await fetch(`${API_BASE}/jobs/submit`, {
            method: 'POST',
            body: formData
        });
        
        const result = await response.json();
        
        if (result.success) {
            showAlert('success', `Print job submitted successfully (ID: ${result.job_id.substring(0, 8)})`);
            clearFile();
            document.getElementById('printForm').reset();
            document.getElementById('copies').value = '1';
            
            // Reset advanced settings to defaults
            resetAdvancedSettings();
            
            // Switch to jobs tab
            document.getElementById('jobs-tab').click();
        } else {
            throw new Error(result.message || 'Failed to submit job');
        }
        
    } catch (error) {
        showAlert('error', `Failed to submit print job: ${error.message}`);
    }
}

// Validate page range format
function validatePageRange(pageRange) {
    // Allow formats like: 1-5, 1,3,5, 1-3,5,7-9
    const pattern = /^\d+(-\d+)?(,\d+(-\d+)?)*$/;
    return pattern.test(pageRange.replace(/\s/g, ''));
}

// Reset advanced settings to defaults
function resetAdvancedSettings() {
    document.getElementById('fitToPage').value = 'fit';
    document.getElementById('customScale').value = '100';
    document.getElementById('pageRangeType').value = 'all';
    document.getElementById('pageRange').value = '';
    document.getElementById('marginTop').value = '20';
    document.getElementById('marginBottom').value = '20';
    document.getElementById('marginLeft').value = '20';
    document.getElementById('marginRight').value = '20';
    document.getElementById('centerHorizontally').checked = true;
    document.getElementById('centerVertically').checked = true;
    
    // Hide conditional divs
    document.getElementById('customScaleDiv').style.display = 'none';
    document.getElementById('pageRangeDiv').style.display = 'none';
}

// Printer actions
async function showPrinterDetails(printerName) {
    try {
        const response = await apiCall(`/printers/${encodeURIComponent(printerName)}`);
        
        const html = `
            <div class="row">
                <div class="col-md-6">
                    <h6>Printer Information</h6>
                    <table class="table table-sm">
                        <tr><td><strong>Name:</strong></td><td>${response.printer.name}</td></tr>
                        <tr><td><strong>Status:</strong></td><td><span class="badge ${getStatusBadgeClass(response.printer.status)}">${response.printer.status}</span></td></tr>
                        <tr><td><strong>Driver:</strong></td><td>${response.printer.driver || 'Unknown'}</td></tr>
                        <tr><td><strong>Location:</strong></td><td>${response.printer.location || 'Local'}</td></tr>
                        <tr><td><strong>Default:</strong></td><td>${response.printer.is_default ? 'Yes' : 'No'}</td></tr>
                    </table>
                </div>
                <div class="col-md-6">
                    <h6>Queue Status</h6>
                    <table class="table table-sm">
                        <tr><td><strong>Queue Length:</strong></td><td>${response.queue_length}</td></tr>
                        <tr><td><strong>Status:</strong></td><td>${response.status}</td></tr>
                    </table>
                    
                    ${response.jobs.length > 0 ? `
                        <h6 class="mt-3">Recent Jobs</h6>
                        <div class="list-group list-group-flush">
                            ${response.jobs.map(job => `
                                <div class="list-group-item px-0 py-2">
                                    <div class="d-flex justify-content-between">
                                        <small>${job.filename}</small>
                                        <span class="job-status job-${job.status}">${job.status}</span>
                                    </div>
                                </div>
                            `).join('')}
                        </div>
                    ` : ''}
                </div>
            </div>
        `;
        
        document.getElementById('printerModalBody').innerHTML = html;
        new bootstrap.Modal(document.getElementById('printerModal')).show();
        
    } catch (error) {
        showAlert('error', `Failed to load printer details: ${error.message}`);
    }
}

async function testPrinter(printerName) {
    try {
        const response = await apiCall(`/printers/${encodeURIComponent(printerName)}/test`, {
            method: 'POST'
        });
        showAlert('success', `Test page sent to ${printerName}`);
        
        // Refresh jobs to show the test print job
        setTimeout(() => {
            loadJobs();
        }, 1000);
    } catch (error) {
        showAlert('error', `Failed to send test page: ${error.message}`);
    }
}

// Job actions
async function cancelJob(jobId) {
    if (!confirm('Are you sure you want to cancel this job?')) return;
    
    try {
        await apiCall(`/jobs/${jobId}/cancel`, { method: 'POST' });
        showAlert('success', 'Job cancelled successfully');
        loadJobs();
    } catch (error) {
        showAlert('error', `Failed to cancel job: ${error.message}`);
    }
}

async function retryJob(jobId) {
    try {
        await apiCall(`/jobs/${jobId}/retry`, { method: 'POST' });
        showAlert('success', 'Job queued for retry');
        loadJobs();
    } catch (error) {
        showAlert('error', `Failed to retry job: ${error.message}`);
    }
}

// Network actions
async function scanNetwork() {
    try {
        const button = event.target;
        const originalText = button.innerHTML;
        button.innerHTML = '<i class="fas fa-spinner fa-spin me-1"></i> Scanning...';
        button.disabled = true;
        
        const response = await apiCall('/discovery/scan', { method: 'POST' });
        showAlert('success', response.message);
        
        // Reload discovered printers
        setTimeout(() => {
            loadNetworkInfo();
        }, 2000);
        
        button.innerHTML = originalText;
        button.disabled = false;
        
    } catch (error) {
        showAlert('error', `Network scan failed: ${error.message}`);
        event.target.disabled = false;
    }
}

// Utility functions
function refreshAll() {
    loadNetworkInfo();
    loadSystemStatus();
    loadPrinters();
    loadJobs();
    showAlert('info', 'Data refreshed');
}

function showAlert(type, message) {
    const alertsContainer = document.getElementById('alerts');
    const alertId = 'alert-' + Date.now();
    
    const alertHtml = `
        <div class="alert alert-${type} alert-dismissible fade show" id="${alertId}">
            ${message}
            <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
        </div>
    `;
    
    alertsContainer.insertAdjacentHTML('beforeend', alertHtml);
    
    // Auto-dismiss after 5 seconds
    setTimeout(() => {
        const alert = document.getElementById(alertId);
        if (alert) {
            bootstrap.Alert.getOrCreateInstance(alert).close();
        }
    }, 5000);
}

function getStatusBadgeClass(status) {
    switch (status) {
        case 'ready': return 'bg-success';
        case 'busy': return 'bg-warning';
        case 'offline': return 'bg-danger';
        case 'error': return 'bg-danger';
        default: return 'bg-secondary';
    }
}

function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

function formatDateTime(dateString) {
    const date = new Date(dateString);
    return date.toLocaleString();
}

function formatUptime(seconds) {
    const hours = Math.floor(seconds / 3600);
    const minutes = Math.floor((seconds % 3600) / 60);
    return `${hours}h ${minutes}m`;
}

function updatePagination(totalPages, totalItems) {
    const paginationList = document.getElementById('paginationList');
    let html = '';
    
    // Previous button
    html += `
        <li class="page-item ${currentPage === 1 ? 'disabled' : ''}">
            <a class="page-link" href="#" onclick="changePage(${currentPage - 1})">
                <i class="fas fa-chevron-left"></i>
            </a>
        </li>
    `;
    
    // Page numbers
    const startPage = Math.max(1, currentPage - 2);
    const endPage = Math.min(totalPages, currentPage + 2);
    
    for (let i = startPage; i <= endPage; i++) {
        html += `
            <li class="page-item ${i === currentPage ? 'active' : ''}">
                <a class="page-link" href="#" onclick="changePage(${i})">${i}</a>
            </li>
        `;
    }
    
    // Next button
    html += `
        <li class="page-item ${currentPage === totalPages ? 'disabled' : ''}">
            <a class="page-link" href="#" onclick="changePage(${currentPage + 1})">
                <i class="fas fa-chevron-right"></i>
            </a>
        </li>
    `;
    
    paginationList.innerHTML = html;
}

function changePage(page) {
    if (page < 1) return;
    currentPage = page;
    loadJobs();
}

// Cleanup on page unload
window.addEventListener('beforeunload', () => {
    if (refreshInterval) {
        clearInterval(refreshInterval);
    }
    if (window.printerStatusInterval) {
        clearInterval(window.printerStatusInterval);
    }
});
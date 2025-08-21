// Configuration management JavaScript
const API_BASE = '/api';

// DOM elements
const currentConfigDiv = document.getElementById('current-config');
const configForm = document.getElementById('config-form');
const reloadButton = document.getElementById('reload-config');
const messageDiv = document.getElementById('message');

// Initialize page
document.addEventListener('DOMContentLoaded', function() {
    loadCurrentConfig();
    setupEventListeners();
});

function setupEventListeners() {
    // Form submission
    configForm.addEventListener('submit', handleConfigUpdate);
    
    // Reload button
    reloadButton.addEventListener('click', handleConfigReload);
    
    // Auto-discovery checkbox change
    const autoDiscoveryCheckbox = document.getElementById('auto-discovery');
    autoDiscoveryCheckbox.addEventListener('change', function() {
        const defaultFields = ['default-name', 'default-id'];
        defaultFields.forEach(fieldId => {
            const field = document.getElementById(fieldId);
            field.disabled = this.checked;
            if (this.checked) {
                field.style.opacity = '0.5';
            } else {
                field.style.opacity = '1';
            }
        });
    });
    
    // Fallback enabled checkbox change
    const fallbackEnabledCheckbox = document.getElementById('fallback-enabled');
    fallbackEnabledCheckbox.addEventListener('change', function() {
        const fallbackKeywords = document.getElementById('fallback-keywords');
        fallbackKeywords.disabled = !this.checked;
        if (!this.checked) {
            fallbackKeywords.style.opacity = '0.5';
        } else {
            fallbackKeywords.style.opacity = '1';
        }
    });
}

async function loadCurrentConfig() {
    try {
        showMessage('Loading configuration...', 'info');
        
        const response = await fetch(`${API_BASE}/config/printers`);
        const result = await response.json();
        
        if (result.success) {
            displayCurrentConfig(result.data);
            populateForm(result.data);
            hideMessage();
        } else {
            throw new Error(result.message || 'Failed to load configuration');
        }
    } catch (error) {
        console.error('Error loading configuration:', error);
        showMessage(`Error loading configuration: ${error.message}`, 'error');
        currentConfigDiv.innerHTML = '<div class="error">Failed to load configuration</div>';
    }
}

function displayCurrentConfig(config) {
    const configHtml = `
        <div class="config-item">
            <strong>Auto Discovery:</strong> ${config.auto_discovery ? 'Enabled' : 'Disabled'}
        </div>
        <div class="config-item">
            <strong>Default Printer:</strong> ${config.default_name || 'Not set'}
        </div>
        <div class="config-item">
            <strong>Default Printer ID:</strong> ${config.default_id || 'Not set'}
        </div>
        <div class="config-item">
            <strong>Refresh Interval:</strong> ${config.refresh_interval || 60} seconds
        </div>
        <div class="config-item">
            <strong>Status Check Interval:</strong> ${config.status_check_interval || 5} seconds
        </div>
        <div class="config-item">
            <strong>Show Connection Status:</strong> ${config.show_connection_status ? 'Yes' : 'No'}
        </div>
        <div class="config-item">
            <strong>Show Visual Indicators:</strong> ${config.show_visual_indicator ? 'Yes' : 'No'}
        </div>
        <div class="config-item">
            <strong>Auto Refresh Status:</strong> ${config.auto_refresh_status ? 'Yes' : 'No'}
        </div>
        <div class="config-item">
            <strong>Fallback Search:</strong> ${config.fallback_enabled ? 'Enabled' : 'Disabled'}
        </div>
        ${config.fallback_enabled ? `
        <div class="config-item">
            <strong>Fallback Keywords:</strong> ${Array.isArray(config.fallback_keywords) ? config.fallback_keywords.join(', ') : (config.fallback_keywords || 'Not set')}
        </div>
        ` : ''}
    `;
    
    currentConfigDiv.innerHTML = configHtml;
}

function populateForm(config) {
    // Populate form fields with current configuration
    document.getElementById('auto-discovery').checked = config.auto_discovery || false;
    document.getElementById('default-name').value = config.default_name || '';
    document.getElementById('default-id').value = config.default_id || '';
    document.getElementById('refresh-interval').value = config.refresh_interval || 60;
    document.getElementById('status-check-interval').value = config.status_check_interval || 5;
    document.getElementById('show-connection-status').checked = config.show_connection_status !== false;
    document.getElementById('show-visual-indicator').checked = config.show_visual_indicator !== false;
    document.getElementById('auto-refresh-status').checked = config.auto_refresh_status !== false;
    document.getElementById('fallback-enabled').checked = config.fallback_enabled || false;
    
    // Handle fallback keywords (could be array or string)
    let fallbackKeywords = '';
    if (Array.isArray(config.fallback_keywords)) {
        fallbackKeywords = config.fallback_keywords.join(', ');
    } else if (config.fallback_keywords) {
        fallbackKeywords = config.fallback_keywords;
    }
    document.getElementById('fallback-keywords').value = fallbackKeywords;
    
    // Trigger change events to update UI state
    document.getElementById('auto-discovery').dispatchEvent(new Event('change'));
    document.getElementById('fallback-enabled').dispatchEvent(new Event('change'));
}

async function handleConfigUpdate(event) {
    event.preventDefault();
    
    try {
        showMessage('Updating configuration...', 'info');
        
        const formData = new FormData(configForm);
        const configData = {};
        
        // Process form data
        for (const [key, value] of formData.entries()) {
            if (key === 'fallback_keywords') {
                // Split comma-separated keywords into array
                configData[key] = value.split(',').map(k => k.trim()).filter(k => k.length > 0);
            } else if (key.endsWith('_interval') || key === 'refresh_interval') {
                configData[key] = parseInt(value);
            } else {
                configData[key] = value;
            }
        }
        
        // Handle checkboxes (they won't be in formData if unchecked)
        const checkboxes = ['auto_discovery', 'show_connection_status', 'show_visual_indicator', 'auto_refresh_status', 'fallback_enabled'];
        checkboxes.forEach(checkbox => {
            if (!(checkbox in configData)) {
                configData[checkbox] = false;
            } else {
                configData[checkbox] = true;
            }
        });
        
        const response = await fetch(`${API_BASE}/config/printers`, {
            method: 'PUT',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(configData)
        });
        
        const result = await response.json();
        
        if (result.success) {
            showMessage('Configuration updated successfully!', 'success');
            // Reload current configuration to reflect changes
            setTimeout(() => {
                loadCurrentConfig();
            }, 1000);
        } else {
            throw new Error(result.message || 'Failed to update configuration');
        }
    } catch (error) {
        console.error('Error updating configuration:', error);
        showMessage(`Error updating configuration: ${error.message}`, 'error');
    }
}

async function handleConfigReload() {
    try {
        showMessage('Reloading configuration from file...', 'info');
        
        const response = await fetch(`${API_BASE}/config/printers/reload`, {
            method: 'POST'
        });
        
        const result = await response.json();
        
        if (result.success) {
            showMessage('Configuration reloaded successfully!', 'success');
            // Reload current configuration to reflect changes
            setTimeout(() => {
                loadCurrentConfig();
            }, 1000);
        } else {
            throw new Error(result.message || 'Failed to reload configuration');
        }
    } catch (error) {
        console.error('Error reloading configuration:', error);
        showMessage(`Error reloading configuration: ${error.message}`, 'error');
    }
}

function showMessage(message, type = 'info') {
    messageDiv.textContent = message;
    messageDiv.className = `message ${type}`;
    messageDiv.style.display = 'block';
    
    // Auto-hide success and info messages after 5 seconds
    if (type === 'success' || type === 'info') {
        setTimeout(() => {
            hideMessage();
        }, 5000);
    }
}

function hideMessage() {
    messageDiv.style.display = 'none';
    messageDiv.textContent = '';
    messageDiv.className = 'message';
}
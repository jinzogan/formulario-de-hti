// Main JavaScript file for Excel processor application

document.addEventListener('DOMContentLoaded', function() {
    // Initialize components
    initFileUpload();
    initFormValidation();
    initTooltips();
    
    // Auto-dismiss alerts after 5 seconds
    setTimeout(function() {
        const alerts = document.querySelectorAll('.alert');
        alerts.forEach(function(alert) {
            const bsAlert = new bootstrap.Alert(alert);
            bsAlert.close();
        });
    }, 5000);
});

function initFileUpload() {
    const fileInput = document.getElementById('file');
    const uploadForm = document.getElementById('uploadForm');
    const uploadBtn = document.getElementById('uploadBtn');
    
    if (!fileInput || !uploadForm || !uploadBtn) return;
    
    // File validation
    fileInput.addEventListener('change', function() {
        const file = this.files[0];
        
        if (file) {
            // Check file size (16MB = 16 * 1024 * 1024 bytes)
            const maxSize = 16 * 1024 * 1024;
            if (file.size > maxSize) {
                showAlert('El archivo es demasiado grande. Tamaño máximo: 16MB', 'error');
                this.value = '';
                return;
            }
            
            // Check file extension
            const allowedExtensions = ['xlsx', 'xls'];
            const fileExtension = file.name.split('.').pop().toLowerCase();
            
            if (!allowedExtensions.includes(fileExtension)) {
                showAlert('Tipo de archivo no permitido. Use archivos .xlsx o .xls', 'error');
                this.value = '';
                return;
            }
            
            // Update button text with filename
            const fileName = file.name.length > 20 ? file.name.substring(0, 20) + '...' : file.name;
            uploadBtn.innerHTML = `<i data-feather="upload"></i> Procesar: ${fileName}`;
            feather.replace();
        } else {
            uploadBtn.innerHTML = '<i data-feather="upload"></i> Procesar Archivo';
            feather.replace();
        }
    });
    
    // Form submission with loading state
    uploadForm.addEventListener('submit', function(e) {
        const file = fileInput.files[0];
        
        if (!file) {
            e.preventDefault();
            showAlert('Por favor seleccione un archivo', 'error');
            return;
        }
        
        // Show loading state
        uploadBtn.classList.add('btn-loading');
        uploadBtn.disabled = true;
        uploadBtn.innerHTML = 'Procesando...';
        
        // Show progress message
        showAlert('Procesando archivo, por favor espere...', 'info');
    });
}

function initFormValidation() {
    const dataForm = document.getElementById('dataForm');
    
    if (!dataForm) return;
    
    dataForm.addEventListener('submit', function(e) {
        const submitBtn = this.querySelector('button[type="submit"]');
        
        if (submitBtn) {
            submitBtn.classList.add('btn-loading');
            submitBtn.disabled = true;
            submitBtn.innerHTML = 'Procesando...';
        }
    });
}

function initTooltips() {
    // Initialize Bootstrap tooltips
    const tooltipTriggerList = [].slice.call(document.querySelectorAll('[data-bs-toggle="tooltip"]'));
    tooltipTriggerList.map(function(tooltipTriggerEl) {
        return new bootstrap.Tooltip(tooltipTriggerEl);
    });
}

function showAlert(message, type = 'info') {
    // Remove existing alerts
    const existingAlerts = document.querySelectorAll('.alert');
    existingAlerts.forEach(alert => alert.remove());
    
    // Create new alert
    const alertDiv = document.createElement('div');
    alertDiv.className = `alert alert-${type === 'error' ? 'danger' : type} alert-dismissible fade show`;
    alertDiv.setAttribute('role', 'alert');
    
    const icon = type === 'error' ? 'alert-circle' : (type === 'success' ? 'check-circle' : 'info');
    
    alertDiv.innerHTML = `
        <i data-feather="${icon}"></i>
        ${message}
        <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
    `;
    
    // Insert at the top of main content
    const mainContent = document.querySelector('main .container');
    if (mainContent) {
        mainContent.insertBefore(alertDiv, mainContent.firstChild);
        feather.replace();
        
        // Auto-dismiss after 5 seconds
        setTimeout(() => {
            if (alertDiv.parentNode) {
                const bsAlert = new bootstrap.Alert(alertDiv);
                bsAlert.close();
            }
        }, 5000);
    }
}

// Utility functions
function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

function validateExcelFile(file) {
    const allowedTypes = [
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // .xlsx
        'application/vnd.ms-excel' // .xls
    ];
    
    const allowedExtensions = ['xlsx', 'xls'];
    const fileExtension = file.name.split('.').pop().toLowerCase();
    
    return allowedTypes.includes(file.type) || allowedExtensions.includes(fileExtension);
}

// Export functions for testing
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        formatFileSize,
        validateExcelFile,
        showAlert
    };
}

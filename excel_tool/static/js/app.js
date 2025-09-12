// 全局JavaScript功能

// 文件拖拽上传功能
document.addEventListener('DOMContentLoaded', function() {
    setupFileDragDrop();
    setupFormValidation();
    setupTooltips();
});

// 设置文件拖拽上传
function setupFileDragDrop() {
    const fileInput = document.getElementById('files');
    const uploadForm = document.getElementById('uploadForm');
    
    if (!fileInput || !uploadForm) return;
    
    const dropArea = uploadForm.querySelector('.card-body');
    if (!dropArea) return;
    
    // 防止默认行为
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropArea.addEventListener(eventName, preventDefaults, false);
        document.body.addEventListener(eventName, preventDefaults, false);
    });
    
    // 添加视觉反馈
    ['dragenter', 'dragover'].forEach(eventName => {
        dropArea.addEventListener(eventName, highlight, false);
    });
    
    ['dragleave', 'drop'].forEach(eventName => {
        dropArea.addEventListener(eventName, unhighlight, false);
    });
    
    // 处理文件拖放
    dropArea.addEventListener('drop', handleDrop, false);
    
    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }
    
    function highlight(e) {
        dropArea.classList.add('drag-over');
    }
    
    function unhighlight(e) {
        dropArea.classList.remove('drag-over');
    }
    
    function handleDrop(e) {
        const dt = e.dataTransfer;
        const files = dt.files;
        
        if (files.length > 0) {
            fileInput.files = files;
            updateFileInputDisplay(files);
        }
    }
}

// 更新文件输入显示
function updateFileInputDisplay(files) {
    const fileInput = document.getElementById('files');
    if (!fileInput) return;
    
    // 创建文件列表显示
    const fileList = Array.from(files).map(file => {
        const size = (file.size / 1024 / 1024).toFixed(2);
        return `${file.name} (${size} MB)`;
    }).join(', ');
    
    // 显示提示信息
    const formText = fileInput.parentElement.querySelector('.form-text');
    if (formText) {
        formText.innerHTML = `已选择 ${files.length} 个文件: ${fileList}`;
        formText.style.color = '#28a745';
    }
}

// 设置表单验证
function setupFormValidation() {
    const forms = document.querySelectorAll('form');
    
    forms.forEach(form => {
        form.addEventListener('submit', function(e) {
            if (!form.checkValidity()) {
                e.preventDefault();
                e.stopPropagation();
            }
            form.classList.add('was-validated');
        });
    });
}

// 设置提示工具
function setupTooltips() {
    const tooltipTriggerList = [].slice.call(document.querySelectorAll('[data-bs-toggle="tooltip"]'));
    tooltipTriggerList.map(function (tooltipTriggerEl) {
        return new bootstrap.Tooltip(tooltipTriggerEl);
    });
}

// 通用AJAX请求函数
function makeRequest(url, options = {}) {
    const defaultOptions = {
        method: 'GET',
        headers: {
            'Content-Type': 'application/json',
        }
    };
    
    const config = { ...defaultOptions, ...options };
    
    return fetch(url, config)
        .then(response => {
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            return response.json();
        });
}

// 显示加载状态
function showLoading(element, message = '加载中...') {
    if (typeof element === 'string') {
        element = document.querySelector(element);
    }
    
    if (!element) return;
    
    const loadingHtml = `
        <div class="d-flex justify-content-center align-items-center p-4">
            <div class="spinner-border me-2" role="status">
                <span class="visually-hidden">Loading...</span>
            </div>
            <span>${message}</span>
        </div>
    `;
    
    element.innerHTML = loadingHtml;
}

// 隐藏加载状态
function hideLoading(element) {
    if (typeof element === 'string') {
        element = document.querySelector(element);
    }
    
    if (!element) return;
    
    const spinner = element.querySelector('.spinner-border');
    if (spinner) {
        spinner.parentElement.remove();
    }
}

// 显示错误消息
function showError(message, container = null) {
    const alertHtml = `
        <div class="alert alert-danger alert-dismissible fade show" role="alert">
            <i class="bi bi-exclamation-triangle me-2"></i>
            ${message}
            <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
        </div>
    `;
    
    if (container) {
        if (typeof container === 'string') {
            container = document.querySelector(container);
        }
        container.insertAdjacentHTML('afterbegin', alertHtml);
    } else {
        // 添加到页面顶部
        const mainContainer = document.querySelector('main.container');
        if (mainContainer) {
            mainContainer.insertAdjacentHTML('afterbegin', alertHtml);
        }
    }
    
    // 自动滚动到错误消息
    const newAlert = document.querySelector('.alert-danger');
    if (newAlert) {
        newAlert.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }
}

// 显示成功消息
function showSuccess(message, container = null) {
    const alertHtml = `
        <div class="alert alert-success alert-dismissible fade show" role="alert">
            <i class="bi bi-check-circle me-2"></i>
            ${message}
            <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
        </div>
    `;
    
    if (container) {
        if (typeof container === 'string') {
            container = document.querySelector(container);
        }
        container.insertAdjacentHTML('afterbegin', alertHtml);
    } else {
        const mainContainer = document.querySelector('main.container');
        if (mainContainer) {
            mainContainer.insertAdjacentHTML('afterbegin', alertHtml);
        }
    }
}

// 格式化文件大小
function formatFileSize(bytes) {
    if (bytes === 0) return '0 B';
    
    const k = 1024;
    const sizes = ['B', 'KB', 'MB', 'GB', 'TB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

// 格式化日期时间
function formatDateTime(dateString) {
    if (!dateString) return '';
    
    const date = new Date(dateString);
    return date.toLocaleString('zh-CN', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit'
    });
}

// 复制到剪贴板
function copyToClipboard(text, successMessage = '已复制到剪贴板') {
    if (navigator.clipboard && window.isSecureContext) {
        navigator.clipboard.writeText(text).then(() => {
            showSuccess(successMessage);
        }).catch(err => {
            console.error('复制失败:', err);
            fallbackCopyTextToClipboard(text, successMessage);
        });
    } else {
        fallbackCopyTextToClipboard(text, successMessage);
    }
}

// 备用复制方法
function fallbackCopyTextToClipboard(text, successMessage) {
    const textArea = document.createElement('textarea');
    textArea.value = text;
    textArea.style.position = 'fixed';
    textArea.style.left = '-999999px';
    textArea.style.top = '-999999px';
    document.body.appendChild(textArea);
    textArea.focus();
    textArea.select();
    
    try {
        document.execCommand('copy');
        showSuccess(successMessage);
    } catch (err) {
        console.error('复制失败:', err);
        showError('复制失败，请手动复制');
    }
    
    document.body.removeChild(textArea);
}

// 防抖函数
function debounce(func, wait, immediate) {
    let timeout;
    return function executedFunction(...args) {
        const later = () => {
            timeout = null;
            if (!immediate) func(...args);
        };
        const callNow = immediate && !timeout;
        clearTimeout(timeout);
        timeout = setTimeout(later, wait);
        if (callNow) func(...args);
    };
}

// 节流函数
function throttle(func, limit) {
    let inThrottle;
    return function(...args) {
        if (!inThrottle) {
            func.apply(this, args);
            inThrottle = true;
            setTimeout(() => inThrottle = false, limit);
        }
    };
}

// 确认对话框
function confirmAction(message, onConfirm, onCancel = null) {
    if (confirm(message)) {
        if (typeof onConfirm === 'function') {
            onConfirm();
        }
    } else {
        if (typeof onCancel === 'function') {
            onCancel();
        }
    }
}

// 页面可见性检测
function handleVisibilityChange() {
    if (document.hidden) {
        // 页面隐藏时的处理
        console.log('页面隐藏');
    } else {
        // 页面显示时的处理
        console.log('页面显示');
    }
}

document.addEventListener('visibilitychange', handleVisibilityChange);

// 网络状态检测
function handleNetworkChange() {
    if (navigator.onLine) {
        showSuccess('网络连接已恢复');
    } else {
        showError('网络连接已断开');
    }
}

window.addEventListener('online', handleNetworkChange);
window.addEventListener('offline', handleNetworkChange);

// 全局错误处理
window.addEventListener('error', function(e) {
    console.error('全局错误:', e.error);
    showError('页面发生错误，请刷新重试');
});

// 全局Promise错误处理
window.addEventListener('unhandledrejection', function(e) {
    console.error('未处理的Promise错误:', e.reason);
    showError('请求失败，请重试');
    e.preventDefault();
});

// 页面加载完成后的初始化
document.addEventListener('DOMContentLoaded', function() {
    // 添加页面加载动画
    document.body.classList.add('fade-in');
    
    // 设置所有外部链接在新窗口打开
    const externalLinks = document.querySelectorAll('a[href^="http"]:not([href*="' + window.location.hostname + '"])');
    externalLinks.forEach(link => {
        link.setAttribute('target', '_blank');
        link.setAttribute('rel', 'noopener noreferrer');
    });
    
    // 设置表格响应式
    const tables = document.querySelectorAll('table:not(.table-responsive table)');
    tables.forEach(table => {
        if (!table.closest('.table-responsive')) {
            const wrapper = document.createElement('div');
            wrapper.className = 'table-responsive';
            table.parentNode.insertBefore(wrapper, table);
            wrapper.appendChild(table);
        }
    });
    
    console.log('页面初始化完成');
});

// 导出全局函数供模板使用
window.ExcelTool = {
    makeRequest,
    showLoading,
    hideLoading,
    showError,
    showSuccess,
    formatFileSize,
    formatDateTime,
    copyToClipboard,
    debounce,
    throttle,
    confirmAction
};
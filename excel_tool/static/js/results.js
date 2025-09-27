// 全局变量
let selectedFiles = new Set();
let currentResults = [];
let currentPreviewData = null;
let pivotChart = null;
let visualChart = null;

// 页面加载完成后初始化
document.addEventListener('DOMContentLoaded', function() {
    loadResults();
    
    // 初始化排序设置事件监听器
    initializeSortSettings();
});

// 初始化排序设置
function initializeSortSettings() {
    // 监听排序方式变化，显示/隐藏自定义排序输入框
    const sortTypeSelect = document.getElementById('sortType');
    if (sortTypeSelect) {
        sortTypeSelect.addEventListener('change', function() {
            toggleCustomSortInput();
        });
    }
}

// 加载结果列表
async function loadResults() {
    try {
        const response = await fetch('/api/results');
        const data = await response.json();
        
        if (data.success) {
            currentResults = data.results;
            renderResults(data.results);
        } else {
            showError('加载结果失败：' + data.message);
        }
    } catch (error) {
        console.error('加载结果失败:', error);
        showError('加载结果失败，请稍后重试');
    }
}

// 渲染结果列表
function renderResults(results) {
    const container = document.getElementById('resultsContainer');
    
    if (!results || results.length === 0) {
        container.innerHTML = `
            <div class="no-results">
                <i class="bi bi-folder-x" style="font-size: 3rem; color: #ddd;"></i>
                <h4 class="mt-3">暂无处理结果</h4>
                <p>请先在<a href="${window.location.origin}">首页</a>处理Excel文件</p>
            </div>
        `;
        return;
    }
    
    const html = results.map(result => `
        <div class="card result-card mb-3" data-filename="${result.filename}">
            <div class="card-body">
                <div class="row align-items-center">
                    <div class="col-md-1">
                        <div class="form-check">
                            <input class="form-check-input" type="checkbox" 
                                   id="check_${result.id}" 
                                   onchange="toggleFileSelection('${result.filename}')">
                        </div>
                    </div>
                    <div class="col-md-5">
                        <h6 class="card-title mb-1">
                            <i class="bi bi-file-earmark-excel text-success"></i>
                            ${result.filename}
                        </h6>
                        <div class="file-size">${result.size_formatted}</div>
                    </div>
                    <div class="col-md-3">
                        <div class="file-date">
                            <i class="bi bi-calendar"></i> ${result.created_time}<br>
                            <i class="bi bi-clock"></i> ${result.modified_time}
                        </div>
                    </div>
                    <div class="col-md-3 text-end">
                        <div class="btn-group-sm" role="group">
                            <button class="btn btn-outline-primary btn-sm" 
                                    onclick="previewFile('${result.filename}')" 
                                    title="预览数据">
                                <i class="bi bi-eye"></i>
                            </button>
                            <button class="btn btn-outline-info btn-sm" 
                                    onclick="showPivotModal('${result.filename}')" 
                                    title="数据透视">
                                <i class="bi bi-table"></i>
                            </button>
                            <button class="btn btn-outline-warning btn-sm" 
                                    onclick="showChartModal('${result.filename}')" 
                                    title="数据可视化">
                                <i class="bi bi-bar-chart"></i>
                            </button>
                            <button class="btn btn-outline-success btn-sm" 
                                    onclick="downloadFile('${result.filename}')" 
                                    title="下载">
                                <i class="bi bi-download"></i>
                            </button>
                            <button class="btn btn-outline-danger btn-sm" 
                                    onclick="deleteFile('${result.filename}')" 
                                    title="删除">
                                <i class="bi bi-trash"></i>
                            </button>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    `).join('');
    
    container.innerHTML = html;
}

// 切换文件选择状态
function toggleFileSelection(filename) {
    if (selectedFiles.has(filename)) {
        selectedFiles.delete(filename);
        document.querySelector(`[data-filename="${filename}"]`).classList.remove('selected');
    } else {
        selectedFiles.add(filename);
        document.querySelector(`[data-filename="${filename}"]`).classList.add('selected');
    }
    
    updateBatchActions();
}

// 更新批量操作面板
function updateBatchActions() {
    const batchActions = document.getElementById('batchActions');
    const selectedCount = document.getElementById('selectedCount');
    
    if (selectedFiles.size > 0) {
        batchActions.classList.add('show');
        selectedCount.textContent = `已选择 ${selectedFiles.size} 个文件`;
    } else {
        batchActions.classList.remove('show');
    }
}

// 全选文件
function selectAll() {
    selectedFiles.clear();
    currentResults.forEach(result => {
        selectedFiles.add(result.filename);
        document.querySelector(`[data-filename="${result.filename}"]`).classList.add('selected');
        const checkbox = document.getElementById(`check_${result.id}`);
        if (checkbox) checkbox.checked = true;
    });
    updateBatchActions();
}

// 清除选择
function clearSelection() {
    selectedFiles.clear();
    document.querySelectorAll('.result-card').forEach(card => {
        card.classList.remove('selected');
    });
    document.querySelectorAll('.form-check-input').forEach(checkbox => {
        checkbox.checked = false;
    });
    updateBatchActions();
}

// 刷新结果列表
function refreshResults() {
    clearSelection();
    loadResults();
}

// 预览文件数据
// 全局变量用于分页
let currentPage = 1;
let currentPageSize = 100;
let currentPreviewFilename = '';

async function previewFile(filename, page = 1, pageSize = 100) {
    currentPreviewFilename = filename;
    currentPage = page;
    currentPageSize = pageSize;
    
    document.getElementById('previewFileName').textContent = filename;
    document.getElementById('previewContent').innerHTML = `
        <div class="loading">
            <div class="spinner-border text-primary" role="status"></div>
            <div class="mt-2">正在加载数据...</div>
        </div>
    `;
    
    const modal = new bootstrap.Modal(document.getElementById('previewModal'));
    if (page === 1) modal.show(); // 只在第一页时显示模态框
    
    try {
        const response = await fetch(`/api/results/preview/${encodeURIComponent(filename)}?page=${page}&page_size=${pageSize}`);
        const data = await response.json();
        
        if (data.success) {
            currentPreviewData = data.data;
            renderPreviewTable(data.data);
        } else {
            document.getElementById('previewContent').innerHTML = `
                <div class="alert alert-danger">
                    <i class="bi bi-exclamation-triangle"></i> ${data.message}
                </div>
            `;
        }
    } catch (error) {
        console.error('预览失败:', error);
        document.getElementById('previewContent').innerHTML = `
            <div class="alert alert-danger">
                <i class="bi bi-exclamation-triangle"></i> 预览失败，请稍后重试
            </div>
        `;
    }
}

// 渲染预览表格
function renderPreviewTable(data) {
    if (!data || !data.columns || !data.data) {
        document.getElementById('previewContent').innerHTML = `
            <div class="alert alert-warning">
                <i class="bi bi-info-circle"></i> 文件中没有数据
            </div>
        `;
        return;
    }
    
    // 分页信息
    const pagination = data.total_pages > 1 ? `
        <div class="d-flex justify-content-between align-items-center mb-3">
            <div class="text-muted">
                显示第 ${((data.page - 1) * data.page_size) + 1} - ${Math.min(data.page * data.page_size, data.total_rows)} 行，共 ${data.total_rows} 行数据
            </div>
            <nav>
                <ul class="pagination pagination-sm mb-0">
                    <li class="page-item ${data.page <= 1 ? 'disabled' : ''}">
                        <a class="page-link" href="#" onclick="previewFile('${currentPreviewFilename}', ${data.page - 1}, ${data.page_size})">上一页</a>
                    </li>
                    <li class="page-item active">
                        <span class="page-link">${data.page} / ${data.total_pages}</span>
                    </li>
                    <li class="page-item ${data.page >= data.total_pages ? 'disabled' : ''}">
                        <a class="page-link" href="#" onclick="previewFile('${currentPreviewFilename}', ${data.page + 1}, ${data.page_size})">下一页</a>
                    </li>
                </ul>
            </nav>
        </div>
    ` : `
        <div class="text-muted mb-3">
            共 ${data.total_rows} 行数据
        </div>
    `;
    
    const table = `
        ${pagination}
        <div class="table-responsive data-preview-table">
            <table class="table table-striped table-hover">
                <thead class="table-dark">
                    <tr>
                        ${data.columns.map(col => `<th>${col}</th>`).join('')}
                    </tr>
                </thead>
                <tbody>
                    ${data.data.map(row => `
                        <tr>
                            ${row.map(cell => `<td>${cell || ''}</td>`).join('')}
                        </tr>
                    `).join('')}
                </tbody>
            </table>
        </div>
        ${data.total_pages > 1 ? `
            <div class="text-center">
                <small class="text-muted">
                    提示：预览模式分页显示，透视表和图表分析使用完整数据
                </small>
            </div>
        ` : ''}
    `;
    
    document.getElementById('previewContent').innerHTML = table;
}

// 显示数据透视模态框
async function showPivotModal(filename) {
    document.getElementById('pivotFileName').textContent = filename;
    
    // 加载完整数据用于透视分析
    try {
        const response = await fetch(`/api/results/preview/${encodeURIComponent(filename)}?show_all=true`);
        const data = await response.json();
        
        if (data.success) {
            currentPreviewData = data.data;
            renderPivotFields(data.data.columns);
            
            // 显示数据加载信息
            console.log(`透视分析：已加载 ${data.data.total_rows} 行完整数据`);
            
            const modal = new bootstrap.Modal(document.getElementById('pivotModal'));
            modal.show();
        } else {
            showError('无法加载数据用于透视分析：' + data.message);
        }
    } catch (error) {
        console.error('加载透视数据失败:', error);
        showError('加载透视数据失败，请稍后重试');
    }
}

// 渲染透视字段
function renderPivotFields(columns) {
    const fieldsContainer = document.getElementById('availableFields');
    
    const html = columns.map(column => `
        <div class="pivot-field" 
             draggable="true" 
             ondragstart="dragStart(event, '${column}')"
             data-field="${column}">
            ${column}
        </div>
    `).join('');
    
    fieldsContainer.innerHTML = html;
}

// 拖拽开始
function dragStart(event, fieldName) {
    event.dataTransfer.setData('text/plain', fieldName);
    event.target.classList.add('dragging');
}

// 允许拖放
function allowDrop(event) {
    event.preventDefault();
    event.currentTarget.classList.add('dragover');
}

// 拖放事件
function drop(event, dropType) {
    event.preventDefault();
    const fieldName = event.dataTransfer.getData('text/plain');
    const dropZone = event.currentTarget;
    
    dropZone.classList.remove('dragover');
    dropZone.classList.add('has-fields');
    
    // 如果是第一次添加字段，清空占位文本
    if (dropZone.querySelector('small.text-muted')) {
        dropZone.innerHTML = '';
    }
    
    // 创建字段标签 - 紧凑设计
    const fieldTag = document.createElement('div');
    fieldTag.className = 'pivot-field';
    fieldTag.innerHTML = `
        ${fieldName}
        <button class="btn btn-sm btn-outline-danger ms-1" onclick="removeField(this)" title="移除字段">&times;</button>
    `;
    fieldTag.dataset.field = fieldName;
    fieldTag.dataset.type = dropType;
    
    dropZone.appendChild(fieldTag);
    
    // 移除拖拽状态
    document.querySelectorAll('.pivot-field.dragging').forEach(el => {
        el.classList.remove('dragging');
    });
}

// 移除字段
function removeField(button) {
    const fieldTag = button.parentElement;
    const dropZone = fieldTag.parentElement;
    
    fieldTag.remove();
    
    // 检查是否还有字段，如果没有则显示占位文本
    if (dropZone.children.length === 0) {
        dropZone.classList.remove('has-fields');
        dropZone.innerHTML = '<small class="text-muted">拖拽字段到此处</small>';
    }
}

// 清除透视字段
function clearPivotFields() {
    const zones = [
        { id: 'rowFields', placeholder: '<small class="text-muted">拖拽字段到此处</small>' },
        { id: 'columnFields', placeholder: '<small class="text-muted">拖拽字段到此处</small>' },
        { id: 'valueFields', placeholder: '<small class="text-muted">拖拽字段到此处</small>' }
    ];
    
    zones.forEach(zone => {
        const element = document.getElementById(zone.id);
        element.innerHTML = zone.placeholder;
        element.classList.remove('has-fields');
    });
    
    // 重置排序状态
    window.currentPivotData = null;
    window.currentSortConfig = { column: null, direction: 'asc' };
    
    // 重置结果显示区域
    document.getElementById('pivotResult').innerHTML = `
        <div class="text-center py-5 text-muted">
            <i class="bi bi-table" style="font-size: 3rem; opacity: 0.3;"></i>
            <h5 class="mt-3">请配置字段并生成透视表</h5>
            <p>从左侧拖拽字段到对应区域，然后点击"生成透视表"</p>
        </div>
    `;
    
    document.getElementById('exportPivot').style.display = 'none';
}

// 生成透视表
function generatePivot() {
    const rowFields = Array.from(document.getElementById('rowFields').children)
        .filter(el => el.classList.contains('pivot-field'))
        .map(el => el.dataset.field);
    
    const columnFields = Array.from(document.getElementById('columnFields').children)
        .filter(el => el.classList.contains('pivot-field'))
        .map(el => el.dataset.field);
    
    const valueFields = Array.from(document.getElementById('valueFields').children)
        .filter(el => el.classList.contains('pivot-field'))
        .map(el => el.dataset.field);
    
    const aggregation = document.getElementById('aggregationMethod').value;
    
    if (rowFields.length === 0 && columnFields.length === 0) {
        showError('请至少选择一个行字段或列字段');
        return;
    }
    
    if (valueFields.length === 0) {
        showError('请选择至少一个值字段');
        return;
    }
    
    // 生成透视表（简化版实现）
    const pivotData = calculatePivot(currentPreviewData, rowFields, columnFields, valueFields, aggregation);
    renderPivotTable(pivotData);
    
    document.getElementById('exportPivot').style.display = 'inline-block';
}

// 计算透视表数据
function calculatePivot(data, rowFields, columnFields, valueFields, aggregation) {
    // 这是一个简化的透视表实现
    const result = {};
    const allColumns = new Set();
    
    data.data.forEach(row => {
        const rowKey = rowFields.map(field => {
            const index = data.columns.indexOf(field);
            return row[index] || '';
        }).join('|');
        
        const columnKey = columnFields.map(field => {
            const index = data.columns.indexOf(field);
            return row[index] || '';
        }).join('|');
        
        if (!result[rowKey]) {
            result[rowKey] = {};
        }
        
        allColumns.add(columnKey);
        
        valueFields.forEach(valueField => {
            const valueIndex = data.columns.indexOf(valueField);
            const value = parseFloat(row[valueIndex]) || 0;
            
            const key = columnKey + '|' + valueField;
            if (!result[rowKey][key]) {
                result[rowKey][key] = [];
            }
            result[rowKey][key].push(value);
        });
    });
    
    // 聚合计算
    Object.keys(result).forEach(rowKey => {
        Object.keys(result[rowKey]).forEach(key => {
            const values = result[rowKey][key];
            switch (aggregation) {
                case 'sum':
                    result[rowKey][key] = values.reduce((a, b) => a + b, 0);
                    break;
                case 'count':
                    result[rowKey][key] = values.length;
                    break;
                case 'avg':
                    result[rowKey][key] = values.reduce((a, b) => a + b, 0) / values.length;
                    break;
                case 'min':
                    result[rowKey][key] = Math.min(...values);
                    break;
                case 'max':
                    result[rowKey][key] = Math.max(...values);
                    break;
            }
        });
    });
    
    return { result, allColumns: Array.from(allColumns), rowFields, columnFields, valueFields };
}

// 渲染透视表
function renderPivotTable(pivotData) {
    const { result, allColumns, rowFields, columnFields, valueFields } = pivotData;
    
    // 将数据存储到全局变量以便排序使用
    window.currentPivotData = pivotData;
    window.currentSortConfig = { column: null, direction: 'asc' };
    
    // 计算汇总统计信息
    const totalRows = Object.keys(result).length;
    const totalColumns = allColumns.length * valueFields.length;
    const aggregationMethod = document.getElementById('aggregationMethod').value;
    const aggregationNames = {
        'sum': '求和',
        'count': '计数', 
        'avg': '平均值',
        'min': '最小值',
        'max': '最大值'
    };
    
    // 计算数据总计和统计信息
    let grandTotal = 0;
    let totalCells = 0;
    let minValue = Infinity;
    let maxValue = -Infinity;
    
    Object.keys(result).forEach(rowKey => {
        allColumns.forEach(col => {
            valueFields.forEach(valueField => {
                const key = col + '|' + valueField;
                const value = result[rowKey][key] || 0;
                if (typeof value === 'number' && !isNaN(value)) {
                    grandTotal += value;
                    totalCells++;
                    minValue = Math.min(minValue, value);
                    maxValue = Math.max(maxValue, value);
                }
            });
        });
    });
    
    const avgValue = totalCells > 0 ? grandTotal / totalCells : 0;
    
    // 优化的数据描述 - 更紧凑的卡片式设计
    let description = `
        <div class="row mb-3">
            <div class="col-12">
                <div class="card border-primary">
                    <div class="card-header bg-primary text-white py-2">
                        <h6 class="mb-0"><i class="bi bi-table me-2"></i>透视表数据分析</h6>
                    </div>
                    <div class="card-body py-2">
                        <div class="row g-2">
                            <div class="col-xl-2 col-lg-3 col-md-4 col-sm-6">
                                <div class="stat-item">
                                    <div class="stat-label">行字段</div>
                                    <div class="stat-value text-primary">${rowFields.join(', ') || '无'}</div>
                                </div>
                            </div>
                            <div class="col-xl-2 col-lg-3 col-md-4 col-sm-6">
                                <div class="stat-item">
                                    <div class="stat-label">列字段</div>
                                    <div class="stat-value text-info">${columnFields.join(', ') || '无'}</div>
                                </div>
                            </div>
                            <div class="col-xl-2 col-lg-3 col-md-4 col-sm-6">
                                <div class="stat-item">
                                    <div class="stat-label">值字段</div>
                                    <div class="stat-value text-success">${valueFields.join(', ')}</div>
                                </div>
                            </div>
                            <div class="col-xl-2 col-lg-3 col-md-4 col-sm-6">
                                <div class="stat-item">
                                    <div class="stat-label">聚合方式</div>
                                    <div class="stat-value text-warning">${aggregationNames[aggregationMethod]}</div>
                                </div>
                            </div>
                            <div class="col-xl-2 col-lg-3 col-md-4 col-sm-6">
                                <div class="stat-item">
                                    <div class="stat-label">数据维度</div>
                                    <div class="stat-value">${totalRows} × ${totalColumns}</div>
                                </div>
                            </div>
                            <div class="col-xl-2 col-lg-3 col-md-4 col-sm-6">
                                <div class="stat-item">
                                    <div class="stat-label">数据量</div>
                                    <div class="stat-value">${totalCells.toLocaleString()} 个</div>
                                </div>
                            </div>
                        </div>
                        <div class="row g-2 mt-1">
                            <div class="col-xl-2 col-lg-3 col-md-4 col-sm-6">
                                <div class="stat-item">
                                    <div class="stat-label">总计</div>
                                    <div class="stat-value text-primary fw-bold">${grandTotal.toLocaleString()}</div>
                                </div>
                            </div>
                            <div class="col-xl-2 col-lg-3 col-md-4 col-sm-6">
                                <div class="stat-item">
                                    <div class="stat-label">平均值</div>
                                    <div class="stat-value">${avgValue.toFixed(2)}</div>
                                </div>
                            </div>
                            <div class="col-xl-2 col-lg-3 col-md-4 col-sm-6">
                                <div class="stat-item">
                                    <div class="stat-label">最小值</div>
                                    <div class="stat-value text-success">${minValue === Infinity ? '无' : minValue.toLocaleString()}</div>
                                </div>
                            </div>
                            <div class="col-xl-2 col-lg-3 col-md-4 col-sm-6">
                                <div class="stat-item">
                                    <div class="stat-label">最大值</div>
                                    <div class="stat-value text-danger">${maxValue === -Infinity ? '无' : maxValue.toLocaleString()}</div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    `;
    
    // 创建响应式表格容器
    let html = description + `
        <div class="pivot-table-container">
            <div class="table-responsive pivot-table-scroll">
                <table class="table table-striped table-hover table-sm pivot-table" id="pivotTable">
    `;
    
    // 表头 - 使用主题色和紧凑设计，添加排序功能
    html += '<thead class="table-primary sticky-top">';
    
    // 主表头行
    html += '<tr>';
    rowFields.forEach((field, index) => {
        const sortIcon = getSortIcon('row', index);
        html += `<th class="pivot-header row-header sortable" onclick="sortPivotTable('row', ${index})" title="点击排序">
            ${field} ${sortIcon}
        </th>`;
    });
    
    // 为列字段创建分组表头
    if (columnFields.length > 0) {
        allColumns.forEach((col, colIndex) => {
            html += `<th colspan="${valueFields.length}" class="pivot-header column-group-header text-center">${col || '空值'}</th>`;
        });
    } else {
        valueFields.forEach((valueField, valueIndex) => {
            const sortIcon = getSortIcon('value', valueIndex);
            html += `<th class="pivot-header value-header sortable" onclick="sortPivotTable('value', ${valueIndex})" title="点击排序">
                ${valueField} ${sortIcon}
            </th>`;
        });
    }
    html += '</tr>';
    
    // 如果有列字段，添加子表头行显示值字段
    if (columnFields.length > 0) {
        html += '<tr>';
        rowFields.forEach(() => {
            html += '<th class="pivot-header row-header-spacer"></th>';
        });
        allColumns.forEach((col, colIndex) => {
            valueFields.forEach((valueField, valueIndex) => {
                const sortIcon = getSortIcon('value', `${colIndex}_${valueIndex}`);
                html += `<th class="pivot-header value-subheader sortable" onclick="sortPivotTable('value', '${colIndex}_${valueIndex}')" title="点击排序">
                    ${valueField} ${sortIcon}
                </th>`;
            });
        });
        html += '</tr>';
    }
    
    html += '</thead>';
    
    // 表体 - 优化数据显示
    html += '<tbody id="pivotTableBody">';
    html += renderPivotTableBody(result, allColumns, rowFields, valueFields, aggregationMethod, avgValue, maxValue);
    html += '</tbody></table></div></div>';
    
    document.getElementById('pivotResult').innerHTML = html;
}

// 渲染透视表表体
function renderPivotTableBody(result, allColumns, rowFields, valueFields, aggregationMethod, avgValue, maxValue) {
    let html = '';
    Object.keys(result).forEach((rowKey, index) => {
        const rowClass = index % 2 === 0 ? '' : 'table-light';
        html += `<tr class="${rowClass}">`;
        
        // 行字段数据
        const rowValues = rowKey.split('|');
        rowValues.forEach((value, idx) => {
            html += `<td class="pivot-cell row-cell">${value}</td>`;
        });
        
        // 数值数据
        allColumns.forEach(col => {
            valueFields.forEach(valueField => {
                const key = col + '|' + valueField;
                const value = result[rowKey][key] || 0;
                let displayValue;
                let cellClass = 'pivot-cell data-cell text-end';
                
                if (typeof value === 'number') {
                    if (aggregationMethod === 'count') {
                        displayValue = value.toLocaleString();
                    } else {
                        displayValue = value.toFixed(2);
                    }
                    
                    // 根据数值大小添加颜色提示
                    if (value > 0) {
                        if (value >= maxValue * 0.8) {
                            cellClass += ' high-value';
                        } else if (value >= avgValue) {
                            cellClass += ' medium-value';
                        } else {
                            cellClass += ' low-value';
                        }
                    }
                } else {
                    displayValue = value;
                }
                
                html += `<td class="${cellClass}" title="${displayValue}">${displayValue}</td>`;
            });
        });
        html += '</tr>';
    });
    return html;
}

// 获取排序图标
function getSortIcon(type, index) {
    if (!window.currentSortConfig) return '<i class="bi bi-arrow-down-up sort-icon"></i>';
    
    const currentSort = window.currentSortConfig;
    const columnKey = `${type}_${index}`;
    
    if (currentSort.column === columnKey) {
        if (currentSort.direction === 'asc') {
            return '<i class="bi bi-sort-up sort-icon active"></i>';
        } else {
            return '<i class="bi bi-sort-down sort-icon active"></i>';
        }
    }
    return '<i class="bi bi-arrow-down-up sort-icon"></i>';
}

// 透视表排序功能
function sortPivotTable(type, index) {
    if (!window.currentPivotData) return;
    
    const columnKey = `${type}_${index}`;
    let direction = 'asc';
    
    // 如果点击的是同一列，切换排序方向
    if (window.currentSortConfig.column === columnKey) {
        direction = window.currentSortConfig.direction === 'asc' ? 'desc' : 'asc';
    }
    
    window.currentSortConfig = { column: columnKey, direction };
    
    const { result, allColumns, rowFields, columnFields, valueFields } = window.currentPivotData;
    const aggregationMethod = document.getElementById('aggregationMethod').value;
    
    // 将数据转换为数组以便排序
    const dataArray = Object.keys(result).map(rowKey => {
        const rowData = { rowKey, rowValues: rowKey.split('|'), data: result[rowKey] };
        return rowData;
    });
    
    // 执行排序
    dataArray.sort((a, b) => {
        let valueA, valueB;
        
        if (type === 'row') {
            // 按行字段排序
            valueA = a.rowValues[index] || '';
            valueB = b.rowValues[index] || '';
            
            // 尝试数值比较
            const numA = parseFloat(valueA);
            const numB = parseFloat(valueB);
            
            if (!isNaN(numA) && !isNaN(numB)) {
                valueA = numA;
                valueB = numB;
            } else {
                valueA = valueA.toString().toLowerCase();
                valueB = valueB.toString().toLowerCase();
            }
        } else if (type === 'value') {
            // 按值字段排序
            if (typeof index === 'string' && index.includes('_')) {
                // 复合列（有列字段的情况）
                const [colIndex, valueIndex] = index.split('_');
                const col = allColumns[colIndex];
                const valueField = valueFields[valueIndex];
                const key = col + '|' + valueField;
                
                valueA = a.data[key] || 0;
                valueB = b.data[key] || 0;
            } else {
                // 简单值字段
                const valueField = valueFields[index];
                const firstCol = allColumns[0] || '';
                const key = firstCol + '|' + valueField;
                
                valueA = a.data[key] || 0;
                valueB = b.data[key] || 0;
            }
        }
        
        if (typeof valueA === 'number' && typeof valueB === 'number') {
            return direction === 'asc' ? valueA - valueB : valueB - valueA;
        } else {
            if (valueA < valueB) return direction === 'asc' ? -1 : 1;
            if (valueA > valueB) return direction === 'asc' ? 1 : -1;
            return 0;
        }
    });
    
    // 重建结果对象
    const sortedResult = {};
    dataArray.forEach(item => {
        sortedResult[item.rowKey] = item.data;
    });
    
    // 计算统计值用于重新渲染
    let grandTotal = 0;
    let totalCells = 0;
    let minValue = Infinity;
    let maxValue = -Infinity;
    
    Object.keys(sortedResult).forEach(rowKey => {
        allColumns.forEach(col => {
            valueFields.forEach(valueField => {
                const key = col + '|' + valueField;
                const value = sortedResult[rowKey][key] || 0;
                if (typeof value === 'number' && !isNaN(value)) {
                    grandTotal += value;
                    totalCells++;
                    minValue = Math.min(minValue, value);
                    maxValue = Math.max(maxValue, value);
                }
            });
        });
    });
    
    const avgValue = totalCells > 0 ? grandTotal / totalCells : 0;
    
    // 只更新表体和表头的排序图标
    const tableBody = document.getElementById('pivotTableBody');
    if (tableBody) {
        tableBody.innerHTML = renderPivotTableBody(sortedResult, allColumns, rowFields, valueFields, aggregationMethod, avgValue, maxValue);
    }
    
    // 更新表头的排序图标
    updateSortIcons();
}

// 更新排序图标
function updateSortIcons() {
    document.querySelectorAll('.sortable').forEach((header, index) => {
        const onclick = header.getAttribute('onclick');
        if (onclick) {
            // 提取排序参数
            const match = onclick.match(/sortPivotTable\('(\w+)', (.+?)\)/);
            if (match) {
                const type = match[1];
                const idx = match[2].replace(/'/g, '');
                const icon = getSortIcon(type, idx);
                
                // 更新图标
                const existingIcon = header.querySelector('.sort-icon');
                if (existingIcon) {
                    const iconElement = document.createElement('div');
                    iconElement.innerHTML = icon;
                    header.replaceChild(iconElement.firstChild, existingIcon);
                }
            }
        }
    });
}

// 显示图表模态框
async function showChartModal(filename) {
    document.getElementById('chartFileName').textContent = filename;
    
    // 加载完整数据用于可视化分析
    try {
        const response = await fetch(`/api/results/preview/${encodeURIComponent(filename)}?show_all=true`);
        const data = await response.json();
        
        if (data.success) {
            currentPreviewData = data.data;
            renderChartFields(data.data.columns);
            
            // 显示数据加载信息
            console.log(`可视化分析：已加载 ${data.data.total_rows} 行完整数据`);
            
            const modal = new bootstrap.Modal(document.getElementById('chartModal'));
            modal.show();
        } else {
            showError('无法加载数据用于可视化：' + data.message);
        }
    } catch (error) {
        console.error('加载可视化数据失败:', error);
        showError('加载可视化数据失败，请稍后重试');
    }
}

// 渲染图表字段
function renderChartFields(columns) {
    const fieldsContainer = document.getElementById('chartAvailableFields');
    
    const html = columns.map(column => `
        <div class="pivot-field" 
             draggable="true" 
             ondragstart="dragStart(event, '${column}')"
             data-field="${column}">
            ${column}
        </div>
    `).join('');
    
    fieldsContainer.innerHTML = html;
}

// 生成图表
async function generateChart() {
    const xAxisFields = Array.from(document.getElementById('xAxisFields').children)
        .filter(el => el.classList.contains('pivot-field'))
        .map(el => el.dataset.field);
    
    const yAxisFields = Array.from(document.getElementById('yAxisFields').children)
        .filter(el => el.classList.contains('pivot-field'))
        .map(el => el.dataset.field);
    
    const chartType = document.getElementById('chartType').value;
    const aggregationMethod = document.getElementById('chartAggregationMethod').value;
    
    if (xAxisFields.length === 0 || yAxisFields.length === 0) {
        showError('请选择X轴和Y轴字段');
        return;
    }
    
    // 验证排序设置
    if (!validateSortSettings()) {
        return;
    }
    
    // 获取排序设置
    const sortSettings = getCurrentSortSettings();
    
    // 使用后端API生成图表数据（支持排序）
    try {
        const currentFileName = document.getElementById('chartFileName').textContent;
        
        const response = await fetch('/api/results/chart-data', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                filename: currentFileName,
                x_field: xAxisFields[0],
                y_field: yAxisFields[0],
                chart_type: chartType,
                aggregation: aggregationMethod,
                sort_field: sortSettings.sortField,
                sort_direction: sortSettings.sortDirection,
                sort_type: sortSettings.sortType,
                custom_sort_order: sortSettings.customSortOrder
            })
        });
        
        const data = await response.json();
        
        if (data.success) {
            renderChart(data.chart_data, chartType);
            document.getElementById('exportChart').style.display = 'inline-block';
        } else {
            showError('生成图表失败：' + data.message);
        }
    } catch (error) {
        console.error('生成图表失败:', error);
        showError('生成图表失败，请稍后重试');
    }
}

// 生成图表数据
function generateChartData(data, xField, yField, chartType, aggregationMethod) {
    const xIndex = data.columns.indexOf(xField);
    const yIndex = data.columns.indexOf(yField);
    
    const labels = [];
    const values = [];
    const dataMap = {};
    
    data.data.forEach(row => {
        const xValue = row[xIndex] || '';
        const yValue = parseFloat(row[yIndex]) || 0;
        
        if (dataMap[xValue]) {
            dataMap[xValue].push(yValue);
        } else {
            dataMap[xValue] = [yValue];
        }
    });
    
    Object.keys(dataMap).forEach(key => {
        labels.push(key);
        const valuesArray = dataMap[key];
        let aggregatedValue;
        
        // 根据聚合方式计算值
        switch (aggregationMethod) {
            case 'sum':
                aggregatedValue = valuesArray.reduce((a, b) => a + b, 0);
                break;
            case 'count':
                aggregatedValue = valuesArray.length;
                break;
            case 'avg':
                aggregatedValue = valuesArray.reduce((a, b) => a + b, 0) / valuesArray.length;
                break;
            case 'min':
                aggregatedValue = Math.min(...valuesArray);
                break;
            case 'max':
                aggregatedValue = Math.max(...valuesArray);
                break;
            default:
                aggregatedValue = valuesArray.reduce((a, b) => a + b, 0) / valuesArray.length;
        }
        
        values.push(aggregatedValue);
    });
    
    return { labels, values, xField, yField, aggregationMethod };
}

// 渲染图表
function renderChart(chartData, chartType) {
    const ctx = document.createElement('canvas');
    document.getElementById('chartContainer').innerHTML = '';
    document.getElementById('chartContainer').appendChild(ctx);
    
    if (visualChart) {
        visualChart.destroy();
    }
    
    // 如果是从后端API获取的数据，使用原有的Chart.js格式
    if (chartData.labels && chartData.datasets) {
        const config = {
            type: chartType,
            data: chartData,
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    title: {
                        display: true,
                        text: chartData.datasets[0].label,
                        font: {
                            size: 16,
                            weight: 'bold'
                        }
                    },
                    legend: {
                        display: chartType === 'pie'
                    }
                },
                scales: chartType !== 'pie' ? {
                    x: {
                        title: {
                            display: true,
                            text: 'X轴'
                        }
                    },
                    y: {
                        beginAtZero: true,
                        title: {
                            display: true,
                            text: 'Y轴'
                        }
                    }
                } : {}
            }
        };
        
        // 为饼图设置颜色
        if (chartType === 'pie') {
            config.data.datasets[0].backgroundColor = generateColors(chartData.labels.length);
            config.data.datasets[0].borderColor = generateColors(chartData.labels.length);
        }
        
        visualChart = new Chart(ctx, config);
        return;
    }
    
    // 兼容旧版本的本地生成数据格式
    const aggregationNames = {
        'sum': '求和',
        'count': '计数',
        'avg': '平均值',
        'min': '最小值',
        'max': '最大值'
    };
    
    const aggregationLabel = aggregationNames[chartData.aggregationMethod] || '平均值';
    const yAxisLabel = `${chartData.yField} (${aggregationLabel})`;
    
    const config = {
        type: chartType,
        data: {
            labels: chartData.labels,
            datasets: [{
                label: yAxisLabel,
                data: chartData.values,
                backgroundColor: chartType === 'pie' ? 
                    generateColors(chartData.values.length) : 
                    'rgba(54, 162, 235, 0.2)',
                borderColor: chartType === 'pie' ? 
                    generateColors(chartData.values.length) : 
                    'rgba(54, 162, 235, 1)',
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                title: {
                    display: true,
                    text: `${chartData.xField} vs ${yAxisLabel}`,
                    font: {
                        size: 16,
                        weight: 'bold'
                    }
                },
                legend: {
                    display: chartType === 'pie'
                }
            },
            scales: chartType !== 'pie' ? {
                x: {
                    title: {
                        display: true,
                        text: chartData.xField
                    }
                },
                y: {
                    beginAtZero: true,
                    title: {
                        display: true,
                        text: yAxisLabel
                    }
                }
            } : {}
        }
    };
    
    visualChart = new Chart(ctx, config);
}

// 生成颜色数组
function generateColors(count) {
    const colors = [];
    for (let i = 0; i < count; i++) {
        const hue = (i * 137.508) % 360; // 黄金角度分布
        colors.push(`hsl(${hue}, 70%, 60%)`);
    }
    return colors;
}

// 单个文件下载
function downloadFile(filename) {
    window.location.href = `/download/${encodeURIComponent(filename)}`;
}

// 批量下载
function batchDownload() {
    if (selectedFiles.size === 0) {
        showError('请先选择要下载的文件');
        return;
    }
    
    // 逐个下载选中的文件
    selectedFiles.forEach(filename => {
        downloadFile(filename);
    });
}

// 单个文件删除
async function deleteFile(filename) {
    if (!confirm(`确定要删除文件 "${filename}" 吗？`)) {
        return;
    }
    
    try {
        const response = await fetch(`/api/results/delete`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ filenames: [filename] })
        });
        
        const data = await response.json();
        
        if (data.success) {
            showSuccess(`成功删除文件: ${filename}`);
            loadResults();
        } else {
            showError('删除失败: ' + data.message);
        }
    } catch (error) {
        console.error('删除文件失败:', error);
        showError('删除文件失败，请稍后重试');
    }
}

// 批量删除
async function batchDelete() {
    if (selectedFiles.size === 0) {
        showError('请先选择要删除的文件');
        return;
    }
    
    if (!confirm(`确定要删除选中的 ${selectedFiles.size} 个文件吗？`)) {
        return;
    }
    
    try {
        const response = await fetch(`/api/results/delete`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ filenames: Array.from(selectedFiles) })
        });
        
        const data = await response.json();
        
        if (data.success) {
            showSuccess(`成功删除 ${data.deleted.length} 个文件`);
            if (data.failed.length > 0) {
                showError(`部分文件删除失败: ${data.failed.join(', ')}`);
            }
            clearSelection();
            loadResults();
        } else {
            showError('批量删除失败: ' + data.message);
        }
    } catch (error) {
        console.error('批量删除失败:', error);
        showError('批量删除失败，请稍后重试');
    }
}

// 导出透视表
function exportPivotData() {
    if (!window.currentPivotData) {
        showError('请先生成透视表');
        return;
    }

    // 显示导出选项模态框
    showExportPivotModal();
}

// 显示导出透视表模态框
function showExportPivotModal() {
    // 创建模态框HTML（如果不存在）
    if (!document.getElementById('exportPivotModal')) {
        const modalHTML = `
            <div class="modal fade" id="exportPivotModal" tabindex="-1">
                <div class="modal-dialog">
                    <div class="modal-content">
                        <div class="modal-header">
                            <h5 class="modal-title">
                                <i class="bi bi-download me-2"></i>导出透视表
                            </h5>
                            <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
                        </div>
                        <div class="modal-body">
                            <div class="mb-3">
                                <label class="form-label">导出格式</label>
                                <select class="form-select" id="exportPivotFormat">
                                    <option value="xlsx">Excel文件 (.xlsx) - 推荐</option>
                                    <option value="csv">CSV文件 (.csv)</option>
                                </select>
                                <div class="form-text">
                                    Excel格式将包含透视表、统计信息和原始数据样本；CSV格式仅包含透视表数据
                                </div>
                            </div>
                            <div class="alert alert-info">
                                <i class="bi bi-info-circle me-2"></i>
                                <strong>导出内容：</strong>
                                <ul class="mb-0 mt-2">
                                    <li>透视表数据</li>
                                    <li>字段配置和统计信息</li>
                                    <li>原始数据样本（Excel格式）</li>
                                </ul>
                            </div>
                        </div>
                        <div class="modal-footer">
                            <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">取消</button>
                            <button type="button" class="btn btn-primary" onclick="performPivotExport()">
                                <i class="bi bi-download me-1"></i>开始导出
                            </button>
                        </div>
                    </div>
                </div>
            </div>
        `;
        document.body.insertAdjacentHTML('beforeend', modalHTML);
    }

    // 显示模态框
    const modal = new bootstrap.Modal(document.getElementById('exportPivotModal'));
    modal.show();
}

// 执行透视表导出
async function performPivotExport() {
    const format = document.getElementById('exportPivotFormat').value;
    const filename = document.getElementById('pivotFileName').textContent;
    
    if (!filename) {
        showError('无法获取文件名');
        return;
    }

    // 获取透视表配置
    const rowFields = Array.from(document.getElementById('rowFields').children)
        .filter(el => el.classList.contains('pivot-field'))
        .map(el => el.dataset.field);
    
    const columnFields = Array.from(document.getElementById('columnFields').children)
        .filter(el => el.classList.contains('pivot-field'))
        .map(el => el.dataset.field);
    
    const valueFields = Array.from(document.getElementById('valueFields').children)
        .filter(el => el.classList.contains('pivot-field'))
        .map(el => el.dataset.field);
    
    const aggregation = document.getElementById('aggregationMethod').value;

    // 显示加载状态
    const exportBtn = document.querySelector('#exportPivotModal .btn-primary');
    const originalText = exportBtn.innerHTML;
    exportBtn.innerHTML = '<span class="spinner-border spinner-border-sm me-1"></span>导出中...';
    exportBtn.disabled = true;

    try {
        const response = await fetch('/api/results/pivot/export', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                filename: filename,
                row_fields: rowFields,
                column_fields: columnFields,
                value_fields: valueFields,
                aggregation: aggregation,
                format: format
            })
        });

        const result = await response.json();

        if (result.success) {
            // 隐藏模态框
            bootstrap.Modal.getInstance(document.getElementById('exportPivotModal')).hide();
            
            // 显示成功消息
            showSuccess(`透视表导出成功！文件：${result.filename}`);
            
            // 自动下载文件
            if (result.download_url) {
                const link = document.createElement('a');
                link.href = result.download_url;
                link.download = result.filename;
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
            }
            
            // 刷新结果列表
            setTimeout(() => {
                loadResults();
            }, 1000);
            
        } else {
            showError('导出失败：' + result.message);
        }
    } catch (error) {
        console.error('导出透视表失败:', error);
        showError('导出失败，请稍后重试');
    } finally {
        // 恢复按钮状态
        exportBtn.innerHTML = originalText;
        exportBtn.disabled = false;
    }
}

// 导出图表
function exportChart() {
    if (!visualChart) {
        showError('请先生成图表');
        return;
    }

    try {
        // 获取当前图表配置信息
        const chartTitle = document.getElementById('visualFileName').textContent;
        const chartType = document.getElementById('chartType').value;
        const timestamp = new Date().toISOString().slice(0, 19).replace(/:/g, '-');
        const filename = `${chartTitle}_${chartType}_${timestamp}.png`;

        // 导出为高质量PNG图片
        const link = document.createElement('a');
        link.download = filename;
        link.href = visualChart.toBase64Image('image/png', 1.0); // 最高质量
        
        // 自动下载
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        
        showSuccess(`图表导出成功：${filename}`);
        
    } catch (error) {
        console.error('导出图表失败:', error);
        showError('图表导出失败，请稍后重试');
    }
}

// 显示成功消息
function showSuccess(message) {
    showToast(message, 'success');
}

// 显示错误消息
function showError(message) {
    showToast(message, 'danger');
}

// 显示信息
function showInfo(message) {
    showToast(message, 'info');
}

// 通用Toast显示函数
function showToast(message, type = 'info') {
    // 创建toast容器（如果不存在）
    let toastContainer = document.getElementById('toastContainer');
    if (!toastContainer) {
        toastContainer = document.createElement('div');
        toastContainer.id = 'toastContainer';
        toastContainer.className = 'toast-container position-fixed top-0 end-0 p-3';
        toastContainer.style.zIndex = '9999';
        document.body.appendChild(toastContainer);
    }

    // 创建toast元素
    const toastId = 'toast-' + Date.now();
    const iconMap = {
        'success': 'bi-check-circle-fill',
        'danger': 'bi-exclamation-triangle-fill', 
        'info': 'bi-info-circle-fill',
        'warning': 'bi-exclamation-triangle-fill'
    };
    
    const toastHTML = `
        <div id="${toastId}" class="toast" role="alert">
            <div class="toast-header">
                <i class="bi ${iconMap[type]} me-2 text-${type}"></i>
                <strong class="me-auto">系统消息</strong>
                <button type="button" class="btn-close" data-bs-dismiss="toast"></button>
            </div>
            <div class="toast-body">
                ${message}
            </div>
        </div>
    `;
    
    toastContainer.insertAdjacentHTML('beforeend', toastHTML);
    
    // 显示toast
    const toastElement = document.getElementById(toastId);
    const toast = new bootstrap.Toast(toastElement, {
        autohide: true,
        delay: type === 'danger' ? 5000 : 3000  // 错误消息显示更久
    });
    
    toast.show();
    
    // 自动清理DOM元素
    toastElement.addEventListener('hidden.bs.toast', function() {
        toastElement.remove();
    });
}

// 处理拖拽离开事件
document.addEventListener('dragleave', function(e) {
    if (e.target.classList.contains('drop-zone')) {
        e.target.classList.remove('dragover');
    }
});

// 排序设置相关函数

// 切换自定义排序输入框的显示/隐藏
function toggleCustomSortInput() {
    const sortType = document.getElementById('sortType').value;
    const customSortContainer = document.getElementById('customSortContainer');
    
    if (sortType === 'custom') {
        customSortContainer.style.display = 'block';
    } else {
        customSortContainer.style.display = 'none';
    }
}

// 清除排序设置
function clearSortSettings() {
    document.getElementById('sortField').value = 'none';
    document.getElementById('sortDirection').value = 'asc';
    document.getElementById('sortType').value = 'alphabetic';
    document.getElementById('customSortOrder').value = '';
    toggleCustomSortInput();
}

// 获取当前排序设置
function getCurrentSortSettings() {
    const sortField = document.getElementById('sortField').value;
    const sortDirection = document.getElementById('sortDirection').value;
    const sortType = document.getElementById('sortType').value;
    
    let customSortOrder = [];
    if (sortType === 'custom') {
        const customSortText = document.getElementById('customSortOrder').value.trim();
        if (customSortText) {
            customSortOrder = customSortText.split('\n').map(item => item.trim()).filter(item => item);
        }
    }
    
    return {
        sortField,
        sortDirection,
        sortType,
        customSortOrder
    };
}

// 验证排序设置
function validateSortSettings() {
    const sortSettings = getCurrentSortSettings();
    
    if (sortSettings.sortType === 'custom' && sortSettings.sortField !== 'none') {
        if (sortSettings.customSortOrder.length === 0) {
            showError('使用自定义排序时，请输入自定义排序顺序');
            return false;
        }
    }
    
    return true;
}

// 重置图表字段（包括排序设置）
function clearChartFields() {
    ['xAxisFields', 'yAxisFields'].forEach(id => {
        const zone = document.getElementById(id);
        zone.innerHTML = id === 'xAxisFields' ? '拖拽字段到此处作为X轴' : '拖拽字段到此处作为Y轴';
        zone.classList.remove('has-fields');
    });
    
    // 重置排序设置
    clearSortSettings();
    
    if (visualChart) {
        visualChart.destroy();
        visualChart = null;
    }
    
    document.getElementById('chartContainer').innerHTML = '';
    document.getElementById('exportChart').style.display = 'none';
}
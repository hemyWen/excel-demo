// 源表行计数器
let sourceRowCount = 3;

// 选择文件
function selectFile(id) {
    const elementId = id === 'b' ? 'fileB' : `file${id}`;
    document.getElementById(elementId).click();
}

// 文件选择后更新显示
function fileSelected(id) {
    const elementId = id === 'b' ? 'fileB' : `file${id}`;
    const fileNameId = id === 'b' ? 'fileNameB' : `fileName${id}`;
    const fileInput = document.getElementById(elementId);
    const fileNameSpan = document.getElementById(fileNameId);
    
    if (fileInput.files && fileInput.files[0]) {
        fileNameSpan.textContent = fileInput.files[0].name;
        fileNameSpan.classList.add('has-file');
    } else {
        fileNameSpan.textContent = '';
        fileNameSpan.classList.remove('has-file');
    }
}

// 添加源表行
function addSourceRow() {
    if (sourceRowCount >= 10) {
        alert('最多只能添加10个源表');
        return;
    }
    
    sourceRowCount++;
    const tbody = document.querySelector('#sourceTable tbody');
    
    const newRow = document.createElement('tr');
    newRow.id = `sourceRow${sourceRowCount}`;
    newRow.innerHTML = `
        <td>
            <div class="file-upload-cell">
                <span class="delete-btn" onclick="deleteSourceRow(${sourceRowCount})">×</span>
                <button type="button" class="file-btn" onclick="selectFile(${sourceRowCount})">选择文件</button>
                <input type="file" id="file${sourceRowCount}" name="file_${sourceRowCount}" class="file-input" accept=".xlsx,.xls" onchange="fileSelected(${sourceRowCount})">
                <span class="file-name" id="fileName${sourceRowCount}"></span>
            </div>
        </td>
        <td><input type="text" id="nameCol${sourceRowCount}" name="name_col_${sourceRowCount}" class="column-input" placeholder="请输入"></td>
        <td><input type="text" id="idCol${sourceRowCount}" name="id_col_${sourceRowCount}" class="column-input" placeholder="请输入"></td>
        <td><input type="text" id="salaryCol${sourceRowCount}" name="salary_col_${sourceRowCount}" class="column-input" placeholder="请输入"></td>
    `;
    
    tbody.appendChild(newRow);
}

// 删除源表行
function deleteSourceRow(rowId) {
    const row = document.getElementById(`sourceRow${rowId}`);
    if (row) {
        row.remove();
    }
}

// 验证表单
function validateForm() {
    // 验证目标表
    const targetFileInput = document.getElementById('fileB');
    const targetNameCol = document.getElementById('nameColB').value.trim();
    const targetIdCol = document.getElementById('idColB').value.trim();
    const targetSalaryCol = document.getElementById('salaryColB').value.trim();
    
    if (!targetFileInput.files || !targetFileInput.files[0]) {
        alert('请选择目标表文件');
        return false;
    }
    if (!targetNameCol) {
        alert('请填写目标表的姓名列名');
        document.getElementById('nameColB').focus();
        return false;
    }
    if (!targetIdCol) {
        alert('请填写目标表的身份证列名');
        document.getElementById('idColB').focus();
        return false;
    }
    if (!targetSalaryCol) {
        alert('请填写目标表的工资列名');
        document.getElementById('salaryColB').focus();
        return false;
    }
    
    // 验证源数据表
    const sourceRows = document.querySelectorAll('#sourceTable tbody tr');
    let hasSourceData = false;
    
    for (let row of sourceRows) {
        const rowId = row.id.replace('sourceRow', '');
        const fileInput = document.getElementById(`file${rowId}`);
        const nameCol = document.getElementById(`nameCol${rowId}`).value.trim();
        const idCol = document.getElementById(`idCol${rowId}`).value.trim();
        const salaryCol = document.getElementById(`salaryCol${rowId}`).value.trim();
        
        // 如果该行选择了文件，则必须填写所有列
        if (fileInput && fileInput.files && fileInput.files[0]) {
            hasSourceData = true;
            if (!nameCol) {
                alert(`请填写第 ${rowId} 个源数据表的姓名列名`);
                document.getElementById(`nameCol${rowId}`).focus();
                return false;
            }
            if (!idCol) {
                alert(`请填写第 ${rowId} 个源数据表的身份证列名`);
                document.getElementById(`idCol${rowId}`).focus();
                return false;
            }
            if (!salaryCol) {
                alert(`请填写第 ${rowId} 个源数据表的工资列名`);
                document.getElementById(`salaryCol${rowId}`).focus();
                return false;
            }
        }
    }
    
    if (!hasSourceData) {
        alert('请至少上传一个源数据表文件');
        return false;
    }
    
    return true;
}

// 表单提交
document.addEventListener('DOMContentLoaded', function() {
    const uploadForm = document.getElementById('uploadForm');
    const submitBtn = document.getElementById('submitBtn');
    const resultArea = document.getElementById('resultArea');
    const errorArea = document.getElementById('errorArea');
    const downloadLink = document.getElementById('downloadLink');
    const errorText = document.getElementById('errorText');

    uploadForm.addEventListener('submit', async function(e) {
        e.preventDefault();
    
        // 验证必填字段
        if (!validateForm()) {
            return;
        }
    
        // 隐藏之前的结果
        resultArea.hidden = true;
        errorArea.hidden = true;
    
        // 显示加载状态
        submitBtn.disabled = true;
        submitBtn.textContent = '处理中...';

        // 创建FormData
        const formData = new FormData();
        
        // 添加源表文件
        for (let i = 1; i <= sourceRowCount; i++) {
            const fileInput = document.getElementById(`file${i}`);
            if (fileInput && fileInput.files[0]) {
                formData.append(`file_${i}`, fileInput.files[0]);
                formData.append(`id_col_${i}`, document.getElementById(`idCol${i}`).value || 'A');
                formData.append(`salary_col_${i}`, document.getElementById(`salaryCol${i}`).value || 'B');
                formData.append(`name_col_${i}`, document.getElementById(`nameCol${i}`).value || '');
            }
        }
        
        // 添加B表
        formData.append('file_b', fileB.files[0]);
        formData.append('id_col_b', document.getElementById('idColB').value || 'A');
        formData.append('salary_col_b', document.getElementById('salaryColB').value || 'B');
        formData.append('name_col_b', document.getElementById('nameColB').value || '');

        try {
            const response = await fetch('/process', {
                method: 'POST',
                body: formData
            });

            const data = await response.json();

            if (response.ok && data.success) {
                showResult(data.download_url);
            } else {
                showError(data.error || '处理失败，请重试');
            }
        } catch (error) {
            showError('网络错误，请检查网络连接');
        } finally {
            submitBtn.disabled = false;
            submitBtn.textContent = '提交处理';
        }
    });

    function showResult(downloadUrl) {
        downloadLink.href = downloadUrl;
        resultArea.hidden = false;
        errorArea.hidden = true;
    }

    function showError(message) {
        errorText.textContent = message;
        errorArea.hidden = false;
        resultArea.hidden = true;
    }
});

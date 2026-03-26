document.addEventListener('DOMContentLoaded', function() {
    // 获取DOM元素
    const uploadForm = document.getElementById('uploadForm');
    const submitBtn = document.getElementById('submitBtn');
    const resultArea = document.getElementById('resultArea');
    const errorArea = document.getElementById('errorArea');
    const downloadLink = document.getElementById('downloadLink');
    const errorText = document.getElementById('errorText');
    const addTableBtn = document.getElementById('addTableBtn');
    const sourceTables = document.getElementById('sourceTables');

    // 最大源表数量
    const MAX_TABLES = 10;
    let tableCount = 1;

    // 初始化第一个表的上传功能
    initTableUpload(1);
    initBTableUpload();

    // 添加源表按钮事件
    addTableBtn.addEventListener('click', function() {
        if (tableCount >= MAX_TABLES) {
            alert(`最多只能添加${MAX_TABLES}个源表`);
            return;
        }
        tableCount++;
        addNewTable(tableCount);
        updateRemoveButtons();
        
        if (tableCount >= MAX_TABLES) {
            addTableBtn.disabled = true;
        }
    });

    // 添加新表
    function addNewTable(id) {
        const tableHtml = `
            <div class="file-section source-section" data-table-id="${id}">
                <div class="source-header">
                    <span class="source-title">表 ${id}</span>
                    <button type="button" class="remove-table-btn" onclick="removeTable(${id})">×</button>
                </div>
                <div class="upload-area" id="uploadArea${id}">
                    <input type="file" id="file${id}" name="file_${id}" accept=".xlsx,.xls" hidden>
                    <div class="upload-content">
                        <span class="upload-icon">&#128193;</span>
                        <p class="upload-text">点击或拖拽上传</p>
                        <p class="file-name" id="fileName${id}"></p>
                    </div>
                </div>
                <div class="config-area">
                    <div class="input-group">
                        <label>身份证号列</label>
                        <input type="text" id="idCol${id}" name="id_col_${id}" value="A" pattern="[A-Za-z]*" title="请输入列名字母，如A、B、C">
                    </div>
                    <div class="input-group">
                        <label>姓名列（可选）</label>
                        <input type="text" id="nameCol${id}" name="name_col_${id}" value="" pattern="[A-Za-z]*" title="请输入列名字母，如A、B、C">
                    </div>
                    <div class="input-group">
                        <label>工资列</label>
                        <input type="text" id="salaryCol${id}" name="salary_col_${id}" value="B" pattern="[A-Za-z]*" title="请输入列名字母，如A、B、C">
                    </div>
                </div>
            </div>
        `;
        sourceTables.insertAdjacentHTML('beforeend', tableHtml);
        initTableUpload(id);
    }

    // 删除表（全局函数）
    window.removeTable = function(id) {
        if (tableCount <= 1) {
            alert('至少保留一个源表');
            return;
        }
        
        const tableElement = document.querySelector(`[data-table-id="${id}"]`);
        if (tableElement) {
            tableElement.remove();
            tableCount--;
            renumberTables();
            updateRemoveButtons();
            addTableBtn.disabled = false;
        }
    };

    // 重新编号表格
    function renumberTables() {
        const tables = sourceTables.querySelectorAll('.source-section');
        tables.forEach((table, index) => {
            const newId = index + 1;
            const oldId = table.getAttribute('data-table-id');
            table.setAttribute('data-table-id', newId);
            
            // 更新标题
            table.querySelector('.source-title').textContent = `表 ${newId}`;
            
            // 更新删除按钮
            const removeBtn = table.querySelector('.remove-table-btn');
            removeBtn.setAttribute('onclick', `removeTable(${newId})`);
            
            // 更新各元素的ID和name
            updateElementId(table, 'uploadArea', oldId, newId);
            updateElementId(table, 'file', oldId, newId);
            updateElementId(table, 'fileName', oldId, newId);
            updateElementId(table, 'idCol', oldId, newId);
            updateElementId(table, 'nameCol', oldId, newId);
            updateElementId(table, 'salaryCol', oldId, newId);
            
            // 更新input的name属性
            const fileInput = table.querySelector(`#file${newId}`);
            if (fileInput) {
                fileInput.name = `file_${newId}`;
            }
            const idColInput = table.querySelector(`#idCol${newId}`);
            if (idColInput) {
                idColInput.name = `id_col_${newId}`;
            }
            const nameColInput = table.querySelector(`#nameCol${newId}`);
            if (nameColInput) {
                nameColInput.name = `name_col_${newId}`;
            }
            const salaryColInput = table.querySelector(`#salaryCol${newId}`);
            if (salaryColInput) {
                salaryColInput.name = `salary_col_${newId}`;
            }
        });
    }

    // 更新元素ID
    function updateElementId(parent, prefix, oldId, newId) {
        const element = parent.querySelector(`#${prefix}${oldId}`);
        if (element) {
            element.id = `${prefix}${newId}`;
        }
    }

    // 更新删除按钮显示状态
    function updateRemoveButtons() {
        const removeButtons = document.querySelectorAll('.remove-table-btn');
        removeButtons.forEach(btn => {
            btn.style.display = tableCount > 1 ? 'flex' : 'none';
        });
    }

    // 初始化表的上传功能
    function initTableUpload(tableId) {
        const uploadArea = document.getElementById(`uploadArea${tableId}`);
        const fileInput = document.getElementById(`file${tableId}`);
        const fileNameElement = document.getElementById(`fileName${tableId}`);

        if (!uploadArea || !fileInput) return;

        uploadArea.addEventListener('click', function() {
            fileInput.click();
        });

        fileInput.addEventListener('change', function() {
            handleFileSelect(this, uploadArea, fileNameElement);
        });

        // 拖拽上传
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            uploadArea.addEventListener(eventName, preventDefaults, false);
        });

        ['dragenter', 'dragover'].forEach(eventName => {
            uploadArea.addEventListener(eventName, () => highlight(uploadArea), false);
        });

        ['dragleave', 'drop'].forEach(eventName => {
            uploadArea.addEventListener(eventName, () => unhighlight(uploadArea), false);
        });

        uploadArea.addEventListener('drop', function(e) {
            const dt = e.dataTransfer;
            const files = dt.files;
            fileInput.files = files;
            handleFileSelect(fileInput, uploadArea, fileNameElement);
        });
    }

    // 初始化B表上传
    function initBTableUpload() {
        const uploadAreaB = document.getElementById('uploadAreaB');
        const fileB = document.getElementById('fileB');
        const fileNameB = document.getElementById('fileNameB');

        if (!uploadAreaB || !fileB) return;

        uploadAreaB.addEventListener('click', function() {
            fileB.click();
        });

        fileB.addEventListener('change', function() {
            handleFileSelect(this, uploadAreaB, fileNameB);
        });

        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            uploadAreaB.addEventListener(eventName, preventDefaults, false);
        });

        ['dragenter', 'dragover'].forEach(eventName => {
            uploadAreaB.addEventListener(eventName, () => highlight(uploadAreaB), false);
        });

        ['dragleave', 'drop'].forEach(eventName => {
            uploadAreaB.addEventListener(eventName, () => unhighlight(uploadAreaB), false);
        });

        uploadAreaB.addEventListener('drop', function(e) {
            const dt = e.dataTransfer;
            const files = dt.files;
            fileB.files = files;
            handleFileSelect(fileB, uploadAreaB, fileNameB);
        });
    }

    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    function highlight(element) {
        element.style.borderColor = '#667eea';
        element.style.background = '#f0f0ff';
    }

    function unhighlight(element) {
        if (!element.classList.contains('has-file')) {
            element.style.borderColor = '#ccc';
            element.style.background = '#fafafa';
        }
    }

    function handleFileSelect(input, uploadArea, fileNameElement) {
        const file = input.files[0];
        if (file) {
            const validTypes = ['.xlsx', '.xls'];
            const fileExtension = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
            
            if (!validTypes.includes(fileExtension)) {
                alert('请上传Excel文件(.xlsx 或 .xls)');
                input.value = '';
                return;
            }
            
            uploadArea.classList.add('has-file');
            fileNameElement.textContent = file.name;
        }
    }

    // 表单提交
    uploadForm.addEventListener('submit', async function(e) {
        e.preventDefault();

        // 验证B表是否已选择
        const fileB = document.getElementById('fileB');
        if (!fileB.files[0]) {
            showError('请上传目标表B');
            return;
        }

        // 验证是否至少上传了一个源表
        let hasSourceFile = false;
        const tables = sourceTables.querySelectorAll('.source-section');
        tables.forEach(table => {
            const tableId = table.getAttribute('data-table-id');
            const fileInput = document.getElementById(`file${tableId}`);
            if (fileInput && fileInput.files[0]) {
                hasSourceFile = true;
            }
        });
        
        if (!hasSourceFile) {
            showError('请至少上传一个源表');
            return;
        }

        // 隐藏之前的结果
        resultArea.hidden = true;
        errorArea.hidden = true;

        // 显示加载状态
        setLoading(true);

        // 创建FormData
        const formData = new FormData();
        
        // 添加源表文件
        tables.forEach(table => {
            const tableId = table.getAttribute('data-table-id');
            const fileInput = document.getElementById(`file${tableId}`);
            if (fileInput && fileInput.files[0]) {
                formData.append(`file_${tableId}`, fileInput.files[0]);
                formData.append(`id_col_${tableId}`, document.getElementById(`idCol${tableId}`).value);
                formData.append(`salary_col_${tableId}`, document.getElementById(`salaryCol${tableId}`).value);
                formData.append(`name_col_${tableId}`, document.getElementById(`nameCol${tableId}`).value);
            }
        });
        
        // 添加B表
        formData.append('file_b', fileB.files[0]);
        formData.append('id_col_b', document.getElementById('idColB').value);
        formData.append('salary_col_b', document.getElementById('salaryColB').value);
        formData.append('name_col_b', document.getElementById('nameColB').value);

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
            setLoading(false);
        }
    });

    function setLoading(loading) {
        submitBtn.disabled = loading;
        submitBtn.querySelector('.btn-text').hidden = loading;
        submitBtn.querySelector('.btn-loading').hidden = !loading;
    }

    function showResult(downloadUrl) {
        downloadLink.href = downloadUrl;
        resultArea.hidden = false;
        errorArea.hidden = true;
        resultArea.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }

    function showError(message) {
        errorText.textContent = message;
        errorArea.hidden = false;
        resultArea.hidden = true;
        errorArea.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }
});

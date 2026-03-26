document.addEventListener('DOMContentLoaded', function() {
    // 获取DOM元素
    const uploadForm = document.getElementById('uploadForm');
    const submitBtn = document.getElementById('submitBtn');
    const resultArea = document.getElementById('resultArea');
    const errorArea = document.getElementById('errorArea');
    const downloadLink = document.getElementById('downloadLink');
    const errorText = document.getElementById('errorText');

    // 初始化5个源表和B表的上传区域
    const tables = [
        { id: 1, area: 'uploadArea1', file: 'file1', name: 'fileName1' },
        { id: 2, area: 'uploadArea2', file: 'file2', name: 'fileName2' },
        { id: 3, area: 'uploadArea3', file: 'file3', name: 'fileName3' },
        { id: 4, area: 'uploadArea4', file: 'file4', name: 'fileName4' },
        { id: 5, area: 'uploadArea5', file: 'file5', name: 'fileName5' },
        { id: 'B', area: 'uploadAreaB', file: 'fileB', name: 'fileNameB', required: true }
    ];

    // 为每个表初始化上传功能
    tables.forEach(table => {
        const uploadArea = document.getElementById(table.area);
        const fileInput = document.getElementById(table.file);
        const fileNameElement = document.getElementById(table.name);

        if (uploadArea && fileInput) {
            // 点击上传区域触发文件选择
            uploadArea.addEventListener('click', function() {
                fileInput.click();
            });

            // 文件选择事件
            fileInput.addEventListener('change', function() {
                handleFileSelect(this, uploadArea, fileNameElement);
            });

            // 拖拽上传支持
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
    });

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

    // 处理文件选择
    function handleFileSelect(input, uploadArea, fileNameElement) {
        const file = input.files[0];
        if (file) {
            // 验证文件类型
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
        for (let i = 1; i <= 5; i++) {
            const fileInput = document.getElementById(`file${i}`);
            if (fileInput && fileInput.files[0]) {
                hasSourceFile = true;
                break;
            }
        }
        
        if (!hasSourceFile) {
            showError('请至少上传一个源表（表1-表5）');
            return;
        }

        // 隐藏之前的结果
        resultArea.hidden = true;
        errorArea.hidden = true;

        // 显示加载状态
        setLoading(true);

        // 创建FormData
        const formData = new FormData();
        
        // 添加源表文件（1-5）
        for (let i = 1; i <= 5; i++) {
            const fileInput = document.getElementById(`file${i}`);
            if (fileInput && fileInput.files[0]) {
                formData.append(`file_${i}`, fileInput.files[0]);
                formData.append(`id_col_${i}`, document.getElementById(`idCol${i}`).value);
                formData.append(`salary_col_${i}`, document.getElementById(`salaryCol${i}`).value);
                formData.append(`name_col_${i}`, document.getElementById(`nameCol${i}`).value);
            }
        }
        
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

    // 设置加载状态
    function setLoading(loading) {
        submitBtn.disabled = loading;
        submitBtn.querySelector('.btn-text').hidden = loading;
        submitBtn.querySelector('.btn-loading').hidden = !loading;
    }

    // 显示成功结果
    function showResult(downloadUrl) {
        downloadLink.href = downloadUrl;
        resultArea.hidden = false;
        errorArea.hidden = true;
        resultArea.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }

    // 显示错误信息
    function showError(message) {
        errorText.textContent = message;
        errorArea.hidden = false;
        resultArea.hidden = true;
        errorArea.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }
});

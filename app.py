import os
import pandas as pd
from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename
import uuid

app = Flask(__name__)

# 配置
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB限制

# 确保上传目录存在
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def excel_column_to_index(col_str):
    """
    将Excel列名（如A、B、C、AA）转换为0-based索引
    """
    col_str = col_str.upper().strip()
    result = 0
    for char in col_str:
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result - 1


def process_single_table(df_source, df_target, id_col_source, salary_col_source, name_col_source,
                         id_col_target, salary_col_target, name_col_target):
    """
    处理单个源表数据，合并到目标表
    :param df_source: 源表DataFrame
    :param df_target: 目标表DataFrame
    :param id_col_source: 源表身份证号列索引
    :param salary_col_source: 源表工资列索引
    :param name_col_source: 源表姓名列索引（可选）
    :param id_col_target: 目标表身份证号列索引
    :param salary_col_target: 目标表工资列索引
    :param name_col_target: 目标表姓名列索引（可选）
    :return: 处理后的目标表DataFrame
    """
    # 将身份证号作为字符串处理，去除空格
    df_source[id_col_source] = df_source[id_col_source].astype(str).str.strip()
    df_target[id_col_target] = df_target[id_col_target].astype(str).str.strip()
    
    # 过滤源表，只保留身份证号为18位的行（中国大陆身份证号标准长度）
    df_source = df_source[df_source[id_col_source].str.match(r'^\d{17}[\dXx]$', na=False)]
    
    # 过滤目标表，只保留身份证号为18位的行或空行
    df_target_original = df_target.copy()  # 保留原始目标表用于保存非18位行
    df_target_valid = df_target[df_target[id_col_target].str.match(r'^\d{17}[\dXx]$', na=False)]
    
    # 将工资列转换为数值类型
    df_source[salary_col_source] = pd.to_numeric(df_source[salary_col_source], errors='coerce').fillna(0)
    df_target_valid[salary_col_target] = pd.to_numeric(df_target_valid[salary_col_target], errors='coerce').fillna(0)
    
    # 按身份证号分组，将源表中相同身份证号的工资相加，同时保留第一个姓名
    if name_col_source is not None:
        name_first = df_source.groupby(id_col_source)[name_col_source].first().reset_index()
        salary_sum = df_source.groupby(id_col_source)[salary_col_source].sum().reset_index()
        salary_sum = salary_sum.merge(name_first, on=id_col_source, how='left')
    else:
        salary_sum = df_source.groupby(id_col_source)[salary_col_source].sum().reset_index()
    
    # 用于记录需要添加到目标表的新行
    new_rows = []
    
    # 遍历源表的汇总数据
    for _, row in salary_sum.iterrows():
        id_card = row[id_col_source]
        salary = row[salary_col_source]
        name = row.get(name_col_source, '') if name_col_source is not None else ''
        
        # 在目标表中查找匹配的身份证号（只匹配18位身份证号）
        mask = df_target_valid[id_col_target] == id_card
        if mask.any():
            # 匹配上，将源表工资加到目标表原有工资上
            df_target_valid.loc[mask, salary_col_target] = df_target_valid.loc[mask, salary_col_target] + salary
        else:
            # 未匹配上，准备新增行
            new_row = {col: '' for col in df_target_valid.columns}
            new_row[id_col_target] = id_card
            new_row[salary_col_target] = salary
            # 如果配置了姓名列，填入姓名
            if name_col_target is not None and name_col_source is not None:
                new_row[name_col_target] = name
            new_rows.append(new_row)
    
    # 将新行添加到目标表
    if new_rows:
        df_new_rows = pd.DataFrame(new_rows)
        df_target_valid = pd.concat([df_target_valid, df_new_rows], ignore_index=True)
    
    # 将非18位身份证号的行添加回结果（保持原有数据不变）
    df_target_invalid = df_target_original[~df_target_original[id_col_target].str.match(r'^\d{17}[\dXx]$', na=False)]
    if not df_target_invalid.empty:
        df_target_valid = pd.concat([df_target_valid, df_target_invalid], ignore_index=True)
    
    return df_target_valid


def process_excel(source_files, file_b_path, source_configs, id_col_b, salary_col_b, name_col_b):
    """
    处理多个Excel文件，按顺序合并到B表
    :param source_files: 源表文件路径列表（表1-表5）
    :param file_b_path: B表路径
    :param source_configs: 源表配置列表，每个元素为(id_col, salary_col, name_col)
    :param id_col_b: B表身份证号列名（如A、B、C）
    :param salary_col_b: B表需要填入工资信息的列名（如A、B、C）
    :param name_col_b: B表姓名列名（如A、B、C）
    :return: 处理后的DataFrame
    """
    # 读取B表
    df_b = pd.read_excel(file_b_path, header=None)
    
    # 将B表列名字母转换为索引
    id_col_b_idx = excel_column_to_index(id_col_b)
    salary_col_b_idx = excel_column_to_index(salary_col_b)
    name_col_b_idx = excel_column_to_index(name_col_b) if name_col_b else None
    
    # 获取B表列名
    id_col_b_name = df_b.columns[id_col_b_idx]
    salary_col_b_name = df_b.columns[salary_col_b_idx]
    name_col_b_name = df_b.columns[name_col_b_idx] if name_col_b_idx is not None else None
    
    # 按顺序处理每个源表
    for i, (source_file, source_config) in enumerate(zip(source_files, source_configs)):
        if source_file is None:
            continue  # 跳过未上传的表
            
        id_col_source, salary_col_source, name_col_source = source_config
        
        # 读取源表
        df_source = pd.read_excel(source_file, header=None)
        
        # 将源表列名字母转换为索引
        id_col_source_idx = excel_column_to_index(id_col_source)
        salary_col_source_idx = excel_column_to_index(salary_col_source)
        name_col_source_idx = excel_column_to_index(name_col_source) if name_col_source else None
        
        # 获取源表列名
        id_col_source_name = df_source.columns[id_col_source_idx]
        salary_col_source_name = df_source.columns[salary_col_source_idx]
        name_col_source_name = df_source.columns[name_col_source_idx] if name_col_source_idx is not None else None
        
        # 处理当前源表，合并到B表
        df_b = process_single_table(
            df_source, df_b,
            id_col_source_name, salary_col_source_name, name_col_source_name,
            id_col_b_name, salary_col_b_name, name_col_b_name
        )
    
    return df_b


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/process', methods=['POST'])
def process():
    try:
        # 检查B表是否上传
        if 'file_b' not in request.files:
            return jsonify({'error': '请上传目标表B'}), 400
        
        file_b = request.files['file_b']
        if file_b.filename == '':
            return jsonify({'error': '请选择目标表B'}), 400
        
        if not allowed_file(file_b.filename):
            return jsonify({'error': '目标表B只允许上传Excel文件(.xlsx, .xls)'}), 400
        
        # 收集所有源表文件（表1-表5）
        source_files = []
        source_configs = []
        saved_source_files = []  # 用于后续清理
        
        import re
        col_pattern = re.compile(r'^[A-Za-z]*$')
        unique_id = str(uuid.uuid4())
        
        for i in range(1, 6):
            file_key = f'file_{i}'
            if file_key in request.files:
                source_file = request.files[file_key]
                if source_file.filename != '' and allowed_file(source_file.filename):
                    # 获取该源表的列配置
                    id_col = request.form.get(f'id_col_{i}', 'A').strip()
                    salary_col = request.form.get(f'salary_col_{i}', 'B').strip()
                    name_col = request.form.get(f'name_col_{i}', '').strip()
                    
                    # 验证列名格式
                    for col, name in [(id_col, f'表{i}身份证号列'), (salary_col, f'表{i}工资列'), 
                                      (name_col, f'表{i}姓名列')]:
                        if col and not col_pattern.match(col):
                            return jsonify({'error': f'{name}必须是字母（如A、B、C、AA）或留空'}), 400
                    
                    # 保存源表文件
                    filename = secure_filename(f"{unique_id}_source_{i}_{source_file.filename}")
                    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                    source_file.save(filepath)
                    saved_source_files.append(filepath)
                    source_files.append(filepath)
                    source_configs.append((id_col, salary_col, name_col))
                else:
                    source_files.append(None)
                    source_configs.append(('A', 'B', ''))  # 默认配置
            else:
                source_files.append(None)
                source_configs.append(('A', 'B', ''))
        
        # 检查是否至少上传了一个源表
        if all(f is None for f in source_files):
            return jsonify({'error': '请至少上传一个源表（表1-表5）'}), 400
        
        # 获取B表列配置
        id_col_b = request.form.get('id_col_b', 'A').strip()
        salary_col_b = request.form.get('salary_col_b', 'B').strip()
        name_col_b = request.form.get('name_col_b', '').strip()
        
        # 验证B表列名格式
        for col, name in [(id_col_b, 'B表身份证号列'), (salary_col_b, 'B表工资列'), (name_col_b, 'B表姓名列')]:
            if col and not col_pattern.match(col):
                return jsonify({'error': f'{name}必须是字母（如A、B、C、AA）或留空'}), 400
        
        # 保存B表文件
        filename_b = secure_filename(f"{unique_id}_b_{file_b.filename}")
        filepath_b = os.path.join(app.config['UPLOAD_FOLDER'], filename_b)
        file_b.save(filepath_b)
        
        # 处理Excel
        result_df = process_excel(source_files, filepath_b, source_configs, 
                                   id_col_b, salary_col_b, name_col_b)
        
        # 保存结果
        output_filename = f"result_{unique_id}.xlsx"
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
        
        # 使用openpyxl引擎保存，保持样式
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            result_df.to_excel(writer, index=False)
        
        # 清理上传的文件
        for filepath in saved_source_files:
            if os.path.exists(filepath):
                os.remove(filepath)
        os.remove(filepath_b)
        
        # 返回下载链接
        return jsonify({
            'success': True,
            'download_url': f'/download/{output_filename}'
        })
        
    except Exception as e:
        return jsonify({'error': f'处理失败: {str(e)}'}), 500


@app.route('/download/<filename>')
def download(filename):
    try:
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        if os.path.exists(file_path):
            return send_file(file_path, as_attachment=True, download_name='处理结果.xlsx')
        else:
            return jsonify({'error': '文件不存在'}), 404
    except Exception as e:
        return jsonify({'error': f'下载失败: {str(e)}'}), 500


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)

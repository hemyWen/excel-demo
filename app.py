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
                         id_col_target, salary_col_target, name_col_target, source_table_name):
    """
    处理单个源表数据，合并到目标表
    :param df_source: 源表 DataFrame
    :param df_target: 目标表 DataFrame（只包含三列：身份证、姓名、工资）
    :param id_col_source: 源表身份证号列索引
    :param salary_col_source: 源表工资列索引
    :param name_col_source: 源表姓名列索引（可选）
    :param id_col_target: 目标表身份证号列索引
    :param salary_col_target: 目标表工资列索引
    :param name_col_target: 目标表姓名列索引（可选）
    :param source_table_name: 源表名称（如"表 1"、"表 2"等）
    :return: 处理后的目标表 DataFrame（包含四列：身份证、姓名、工资、来源表）
    """
    # 将身份证号作为字符串处理，去除空格
    df_source[id_col_source] = df_source[id_col_source].astype(str).str.strip()
    df_target[id_col_target] = df_target[id_col_target].astype(str).str.strip()
    
    # 过滤源表，只保留身份证号为 18 位的行（中国大陆身份证号标准长度）
    df_source = df_source[df_source[id_col_source].str.match(r'^\d{17}[\dXx]$', na=False)]
    
    # 过滤目标表，只保留身份证号为 18 位的行
    df_target = df_target[df_target[id_col_target].str.match(r'^\d{17}[\dXx]$', na=False)]
    
    # 将工资列转换为浮点型数值
    df_source[salary_col_source] = pd.to_numeric(df_source[salary_col_source], errors='coerce').fillna(0.0).astype(float)
    df_target[salary_col_target] = pd.to_numeric(df_target[salary_col_target], errors='coerce').fillna(0.0).astype(float)
    
    # 不分组，直接按源表原始顺序处理每一行
    # 用于记录需要添加到目标表的新行
    new_rows = []
    added_ids = set()  # 记录本次新增的身份证号，避免重复添加
    
    # 遍历源表的每一行（保持原始顺序）
    for idx, row in df_source.iterrows():
        id_card = row[id_col_source]
        salary = row[salary_col_source]
        name = row[name_col_source] if name_col_source is not None else ''
        
        # 在目标表中查找匹配的身份证号（只匹配 18 位身份证号）
        mask = df_target[id_col_target] == id_card
        if mask.any():
            # 匹配上，将源表工资加到目标表原有工资上
            df_target.loc[mask, salary_col_target] = df_target.loc[mask, salary_col_target] + salary
            # 如果目标表姓名为空，填入源表姓名
            if name_col_target is not None and name_col_source is not None:
                df_target.loc[mask, name_col_target] = df_target.loc[mask, name_col_target].replace('', name)
            # 注意：匹配到的行不修改来源表列
        else:
            # 未匹配上，且本次还未添加过这个身份证号，才准备新增行
            if id_card not in added_ids:
                # 准备新增行
                new_row = {id_col_target: id_card, salary_col_target: salary}
                # 如果配置了姓名列，填入姓名
                if name_col_target is not None and name_col_source is not None:
                    new_row[name_col_target] = name
                else:
                    new_row[name_col_target] = ''
                # 填入来源表名称（只有新增的行才填）
                if '来源表' in df_target.columns:
                    new_row['来源表'] = source_table_name
                new_rows.append(new_row)
                added_ids.add(id_card)
    
    # 将新行添加到目标表
    if new_rows:
        df_new_rows = pd.DataFrame(new_rows)
        df_target = pd.concat([df_target, df_new_rows], ignore_index=True)
    
    return df_target


def process_excel(source_files, file_b_path, id_col_b, salary_col_b, name_col_b):
    """
    处理多个 Excel 文件，按顺序合并到 B 表
    :param source_files: 源表文件列表，每个元素为 (文件路径, 列配置, 文件名) 的元组
    :param file_b_path: B 表路径
    :param id_col_b: B 表身份证号列名（如 A、B、C）
    :param salary_col_b: B 表需要填入工资信息的列名（如 A、B、C）
    :param name_col_b: B 表姓名列名（如 A、B、C）
    :return: 处理后的 DataFrame（包含四列：身份证、姓名、工资、来源表）
    """
    # 读取B表
    df_b_full = pd.read_excel(file_b_path, header=None)
    
    # 将B表列名字母转换为索引
    id_col_b_idx = excel_column_to_index(id_col_b)
    salary_col_b_idx = excel_column_to_index(salary_col_b)
    name_col_b_idx = excel_column_to_index(name_col_b) if name_col_b else None
    
    # 获取B表列名
    id_col_b_name = df_b_full.columns[id_col_b_idx]
    salary_col_b_name = df_b_full.columns[salary_col_b_idx]
    name_col_b_name = df_b_full.columns[name_col_b_idx] if name_col_b_idx is not None else None
    
    # 创建只包含三列的新 DataFrame
    columns_to_keep = [id_col_b_name, salary_col_b_name]
    if name_col_b_name is not None:
        columns_to_keep.append(name_col_b_name)
        
    df_b = df_b_full[columns_to_keep].copy()
        
    # 添加来源表列，初始为空字符串
    df_b['来源表'] = ''
    
    # 按顺序处理每个源表
    for source_file_path, source_config, source_table_name in source_files:
        id_col_source, salary_col_source, name_col_source = source_config
        
        # 读取源表
        df_source = pd.read_excel(source_file_path, header=None)
        
        # 将源表列名字母转换为索引
        id_col_source_idx = excel_column_to_index(id_col_source)
        salary_col_source_idx = excel_column_to_index(salary_col_source)
        name_col_source_idx = excel_column_to_index(name_col_source) if name_col_source else None
        
        # 获取源表列名
        id_col_source_name = df_source.columns[id_col_source_idx]
        salary_col_source_name = df_source.columns[salary_col_source_idx]
        name_col_source_name = df_source.columns[name_col_source_idx] if name_col_source_idx is not None else None
                        
        # 处理当前源表，合并到 B 表
        df_b = process_single_table(
            df_source, df_b,
            id_col_source_name, salary_col_source_name, name_col_source_name,
            id_col_b_name, salary_col_b_name, name_col_b_name,
            source_table_name
        )
    
    # 只保留四列，并按顺序重命名
    result_columns = [id_col_b_name]
    if name_col_b_name is not None:
        result_columns.append(name_col_b_name)
    result_columns.append(salary_col_b_name)
    result_columns.append('来源表')
    
    df_result = df_b[result_columns].copy()
    
    # 重命名列为中文表头
    new_column_names = ['身份证号']
    if name_col_b_name is not None:
        new_column_names.append('姓名')
    new_column_names.append('工资')
    new_column_names.append('来源表')
    
    df_result.columns = new_column_names
    
    return df_result


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/match')
def match():
    return render_template('match.html')


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
        
        # 动态获取所有上传的源表（从 file_1 到 file_10）
        source_files = []  # (文件路径, 列配置, 文件名) 的元组列表
        saved_source_files = []  # 用于后续清理
                
        import re
        col_pattern = re.compile(r'^[A-Za-z]*$')
        unique_id = str(uuid.uuid4())
                
        for i in range(1, 11):
            file_key = f'file_{i}'
            if file_key in request.files:
                source_file = request.files[file_key]
                if source_file.filename != '' and allowed_file(source_file.filename):
                    # 获取该源表的列配置
                    id_col = request.form.get(f'id_col_{i}', 'A').strip()
                    salary_col = request.form.get(f'salary_col_{i}', 'B').strip()
                    name_col = request.form.get(f'name_col_{i}', '').strip()
                            
                    # 验证列名格式
                    for col, name in [(id_col, f'源表{i}身份证号列'), (salary_col, f'源表{i}工资列'), 
                                      (name_col, f'源表{i}姓名列')]:
                        if col and not col_pattern.match(col):
                            return jsonify({'error': f'{name}必须是字母（如 A、B、C、AA）或留空'}), 400
                            
                    # 保存真实文件名（去掉后缀）
                    import os
                    filename_without_ext = os.path.splitext(source_file.filename)[0]
                            
                    # 保存源表文件
                    filename = secure_filename(f"{unique_id}_source_{i}_{source_file.filename}")
                    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                    source_file.save(filepath)
                    saved_source_files.append(filepath)
                    
                    # 保存文件路径、配置和文件名
                    source_files.append((filepath, (id_col, salary_col, name_col), filename_without_ext))
        
        # 检查是否至少上传了一个源表
        if not source_files:
            return jsonify({'error': '请至少上传一个源数据表文件'}), 400
        
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
        
        # 处理 Excel
        result_df = process_excel(source_files, filepath_b, 
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

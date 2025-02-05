from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
import os
from werkzeug.utils import secure_filename
import uuid

app = Flask(__name__)
# 使用绝对路径
app.config['UPLOAD_FOLDER'] = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max-limit

# 确保上传目录存在
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    if 'files[]' not in request.files:
        return jsonify({'error': '没有选择文件'}), 400
    
    files = request.files.getlist('files[]')
    merge_type = request.form.get('merge_type', 'sheet')  # 可选值: 'sheet'(分sheet保存), 'same_sheet'(同名sheet合并), 'single_sheet'(合并成一个sheet)
    output_format = request.form.get('output_format', 'xlsx')
    custom_filename = request.form.get('custom_filename', '')
    
    if not files or files[0].filename == '':
        return jsonify({'error': '没有选择文件'}), 400

    # 保存上传的文件并处理
    saved_files = []
    try:
        for file in files:
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                temp_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(temp_path)
                saved_files.append(temp_path)

        # 使用自定义文件名或生成默认文件名
        if custom_filename:
            output_filename = secure_filename(custom_filename)
        else:
            output_filename = f"merged_{uuid.uuid4().hex[:8]}.{output_format}"
        
        output_path = merge_files(saved_files, merge_type, output_format, output_filename)
        
        # 清理临时文件
        for file_path in saved_files:
            try:
                if os.path.exists(file_path):
                    os.remove(file_path)
            except Exception as e:
                print(f"清理临时文件失败: {str(e)}")
            
        return jsonify({
            'success': True,
            'download_path': output_path
        })
    except Exception as e:
        # 发生错误时也要清理临时文件
        for file_path in saved_files:
            try:
                if os.path.exists(file_path):
                    os.remove(file_path)
            except:
                pass
        return jsonify({'error': str(e)}), 500

@app.route('/download/<filename>')
def download_file(filename):
    try:
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        if not os.path.exists(file_path):
            return jsonify({'error': '文件不存在'}), 404
        return send_file(file_path, as_attachment=True)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/preview', methods=['POST'])
def preview_sheets():
    if 'files[]' not in request.files:
        return jsonify({'error': '没有选择文件'}), 400
    
    files = request.files.getlist('files[]')
    merge_type = request.form.get('merge_type', 'sheet')
    
    if not files or files[0].filename == '':
        return jsonify({'error': '没有选择文件'}), 400

    # 保存上传的文件并获取预览信息
    saved_files = []
    try:
        sheet_names = []
        for file in files:
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                temp_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(temp_path)
                saved_files.append(temp_path)
                
                # 获取预览的sheet名
                if merge_type == 'sheet':
                    if file.filename.endswith('.csv'):
                        base_name = os.path.splitext(filename)[0]
                        sheet_names.append({
                            'original': base_name,
                            'sanitized': sanitize_sheet_name(base_name),
                            'file': filename
                        })
                    else:
                        file_base_name = os.path.splitext(filename)[0]
                        sheet_name = sanitize_sheet_name(file_base_name)
                        sheet_names.append({
                            'original': file_base_name,
                            'sanitized': sheet_name,
                            'file': filename
                        })

        # 清理临时文件
        for file_path in saved_files:
            try:
                if os.path.exists(file_path):
                    os.remove(file_path)
            except:
                pass

        return jsonify({
            'success': True,
            'sheet_names': sheet_names
        })
    except Exception as e:
        # 清理临时文件
        for file_path in saved_files:
            try:
                if os.path.exists(file_path):
                    os.remove(file_path)
            except:
                pass
        return jsonify({'error': str(e)}), 500

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in {'xlsx', 'xls', 'csv'}

def sanitize_sheet_name(sheet_name):
    """
    处理sheet名称，确保符合Excel的要求：
    1. 长度不超过31个字符
    2. 只替换Excel不允许的特殊字符 [ ] : * ? / \
    3. 保留其他字符，如 - 和 _
    """
    # 只替换Excel不允许的字符
    invalid_chars = ['[', ']', ':', '*', '?', '/', '\\']
    for char in invalid_chars:
        sheet_name = sheet_name.replace(char, '-')
    
    # 合并连续的 - 和 _
    while '--' in sheet_name:
        sheet_name = sheet_name.replace('--', '-')
    while '__' in sheet_name:
        sheet_name = sheet_name.replace('__', '_')
    
    # 提取共同部分
    # 例如：从 "0201-0205搜索词-投放数据（MV-华北-新东方在线考研-02-b）" 提取 "0201-0205-MV-b"
    parts = sheet_name.split('（')
    if len(parts) > 1:
        prefix = parts[0].split('-')[0:2]  # 取前两部分 "0201-0205"
        suffix = parts[1].split('-')[-2:]   # 取后两部分 "02-b）"
        sheet_name = '-'.join(prefix + ['MV'] + [s.rstrip('）') for s in suffix])
    
    # 截断长度（预留空间给可能的数字后缀）
    if len(sheet_name) > 28:
        sheet_name = sheet_name[:28]
    
    return sheet_name

def get_unique_sheet_name(sheet_name, existing_names):
    """
    确保sheet名称唯一，如果重复则添加数字后缀
    """
    base_name = sanitize_sheet_name(sheet_name)
    final_name = base_name
    counter = 1
    
    while final_name in existing_names:
        final_name = f"{base_name}_{counter}"
        counter += 1
    
    return final_name

def merge_files(file_paths, merge_type, output_format='xlsx', output_filename=None):
    if not output_filename:
        output_filename = f"merged_{uuid.uuid4().hex[:8]}.{output_format}"
    
    output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
    
    if merge_type == 'single_sheet':
        # 合并成单个sheet
        all_data = []
        for file_path in file_paths:
            try:
                if file_path.endswith('.csv'):
                    df = pd.read_csv(file_path)
                    df['来源文件'] = os.path.basename(file_path)
                    all_data.append(df)
                else:
                    xls = pd.ExcelFile(file_path)
                    for sheet in xls.sheet_names:
                        df = pd.read_excel(file_path, sheet_name=sheet)
                        df['来源文件'] = f"{os.path.basename(file_path)} - {sheet}"
                        all_data.append(df)
            except Exception as e:
                raise Exception(f"处理文件 {os.path.basename(file_path)} 时出错: {str(e)}")
        
        if all_data:
            merged_df = pd.concat(all_data, ignore_index=True)
            if output_format == 'csv':
                merged_df.to_csv(output_path, index=False, encoding='utf-8-sig')
            else:
                merged_df.to_excel(output_path, sheet_name='合并数据', index=False)
        
        return output_filename
    
    elif merge_type == 'sheet':
        if output_format == 'csv':
            # CSV格式：合并所有数据到一个CSV文件
            all_data = []
            for file_path in file_paths:
                try:
                    if file_path.endswith('.csv'):
                        df = pd.read_csv(file_path)
                    else:
                        df = pd.read_excel(file_path)
                    # 添加文件名作为来源列
                    df['来源文件'] = os.path.basename(file_path)
                    all_data.append(df)
                except Exception as e:
                    raise Exception(f"处理文件 {os.path.basename(file_path)} 时出错: {str(e)}")
            
            if all_data:
                merged_df = pd.concat(all_data, ignore_index=True)
                merged_df.to_csv(output_path, index=False, encoding='utf-8-sig')
        else:
            # Excel格式：按原来的处理方式
            with pd.ExcelWriter(output_path) as writer:
                existing_names = set()
                for file_path in file_paths:
                    try:
                        if file_path.endswith('.csv'):
                            df = pd.read_csv(file_path)
                            base_name = os.path.splitext(os.path.basename(file_path))[0]
                            sheet_name = get_unique_sheet_name(base_name, existing_names)
                            existing_names.add(sheet_name)
                            df.to_excel(writer, sheet_name=sheet_name, index=False)
                        else:
                            xls = pd.ExcelFile(file_path)
                            file_base_name = os.path.splitext(os.path.basename(file_path))[0]
                            sheet_name = sanitize_sheet_name(file_base_name)
                            if sheet_name not in existing_names:
                                df = pd.read_excel(file_path)
                                existing_names.add(sheet_name)
                                df.to_excel(writer, sheet_name=sheet_name, index=False)
                    except Exception as e:
                        raise Exception(f"处理文件 {os.path.basename(file_path)} 时出错: {str(e)}")
    else:
        # 相同sheet名合并
        if output_format == 'csv':
            # CSV格式：合并所有数据到一个CSV文件
            all_data = []
            for file_path in file_paths:
                try:
                    if file_path.endswith('.csv'):
                        df = pd.read_csv(file_path)
                        df['来源文件'] = os.path.basename(file_path)
                        all_data.append(df)
                    else:
                        xls = pd.ExcelFile(file_path)
                        for sheet in xls.sheet_names:
                            df = pd.read_excel(file_path, sheet_name=sheet)
                            df['来源文件'] = f"{os.path.basename(file_path)} - {sheet}"
                            all_data.append(df)
                except Exception as e:
                    raise Exception(f"处理文件 {os.path.basename(file_path)} 时出错: {str(e)}")
            
            if all_data:
                merged_df = pd.concat(all_data, ignore_index=True)
                merged_df.to_csv(output_path, index=False, encoding='utf-8-sig')
        else:
            # Excel格式：按原来的处理方式
            sheet_data = {}
            for file_path in file_paths:
                try:
                    if file_path.endswith('.csv'):
                        df = pd.read_csv(file_path)
                        sheet_name = 'Sheet1'
                        if sheet_name not in sheet_data:
                            sheet_data[sheet_name] = []
                        sheet_data[sheet_name].append(df)
                    else:
                        xls = pd.ExcelFile(file_path)
                        for sheet in xls.sheet_names:
                            df = pd.read_excel(file_path, sheet_name=sheet)
                            if sheet not in sheet_data:
                                sheet_data[sheet] = []
                            sheet_data[sheet].append(df)
                except Exception as e:
                    raise Exception(f"处理文件 {os.path.basename(file_path)} 时出错: {str(e)}")
            
            with pd.ExcelWriter(output_path) as writer:
                for sheet, dfs in sheet_data.items():
                    merged_df = pd.concat(dfs, ignore_index=True)
                    merged_df.to_excel(writer, sheet_name=sheet, index=False)
    
    return output_filename

if __name__ == '__main__':
    app.run(debug=True) 
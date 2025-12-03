"""
Excel拆分工具 - Web界面
提供文件上传、拆分配置和下载功能
"""
from flask import Flask, render_template, request, send_file, jsonify, send_from_directory
import os
import shutil
from pathlib import Path
from werkzeug.utils import secure_filename
from excel_splitter import ExcelSplitter
from excel_merger import ExcelMerger
import zipfile
from datetime import datetime
import pandas as pd

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB最大文件大小
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'outputs'

# 创建必要的文件夹
Path(app.config['UPLOAD_FOLDER']).mkdir(parents=True, exist_ok=True)
Path(app.config['OUTPUT_FOLDER']).mkdir(parents=True, exist_ok=True)

ALLOWED_EXTENSIONS = {'xlsx', 'xls'}


def allowed_file(filename):
    """检查文件扩展名是否允许"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def index():
    """主页"""
    return render_template('index.html')


@app.route('/merger')
def merger():
    """合并页面"""
    return render_template('merger.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    """处理文件上传"""
    if 'file' not in request.files:
        return jsonify({'error': '没有文件'}), 400
    
    file = request.files['file']
    
    if file.filename == '':
        return jsonify({'error': '未选择文件'}), 400
    
    if not allowed_file(file.filename):
        return jsonify({'error': '只支持 .xlsx 和 .xls 文件'}), 400
    
    # 保存文件
    filename = secure_filename(file.filename)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    unique_filename = f"{timestamp}_{filename}"
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], unique_filename)
    file.save(filepath)
    
    # 读取sheet和列信息
    try:
        import pandas as pd
        excel_file = pd.ExcelFile(filepath)
        sheets = excel_file.sheet_names
        
        # 读取第一个sheet的列名
        df = pd.read_excel(filepath, sheet_name=sheets[0])
        columns = df.columns.tolist()
        
        return jsonify({
            'filename': unique_filename,
            'sheets': sheets,
            'columns': columns
        })
    except Exception as e:
        # 清理上传的文件
        if os.path.exists(filepath):
            os.remove(filepath)
        return jsonify({'error': f'读取Excel文件失败: {str(e)}'}), 400


@app.route('/preview', methods=['POST'])
def preview_split():
    """预览拆分结果"""
    data = request.json
    filename = data.get('filename')
    split_column = data.get('split_column')
    
    if not filename or not split_column:
        return jsonify({'error': '缺少必要参数'}), 400
    
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    
    if not os.path.exists(filepath):
        return jsonify({'error': '文件不存在'}), 404
    
    try:
        splitter = ExcelSplitter(filepath, split_column, app.config['OUTPUT_FOLDER'])
        sheets = splitter.read_all_sheets()
        unique_values = splitter.get_unique_values(sheets)
        
        # 统计每个值在各个sheet中的数据量
        preview_data = []
        for value in unique_values:
            value_data = {
                'value': str(value),
                'sheets': {}
            }
            total_rows = 0
            
            for sheet_name, df in sheets.items():
                if split_column in df.columns:
                    filtered_df = df[df[split_column] == value]
                    row_count = len(filtered_df)
                    if row_count > 0:
                        value_data['sheets'][sheet_name] = row_count
                        total_rows += row_count
            
            value_data['total_rows'] = total_rows
            preview_data.append(value_data)
        
        return jsonify({
            'preview': preview_data,
            'total_files': len(unique_values),
            'sheet_names': list(sheets.keys())
        })
    except Exception as e:
        return jsonify({'error': f'预览失败: {str(e)}'}), 400


@app.route('/split', methods=['POST'])
def split_file():
    """执行拆分"""
    data = request.json
    filename = data.get('filename')
    split_column = data.get('split_column')
    
    if not filename or not split_column:
        return jsonify({'error': '缺少必要参数'}), 400
    
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    
    if not os.path.exists(filepath):
        return jsonify({'error': '文件不存在'}), 404
    
    # 创建唯一的输出目录
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_dir = os.path.join(app.config['OUTPUT_FOLDER'], timestamp)
    
    try:
        splitter = ExcelSplitter(filepath, split_column, output_dir)
        output_files = splitter.split_and_save()
        
        # 创建ZIP文件
        zip_filename = f"拆分结果_{timestamp}.zip"
        zip_path = os.path.join(app.config['OUTPUT_FOLDER'], zip_filename)
        
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for value, file_path in output_files.items():
                # 添加文件到ZIP，使用相对路径
                arcname = os.path.basename(file_path)
                zipf.write(file_path, arcname)
        
        # 清理临时文件
        shutil.rmtree(output_dir)
        
        return jsonify({
            'success': True,
            'download_url': f'/download/{zip_filename}',
            'file_count': len(output_files),
            'files': list(output_files.keys())
        })
    except Exception as e:
        # 清理可能的临时文件
        if os.path.exists(output_dir):
            shutil.rmtree(output_dir)
        return jsonify({'error': f'拆分失败: {str(e)}'}), 400


@app.route('/download/<filename>')
def download_file(filename):
    """下载拆分结果"""
    try:
        file_path = os.path.join(app.config['OUTPUT_FOLDER'], filename)
        if os.path.exists(file_path):
            return send_file(file_path, as_attachment=True, download_name=filename)
        else:
            return jsonify({'error': '文件不存在'}), 404
    except Exception as e:
        return jsonify({'error': f'下载失败: {str(e)}'}), 400


@app.route('/cleanup', methods=['POST'])
def cleanup():
    """清理临时文件"""
    data = request.json
    filename = data.get('filename')
    
    if filename:
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        if os.path.exists(filepath):
            os.remove(filepath)
    
    return jsonify({'success': True})


@app.route('/upload-multiple', methods=['POST'])
def upload_multiple_files():
    """处理多文件上传（用于合并）"""
    if 'files[]' not in request.files:
        return jsonify({'error': '没有文件'}), 400
    
    files = request.files.getlist('files[]')
    
    if len(files) == 0:
        return jsonify({'error': '未选择文件'}), 400
    
    if len(files) < 2:
        return jsonify({'error': '至少需要上传2个文件进行合并'}), 400
    
    uploaded_files = []
    all_sheets = set()
    
    try:
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        
        for idx, file in enumerate(files):
            if file.filename == '':
                continue
                
            if not allowed_file(file.filename):
                return jsonify({'error': f'文件 {file.filename} 格式不支持，只支持 .xlsx 和 .xls'}), 400
            
            # 保存文件
            filename = secure_filename(file.filename)
            unique_filename = f"{timestamp}_{idx}_{filename}"
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], unique_filename)
            file.save(filepath)
            
            # 读取sheet信息
            excel_file = pd.ExcelFile(filepath)
            sheets = excel_file.sheet_names
            all_sheets.update(sheets)
            
            uploaded_files.append({
                'original_name': file.filename,
                'saved_name': unique_filename,
                'sheets': sheets
            })
        
        return jsonify({
            'files': uploaded_files,
            'total_files': len(uploaded_files),
            'all_sheets': sorted(list(all_sheets))
        })
        
    except Exception as e:
        # 清理已上传的文件
        for file_info in uploaded_files:
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], file_info['saved_name'])
            if os.path.exists(filepath):
                os.remove(filepath)
        return jsonify({'error': f'上传失败: {str(e)}'}), 400


@app.route('/preview-merge', methods=['POST'])
def preview_merge():
    """预览合并结果"""
    data = request.json
    files = data.get('files', [])
    
    if not files or len(files) < 2:
        return jsonify({'error': '至少需要2个文件进行合并'}), 400
    
    file_paths = []
    for file_info in files:
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], file_info['saved_name'])
        if not os.path.exists(filepath):
            return jsonify({'error': f'文件 {file_info["original_name"]} 不存在'}), 404
        file_paths.append(filepath)
    
    try:
        merger = ExcelMerger(file_paths, "temp.xlsx")
        sheet_files = merger.get_all_sheets_info()
        
        # 统计每个sheet的数据
        preview_data = []
        for sheet_name, file_list in sorted(sheet_files.items()):
            sheet_info = {
                'sheet_name': sheet_name,
                'file_count': len(file_list),
                'files': [],
                'total_rows': 0
            }
            
            for file_path in file_list:
                try:
                    df = pd.read_excel(file_path, sheet_name=sheet_name)
                    row_count = len(df)
                    sheet_info['total_rows'] += row_count
                    sheet_info['files'].append({
                        'name': os.path.basename(file_path),
                        'rows': row_count
                    })
                except:
                    pass
            
            preview_data.append(sheet_info)
        
        return jsonify({
            'preview': preview_data,
            'total_sheets': len(sheet_files)
        })
        
    except Exception as e:
        return jsonify({'error': f'预览失败: {str(e)}'}), 400


@app.route('/merge', methods=['POST'])
def merge_files():
    """执行合并"""
    data = request.json
    files = data.get('files', [])
    
    if not files or len(files) < 2:
        return jsonify({'error': '至少需要2个文件进行合并'}), 400
    
    file_paths = []
    for file_info in files:
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], file_info['saved_name'])
        if not os.path.exists(filepath):
            return jsonify({'error': f'文件 {file_info["original_name"]} 不存在'}), 404
        file_paths.append(filepath)
    
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_filename = f"合并结果_{timestamp}.xlsx"
    output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)
    
    try:
        merger = ExcelMerger(file_paths, output_path)
        result_stats = merger.merge_and_save()
        
        return jsonify({
            'success': True,
            'download_url': f'/download/{output_filename}',
            'sheet_count': len(result_stats),
            'stats': result_stats
        })
        
    except Exception as e:
        # 清理可能的临时文件
        if os.path.exists(output_path):
            os.remove(output_path)
        return jsonify({'error': f'合并失败: {str(e)}'}), 400


@app.route('/cleanup-merge', methods=['POST'])
def cleanup_merge():
    """清理合并临时文件"""
    data = request.json
    files = data.get('files', [])
    
    for file_info in files:
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], file_info['saved_name'])
        if os.path.exists(filepath):
            os.remove(filepath)
    
    return jsonify({'success': True})


if __name__ == '__main__':
    # 获取端口（支持云端环境变量）
    port = int(os.environ.get('PORT', 5001))
    
    print("=" * 60)
    print("Excel拆分/合并工具已启动")
    print(f"拆分功能: http://0.0.0.0:{port}")
    print(f"合并功能: http://0.0.0.0:{port}/merger")
    print("=" * 60)
    app.run(debug=False, host='0.0.0.0', port=port)

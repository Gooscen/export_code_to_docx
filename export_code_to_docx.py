import os
import argparse
import re
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

# 默认支持的文件扩展名
DEFAULT_EXTENSIONS = [
    'py', 'js', 'ts', 'java', 'c', 'cpp', 'go', 
    'html', 'css', 'xml', 'json', 'yaml', 'yml',
    'sh', 'bash', 'sql', 'md', 'txt'
]

def is_comment_line(line, ext):
    line = line.strip()
    if not line:
        return False

    # 单行注释
    if ext in ['go', 'js', 'ts', 'java', 'c', 'cpp']:
        return line.startswith('//') or line.startswith('/*') or line.startswith('*') or line.startswith('*/')
    elif ext in ['py', 'sh', 'yaml', 'yml']:
        return line.startswith('#')
    elif ext in ['html', 'xml']:
        return '<!--' in line and '-->' in line
    return False

def remove_comments(content, ext):
    # 简单处理多行注释
    if ext in ['go', 'js', 'ts', 'java', 'c', 'cpp']:
        content = re.sub(r'/\*[\s\S]*?\*/', '', content)
    elif ext in ['html', 'xml']:
        content = re.sub(r'<!--[\s\S]*?-->', '', content)

    # 删除单行注释
    lines = content.splitlines()
    filtered = []
    for line in lines:
        if not is_comment_line(line, ext):
            filtered.append(line)
    return '\n'.join(filtered)

def add_code_to_docx(doc, file_path, show_filename=True, keep_comments=True):
    ext = file_path.split('.')[-1].lower()
    if show_filename:
        doc.add_heading(file_path, level=2)

    try:
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()
    except Exception as e:
        print(f"Error reading file {file_path}: {e}")
        return

    if not keep_comments:
        content = remove_comments(content, ext)

    paragraph = doc.add_paragraph()
    run = paragraph.add_run(content)
    run.font.name = 'Courier New'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Courier New')
    run.font.size = Pt(10)

def collect_valid_files(include_paths, exclude_paths, allowed_exts):
    included_files = []
    exclude_paths = set(os.path.abspath(p) for p in exclude_paths)

    for path in include_paths:
        abs_path = os.path.abspath(path)
        if os.path.isfile(abs_path):
            if should_include(abs_path, exclude_paths, allowed_exts):
                included_files.append(abs_path)
        else:
            for root, dirs, files in os.walk(abs_path):
                if any(os.path.abspath(root).startswith(e) for e in exclude_paths):
                    continue
                for file in files:
                    file_path = os.path.join(root, file)
                    if should_include(file_path, exclude_paths, allowed_exts):
                        included_files.append(file_path)
    return included_files

def should_include(file_path, exclude_paths, allowed_exts):
    if any(os.path.abspath(file_path).startswith(p) for p in exclude_paths):
        return False
    return file_path.lower().endswith(tuple('.' + ext.lower() for ext in allowed_exts))

def main():
    parser = argparse.ArgumentParser(description="将源代码导出为Word文档(.docx)")
    parser.add_argument('--include', nargs='+', required=True, help='要导出的文件或目录路径')
    parser.add_argument('--exclude', nargs='*', default=[], help='要排除的文件或目录路径')
    parser.add_argument('--ext', nargs='+', default=DEFAULT_EXTENSIONS, help=f'要包含的文件扩展名 (默认: {", ".join(DEFAULT_EXTENSIONS)})')
    parser.add_argument('--output', default='code_export.docx', help='输出文件名 (默认: code_export.docx)')
    parser.add_argument('--show-filename', action='store_true', default=False, help='是否显示文件名 (默认: 否)')
    parser.add_argument('--no-comments', action='store_true', default=False, help='是否删除注释 (默认: 否)')

    args = parser.parse_args()

    files = collect_valid_files(args.include, args.exclude, args.ext)
    print(f"将导出 {len(files)} 个文件到 {args.output}")

    doc = Document()
    # doc.add_heading('源码导出文档', level=1)

    for file_path in files:
        add_code_to_docx(doc, file_path, show_filename=args.show_filename, keep_comments=not args.no_comments)

    doc.save(args.output)
    print(f"导出成功: {args.output}")

if __name__ == "__main__":
    main()
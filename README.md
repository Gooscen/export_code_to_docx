# 📦 Export Code to Word (.docx)

这是一个用于将指定项目中的源代码文件（如 `.go`, `.html`, `.js` 等）导出为 Microsoft Word (`.docx`) 文件的 Python 工具，支持注释过滤、文件选择和中文路径。

---

## ✅ 功能特性

- 支持指定多个文件或目录作为导出来源；
- 支持指定多个文件扩展名（如 `go`, `html`, `js`, `py` 等）；
- 支持排除文件或目录；
- 支持保留或删除注释；
- 支持中文路径和文件名；
- 支持在导出文档中显示文件路径标题。

---

## 📦 安装依赖

你需要安装 `python-docx` 库：

```bash
pip install python-docx
```

## 🚀 使用方式
```bash
python export_code_to_docx.py \
  --include [包含的路径1 路径2 ...] \
  --exclude [排除的路径1 路径2 ...] \
  --ext [文件扩展名1 扩展名2 ...] \
  --output [导出的.docx文件名] \
  [--show-filename] \
  [--keep-comments]
```

### ⭐️使用示例
```bash
python export_code_to_docx.py \
  --include backend frontend config.yaml \            # 包含这些文件和文件夹
  --exclude frontend/static README.md \               # 排除静态资源和 README
  --ext go html js \                                  # 只处理 .go, .html, .js 文件
  --output output_code.docx \                         # 导出为 output_code.docx 文件
  --show-filename \                                   # 显示每个文件的路径标题
  # 注释掉下一行等于默认删除注释，如果要保留注释就加上这行：
  --keep-comments
  ```

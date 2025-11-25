# PDF 转 Word 神器

本项目是一个基于 PyQt5 的桌面应用，支持将 PDF 批量转换为 Word（`.docx`）。支持拖拽添加、文件夹扫描、进度条展示，转换结果与原文件同目录保存。

## 功能特性
- 拖拽文件/文件夹到窗口即可添加
- 选择单个或多个 PDF 文件
- 扫描文件夹并批量添加 PDF
- 转换进度与状态实时显示
- 转换后 `.docx` 与 PDF 同名、保存在原目录

## 环境要求
- Python 3.8+
- Windows / macOS / Linux（以 Windows 打包为主）

## 安装
```bash
pip install -r requirements.txt
```

## 运行
```bash
python PDF_2_Word.py
```

## 使用步骤
- 通过顶部按钮或拖拽添加 PDF 文件／文件夹
- 点击 `开始转换`
- 等待进度完成，`.docx` 将保存在对应 PDF 的目录下

## 打包为 Windows 可执行文件
- 项目已包含 `pyinstaller` 依赖（见 `requirements.txt`）
- 在项目根目录执行：
```bash
pyinstaller --noconsole --onefile --windowed --icon Pdf.png PDF_2_Word.py
```
- 生成的 `PDF_2_Word.exe` 位于 `dist/` 目录

## 依赖
- PyQt5>=5.15.0
- pdf2docx>=0.5.0
- pyinstaller

## 常见问题
- 扫描版（图片型）PDF 需先 OCR，否则生成的 Word 可编辑性有限
- 复杂排版（表格/公式）可能需要在 Word 中微调
- 请确保目标目录具备写入权限

## 作者
- 窗口标题标注：`作者:小庄-Python办公`（见 `PDF_2_Word.py`）

## 代码入口
- 主程序入口：`PDF_2_Word.py` 中的 `if __name__ == "__main__":` 段。


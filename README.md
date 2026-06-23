# FileToPDF - 文件转 PDF 工具

支持将 DOCX、DOC、WPS、PPTX、PPT、DPS 等格式文件转换为图片格式的 PDF，确保排版 100% 保真。

## 使用方法

### 单文件转换

```bash
python FileToPDF.py input.docx output.pdf --dpi 200
```

### 批量转换

```bash
python FileToPDF.py --batch --input-dir ./docx --output-dir ./pdf
```

### 添加水印

```bash
python FileToPDF.py input.docx output.pdf --watermark "机密" --watermark-pos tiled
```

## 命令行参数

| 参数 | 说明 | 默认值 |
|------|------|--------|
| `--dpi` | 输出分辨率 | 200 |
| `--watermark` | 水印文字 | None |
| `--watermark-pos` | 水印位置 | center |
| `--watermark-opacity` | 水印透明度 | 0.2 |
| `--watermark-font-size` | 水印字号 | 0（自动） |

## 依赖安装

```bash
pip install pywin32 pymupdf img2pdf pillow
```

## 原理

1. **WPS COM 导出 PDF**：使用 WPS 自身引擎渲染，确保格式保真
2. **PyMuPDF 逐页转图**：将 PDF 每页渲染为高分辨率图片
3. **img2pdf 打包**：将图片重新组合为 PDF 文件

## 支持格式

- 文字文档：DOCX、DOC、WPS
- 演示文稿：PPTX、PPT、DPS

## 注意事项

- 需要安装 WPS Office 或 Microsoft Office
- 仅支持 Windows 系统（依赖 COM 接口）
- 建议使用 DPI 200，平衡文件大小和清晰度
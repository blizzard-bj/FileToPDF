"""
文件转 PDF 工具（FileToPDF）
支持格式：DOCX、DOC、WPS、PPTX、PPT、DPS
原理：WPS COM 导出 PDF（保真渲染） → PyMuPDF 逐页转图 → img2pdf 打包
"""

import os
import sys
import time
import shutil
import tempfile
from pathlib import Path

import win32com.client
from pywintypes import com_error
import fitz            # PyMuPDF
import img2pdf
from PIL import Image, ImageDraw, ImageFont


class FileToPDF:
    """
    通过 WPS COM 自动化，将 Office 文件转换为图片格式的 PDF
    确保排版 100% 保真。
    """

    # ── WPS COM 常量 ──
    WD_EXPORT_FORMAT_PDF = 17         # wdExportFormatPDF
    WD_EXPORT_OPT_PUBLISHING = 1      # wdExportOptimizeForPrint（高质量）
    WD_EXPORT_RANGE_ALL = 0           # wdExportAllDocument
    WD_EXPORT_ITEM_DOCUMENT = 0       # wdExportDocumentContent
    WD_EXPORT_CREATE_NO_BOOKMARK = 0  # 不创建书签

    def __init__(self, dpi=200, watermark_text=None, watermark_pos='center', watermark_opacity=0.2, watermark_font_size=0):
        """
        参数：
            dpi: 输出图片的分辨率
                 150 — 屏幕阅读，文件较小
                 200 — 日常使用（推荐）
                 300 — 打印质量
                 400 — 高清打印
                 600 — 专业印刷
            watermark_text: 水印文字，None 表示不添加水印
            watermark_pos: 水印位置，可选值: 'center', 'top-left', 'top-right', 'bottom-left', 'bottom-right', 'tiled'
            watermark_opacity: 水印透明度，0.0-1.0，默认 0.2
            watermark_font_size: 水印字号，默认 0 表示自动计算
        """
        self.dpi = dpi
        self.watermark_text = watermark_text
        self.watermark_pos = watermark_pos
        self.watermark_opacity = watermark_opacity
        self.watermark_font_size = watermark_font_size
        self.wps = None
        self.doc = None

    def convert(self, input_docx: str, output_pdf: str) -> dict:
        """
        执行完整转换流程。
        参数：
            input_docx:  输入 docx 文件路径
            output_pdf:  输出图片 PDF 路径
        返回：
            dict: { pages, file_size_mb, output_path, dpi }
        """
        input_path = Path(input_docx).resolve()
        output_path = Path(output_pdf).resolve()

        if not input_path.exists():
            raise FileNotFoundError(f"输入文件不存在: {input_path}")

        if not input_path.suffix.lower() in ('.docx', '.doc', '.wps', '.pptx', '.ppt', '.dps'):
            raise ValueError(f"不支持的文件格式: {input_path.suffix}")

        # 创建临时目录
        tmpdir = Path(tempfile.mkdtemp(prefix="wps2imgpdf_"))
        temp_pdf = tmpdir / "wps_export.pdf"

        # 判断文件类型
        ppt_extensions = ('.pptx', '.ppt', '.dps')
        file_type = 'ppt' if input_path.suffix.lower() in ppt_extensions else 'doc'

        try:
            # ── 第一阶段：WPS COM 导出 PDF ──
            print("[1/4] 启动 WPS 并打开文档...")
            self._open_wps(input_path)

            print("[2/4] WPS 导出为 PDF（自身渲染，格式保真）...")
            self._export_pdf(temp_pdf, file_type=file_type)

            # 关闭文档和 WPS（尽早释放资源）
            print("[3/4] 关闭 WPS，开始转换图片...")
            self._close_wps()

            # ── 第二阶段：PDF → 图片 → 图片PDF ──
            page_count = self._pdf_to_image_pdf(temp_pdf, output_path)

            # 统计结果
            file_size = output_path.stat().st_size / (1024 * 1024)
            result = {
                "pages": page_count,
                "file_size_mb": round(file_size, 2),
                "output_path": str(output_path),
                "dpi": self.dpi,
            }

            print("[4/4] 转换完成！")
            print(f"      页数: {result['pages']}")
            print(f"      大小: {result['file_size_mb']} MB")
            print(f"      路径: {result['output_path']}")

            return result

        except Exception as e:
            self._close_wps()
            raise RuntimeError(f"转换失败: {e}") from e

        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)

    # ───────────────────────────────────────────
    #  WPS COM 操作
    # ───────────────────────────────────────────

    def _open_wps(self, file_path: Path):
        """启动 WPS 并打开文档"""
        file_url = str(file_path).replace("/", "\\")
        extension = file_path.suffix.lower()

        ppt_extensions = ('.pptx', '.ppt', '.dps')
        doc_extensions = ('.docx', '.doc', '.wps')

        if extension in ppt_extensions:
            # WPS 演示 COM 类名列表（按优先级排序）
            ppt_progids = [
                "kpp.Application",           # WPS 演示
                "WPS.Presentation",          # WPS 演示（另一种格式）
                "WPS.Application",           # WPS 通用（可能支持演示）
                "PowerPoint.Application",   # Microsoft PowerPoint
            ]

            self.wps = None
            for progid in ppt_progids:
                try:
                    self.wps = win32com.client.Dispatch(progid)
                    break
                except com_error:
                    continue

            if not self.wps:
                raise RuntimeError(
                    "无法启动演示文稿程序。请确保已安装 WPS 演示或 Microsoft PowerPoint。\n"
                    "尝试的 COM 组件: " + ", ".join(ppt_progids)
                )

            self.wps.Visible = True
            self.wps.DisplayAlerts = False
            self.doc = self.wps.Presentations.Open(
                FileName=file_url,
                ReadOnly=True,
                WithWindow=True,
            )
            # 保存文件路径供后续使用
            self._current_file = file_path
        elif extension in doc_extensions:
            doc_progids = [
                "kwps.Application",         # WPS 文字
                "WPS.Document",             # WPS 文字（另一种格式）
                "Word.Application",         # Microsoft Word
            ]

            self.wps = None
            for progid in doc_progids:
                try:
                    self.wps = win32com.client.Dispatch(progid)
                    break
                except com_error:
                    continue

            if not self.wps:
                raise RuntimeError(
                    "无法启动文字处理程序。请确保已安装 WPS 文字或 Microsoft Word。\n"
                    "尝试的 COM 组件: " + ", ".join(doc_progids)
                )

            self.wps.Visible = True
            self.wps.DisplayAlerts = False
            self.doc = self.wps.Documents.Open(
                FileName=file_url,
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
            )
            # 保存文件路径供后续使用
            self._current_file = file_path

        time.sleep(3)
        print(f"      已打开: {file_path.name}")

    def _export_pdf(self, pdf_path: Path, file_type: str = 'doc'):
        """将当前文档导出为 PDF（WPS 自身引擎渲染）"""
        import os
        import subprocess
        import time

        pdf_path_obj = Path(pdf_path)
        pdf_path_str = str(pdf_path_obj.resolve()).replace("/", "\\")
        pdf_dir = str(pdf_path_obj.parent.resolve()).replace("/", "\\")

        if file_type == 'ppt':
            # ── PPT 导出方式 ──

            # 方法1: 使用 WPS 打印功能（最可靠）
            try:
                # 获取 WPS 演示的路径
                wps_exe = None
                possible_paths = [
                    r"C:\Program Files\Kingsoft\WPS Office\11.1.0.11690\office6\wpp.exe",
                    r"C:\Program Files (x86)\Kingsoft\WPS Office\11.1.0.11690\office6\wpp.exe",
                    r"C:\Program Files\Kingsoft\WPS Office\office6\wpp.exe",
                    r"C:\Program Files (x86)\Kingsoft\WPS Office\office6\wpp.exe",
                ]
                for path in possible_paths:
                    if os.path.exists(path):
                        wps_exe = path
                        break

                if not wps_exe:
                    # 尝试从注册表获取路径
                    try:
                        import winreg
                        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\wpp.exe")
                        wps_exe, _ = winreg.QueryValueEx(key, "")
                        winreg.CloseKey(key)
                    except:
                        pass

                if wps_exe and os.path.exists(wps_exe):
                    # 使用 WPS 命令行参数打印到 PDF
                    # /t 表示打印，/p 表示输出打印机，/d 是输出路径
                    ppt_path = str(self._current_file.resolve()) if hasattr(self, '_current_file') else ""

                    # 先尝试直接导出
                    result = subprocess.run([
                        wps_exe,
                        pdf_path_str,
                        "/convert",  # 转换模式
                        "/pdf",       # 输出为 PDF
                    ], capture_output=True, text=True, timeout=120)

                    if pdf_path_obj.exists():
                        print(f"      PDF 导出成功（方法1: WPS命令行）")
                        return

                    # 如果命令行方式失败，尝试打印方式
                    print(f"      WPS 命令行方式未生成文件，尝试打印方式...")
            except Exception as e:
                print(f"      方法1失败: {str(e)[:50]}...")

            # 方法2: 使用 WPS 打印到 Microsoft Print to PDF
            try:
                # 获取当前文件路径
                if hasattr(self, '_current_file'):
                    ppt_file = str(self._current_file.resolve())
                else:
                    ppt_file = ""

                # 使用 PowerShell 调用 WPS 打印
                ps_script = f'''
$wps = New-Object -ComObject WPS.Application
$pres = $wps.Presentations.Open("{ppt_file.replace("\\", "\\\\")}", $true)
$pres.SaveCopyAs("{pdf_path_str.replace("\\", "\\\\")}", 32)  # ppSaveAsPDF = 32
$pres.Close()
$wps.Quit()
'''
                import tempfile
                ps_file = tempfile.NamedTemporaryFile(mode='w', suffix='.ps1', delete=False, encoding='utf-8')
                ps_file.write(ps_script)
                ps_file.close()

                result = subprocess.run([
                    'powershell', '-ExecutionPolicy', 'Bypass', '-File', ps_file.name
                ], capture_output=True, text=True, timeout=120)

                os.unlink(ps_file.name)

                if pdf_path_obj.exists():
                    print(f"      PDF 导出成功（方法2: PowerShell COM）")
                    return
            except Exception as e:
                print(f"      方法2失败: {str(e)[:50]}...")

            # 方法3: 直接使用 COM 对象的正确方法
            try:
                # WPS 演示文稿应该使用 SaveCopyAs 方法
                if hasattr(self.doc, 'SaveCopyAs'):
                    self.doc.SaveCopyAs(pdf_path_str, 32)  # ppSaveAsPDF = 32
                    time.sleep(1)
                    if pdf_path_obj.exists():
                        print(f"      PDF 导出成功（方法3: SaveCopyAs）")
                        return
                    print(f"      方法3执行成功但文件未生成")
            except Exception as e:
                print(f"      方法3失败: {str(e)[:50]}...")

            # 方法4: 尝试使用 Export 方法
            try:
                if hasattr(self.doc, 'Export'):
                    self.doc.Export(pdf_path_str, "PDF")
                    time.sleep(1)
                    if pdf_path_obj.exists():
                        print(f"      PDF 导出成功（方法4: Export）")
                        return
                    print(f"      方法4执行成功但文件未生成")
            except Exception as e:
                print(f"      方法4失败: {str(e)[:50]}...")

            # 方法5: SaveAs 不带格式参数
            try:
                self.doc.SaveAs(pdf_path_str)
                time.sleep(1)
                if pdf_path_obj.exists():
                    print(f"                      PDF 导出成功（方法5: SaveAs无参数）")
                    return
                print(f"      方法5执行成功但文件未生成")
            except Exception as e:
                print(f"      方法5失败: {str(e)[:50]}...")

            # 方法6: 尝试使用 WPS 的打印功能
            try:
                # 设置默认打印机为 Microsoft Print to PDF
                import win32print
                default_printer = win32print.GetDefaultPrinter()
                win32print.SetDefaultPrinter("Microsoft Print to PDF")

                # 打印演示文稿
                self.doc.PrintOut()

                # 等待打印完成
                time.sleep(3)

                # 恢复默认打印机
                win32print.SetDefaultPrinter(default_printer)

                # 检查是否有 PDF 生成
                output_files = list(Path(pdf_dir).glob("*.pdf"))
                if output_files:
                    # 移动到目标位置
                    latest_pdf = max(output_files, key=lambda p: p.stat().st_mtime)
                    shutil.move(str(latest_pdf), str(pdf_path))
                    print(f"      PDF 导出成功（方法6: 打印到PDF）")
                    return
            except Exception as e:
                print(f"      方法6失败: {str(e)[:50]}...")

        else:
            # ── DOC 导出方式 ──
            # 方法1: SaveAs
            try:
                self.doc.SaveAs(pdf_path_str, 17)
                if pdf_path_obj.exists():
                    print(f"      PDF 导出成功（方法1: SaveAs）")
                    return
                print(f"      方法1执行成功但文件未生成")
            except Exception as e:
                print(f"      方法1失败: {str(e)[:50]}...")

            # 方法2: ActiveDocument
            try:
                self.wps.ActiveDocument.SaveAs(pdf_path_str, 17)
                if pdf_path_obj.exists():
                    print(f"      PDF 导出成功（方法2: ActiveDocument.SaveAs）")
                    return
                print(f"      方法2执行成功但文件未生成")
            except Exception as e:
                print(f"      方法2失败: {str(e)[:50]}...")

            # 方法3: ExportAsFixedFormat
            try:
                self.doc.ExportAsFixedFormat(
                    OutputFileName=pdf_path_str,
                    ExportFormat=17,
                )
                if pdf_path_obj.exists():
                    print(f"      PDF 导出成功（方法3: ExportAsFixedFormat）")
                    return
                print(f"      方法3执行成功但文件未生成")
            except Exception as e:
                print(f"      方法3失败: {str(e)[:50]}...")

        raise RuntimeError(f"所有导出方法均失败，请检查 WPS/PowerPoint 安装")

    def _close_wps(self):
        """安全关闭文档和 WPS"""
        try:
            if self.doc:
                # 尝试多种关闭方式
                try:
                    self.doc.Close()
                except:
                    try:
                        self.doc.Close(SaveChanges=False)
                    except:
                        pass
                self.doc = None
        except com_error:
            pass
        try:
            if self.wps:
                self.wps.Quit()
                self.wps = None
        except com_error:
            pass

    def _add_watermark(self, image: Image.Image) -> Image.Image:
        """
        为图片添加文字水印（优化版本）
        """
        if not self.watermark_text:
            return image

        width, height = image.size

        try:
            if self.watermark_font_size > 0:
                font_size = self.watermark_font_size
            else:
                font_size = int(min(width, height) / 12)
            font = ImageFont.truetype('simhei.ttf', font_size)
        except:
            try:
                if self.watermark_font_size > 0:
                    font_size = self.watermark_font_size
                else:
                    font_size = int(min(width, height) / 12)
                font = ImageFont.truetype('msyh.ttc', font_size)
            except:
                font = ImageFont.load_default()
                font_size = int(min(width, height) / 15)

        text_bbox = ImageDraw.Draw(image).textbbox((0, 0), self.watermark_text, font)
        text_width = text_bbox[2] - text_bbox[0]
        text_height = text_bbox[3] - text_bbox[1]

        # 水印颜色 (RGBA)
        opacity = int(255 * self.watermark_opacity)
        color = (128, 128, 128, opacity)

        if self.watermark_pos == 'tiled':
            # ── 优化的平铺水印 ──
            # 1. 创建一个小尺寸的平铺图案
            # 2. 缩放到目标图片大小
            # 3. 叠加到原图

            angle = -45  # 倾斜角度

            # 计算平铺图案的尺寸（倾斜后需要更大的画布）
            tile_width = text_width * 4
            tile_height = text_height * 4

            # 创建单个平铺单元
            tile = Image.new('RGBA', (tile_width, tile_height), (0, 0, 0, 0))
            tile_draw = ImageDraw.Draw(tile)

            # 计算文字起始位置（居中）
            start_x = (tile_width - text_width) // 2
            start_y = (tile_height - text_height) // 2

            # 绘制文字
            tile_draw.text((start_x, start_y), self.watermark_text, font=font, fill=color)

            # 旋转平铺单元
            rotated_tile = tile.rotate(angle, expand=True, resample=Image.BICUBIC)

            # 创建完整水印层（与目标图片同尺寸）
            watermark_layer = Image.new('RGBA', (width, height), (0, 0, 0, 0))

            # 平铺旋转后的图案
            pattern_width, pattern_height = rotated_tile.size

            # 使用更高效的平铺方式
            for y in range(0, height + pattern_height, pattern_height):
                for x in range(0, width + pattern_width, pattern_width):
                    watermark_layer.paste(rotated_tile, (x, y), rotated_tile)

            # 将水印叠加到原图
            if image.mode != 'RGBA':
                image = image.convert('RGBA')

            image = Image.alpha_composite(image, watermark_layer)
            image = image.convert('RGB')

        else:
            # ── 单位置水印 ──
            draw = ImageDraw.Draw(image)

            if self.watermark_pos == 'center':
                x = (width - text_width) // 2
                y = (height - text_height) // 2
            elif self.watermark_pos == 'top-left':
                x = 50
                y = 50
            elif self.watermark_pos == 'top-right':
                x = width - text_width - 50
                y = 50
            elif self.watermark_pos == 'bottom-left':
                x = 50
                y = height - text_height - 50
            elif self.watermark_pos == 'bottom-right':
                x = width - text_width - 50
                y = height - text_height - 50
            else:
                x = (width - text_width) // 2
                y = (height - text_height) // 2

            draw.text((x, y), self.watermark_text, font=font, fill=(128, 128, 128, opacity))

        if image.mode != 'RGB':
            image = image.convert('RGB')

        return image

    # ───────────────────────────────────────────
    #  PDF → 图片 → 图片PDF
    # ───────────────────────────────────────────

    def _pdf_to_image_pdf(self, temp_pdf: Path, output_pdf: Path) -> int:
        """
        将 WPS 导出的文本 PDF 转换为图片格式 PDF。
        此步不涉及排版，仅做像素渲染。
        """
        pdf_doc = fitz.open(str(temp_pdf))
        total_pages = len(pdf_doc)
        print(f"      PDF 共 {total_pages} 页，逐页渲染中...")

        tmpdir = temp_pdf.parent
        image_paths = []
        page_sizes = []

        for page_idx in range(total_pages):
            page = pdf_doc[page_idx]

            # DPI → 缩放矩阵（PDF原始 72 DPI）
            zoom = self.dpi / 72.0
            matrix = fitz.Matrix(zoom, zoom)

            # 渲染为 RGB 像素图
            pixmap = page.get_pixmap(matrix=matrix, alpha=False)

            # 获取页面实际尺寸（点）
            rect = page.rect
            page_sizes.append((rect.width, rect.height))

            # 保存为 PNG
            img_path = tmpdir / f"page_{page_idx + 1:04d}.png"
            pixmap.save(str(img_path))

            # 添加水印
            if self.watermark_text:
                img = Image.open(str(img_path))
                img = self._add_watermark(img)
                img.save(str(img_path))

            image_paths.append(str(img_path))

            progress = (page_idx + 1) / total_pages * 100
            print(f"      渲染: {page_idx + 1}/{total_pages} ({progress:.0f}%)", end="\r")

        pdf_doc.close()
        print()

        # ── PNG → 图片PDF ──
        print(f"      正在组装图片 PDF...")

        # 使用第一页尺寸作为 PDF 页面尺寸
        if page_sizes:
            layout_fun = img2pdf.get_layout_fun(page_sizes[0])
        else:
            # 降级到 A4 尺寸
            a4_pt = (img2pdf.mm_to_pt(210), img2pdf.mm_to_pt(297))
            layout_fun = img2pdf.get_layout_fun(a4_pt)

        with open(str(output_pdf), "wb") as f:
            f.write(img2pdf.convert(
                image_paths,
                layout_fun=layout_fun,
                dpi=self.dpi,
            ))

        return total_pages


# ───────────────────────────────────────────
#  命令行入口
# ───────────────────────────────────────────

def batch_convert(input_dir: str, output_dir: str, dpi=200, watermark_text=None, 
                  watermark_pos='center', watermark_opacity=0.2, watermark_font_size=0) -> dict:
    """
    批量转换目录中的所有 docx/doc/wps 文件
    """
    input_path = Path(input_dir).resolve()
    output_path = Path(output_dir).resolve()

    if not input_path.exists():
        raise FileNotFoundError(f"输入目录不存在: {input_path}")

    if not output_path.exists():
        output_path.mkdir(parents=True, exist_ok=True)

    supported_extensions = ('.docx', '.doc', '.wps', '.pptx', '.ppt', '.dps')
    files = sorted(input_path.glob('*'))
    docx_files = [f for f in files if f.suffix.lower() in supported_extensions]

    if not docx_files:
        raise ValueError(f"输入目录中没有找到支持的文件格式: {supported_extensions}")

    print(f"发现 {len(docx_files)} 个待转换文件")
    print("-" * 50)

    results = {
        "total_files": len(docx_files),
        "success": 0,
        "failed": 0,
        "details": [],
    }

    converter = FileToPDF(
        dpi=dpi,
        watermark_text=watermark_text,
        watermark_pos=watermark_pos,
        watermark_opacity=watermark_opacity,
        watermark_font_size=watermark_font_size,
    )

    for idx, input_file in enumerate(docx_files, 1):
        output_file = output_path / f"{input_file.stem}.pdf"
        print(f"[{idx}/{len(docx_files)}] 转换: {input_file.name}")

        try:
            result = converter.convert(str(input_file), str(output_file))
            results["success"] += 1
            results["details"].append({
                "file": input_file.name,
                "status": "success",
                "pages": result["pages"],
                "size_mb": result["file_size_mb"],
            })
        except Exception as e:
            results["failed"] += 1
            results["details"].append({
                "file": input_file.name,
                "status": "failed",
                "error": str(e),
            })
            print(f"      ❌ 失败: {e}")

    return results


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="文件转 PDF 工具（支持 DOCX/DOC/WPS/PPTX/PPT/DPS）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  单文件转换:
    python FileToPDF.py input.docx output.pdf
    python FileToPDF.py input.docx output.pdf --dpi 300

  添加水印:
    python FileToPDF.py input.docx output.pdf --watermark "机密"
    python FileToPDF.py input.docx output.pdf --watermark "内部文件" --watermark-pos tiled --watermark-opacity 0.3

  批量转换:
    python FileToPDF.py --batch --input-dir ./docx --output-dir ./pdf
    python FileToPDF.py --batch --input-dir ./docx --output-dir ./pdf --dpi 300 --watermark "机密"
        """
    )

    parser.add_argument("--batch", action="store_true", help="启用批量转换模式")
    parser.add_argument("--input-dir", help="批量转换时的输入目录（需要配合 --batch）")
    parser.add_argument("--output-dir", help="批量转换时的输出目录（需要配合 --batch）")

    parser.add_argument("input", nargs='?', help="输入 docx/doc/wps/pptx/ppt/dps 文件路径（单文件模式）")
    parser.add_argument("output", nargs='?', help="输出 PDF 文件路径（单文件模式）")

    parser.add_argument(
        "--dpi",
        type=int,
        default=200,
        choices=[150, 200, 300, 400, 600],
        help="输出分辨率: 150(屏幕) / 200(推荐) / 300(打印) / 400(高清) / 600(专业)，默认 200"
    )

    parser.add_argument("--watermark", help="水印文字，如：机密、内部文件")
    parser.add_argument(
        "--watermark-pos",
        default='center',
        choices=['center', 'top-left', 'top-right', 'bottom-left', 'bottom-right', 'tiled'],
        help="水印位置: center(居中) / top-left(左上) / top-right(右上) / bottom-left(左下) / bottom-right(右下) / tiled(平铺)，默认 center"
    )
    parser.add_argument(
        "--watermark-opacity",
        type=float,
        default=0.2,
        help="水印透明度，0.0-1.0，默认 0.2"
    )
    parser.add_argument(
        "--watermark-font-size",
        type=int,
        default=0,
        help="水印字号，默认 0 表示自动计算"
    )

    args = parser.parse_args()

    print("=" * 60)
    print("文件转 PDF 工具（FileToPDF）")
    print("=" * 60)

    if args.batch:
        if not args.input_dir or not args.output_dir:
            parser.error("批量模式需要同时指定 --input-dir 和 --output-dir")

        print(f"模式: 批量转换")
        print(f"输入目录: {args.input_dir}")
        print(f"输出目录: {args.output_dir}")
        print(f"DPI: {args.dpi}")
        if args.watermark:
            print(f"水印: {args.watermark}")
            print(f"水印位置: {args.watermark_pos}")
            print(f"水印透明度: {args.watermark_opacity}")
            if args.watermark_font_size > 0:
                print(f"水印字号: {args.watermark_font_size}")
        print("-" * 60)

        try:
            results = batch_convert(
                args.input_dir,
                args.output_dir,
                dpi=args.dpi,
                watermark_text=args.watermark,
                watermark_pos=args.watermark_pos,
                watermark_opacity=args.watermark_opacity,
                watermark_font_size=args.watermark_font_size,
            )

            print("\n" + "=" * 60)
            print("📊 批量转换结果")
            print("=" * 60)
            print(f"总文件数: {results['total_files']}")
            print(f"成功: {results['success']}")
            print(f"失败: {results['failed']}")

            if results['failed'] > 0:
                print("\n❌ 失败文件列表:")
                for detail in results['details']:
                    if detail['status'] == 'failed':
                        print(f"  - {detail['file']}: {detail['error']}")

        except Exception as e:
            print("\n" + "=" * 60)
            print("❌ 批量转换失败")
            print("=" * 60)
            print(f"错误信息: {e}")
            import traceback
            traceback.print_exc()

    else:
        if not args.input or not args.output:
            parser.error("单文件模式需要指定 input 和 output 参数")

        print(f"模式: 单文件转换")
        print(f"输入文件: {args.input}")
        print(f"输出文件: {args.output}")
        print(f"DPI: {args.dpi}")
        if args.watermark:
            print(f"水印: {args.watermark}")
            print(f"水印位置: {args.watermark_pos}")
            print(f"水印透明度: {args.watermark_opacity}")
            if args.watermark_font_size > 0:
                print(f"水印字号: {args.watermark_font_size}")
        print("-" * 60)

        try:
            converter = FileToPDF(
                dpi=args.dpi,
                watermark_text=args.watermark,
                watermark_pos=args.watermark_pos,
                watermark_opacity=args.watermark_opacity,
                watermark_font_size=args.watermark_font_size,
            )
            result = converter.convert(args.input, args.output)

            print("\n" + "=" * 60)
            print("✅ 转换成功！")
            print("=" * 60)
            print(f"页数: {result['pages']}")
            print(f"文件大小: {result['file_size_mb']} MB")
            print(f"输出路径: {result['output_path']}")
            print(f"DPI: {result['dpi']}")

        except Exception as e:
            print("\n" + "=" * 60)
            print("❌ 转换失败")
            print("=" * 60)
            print(f"错误信息: {e}")
            import traceback
            traceback.print_exc()
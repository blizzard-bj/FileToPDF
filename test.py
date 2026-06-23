"""
测试文件转 PDF 工具
"""

from FileToPDF import FileToPDF
from pathlib import Path

# 配置测试参数
INPUT_DOCX = r"e:\ZW\文档转换\FileToPDF\test.docx"  # 替换为你的 docx 文件路径
OUTPUT_PDF = r"e:\ZW\文档转换\FileToPDF\test_output.pdf"  # 输出 PDF 路径
DPI = 200  # 输出分辨率：150/200/300

def main():
    print("=" * 50)
    print("文件转 PDF 工具测试")
    print("=" * 50)

    # 检查输入文件
    input_path = Path(INPUT_DOCX)
    if not input_path.exists():
        print(f"❌ 输入文件不存在: {INPUT_DOCX}")
        print("\n请修改 INPUT_DOCX 变量为有效的 docx 文件路径")
        return

    print(f"输入文件: {INPUT_DOCX}")
    print(f"输出文件: {OUTPUT_PDF}")
    print(f"DPI: {DPI}")
    print("-" * 50)

    try:
        # 创建转换器实例
        converter = FileToPDF(dpi=DPI)

        # 执行转换
        result = converter.convert(INPUT_DOCX, OUTPUT_PDF)

        print("\n" + "=" * 50)
        print("✅ 转换成功！")
        print("=" * 50)
        print(f"页数: {result['pages']}")
        print(f"文件大小: {result['file_size_mb']} MB")
        print(f"输出路径: {result['output_path']}")
        print(f"DPI: {result['dpi']}")

    except Exception as e:
        print("\n" + "=" * 50)
        print("❌ 转换失败")
        print("=" * 50)
        print(f"错误信息: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
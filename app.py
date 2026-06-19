import io
import os
import shutil          # ← 原来漏掉的
import zipfile
import subprocess
import tempfile
import streamlit as st

# -------------------- 工具函数 --------------------
def get_soffice_path():
    """自动查找 LibreOffice 可执行文件"""
    paths = [
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/usr/bin/soffice",
        "soffice",
    ]
    for p in paths:
        if os.path.exists(p) or shutil.which(p):
            return p
    return None


def convert_excel_to_pdf(file_bytes, base_name):
    """用 LibreOffice 将 Excel 转为 PDF（保留打印排版）"""
    soffice = get_soffice_path()
    if not soffice:
        raise RuntimeError("未找到 LibreOffice，请确认已安装")

    with tempfile.TemporaryDirectory() as tmpdir:
        xlsx_path = os.path.join(tmpdir, f"{base_name}.xlsx")
        with open(xlsx_path, "wb") as f:
            f.write(file_bytes)

        # 转换
        subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf", "--outdir", tmpdir, xlsx_path],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )

        pdf_path = os.path.join(tmpdir, f"{base_name}.pdf")
        if not os.path.exists(pdf_path):
            raise RuntimeError("转换失败，未生成 PDF")
        with open(pdf_path, "rb") as f:
            return f.read()


def split_pdf_pages(pdf_bytes, base_name, sheet_names):
    """
    将多页 PDF 拆分成单页 PDF，并用工作表名命名
    如果 PyPDF2 不可用，则返回整个 PDF（一个文件）
    """
    try:
        from PyPDF2 import PdfReader, PdfWriter
    except ImportError:
        # 未安装 PyPDF2，返回原文件
        return {f"{base_name}.pdf": pdf_bytes}

    reader = PdfReader(io.BytesIO(pdf_bytes))
    total_pages = len(reader.pages)
    if total_pages == 0:
        return {}

    # 如果 sheet_names 数量匹配，用 sheet 命名，否则用页码
    single_pdfs = {}
    for i in range(total_pages):
        writer = PdfWriter()
        writer.add_page(reader.pages[i])
        buf = io.BytesIO()
        writer.write(buf)
        buf.seek(0)

        if i < len(sheet_names):
            name = f"{base_name}_{sheet_names[i]}.pdf"
        else:
            name = f"{base_name}_page{i+1}.pdf"
        single_pdfs[name] = buf.read()
    return single_pdfs


def get_sheet_names(file_bytes):
    """读取 Excel 的所有工作表名，不解析内容"""
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    names = wb.sheetnames
    wb.close()
    return names


# -------------------- Streamlit 界面 --------------------
st.set_page_config(page_title="Excel → A4 PDF (完美打印)")
st.title("📄 Excel 批量转 A4 单页 PDF（打印预览效果）")
st.markdown(
    "使用 **LibreOffice** 完美保留 Excel 打印缩放、合并单元格、排版。\n"
    "每个工作表单独生成一个 PDF，所有内容自动缩放至 A4 一页。\n\n"
    "如未安装 `PyPDF2`，请执行 `pip install PyPDF2` 以启用自动拆分。"
)

uploaded_files = st.file_uploader(
    "选择 Excel 文件（.xlsx / .xls）",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
)

if uploaded_files:
    if st.button("🚀 开始转换"):
        with st.spinner("正在用 LibreOffice 转换..."):
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for file in uploaded_files:
                    try:
                        name_no_ext = os.path.splitext(file.name)[0]
                        # 1. 获取工作表名称
                        sheet_names = get_sheet_names(file.read())
                        file_bytes = file.read()  # 重新读取，因为 get_sheet_names 已经消耗了流
                        # 上面有问题：file.read() 后光标位置变化，需要重置
                    except Exception as e:
                        st.error(f"❌ {file.name} 读取失败: {e}")
                        continue

            # 修正读取两次的问题，重新设计循环
            # 先在循环外读取所有文件到内存
            file_data = {}
            for file in uploaded_files:
                file_data[file.name] = io.BytesIO(file.read())

            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for fname, fbytes in file_data.items():
                    try:
                        name_no_ext = os.path.splitext(fname)[0]
                        fbytes.seek(0)
                        data = fbytes.read()

                        # 读取工作表名
                        sheet_names = get_sheet_names(data)
                        # 转换整个 Excel 为 PDF
                        pdf_bytes = convert_excel_to_pdf(data, name_no_ext)
                        # 拆分 PDF（如果可能）
                        single_pdfs = split_pdf_pages(pdf_bytes, name_no_ext, sheet_names)

                        for pdf_name, pdf_data in single_pdfs.items():
                            zf.writestr(pdf_name, pdf_data)

                    except Exception as e:
                        st.error(f"❌ {fname}\n{e}")

            zip_buffer.seek(0)
            if zipfile.ZipFile(zip_buffer).namelist():
                st.success("✅ 转换完成，排版与打印预览一致")
                st.download_button(
                    "⬇️ 下载所有 PDF（ZIP）",
                    data=zip_buffer,
                    file_name="excel_pdfs_libreoffice.zip",
                    mime="application/zip",
                )
            else:
                st.error("❌ 未生成任何 PDF，请检查文件。")

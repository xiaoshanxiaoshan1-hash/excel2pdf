import io
import os
import zipfile
import subprocess
import tempfile
import shutil
import streamlit as st

def convert_with_libreoffice(file_bytes, base_name):
    """使用 LibreOffice 命令行将 Excel 转为 PDF（保留原格式）"""
    with tempfile.TemporaryDirectory() as tmpdir:
        xlsx_path = os.path.join(tmpdir, f"{base_name}.xlsx")
        with open(xlsx_path, "wb") as f:
            f.write(file_bytes)
        # 调用 soffice 转换为 PDF，输出到临时目录
        subprocess.run([
            "soffice",
            "--headless",
            "--convert-to", "pdf",
            "--outdir", tmpdir,
            xlsx_path
        ], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        pdf_path = os.path.join(tmpdir, f"{base_name}.pdf")
        if not os.path.exists(pdf_path):
            raise RuntimeError("LibreOffice 转换失败，未生成 PDF")
        with open(pdf_path, "rb") as f:
            pdf_bytes = f.read()
    return pdf_bytes

def split_pdf_pages(pdf_bytes):
    """将一个多页 PDF 拆分成多个单页 PDF 字节流（每个工作表一页）"""
    from PyPDF2 import PdfReader, PdfWriter
    reader = PdfReader(io.BytesIO(pdf_bytes))
    pages = []
    for i, page in enumerate(reader.pages):
        writer = PdfWriter()
        writer.add_page(page)
        buf = io.BytesIO()
        writer.write(buf)
        buf.seek(0)
        pages.append((f"page_{i+1}.pdf", buf.read()))
    return pages

# Streamlit 界面
st.set_page_config(page_title="Excel → A4 PDF (完美格式)")
st.title("📄 Excel 转 A4 PDF（保留原始排版）")
st.markdown("依赖 **LibreOffice** 实现完美格式转换，一个工作表生成一个 PDF。")

uploaded_files = st.file_uploader("选择 Excel 文件", type=["xlsx","xls"], accept_multiple_files=True)

if uploaded_files:
    if st.button("🚀 转换"):
        with st.spinner("正在调用 LibreOffice 转换…"):
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for file in uploaded_files:
                    try:
                        name_no_ext = os.path.splitext(file.name)[0]
                        pdf_bytes = convert_with_libreoffice(file.read(), name_no_ext)
                        # 拆分多页
                        pages = split_pdf_pages(pdf_bytes)
                        for page_name, page_data in pages:
                            zf.writestr(f"{name_no_ext}_{page_name}", page_data)
                    except Exception as e:
                        st.error(f"❌ {file.name}: {e}")
            zip_buffer.seek(0)
            if zipfile.ZipFile(zip_buffer).namelist():
                st.success("转换完成")
                st.download_button("⬇️ 下载 ZIP", data=zip_buffer, file_name="excel_pdfs_libre.zip", mime="application/zip")
            else:
                st.error("无有效输出")

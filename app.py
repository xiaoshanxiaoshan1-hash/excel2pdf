import io
import os
import shutil
import zipfile
import subprocess
import tempfile
import streamlit as st

# -------------------- LibreOffice 路径 --------------------
def get_soffice_path():
    paths = [
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",  # macOS
        "/usr/bin/soffice",                                     # Linux
        "soffice",                                              # 已在 PATH
    ]
    for p in paths:
        if os.path.exists(p) or shutil.which(p):
            return p
    return None


def convert_with_libreoffice(file_bytes, base_name):
    soffice = get_soffice_path()
    if not soffice:
        raise RuntimeError("❌ 服务器未安装 LibreOffice，请联系管理员")
    with tempfile.TemporaryDirectory() as tmpdir:
        xlsx_path = os.path.join(tmpdir, f"{base_name}.xlsx")
        with open(xlsx_path, "wb") as f:
            f.write(file_bytes)
        subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf", "--outdir", tmpdir, xlsx_path],
            check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
        )
        pdf_path = os.path.join(tmpdir, f"{base_name}.pdf")
        if not os.path.exists(pdf_path):
            raise RuntimeError("转换失败，未生成 PDF")
        with open(pdf_path, "rb") as f:
            return f.read()


def get_sheet_names(file_bytes):
    """用 openpyxl 获取工作表名（只读，不解析内容）"""
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    names = wb.sheetnames
    wb.close()
    return names


def split_pdf_pages(pdf_bytes, base_name, sheet_names):
    """尝试拆分多页 PDF，拆不开则直接返回原文件"""
    try:
        from PyPDF2 import PdfReader, PdfWriter
    except ImportError:
        return {f"{base_name}.pdf": pdf_bytes}

    reader = PdfReader(io.BytesIO(pdf_bytes))
    total = len(reader.pages)
    if total == 0:
        return {}
    pdfs = {}
    for i in range(total):
        writer = PdfWriter()
        writer.add_page(reader.pages[i])
        buf = io.BytesIO()
        writer.write(buf)
        buf.seek(0)
        name = f"{base_name}_{sheet_names[i]}.pdf" if i < len(sheet_names) else f"{base_name}_page{i+1}.pdf"
        pdfs[name] = buf.read()
    return pdfs


# -------------------- Streamlit 界面 --------------------
st.set_page_config(page_title="Excel → A4 PDF")
st.title("📄 Excel 批量转 A4 单页 PDF（打印预览效果）")
st.markdown("依赖 LibreOffice 完美还原排版，每个工作表一个 PDF。")

uploaded = st.file_uploader("选择 Excel 文件", type=["xlsx","xls"], accept_multiple_files=True)

if uploaded and st.button("🚀 开始转换"):
    with st.spinner("转换中…"):
        zip_buf = io.BytesIO()
        file_data = {f.name: io.BytesIO(f.read()) for f in uploaded}

        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for fname, buf in file_data.items():
                try:
                    buf.seek(0)
                    raw = buf.read()
                    name_no_ext = os.path.splitext(fname)[0]
                    sheet_names = get_sheet_names(raw)
                    pdf_bytes = convert_with_libreoffice(raw, name_no_ext)
                    single_pdfs = split_pdf_pages(pdf_bytes, name_no_ext, sheet_names)
                    for pname, pdata in single_pdfs.items():
                        zf.writestr(pname, pdata)
                except Exception as e:
                    st.error(f"❌ {fname}\n{e}")

        zip_buf.seek(0)
        if zipfile.ZipFile(zip_buf).namelist():
            st.success("✅ 完成，排版与打印预览一致")
            st.download_button("⬇️ 下载 ZIP", data=zip_buf, file_name="pdfs.zip", mime="application/zip")
        else:
            st.error("❌ 未生成任何 PDF")

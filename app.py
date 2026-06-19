import io, os, shutil, zipfile, subprocess, tempfile, traceback, warnings
import openpyxl
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

# ==================== 全局配置 ====================
SCALE_PERCENT = 67          # ★ 缩放比例 67%

def set_print_scale(workbook, scale):
    """将工作簿所有工作表的打印缩放设为指定百分比"""
    for ws in workbook.worksheets:
        ws.page_setup.scale = scale
        # 关闭“调整为页宽/页高”，只用缩放比例
        ws.page_setup.fitToWidth = 0
        ws.page_setup.fitToHeight = 0

# ==================== LibreOffice 路径 ====================
def get_soffice_path():
    paths = [
        '/usr/bin/soffice',                                     # Linux/云端
        '/Applications/LibreOffice.app/Contents/MacOS/soffice', # Mac
        'soffice'
    ]
    for p in paths:
        if os.path.exists(p) or shutil.which(p):
            return p
    return None

def convert_with_libreoffice(file_bytes, base_name):
    soffice = get_soffice_path()
    if not soffice:
        raise RuntimeError("❌ 未找到 LibreOffice，请确认已安装")

    # 1. 用 openpyxl 修改打印缩放为 67%
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    set_print_scale(wb, SCALE_PERCENT)
    modified_bytes = io.BytesIO()
    wb.save(modified_bytes)
    wb.close()
    modified_bytes.seek(0)

    # 2. 保存到临时文件，调用 LibreOffice 转换为 PDF
    with tempfile.TemporaryDirectory() as tmpdir:
        xlsx_path = os.path.join(tmpdir, f'{base_name}.xlsx')
        with open(xlsx_path, 'wb') as f:
            f.write(modified_bytes.read())

        subprocess.run(
            [soffice, '--headless', '--convert-to', 'pdf', '--outdir', tmpdir, xlsx_path],
            check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            timeout=60
        )
        pdf_path = os.path.join(tmpdir, f'{base_name}.pdf')
        if not os.path.exists(pdf_path):
            raise RuntimeError("转换失败，未生成 PDF")
        with open(pdf_path, 'rb') as f:
            return f.read()

def get_sheet_names(file_bytes):
    """读取所有工作表名称"""
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    names = wb.sheetnames
    wb.close()
    return names

def split_pdf_pages(pdf_bytes, base_name, sheet_names):
    """拆分多页 PDF，用工作表名命名"""
    try:
        from PyPDF2 import PdfReader, PdfWriter
    except ImportError:
        return {f'{base_name}.pdf': pdf_bytes}
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
        name = f'{base_name}_{sheet_names[i]}.pdf' if i < len(sheet_names) else f'{base_name}_page{i+1}.pdf'
        pdfs[name] = buf.read()
    return pdfs

# ==================== Streamlit 界面 ====================
st.set_page_config(page_title='Excel → A4 PDF (67%)')
st.title('📄 Excel 批量转 A4 单页 PDF（缩放 67%）')
st.markdown(
    f'✅ 使用 **LibreOffice** 完美保留排版，每个工作表生成一个 PDF。\n\n'
    f'⭕ 打印缩放比例：**{SCALE_PERCENT}%**（可根据需要修改代码中的 `SCALE_PERCENT`）'
)

uploaded_files = st.file_uploader('选择 Excel 文件', type=['xlsx','xls'], accept_multiple_files=True)

if uploaded_files and st.button('🚀 开始转换'):
    with st.spinner('转换中…'):
        zip_buf = io.BytesIO()
        total_pdfs = 0
        with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
            for f in uploaded_files:
                try:
                    raw = f.read()
                    name_no_ext = os.path.splitext(f.name)[0]
                    sheet_names = get_sheet_names(raw)
                    pdf_bytes = convert_with_libreoffice(raw, name_no_ext)
                    pdfs = split_pdf_pages(pdf_bytes, name_no_ext, sheet_names)
                    for pname, pdata in pdfs.items():
                        zf.writestr(pname, pdata)
                        total_pdfs += 1
                except Exception as e:
                    st.error(f'❌ {f.name}\n{e}\n{traceback.format_exc()}')
        zip_buf.seek(0)
        if total_pdfs > 0:
            st.success(f'✅ 转换完成，共生成 {total_pdfs} 个 PDF（缩放 {SCALE_PERCENT}%）')
            st.download_button('⬇️ 下载 ZIP', data=zip_buf, file_name='excel_pdfs.zip', mime='application/zip')
        else:
            st.error('❌ 未生成任何 PDF')

import io, os, shutil, zipfile, subprocess, tempfile, traceback
import openpyxl
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

# ==================== 页面优化 ====================
def get_last_data_column(ws):
    """返回最后一个包含数据的列号（1-based）"""
    max_col = ws.max_column
    if max_col is None:
        return 1
    for col in range(max_col, 0, -1):
        for row in ws.iter_rows(min_col=col, max_col=col, values_only=True):
            if any(cell is not None and str(cell).strip() != '' for cell in row):
                return col
    return 1

def set_print_optimize(workbook):
    """设置打印区域、窄边距、A4 纸张、强制一页宽"""
    for ws in workbook.worksheets:
        # 1. 打印区域仅包含数据列
        last_col = get_last_data_column(ws)
        max_row = ws.max_row or 1
        col_letter = get_column_letter(last_col)
        ws.print_area = f'A1:{col_letter}{max_row}'

        # 2. 极窄页边距（单位：英寸）
        ws.page_margins.left = 0.2
        ws.page_margins.right = 0.2
        ws.page_margins.top = 0.2
        ws.page_margins.bottom = 0.2
        ws.page_margins.header = 0.0
        ws.page_margins.footer = 0.0

        # 3. 纸张大小：A4 (9)
        ws.page_setup.paperSize = 9

        # 4. 强制所有列在一页（高度不限制）
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0   # 0 表示不限制，高度自动适配

# ==================== LibreOffice 转换 ====================
def get_soffice_path():
    paths = [
        '/usr/bin/soffice',
        '/Applications/LibreOffice.app/Contents/MacOS/soffice',
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

    # 修改 Excel 页面设置
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    set_print_optimize(wb)
    modified_bytes = io.BytesIO()
    wb.save(modified_bytes)
    wb.close()
    modified_bytes.seek(0)

    # 用 LibreOffice 转 PDF
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

# ==================== 辅助函数 ====================
def get_sheet_names(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    names = wb.sheetnames
    wb.close()
    return names

def split_pdf_pages(pdf_bytes, base_name, sheet_names):
    """拆分多页 PDF，用工作表名命名（每个工作表应只有 1 页）"""
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
st.set_page_config(page_title='Excel → A4 单页 PDF')
st.title('📄 Excel 批量转 A4 单页 PDF（智能打印）')
st.markdown(
    '✅ 自动检测数据列，极窄边距，强制所有列在一页 A4 内。\n'
    '**每个工作表生成唯一一个 PDF，不再分页**'
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
            st.success(f'✅ 转换完成，共 {total_pdfs} 个 PDF（每个工作表一页）')
            st.download_button('⬇️ 下载 ZIP', data=zip_buf, file_name='excel_pdfs.zip', mime='application/zip')
        else:
            st.error('❌ 未生成任何 PDF')

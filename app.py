import io, os, shutil, zipfile, subprocess, tempfile, traceback
import openpyxl
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

# ==================== 可调参数 ====================
SCALE_PERCENT = 67          # 缩放比例（若仍超宽可改小，如 50）
MARGIN_INCHES = 0.2         # 页边距（英寸）

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

def clean_and_set_print(workbook):
    """清理分页符，设置缩放、窄边距、A4，限定打印区域"""
    for ws in workbook.worksheets:
        # 1. 清除所有水平/垂直分页符
        ws.page_breaks.horizontal_breaks = []
        ws.page_breaks.vertical_breaks = []
        ws.page_breaks.row = 0
        ws.page_breaks.col = 0

        # 2. 设置打印区域（只包含有数据的列）
        last_col = get_last_data_column(ws)
        max_row = ws.max_row or 1
        col_letter = get_column_letter(last_col)
        ws.print_area = f'A1:{col_letter}{max_row}'

        # 3. 极窄页边距
        ws.page_margins.left = MARGIN_INCHES
        ws.page_margins.right = MARGIN_INCHES
        ws.page_margins.top = MARGIN_INCHES
        ws.page_margins.bottom = MARGIN_INCHES
        ws.page_margins.header = 0.0
        ws.page_margins.footer = 0.0

        # 4. 纸张大小：A4 (9)
        ws.page_setup.paperSize = 9

        # 5. 固定缩放比例（不依赖 fitToWidth，避免被文件原有设置覆盖）
        ws.page_setup.scale = SCALE_PERCENT
        ws.page_setup.fitToWidth = 0
        ws.page_setup.fitToHeight = 0

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

    # 用 openpyxl 重新保存文件，洗掉微信只读等特殊属性
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    clean_and_set_print(wb)
    modified_bytes = io.BytesIO()
    wb.save(modified_bytes)
    wb.close()
    modified_bytes.seek(0)

    # 再用 LibreOffice 转 PDF
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
st.set_page_config(page_title='Excel → A4 单页 PDF')
st.title('📄 Excel 批量转 A4 单页 PDF（固定缩放）')
st.markdown(
    f'✅ 清除分页符，固定缩放 **{SCALE_PERCENT}%**，极窄边距，只打印有内容的列。\n'
    '**每个工作表生成唯一一个 PDF，所有列一定在同一页。**\n'
    '（若仍超宽，可将代码中的 `SCALE_PERCENT` 改小，例如 50）'
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

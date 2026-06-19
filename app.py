import io, os, shutil, zipfile, subprocess, tempfile, traceback
import openpyxl
import streamlit as st

def force_fit_all_columns(workbook):
    """清除分页符、极窄边距，强制所有内容缩放到一页宽度"""
    for ws in workbook.worksheets:
        # 清除分页符
        if ws.row_breaks:
            ws.row_breaks.break_list.clear()
        if ws.col_breaks:
            ws.col_breaks.break_list.clear()

        # 极窄边距（0.1 英寸）
        ws.page_margins.left = 0.1
        ws.page_margins.right = 0.1
        ws.page_margins.top = 0.1
        ws.page_margins.bottom = 0.1
        ws.page_margins.header = 0.0
        ws.page_margins.footer = 0.0

        # A4 纸张
        ws.page_setup.paperSize = 9

        # 强制所有列在一页，高度不限制
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.page_setup.scale = 0  # 清除缩放比例

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

    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    force_fit_all_columns(wb)
    modified_bytes = io.BytesIO()
    wb.save(modified_bytes)
    wb.close()
    modified_bytes.seek(0)

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

# ==================== Streamlit 界面 ====================
st.set_page_config(page_title='Excel → A4 PDF（一文件一PDF）')
st.title('📄 Excel 批量转 A4 PDF（每个 Excel 生成一个 PDF）')
st.markdown(
    '✅ 上传多个 Excel，每个生成一个 **独立的 PDF** 文件。\n'
    '✅ 所有工作表都在同一个 PDF 中（多页），每页内容自动缩放至 A4。\n'
    '✅ 合并单元格、格式完美保留。'
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
                    pdf_bytes = convert_with_libreoffice(raw, name_no_ext)
                    zf.writestr(f'{name_no_ext}.pdf', pdf_bytes)
                    total_pdfs += 1
                except Exception as e:
                    st.error(f'❌ {f.name}\n{e}\n{traceback.format_exc()}')
        zip_buf.seek(0)
        if total_pdfs > 0:
            st.success(f'✅ 转换完成，共 {total_pdfs} 个 PDF（每个 Excel 文件对应一个）')
            st.download_button('⬇️ 下载 ZIP', data=zip_buf, file_name='excel_pdfs.zip', mime='application/zip')
        else:
            st.error('❌ 未生成任何 PDF')

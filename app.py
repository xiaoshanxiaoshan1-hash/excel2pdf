import io, os, shutil, zipfile, subprocess, tempfile, traceback
import openpyxl
import streamlit as st

def apply_fixed_scaling(workbook, scale_percent):
    """只使用固定缩放，关闭自适应，窄边距，清除分页符"""
    for ws in workbook.worksheets:
        # 1. 清空所有分页符
        if ws.row_breaks:
            ws.row_breaks.break_list.clear()
        if ws.col_breaks:
            ws.col_breaks.break_list.clear()

        # 2. 极窄边距（0.1 英寸 ≈ 2.5mm）
        ws.page_margins.left = 0.1
        ws.page_margins.right = 0.1
        ws.page_margins.top = 0.1
        ws.page_margins.bottom = 0.1
        ws.page_margins.header = 0.0
        ws.page_margins.footer = 0.0

        # 3. 纸张 A4
        ws.page_setup.paperSize = 9

        # 4. ⚠️ 关键：完全关闭自适应，只用固定缩放
        ws.page_setup.fitToWidth = 0
        ws.page_setup.fitToHeight = 0
        ws.page_setup.scale = scale_percent

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

def convert_with_libreoffice(file_bytes, base_name, scale):
    soffice = get_soffice_path()
    if not soffice:
        raise RuntimeError("❌ 未找到 LibreOffice，请确认已安装")

    # 1. 修改 Excel 打印设置
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    apply_fixed_scaling(wb, scale)
    modified_bytes = io.BytesIO()
    wb.save(modified_bytes)
    wb.close()
    modified_bytes.seek(0)

    # 2. 用 LibreOffice 转 PDF
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
st.set_page_config(page_title='Excel → A4 PDF（固定缩放）')
st.title('📄 Excel 批量转 A4 PDF（缩放可调）')
st.markdown('每个 Excel 生成一个 PDF（多工作表合成一个文件），**所有列缩放至一页**。')

# 让同事自己拖拽缩放比例（默认 67%）
scale = st.slider('缩放比例 (%)', min_value=30, max_value=100, value=67, step=1,
                  help='数字越小，内容越小，越容易装下所有列。若表格很宽可调至 50% 或更低')

uploaded_files = st.file_uploader('选择 Excel 文件', type=['xlsx','xls'], accept_multiple_files=True)

if uploaded_files and st.button('🚀 开始转换'):
    with st.spinner(f'正在以 {scale}% 缩放转换…'):
        zip_buf = io.BytesIO()
        total_pdfs = 0
        with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
            for f in uploaded_files:
                try:
                    raw = f.read()
                    name_no_ext = os.path.splitext(f.name)[0]
                    pdf_bytes = convert_with_libreoffice(raw, name_no_ext, scale)
                    zf.writestr(f'{name_no_ext}.pdf', pdf_bytes)
                    total_pdfs += 1
                except Exception as e:
                    st.error(f'❌ {f.name}\n{e}\n{traceback.format_exc()}')
        zip_buf.seek(0)
        if total_pdfs > 0:
            st.success(f'✅ 转换完成，共 {total_pdfs} 个 PDF')
            st.download_button('⬇️ 下载 ZIP', data=zip_buf, file_name='excel_pdfs.zip', mime='application/zip')
        else:
            st.error('❌ 未生成任何 PDF')

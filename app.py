import io, os, shutil, zipfile, subprocess, tempfile, traceback
import openpyxl
import streamlit as st

# ==================== 三语翻译 ====================
LANG = {
    'zh': {
        'title': '📄 Excel 批量转 A4 PDF',
        'desc': '每个 Excel 生成一个 PDF，内容居中，所有列缩放至一页。',
        'scale_label': '缩放比例 (%)',
        'scale_help': '输入 30-100，表格越宽数值越小',
        'upload_label': '选择 Excel 文件',
        'button': '🚀 开始转换',
        'clear_button': '🗑️ 一键清除所有文件',
        'spinner': '正在以 {scale}% 缩放转换…',
        'success': '✅ 转换完成，共 {count} 个 PDF',
        'download': '⬇️ 下载 ZIP',
        'error_no_file': '❌ 未生成任何 PDF',
        'libre_error': '❌ 未找到 LibreOffice，请确认已安装',
        'convert_error': '转换失败，未生成 PDF',
    },
    'en': {
        'title': '📄 Excel to A4 PDF Batch Converter',
        'desc': 'Each Excel generates one PDF, content centered, all columns scaled to one page.',
        'scale_label': 'Scale (%)',
        'scale_help': 'Enter 30-100, smaller for wider sheets',
        'upload_label': 'Choose Excel files',
        'button': '🚀 Convert',
        'clear_button': '🗑️ Clear All Files',
        'spinner': 'Converting at {scale}% scale…',
        'success': '✅ Done! {count} PDF(s) generated.',
        'download': '⬇️ Download ZIP',
        'error_no_file': '❌ No PDF generated',
        'libre_error': '❌ LibreOffice not found, please install it.',
        'convert_error': 'Conversion failed, no PDF generated.',
    },
    'th': {
        'title': '📄 แปลง Excel เป็น PDF A4 แบบทีละไฟล์',
        'desc': 'แต่ละไฟล์ Excel สร้าง PDF หนึ่งไฟล์ จัดกึ่งกลาง คอลัมน์ทั้งหมดอยู่ในหนึ่งหน้า',
        'scale_label': 'เปอร์เซ็นต์การย่อ (%)',
        'scale_help': 'ป้อน 30-100, ตารางกว้างใช้ค่าน้อย',
        'upload_label': 'เลือกไฟล์ Excel',
        'button': '🚀 เริ่มแปลง',
        'clear_button': '🗑️ ล้างไฟล์ทั้งหมด',
        'spinner': 'กำลังแปลงที่ {scale}%…',
        'success': '✅ เสร็จสิ้น สร้าง PDF ทั้งหมด {count} ไฟล์',
        'download': '⬇️ ดาวน์โหลด ZIP',
        'error_no_file': '❌ ไม่มี PDF ถูกสร้าง',
        'libre_error': '❌ ไม่พบ LibreOffice กรุณาติดตั้ง',
        'convert_error': 'การแปลงล้มเหลว ไม่มี PDF ถูกสร้าง',
    }
}

# ==================== 页面配置 ====================
st.set_page_config(page_title='Excel to A4 PDF', layout='centered')

# ==================== 水印 ====================
st.markdown("""
<style>
.watermark {
    position: fixed;
    left: 15px;
    bottom: 15px;
    font-size: 12px;
    color: rgba(150, 150, 150, 0.5);
    z-index: 9999;
    font-family: Arial, sans-serif;
}
</style>
<div class="watermark">MADEBYCX</div>
""", unsafe_allow_html=True)

# ==================== 核心功能 ====================
def apply_fixed_scaling(workbook, scale_percent):
    for ws in workbook.worksheets:
        if ws.row_breaks:
            ws.row_breaks.break_list.clear()
        if ws.col_breaks:
            ws.col_breaks.break_list.clear()

        ws.page_margins.left = 0.3
        ws.page_margins.right = 0.3
        ws.page_margins.top = 0.3
        ws.page_margins.bottom = 0.3
        ws.page_margins.header = 0.0
        ws.page_margins.footer = 0.0

        ws.page_setup.horizontalCentered = True
        ws.page_setup.verticalCentered = True
        ws.page_setup.paperSize = 9

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
        raise RuntimeError(LANG[lang]['libre_error'])

    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    apply_fixed_scaling(wb, scale)
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
            raise RuntimeError(LANG[lang]['convert_error'])
        with open(pdf_path, 'rb') as f:
            return f.read()

# ==================== 界面 ====================
lang = st.selectbox('🌐 语言 / Language / ภาษา', ['zh', 'en', 'th'], index=0)
t = LANG[lang]

st.title(t['title'])
st.markdown(t['desc'])

scale = st.number_input(
    t['scale_label'],
    min_value=30,
    max_value=100,
    value=67,
    step=1,
    help=t['scale_help']
)

# ---------- 文件上传（带清除功能） ----------
if 'clear_counter' not in st.session_state:
    st.session_state.clear_counter = 0

# 清除按钮：点击后递增计数器，从而改变 file_uploader 的 key，清空文件列表
if st.button(t['clear_button']):
    st.session_state.clear_counter += 1
    st.rerun()

# file_uploader 的 key 随计数器变化，实现清空
uploaded_files = st.file_uploader(
    t['upload_label'],
    type=['xlsx', 'xls'],
    accept_multiple_files=True,
    key=f"file_uploader_{st.session_state.clear_counter}"
)

# ---------- 转换 ----------
if uploaded_files and st.button(t['button']):
    with st.spinner(t['spinner'].format(scale=scale)):
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
            st.success(t['success'].format(count=total_pdfs))
            st.download_button(t['download'], data=zip_buf, file_name='excel_pdfs.zip', mime='application/zip')
        else:
            st.error(t['error_no_file'])

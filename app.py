import io
import os
import zipfile
import subprocess
import tempfile
import streamlit as st

# 检测 LibreOffice 可执行文件路径
def get_soffice_path():
    # 常见路径列表
    paths = [
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",  # macOS Homebrew cask
        "/usr/bin/soffice",
        "soffice",  # 如果已在 PATH 里
    ]
    for p in paths:
        if os.path.exists(p) or shutil.which(p):
            return p
    return None

def convert_with_libreoffice(file_bytes, base_name):
    soffice = get_soffice_path()
    if not soffice:
        raise RuntimeError("未找到 LibreOffice，请安装后重试")

    with tempfile.TemporaryDirectory() as tmpdir:
        xlsx_path = os.path.join(tmpdir, f"{base_name}.xlsx")
        with open(xlsx_path, "wb") as f:
            f.write(file_bytes)

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


# Streamlit 界面
st.set_page_config(page_title="Excel → A4 PDF (完美打印)")
st.title("📄 Excel 批量转 A4 单页 PDF（打印预览效果）")
st.markdown("依赖 LibreOffice，自动检测安装路径。")

uploaded_files = st.file_uploader("选择 Excel 文件", type=["xlsx","xls"], accept_multiple_files=True)

if uploaded_files:
    if st.button("🚀 开始转换"):
        with st.spinner("转换中…"):
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for file in uploaded_files:
                    try:
                        name = os.path.splitext(file.name)[0]
                        pdf_bytes = convert_with_libreoffice(file.read(), name)
                        zf.writestr(f"{name}.pdf", pdf_bytes)
                    except Exception as e:
                        st.error(f"❌ {file.name}: {e}")
            zip_buffer.seek(0)
            if zipfile.ZipFile(zip_buffer).namelist():
                st.success("✅ 完成，排版与打印预览一致")
                st.download_button("⬇️ 下载 ZIP", data=zip_buffer, file_name="pdfs.zip", mime="application/zip")
            else:
                st.error("未生成任何 PDF")

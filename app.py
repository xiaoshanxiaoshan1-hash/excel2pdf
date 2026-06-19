import io
import os
import zipfile
import subprocess
import tempfile
import shutil
import streamlit as st

def convert_with_libreoffice(file_bytes, base_name):
    """用 LibreOffice 将 Excel 转为 PDF（保留页面设置、合并单元格）"""
    with tempfile.TemporaryDirectory() as tmpdir:
        xlsx_path = os.path.join(tmpdir, f"{base_name}.xlsx")
        with open(xlsx_path, "wb") as f:
            f.write(file_bytes)

        # 调用 LibreOffice 无头模式转换
        subprocess.run(
            [
                "soffice",
                "--headless",
                "--convert-to", "pdf",
                "--outdir", tmpdir,
                xlsx_path
            ],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )

        pdf_path = os.path.join(tmpdir, f"{base_name}.pdf")
        if not os.path.exists(pdf_path):
            raise RuntimeError("LibreOffice 转换失败，未生成 PDF")
        with open(pdf_path, "rb") as f:
            pdf_bytes = f.read()
    return pdf_bytes


# Streamlit 界面
st.set_page_config(page_title="Excel → A4 PDF (完美打印)")
st.title("📄 Excel 批量转 A4 单页 PDF（打印预览效果）")
st.markdown(
    "依赖 **LibreOffice** 实现与 Excel 打印预览完全一致的排版。\n\n"
    "**使用前请确认已安装 LibreOffice**，安装命令：\n"
    "`brew install --cask libreoffice`"
)

uploaded_files = st.file_uploader(
    "选择 Excel 文件（.xlsx / .xls）",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
)

if uploaded_files:
    if st.button("🚀 开始转换"):
        with st.spinner("正在用 LibreOffice 转换…"):
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for file in uploaded_files:
                    try:
                        name_no_ext = os.path.splitext(file.name)[0]
                        pdf_bytes = convert_with_libreoffice(file.read(), name_no_ext)
                        zf.writestr(f"{name_no_ext}.pdf", pdf_bytes)
                    except Exception as e:
                        st.error(f"❌ {file.name}\n{e}")
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
                st.error("❌ 未生成任何 PDF，请检查文件是否正确。")

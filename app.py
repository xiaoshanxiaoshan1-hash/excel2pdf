import io
import os
import zipfile
import traceback
import pandas as pd
import matplotlib.pyplot as plt
import img2pdf
from PIL import Image
import streamlit as st

A4_WIDTH_PT  = 595.276
A4_HEIGHT_PT = 841.890

def sheet_to_pdf(df):
    raw = df.values.tolist()
    data = [[str(cell) for cell in row] for row in raw]

    nrows = len(data)
    ncols = len(data[0]) if nrows > 0 else 0

    font_size_w = (8.27 * 72) / ncols / 0.5
    font_size_h = (11.69 * 72) / nrows / 1.0
    font_size = min(font_size_w, font_size_h, 10)
    font_size = max(font_size, 4)

    fig, ax = plt.subplots(figsize=(8.27, 11.69))
    ax.axis("off")
    fig.subplots_adjust(left=0.05, right=0.95, bottom=0.05, top=0.95)

    table = ax.table(cellText=data, bbox=[0, 0, 1, 1])
    table.auto_set_font_size(False)
    table.set_fontsize(font_size)

    for key, cell in table.get_celld().items():
        cell.set_width(1.0 / ncols)
        cell.set_height(1.0 / nrows)
        cell.set_linewidth(0)
        cell.set_edgecolor("none")
        cell.set_text_props(ha="center", va="center")

    buf = io.BytesIO()
    try:
        fig.savefig(buf, format="png", dpi=200, pad_inches=0, bbox_inches="tight")
    finally:
        plt.close(fig)
    buf.seek(0)

    image = Image.open(buf)
    if image.mode in ("RGBA", "LA", "P"):
        rgb_image = Image.new("RGB", image.size, (255, 255, 255))
        rgb_image.paste(image, mask=image.split()[-1] if image.mode == "RGBA" else None)
        image = rgb_image

    img_bytes = io.BytesIO()
    image.save(img_bytes, format="PNG")
    img_bytes.seek(0)

    layout = img2pdf.get_layout_fun(pagesize=(A4_WIDTH_PT, A4_HEIGHT_PT))
    pdf_bytes = img2pdf.convert(img_bytes.getvalue(), layout_fun=layout)
    return pdf_bytes


def excel_to_pdfs(file_bytes, base_name):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    sheet_names = xls.sheet_names
    pdfs = {}
    skipped = []

    for sheet in sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet, header=None,
                           dtype=str, keep_default_na=False)
        if df.empty or df.size == 0:
            skipped.append(sheet)
            continue
        try:
            pdf_bytes = sheet_to_pdf(df)
        except Exception as e:
            raise RuntimeError(f"工作表 “{sheet}” 转换失败:\n{e}\n{traceback.format_exc()}")
        pdf_name = f"{base_name}_{sheet}.pdf" if len(sheet_names) > 1 else f"{base_name}.pdf"
        pdfs[pdf_name] = pdf_bytes
    return pdfs, skipped


st.set_page_config(page_title="Excel → A4 PDF")
st.title("📄 Excel 批量转 A4 单页 PDF（无网格线）")
st.markdown("上传 Excel 文件，自动缩放至 A4 打印。**无表格框线**，干净整洁。")

uploaded_files = st.file_uploader("选择 Excel 文件", type=["xlsx","xls"], accept_multiple_files=True)

if uploaded_files:
    if st.button("🚀 转换"):
        with st.spinner("转换中…"):
            zip_buffer = io.BytesIO()
            all_skipped = {}
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for file in uploaded_files:
                    try:
                        name_no_ext = os.path.splitext(file.name)[0]
                        pdfs, skipped = excel_to_pdfs(file.read(), name_no_ext)
                        for pname, pdata in pdfs.items():
                            zf.writestr(pname, pdata)
                        if skipped:
                            all_skipped[file.name] = skipped
                    except Exception as e:
                        st.error(f"❌ {file.name}\n{e}")
            zip_buffer.seek(0)
            if all_skipped:
                msg = "⚠️ 跳过的空工作表：\n" + "\n".join([f"{k}: {v}" for k,v in all_skipped.items()])
                st.warning(msg)
            if zipfile.ZipFile(zip_buffer).namelist():
                st.success("完成")
                st.download_button("⬇️ 下载 ZIP", data=zip_buffer, file_name="excel_pdfs.zip", mime="application/zip")
            else:
                st.error("无内容可转换")

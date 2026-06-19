import io
import os
import zipfile
import pandas as pd
import matplotlib.pyplot as plt
import img2pdf
from PIL import Image
import streamlit as st

def sheet_to_pdf(df):
    """将单个 DataFrame 转换为 A4 单页 PDF 的字节流"""
    data = df.astype(str).values.tolist()
    nrows = len(data)
    ncols = len(data[0]) if nrows > 0 else 0

    if nrows == 0 or ncols == 0:
        # 空表 → 空白 A4 PDF
        return img2pdf.convert(
            [], layout_fun=img2pdf.get_layout_fun(pagesize="A4")
        )

    # 根据行列数自动估算字体大小（点）
    font_size_w = (8.27 * 72) / ncols / 0.6
    font_size_h = (11.69 * 72) / nrows / 1.2
    font_size = min(font_size_w, font_size_h, 12)
    font_size = max(font_size, 4)

    fig, ax = plt.subplots(figsize=(8.27, 11.69))
    ax.axis("off")
    fig.subplots_adjust(left=0, right=1, bottom=0, top=1)

    # 表格铺满整个 axes
    table = ax.table(cellText=data, bbox=[0, 0, 1, 1])
    table.auto_set_font_size(False)
    table.set_fontsize(font_size)

    # 强制单元格等宽等高，均匀分布
    for key, cell in table.get_celld().items():
        cell.set_width(1.0 / ncols)
        cell.set_height(1.0 / nrows)
        cell.set_linewidth(0.3)
        cell.set_text_props(ha="center", va="center")

    # 先输出为高分辨率 PNG
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=200, pad_inches=0, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)

    # 转为 A4 PDF（自动缩放填满整页）
    image = Image.open(buf)
    if image.mode in ("RGBA", "LA", "P"):
        rgb_image = Image.new("RGB", image.size, (255, 255, 255))
        rgb_image.paste(image, mask=image.split()[-1] if image.mode == "RGBA" else None)
        image = rgb_image

    img_bytes = io.BytesIO()
    image.save(img_bytes, format="PNG")
    img_bytes.seek(0)

    pdf_bytes = img2pdf.convert(
        img_bytes.getvalue(),
        layout_fun=img2pdf.get_layout_fun(pagesize="A4"),
    )
    return pdf_bytes


def excel_to_pdfs(file_bytes, base_name):
    """将一个 Excel 文件的所有工作表分别转为 PDF，返回 {文件名: PDF字节}"""
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    sheet_names = xls.sheet_names
    pdfs = {}

    for sheet in sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet, header=None, dtype=str, keep_default_na=False)
        pdf_bytes = sheet_to_pdf(df)

        if len(sheet_names) == 1:
            pdf_name = f"{base_name}.pdf"
        else:
            pdf_name = f"{base_name}_{sheet}.pdf"
        pdfs[pdf_name] = pdf_bytes

    return pdfs


# -------------------- Streamlit 界面 --------------------
st.set_page_config(page_title="Excel → A4 PDF")
st.title("📄 Excel 批量转 A4 单页 PDF")
st.markdown(
    "上传多个 Excel 文件，每个工作表会单独生成一个 **A4 单页 PDF**，"
    "内容自动缩放，适合打印。"
)

uploaded_files = st.file_uploader(
    "选择 Excel 文件（.xlsx / .xls）",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
)

if uploaded_files:
    if st.button("🚀 开始转换"):
        with st.spinner("转换中…"):
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for file in uploaded_files:
                    try:
                        name_no_ext = os.path.splitext(file.name)[0]
                        pdfs = excel_to_pdfs(file.read(), name_no_ext)
                        for pdf_name, pdf_data in pdfs.items():
                            zf.writestr(pdf_name, pdf_data)
                    except Exception as e:
                        st.error(f"❌ {file.name} 转换失败: {e}")
            zip_buffer.seek(0)
            st.success("✅ 转换完成！")
            st.download_button(
                "⬇️ 下载所有 PDF（ZIP）",
                data=zip_buffer,
                file_name="excel_pdfs.zip",
                mime="application/zip",
            )

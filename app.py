import io
import os
import zipfile
import traceback
import pandas as pd
import matplotlib.pyplot as plt
import img2pdf
from PIL import Image
import streamlit as st

def sheet_to_pdf(df):
    """将单个 DataFrame 转换为 A4 单页 PDF 的字节流"""
    # ★ 确保所有单元格内容都是纯粹的字符串，避免类型混用
    raw = df.values.tolist()
    data = [[str(cell) for cell in row] for row in raw]

    nrows = len(data)
    ncols = len(data[0]) if nrows > 0 else 0

    if nrows == 0 or ncols == 0:
        return img2pdf.convert(
            [], layout_fun=img2pdf.get_layout_fun(pagesize="A4")
        )

    # 自适应字体大小
    font_size_w = (8.27 * 72) / ncols / 0.6
    font_size_h = (11.69 * 72) / nrows / 1.2
    font_size = min(font_size_w, font_size_h, 12)
    font_size = max(font_size, 4)

    fig, ax = plt.subplots(figsize=(8.27, 11.69))
    ax.axis("off")
    fig.subplots_adjust(left=0, right=1, bottom=0, top=1)

    table = ax.table(cellText=data, bbox=[0, 0, 1, 1])
    table.auto_set_font_size(False)
    table.set_fontsize(font_size)

    # 均匀分布行列
    for key, cell in table.get_celld().items():
        cell.set_width(1.0 / ncols)
        cell.set_height(1.0 / nrows)
        cell.set_linewidth(0.3)
        cell.set_text_props(ha="center", va="center")

    # 渲染为高分辨率 PNG
    buf = io.BytesIO()
    try:
        fig.savefig(buf, format="png", dpi=200, pad_inches=0, bbox_inches="tight")
    finally:
        plt.close(fig)
    buf.seek(0)

    # PNG → A4 PDF
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
    """一个 Excel 文件 → 多个 PDF（一个工作表一个文件）"""
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    sheet_names = xls.sheet_names
    pdfs = {}

    for sheet in sheet_names:
        # header=None 保证不丢失首行；dtype=str 把数据当作文本
        df = pd.read_excel(
            xls, sheet_name=sheet, header=None, dtype=str, keep_default_na=False
        )
        try:
            pdf_bytes = sheet_to_pdf(df)
        except Exception as e:
            # 输出完整错误栈，便于定位问题
            err_msg = f"{e}\n{traceback.format_exc()}"
            raise RuntimeError(f"工作表 “{sheet}” 转换失败:\n{err_msg}")

        pdf_name = f"{base_name}_{sheet}.pdf" if len(sheet_names) > 1 else f"{base_name}.pdf"
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
                        st.error(f"❌ {file.name}\n{e}")
                        # 即使某个文件失败也继续处理剩余文件
            zip_buffer.seek(0)
            st.success("✅ 转换完成（有错误的文件已跳过）")
            st.download_button(
                "⬇️ 下载所有 PDF（ZIP）",
                data=zip_buffer,
                file_name="excel_pdfs.zip",
                mime="application/zip",
            )

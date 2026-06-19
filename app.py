import io
import os
import zipfile
import traceback
import pandas as pd
import matplotlib.pyplot as plt
import img2pdf
from PIL import Image
import streamlit as st

# A4 尺寸（点）
A4_WIDTH_PT  = 595.276
A4_HEIGHT_PT = 841.890

# 字体度量系数（经验值）
CHAR_WIDTH_FACTOR = 0.5   # 每个字符约占字体大小的 0.5 倍宽度（点）
LINE_HEIGHT_FACTOR = 1.4  # 每行高度约为字体大小的 1.4 倍

def sheet_to_pdf(df):
    """
    将 DataFrame 转换为 A4 单页 PDF，内容自适应缩放，完整显示
    """
    raw = df.values.tolist()
    data = [[str(cell) for cell in row] for row in raw]
    nrows = len(data)
    ncols = len(data[0]) if nrows > 0 else 0

    # --- 1. 分析每列最大字符数 ---
    max_chars_per_col = [0] * ncols
    for row in data:
        for c in range(ncols):
            # 中文字符视觉宽度约等于英文字符的2倍，这里简单用 len() 统计字符数
            char_count = len(row[c])
            if char_count > max_chars_per_col[c]:
                max_chars_per_col[c] = char_count

    # --- 2. 根据内容长度计算最大允许字体大小 ---
    # 宽度约束：所有列的宽度之和 <= A4 宽度（减去少量边距）
    margin_pt = 20  # 左右各留 10pt 边距
    avail_width = A4_WIDTH_PT - margin_pt
    total_char_weight = sum(max_chars_per_col)
    if total_char_weight > 0:
        font_size_from_width = avail_width / total_char_weight / CHAR_WIDTH_FACTOR
    else:
        font_size_from_width = 12

    # 高度约束：所有行的高度之和 <= A4 高度（减去边距）
    avail_height = A4_HEIGHT_PT - margin_pt
    font_size_from_height = avail_height / nrows / LINE_HEIGHT_FACTOR

    # 取最小值，并限制在合理范围
    font_size = min(font_size_from_width, font_size_from_height, 10)  # 最大10pt
    font_size = max(font_size, 3)  # 最小3pt，保证可读性

    # --- 3. 计算每列的宽度比例（根据内容长度） ---
    col_widths = [chars / total_char_weight if total_char_weight > 0 else 1.0/ncols
                  for chars in max_chars_per_col]
    # 归一化，使总和为1
    col_widths = [w / sum(col_widths) for w in col_widths]

    # --- 4. 绘制表格 ---
    fig, ax = plt.subplots(figsize=(8.27, 11.69))
    ax.axis("off")
    # 留出边距，保证内容不被切
    fig.subplots_adjust(left=0.03, right=0.97, bottom=0.03, top=0.97)

    table = ax.table(cellText=data, bbox=[0, 0, 1, 1])
    table.auto_set_font_size(False)
    table.set_fontsize(font_size)

    # 为每个单元格设置宽度（按列比例）和统一高度
    for row_idx in range(nrows):
        for col_idx in range(ncols):
            cell = table[row_idx, col_idx]
            cell.set_width(col_widths[col_idx])
            cell.set_height(1.0 / nrows)
            cell.set_linewidth(0)          # 无框线
            cell.set_edgecolor("none")
            cell.set_text_props(ha="center", va="center")

    # 保存为高分辨率 PNG
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


# -------------------- Streamlit 界面 --------------------
st.set_page_config(page_title="Excel → A4 PDF")
st.title("📄 Excel 批量转 A4 单页 PDF（内容自适应）")
st.markdown(
    "上传 Excel 文件，每个工作表生成一个 **A4 单页 PDF**，"
    "根据内容自动缩放，保证文字完整显示，适合打印。"
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
                msg = "⚠️ 以下空工作表已跳过：\n" + "\n".join(
                    [f"{k}: {v}" for k, v in all_skipped.items()]
                )
                st.warning(msg)
            if zipfile.ZipFile(zip_buffer).namelist():
                st.success("✅ 转换完成")
                st.download_button(
                    "⬇️ 下载所有 PDF（ZIP）",
                    data=zip_buffer,
                    file_name="excel_pdfs.zip",
                    mime="application/zip",
                )
            else:
                st.error("❌ 所有工作表均为空，未生成任何 PDF。")

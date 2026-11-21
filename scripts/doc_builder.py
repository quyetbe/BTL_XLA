# -*- coding: utf-8 -*-
"""
Script: scripts/doc_builder.py
Mô tả: Các hàm hỗ trợ định dạng chuẩn học thuật vào file Word BTL_XLA.docx
       (Times New Roman, cỡ 13pt, giãn dòng 1.3, căn đều 2 bên, bảng biểu, hộp công thức).
"""

import docx
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def set_cell_background(cell, fill_hex: str) -> None:
    """Đặt màu nền cho ô trong bảng Word."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)
    tcPr.append(shd)


def set_table_borders(table, color: str = "4A5568", sz: str = "4") -> None:
    """Thiết lập đường viền chuẩn cho bảng trong Word."""
    tblPr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for border_name in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border = OxmlElement(f"w:{border_name}")
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), sz)
        border.set(qn("w:space"), "0")
        border.set(qn("w:color"), color)
        borders.append(border)
    tblPr.append(borders)


def set_row_cant_split(row) -> None:
    """Ngăn ngắt giữa chừng một dòng bảng qua hai trang."""
    trPr = row._tr.get_or_add_trPr()
    if trPr.find(qn("w:cantSplit")) is None:
        trPr.append(OxmlElement("w:cantSplit"))


def set_repeat_table_header(row) -> None:
    """Lặp lại dòng tiêu đề bảng khi bảng kéo dài qua nhiều trang."""
    trPr = row._tr.get_or_add_trPr()
    if trPr.find(qn("w:tblHeader")) is None:
        trPr.append(OxmlElement("w:tblHeader"))


def configure_styles(doc: docx.Document) -> None:
    """Chuẩn hóa font chữ Times New Roman cho các style chính."""
    for style_name in ("Normal", "Heading 1", "Heading 2", "Heading 3"):
        if style_name in doc.styles:
            style = doc.styles[style_name]
            style.font.name = "Times New Roman"
            style.font.color.rgb = RGBColor(0, 0, 0)
            if style_name == "Normal":
                style.font.size = Pt(13)
            elif style_name == "Heading 1":
                style.font.size = Pt(15)
                style.font.bold = True
            elif style_name == "Heading 2":
                style.font.size = Pt(13.5)
                style.font.bold = True
            elif style_name == "Heading 3":
                style.font.size = Pt(13)
                style.font.bold = True


def add_h1(doc: docx.Document, text: str, page_break_before: bool = True):
    """Thêm tiêu đề cấp 1 (Chương / Phần chính)."""
    p = doc.add_paragraph(style="Heading 1")
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.page_break_before = page_break_before
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.line_spacing = 1.3
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    run.font.name = "Times New Roman"
    run.font.size = Pt(15)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0, 0, 0)
    return p


def add_h2(doc: docx.Document, text: str):
    """Thêm tiêu đề cấp 2 (Mục lớn 1.1, 1.2...)."""
    p = doc.add_paragraph(style="Heading 2")
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.3
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    run.font.name = "Times New Roman"
    run.font.size = Pt(13.5)
    run.font.bold = True
    run.font.italic = False
    run.font.color.rgb = RGBColor(0, 0, 0)
    return p


def add_h3(doc: docx.Document, text: str):
    """Thêm tiêu đề cấp 3 (Tiểu mục 1.1.1, 1.1.2...)."""
    p = doc.add_paragraph(style="Heading 3")
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.3
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    run.font.name = "Times New Roman"
    run.font.size = Pt(13)
    run.font.bold = True
    run.font.italic = True
    run.font.color.rgb = RGBColor(0, 0, 0)
    return p


def add_para(doc: docx.Document, text: str, bold_prefix: str = "", first_line_indent: float = 1.0):
    """Thêm đoạn văn bản chuẩn học thuật (căn đều 2 bên, thụt đầu dòng 1cm)."""
    p = doc.add_paragraph(style="Normal")
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.first_line_indent = Cm(first_line_indent)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.3

    if bold_prefix:
        r_bold = p.add_run(bold_prefix)
        r_bold.font.name = "Times New Roman"
        r_bold.font.size = Pt(13)
        r_bold.font.bold = True

    r = p.add_run(text)
    r.font.name = "Times New Roman"
    r.font.size = Pt(13)
    return p


def add_bullet(doc: docx.Document, text: str, bold_prefix: str = ""):
    """Thêm gạch đầu dòng chuẩn."""
    p = doc.add_paragraph(style="Normal")
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.left_indent = Cm(0.8)
    p.paragraph_format.first_line_indent = Cm(-0.4)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.3

    r_dash = p.add_run("- ")
    r_dash.font.name = "Times New Roman"
    r_dash.font.size = Pt(13)
    r_dash.font.bold = True

    if bold_prefix:
        r_b = p.add_run(bold_prefix)
        r_b.font.name = "Times New Roman"
        r_b.font.size = Pt(13)
        r_b.font.bold = True

    r = p.add_run(text)
    r.font.name = "Times New Roman"
    r.font.size = Pt(13)
    return p


def add_formula(doc: docx.Document, formula_text: str, eq_number: str = ""):
    """Thêm dòng công thức toán học căn giữa rõ ràng kèm số thứ tự công thức."""
    p = doc.add_paragraph(style="Normal")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.25

    r = p.add_run(formula_text)
    r.font.name = "Cambria Math"
    r.font.size = Pt(12.5)
    r.font.bold = True

    if eq_number:
        r_num = p.add_run(f"      ({eq_number})")
        r_num.font.name = "Times New Roman"
        r_num.font.size = Pt(12.5)
        r_num.font.italic = True
    return p


def set_cell_margins(cell, top: int = 50, bottom: int = 50, left: int = 90, right: int = 90) -> None:
    """Thiết lập khoảng đệm (dxa) bên trong ô bảng để bảng gọn, tránh xuống dòng vỡ chữ."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = tcPr.find(qn("w:tcMar"))
    if tcMar is None:
        tcMar = OxmlElement("w:tcMar")
        tcPr.append(tcMar)
    for side, val in (("top", top), ("bottom", bottom), ("left", left), ("right", right)):
        el = tcMar.find(qn(f"w:{side}"))
        if el is None:
            el = OxmlElement(f"w:{side}")
            tcMar.append(el)
        el.set(qn("w:w"), str(val))
        el.set(qn("w:type"), "dxa")


def add_code_block(
    doc: docx.Document,
    code_lines: list,
    caption: str = "",
    caption_style: str = "Normal",
    font_size: float = 10.5,
):
    """Thêm khung mã giả hoặc mã lệnh kẻ khung Table Grid đen trắng chuẩn học thuật, không tô màu nền."""
    if caption:
        cap_style = caption_style if caption_style in doc.styles else "Normal"
        p_cap = doc.add_paragraph(style=cap_style)
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER if cap_style == "Hình" else WD_ALIGN_PARAGRAPH.LEFT
        p_cap.paragraph_format.space_before = Pt(6)
        p_cap.paragraph_format.space_after = Pt(2)
        p_cap.paragraph_format.keep_with_next = True
        r_cap = p_cap.add_run(caption)
        r_cap.font.name = "Times New Roman"
        r_cap.font.size = Pt(12)
        r_cap.font.bold = True
        r_cap.font.italic = True

    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    try:
        tbl.style = "Table Grid"
    except KeyError:
        pass
    set_table_borders(tbl, color="000000", sz="4")
    set_row_cant_split(tbl.rows[0])
    cell = tbl.cell(0, 0)

    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.15

    full_text = "\n".join(code_lines)
    r = p.add_run(full_text)
    r.font.name = "Consolas"
    r.font.size = Pt(font_size)
    r.font.color.rgb = RGBColor(0, 0, 0)

    p_space = doc.add_paragraph(style="Normal")
    p_space.paragraph_format.space_before = Pt(0)
    p_space.paragraph_format.space_after = Pt(4)


def add_styled_table(
    doc: docx.Document,
    headers: list,
    rows: list,
    caption: str = "",
    col_aligns: list = None,
    col_widths_cm: list = None,
    font_size: float = 12.0,
):
    """Thêm bảng dữ liệu chuẩn Table Grid đen trắng (không dùng màu nền)."""
    if caption:
        cap_style = "Bảng" if "Bảng" in doc.styles else "Normal"
        p_cap = doc.add_paragraph(style=cap_style)
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_before = Pt(8)
        p_cap.paragraph_format.space_after = Pt(4)
        p_cap.paragraph_format.keep_with_next = True
        r_cap = p_cap.add_run(caption)
        r_cap.font.name = "Times New Roman"
        r_cap.font.size = Pt(12.5)
        r_cap.font.bold = True
        r_cap.font.color.rgb = RGBColor(0, 0, 0)

    tbl = doc.add_table(rows=len(rows) + 1, cols=len(headers))
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    try:
        tbl.style = "Table Grid"
    except KeyError:
        pass
    set_table_borders(tbl, color="000000", sz="4")

    if col_widths_cm:
        tbl.autofit = False

    is_dense = len(headers) >= 8 or font_size <= 10.5
    pad_lr = 50 if is_dense else 90
    pad_tb = 40 if is_dense else 50

    # Header row (chữ in đậm, nền trắng, viền đen chuẩn Table Grid)
    set_repeat_table_header(tbl.rows[0])
    set_row_cant_split(tbl.rows[0])
    hdr_cells = tbl.rows[0].cells
    for j, h_text in enumerate(headers):
        cell = hdr_cells[j]
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        set_cell_margins(cell, top=pad_tb, bottom=pad_tb, left=pad_lr, right=pad_lr)
        if col_widths_cm and j < len(col_widths_cm):
            cell.width = Cm(col_widths_cm[j])
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.1
        r = p.add_run(str(h_text))
        r.font.name = "Times New Roman"
        r.font.size = Pt(font_size)
        r.font.bold = True
        r.font.color.rgb = RGBColor(0, 0, 0)

    # Data rows (nền trắng hoàn toàn, không tô màu xen kẽ)
    for i, row_data in enumerate(rows):
        set_row_cant_split(tbl.rows[i + 1])
        row_cells = tbl.rows[i + 1].cells
        for j, val in enumerate(row_data):
            cell = row_cells[j]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_margins(cell, top=pad_tb, bottom=pad_tb, left=pad_lr, right=pad_lr)
            if col_widths_cm and j < len(col_widths_cm):
                cell.width = Cm(col_widths_cm[j])
            p = cell.paragraphs[0]
            if col_aligns and j < len(col_aligns):
                p.alignment = col_aligns[j]
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT if j > 0 else WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.15
            r = p.add_run(str(val))
            r.font.name = "Times New Roman"
            r.font.size = Pt(font_size)
            r.font.color.rgb = RGBColor(0, 0, 0)

    p_after = doc.add_paragraph(style="Normal")
    p_after.paragraph_format.space_before = Pt(0)
    p_after.paragraph_format.space_after = Pt(6)
    return tbl


def add_figure(doc: docx.Document, image_path: str, caption: str, width_inches: float = 5.8):
    """Chèn hình ảnh căn giữa kèm chú thích hình bên dưới."""
    p_img = doc.add_paragraph(style="Normal")
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img.paragraph_format.space_before = Pt(8)
    p_img.paragraph_format.space_after = Pt(2)
    p_img.paragraph_format.keep_with_next = True
    run_img = p_img.add_run()
    run_img.add_picture(image_path, width=Inches(width_inches))

    cap_style = "Hình" if "Hình" in doc.styles else "Normal"
    p_cap = doc.add_paragraph(style=cap_style)
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap.paragraph_format.space_before = Pt(2)
    p_cap.paragraph_format.space_after = Pt(8)
    r_cap = p_cap.add_run(caption)
    r_cap.font.name = "Times New Roman"
    r_cap.font.size = Pt(12)
    r_cap.font.italic = True
    r_cap.font.bold = True

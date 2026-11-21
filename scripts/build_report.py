# -*- coding: utf-8 -*-
"""
Script: scripts/build_report.py
Mô tả: Tạo lại toàn bộ nội dung báo cáo BTL_XLA.docx theo chuẩn cấu trúc mới của
       template.docx (4 Chương + Kết luận & Hướng phát triển + Tài liệu tham khảo & Phụ lục)
       và đáp ứng đầy đủ mọi tiêu chí trong Phiếu chấm.docx.
       Giữ nguyên Trang bìa, Mục lục tự động (TOC), Danh mục từ viết tắt,
       Danh mục hình ảnh, Danh mục bảng biểu và ngắt section đánh số trang La Mã / Ả Rập.
"""

import os
import sys
import csv
import docx
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from scripts.doc_builder import (
    configure_styles,
    add_h1,
    add_h2,
    add_h3,
    add_para,
    add_bullet,
    add_formula,
    add_code_block,
    add_styled_table,
    add_figure,
    set_row_cant_split,
    set_repeat_table_header,
)


def load_csv_rows(csv_path: str) -> list:
    with open(csv_path, mode="r", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def reset_body_keeping_front_matter(doc: docx.Document) -> None:
    """
    Giữ lại toàn bộ Trang bìa (Sec 0), Mục lục tự động, Danh mục từ viết tắt,
    Danh mục hình ảnh, Danh mục bảng biểu và Section Break chuyển sang đánh số trang Ả Rập (Sec 1 -> Sec 2).
    Xóa phần thân bài cũ từ LỜI MỞ ĐẦU / CHƯƠNG 1 trở về sau để ghi mới hoàn toàn.
    """
    body = doc.element.body
    children = list(body)

    cut_idx = None
    for idx, child in enumerate(children):
        if child.tag.endswith("p"):
            txt = "".join(child.itertext()).strip().upper()
            if txt.startswith("LỜI MỞ ĐẦU") or txt.startswith("MỞ ĐẦU") or txt.startswith("CHƯƠNG 1"):
                cut_idx = idx
                break

    if cut_idx is not None:
        # Bảo đảm đoạn văn bản chứa <w:sectPr> ngay trước cut_idx không mang style Heading 1
        # (tránh sinh dòng trắng trong Mục lục tự động)
        if cut_idx - 1 >= 0:
            prev_el = children[cut_idx - 1]
            pPr = prev_el.find(qn("w:pPr"))
            if pPr is not None:
                pStyle = pPr.find(qn("w:pStyle"))
                if pStyle is not None and pStyle.get(qn("w:val")) == "Heading1":
                    pPr.remove(pStyle)

        for child in children[cut_idx:]:
            if not child.tag.endswith("sectPr"):
                body.remove(child)


def update_abbreviations_table(doc: docx.Document) -> None:
    """Cập nhật Bảng Danh mục các từ viết tắt (doc.tables[1]) đầy đủ cho báo cáo mới."""
    if len(doc.tables) < 2:
        return

    abbr_rows = [
        ["2D", "Two-Dimensional", "Hai chiều (không gian phẳng gồm chiều cao và chiều rộng)"],
        ["ADC", "Analog-to-Digital Converter", "Bộ chuyển đổi tín hiệu tương tự sang tín hiệu số"],
        ["AMF", "Adaptive Median Filter", "Bộ lọc trung vị thích nghi"],
        ["API", "Application Programming Interface", "Giao diện lập trình ứng dụng"],
        ["ASGI", "Asynchronous Server Gateway Interface", "Chuẩn giao diện máy chủ web bất đồng bộ trong Python"],
        ["AWGN", "Additive White Gaussian Noise", "Nhiễu cộng trắng phân phối chuẩn Gaussian"],
        ["CCD / CMOS", "Charge-Coupled Device / Complementary Metal-Oxide-Semiconductor", "Cảm biến quang học thu nhận hình ảnh trong máy ảnh số"],
        ["CSV", "Comma-Separated Values", "Định dạng tập tin bảng số liệu phân cách bằng dấu phẩy"],
        ["dB", "Decibel", "Đơn vị đo lường tỷ số tín hiệu theo thang logarit"],
        ["Edge-MSE", "Edge Mean Squared Error", "Sai số bình phương trung bình giữa hai bản đồ độ lớn biên"],
        ["F1-Score", "Harmonic Mean of Precision and Recall", "Trung bình điều hòa giữa độ chính xác (Precision) và độ phủ (Recall)"],
        ["FN / FP / TP", "False Negative / False Positive / True Positive", "Âm tính giả / Dương tính giả / Dương tính thật (trong phát hiện biên)"],
        ["JSON", "JavaScript Object Notation", "Định dạng dữ liệu trao đổi có cấu trúc dạng khóa - giá trị"],
        ["MLE", "Maximum Likelihood Estimation", "Ước lượng hợp lý cực đại"],
        ["MSE", "Mean Squared Error", "Sai số bình phương trung bình"],
        ["NLM", "Non-Local Means", "Thuật toán lọc trung bình phi cục bộ"],
        ["PDF", "Probability Density Function", "Hàm mật độ xác suất"],
        ["PNG", "Portable Network Graphics", "Định dạng ảnh nén không mất dữ liệu (Lossless)"],
        ["PSNR", "Peak Signal-to-Noise Ratio", "Tỷ số tín hiệu cực đại trên nhiễu"],
        ["RAM", "Random Access Memory", "Bộ nhớ truy cập ngẫu nhiên"],
        ["REST", "Representational State Transfer", "Kiến trúc thiết kế dịch vụ web thông qua giao thức HTTP"],
        ["RGB", "Red - Green - Blue", "Không gian màu Đỏ - Lục - Lam"],
        ["S&P", "Salt-and-Pepper Noise", "Nhiễu muối tiêu (nhiễu xung lưỡng cực)"],
        ["UI", "User Interface", "Giao diện người dùng"],
    ]

    tbl = doc.tables[1]
    # Xóa tất cả các dòng dữ liệu cũ (giữ lại dòng header 0)
    for row in list(tbl.rows)[1:]:
        tbl._tbl.remove(row._tr)

    set_repeat_table_header(tbl.rows[0])
    set_row_cant_split(tbl.rows[0])

    for abbr, eng, vie in abbr_rows:
        row = tbl.add_row()
        set_row_cant_split(row)
        vals = [abbr, eng, vie]
        aligns = [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT]
        for j, (val, align) in enumerate(zip(vals, aligns)):
            cell = row.cells[j]
            p = cell.paragraphs[0]
            p.alignment = align
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.line_spacing = 1.15
            r = p.add_run(val)
            r.font.name = "Times New Roman"
            r.font.size = Pt(12)
            r.font.color.rgb = RGBColor(0, 0, 0)


def enable_update_fields_on_open(doc: docx.Document) -> None:
    """Bật cờ updateFields trong settings.xml để Word tự hỏi cập nhật Mục lục khi mở."""
    settings = doc.settings.element
    if settings.find(qn("w:updateFields")) is None:
        el = OxmlElement("w:updateFields")
        el.set(qn("w:val"), "true")
        settings.append(el)


def write_preface(doc: docx.Document) -> None:
    """Ghi phần LỜI MỞ ĐẦU (Thông tin học phần và phân công thành viên Nhóm 10)."""
    add_h1(doc, "LỜI MỞ ĐẦU", page_break_before=False)

    add_h2(doc, "1. Giới thiệu tổng quan học phần")
    add_para(
        doc,
        "Trong kỷ nguyên chuyển đổi số, dữ liệu thị giác (hình ảnh tĩnh và luồng video số) đóng vai trò trung tâm trong nhiều lĩnh vực từ y tế, tự động hóa công nghiệp, giám sát an ninh đến xe tự hành. Học phần Xử lý ảnh và thị giác máy tính tại Khoa Công nghệ Thông tin – Trường Đại học Công nghệ Đông Á cung cấp cho sinh viên nền tảng toán học, tư duy thuật toán và kỹ năng lập trình thực nghiệm để xử lý, cải thiện chất lượng ảnh và trích xuất thông tin hình học từ ảnh số."
    )
    add_para(
        doc,
        "Thực hiện Đề tài 7: \"Khử nhiễu ảnh bằng các bộ lọc trong miền không gian\", Nhóm 10 tập trung nghiên cứu, tự cài đặt hoàn toàn bằng NumPy hai phương pháp lọc kinh điển (Bộ lọc Trung bình và Bộ lọc Trung vị) kết hợp toán tử phát hiện biên Sobel nhằm đánh giá toàn diện cả chất lượng khôi phục mức xám lẫn khả năng bảo toàn cấu trúc biên ảnh, đồng thời xây dựng ứng dụng Web tương tác phục vụ kiểm thử trực quan."
    )

    add_h2(doc, "2. Thành viên Nhóm 10 và phân công công việc chi tiết")
    add_para(
        doc,
        "Bài tập lớn được thực hiện bởi Nhóm 10 thuộc lớp tín chỉ Xử lý ảnh và thị giác máy tính-1-1-26(N05.THTT_IN_ELN.11), dưới sự hướng dẫn của Giảng viên Nguyễn Phồn Lữa. Bảng 0.1 trình bày phân công nhiệm vụ cụ thể và tỷ lệ đóng góp của từng thành viên."
    )

    member_headers = [
        "TT",
        "Mã SV",
        "Họ và tên",
        "Lớp HC",
        "Vai trò",
        "Nhiệm vụ phụ trách chi tiết",
        "Tỷ lệ hoàn thành",
    ]
    member_rows = [
        [
            "1",
            "20231053",
            "Nguyễn Mạnh Quyết",
            "DCCNTT.14.3",
            "Nhóm trưởng",
            "- Nghiên cứu cơ sở lý thuyết biểu diễn ảnh số, phép tích chập 2D, bộ lọc trung bình và toán tử biên Sobel.\n"
            "- Tự cài đặt các module cốt lõi: mở rộng biên (src/padding.py), tích chập & lọc không gian (src/filters.py), toán tử Sobel (src/sobel.py).\n"
            "- Xây dựng kịch bản điều phối thực nghiệm tự động (run_experiments.py), đối chiếu OpenCV (src/verification.py) và giao diện Web (web_app.py).\n"
            "- Biên soạn Lời mở đầu, Chương 1, Chương 2 (Mục 2.1.1–2.1.3), Chương 3 và chuẩn hóa định dạng báo cáo.",
            "100%\n(Đóng góp 50%)",
        ],
        [
            "2",
            "20231045",
            "Hà Thái Bảo",
            "DCCNTT.14.3",
            "Thành viên",
            "- Nghiên cứu cơ sở lý thuyết mô hình nhiễu (Muối tiêu, Gaussian), bộ lọc trung vị và các độ đo MSE, PSNR, Edge-MSE, Edge-F1.\n"
            "- Tự cài đặt các module: giả lập nhiễu có seed cố định (src/noise.py), tính chỉ số đánh giá (src/metrics.py), chuẩn bị dữ liệu (src/dataset.py) và vẽ biểu đồ (src/visualization.py).\n"
            "- Chạy thực nghiệm trên 6 ảnh chuẩn, thu thập số liệu CSV, kiểm thử Web UI và xuất hệ thống hình ảnh so sánh.\n"
            "- Biên soạn Chương 2 (Mục 2.1.4–2.2), Chương 4, Kết luận và Hướng phát triển, Tài liệu tham khảo & Phụ lục.",
            "100%\n(Đóng góp 50%)",
        ],
    ]
    add_styled_table(
        doc,
        headers=member_headers,
        rows=member_rows,
        caption="Bảng 0.1. Danh sách thành viên Nhóm 10 và bảng phân công nhiệm vụ chi tiết",
        col_aligns=[
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.CENTER,
        ],
        col_widths_cm=[0.9, 2.0, 2.4, 2.6, 1.7, 4.6, 1.8],
        font_size=10.5,
    )


def write_chapter_1(doc: docx.Document) -> None:
    """CHƯƠNG 1: TỔNG QUAN BÀI TOÁN VÀ YÊU CẦU HỆ THỐNG (Tiêu chí 2.2 Phiếu chấm)."""
    add_h1(doc, "CHƯƠNG 1: TỔNG QUAN BÀI TOÁN VÀ YÊU CẦU HỆ THỐNG", page_break_before=True)

    # 1.1
    add_h2(doc, "1.1. Đặt vấn đề và Bối cảnh ứng dụng")
    add_para(
        doc,
        "Trong thực tế, không có hệ thống thu nhận và truyền tải ảnh số nào là lý tưởng tuyệt đối. Trong quá trình chụp ảnh qua cảm biến quang học CCD/CMOS, lượng tử hóa tín hiệu qua bộ chuyển đổi tương tự – số (ADC), lưu trữ và truyền dẫn qua kênh truyền, ảnh số luôn bị suy giảm bởi các thành phần nhiễu ngẫu nhiên. Cụ thể, dao động nhiệt điện tử hoặc điều kiện chụp thiếu sáng (ISO cao) sinh ra nhiễu cộng phân phối chuẩn (Gaussian Noise) phủ hạt mịn lên toàn bộ khung hình; trong khi lỗi xung đóng mở mạch phần cứng, bụi cảm biến, điểm ảnh chết (dead pixels) hoặc lỗi bit truyền dẫn lại tạo ra nhiễu xung lưỡng cực dạng chấm trắng – đen cực trị, gọi là nhiễu Muối tiêu (Salt-and-Pepper Noise)."
    )
    add_para(
        doc,
        "Dưới góc độ kiến trúc hệ thống thị giác máy tính, tiến trình xử lý thông tin thị giác gồm ba tầng liên kết chặt chẽ: (1) Xử lý ảnh mức thấp (thu nhận, khử nhiễu, tăng cường ảnh); (2) Xử lý ảnh mức trung gian (phát hiện biên, phân đoạn vùng, trích xuất đặc trưng); và (3) Thị giác máy tính mức cao (nhận dạng, phân loại và hiểu ngữ nghĩa). Trong chuỗi xử lý này, khâu khử nhiễu ở tầng mức thấp giữ vai trò nền móng quyết định: nếu ảnh đầu vào còn lẫn nhiễu nặng, các toán tử đạo hàm ở tầng trung gian (như phát hiện biên Sobel, Prewitt, Canny) sẽ khuếch đại nhiễu thành hàng loạt điểm biên giả (false edges); ngược lại, nếu khử nhiễu bằng bộ lọc không phù hợp hoặc kích thước cửa sổ quá lớn, các đường biên thật và chi tiết mảnh của vật thể sẽ bị làm nhòe hoặc xóa mất."
    )
    add_para(
        doc,
        "Do đó, bài toán thực tế mà Đề tài 7 giải quyết là: xây dựng một hệ thống phần mềm tự cài đặt hoàn chỉnh các bộ lọc khử nhiễu trong miền không gian (Bộ lọc Trung bình tuyến tính và Bộ lọc Trung vị phi tuyến) kết hợp toán tử biên Sobel, nhằm trả lời định lượng câu hỏi kỹ thuật: Bộ lọc nào và kích thước cửa sổ nào tối ưu cho từng loại nhiễu và cấu trúc ảnh khác nhau mà vẫn bảo toàn tốt nhất đường biên đối tượng?"
    )

    # 1.2
    add_h2(doc, "1.2. Mục tiêu và Phạm vi đề tài")
    add_h3(doc, "1.2.1. Mục tiêu và đặc tả Đầu vào – Đầu ra (Input / Output)")
    add_para(
        doc,
        "Mục tiêu cốt lõi của đề tài là nắm vững bản chất toán học của lọc không gian, tự cài đặt được toàn bộ chuỗi thuật toán xử lý ảnh bằng NumPy (không phụ thuộc hàm lọc đóng gói sẵn), thực nghiệm đối chứng khách quan trên đa dạng dữ liệu và xây dựng giao diện Web trực quan. Đặc tả luồng Đầu vào – Đầu ra của chương trình được xác định như sau:"
    )
    add_bullet(
        doc,
        "Ảnh đa mức xám 8-bit (8-bit Grayscale Image) dạng ma trận 2 chiều kích thước (H, W), kiểu dữ liệu uint8 (giá trị mức xám thuộc [0, 255]). Trong chế độ thực nghiệm hàng loạt, đầu vào gồm 6 ảnh chuẩn kích thước 256 x 256 pixel; trong chế độ Web tương tác, người dùng có thể chọn ảnh mẫu hoặc tải lên ảnh màu/ảnh xám bất kỳ (.PNG, .JPG) để hệ thống tự động chuyển về ảnh xám 8-bit và chuẩn hóa về kích thước 256 x 256 hoặc 512 x 512 pixel, kèm các tham số cấu hình (loại nhiễu, mức nhiễu, kích thước cửa sổ k và chế độ đệm biên).",
        bold_prefix="Đầu vào (Input): "
    )
    add_bullet(
        doc,
        "(1) Ma trận ảnh sau khử nhiễu bằng Bộ lọc Trung bình và Bộ lọc Trung vị (giữ nguyên kích thước H x W, kiểu uint8); (2) Bản đồ độ lớn Gradient và bản đồ biên nhị phân Sobel của ảnh gốc, ảnh nhiễu và các ảnh sau lọc; (3) Hệ thống chỉ số đánh giá định lượng (MSE, PSNR, Edge-MSE, Edge-Precision, Edge-Recall, Edge-F1, thời gian thực thi ms) xuất ra các bảng dữ liệu CSV, biểu đồ phân tích PNG và hiển thị trực quan 8 khung hình trên giao diện Web.",
        bold_prefix="Đầu ra (Output): "
    )

    add_h3(doc, "1.2.2. Giới hạn phạm vi và các ràng buộc kỹ thuật")
    add_bullet(
        doc,
        "Đối tượng xử lý là ảnh tĩnh đơn kênh đa mức xám (2D Grayscale); chưa áp dụng trực tiếp cho luồng video thời gian thực hay xử lý tách kênh màu 3D (RGB/HSV).",
        bold_prefix="Phạm vi dữ liệu: "
    )
    add_bullet(
        doc,
        "Khảo sát hai mô hình nhiễu đại diện là nhiễu Muối tiêu (mật độ 1%, 5%, 10%, tỷ lệ muối/tiêu 0.5) và nhiễu cộng trắng Gaussian (kỳ vọng μ = 0, độ lệch chuẩn σ = 10, 25, 50); hai bộ lọc miền không gian là Bộ lọc Trung bình và Bộ lọc Trung vị với ba kích thước cửa sổ vuông lẻ k ∈ {3, 5, 7} và ba chế độ đệm biên ('reflect', 'replicate', 'zero').",
        bold_prefix="Phạm vi thuật toán và tham số: "
    )
    add_bullet(
        doc,
        "Toàn bộ thuật toán cốt lõi (Padding, sinh nhiễu có seed = 42 cố định để tái lập 100%, tích chập 2D, lọc Trung bình, lọc Trung vị, toán tử Sobel và các hàm tính MSE/PSNR/Edge-F1) phải được nhóm tự cài đặt từ đầu (from scratch) bằng NumPy. OpenCV và Pillow chỉ dùng để đọc/ghi tệp ảnh và kiểm chứng đối chiếu sai số trong module độc lập (src/verification.py).",
        bold_prefix="Ràng buộc cài đặt (From Scratch): "
    )

    # 1.3
    add_h2(doc, "1.3. Đặc tả chức năng của chương trình")
    add_para(
        doc,
        "Chương trình được thiết kế gồm 8 phân hệ chức năng chính phục vụ cả hai chế độ vận hành: chạy thực nghiệm tự động hàng loạt (CLI script run_experiments.py) và trình diễn tương tác trực quan qua trình duyệt Web (web_app.py). Bảng 1.1 mô tả chi tiết các chức năng của chương trình."
    )

    func_headers = ["TT", "Chức năng chính", "Module đảm nhiệm", "Mô tả chi tiết Đầu vào – Xử lý – Đầu ra"]
    func_rows = [
        [
            "1",
            "Nạp và chuẩn hóa dữ liệu ảnh",
            "src/dataset.py",
            "Đọc hoặc trích xuất 6 ảnh chuẩn (cameraman, coins, astronaut, text, checkerboard, brick), chuyển về ảnh xám 8-bit kích thước 256x256, lưu trữ dưới dạng PNG trong data/original/.",
        ],
        [
            "2",
            "Mở rộng biên ảnh 2D (Padding)",
            "src/padding.py",
            "Tự cài đặt 3 chế độ đệm biên ('zero', 'replicate', 'reflect') với bán kính p = ⌊k/2⌋ để bảo toàn kích thước (H, W) sau lọc và tránh nhiễu giả ở viền ảnh.",
        ],
        [
            "3",
            "Giả lập nhiễu có kiểm soát",
            "src/noise.py",
            "Sinh nhiễu Muối tiêu (mật độ 0.0–1.0) và nhiễu cộng Gaussian (σ ≥ 0) với hạt giống cố định seed = 42, bảo đảm tính tái lập 100% giữa mọi lần chạy.",
        ],
        [
            "4",
            "Khử nhiễu trong miền không gian",
            "src/filters.py",
            "Tự cài đặt phép tích chập 2D tổng quát (convolve2d), Bộ lọc Trung bình (mean_filter) và Bộ lọc Trung vị (median_filter) hỗ trợ cửa sổ lẻ 3x3, 5x5, 7x7.",
        ],
        [
            "5",
            "Phát hiện biên bằng toán tử Sobel",
            "src/sobel.py",
            "Tính đạo hàm Gx, Gy bằng tích chập tự cài đặt với cặp mặt nạ Sobel 3x3, tính độ lớn Gradient và ngưỡng hóa theo phân vị thứ 85 (P85) của ảnh gốc.",
        ],
        [
            "6",
            "Đo lường chất lượng & Kiểm chứng OpenCV",
            "src/metrics.py &\nsrc/verification.py",
            "Tính MSE, PSNR (dB), Edge-MSE, Precision, Recall, Edge-F1; đo thời gian thực thi (ms) và kiểm chứng sai số tuyệt đối của code tự cài đặt với OpenCV.",
        ],
        [
            "7",
            "Điều phối thực nghiệm & Xuất báo cáo số liệu",
            "run_experiments.py &\nsrc/visualization.py",
            "Tự động thực thi toàn bộ 216 cấu hình lọc (6 ảnh × 6 mức nhiễu × 2 bộ lọc × 3 kích thước), xuất 3 tệp CSV và vẽ hệ thống biểu đồ, ảnh so sánh PNG.",
        ],
        [
            "8",
            "Ứng dụng Web tương tác trên Localhost",
            "web_app.py\n(FastAPI + UI)",
            "Cho phép chọn ảnh mẫu hoặc tải ảnh từ máy tính, chỉnh slider nhiễu, đổi kernel/padding theo thời gian thực, hiển thị 8 khung ảnh Trước/Sau & Biên Sobel, nút Reset ảnh gốc và tải ảnh kết quả.",
        ],
    ]
    add_styled_table(
        doc,
        headers=func_headers,
        rows=func_rows,
        caption="Bảng 1.1. Đặc tả chi tiết các phân hệ chức năng của chương trình",
        col_aligns=[
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.LEFT,
        ],
    )


def write_chapter_2(doc: docx.Document) -> None:
    """CHƯƠNG 2: CƠ SỞ LÝ THUYẾT VÀ CÔNG CỤ THỰC HIỆN (Tiêu chí 2.1 Phiếu chấm)."""
    add_h1(doc, "CHƯƠNG 2: CƠ SỞ LÝ THUYẾT VÀ CÔNG CỤ THỰC HIỆN", page_break_before=True)

    # 2.1
    add_h2(doc, "2.1. Cơ sở toán học và Lý thuyết thuật toán")

    # 2.1.1
    add_h3(doc, "2.1.1. Biểu diễn ảnh số trong miền không gian và lân cận điểm ảnh")
    add_para(
        doc,
        "Về mặt vật lý, một bức ảnh tự nhiên là một hàm cường độ sáng liên tục hai chiều f(x, y). Để máy tính lưu trữ và xử lý được, hàm f(x, y) trải qua hai quá trình rời rạc hóa: Lấy mẫu (Sampling) nhằm rời rạc hóa tọa độ không gian thành lưới M hàng × N cột điểm ảnh (pixel), và Lượng tử hóa (Quantization) nhằm rời rạc hóa biên độ sáng thành L = 2^8 = 256 mức nguyên đối với ảnh đa mức xám 8-bit:"
    )
    add_formula(
        doc,
        "f(x, y) ∈ {0, 1, 2, ..., 255},   với x ∈ {0, 1, ..., M - 1},  y ∈ {0, 1, ..., N - 1}",
        eq_number="2.1",
    )
    add_para(
        doc,
        "Trong hệ tọa độ chuẩn của xử lý ảnh số (và thư viện NumPy/OpenCV), gốc tọa độ (0, 0) đặt ở góc trên bên trái; chỉ số x tăng từ trên xuống dưới (hàng) và chỉ số y tăng từ trái sang phải (cột). Các thuật toán xử lý trong miền không gian (Spatial Domain) thao tác trực tiếp trên ma trận điểm ảnh thông qua toán tử T tác động lên tập lân cận S_xy kích thước lẻ k × k (với k = 2a + 1, bán kính a = (k - 1)/2) có tâm đặt tại (x, y):"
    )
    add_formula(
        doc,
        "g(x, y) = T [ f(x, y) ],     với S_xy = { (x + s, y + t) | -a ≤ s ≤ a,  -a ≤ t ≤ a }",
        eq_number="2.2",
    )

    # 2.1.2
    add_h3(doc, "2.1.2. Các mô hình nhiễu điển hình: Nhiễu Muối tiêu và Nhiễu Gaussian")
    add_para(
        doc,
        "Nhiễu ảnh là những biến thiên ngẫu nhiên không mong muốn chồng lấn lên tín hiệu ảnh thực f(x, y) tạo thành ảnh quan sát g(x, y). Hai mô hình nhiễu phổ biến nhất trong miền không gian bao gồm:"
    )
    add_bullet(
        doc,
        "Còn gọi là nhiễu xung lưỡng cực (Bipolar Impulse Noise), xuất hiện dưới dạng các hạt trắng (giá trị cực đại 255) và hạt đen (giá trị cực tiểu 0) rải rác ngẫu nhiên do lỗi chuyển mạch ADC, lỗi bit truyền dẫn hoặc điểm ảnh chết. Với mật độ nhiễu tổng cộng p ∈ [0, 1] và tỷ lệ cân bằng p_s = p_p = p/2, giá trị điểm ảnh g(x, y) tuân theo phân phối rời rạc:",
        bold_prefix="Nhiễu Muối tiêu (Salt-and-Pepper Noise): "
    )
    add_formula(
        doc,
        "g(x, y) = { 255 (với xác suất p/2);   0 (với xác suất p/2);   f(x, y) (với xác suất 1 - p) }",
        eq_number="2.3",
    )
    add_para(
        doc,
        "Đặc tính cốt lõi của nhiễu Muối tiêu là chỉ phá hủy cục bộ p% số pixel thành cực trị 0 hoặc 255, trong khi (1 - p)% pixel còn lại giữ nguyên hoàn toàn giá trị gốc f(x, y)."
    )
    add_bullet(
        doc,
        "Là nhiễu cộng liên tục tác động lên gần như 100% số điểm ảnh do nhiễu nhiệt mạch khuếch đại trên cảm biến CCD/CMOS hoặc chụp thiếu sáng. Thành phần nhiễu η(x, y) tại mỗi pixel là biến ngẫu nhiên độc lập tuân theo phân phối chuẩn N(μ, σ^2) với hàm mật độ xác suất (PDF):",
        bold_prefix="Nhiễu cộng trắng Gaussian (AWGN): "
    )
    add_formula(
        doc,
        "p(η) = [ 1 / (√(2π) · σ) ] · exp( - (η - μ)^2 / (2σ^2) )",
        eq_number="2.4",
    )
    add_para(
        doc,
        "Với kỳ vọng μ = 0, cường độ nhiễu được điều khiển bởi độ lệch chuẩn σ (σ = 10 nhẹ, σ = 25 trung bình, σ = 50 nặng). Do phép cộng có thể vượt ngoài [0, 255], ảnh nhiễu 8-bit được tạo theo công thức g(x, y) = clip(round(f(x, y) + η(x, y)), 0, 255)."
    )

    noise_comp_headers = ["Tiêu chí so sánh", "Nhiễu Muối tiêu (Salt-and-Pepper)", "Nhiễu Gaussian (AWGN)"]
    noise_comp_rows = [
        ["Bản chất toán học", "Nhiễu xung lưỡng cực rời rạc (thay thế giá trị pixel bằng cực trị 0 hoặc 255).", "Nhiễu cộng liên tục tuân theo phân phối chuẩn N(0, σ^2)."],
        ["Phạm vi tác động", "Chỉ ảnh hưởng cục bộ tỷ lệ p% điểm ảnh; (100 - p)% điểm ảnh còn lại nguyên vẹn.", "Tác động toàn cục lên gần như 100% số điểm ảnh trên toàn bức ảnh."],
        ["Biểu hiện thị giác", "Các chấm trắng sáng và đen sẫm tương phản mạnh rải rác độc lập.", "Lớp hạt mịn phủ đều khắp ảnh, làm giảm độ trong trẻo và nhòe chi tiết."],
        ["Tham số điều khiển", "Mật độ nhiễu p (density ∈ [0, 1]).", "Độ lệch chuẩn σ (standard deviation ≥ 0)."],
    ]
    add_styled_table(
        doc,
        headers=noise_comp_headers,
        rows=noise_comp_rows,
        caption="Bảng 2.1. So sánh đặc trưng lý thuyết giữa nhiễu Muối tiêu và nhiễu Gaussian",
        col_aligns=[WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT],
    )

    # 2.1.3
    add_h3(doc, "2.1.3. Phép tương quan, Tích chập hai chiều và Kỹ thuật mở rộng biên ảnh")
    add_para(
        doc,
        "Trong lọc không gian tuyến tính với mặt nạ nhân lọc (kernel) h(s, t) kích thước k × k (k = 2a + 1), phép Tương quan không gian (Correlation, ký hiệu f ⊗ h) và phép Tích chập hai chiều (2D Convolution, ký hiệu f * h) được định nghĩa như sau:"
    )
    add_formula(
        doc,
        "(f ⊗ h)(x, y) = ∑_{s=-a}^{a} ∑_{t=-a}^{a} h(s, t) · f(x + s, y + t)",
        eq_number="2.5",
    )
    add_formula(
        doc,
        "(f * h)(x, y) = ∑_{s=-a}^{a} ∑_{t=-a}^{a} h(s, t) · f(x - s, y - t)",
        eq_number="2.6",
    )
    add_para(
        doc,
        "Điểm khác biệt cốt lõi là tích chập 2D thực hiện lật đối xứng nhân lọc 180° theo cả hai trục trước khi tính tổng tích vô hướng. Khi nhân lọc đối xứng tâm (như nhân lọc Trung bình), tương quan và tích chập cho kết quả trùng nhau; nhưng với nhân lọc bất đối xứng tính đạo hàm (như Sobel), phải kiểm soát chặt hướng lật nhân để bảo toàn đúng dấu vector Gradient."
    )
    add_para(
        doc,
        "Khi cửa sổ trượt k × k đặt tâm tại các điểm ảnh sát mép ảnh, một phần cửa sổ tràn ra ngoài biên ma trận. Để giữ nguyên kích thước ảnh đầu ra (M × N), ta cần mở rộng thêm a hàng/cột xung quanh ảnh gốc thành ma trận (M + 2a) × (N + 2a) bằng kỹ thuật đệm biên (Boundary Padding) gồm 3 chế độ phổ biến: (1) Zero-padding (điền số 0 — đơn giản nhưng tạo bước nhảy tối tại biên và sinh biên giả mạnh khi chạy Sobel); (2) Replicate-padding (lặp lại hàng/cột ngoài cùng); và (3) Reflect-padding (phản chiếu gương các điểm ảnh sát biên qua trục mép ảnh, duy trì tính liên tục của mức xám và triệt tiêu biên giả)."
    )

    # 2.1.4
    add_h3(doc, "2.1.4. Bộ lọc Trung bình và Bộ lọc Trung vị (kèm ví dụ tính tay 3x3)")
    add_para(
        doc,
        "Hai bộ lọc khử nhiễu nền tảng đại diện cho hai trường phái đối lập nhau (tuyến tính và phi tuyến) trong miền không gian là:"
    )
    add_bullet(
        doc,
        "Là bộ lọc không gian tuyến tính (Box Filter), thay thế mỗi điểm ảnh bằng trung bình cộng số học của k^2 điểm ảnh trong lân cận S_xy, tương đương tích chập ảnh với ma trận nhân lọc H_{k×k} có mọi phần tử bằng 1/k^2:",
        bold_prefix="Bộ lọc Trung bình (Arithmetic Mean Filter): "
    )
    add_formula(
        doc,
        "g_mean(x, y) = (1 / k^2) · ∑_{(r, c) ∈ S_xy} f(r, c),     H_{3x3} = (1/9) · [1]_{3x3}",
        eq_number="2.7",
    )
    add_para(
        doc,
        "Với nhiễu cộng trắng Gaussian độc lập η_i ~ N(0, σ^2) trên vùng ảnh phẳng, bộ lọc Trung bình là ước lượng hợp lý cực đại (MLE) giúp giảm phương sai nhiễu đi đúng k^2 lần: Var(g_mean) = σ^2 / k^2 (độ lệch chuẩn giảm k lần). Tuy nhiên, bộ lọc này làm nhòe các đường biên sắc nét (edge blurring) và khi gặp điểm nhiễu xung cực trị (0 hoặc 255) sẽ phát tán điểm nhiễu đó thành đốm mờ kích thước k × k."
    )
    add_bullet(
        doc,
        "Là bộ lọc phi tuyến dựa trên thống kê thứ tự (Order-Statistic Filter). Tại mỗi vị trí (x, y), thuật toán thu thập k^2 giá trị mức xám trong S_xy, sắp xếp thành dãy không giảm I_(0) ≤ I_(1) ≤ ... ≤ I_(k^2 - 1) và chọn phần tử nằm chính giữa:",
        bold_prefix="Bộ lọc Trung vị (Median Filter): "
    )
    add_formula(
        doc,
        "g_median(x, y) = median { f(r, c) | (r, c) ∈ S_xy } = I_( (k^2 - 1) / 2 )",
        eq_number="2.8",
    )
    add_para(
        doc,
        "Do các hạt muối (255) và tiêu (0) là hai cực trị nên sau khi sắp xếp chúng luôn bị dạt về hai đầu mút của dãy. Chỉ cần số điểm nhiễu trong cửa sổ nhỏ hơn một nửa diện tích cửa sổ (≤ (k^2 - 1)/2 điểm, tức điểm gãy lý thuyết gần 50%), phần tử trung vị hoàn toàn là một điểm ảnh sạch, triệt tiêu triệt để nhiễu xung mà không tạo ra mức xám trung gian mới, nhờ đó bảo toàn sắc nét các đường biên bước nhảy (step edges)."
    )

    add_code_block(
        doc,
        [
            "Ma trận vùng ảnh lân cận 3x3 (nền phẳng ~100, tâm bị nhiễu muối = 255):",
            "    [ 100   102    98 ]",
            "    [  99   255   101 ]",
            "    [ 100   101    99 ]",
            "",
            "1) Bộ lọc Trung bình 3x3: g_mean = round((100+102+98+99+255+101+100+101+99)/9) = round(117.22) = 117.",
            "   -> Sai lệch +17 mức xám và phát tán giá trị 255 sang cả 8 điểm lân cận thành đốm mờ 3x3.",
            "",
            "2) Bộ lọc Trung vị 3x3: Sắp xếp tăng dần -> [98, 99, 99, 100, (100), 101, 101, 102, 255].",
            "   -> Chọn phần tử chỉ số 4 ở giữa: g_median = 100 (khôi phục chính xác 100%, sai số bằng 0).",
        ],
        caption="Ví dụ 2.1. Minh họa tính tay kết quả lọc Trung bình và lọc Trung vị trên cửa sổ 3x3 có nhiễu xung",
    )

    filter_comp_headers = ["Đặc tính kỹ thuật", "Bộ lọc Trung bình (Mean Filter)", "Bộ lọc Trung vị (Median Filter)"]
    filter_comp_rows = [
        ["Phân loại toán học", "Lọc không gian tuyến tính (thực hiện bằng tích chập 2D).", "Lọc không gian phi tuyến (dựa trên sắp xếp thống kê thứ tự)."],
        ["Cơ chế tính đầu ra", "Trung bình cộng có trọng số đều 1/k^2 của k^2 pixel lân cận.", "Sắp xếp tăng dần k^2 pixel và chọn phần tử ở chỉ số ⌊k^2/2⌋."],
        ["Khả năng sinh mức xám mới", "Có thể sinh giá trị trung gian mới không có trong ảnh gốc.", "Không sinh mức xám mới; luôn lấy một giá trị thực có sẵn."],
        ["Hiệu quả với nhiễu Muối tiêu", "Kém (phát tán nhiễu cực trị thành đốm mờ kích thước k × k).", "Rất xuất sắc (điểm gãy ~44.4%–48%, xóa triệt để nhiễu xung)."],
        ["Hiệu quả với nhiễu Gaussian", "Tốt (tối ưu MLE, phương sai còn lại σ^2/k^2).", "Khá (phương sai còn lại ≈ (π/2)·(σ^2/k^2) ≈ 1.57 lần Mean)."],
        ["Bảo toàn biên ảnh", "Kém (làm suy giảm độ dốc gradient và nhòe biên).", "Tốt (bảo toàn độ sắc nét của biên bước nhảy)."],
        ["Độ phức tạp thời gian", "O(H · W · k^2) phép tính cộng/nhân.", "O(H · W · k^2 · log(k)) do chi phí sắp xếp tại mỗi pixel."],
    ]
    add_styled_table(
        doc,
        headers=filter_comp_headers,
        rows=filter_comp_rows,
        caption="Bảng 2.2. Bảng đối sánh lý thuyết giữa Bộ lọc Trung bình và Bộ lọc Trung vị",
        col_aligns=[WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT],
    )

    # 2.1.5
    add_h3(doc, "2.1.5. Toán tử phát hiện biên Sobel và Các độ đo đánh giá chất lượng ảnh")
    add_para(
        doc,
        "Biên ảnh tương ứng với các vị trí có sự biến thiên cường độ sáng đột ngột, được đo bằng vector Gradient ∇f(x, y) = [G_x, G_y]^T. Toán tử Sobel 3x3 kết hợp lấy sai phân trung tâm theo một trục và làm trơn theo trục vuông góc với trọng số [1, 2, 1] thông qua cặp mặt nạ đạo hàm S_x, S_y:"
    )
    add_formula(
        doc,
        "S_x = [ [-1, 0, +1], [-2, 0, +2], [-1, 0, +1] ],    S_y = [ [-1, -2, -1], [0, 0, 0], [+1, +2, +1] ]",
        eq_number="2.9",
    )
    add_formula(
        doc,
        "G(x, y) = √( G_x(x, y)^2 + G_y(x, y)^2 ),     E(x, y) = 1 nếu G(x, y) ≥ T_edge ngược lại 0",
        eq_number="2.10",
    )
    add_para(
        doc,
        "Trong đề tài này, ngưỡng cắt biên T_edge của mỗi ảnh được xác định tự động theo phân vị thứ 85 (85th percentile) của độ lớn Gradient trên ảnh gốc sạch (top 15% điểm biên mạnh nhất) và giữ cố định cho ảnh nhiễu cũng như mọi ảnh sau lọc để so sánh công bằng."
    )
    add_para(
        doc,
        "Để đánh giá khách quan cả chất lượng mức xám toàn cục lẫn khả năng bảo toàn biên sau lọc, nhóm áp dụng hệ thống 5 chỉ số định lượng sau (với ảnh gốc f, ảnh sau xử lý g kích thước M × N, MAX_I = 255):"
    )
    add_formula(
        doc,
        "MSE(f, g) = (1 / (M·N)) ∑_{x,y} [f(x,y) - g(x,y)]^2 ,     PSNR = 10 · log_10( 255^2 / MSE )  (dB)",
        eq_number="2.11",
    )
    add_formula(
        doc,
        "Edge-MSE = (1 / (M·N)) ∑_{x,y} [G_orig(x,y) - G_filt(x,y)]^2",
        eq_number="2.12",
    )
    add_formula(
        doc,
        "Precision = TP / (TP + FP),    Recall = TP / (TP + FN),    Edge-F1 = (2 · Precision · Recall) / (Precision + Recall)",
        eq_number="2.13",
    )
    add_para(
        doc,
        "Trong đó, TP là số điểm biên đúng trùng với ảnh gốc, FP là số điểm biên giả sinh ra do nhiễu sót hoặc nhòe biên, FN là số điểm biên thật bị mất sau lọc. Edge-F1 ∈ [0, 1] càng gần 1 chứng tỏ bộ lọc vừa triệt tiêu tốt biên giả vừa không làm đứt gãy biên thật."
    )

    # 2.2
    add_h2(doc, "2.2. Công cụ, Ngôn ngữ và Thư viện lập trình")
    add_para(
        doc,
        "Hệ thống được phát triển bằng ngôn ngữ Python 3 (phiên bản 3.12.9) trên hệ điều hành Windows 11 64-bit. Nhằm tuân thủ nghiêm ngặt quy định không lạm dụng hàm lọc có sẵn, vai trò và phạm vi cho phép của từng công cụ, thư viện được phân định rạch ròi tại Bảng 2.3."
    )

    tool_headers = ["Công cụ / Thư viện", "Phiên bản", "Vai trò và phạm vi sử dụng trong đề tài"]
    tool_rows = [
        ["Python 3", "3.12.9", "Ngôn ngữ lập trình chính, tổ chức kiến trúc module hóa (src/), kịch bản thực nghiệm và máy chủ Web."],
        ["NumPy", "2.5.4", "Nền tảng tính toán ma trận ndarray; cài đặt 100% thuật toán cốt lõi (padding, noise, convolve2d, mean/median filter, Sobel, metrics) bằng vector hóa."],
        ["OpenCV (cv2)", "5.0.0", "Chỉ dùng để đọc/ghi ảnh PNG, resize ảnh và làm chuẩn đối chiếu sai số (cv2.blur, cv2.medianBlur, cv2.Sobel) trong module kiểm thử verification.py."],
        ["Pillow (PIL)", "13.0.0", "Hỗ trợ giải mã ảnh upload từ người dùng và mã hóa ma trận ảnh NumPy sang chuỗi PNG Base64 trong web_app.py."],
        ["scikit-image", "skimage.data", "Cung cấp bộ ảnh chuẩn quốc tế (camera, coins, astronaut, page, checkerboard, brick) trong bước khởi tạo dữ liệu."],
        ["Matplotlib", "3.11.2", "Vẽ sơ đồ quy trình, lưới ảnh đối sánh trước/sau lọc, bản đồ biên Sobel và các biểu đồ PSNR, Edge-F1, thời gian thực thi."],
        ["FastAPI & Uvicorn", "0.162.2 / 0.54.0", "Xây dựng RESTful API backend và máy chủ ASGI chạy ứng dụng Web tương tác trên Localhost (cổng 8000)."],
    ]
    add_styled_table(
        doc,
        headers=tool_headers,
        rows=tool_rows,
        caption="Bảng 2.3. Tổng hợp ngôn ngữ lập trình, thư viện và phạm vi sử dụng trong đề tài",
        col_aligns=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT],
    )


def write_chapter_3(doc: docx.Document, fig_dir: str) -> None:
    """CHƯƠNG 3: PHÂN TÍCH, THIẾT KẾ & XÂY DỰNG CHƯƠNG TRÌNH (Tiêu chí 2.3, 2.4, 2.5 Phiếu chấm)."""
    add_h1(doc, "CHƯƠNG 3: PHÂN TÍCH, THIẾT KẾ & XÂY DỰNG CHƯƠNG TRÌNH", page_break_before=True)

    # 3.1
    add_h2(doc, "3.1. Quy trình (Luồng xử lý) của chương trình")
    add_h3(doc, "3.1.1. Sơ đồ khối luồng xử lý tổng thể (Pipeline)")
    add_para(
        doc,
        "Luồng xử lý dữ liệu của chương trình thực nghiệm được thiết kế khép kín, tự động hóa hoàn toàn từ khâu nạp dữ liệu đến khâu đánh giá chỉ số và xuất kết quả. Hình 3.1 mô tả sơ đồ khối tổng thể (Pipeline) của chương trình."
    )

    pipeline_img = os.path.join(fig_dir, "ch2_pipeline_diagram.png")
    if os.path.exists(pipeline_img):
        add_figure(
            doc,
            pipeline_img,
            caption="Hình 3.1. Sơ đồ khối quy trình xử lý tổng thể (Pipeline) của chương trình thực nghiệm",
            width_inches=6.0,
        )

    add_h3(doc, "3.1.2. Giải thích chi tiết nhiệm vụ và định dạng dữ liệu qua từng bước")
    add_bullet(
        doc,
        "Đọc 6 ảnh chuẩn từ thư mục data/original/ (hoặc trích xuất từ skimage.data / bộ tạo dự phòng nếu chưa có trên đĩa), chuyển về ảnh xám 1 kênh chuẩn hóa kích thước (256, 256), kiểu np.uint8. Đồng thời, chạy kiểm chứng độ chính xác các hàm tự cài đặt so với OpenCV trên ảnh cameraman.",
        bold_prefix="Bước 1 – Chuẩn bị dữ liệu & Kiểm chứng thuật toán: "
    )
    add_bullet(
        doc,
        "Chạy toán tử Sobel 3x3 (chế độ đệm 'reflect') trên ảnh gốc sạch để thu được ma trận độ lớn Gradient tham chiếu orig_grad_mag (float64), bản đồ biên nhị phân chuẩn orig_bin_edge (uint8 ∈ {0, 1}) và ngưỡng chuẩn ref_threshold (giá trị vô hướng float tại phân vị thứ 85).",
        bold_prefix="Bước 2 – Trích xuất biên chuẩn tham chiếu (Ground-Truth): "
    )
    add_bullet(
        doc,
        "Khởi tạo bộ sinh ngẫu nhiên np.random.default_rng(seed=42) cho mỗi lần thêm nhiễu để tạo 6 phiên bản ảnh nhiễu uint8 (256, 256): Muối tiêu (1%, 5%, 10%) và Gaussian (σ = 10, 25, 50). Tính sẵn các chỉ số MSE, PSNR và Edge-F1 của ảnh nhiễu trước khi lọc.",
        bold_prefix="Bước 3 – Giả lập nhiễu có kiểm soát: "
    )
    add_bullet(
        doc,
        "Với mỗi ảnh nhiễu và mỗi kích thước cửa sổ k ∈ {3, 5, 7}, thực hiện mở rộng biên pad_image(..., mode='reflect') thành ma trận (256+2p, 256+2p), sau đó chạy song song Bộ lọc Trung bình (tích chập 2D trên float64 rồi ép về uint8) và Bộ lọc Trung vị (sắp xếp trượt trên uint8), đồng thời bấm giờ chính xác bằng time.perf_counter() (đơn vị ms).",
        bold_prefix="Bước 4 – Mở rộng biên và Khử nhiễu song song: "
    )
    add_bullet(
        doc,
        "Đưa từng ảnh sau lọc qua toán tử Sobel 3x3 sử dụng lại đúng ngưỡng chuẩn ref_threshold của ảnh gốc; tính MSE, PSNR (so với ảnh gốc) và Edge-MSE, Precision, Recall, Edge-F1 (so với biên ảnh gốc).",
        bold_prefix="Bước 5 – Hậu xử lý biên Sobel và Đo lường chất lượng: "
    )
    add_bullet(
        doc,
        "Tổng hợp toàn bộ 216 bản ghi chi tiết và 36 bản ghi trung bình thành tệp .CSV trong results/tables/, xuất hệ thống lưới ảnh đối sánh và biểu đồ .PNG trong results/figures/; hoặc mã hóa Base64 PNG trả về JSON cho giao diện Web.",
        bold_prefix="Bước 6 – Lưu trữ, Trực quan hóa và Phản hồi Web: "
    )

    # 3.2
    add_h2(doc, "3.2. Giải pháp kỹ thuật chi tiết")

    add_h3(doc, "3.2.1. Giải pháp xử lý biên ảnh (Padding) và chuẩn hóa kiểu dữ liệu (tránh tràn số uint8)")
    add_para(
        doc,
        "Khi lọc với cửa sổ k × k (k ∈ {3, 5, 7}), bán kính đệm p = (k - 1) // 2 lần lượt bằng 1, 2 và 3 pixel. Hàm pad_image(image, pad_h, pad_w, mode) trong src/padding.py cấp phát ma trận P kích thước (H + 2p, W + 2p) và sao chép ảnh gốc vào vùng trung tâm P[p:p+H, p:p+W]. Nhóm cài đặt đủ 3 chế độ ('zero', 'replicate', 'reflect') nhưng chọn 'reflect' làm mặc định trong mọi thực nghiệm chính vì phản chiếu gương qua đường biên giúp triệt tiêu hiện tượng bước nhảy sáng về 0 và loại bỏ biên giả ở 4 cạnh ảnh khi chạy Sobel (Hình 3.2)."
    )

    pad_img = os.path.join(fig_dir, "ch2_padding_modes.png")
    if os.path.exists(pad_img):
        add_figure(
            doc,
            pad_img,
            caption="Hình 3.2. Minh họa trực quan 3 chế độ mở rộng biên (Zero, Replicate, Reflect) với bán kính đệm p = 1",
            width_inches=5.9,
        )

    add_para(
        doc,
        "Một kỹ thuật lập trình tối quan trọng trong xử lý ảnh bằng NumPy là kiểm soát kiểu dữ liệu (dtype) để tránh lỗi tràn số nguyên không dấu 8-bit (uint8 overflow/underflow). Nếu thực hiện phép cộng nhiễu Gaussian, phép nhân tích chập, phép trừ đạo hàm Sobel (có giá trị âm) hoặc bình phương hiệu (f - g)^2 trực tiếp trên mảng uint8, giá trị vượt quá 255 hoặc nhỏ hơn 0 sẽ bị quay vòng modulo 256 làm sai hỏng hoàn toàn kết quả. Do đó, mọi phép tính số học trong src/noise.py, src/filters.py, src/sobel.py và src/metrics.py đều ép kiểu tường minh sang số thực dấu phẩy động độ chính xác kép (np.float64), sau đó mới làm tròn (np.rint) và cắt ngưỡng về đoạn [0, 255] bằng np.clip trước khi ép ngược về np.uint8."
    )

    add_code_block(
        doc,
        [
            "Thuật toán 3.1: Mở rộng biên ảnh 2 chiều (pad_image trong src/padding.py)",
            "-------------------------------------------------------------------------",
            "Đầu vào : Ma trận ảnh xám 2D I (H x W), bán kính đệm p, chế độ mode.",
            "Đầu ra  : Ma trận ảnh đã đệm P kích thước (H + 2p) x (W + 2p).",
            "1. Khởi tạo P = zeros((H + 2p, W + 2p), dtype=I.dtype); đặt I vào P[p:p+H, p:p+W].",
            "2. Nếu mode == 'zero': trả về P.",
            "3. Nếu mode in ('replicate', 'edge'): sao chép hàng 0 và hàng H-1 lên vùng đệm trên/dưới;",
            "   sau đó sao chép cột p và cột p+W-1 ra toàn bộ đệm trái/phải (gồm 4 góc).",
            "4. Nếu mode == 'reflect': với i = 0..p-1 lấy đối xứng src_top = min(p - i, H - 1),",
            "   src_bot = max(H - 2 - i, 0) cho viền trên/dưới; tương tự với trục cột cho viền trái/phải.",
            "5. Trả về P.",
        ],
        caption="Mã giả 3.1. Thuật toán mở rộng biên ảnh 2 chiều tự cài đặt trong src/padding.py",
    )

    add_h3(doc, "3.2.2. Kỹ thuật tối ưu hóa tốc độ bằng Vector hóa NumPy và Phân tích độ phức tạp")
    add_para(
        doc,
        "Nếu cài đặt phép tích chập 2D và lọc Trung vị bằng 4 vòng lặp for lồng nhau thuần Python trên ảnh 256×256, thời gian xử lý một ảnh sẽ kéo dài tới hàng trăm mili-giây đến vài giây do chi phí thông dịch vòng lặp. Để tối ưu tốc độ thực thi mà vẫn giữ nguyên 100% thuật toán tự cài đặt, nhóm áp dụng kỹ thuật vector hóa mảng (Vectorization) kết hợp khung nhìn cửa sổ trượt không sao chép bộ nhớ (zero-copy memory view) trong NumPy:"
    )
    add_bullet(
        doc,
        "Hàm numpy.lib.stride_tricks.sliding_window_view(padded, window_shape=(k, k)) tạo một tensor 4 chiều kích thước (H, W, k, k) ánh xạ trực tiếp tới vùng nhớ của ảnh đệm bằng phép tính bước nhảy (strides) mà không nhân bản dữ liệu.",
        bold_prefix="Khung nhìn cửa sổ trượt 4D: "
    )
    add_bullet(
        doc,
        "Thay thế 4 vòng lặp tích chập bằng phép co rút chỉ số Einstein np.einsum('ijkl,kl->ij', windows, flipped_kernel), tính đồng thời toàn bộ H×W tích vô hướng trên tập lệnh vector C/SIMD của CPU.",
        bold_prefix="Tích chập vector hóa (convolve2d): "
    )
    add_bullet(
        doc,
        "Biến đổi tensor cửa sổ (H, W, k, k) thành mảng 3 chiều (H, W, k^2), rồi gọi np.sort(flat_windows, axis=-1) dọc theo trục cuối cùng (sử dụng thuật toán Introsort) và lấy lát cắt [:, :, k^2 // 2].",
        bold_prefix="Lọc trung vị vector hóa (median_filter): "
    )
    add_para(
        doc,
        "Về độ phức tạp tính toán lý thuyết: Bộ lọc Trung bình có độ phức tạp O(H · W · k^2) phép cộng/nhân, trong khi Bộ lọc Trung vị có độ phức tạp O(H · W · k^2 · log(k)) do phải sắp xếp k^2 phần tử tại mỗi trong H·W điểm ảnh. Độ phức tạp bộ nhớ trung gian của tensor view/mảng sắp xếp là O(H · W · k^2)."
    )

    add_h3(doc, "3.2.3. Thiết kế chi tiết các thuật toán cốt lõi (Sinh nhiễu, Bộ lọc, Sobel & Độ đo)")
    add_para(
        doc,
        "Dưới đây là mã giả đặc tả chính xác các thuật toán đã được hiện thực trong src/noise.py, src/filters.py, src/sobel.py và src/metrics.py."
    )

    add_code_block(
        doc,
        [
            "Thuật toán 3.2: Sinh nhiễu Muối tiêu & Gaussian có seed cố định (src/noise.py)",
            "-------------------------------------------------------------------------",
            "A. ADD_SALT_AND_PEPPER_NOISE(I, density d, salt_vs_pepper=0.5, seed=42):",
            "   1. rng <- default_rng(seed); R <- rng.random(I.shape).",
            "   2. p_salt <- d * salt_vs_pepper; p_pepper <- d * (1 - salt_vs_pepper).",
            "   3. G <- copy(I).uint8; G[R < p_salt] <- 255; G[p_salt <= R < p_salt + p_pepper] <- 0.",
            "   4. Trả về G.",
            "B. ADD_GAUSSIAN_NOISE(I, sigma, mean=0.0, seed=42):",
            "   1. rng <- default_rng(seed); Eta <- rng.normal(loc=mean, scale=sigma, size=I.shape).",
            "   2. Trả về clip(rint(float64(I) + Eta), 0, 255).astype(uint8).",
        ],
        caption="Mã giả 3.2. Thuật toán sinh nhiễu Muối tiêu và nhiễu Gaussian trong src/noise.py",
    )

    add_code_block(
        doc,
        [
            "Thuật toán 3.3: Tích chập 2D, Bộ lọc Trung bình và Bộ lọc Trung vị (src/filters.py)",
            "-------------------------------------------------------------------------",
            "A. CONVOLVE2D(I, K (k_h x k_w), pad_mode='reflect'):",
            "   1. K_flip <- flip(flip(K, axis=0), axis=1).astype(float64).",
            "   2. P <- pad_image(float64(I), k_h // 2, k_w // 2, mode=pad_mode).",
            "   3. W_view <- sliding_window_view(P, (k_h, k_w)).",
            "   4. Trả về einsum('ijkl,kl->ij', W_view, K_flip) (kiểu float64).",
            "B. MEAN_FILTER(I, k, pad_mode='reflect'):",
            "   1. K_mean <- ones((k, k), float64) / (k * k).",
            "   2. O <- CONVOLVE2D(I, K_mean, pad_mode).",
            "   3. Trả về clip(rint(O), 0, 255).astype(uint8).",
            "C. MEDIAN_FILTER(I, k, pad_mode='reflect'):",
            "   1. P <- pad_image(I, k // 2, k // 2, mode=pad_mode).",
            "   2. flat <- sliding_window_view(P, (k, k)).reshape(H, W, k * k).",
            "   3. Trả về sort(flat, axis=-1)[:, :, (k * k) // 2].astype(uint8).",
        ],
        caption="Mã giả 3.3. Thuật toán tích chập 2D, Bộ lọc Trung bình và Bộ lọc Trung vị trong src/filters.py",
    )

    add_code_block(
        doc,
        [
            "Thuật toán 3.4: Toán tử Sobel 3x3 & Đánh giá bảo toàn biên (src/sobel.py, src/metrics.py)",
            "-------------------------------------------------------------------------",
            "A. SOBEL_EDGE_DETECTION(I, threshold=None, pad_mode='reflect'):",
            "   1. Đảo chiều sẵn S_x, S_y (vì convolve2d sẽ lật 180° lại) để khớp công thức tương quan.",
            "   2. G_x <- CONVOLVE2D(I, flip180(S_x), pad_mode); G_y <- CONVOLVE2D(I, flip180(S_y), pad_mode).",
            "   3. grad_mag <- sqrt(G_x^2 + G_y^2); grad_vis <- uint8(clip(rint(grad_mag), 0, 255)).",
            "   4. Nếu threshold là None: T_edge <- percentile(grad_mag, 85.0); ngược lại T_edge <- threshold.",
            "   5. binary_edge <- (grad_mag >= T_edge).astype(uint8).",
            "   6. Trả về (grad_mag, grad_vis, binary_edge, (G_x, G_y), T_edge).",
            "B. EVALUATE_EDGE_PRESERVATION(ref_mag, ref_bin, test_mag, test_bin):",
            "   1. Cắt ngưỡng ref_mag, test_mag về [0, 255]; tính Edge-MSE <- calculate_mse(ref_norm, test_norm).",
            "   2. Tính TP, FP, FN trên ref_bin và test_bin -> Precision, Recall, F1-Score.",
        ],
        caption="Mã giả 3.4. Thuật toán phát hiện biên Sobel và đánh giá bảo toàn biên",
    )

    add_h3(doc, "3.2.4. Xử lý các trường hợp ngoại lệ (Edge Cases & Exception Handling)")
    add_para(
        doc,
        "Chương trình cài đặt đầy đủ các chốt kiểm tra đầu vào và xử lý ngoại lệ nhằm bảo đảm ổn định khi chạy thực nghiệm lẫn khi người dùng thao tác tự do trên Web:"
    )
    add_bullet(
        doc,
        "Các hàm trong src/padding.py và src/filters.py kiểm tra kích thước image.ndim != 2 (bắt buộc ảnh đơn kênh 2D), kiểm tra kernel_size % 2 == 0 hoặc < 1 (chỉ nhận cửa sổ lẻ dương) và kiểm tra mode hợp lệ trong {'zero', 'replicate', 'reflect'}. Trên web_app.py, mọi ảnh màu RGB/RGBA hoặc khác kích thước khi người dùng tải lên đều được Pillow chuyển về chế độ 'L' (ảnh xám 8-bit) và chuẩn hóa kích thước bằng cv2.resize(..., interpolation=cv2.INTER_AREA).",
        bold_prefix="Ảnh khác kích thước, đa kênh vs đơn kênh và tham số kernel: "
    )
    add_bullet(
        doc,
        "Trong src/metrics.py, khi ảnh gốc và ảnh xử lý giống hệt nhau (mse == 0.0, ví dụ khi nhấn nút Reset Ảnh Gốc trên Web), hàm calculate_psnr trả về float('inf') thay vì gây lỗi chia cho 0 (ZeroDivisionError), và Web UI hiển thị ký hiệu '∞'. Tương tự, evaluate_edge_preservation gán Precision = 0.0 nếu TP + FP == 0 (ví dụ trường hợp ảnh text bị làm mờ mất hết biên ở cửa sổ 7x7) và Recall = 0.0 nếu TP + FN == 0.",
        bold_prefix="Tránh lỗi chia cho 0 trong PSNR và Edge-F1: "
    )
    add_bullet(
        doc,
        "Trong src/dataset.py, nếu máy đích chưa cài đặt scikit-image, hàm load_or_prepare_dataset tự động kích hoạt _create_fallback_standard_images(size=256) để sinh bộ 6 ảnh xám chuẩn tổng hợp bảo đảm chương trình luôn chạy thông suốt; khi upload ảnh rỗng/lỗi định dạng trên Web, endpoint trả HTTP 400 kèm thông báo lỗi JSON.",
        bold_prefix="Dự phòng thiếu dữ liệu và tệp upload không hợp lệ: "
    )

    add_h3(doc, "3.2.5. Phân định phần thuật toán tự cài đặt (From Scratch) và phần gọi thư viện")
    add_para(
        doc,
        "Theo khuyến nghị của đề tài nhằm làm rõ đóng góp kỹ thuật cốt lõi khi bảo vệ vấn đáp, Bảng 3.1 phân tách minh bạch phần thuật toán nhóm tự lập trình từ đầu và phần chỉ sử dụng thư viện hỗ trợ."
    )

    scratch_headers = ["Thành phần / Tác vụ", "Phần tự cài đặt (From Scratch - NumPy)", "Phần sử dụng thư viện hỗ trợ"]
    scratch_rows = [
        ["Mở rộng biên ảnh", "pad_image() với 3 chế độ zero, replicate, reflect (src/padding.py).", "Không dùng cv2.copyMakeBorder hay np.pad."],
        ["Sinh nhiễu ngẫu nhiên", "add_salt_and_pepper_noise(), add_gaussian_noise() với default_rng(42) (src/noise.py).", "Không dùng skimage.util.random_noise."],
        ["Tích chập 2D & Bộ lọc", "convolve2d(), mean_filter(), median_filter() dùng sliding_window_view, einsum, sort (src/filters.py).", "Không dùng cv2.blur, cv2.medianBlur, scipy.ndimage."],
        ["Phát hiện biên Sobel", "sobel_edge_detection() tái sử dụng convolve2d() tự viết & ngưỡng P85 (src/sobel.py).", "Không dùng cv2.Sobel trong luồng chính."],
        ["Tính các độ đo chất lượng", "calculate_mse(), calculate_psnr(), evaluate_edge_preservation() (src/metrics.py).", "Không dùng skimage.metrics hay sklearn.metrics."],
        ["Đọc/Ghi, Kiểm chứng & Web UI", "Cấu trúc logic điều phối (run_experiments.py) và xử lý luồng API (web_app.py).", "OpenCV/Pillow (đọc/ghi/resize/Base64), verification.py (gọi cv2 đối chiếu), Matplotlib (vẽ hình), FastAPI (Web)."],
    ]
    add_styled_table(
        doc,
        headers=scratch_headers,
        rows=scratch_rows,
        caption="Bảng 3.1. Phân định phần thuật toán tự cài đặt (From Scratch) và phần sử dụng thư viện hỗ trợ",
        col_aligns=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT],
    )

    # 3.3
    add_h2(doc, "3.3. Kiến trúc mã nguồn và Tổ chức dữ liệu")
    add_h3(doc, "3.3.1. Cấu trúc thư mục dự án và chức năng các module")
    add_para(
        doc,
        "Dự án BTL_XLA được tổ chức theo kiến trúc module hóa tách biệt rõ ràng giữa lõi thuật toán (src/), dữ liệu ảnh gốc (data/), kịch bản chạy thực nghiệm (run_experiments.py), ứng dụng Web (web_app.py), script hỗ trợ (scripts/) và kết quả đầu ra (results/). Hình 3.3 trình bày cây thư mục đầy đủ của dự án."
    )

    add_code_block(
        doc,
        [
            "BTL_XLA/",
            "├── BTL_XLA.docx                 # Báo cáo Bài tập lớn hoàn chỉnh Nhóm 10",
            "├── requirements.txt             # Danh sách các gói thư viện Python",
            "├── run_experiments.py           # Script điều phối thực nghiệm tự động",
            "├── web_app.py                   # Ứng dụng Web Localhost (cổng 8000)",
            "├── src/                         # Package chứa thuật toán tự cài đặt",
            "│   ├── __init__.py              # Khởi tạo package src",
            "│   ├── padding.py               # Mở rộng biên (zero, replicate, reflect)",
            "│   ├── noise.py                 # Sinh nhiễu Muối tiêu & Gaussian (seed=42)",
            "│   ├── filters.py               # Tích chập 2D, Lọc Trung bình & Trung vị",
            "│   ├── sobel.py                 # Toán tử Sobel 3x3 và ngưỡng hóa P85",
            "│   ├── metrics.py               # MSE, PSNR, Edge-MSE, Precision/Recall/F1",
            "│   ├── dataset.py               # Chuẩn bị 6 ảnh xám chuẩn 256x256",
            "│   ├── verification.py          # Kiểm chứng đối chiếu sai số với OpenCV",
            "│   └── visualization.py         # Xuất lưới ảnh đối sánh và biểu đồ (.PNG)",
            "├── data/original/               # Lưu 6 ảnh xám gốc chuẩn (256x256, PNG)",
            "├── scripts/                     # Script vẽ sơ đồ, chụp Web UI, xuất DOCX",
            "└── results/",
            "    ├── tables/                  # 3 bảng kết quả .CSV",
            "    └── figures/                 # Ảnh đối sánh, bản đồ Sobel & biểu đồ",
        ],
        caption="Hình 3.3. Sơ đồ tổ chức cây thư mục mã nguồn và dữ liệu của dự án BTL_XLA",
        caption_style="Hình",
        font_size=9.5,
    )

    add_h3(doc, "3.3.2. Kiến trúc phần mềm và Giao diện người dùng Web Localhost")
    add_para(
        doc,
        "Kiến trúc phần mềm tuân theo mô hình phân lớp 3 tầng: Tầng Lõi thuật toán (src/) đóng vai trò thư viện xử lý thuần NumPy độc lập; Tầng Điều phối (run_experiments.py cho xử lý hàng loạt CLI và web_app.py cho xử lý tương tác trực tuyến); và Tầng Trình bày (bảng CSV/PNG và HTML5 Dashboard)."
    )
    add_para(
        doc,
        "Ứng dụng Web tại web_app.py sử dụng FastAPI + Uvicorn (http://127.0.0.1:8000) với 5 RESTful API endpoints: GET '/' (giao diện Dashboard), POST '/api/select_sample' (chọn ảnh mẫu), POST '/api/upload_image' (tải ảnh ngoài và chuẩn hóa 256x256 hoặc 512x512), POST '/api/process' (thêm nhiễu, lọc, chạy Sobel và tính chỉ số) và POST '/api/reset' (xóa nhiễu khôi phục ảnh gốc sạch). Hình 3.4 minh họa giao diện tổng thể của ứng dụng Web."
    )

    web_fig1 = os.path.join(fig_dir, "web_ui_salt_pepper.png")
    if os.path.exists(web_fig1):
        add_figure(
            doc,
            web_fig1,
            caption="Hình 3.4. Giao diện tổng thể của ứng dụng Web tương tác chạy trên Localhost (http://127.0.0.1:8000)",
            width_inches=6.1,
        )

    add_h3(doc, "3.3.3. Cấu trúc dữ liệu lưu trữ và trao đổi (Tensor ảnh, CSV, JSON API)")
    add_para(
        doc,
        "Để bảo đảm sự nhất quán giữa các module tính toán, lưu trữ ngoại tuyến và giao tiếp Web API, các cấu trúc dữ liệu trong chương trình được chuẩn hóa như Bảng 3.2."
    )

    ds_struct_headers = ["Đối tượng dữ liệu", "Cấu trúc / Định dạng", "Kích thước (Shape) & Kiểu (Dtype) / Các trường chính"]
    ds_struct_rows = [
        ["Ảnh gốc, Ảnh nhiễu, Ảnh sau lọc", "NumPy ndarray (2D)", "Shape: (H, W) = (256, 256) (hoặc 512x512 trên Web); Dtype: uint8 ∈ [0, 255]."],
        ["Ảnh mở rộng biên (Padding)", "NumPy ndarray (2D)", "Shape: (H + 2p, W + 2p) với p ∈ {1, 2, 3}; Dtype: uint8 hoặc float64 (trong convolve2d)."],
        ["Cửa sổ trượt (Sliding Window)", "NumPy ndarray view (4D / 3D)", "4D view (H, W, k, k) cho tích chập; reshape thành 3D (H, W, k^2) cho hàm sắp xếp lọc trung vị."],
        ["Bản đồ Gradient & Biên Sobel", "Tuple (ndarray, uint8, bin)", "gx, gy, grad_mag: (H, W) float64; grad_mag_uint8: (H, W) uint8; binary_edge: (H, W) uint8 ∈ {0, 1}."],
        ["Bảng verification_results.csv", "CSV (UTF-8-SIG), 7 dòng", "Cột: algorithm, reference, max_abs_error, mean_abs_error, status."],
        ["Bảng detailed_results.csv", "CSV (UTF-8-SIG), 216 dòng", "14 cột: image, noise_type, noise_level, noisy_mse, noisy_psnr, filter_type, kernel_size, mse, psnr, edge_mse, edge_precision, edge_recall, edge_f1, time_ms."],
        ["Bảng summary_by_noise_and_filter.csv", "CSV (UTF-8-SIG), 36 dòng", "12 cột trung bình trên 6 ảnh theo từng (noise_type, noise_level, filter_type, kernel_size)."],
        ["Trạng thái & JSON API Web", "Dict STATE & JSON Response", "STATE lưu dataset, original_img, working_noisy_img, noise_info; JSON trả shape, sobel_threshold, 8 ảnh Base64 PNG và object metrics."],
    ]
    add_styled_table(
        doc,
        headers=ds_struct_headers,
        rows=ds_struct_rows,
        caption="Bảng 3.2. Đặc tả cấu trúc dữ liệu mảng đa chiều (Tensor), tệp CSV và JSON API",
        col_aligns=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT],
    )


def write_chapter_4(doc: docx.Document, tables_dir: str, fig_dir: str) -> None:
    """CHƯƠNG 4: THỰC NGHIỆM VÀ ĐÁNH GIÁ KẾT QUẢ (Tiêu chí 2.6, 2.7, 2.8, 2.9 Phiếu chấm)."""
    ver_rows = load_csv_rows(os.path.join(tables_dir, "verification_results.csv"))
    sum_rows = load_csv_rows(os.path.join(tables_dir, "summary_by_noise_and_filter.csv"))
    det_rows = load_csv_rows(os.path.join(tables_dir, "detailed_results.csv"))

    add_h1(doc, "CHƯƠNG 4: THỰC NGHIỆM VÀ ĐÁNH GIÁ KẾT QUẢ", page_break_before=True)

    # 4.1
    add_h2(doc, "4.1. Bộ dữ liệu thực nghiệm")
    add_h3(doc, "4.1.1. Môi trường phần cứng, phần mềm và cấu hình tái lập")
    add_para(
        doc,
        "Toàn bộ các kịch bản thực nghiệm định lượng và kiểm thử Web UI được thực hiện trên hệ điều hành Windows 11 64-bit, trình thông dịch Python 3.12.9 với các gói thư viện: NumPy 2.5.4 (lõi thuật toán tự cài đặt), OpenCV 5.0.0 và Pillow 13.0.0 (đọc/ghi ảnh và kiểm chứng đối chiếu), Matplotlib 3.11.2 (trực quan hóa biểu đồ), FastAPI 0.162.2 và Uvicorn 0.54.0 (ứng dụng Web Localhost). Hạt giống bộ sinh số giả ngẫu nhiên được cố định toàn cục ở giá trị SEED = 42 để bảo đảm tính tái lập 100%."
    )

    add_h3(doc, "4.1.2. Nguồn gốc và thống kê chi tiết bộ dữ liệu 6 ảnh xám chuẩn")
    add_para(
        doc,
        "Bộ dữ liệu thực nghiệm gồm 6 ảnh chuẩn quốc tế trích xuất từ kho ảnh tiêu chuẩn của thư viện scikit-image (skimage.data / USC-SIPI, kèm bộ tạo dự phòng trong src/dataset.py). Tất cả 6 ảnh được chuẩn hóa về ảnh đa mức xám đơn kênh 8-bit (uint8, dải mức xám [0, 255]), định dạng tệp PNG không mất dữ liệu, độ phân giải đồng nhất 256 × 256 pixel (65.536 điểm ảnh/ảnh). Không chia tập Train/Test vì đây là bài toán lọc không gian trực tiếp (non-learning spatial filtering). Sáu ảnh được chọn chủ đích để bao phủ đa dạng đặc trưng phân bố không gian (vùng phẳng, biên cong, biên thẳng góc vuông, nét chữ mảnh 1–2 pixel và kết cấu lặp), chi tiết tại Bảng 4.1 và trực quan hóa tại Hình 4.1."
    )

    ds_headers = ["TT", "Tên tệp ảnh", "Độ phân giải & Định dạng", "Đặc trưng phân bố không gian và mục đích kiểm thử"]
    ds_rows = [
        ["1", "cameraman.png", "256 × 256 px\n8-bit Grayscale (.png)", "Kết hợp vùng bầu trời phẳng rộng (tần số thấp), vùng tối áo khoác và các cạnh thẳng sắc nét của chân máy ảnh (tần số cao)."],
        ["2", "coins.png", "256 × 256 px\n8-bit Grayscale (.png)", "Chứa nhiều đồng xu tròn tách biệt trên nền tối đồng nhất; kiểm thử bảo toàn đường biên cong khép kín và chi tiết chạm khắc."],
        ["3", "astronaut.png", "256 × 256 px\n8-bit Grayscale (.png)", "Ảnh chân dung có dải chuyển tiếp độ sáng mượt trên khuôn mặt kết hợp các chi tiết đa tỷ lệ trên mũ bảo hộ và phù hiệu."],
        ["4", "text.png", "256 × 256 px\n8-bit Grayscale (.png)", "Trang văn bản gồm các nét chữ in mảnh (độ dày 1–2 pixel) tương phản cao trên nền sáng; phép thử khắc nghiệt nhất cho cửa sổ lớn."],
        ["5", "checkerboard.png", "256 × 256 px\n8-bit Grayscale (.png)", "Bàn cờ hình học gồm các ô sáng/tối tuần hoàn, biên bước nhảy (step edges) tuyệt đối và các góc vuông 90°."],
        ["6", "brick.png", "256 × 256 px\n8-bit Grayscale (.png)", "Kết cấu tường gạch lặp tuần hoàn với nhiều rãnh vữa hẹp và bề mặt nhám; khảo sát hiện tượng suy giảm kết cấu (texture degradation)."],
    ]
    add_styled_table(
        doc,
        headers=ds_headers,
        rows=ds_rows,
        caption="Bảng 4.1. Mô tả chi tiết bộ dữ liệu 6 ảnh xám chuẩn sử dụng trong thực nghiệm",
        col_aligns=[WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT],
    )

    ds_fig = os.path.join(fig_dir, "dataset_overview.png")
    if os.path.exists(ds_fig):
        add_figure(
            doc,
            ds_fig,
            caption="Hình 4.1. Trực quan hóa bộ dữ liệu 6 ảnh xám chuẩn kích thước 256x256 sử dụng trong thực nghiệm",
            width_inches=5.8,
        )

    # 4.2
    add_h2(doc, "4.2. Kịch bản thực nghiệm và Các độ đo đánh giá")
    add_h3(doc, "4.2.1. Thiết kế các nhóm kịch bản thực nghiệm và kiểm thử")
    add_para(
        doc,
        "Để chứng minh toàn diện tính đúng đắn của thuật toán, hiệu năng khử nhiễu và độ ổn định của ứng dụng, nhóm thiết kế 3 nhóm kịch bản thực nghiệm và kiểm thử rõ ràng (Bảng 4.2):"
    )

    sc_headers = ["Nhóm kịch bản", "Mục tiêu kiểm chứng", "Cấu hình tham số & Phạm vi thực thi"]
    sc_rows = [
        [
            "Kịch bản 1:\nKhảo sát tham số\n(Parameter Study)",
            "Đánh giá ảnh hưởng của kích thước cửa sổ lọc và cường độ nhiễu lên chất lượng ảnh & biên.",
            "- 6 ảnh chuẩn × 2 loại nhiễu × 3 mức nhiễu (S&P: 1%, 5%, 10%; Gaussian: σ = 10, 25, 50) × 2 bộ lọc × 3 kích thước cửa sổ (3x3, 5x5, 7x7) = 216 cấu hình (108 S&P + 108 Gaussian).",
        ],
        [
            "Kịch bản 2:\nSo sánh đối chiếu\n(Comparison)",
            "(a) Kiểm chứng độ chính xác code tự cài đặt với OpenCV.\n(b) Đối sánh hiệu năng Lọc Trung bình vs Trung vị.",
            "- (a) Chạy 7 phép kiểm thử trên cameraman.png (Mean/Median 3x3, 5x5, 7x7 và Sobel 3x3) đối chiếu cv2.blur, cv2.medianBlur, cv2.Sobel.\n- (b) So sánh trực tiếp Mean vs Median trên cùng ma trận nhiễu (seed=42) về MSE, PSNR, Edge-F1 và thời gian (ms).",
        ],
        [
            "Kịch bản 3:\nKiểm thử điều kiện khắc nghiệt & Web UI",
            "(a) Phát hiện giới hạn và trường hợp thất bại (Failure cases).\n(b) Kiểm thử toàn bộ chức năng ứng dụng Web.",
            "- (a) Thử nghiệm nhiễu nặng (S&P 10%, Gaussian σ = 50) và ảnh nét mảnh/góc vuông (text, checkerboard) ở cửa sổ 5x5, 7x7.\n- (b) Thực thi 6 ca kiểm thử (TC-01 đến TC-06) trên Localhost: chọn ảnh, upload ảnh ngoài, kéo slider nhiễu, đổi kernel/padding, Reset và tải ảnh.",
        ],
    ]
    add_styled_table(
        doc,
        headers=sc_headers,
        rows=sc_rows,
        caption="Bảng 4.2. Tổng hợp các kịch bản thực nghiệm và kiểm thử (Test Cases)",
        col_aligns=[WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT],
    )

    add_h3(doc, "4.2.2. Hệ thống độ đo định lượng (Metrics)")
    add_para(
        doc,
        "Chương trình đánh giá đồng thời 3 khía cạnh bằng các công thức đã trình bày tại Mục 2.1.5: (1) Chất lượng khôi phục mức xám toàn cục qua MSE (càng nhỏ càng tốt) và PSNR tính bằng dB (càng cao càng tốt); (2) Khả năng bảo toàn biên Sobel tại ngưỡng cố định P85 của ảnh gốc qua Edge-MSE, Edge-Precision, Edge-Recall và Edge-F1 (càng gần 1 càng tốt); và (3) Hiệu năng tính toán qua thời gian xử lý trung bình (ms) đo bằng time.perf_counter()."
    )

    # 4.3
    add_h2(doc, "4.3. Kết quả thực nghiệm")

    # 4.3.1
    add_h3(doc, "4.3.1. Kết quả kiểm chứng tính đúng đắn của thuật toán tự cài đặt so với OpenCV")
    add_para(
        doc,
        "Kết quả chạy module kiểm thử độc lập src/verification.py trên ảnh chuẩn cameraman.png được ghi nhận tại Bảng 4.3. Cả 7 phép đối chiếu giữa hàm tự viết bằng NumPy và hàm chuẩn OpenCV (trong cùng chế độ mở rộng biên tương ứng: BORDER_REFLECT_101 cho Mean/Sobel và BORDER_REPLICATE cho Median) đều đạt sai số cực đại bằng 0.0 và sai số trung bình bằng 0.0 (khớp tuyệt đối 100%)."
    )

    ver_headers = ["Thuật toán tự cài đặt (NumPy)", "Hàm chuẩn đối chiếu (OpenCV)", "Sai số cực đại", "Sai số trung bình", "Kết luận"]
    ver_table_rows = [
        [r["algorithm"], r["reference"], r["max_abs_error"], r["mean_abs_error"], r["status"]]
        for r in ver_rows
    ]
    add_styled_table(
        doc,
        headers=ver_headers,
        rows=ver_table_rows,
        caption="Bảng 4.3. Kết quả kiểm chứng độ chính xác của các thuật toán tự cài đặt so với thư viện OpenCV",
        col_aligns=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER],
    )

    # 4.3.2
    add_h3(doc, "4.3.2. Kết quả thực nghiệm trên nhiễu Muối tiêu (1%, 5%, 10%)")
    add_para(
        doc,
        "Trên bộ 6 ảnh chuẩn bị nhiễu Muối tiêu ở 3 mức mật độ 1%, 5% và 10%, PSNR trung bình của ảnh nhiễu ban đầu lần lượt là 24.59 dB, 17.79 dB và 14.81 dB. Bảng 4.4 tổng hợp giá trị trung bình trên 6 ảnh, Bảng 4.5 trích xuất số liệu chi tiết trên ảnh cameraman, và các Hình 4.2, 4.3, 4.4 minh họa trực quan cùng biểu đồ xu hướng."
    )

    sp_sum_headers = [
        "Mức nhiễu\n(PSNR nhiễu)",
        "Bộ lọc",
        "Cửa sổ",
        "MSE\n(Sau lọc)",
        "PSNR\n(dB)",
        "Edge-MSE",
        "Edge\nPrec.",
        "Edge\nRecall",
        "Edge\nF1",
        "Time\n(ms)",
    ]
    dense_widths = [2.2, 1.8, 1.1, 1.5, 1.4, 1.7, 1.6, 1.6, 1.6, 1.5]
    sp_sum_rows = []
    for r in sum_rows:
        if r["noise_type"] == "Salt & Pepper":
            sp_sum_rows.append([
                f"{r['noise_level']}\n({r['noisy_psnr_mean']} dB)",
                "Trung bình" if r["filter_type"] == "Mean" else "Trung vị",
                f"{r['kernel_size']}x{r['kernel_size']}",
                r["mse_mean"],
                r["psnr_mean"],
                r["edge_mse_mean"],
                r["edge_precision_mean"],
                r["edge_recall_mean"],
                r["edge_f1_mean"],
                r["time_ms_mean"],
            ])
    add_styled_table(
        doc,
        headers=sp_sum_headers,
        rows=sp_sum_rows,
        caption="Bảng 4.4. Kết quả trung bình trên 6 ảnh chuẩn khi khử nhiễu Muối tiêu (1%, 5%, 10%)",
        col_aligns=[WD_ALIGN_PARAGRAPH.CENTER] * 10,
        col_widths_cm=dense_widths,
        font_size=10.0,
    )

    cam_sp_rows = []
    for r in det_rows:
        if r["image"] == "cameraman" and r["noise_type"] == "Salt & Pepper":
            cam_sp_rows.append([
                f"{r['noise_level']}\n({r['noisy_psnr']} dB)",
                "Trung bình" if r["filter_type"] == "Mean" else "Trung vị",
                f"{r['kernel_size']}x{r['kernel_size']}",
                r["mse"],
                r["psnr"],
                r["edge_mse"],
                r["edge_precision"],
                r["edge_recall"],
                r["edge_f1"],
                r["time_ms"],
            ])
    add_styled_table(
        doc,
        headers=sp_sum_headers,
        rows=cam_sp_rows,
        caption="Bảng 4.5. Số liệu thực nghiệm chi tiết trên ảnh 'cameraman' với nhiễu Muối tiêu",
        col_aligns=[WD_ALIGN_PARAGRAPH.CENTER] * 10,
        col_widths_cm=dense_widths,
        font_size=10.0,
    )

    for fig_name, cap in [
        ("compare_cameraman_sa_5pct.png", "Hình 4.2. So sánh trực quan kết quả khử nhiễu Muối tiêu 5% trên ảnh 'cameraman'"),
        ("compare_coins_sa_10pct.png", "Hình 4.3. So sánh trực quan kết quả khử nhiễu Muối tiêu 10% trên ảnh 'coins'"),
        ("chart_salt_pepper_psnr_f1.png", "Hình 4.4. Biểu đồ PSNR trung bình và Edge F1-Score trên nhiễu Muối tiêu theo mật độ nhiễu và kích thước cửa sổ"),
    ]:
        p = os.path.join(fig_dir, fig_name)
        if os.path.exists(p):
            add_figure(doc, p, caption=cap, width_inches=5.9)

    # 4.3.3
    add_h3(doc, "4.3.3. Kết quả thực nghiệm trên nhiễu Gaussian (σ = 10, 25, 50)")
    add_para(
        doc,
        "Ở nhóm kịch bản nhiễu cộng trắng Gaussian với độ lệch chuẩn σ = 10, 25 và 50, PSNR trung bình của ảnh nhiễu ban đầu lần lượt là 28.57 dB, 20.81 dB và 15.24 dB. Bảng 4.6 trình bày kết quả trung bình trên 6 ảnh chuẩn, Bảng 4.7 trình bày chi tiết trên ảnh cameraman, và các Hình 4.5, 4.6, 4.7 trực quan hóa kết quả."
    )

    ga_sum_rows = []
    for r in sum_rows:
        if r["noise_type"] == "Gaussian":
            ga_sum_rows.append([
                f"{r['noise_level']}\n({r['noisy_psnr_mean']} dB)",
                "Trung bình" if r["filter_type"] == "Mean" else "Trung vị",
                f"{r['kernel_size']}x{r['kernel_size']}",
                r["mse_mean"],
                r["psnr_mean"],
                r["edge_mse_mean"],
                r["edge_precision_mean"],
                r["edge_recall_mean"],
                r["edge_f1_mean"],
                r["time_ms_mean"],
            ])
    add_styled_table(
        doc,
        headers=sp_sum_headers,
        rows=ga_sum_rows,
        caption="Bảng 4.6. Kết quả trung bình trên 6 ảnh chuẩn khi khử nhiễu Gaussian (σ = 10, 25, 50)",
        col_aligns=[WD_ALIGN_PARAGRAPH.CENTER] * 10,
        col_widths_cm=dense_widths,
        font_size=10.0,
    )

    cam_ga_rows = []
    for r in det_rows:
        if r["image"] == "cameraman" and r["noise_type"] == "Gaussian":
            cam_ga_rows.append([
                f"{r['noise_level']}\n({r['noisy_psnr']} dB)",
                "Trung bình" if r["filter_type"] == "Mean" else "Trung vị",
                f"{r['kernel_size']}x{r['kernel_size']}",
                r["mse"],
                r["psnr"],
                r["edge_mse"],
                r["edge_precision"],
                r["edge_recall"],
                r["edge_f1"],
                r["time_ms"],
            ])
    add_styled_table(
        doc,
        headers=sp_sum_headers,
        rows=cam_ga_rows,
        caption="Bảng 4.7. Số liệu thực nghiệm chi tiết trên ảnh 'cameraman' với nhiễu Gaussian",
        col_aligns=[WD_ALIGN_PARAGRAPH.CENTER] * 10,
        col_widths_cm=dense_widths,
        font_size=10.0,
    )

    for fig_name, cap in [
        ("compare_cameraman_ga_sigma25.png", "Hình 4.5. So sánh trực quan kết quả khử nhiễu Gaussian (σ = 25) trên ảnh 'cameraman'"),
        ("compare_cameraman_ga_sigma50.png", "Hình 4.6. So sánh trực quan kết quả khử nhiễu Gaussian nặng (σ = 50) trên ảnh 'cameraman'"),
        ("chart_gaussian_psnr_f1.png", "Hình 4.7. Biểu đồ PSNR trung bình và Edge F1-Score trên nhiễu Gaussian theo độ lệch chuẩn σ và kích thước cửa sổ"),
    ]:
        p = os.path.join(fig_dir, fig_name)
        if os.path.exists(p):
            add_figure(doc, p, caption=cap, width_inches=5.9)

    # 4.3.4
    add_h3(doc, "4.3.4. Kết quả đánh giá bảo toàn biên Sobel và Thời gian xử lý")
    add_para(
        doc,
        "Hình 4.8 và Hình 4.9 trực quan hóa bản đồ biên Sobel thu được trên ảnh cameraman (nhiễu Muối tiêu 5% và Gaussian σ = 25), trong khi Hình 4.10 minh họa trường hợp khắc nghiệt trên ảnh văn bản text.png khi tăng kích thước cửa sổ lên 5x5 và 7x7. Tiếp đó, Bảng 4.8 và Hình 4.11 tổng hợp thời gian xử lý trung bình (ms) tính trên toàn bộ 36 kịch bản của mỗi kích thước cửa sổ."
    )

    for fig_name, cap in [
        ("sobel_cameraman_sa_5pct.png", "Hình 4.8. So sánh bản đồ biên Sobel trên ảnh 'cameraman' với nhiễu Muối tiêu 5%"),
        ("sobel_cameraman_ga_sigma25.png", "Hình 4.9. So sánh bản đồ biên Sobel trên ảnh 'cameraman' với nhiễu Gaussian (σ = 25)"),
        ("compare_text_sa_5pct.png", "Hình 4.10. Minh họa hiện tượng mất nét chữ mảnh trên ảnh 'text' khi tăng kích thước cửa sổ lọc lên 5x5 và 7x7"),
    ]:
        p = os.path.join(fig_dir, fig_name)
        if os.path.exists(p):
            add_figure(doc, p, caption=cap, width_inches=5.9)

    time_table_rows = []
    time_stats = {}
    for k in (3, 5, 7):
        m_times = [float(r["time_ms_mean"]) for r in sum_rows if r["filter_type"] == "Mean" and int(r["kernel_size"]) == k]
        med_times = [float(r["time_ms_mean"]) for r in sum_rows if r["filter_type"] == "Median" and int(r["kernel_size"]) == k]
        avg_m = sum(m_times) / len(m_times)
        avg_med = sum(med_times) / len(med_times)
        ratio = avg_med / avg_m if avg_m > 0 else 0.0
        time_stats[k] = (avg_m, avg_med, ratio)
        time_table_rows.append([
            f"{k} x {k}",
            f"{k * k} phần tử",
            f"{avg_m:.3f} ms",
            f"{avg_med:.3f} ms",
            f"Lọc Trung vị chậm hơn {ratio:.2f} lần",
        ])

    add_styled_table(
        doc,
        headers=[
            "Kích thước cửa sổ",
            "Số điểm ảnh lân cận (N = k^2)",
            "Thời gian Lọc Trung bình (ms)",
            "Thời gian Lọc Trung vị (ms)",
            "Tỷ lệ chênh lệch thời gian",
        ],
        rows=time_table_rows,
        caption="Bảng 4.8. So sánh thời gian xử lý trung bình (ms) trên ảnh 256x256 giữa hai bộ lọc",
        col_aligns=[WD_ALIGN_PARAGRAPH.CENTER] * 5,
    )

    time_fig = os.path.join(fig_dir, "chart_execution_time.png")
    if os.path.exists(time_fig):
        add_figure(
            doc,
            time_fig,
            caption="Hình 4.11. Biểu đồ so sánh thời gian thực thi trung bình (ms) theo kích thước cửa sổ lọc",
            width_inches=5.2,
        )

    # 4.3.5
    add_h3(doc, "4.3.5. Kết quả kiểm thử chức năng ứng dụng Web tương tác trên Localhost")
    add_para(
        doc,
        "Bảng 4.9 ghi nhận kết quả thực thi 6 ca kiểm thử chức năng trên ứng dụng Web chạy tại http://127.0.0.1:8000. Hình 4.12 minh họa kết quả kiểm thử khử nhiễu Gaussian (σ = 25) trên ảnh coins.png (khung Lọc Trung bình 3x3 được viền xanh nổi bật với PSNR = 26.07 dB, MSE = 160.62, Edge-F1 = 0.7503 so với 25.68 dB, MSE = 175.79, Edge-F1 = 0.7443 của Trung vị 3x3), và Hình 4.13 minh họa trạng thái sau khi nhấn 'Reset Ảnh Gốc'."
    )

    tc_headers = ["Mã TC", "Tên chức năng kiểm thử", "Thao tác thực hiện trên Web", "Kết quả thực tế ghi nhận", "Đánh giá"]
    tc_rows = [
        ["TC-01", "Chọn ảnh từ bộ 6 ảnh chuẩn", "Chọn lần lượt 6 ảnh chuẩn ở danh sách thả xuống.", "Tải ảnh tức thì, cập nhật đúng kích thước 256x256 px và ngưỡng Sobel P85 tương ứng.", "ĐẠT (PASS)"],
        ["TC-02", "Upload ảnh bất kỳ từ máy tính", "Chọn tệp ảnh màu/xám (.PNG, .JPG) và chọn chuẩn hóa 256x256 hoặc 512x512.", "Backend chuyển sang ảnh xám 8-bit, chuẩn hóa kích thước và hiển thị đủ 8 khung ảnh.", "ĐẠT (PASS)"],
        ["TC-03", "Thêm nhiễu Muối tiêu & Gaussian", "Kéo slider mật độ Muối tiêu (0.5%–30%) và độ lệch chuẩn Gaussian (σ = 2–80).", "Ảnh nhiễu cập nhật ngay theo seed cố định; MSE/PSNR khớp 100% với kết quả CSV.", "ĐẠT (PASS)"],
        ["TC-04", "Đổi cửa sổ lọc & Padding", "Chọn cửa sổ 3x3, 5x5, 7x7 và chế độ Reflect / Replicate / Zero.", "8 khung hình và bảng chỉ số cập nhật theo thời gian thực; bộ lọc tốt hơn sáng viền xanh.", "ĐẠT (PASS)"],
        ["TC-05", "Nút 'Reset Ảnh Gốc'", "Nhấn nút 'Reset Ảnh Gốc' khi ảnh đang nhiễm nhiễu nặng.", "Khôi phục nguyên trạng ảnh gốc sạch (MSE = 0, PSNR = ∞, Edge-F1 = 1.0000).", "ĐẠT (PASS)"],
        ["TC-06", "Tải ảnh kết quả về máy", "Nhấn liên kết tải ảnh dưới từng khung (ảnh lọc & bản đồ biên Sobel).", "Trình duyệt tải về tệp ảnh PNG đúng kích thước và định dạng 8-bit.", "ĐẠT (PASS)"],
    ]
    add_styled_table(
        doc,
        headers=tc_headers,
        rows=tc_rows,
        caption="Bảng 4.9. Bảng kịch bản kiểm thử các chức năng của ứng dụng Web trên Localhost",
        col_aligns=[WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER],
    )

    for fig_name, cap in [
        ("web_ui_gaussian.png", "Hình 4.12. Kiểm thử chức năng khử nhiễu Gaussian (σ = 25) trên ảnh 'coins' qua giao diện Web Localhost"),
        ("web_ui_reset_clean.png", "Hình 4.13. Kiểm thử chức năng nút 'Reset Ảnh Gốc' khôi phục trạng thái ảnh sạch ban đầu trên giao diện Web"),
    ]:
        p = os.path.join(fig_dir, fig_name)
        if os.path.exists(p):
            add_figure(doc, p, caption=cap, width_inches=6.1)

    # 4.4
    add_h2(doc, "4.4. Phân tích và Đánh giá kết quả")
    add_h3(doc, "4.4.1. Giải thích bản chất toán học của các hiện tượng thực nghiệm")
    add_bullet(
        doc,
        "Trên cả 3 mức mật độ nhiễu Muối tiêu (1%, 5%, 10%), Bộ lọc Trung vị 3x3 luôn vượt trội tuyệt đối: đạt PSNR trung bình 31.49 dB (mức 1%), 29.43 dB (mức 5%) và 28.04 dB (mức 10%), đồng thời giữ Edge-F1 cao từ 0.7771 đến 0.7971 (Bảng 4.4). Đặc biệt trên ảnh checkerboard (1% S&P), Trung vị 3x3 đạt PSNR 43.07 dB (MSE = 3.20) so với 24.67 dB (MSE = 221.99) của Trung bình 3x3. Lý do gốc rễ nằm ở điểm gãy thống kê (Breakdown Point): Lọc Trung vị có điểm gãy gần 50% (chịu được tối đa 4 hạt nhiễu trong cửa sổ 3x3) và cô lập hoàn toàn cực trị 0/255 ở hai đầu dãy sắp xếp, trong khi Lọc Trung bình có điểm gãy 0%, làm phát tán mỗi hạt nhiễu thành đốm mờ k × k khiến bản đồ Sobel phát sinh hàng loạt biên giả (trên cameraman 5% S&P, Edge-Precision của Trung bình 3x3 tụt xuống 0.3830 so với 0.9574 của Trung vị 3x3).",
        bold_prefix="1) Vì sao Bộ lọc Trung vị thắng tuyệt đối trên nhiễu Muối tiêu: "
    )
    add_bullet(
        doc,
        "Với nhiễu Gaussian vừa và nặng (σ = 25, 50), trên các ảnh có nhiều vùng chuyển tiếp mượt như cameraman và coins (Bảng 4.7), Bộ lọc Trung bình 3x3 đạt PSNR cao hơn Trung vị 3x3: tại σ = 25 trên cameraman đạt 25.99 dB vs 25.60 dB (Edge-F1: 0.6909 vs 0.5963); tại σ = 50 đạt 22.53 dB vs 21.26 dB (tính trung bình 6 ảnh là 21.64 dB vs 21.15 dB). Điều này hoàn toàn khớp với lý thuyết MLE: trung bình mẫu là bộ ước lượng không chệch có phương sai nhỏ nhất σ^2/k^2, trong khi phương sai của trung vị mẫu lớn hơn xấp xỉ π/2 ≈ 1.57 lần. Ngoài ra, ở mức rất nặng σ = 50 trên cameraman và coins, cửa sổ Trung bình 5x5 tiếp tục tăng PSNR lên 22.86 dB và 23.09 dB nhờ giảm phương sai nhiễu đi 25 lần. Ngược lại, ở mức nhiễu Gaussian rất nhẹ σ = 10 (ảnh nhiễu vốn đạt 28.57 dB), sai số nhòe biên của Lọc Trung bình trên các ảnh cạnh gắt (checkerboard, text) lớn hơn lượng nhiễu khử được, nên PSNR trung bình bị tụt xuống 26.55 dB so với 29.11 dB của Trung vị 3x3.",
        bold_prefix="2) Vì sao Bộ lọc Trung bình tối ưu giảm phương sai trên nhiễu Gaussian nặng nhưng bị thiệt khi nhiễu nhẹ: "
    )
    add_bullet(
        doc,
        "Tăng cửa sổ từ 3x3 lên 5x5 và 7x7 làm tăng diện tích lấy mẫu từ 9 lên 25 và 49 pixel, giúp san phẳng vùng nhiễu cực nặng nhưng đồng thời làm giảm mạnh độ phủ biên (Edge-Recall) và độ phân giải chi tiết, đẩy thời gian chạy của Lọc Trung vị tăng vọt từ "
        f"{time_stats[3][1]:.2f} ms (chậm hơn {time_stats[3][2]:.2f}x) lên {time_stats[5][1]:.2f} ms ({time_stats[5][2]:.2f}x) và {time_stats[7][1]:.2f} ms ({time_stats[7][2]:.2f}x) so với Lọc Trung bình ({time_stats[3][0]:.2f} ms, {time_stats[5][0]:.2f} ms, {time_stats[7][0]:.2f} ms) tại Bảng 4.8 — khớp hoàn toàn với phân tích độ phức tạp O(H·W·k^2·log k) so với O(H·W·k^2).",
        bold_prefix="3) Quy luật đánh đổi theo kích thước cửa sổ (3x3, 5x5, 7x7) và thời gian tính toán: "
    )

    add_h3(doc, "4.4.2. Đánh giá ưu, nhược điểm và Các trường hợp thất bại (Failure Cases)")
    add_para(
        doc,
        "Để đánh giá khách quan giới hạn của chương trình, Bảng 4.10 tổng hợp các kịch bản hoạt động tốt và các trường hợp thuật toán thất bại (Failure Cases) quan sát trực tiếp từ bảng dữ liệu detailed_results.csv kèm nguyên nhân gốc rễ."
    )

    fail_headers = ["Bộ lọc", "Trường hợp hoạt động tốt", "Trường hợp thất bại (Failure Cases) quan sát được", "Nguyên nhân gốc rễ"]
    fail_rows = [
        [
            "Bộ lọc Trung bình\n(Mean Filter)",
            "- Nhiễu cộng Gaussian vừa và nặng (σ = 25, 50) trên ảnh vùng nền rộng (cameraman, coins).\n- Yêu cầu tốc độ cao (~3–6.6 ms/ảnh 256x256).",
            "1) Thất bại trên nhiễu Muối tiêu: ở mức 1% S&P với cửa sổ 5x5 và 7x7, PSNR tụt xuống 23.06 dB và 21.42 dB (thấp hơn cả ảnh nhiễu 24.59 dB).\n2) Phá hủy nét chữ mảnh trên text.png: với cửa sổ 5x5/7x7 (S&P 5% và Gaussian σ=25), Edge-F1 rơi xuống 0.0012 và 0.0000 (mất 100% biên chữ).",
            "1) Điểm gãy 0%: giá trị cực trị 0/255 bị chia đều ra cả vùng lân cận tạo đốm mờ.\n2) Cơ chế làm trơn đồng đều không phân biệt biên làm san phẳng hoàn toàn các nét mảnh 1–2 pixel xuống dưới ngưỡng Sobel.",
        ],
        [
            "Bộ lọc Trung vị\n(Median Filter)",
            "- Nhiễu Muối tiêu ở cả 3 mức 1%, 5%, 10% với cửa sổ 3x3 (PSNR 28.04–31.49 dB, Edge-F1 ~0.78–0.80).\n- Bảo toàn biên bước nhảy và góc vuông sắc nét (checkerboard).",
            "1) Thất bại trên nét chữ mảnh ở cửa sổ lớn: trên text.png (S&P 5%, 7x7), Edge-Recall tụt từ 0.4605 (3x3) xuống còn 0.0366 (Edge-F1 = 0.0701).\n2) Kém hơn Mean trên Gaussian nặng (σ = 50) với 3x3 (21.15 dB vs 21.64 dB) và sót cụm nhiễu khi hạt muối/tiêu tụ thành đám ≥ 5 pixel trong 3x3.",
            "1) Với nét chữ mảnh dày 1–2 px (chiếm < 50% diện tích cửa sổ 7x7 = 49 px), toàn bộ nét chữ bị coi như điểm ngoại lai và bị nền trắng nuốt mất.\n2) Phương sai ước lượng trung vị cao hơn ~57% và cửa sổ 3x3 bị phá vỡ khi số điểm xung vượt quá 4/9.",
        ],
    ]
    add_styled_table(
        doc,
        headers=fail_headers,
        rows=fail_rows,
        caption="Bảng 4.10. Tổng hợp ưu điểm, nhược điểm và các trường hợp thất bại (Failure cases) của hai bộ lọc",
        col_aligns=[WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT],
    )


def write_conclusion_and_appendix(doc: docx.Document, tables_dir: str) -> None:
    """KẾT LUẬN VÀ HƯỚNG PHÁT TRIỂN + TÀI LIỆU THAM KHẢO & PHỤ LỤC."""
    add_h1(doc, "KẾT LUẬN VÀ HƯỚNG PHÁT TRIỂN", page_break_before=True)

    add_h2(doc, "1. Tóm tắt kết quả đạt được")
    add_para(
        doc,
        "Đối chiếu với các mục tiêu và phạm vi đề ra ở Chương 1, Nhóm 10 đã hoàn thành toàn diện Đề tài 7: \"Khử nhiễu ảnh bằng các bộ lọc trong miền không gian\" trên cả ba phương diện:"
    )
    add_bullet(
        doc,
        "Hệ thống hóa chặt chẽ cơ sở toán học của biểu diễn ảnh số, mô hình nhiễu Muối tiêu & Gaussian, phép tích chập 2D, 3 kỹ thuật mở rộng biên, hai bộ lọc không gian tuyến tính/phi tuyến, toán tử phát hiện biên Sobel và hệ thống 5 độ đo định lượng (MSE, PSNR, Edge-MSE, Edge-Precision/Recall/F1).",
        bold_prefix="Về mặt lý thuyết: "
    )
    add_bullet(
        doc,
        "Tự cài đặt thành công 100% thuật toán cốt lõi bằng NumPy theo kiến trúc module hóa (src/), không phụ thuộc hàm lọc đóng gói sẵn của OpenCV. Kết quả kiểm chứng trên 7 bài test đối chiếu với OpenCV đều đạt sai số tuyệt đối bằng 0.0. Đồng thời xây dựng thành công ứng dụng Web tương tác trên Localhost (web_app.py) vượt qua 6/6 ca kiểm thử chức năng.",
        bold_prefix="Về mặt kỹ thuật xây dựng chương trình: "
    )
    add_bullet(
        doc,
        "Thực thi trọn vẹn 216 cấu hình thực nghiệm trên 6 ảnh chuẩn với seed cố định (SEED = 42), chứng minh bằng số liệu định lượng: Bộ lọc Trung vị 3x3 đặc trị nhiễu Muối tiêu (PSNR 28.04–31.49 dB, Edge-F1 ~0.78–0.80); Bộ lọc Trung bình tối ưu giảm phương sai trên nhiễu Gaussian vừa và nặng với tốc độ nhanh hơn từ 2.85 đến 10.95 lần; đồng thời chỉ rõ các trường hợp thất bại khi dùng cửa sổ 5x5, 7x7 trên ảnh nét mảnh.",
        bold_prefix="Về mặt thực nghiệm và phân tích kết quả: "
    )

    add_h2(doc, "2. Đề xuất hướng cải tiến trong tương lai")
    add_bullet(
        doc,
        "Thay thế nhân lọc đồng đều 1/k^2 bằng Bộ lọc Gaussian (trọng số giảm theo khoảng cách), Bộ lọc Song phương (Bilateral Filter) hoặc Non-Local Means (NLM) kết hợp độ tương đồng mức xám để khử nhiễu Gaussian hiệu quả mà vẫn giữ sắc nét biên vật thể.",
        bold_prefix="Cải tiến khả năng giữ biên cho nhóm lọc tuyến tính: "
    )
    add_bullet(
        doc,
        "Cài đặt Bộ lọc Trung vị Thích nghi (Adaptive Median Filter - AMF) tự động co giãn cửa sổ tùy theo mật độ nhiễu cục bộ và chỉ thay thế những điểm ảnh thực sự bị nhiễu xung, khắc phục hiện tượng mất nét chữ mảnh và vỡ chi tiết khi mật độ nhiễu cao > 20%.",
        bold_prefix="Cải tiến Bộ lọc Trung vị cho nhiễu mật độ cao & nét mảnh: "
    )
    add_bullet(
        doc,
        "Ứng dụng kỹ thuật Ảnh tích phân (Integral Image) để đưa độ phức tạp Bộ lọc Trung bình về O(1) độc lập với k, và thuật toán trượt Histogram của Huang để tăng tốc Bộ lọc Trung vị phục vụ xử lý ảnh độ phân giải cao (4K) hoặc luồng video thời gian thực.",
        bold_prefix="Tối ưu hiệu năng tính toán thời gian thực: "
    )

    # TÀI LIỆU THAM KHẢO & PHỤ LỤC
    add_h1(doc, "TÀI LIỆU THAM KHẢO & PHỤ LỤC", page_break_before=True)
    add_h2(doc, "1. Tài liệu tham khảo")
    refs = [
        "[1] Rafael C. Gonzalez, Richard E. Woods, \"Digital Image Processing\", 4th Edition, Pearson Education, New York, 2018 (Chương 3: Intensity Transformations and Spatial Filtering; Chương 5: Image Restoration and Reconstruction; Chương 10: Image Segmentation).",
        "[2] Lương Mạnh Bá, Nguyễn Thanh Thủy, \"Nhập môn Xử lý ảnh số\", Nhà xuất bản Khoa học và Kỹ thuật, Hà Nội.",
        "[3] Đỗ Năng Toàn, Phạm Việt Bình, \"Giáo trình Xử lý ảnh\", Khoa Công nghệ Thông tin – Đại học Thái Nguyên / Viện Công nghệ Thông tin.",
        "[4] Bài giảng học phần Xử lý ảnh và thị giác máy tính (IT3243), Khoa Công nghệ Thông tin, Trường Đại học Công nghệ Đông Á, năm học 2025 – 2026.",
        "[5] NumPy Developers, \"NumPy Reference & User Guide\", https://numpy.org/doc/stable/.",
        "[6] OpenCV Team, \"OpenCV Image Filtering & Sobel Derivatives Documentation\", https://docs.opencv.org/4.x/d4/d86/group__imgproc__filter.html.",
        "[7] scikit-image Team, \"Module: skimage.data (Standard Test Images)\", https://scikit-image.org/docs/stable/api/skimage.data.html.",
    ]
    for ref in refs:
        add_para(doc, ref, first_line_indent=0.0)

    add_h2(doc, "2. Phụ lục mã nguồn (GitHub) và hướng dẫn tái lập thực nghiệm")
    add_bullet(
        doc,
        "https://github.com/NguyenManhQuyet20231053/BTL_XLA_Nhom10_DeTai7 (Chứa toàn bộ mã nguồn tự cài đặt trong src/, kịch bản run_experiments.py, ứng dụng Web web_app.py, dữ liệu 6 ảnh chuẩn và kết quả CSV/PNG).",
        bold_prefix="Kho lưu trữ mã nguồn GitHub của Nhóm 10: "
    )
    add_code_block(
        doc,
        [
            "# Bước 1: Cài đặt các thư viện cần thiết",
            "pip install -r requirements.txt",
            "",
            "# Bước 2: Chạy toàn bộ 216 cấu hình thực nghiệm (cố định seed = 42)",
            "python run_experiments.py",
            "",
            "# Bước 3: Khởi chạy giao diện Web tương tác trên Localhost (http://127.0.0.1:8000)",
            "python web_app.py",
        ],
        caption="Hướng dẫn dòng lệnh tái lập toàn bộ thực nghiệm và khởi chạy Web App",
    )

    add_h2(doc, "3. Phụ lục Bảng số liệu thực nghiệm trích xuất trên các ảnh còn lại")
    add_para(
        doc,
        "Bảng P.1 trích xuất số liệu thực nghiệm từ tệp results/tables/detailed_results.csv ở cửa sổ 3x3 và 5x5 cho 5 ảnh còn lại trong bộ dữ liệu (coins, astronaut, text, checkerboard, brick)."
    )

    det_rows = load_csv_rows(os.path.join(tables_dir, "detailed_results.csv"))
    app_headers = [
        "Tên ảnh",
        "Loại nhiễu",
        "Mức nhiễu",
        "Bộ lọc",
        "Cửa sổ",
        "MSE",
        "PSNR (dB)",
        "Edge-MSE",
        "Edge-F1",
        "Thời gian (ms)",
    ]
    app_rows = []
    for r in det_rows:
        if r["image"] != "cameraman" and r["kernel_size"] in ("3", "5") and r["noise_level"] in ("5%", "10%", "sigma=25", "sigma=50"):
            app_rows.append([
                r["image"],
                r["noise_type"],
                r["noise_level"],
                "Trung bình" if r["filter_type"] == "Mean" else "Trung vị",
                f"{r['kernel_size']}x{r['kernel_size']}",
                r["mse"],
                r["psnr"],
                r["edge_mse"],
                r["edge_f1"],
                r["time_ms"],
            ])
    add_styled_table(
        doc,
        headers=app_headers,
        rows=app_rows,
        caption="Bảng P.1. Trích xuất kết quả thực nghiệm chi tiết trên các ảnh coins, astronaut, text, checkerboard và brick",
        col_aligns=[WD_ALIGN_PARAGRAPH.CENTER] * 10,
        col_widths_cm=[2.5, 2.0, 1.8, 1.9, 1.1, 1.4, 1.3, 1.6, 1.2, 1.2],
        font_size=10.0,
    )


def main() -> None:
    root_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    doc_path = os.path.join(root_dir, "BTL_XLA.docx")
    upd_path = os.path.join(root_dir, "BTL_XLA_updated.docx")
    tables_dir = os.path.join(root_dir, "results", "tables")
    fig_dir = os.path.join(root_dir, "results", "figures")

    doc = docx.Document(doc_path)
    configure_styles(doc)
    reset_body_keeping_front_matter(doc)
    update_abbreviations_table(doc)
    enable_update_fields_on_open(doc)

    write_preface(doc)
    write_chapter_1(doc)
    write_chapter_2(doc)
    write_chapter_3(doc, fig_dir)
    write_chapter_4(doc, tables_dir, fig_dir)
    write_conclusion_and_appendix(doc, tables_dir)

    try:
        doc.save(doc_path)
        print(f"Đã cập nhật thành công toàn bộ báo cáo vào: {doc_path}")
    except PermissionError:
        doc.save(upd_path)
        print(f"Lưu ý: BTL_XLA.docx đang mở trong Word nên đã lưu bản cập nhật tại: {upd_path}")


if __name__ == "__main__":
    main()

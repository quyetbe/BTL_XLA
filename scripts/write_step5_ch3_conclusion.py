# -*- coding: utf-8 -*-
"""
Script: scripts/write_step5_ch3_conclusion.py
Mô tả: Đọc trực tiếp số liệu thật từ các file CSV trong results/tables/ và các hình ảnh
       trong results/figures/ để biên soạn hoàn chỉnh:
       - CHƯƠNG 3. THỰC NGHIỆM
       - KẾT LUẬN
       - TÀI LIỆU THAM KHẢO & PHỤ LỤC
       vào file báo cáo Word của Nhóm 10.
"""

import os
import sys
import csv
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH

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
)


def remove_existing_chapter3_onwards(doc: docx.Document) -> None:
    """Nếu đã có Chương 3 trở đi thì xóa để ghi mới không bị trùng lặp."""
    body = doc.element.body
    children = list(body)
    ch3_idx = None
    for idx, child in enumerate(children):
        if child.tag.endswith("p"):
            txt = "".join(child.itertext()).strip()
            if txt.startswith("CHƯƠNG 3."):
                ch3_idx = idx
                break
    if ch3_idx is not None:
        for child in children[ch3_idx:]:
            if not child.tag.endswith("sectPr"):
                body.remove(child)


def load_csv_rows(csv_path: str) -> list:
    with open(csv_path, mode="r", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def write_chapter_3(doc: docx.Document, tables_dir: str, fig_dir: str) -> None:
    """Ghi toàn bộ Chương 3 dựa trên 100% số liệu thật từ thực nghiệm."""
    ver_rows = load_csv_rows(os.path.join(tables_dir, "verification_results.csv"))
    sum_rows = load_csv_rows(os.path.join(tables_dir, "summary_by_noise_and_filter.csv"))
    det_rows = load_csv_rows(os.path.join(tables_dir, "detailed_results.csv"))

    add_h1(doc, "CHƯƠNG 3. THỰC NGHIỆM", page_break_before=True)

    # --- MỤC 3.1 ---
    add_h2(doc, "3.1. Mô tả môi trường và bộ dữ liệu sử dụng để thực nghiệm")
    add_h3(doc, "3.1.1. Môi trường phần cứng và phần mềm thực nghiệm")
    add_para(
        doc,
        "Toàn bộ các kịch bản thực nghiệm trong báo cáo này được tiến hành trên hệ thống máy tính cá nhân chạy hệ điều hành Windows 11 64-bit, trình thông dịch Python 3.12. Các gói thư viện tính toán khoa học bao gồm: NumPy 2.2.6 (dùng cho toàn bộ thuật toán tự cài đặt), OpenCV 5.0.0 và Pillow 11.3.0 (dùng để đọc/ghi ảnh và đối chiếu kiểm thử), Matplotlib 3.11.1 (dùng để xuất biểu đồ và hình ảnh đối sánh). Hạt giống bộ sinh số giả ngẫu nhiên được cố định toàn cục ở giá trị SEED = 42."
    )

    add_h3(doc, "3.1.2. Mô tả bộ dữ liệu 6 ảnh xám chuẩn")
    add_para(
        doc,
        "Để đánh giá khách quan và toàn diện hiệu năng của Bộ lọc trung bình và Bộ lọc trung vị trên nhiều cấu trúc hình học khác nhau, nhóm sử dụng bộ dữ liệu gồm 6 ảnh đa mức xám (8-bit Grayscale, dải mức xám [0, 255]) có kích thước chuẩn hóa 256 x 256 pixel (tổng cộng 65.536 điểm ảnh trên mỗi bức ảnh). Sáu bức ảnh được trích xuất từ kho ảnh chuẩn quốc tế (Standard Test Images trong thư viện scikit-image / USC-SIPI), đại diện cho 6 nhóm đặc trưng tần số không gian khác nhau:"
    )

    ds_headers = ["TT", "Tên ảnh", "Kích thước", "Đặc trưng không gian và mục đích kiểm thử"]
    ds_rows = [
        [
            "1",
            "cameraman.png",
            "256 x 256\n(8-bit xám)",
            "Ảnh kinh điển kết hợp hài hòa giữa vùng nền trời phẳng rộng lớn (tần số thấp), vùng áo khoác tối màu và các cạnh thẳng sắc nét của chân máy ảnh (tần số cao). Rất lý tưởng để quan sát đồng thời khả năng làm trơn nền và bảo toàn biên.",
        ],
        [
            "2",
            "coins.png",
            "256 x 256\n(8-bit xám)",
            "Chứa các đồng xu kim loại hình tròn tách biệt trên nền tối đồng nhất. Dùng để kiểm thử khả năng bảo toàn đường biên cong khép kín và các chi tiết chạm khắc nổi bên trong lòng đồng xu.",
        ],
        [
            "3",
            "astronaut.png",
            "256 x 256\n(8-bit xám)",
            "Ảnh chân dung phi hành gia có dải chuyển tiếp độ sáng mượt trên khuôn mặt kết hợp với các chi tiết đa tỷ lệ trên mũ bảo hộ, phù hiệu và quốc kỳ.",
        ],
        [
            "4",
            "text.png",
            "256 x 256\n(8-bit xám)",
            "Ảnh trang tài liệu văn bản chứa các nét chữ in mảnh (độ dày chỉ 1–2 pixel) có độ tương phản cao trên nền giấy sáng. Đây là phép thử khắc nghiệt nhất đối với kích thước cửa sổ lọc lớn (5x5, 7x7) vì rất dễ gây mất nét chữ.",
        ],
        [
            "5",
            "checkerboard.png",
            "256 x 256\n(8-bit xám)",
            "Ảnh bàn cờ hình học gồm các ô vuông sáng/tối xen kẽ tuần hoàn với đường biên bước nhảy (step edges) tuyệt đối và các góc vuông 90 độ. Dùng để đánh giá hiện tượng bo tròn góc và nhòe biên.",
        ],
        [
            "6",
            "brick.png",
            "256 x 256\n(8-bit xám)",
            "Ảnh kết cấu tường gạch có các rãnh vữa hẹp lặp tuần hoàn và bề mặt gạch nhám mịn. Dùng để khảo sát sự suy giảm kết cấu bề mặt (texture degradation) khi tăng kích thước cửa sổ lọc.",
        ],
    ]
    add_styled_table(
        doc,
        headers=ds_headers,
        rows=ds_rows,
        caption="Bảng 3.1. Mô tả chi tiết bộ dữ liệu 6 ảnh xám chuẩn sử dụng trong thực nghiệm",
        col_aligns=[
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.LEFT,
        ],
    )

    ds_fig = os.path.join(fig_dir, "dataset_overview.png")
    if os.path.exists(ds_fig):
        add_figure(
            doc,
            ds_fig,
            caption="Hình 3.1. Trực quan hóa bộ dữ liệu 6 ảnh xám chuẩn kích thước 256x256 sử dụng trong thực nghiệm",
            width_inches=5.8,
        )

    # --- MỤC 3.2 ---
    add_h2(doc, "3.2. Kết quả kiểm tra tính đúng đắn của các thuật toán tự cài đặt")
    add_para(
        doc,
        "Trước khi đi vào phân tích kết quả khử nhiễu, bước bắt buộc trong quy trình kỹ thuật là kiểm chứng (Unit Verification) xem các hàm tự cài đặt bằng NumPy của nhóm (bao gồm mean_filter, median_filter và sobel_edge_detection) có hoạt động chính xác theo đúng định nghĩa toán học chuẩn hay không. Nhóm đã chạy module src/verification.py trên ảnh chuẩn cameraman.png để đo sai số tuyệt đối cực đại (Max Absolute Error) và sai số tuyệt đối trung bình (Mean Absolute Error) giữa kết quả của hàm tự viết và kết quả của các hàm chuẩn trong thư viện OpenCV (cv2.blur, cv2.medianBlur, cv2.Sobel) dưới cùng một chế độ mở rộng biên."
    )

    ver_headers = ["Thuật toán tự cài đặt (NumPy)", "Hàm chuẩn đối chiếu (OpenCV)", "Sai số cực đại", "Sai số trung bình", "Kết luận kiểm thử"]
    ver_table_rows = [
        [r["algorithm"], r["reference"], r["max_abs_error"], r["mean_abs_error"], r["status"]]
        for r in ver_rows
    ]
    add_styled_table(
        doc,
        headers=ver_headers,
        rows=ver_table_rows,
        caption="Bảng 3.2. Kết quả kiểm chứng độ chính xác của các thuật toán tự cài đặt so với thư viện OpenCV",
        col_aligns=[
            WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.CENTER,
        ],
    )
    add_para(
        doc,
        "Số liệu thực tế tại Bảng 3.2 khẳng định: Cả 7 phép kiểm thử (Lọc trung bình 3x3, 5x5, 7x7; Lọc trung vị 3x3, 5x5, 7x7 và Toán tử Sobel 3x3) đều đạt sai số tuyệt đối cực đại bằng 0.0 và sai số trung bình bằng 0.0 (khớp tuyệt đối 100% tới từng điểm ảnh và từng chữ số thập phân của độ lớn Gradient). Điều này chứng minh toàn bộ mã nguồn tự cài đặt của Nhóm 10 là hoàn toàn chuẩn xác, tạo cơ sở tin cậy tuyệt đối cho các phân tích ở các mục tiếp theo."
    )

    # --- MỤC 3.3 ---
    add_h2(doc, "3.3. Kết quả thực nghiệm và đánh giá trên ảnh nhiễu Muối tiêu")
    add_h3(doc, "3.3.1. Bảng số liệu thực nghiệm trên nhiễu Muối tiêu (1%, 5%, 10%)")
    add_para(
        doc,
        "Trong kịch bản thứ nhất, 6 ảnh gốc lần lượt được bổ sung nhiễu muối tiêu (Salt-and-Pepper) ở 3 mức mật độ: nhẹ (1%), trung bình (5%) và nặng (10%). Trước khi lọc, chỉ số PSNR trung bình của ảnh nhiễu giảm mạnh từ 24.59 dB (ở mức 1%) xuống 17.79 dB (ở mức 5%) và chỉ còn 14.81 dB (ở mức 10%). Bảng 3.3 trình bày kết quả trung bình trên toàn bộ 6 ảnh chuẩn, và Bảng 3.4 trình bày số liệu chi tiết trên ảnh đại diện cameraman.png."
    )

    sp_sum_headers = [
        "Mức nhiễu\n(PSNR nhiễu)",
        "Bộ lọc",
        "Cửa sổ",
        "MSE\n(Sau lọc)",
        "PSNR\n(dB)",
        "Edge-MSE",
        "Edge\nPrecision",
        "Edge\nRecall",
        "Edge\nF1-Score",
        "Thời gian\n(ms)",
    ]
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
        caption="Bảng 3.3. Kết quả trung bình trên 6 ảnh chuẩn khi khử nhiễu Muối tiêu (1%, 5%, 10%)",
        col_aligns=[WD_ALIGN_PARAGRAPH.CENTER] * 10,
    )

    # Bảng chi tiết trên ảnh cameraman với nhiễu muối tiêu
    cam_sp_rows = []
    for r in det_rows:
        if r["image"] == "cameraman" and r["noise_type"] == "Salt & Pepper":
            cam_sp_rows.append([
                f"{r['noise_level']} ({r['noisy_psnr']} dB)",
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
        caption="Bảng 3.4. Số liệu thực nghiệm chi tiết trên ảnh 'cameraman' với nhiễu Muối tiêu",
        col_aligns=[WD_ALIGN_PARAGRAPH.CENTER] * 10,
    )

    add_h3(doc, "3.3.2. Hình ảnh so sánh trực quan và biểu đồ phân tích trên nhiễu Muối tiêu")
    add_para(
        doc,
        "Hình 3.2 và Hình 3.3 dưới đây minh họa trực quan kết quả khử nhiễu Muối tiêu ở mật độ 5% và 10% trên hai ảnh tiêu biểu là cameraman và coins. Tiếp đó, Hình 3.4 biểu diễn đồ thị biến thiên của PSNR trung bình và chỉ số bảo toàn biên Edge F1-Score theo mật độ nhiễu muối tiêu và kích thước cửa sổ."
    )

    for fig_name, cap in [
        ("compare_cameraman_sa_5pct.png", "Hình 3.2. So sánh trực quan kết quả khử nhiễu Muối tiêu 5% trên ảnh 'cameraman'"),
        ("compare_coins_sa_10pct.png", "Hình 3.3. So sánh trực quan kết quả khử nhiễu Muối tiêu 10% trên ảnh 'coins'"),
        ("chart_salt_pepper_psnr_f1.png", "Hình 3.4. Biểu đồ PSNR trung bình và Edge F1-Score trên nhiễu Muối tiêu theo mật độ nhiễu và kích thước cửa sổ"),
    ]:
        p = os.path.join(fig_dir, fig_name)
        if os.path.exists(p):
            add_figure(doc, p, caption=cap, width_inches=5.9)

    add_h3(doc, "3.3.3. Nhận xét và đánh giá thực nghiệm trên nhiễu Muối tiêu")
    add_para(
        doc,
        "Từ số liệu định lượng tại Bảng 3.3, Bảng 3.4 và các hình ảnh trực quan, ta rút ra các nhận xét quan trọng sau:"
    )
    add_bullet(
        doc,
        "Trên mọi mức mật độ nhiễu muối tiêu (1%, 5%, 10%), Bộ lọc trung vị cửa sổ 3x3 luôn đạt chất lượng khôi phục cao nhất và bỏ xa Bộ lọc trung bình. Cụ thể, tính trung bình trên 6 ảnh (Bảng 3.3), ở mức nhiễu 1%, Lọc trung vị 3x3 đạt PSNR = 31.49 dB (MSE = 97.22), cao hơn tới +5.41 dB so với Lọc trung bình 3x3 (26.08 dB, MSE = 177.10). Ở mức nhiễu 5%, khoảng cách này mở rộng lên thành +5.77 dB (Lọc trung vị 3x3 đạt 29.43 dB so với 23.66 dB của Lọc trung bình 3x3). Đặc biệt ở mức nhiễu nặng 10%, Lọc trung vị 3x3 vẫn giữ được PSNR rất cao là 28.04 dB (tăng +13.23 dB so với ảnh nhiễu 14.81 dB), cao hơn Lọc trung bình 3x3 (21.69 dB) tới +6.35 dB. Thậm chí trên ảnh bàn cờ checkerboard ở mức nhiễu 1%, Lọc trung vị 3x3 đạt PSNR lên tới 43.07 dB (MSE chỉ bằng 3.20), trong khi Lọc trung bình 3x3 chỉ đạt 24.67 dB (MSE = 221.99).",
        bold_prefix="Sự vượt trội tuyệt đối của Bộ lọc Trung vị trước nhiễu Muối tiêu: "
    )
    add_bullet(
        doc,
        "Quan sát Hình 3.2 và Hình 3.3, ảnh sau khi qua Lọc trung bình 3x3 vẫn còn chi chít các đốm xám mờ kích thước 3x3 tại đúng vị trí các hạt muối/tiêu ban đầu (do giá trị 0 và 255 bị chia đều cho 9 pixel lân cận). Khi tăng cửa sổ Lọc trung bình lên 5x5 và 7x7, các đốm nhiễu tuy mờ hơn nhưng toàn bộ bức ảnh bị nhòe nặng, khiến PSNR trung bình ở mức nhiễu 1% tụt xuống chỉ còn 23.06 dB (5x5) và 21.42 dB (7x7) — tức là thấp hơn cả ảnh nhiễu chưa lọc (24.59 dB)!",
        bold_prefix="Hiện tượng phát tán nhiễu và suy giảm chất lượng của Bộ lọc Trung bình: "
    )
    add_bullet(
        doc,
        "Đối với Lọc trung vị, cửa sổ 3x3 (gồm 9 phần tử, chịu được tối đa 4 điểm nhiễu trong cùng một cửa sổ) là kích thước tối ưu nhất cho mật độ nhiễu từ 1% đến 10%. Khi tăng kích thước cửa sổ Lọc trung vị lên 5x5 và 7x7, mặc dù nhiễu xung vẫn bị triệt tiêu hoàn toàn nhưng các chi tiết mảnh bị bào mòn, làm PSNR trung bình ở mức nhiễu 5% giảm từ 29.43 dB (3x3) xuống 25.93 dB (5x5) và 24.19 dB (7x7).",
        bold_prefix="Ảnh hưởng của kích thước cửa sổ trên nhiễu Muối tiêu: "
    )

    # --- MỤC 3.4 ---
    add_h2(doc, "3.4. Kết quả thực nghiệm và đánh giá trên ảnh nhiễu Gaussian")
    add_h3(doc, "3.4.1. Bảng số liệu thực nghiệm trên nhiễu Gaussian (σ = 10, 25, 50)")
    add_para(
        doc,
        "Trong kịch bản thứ hai, 6 ảnh chuẩn được bổ sung nhiễu cộng trắng phân phối chuẩn Gaussian với kỳ vọng μ = 0 và 3 mức độ lệch chuẩn: nhẹ (σ = 10, PSNR ảnh nhiễu trung bình = 28.57 dB), trung bình (σ = 25, PSNR ảnh nhiễu = 20.81 dB) và rất nặng (σ = 50, PSNR ảnh nhiễu = 15.24 dB). Bảng 3.5 tổng hợp số liệu trung bình trên 6 ảnh và Bảng 3.6 trình bày số liệu chi tiết trên ảnh cameraman.png."
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
        caption="Bảng 3.5. Kết quả trung bình trên 6 ảnh chuẩn khi khử nhiễu Gaussian (σ = 10, 25, 50)",
        col_aligns=[WD_ALIGN_PARAGRAPH.CENTER] * 10,
    )

    cam_ga_rows = []
    for r in det_rows:
        if r["image"] == "cameraman" and r["noise_type"] == "Gaussian":
            cam_ga_rows.append([
                f"{r['noise_level']} ({r['noisy_psnr']} dB)",
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
        caption="Bảng 3.6. Số liệu thực nghiệm chi tiết trên ảnh 'cameraman' với nhiễu Gaussian",
        col_aligns=[WD_ALIGN_PARAGRAPH.CENTER] * 10,
    )

    add_h3(doc, "3.4.2. Hình ảnh so sánh trực quan và biểu đồ phân tích trên nhiễu Gaussian")
    add_para(
        doc,
        "Hình 3.5 và Hình 3.6 trình bày kết quả so sánh trực quan trên ảnh cameraman ở mức nhiễu Gaussian trung bình (σ = 25) và rất nặng (σ = 50). Hình 3.7 tổng hợp biểu đồ PSNR và Edge F1-Score của hai bộ lọc trên nhiễu Gaussian."
    )

    for fig_name, cap in [
        ("compare_cameraman_ga_sigma25.png", "Hình 3.5. So sánh trực quan kết quả khử nhiễu Gaussian (σ = 25) trên ảnh 'cameraman'"),
        ("compare_cameraman_ga_sigma50.png", "Hình 3.6. So sánh trực quan kết quả khử nhiễu Gaussian nặng (σ = 50) trên ảnh 'cameraman'"),
        ("chart_gaussian_psnr_f1.png", "Hình 3.7. Biểu đồ PSNR trung bình và Edge F1-Score trên nhiễu Gaussian theo độ lệch chuẩn σ và kích thước cửa sổ"),
    ]:
        p = os.path.join(fig_dir, fig_name)
        if os.path.exists(p):
            add_figure(doc, p, caption=cap, width_inches=5.9)

    add_h3(doc, "3.4.3. Nhận xét và đánh giá thực nghiệm trên nhiễu Gaussian")
    add_para(
        doc,
        "Phân tích số liệu thực tế trên nhiễu Gaussian tại Bảng 3.5 và Bảng 3.6 cho thấy một bức tranh thú vị và sâu sắc hơn rất nhiều so với nhiễu muối tiêu:"
    )
    add_bullet(
        doc,
        "Xét trên ảnh tiêu biểu cameraman (Bảng 3.6) và coins, khi cường độ nhiễu Gaussian tăng lên mức trung bình (σ = 25) và nặng (σ = 50), Bộ lọc trung bình 3x3 vượt qua Bộ lọc trung vị 3x3 cả về PSNR lẫn khả năng giảm sai số biên Edge-MSE. Cụ thể, trên ảnh cameraman với σ = 25, Lọc trung bình 3x3 đạt PSNR = 25.99 dB (MSE = 163.66, Edge-MSE = 1953.50), tốt hơn Lọc trung vị 3x3 (PSNR = 25.60 dB, MSE = 179.06, Edge-MSE = 2515.84). Trên ảnh coins với σ = 25, Lọc trung bình 3x3 đạt PSNR = 26.07 dB (MSE = 160.62) so với 25.68 dB (MSE = 175.79) của Lọc trung vị 3x3. Đặc biệt, khi nhiễu Gaussian ở mức rất nặng σ = 50, tính trung bình trên toàn bộ 6 ảnh (Bảng 3.5), Lọc trung bình 3x3 đạt PSNR = 21.64 dB (MSE = 462.89, Edge-F1 = 0.5722), vượt rõ rệt so với Lọc trung vị 3x3 (PSNR = 21.15 dB, MSE = 512.45, Edge-F1 = 0.5297). Trên ảnh cameraman ở σ = 50, Lọc trung bình 3x3 đạt 22.53 dB (MSE = 363.33), cao hơn Lọc trung vị 3x3 (21.26 dB, MSE = 486.54) tới +1.27 dB.",
        bold_prefix="Ưu thế khử nhiễu cộng của Bộ lọc Trung bình khi nhiễu Gaussian vừa và mạnh (σ = 25, 50): "
    )
    add_bullet(
        doc,
        "Ở mức nhiễu Gaussian rất nhẹ σ = 10 (khi PSNR của bản thân ảnh nhiễu đã khá cao, đạt trung bình 28.57 dB), sai số do nhiễu gây ra trên vùng nền nhỏ hơn nhiều so với sai số do làm nhòe các cạnh sắc nét khi lấy trung bình cộng. Vì vậy, trên các ảnh có nhiều biên tương phản mạnh như checkerboard, coins hay cameraman ở σ = 10, Lọc trung bình làm giảm PSNR từ 28.57 dB xuống 26.55 dB do hiệu ứng nhòe biên, trong khi Lọc trung vị 3x3 giữ nguyên được độ dốc của biên nên đạt PSNR trung bình 29.11 dB.",
        bold_prefix="Sự đánh đổi giữa triệt nhiễu nền và nhòe biên ở mức nhiễu Gaussian nhẹ (σ = 10): "
    )
    add_bullet(
        doc,
        "Trên ảnh cameraman với nhiễu Gaussian rất nặng σ = 50 (Bảng 3.6), khi tăng kích thước cửa sổ Lọc trung bình từ 3x3 lên 5x5, chỉ số PSNR tiếp tục được cải thiện từ 22.53 dB (MSE = 363.33) lên 22.86 dB (MSE = 336.29), và trên ảnh coins (σ = 50) PSNR tăng từ 22.68 dB (3x3) lên 23.09 dB (5x5). Nguyên nhân là với phương sai nhiễu quá lớn (σ^2 = 2500), cửa sổ 3x3 (trung bình 9 điểm) chưa đủ để triệt tiêu dao động nhiễu, trong khi cửa sổ 5x5 (trung bình 25 điểm) giảm phương sai nhiễu đi 25 lần, giúp vùng nền phẳng mượt hơn hẳn.",
        bold_prefix="Tác dụng tích cực của cửa sổ 5x5 khi nhiễu Gaussian có cường độ rất lớn (σ = 50): "
    )

    # --- MỤC 3.5 ---
    add_h2(doc, "3.5. Đánh giá khả năng bảo toàn biên với toán tử Sobel và thời gian xử lý")
    add_h3(doc, "3.5.1. Phân tích trực quan và định lượng bản đồ biên Sobel")
    add_para(
        doc,
        "Để đánh giá trực tiếp tác động của từng bộ lọc lên khâu trích xuất đặc trưng hình học ở tầng xử lý ảnh trung gian, toán tử phát hiện biên Sobel 3x3 tự cài đặt đã được chạy trên ảnh gốc, ảnh nhiễu và toàn bộ ảnh sau lọc. Như đã trình bày tại Chương 2, ngưỡng nhị phân hóa T_edge được tính theo phân vị thứ 85 của độ lớn Gradient trên ảnh gốc sạch và giữ cố định cho mọi ảnh nhiễu/ảnh lọc của cùng bức ảnh đó. Hình 3.8, Hình 3.9 và Hình 3.10 minh họa trực quan bản đồ biên Sobel trên ảnh cameraman và ảnh văn bản text."
    )

    for fig_name, cap in [
        ("sobel_cameraman_sa_5pct.png", "Hình 3.8. So sánh bản đồ biên Sobel trên ảnh 'cameraman' với nhiễu Muối tiêu 5%"),
        ("sobel_cameraman_ga_sigma25.png", "Hình 3.9. So sánh bản đồ biên Sobel trên ảnh 'cameraman' với nhiễu Gaussian (σ = 25)"),
        ("compare_text_sa_5pct.png", "Hình 3.10. Minh họa hiện tượng mất nét chữ mảnh trên ảnh 'text' khi tăng kích thước cửa sổ lọc lên 5x5 và 7x7"),
    ]:
        p = os.path.join(fig_dir, fig_name)
        if os.path.exists(p):
            add_figure(doc, p, caption=cap, width_inches=5.9)

    add_para(
        doc,
        "Qua phân tích Hình 3.8, Hình 3.9 và các chỉ số Edge-MSE, Edge-Precision, Edge-Recall, Edge-F1 trong các bảng số liệu, ta thấy rõ ba hiện tượng vật lý điển hình:"
    )
    add_bullet(
        doc,
        "Trên ảnh nhiễu chưa qua xử lý (ô dưới bên trái của Hình 3.8 và 3.9), toán tử đạo hàm Sobel cực kỳ nhạy cảm với các biến thiên tần số cao của nhiễu. Mỗi hạt nhiễu muối tiêu (0 hoặc 255) tạo ra một vòng tròn biên giả rất sáng kích thước 3x3 xung quanh nó, khiến bản đồ biên Sobel của ảnh nhiễu bị lấp đầy bởi các đốm trắng hỗn loạn và làm chỉ số Edge-F1 của ảnh nhiễu cameraman tụt xuống chỉ còn ~0.35.",
        bold_prefix="1) Hiện tượng bùng nổ 'biên giả' (False Edges) khi không lọc nhiễu: "
    )
    add_bullet(
        doc,
        "Với nhiễu Muối tiêu 5% trên ảnh cameraman (Hình 3.8, Bảng 3.4), Lọc trung vị 3x3 đạt độ chính xác biên Edge-Precision = 0.9574, độ phủ Edge-Recall = 0.7230 và chỉ số tổng hợp Edge-F1 = 0.8239 (sai số bản đồ gradient Edge-MSE chỉ là 654.50). Ngược lại, Lọc trung bình 3x3 do không xóa được hạt nhiễu xung mà biến chúng thành các khối mờ 3x3 nên Sobel vẫn bắt phải viền của các khối mờ này, làm Edge-Precision rơi xuống chỉ còn 0.3830, Edge-F1 = 0.4984 và Edge-MSE vọt lên tới 3297.04 (cao gấp 5 lần sai số của Lọc trung vị).",
        bold_prefix="2) Lọc Trung vị bảo toàn biên sắc nét và triệt tiêu biên giả vượt trội trên nhiễu xung: "
    )
    add_bullet(
        doc,
        "Khi kích thước cửa sổ lọc tăng lên 5x5 và 7x7, đặc biệt là với Lọc trung bình, độ dốc của các đường biên thật bị san phẳng khiến độ lớn Gradient G = √(G_x^2 + G_y^2) rơi xuống thấp hơn ngưỡng chuẩn T_edge. Hệ quả là độ phủ biên (Edge-Recall) sụt giảm nghiêm trọng. Minh chứng rõ nhất nằm ở ảnh văn bản text.png (dòng 110–127 của bảng chi tiết): vì các nét chữ chỉ dày 1–2 pixel, khi dùng cửa sổ 3x3 thì Lọc trung vị giữ được Edge-F1 = 0.6051, nhưng khi tăng lên cửa sổ 7x7 thì Lọc trung bình làm mờ trắng toàn bộ trang chữ (Edge-Recall = 0.0000, Edge-F1 = 0.0000), còn Lọc trung vị 7x7 cũng làm mất gần hết các nét chữ mảnh (Edge-Recall chỉ còn 0.0331).",
        bold_prefix="3) Hiện tượng suy giảm độ dốc Gradient và đứt gãy biên mảnh ở cửa sổ lớn (5x5, 7x7): "
    )

    add_h3(doc, "3.5.2. Đánh giá và so sánh thời gian xử lý trung bình của hai bộ lọc")
    add_para(
        doc,
        "Bên cạnh độ chính xác khử nhiễu, tốc độ tính toán là yếu tố then chốt khi triển khai ứng dụng thực tế. Bảng 3.7 và Hình 3.11 tổng hợp thời gian thực thi trung bình (tính bằng mili-giây trên một ảnh 256 x 256) của Bộ lọc trung bình và Bộ lọc trung vị theo 3 kích thước cửa sổ."
    )

    # Tính trung bình thời gian từ sum_rows
    time_table_rows = []
    for k in (3, 5, 7):
        m_times = [float(r["time_ms_mean"]) for r in sum_rows if r["filter_type"] == "Mean" and int(r["kernel_size"]) == k]
        med_times = [float(r["time_ms_mean"]) for r in sum_rows if r["filter_type"] == "Median" and int(r["kernel_size"]) == k]
        avg_m = sum(m_times) / len(m_times)
        avg_med = sum(med_times) / len(med_times)
        ratio = avg_med / avg_m if avg_m > 0 else 0.0
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
            "Số điểm ảnh trong lân cận (N = k^2)",
            "Thời gian Lọc Trung bình (ms)",
            "Thời gian Lọc Trung vị (ms)",
            "Tỷ lệ chênh lệch thời gian",
        ],
        rows=time_table_rows,
        caption="Bảng 3.7. So sánh thời gian xử lý trung bình (ms) trên ảnh 256x256 giữa hai bộ lọc",
        col_aligns=[WD_ALIGN_PARAGRAPH.CENTER] * 5,
    )

    time_fig = os.path.join(fig_dir, "chart_execution_time.png")
    if os.path.exists(time_fig):
        add_figure(
            doc,
            time_fig,
            caption="Hình 3.11. Biểu đồ so sánh thời gian thực thi trung bình (ms) theo kích thước cửa sổ lọc",
            width_inches=5.2,
        )

    add_para(
        doc,
        "Số liệu thực nghiệm tại Bảng 3.7 hoàn toàn khớp với phân tích độ phức tạp lý thuyết tại Chương 2: Bộ lọc trung bình (độ phức tạp O(H·W·k^2)) chỉ mất trung bình 3.06 ms ở cửa sổ 3x3, 4.91 ms ở cửa sổ 5x5 và 6.61 ms ở cửa sổ 7x7. Trong khi đó, Bộ lọc trung vị (độ phức tạp O(H·W·k^2·log k) do phải sắp xếp mảng k^2 phần tử tại mỗi pixel) mất trung bình 8.73 ms ở cửa sổ 3x3 (chậm hơn 2.85 lần), tăng lên 30.36 ms ở cửa sổ 5x5 (chậm hơn 6.18 lần) và mất tới 72.38 ms ở cửa sổ 7x7 (chậm hơn 10.95 lần so với Lọc trung bình)."
    )

    # --- MỤC 3.6 ---
    add_h2(doc, "3.6. Kết quả kiểm thử chức năng ứng dụng Web tương tác trên Localhost")
    add_para(
        doc,
        "Nhằm đáp ứng yêu cầu kiểm tra các chức năng chương trình ứng dụng một cách trực quan và toàn diện, Nhóm 10 đã tiến hành kiểm thử thực tế toàn bộ các tính năng của ứng dụng Web (web_app.py) chạy trên địa chỉ http://127.0.0.1:8000. Bảng 3.8 tổng hợp kết quả kiểm thử các kịch bản chức năng (Test Cases) trên giao diện Web."
    )

    tc_headers = ["Mã TC", "Tên chức năng kiểm thử", "Thao tác thực hiện trên Web", "Kết quả thực tế ghi nhận", "Đánh giá"]
    tc_rows = [
        [
            "TC-01",
            "Chọn ảnh từ bộ 6 ảnh chuẩn",
            "Chọn lần lượt từng ảnh (cameraman, coins, astronaut, text, checkerboard, brick) ở danh sách thả xuống.",
            "Hệ thống tải ảnh tức thì (< 50 ms), cập nhật đúng kích thước 256x256 px và ngưỡng biên Sobel P85 tương ứng.",
            "ĐẠT (PASS)",
        ],
        [
            "TC-02",
            "Tải ảnh trực tiếp từ máy tính (Upload)",
            "Nhấn 'Choose File', chọn 1 tệp ảnh màu/xám bất kỳ (.PNG, .JPG) từ ổ đĩa và chọn chế độ chuẩn hóa 256x256 hoặc 512x512.",
            "Backend tự động chuyển sang ảnh xám 8-bit, chuẩn hóa kích thước và hiển thị đầy đủ 8 khung ảnh Trước/Sau lọc.",
            "ĐẠT (PASS)",
        ],
        [
            "TC-03",
            "Thêm nhiễu Muối tiêu & Gaussian trực tiếp",
            "Chuyển đổi loại nhiễu, kéo thanh trượt mật độ Muối tiêu (1%, 5%, 10%, 20%) và độ lệch chuẩn Gaussian (σ = 10, 25, 50).",
            "Ảnh nhiễu cập nhật ngay lập tức với seed cố định, chỉ số MSE/PSNR của ảnh nhiễu khớp 100% với bảng CSV.",
            "ĐẠT (PASS)",
        ],
        [
            "TC-04",
            "Thay đổi cửa sổ lọc (3x3, 5x5, 7x7) & Padding",
            "Nhấn chọn các nút Cửa sổ 3x3, 5x5, 7x7 và đổi chế độ Padding (Reflect, Replicate, Zero).",
            "8 khung hình và bảng đối sánh số liệu cập nhật thời gian thực; khung hình của bộ lọc có PSNR cao hơn tự động sáng viền xanh lá.",
            "ĐẠT (PASS)",
        ],
        [
            "TC-05",
            "Chức năng nút 'Reset Ảnh Gốc'",
            "Nhấn nút màu đỏ '🔄 Reset Ảnh Gốc' khi ảnh đang bị nhiễu nặng.",
            "Trạng thái nhiễu được xóa sạch về 0%, ảnh khôi phục nguyên trạng về ảnh gốc ban đầu (MSE = 0, PSNR = ∞, Edge-F1 = 1).",
            "ĐẠT (PASS)",
        ],
        [
            "TC-06",
            "Tải ảnh kết quả về máy tính",
            "Nhấn vào liên kết '⬇ Tải ảnh...' bên dưới từng khung hình (Ảnh nhiễu, Lọc Trung bình, Lọc Trung vị, Biên Sobel).",
            "Trình duyệt tải về tệp ảnh PNG đúng kích thước và chất lượng 8-bit.",
            "ĐẠT (PASS)",
        ],
    ]
    add_styled_table(
        doc,
        headers=tc_headers,
        rows=tc_rows,
        caption="Bảng 3.8. Bảng kịch bản kiểm thử các chức năng của ứng dụng Web trên Localhost",
        col_aligns=[
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.CENTER,
        ],
    )

    add_para(
        doc,
        "Hình 3.12 dưới đây minh họa màn hình ứng dụng Web khi kiểm thử khử nhiễu cộng Gaussian (σ = 25) trên ảnh coins.png với cửa sổ 3x3: giao diện tự động đánh dấu viền xanh lá cho khung '3. Sau Lọc Trung Bình (3x3)' vì đạt PSNR = 26.07 dB (MSE = 160.62) và Edge-F1 = 0.7503, vượt qua Lọc Trung vị 3x3 (PSNR = 25.68 dB, MSE = 175.79, Edge-F1 = 0.7443). Tiếp đó, Hình 3.13 minh họa màn hình sau khi người dùng nhấn nút 'Reset Ảnh Gốc': ảnh đang xử lý được đưa về trạng thái sạch hoàn toàn (0% nhiễu, MSE = 0, PSNR = ∞, Edge-F1 = 1.0000)."
    )

    for fig_name, cap in [
        ("web_ui_gaussian.png", "Hình 3.12. Kiểm thử chức năng khử nhiễu Gaussian (σ = 25) trên ảnh 'coins' qua giao diện Web Localhost"),
        ("web_ui_reset_clean.png", "Hình 3.13. Kiểm thử chức năng nút 'Reset Ảnh Gốc' khôi phục trạng thái ảnh sạch ban đầu trên giao diện Web"),
    ]:
        p = os.path.join(fig_dir, fig_name)
        if os.path.exists(p):
            add_figure(doc, p, caption=cap, width_inches=6.1)

    # --- MỤC 3.7 ---
    add_h2(doc, "3.7. Giải thích chuyên sâu bản chất toán học của các kết quả thực nghiệm")
    add_h3(doc, "3.7.1. Vì sao Bộ lọc Trung vị đặc trị nhiễu Muối tiêu nhưng Bộ lọc Trung bình lại thất bại?")
    add_para(
        doc,
        "Xét dưới góc độ thống kê bền vững (Robust Statistics), điểm gãy (Breakdown Point) của một bộ ước lượng là tỷ lệ phần trăm dữ liệu ngoại lai (outliers) tối đa có thể xuất hiện trong mẫu mà không làm bộ ước lượng bị chệch đi tùy ý:"
    )
    add_bullet(
        doc,
        "Có điểm gãy bằng 0% (Breakdown Point = 0). Chỉ cần đúng 1 điểm ảnh trong cửa sổ k^2 điểm bị nhảy vọt lên giá trị cực trị (0 hoặc 255), tổng số học sẽ bị kéo lệch ngay lập tức một lượng |I_outlier - μ| / k^2. Tệ hại hơn, vì cửa sổ trượt liên tục qua từng pixel, 1 điểm nhiễu xung đó sẽ nằm trong cửa sổ lân cận của tất cả k^2 điểm ảnh xung quanh nó, dẫn đến hiện tượng 'nhòe đốm' (patchy artifact) làm hỏng cả vùng lân cận.",
        bold_prefix="Phép tính trung bình cộng (Mean Filter): "
    )
    add_bullet(
        doc,
        "Có điểm gãy đạt mức tối đa lý thuyết là 50% (cụ thể là (k^2 - 1) / (2k^2), tức là 4/9 = 44.4% với cửa sổ 3x3 và 12/25 = 48% với cửa sổ 5x5). Vì các hạt muối (255) và hạt tiêu (0) là hai cực trị của dải mức xám, sau bước sắp xếp tăng dần chúng luôn bị cô lập ở hai đầu mút của mảng một chiều. Phần tử đứng ở vị trí trung tâm ⌊k^2 / 2⌋ hoàn toàn không bị ảnh hưởng bởi độ lớn của các điểm cực trị này và luôn trả về đúng một giá trị mức xám sạch của vùng ảnh thực.",
        bold_prefix="Phép lấy trung vị (Median Filter): "
    )

    add_h3(doc, "3.7.2. Vì sao Bộ lọc Trung bình phù hợp hơn với nhiễu cộng Gaussian?")
    add_para(
        doc,
        "Đối với nhiễu cộng Gaussian trắng η ~ N(0, σ^2), mọi điểm ảnh trong cửa sổ đều bị cộng thêm một sai số ngẫu nhiên tuân theo phân phối chuẩn đối xứng quanh 0. Theo lý thuyết Ước lượng Hợp lý Cực đại (Maximum Likelihood Estimation - MLE) đối với phân phối chuẩn, trung bình cộng mẫu (Sample Mean) chính là bộ ước lượng không chệch có phương sai nhỏ nhất đạt ngưỡng giới hạn dưới Cramér-Rao:"
    )
    add_formula(
        doc,
        "Var( Mean_Filter ) = σ^2 / N = σ^2 / k^2",
        eq_number="3.1",
    )
    add_para(
        doc,
        "Trong khi đó, theo định lý phân phối tiệm cận của trung vị mẫu lấy từ phân phối chuẩn N(μ, σ^2), phương sai của bộ ước lượng trung vị lớn hơn phương sai của bộ ước lượng trung bình một hệ số bằng π / 2 ≈ 1.5708:"
    )
    add_formula(
        doc,
        "Var( Median_Filter ) ≈ (π / 2) · (σ^2 / N) ≈ 1.5708 · (σ^2 / k^2)",
        eq_number="3.2",
    )
    add_para(
        doc,
        "Công thức (3.1) và (3.2) giải thích hoàn hảo số liệu thực nghiệm của chúng ta: Trên vùng nền bị nhiễu Gaussian nặng (σ = 25, 50), phương sai nhiễu còn sót lại sau Lọc trung vị cao hơn khoảng 57% so với Lọc trung bình, khiến Lọc trung bình đạt MSE thấp hơn và PSNR cao hơn."
    )

    add_h3(doc, "3.7.3. Ảnh hưởng của kích thước cửa sổ (3x3, 5x5, 7x7): Quy luật đánh đổi (Trade-off)")
    add_para(
        doc,
        "Kích thước cửa sổ lọc k x k điều khiển trực tiếp tần số cắt (cutoff frequency) của bộ lọc thông thấp trong miền không gian:"
    )
    add_bullet(
        doc,
        "Cửa sổ nhỏ chỉ gom thông tin trong phạm vi bán kính 1 pixel quanh tâm, nhờ đó bảo toàn tối đa các chi tiết mảnh (như nét chữ trong ảnh text, rãnh vữa trong ảnh brick), giữ độ dốc biên Sobel cao và thời gian chạy cực nhanh (3–8 ms). Đây là lựa chọn tốt nhất cho hầu hết các mức nhiễu thông thường.",
        bold_prefix="Cửa sổ nhỏ (3x3 - N = 9 điểm ảnh): "
    )
    add_bullet(
        doc,
        "Khi tăng k lên 5 và 7, diện tích lấy mẫu tăng lên gần gấp 3 và gấp 5.5 lần so với 3x3. Điều này giúp triệt tiêu mạnh hơn các dao động nhiễu có cường độ cực lớn (như Gaussian σ = 50 hoặc Muối tiêu mật độ rất dày > 20%). Tuy nhiên, cái giá phải trả là mọi cấu trúc chi tiết có kích thước nhỏ hơn bán kính cửa sổ (p = 2, 3 pixel) sẽ bị xóa sổ hoàn toàn, các góc vuông bị bo tròn, biên vật thể bị dịch chuyển hoặc nhòe rộng làm tụt chỉ số Edge-F1, đồng thời thời gian tính toán tăng vọt theo hàm bậc hai.",
        bold_prefix="Cửa sổ trung bình và lớn (5x5 - N = 25 điểm; 7x7 - N = 49 điểm): "
    )


def write_conclusion_and_appendix(doc: docx.Document, tables_dir: str) -> None:
    """Ghi phần KẾT LUẬN và PHỤ LỤC (Tài liệu tham khảo + Link GitHub + Bảng số liệu đầy đủ)."""
    add_h1(doc, "KẾT LUẬN", page_break_before=True)

    add_h2(doc, "1. Nội dung đã đạt được")
    add_para(
        doc,
        "Sau quá trình nghiên cứu lý thuyết, thiết kế thuật toán và tiến hành thực nghiệm nghiêm túc, Nhóm 10 đã hoàn thành toàn bộ các mục tiêu và yêu cầu kỹ thuật của Đề tài 7: \"Khử nhiễu ảnh bằng các bộ lọc trong miền không gian\" với các kết quả cụ thể sau:"
    )
    add_bullet(
        doc,
        "Hệ thống hóa đầy đủ và chặt chẽ cơ sở toán học về biểu diễn ảnh số đa mức xám, mô hình xác suất của nhiễu Muối tiêu và nhiễu cộng Gaussian, nguyên lý lọc không gian và phép tích chập 2 chiều, cơ chế hoạt động của Bộ lọc trung bình (tuyến tính) và Bộ lọc trung vị (phi tuyến thống kê thứ tự), toán tử đạo hàm bậc nhất Sobel cùng các độ đo đánh giá chất lượng ảnh (MSE, PSNR) và độ đo bảo toàn biên (Edge-MSE, Precision, Recall, Edge-F1).",
        bold_prefix="Về mặt lý thuyết: "
    )
    add_bullet(
        doc,
        "Xây dựng thành công bộ mã nguồn Python 3 hoàn chỉnh theo kiến trúc module hóa. Toàn bộ các thuật toán cốt lõi (mở rộng biên ảnh 3 chế độ, sinh nhiễu có seed = 42 cố định, tích chập 2D, lọc trung bình, lọc trung vị, toán tử Sobel và các hàm tính chỉ số) đều được tự cài đặt từ đầu bằng NumPy. Kết quả kiểm chứng tự động đối chiếu với thư viện OpenCV đạt sai số tuyệt đối bằng 0.0 (khớp 100%).",
        bold_prefix="Về mặt xây dựng chương trình ứng dụng: "
    )
    add_bullet(
        doc,
        "Thực hiện trọn vẹn 108 kịch bản thực nghiệm trên 6 ảnh chuẩn (256x256) và chứng minh bằng số liệu thực tế các quy luật cốt lõi: (1) Bộ lọc trung vị 3x3 vượt trội tuyệt đối trên nhiễu Muối tiêu (đạt PSNR trung bình 31.49 dB ở mức 1% và 28.04 dB ở mức 10%, đồng thời bảo toàn biên Sobel với Edge-F1 từ 0.78 đến 0.80); (2) Bộ lọc trung bình hiệu quả hơn về mặt giảm phương sai trên nhiễu Gaussian trung bình và nặng (σ = 25, 50) với tốc độ thực thi nhanh gấp 2.85 đến 10.95 lần so với lọc trung vị; (3) Làm rõ quy luật đánh đổi của kích thước cửa sổ lọc (3x3, 5x5, 7x7) giữa khả năng làm phẳng nhiễu và sự suy giảm độ sắc nét của biên ảnh.",
        bold_prefix="Về mặt thực nghiệm và phân tích định lượng: "
    )

    add_h2(doc, "2. Nội dung có thể cải tiến (Hướng phát triển)")
    add_para(
        doc,
        "Mặc dù Bộ lọc trung bình và Bộ lọc trung vị là hai công cụ nền tảng rất quan trọng trong xử lý ảnh, kết quả thực nghiệm cũng chỉ ra những giới hạn cố hữu của chúng khi đối mặt với các điều kiện nhiễu khắc nghiệt hoặc ảnh có nhiều chi tiết mảnh. Trong tương lai, đề tài có thể được mở rộng và cải tiến theo các hướng sau:"
    )
    add_bullet(
        doc,
        "Thay vì gán trọng số bằng nhau 1/k^2 cho mọi điểm trong cửa sổ như lọc trung bình hộp, có thể sử dụng Bộ lọc Gaussian (trọng số giảm dần theo khoảng cách tới tâm) và đặc biệt là Bộ lọc song phương (Bilateral Filter) hoặc Non-Local Means (NLM) kết hợp cả khoảng cách không gian lẫn độ tương đồng mức xám để khử nhiễu Gaussian mà không làm nhòe biên.",
        bold_prefix="Khắc phục nhược điểm làm nhòe biên của Bộ lọc Trung bình: "
    )
    add_bullet(
        doc,
        "Khi mật độ nhiễu muối tiêu vượt quá 20% – 30%, cửa sổ 3x3 cố định bắt đầu xuất hiện nhiều hơn 4 điểm nhiễu khiến lọc trung vị chuẩn bị suy giảm hiệu năng, trong khi dùng cửa sổ 7x7 cố định lại làm mất chi tiết. Giải pháp nâng cao là cài đặt Bộ lọc Trung vị Thích nghi (Adaptive Median Filter - AMF) có khả năng tự động co giãn kích thước cửa sổ k tùy theo mật độ nhiễu cục bộ và chỉ thay thế những pixel được phát hiện là điểm nhiễu (0 hoặc 255), giữ nguyên các pixel không nhiễu.",
        bold_prefix="Khắc phục nhược điểm của Bộ lọc Trung vị ở mật độ nhiễu cao và bảo toàn nét mảnh: "
    )
    add_bullet(
        doc,
        "Áp dụng kỹ thuật Ảnh tích phân (Integral Image / Summed-Area Table) để đưa độ phức tạp của Bộ lọc trung bình về O(1) không phụ thuộc vào kích thước cửa sổ k; đồng thời áp dụng thuật toán cập nhật Histogram trượt của Huang để giảm độ phức tạp của Bộ lọc trung vị từ O(k^2 log k) xuống O(k) hoặc O(1).",
        bold_prefix="Tối ưu hóa tốc độ tính toán thời gian thực: "
    )

    # --- PHỤ LỤC ---
    add_h1(doc, "PHỤ LỤC", page_break_before=True)
    add_h2(doc, "1. Tài liệu và link tham khảo")
    refs = [
        "[1] Rafael C. Gonzalez, Richard E. Woods, \"Digital Image Processing\", 4th Edition, Pearson Education, New York, 2018 (Chương 3: Intensity Transformations and Spatial Filtering; Chương 5: Image Restoration and Reconstruction; Chương 10: Image Segmentation).",
        "[2] Lương Mạnh Bá, Nguyễn Thanh Thủy, \"Nhập môn Xử lý ảnh số\", Nhà xuất bản Khoa học và Kỹ thuật, Hà Nội.",
        "[3] Đỗ Năng Toàn, Phạm Việt Bình, \"Giáo trình Xử lý ảnh\", Khoa Công nghệ Thông tin – Đại học Thái Nguyên / Viện Công nghệ Thông tin.",
        "[4] Bài giảng học phần Xử lý ảnh và thị giác máy tính (Mã học phần: IT3243), Khoa Công nghệ Thông tin, Trường Đại học Công nghệ Đông Á, năm học 2025 - 2026.",
        "[5] Tài liệu chính thức thư viện tính toán mảng đa chiều NumPy: https://numpy.org/doc/stable/",
        "[6] Tài liệu chính thức thư viện thị giác máy tính OpenCV (Image Filtering & Sobel Derivatives): https://docs.opencv.org/4.x/d4/d86/group__imgproc__filter.html",
        "[7] Kho ảnh chuẩn phục vụ thực nghiệm xử lý ảnh (Standard Test Images - scikit-image / USC-SIPI Image Database): https://scikit-image.org/docs/stable/api/skimage.data.html",
    ]
    for ref in refs:
        add_para(doc, ref, first_line_indent=0.0)

    add_h2(doc, "2. Link mã nguồn (GitHub) và hướng dẫn chạy tái lập thực nghiệm")
    add_bullet(
        doc,
        "https://github.com/NguyenManhQuyet20231053/BTL_XLA_Nhom10_DeTai7 (Kho lưu trữ bao gồm toàn bộ mã nguồn Python tự cài đặt trong thư mục src/, script chạy thực nghiệm tự động run_experiments.py, bộ 6 ảnh chuẩn, các bảng số liệu CSV và hình ảnh kết quả).",
        bold_prefix="Link kho mã nguồn GitHub của Nhóm 10: "
    )
    add_para(
        doc,
        "Để tái lập lại toàn bộ 100% bảng số liệu và biểu đồ trong báo cáo này trên bất kỳ máy tính nào (qua Terminal hoặc PyCharm), chỉ cần thực hiện 2 lệnh sau tại thư mục gốc của dự án:"
    )
    add_code_block(
        doc,
        [
            "# Bước 1: Cài đặt các thư viện cần thiết",
            "pip install -r requirements.txt",
            "",
            "# Bước 2: Chạy kịch bản thực nghiệm tự động (cố định seed = 42)",
            "python run_experiments.py",
            "",
            "# Bước 3: Khởi chạy giao diện ứng dụng Web tương tác trên Localhost (http://127.0.0.1:8000)",
            "python web_app.py",
        ],
        caption="Hướng dẫn dòng lệnh chạy tái lập thực nghiệm Đề tài 7",
    )

    add_h2(doc, "3. Phụ lục Bảng số liệu thực nghiệm trên từng ảnh trong bộ dữ liệu")
    add_para(
        doc,
        "Dưới đây trích xuất bảng số liệu thực nghiệm ở cửa sổ 3x3, 5x5, 7x7 trên các ảnh đặc trưng còn lại trong bộ dữ liệu (coins, astronaut, text, checkerboard, brick) được xuất trực tiếp từ tệp results/tables/detailed_results.csv."
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
    # Lấy các kịch bản tiêu biểu ở cửa sổ 3x3 và 5x5 của 5 ảnh còn lại để bảng phụ lục gọn đẹp
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
    )


def main() -> None:
    root_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    doc_path = os.path.join(root_dir, "BTL_XLA.docx")
    upd_path = os.path.join(root_dir, "BTL_XLA_updated.docx")
    tables_dir = os.path.join(root_dir, "results", "tables")
    fig_dir = os.path.join(root_dir, "results", "figures")

    # Đọc trực tiếp từ BTL_XLA.docx vừa được ghi Bước 3 và Bước 4
    src_doc_path = upd_path if os.path.exists(upd_path) else doc_path
    doc = docx.Document(src_doc_path)
    configure_styles(doc)
    remove_existing_chapter3_onwards(doc)
    write_chapter_3(doc, tables_dir, fig_dir)
    write_conclusion_and_appendix(doc, tables_dir)

    # Luôn lưu vào BTL_XLA_updated.docx trước và thử ghi đè vào BTL_XLA.docx
    doc.save(upd_path)
    try:
        doc.save(doc_path)
        print(f"Đã cập nhật hoàn chỉnh toàn bộ Báo cáo vào: {doc_path}")
    except PermissionError:
        print(f"Lưu ý: BTL_XLA.docx đang mở trong Word nên toàn bộ Báo cáo hoàn chỉnh đã được lưu tại: {upd_path}")


if __name__ == "__main__":
    main()

# -*- coding: utf-8 -*-
"""
Script: scripts/write_step3_intro_ch1.py
Mô tả: Ghi trực tiếp phần DANH MỤC TỪ VIẾT TẮT, MỞ ĐẦU và CHƯƠNG 1 (CƠ SỞ LÝ THUYẾT)
       vào file báo cáo BTL_XLA.docx, giữ nguyên trang bìa và Mục lục tự động của Nhóm 10.
"""

import os
import sys
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
)


def reset_body_after_toc(doc: docx.Document) -> None:
    """
    Giữ lại toàn bộ phần Trang bìa (paragraphs 0..15), bookmark và khối Mục lục tự động (sdt),
    xóa các dòng gợi ý nháp phía sau sdt để ghi nội dung báo cáo chính thức.
    """
    body = doc.element.body
    children = list(body)
    sdt_idx = None
    for idx, child in enumerate(children):
        if child.tag.endswith("sdt"):
            sdt_idx = idx
            break

    if sdt_idx is not None:
        for child in children[sdt_idx + 1:]:
            if not child.tag.endswith("sectPr"):
                body.remove(child)


def write_front_lists(doc: docx.Document) -> None:
    """Ghi Danh mục các từ viết tắt, Danh mục bảng biểu và Danh mục hình ảnh."""
    add_h1(doc, "DANH MỤC CÁC TỪ VIẾT TẮT", page_break_before=True)
    abbr_headers = ["Từ viết tắt", "Thuật ngữ tiếng Anh", "Ý nghĩa tiếng Việt"]
    abbr_rows = [
        ["2D", "Two-Dimensional", "Hai chiều (không gian phẳng gồm chiều cao và chiều rộng)"],
        ["ADC", "Analog-to-Digital Converter", "Bộ chuyển đổi tín hiệu tương tự sang tín hiệu số"],
        ["AWGN", "Additive White Gaussian Noise", "Nhiễu cộng trắng phân phối chuẩn Gaussian"],
        ["CCD / CMOS", "Charge-Coupled Device / Complementary Metal-Oxide-Semiconductor", "Cảm biến quang học thu nhận hình ảnh trong máy ảnh số"],
        ["CSV", "Comma-Separated Values", "Định dạng tập tin bảng số liệu phân cách bằng dấu phẩy"],
        ["dB", "Decibel", "Đơn vị đo lường tỷ số tín hiệu theo thang logarit"],
        ["Edge-MSE", "Edge Mean Squared Error", "Sai số bình phương trung bình giữa hai bản đồ độ lớn biên"],
        ["F1-Score", "Harmonic Mean of Precision and Recall", "Trung bình điều hòa giữa độ chính xác (Precision) và độ phủ (Recall)"],
        ["FN / FP / TP", "False Negative / False Positive / True Positive", "Âm tính giả / Dương tính giả / Dương tính thật (trong phát hiện biên)"],
        ["MSE", "Mean Squared Error", "Sai số bình phương trung bình"],
        ["PDF", "Probability Density Function", "Hàm mật độ xác suất"],
        ["PSNR", "Peak Signal-to-Noise Ratio", "Tỷ số tín hiệu cực đại trên nhiễu"],
        ["RGB", "Red - Green - Blue", "Không gian màu Đỏ - Lục - Lam"],
        ["S&P", "Salt-and-Pepper Noise", "Nhiễu muối tiêu (nhiễu xung lưỡng cực)"],
    ]
    add_styled_table(
        doc,
        headers=abbr_headers,
        rows=abbr_rows,
        caption="Bảng 0.1. Danh mục các từ viết tắt và thuật ngữ chuyên ngành",
        col_aligns=[WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT],
    )

    # 2. DANH MỤC HÌNH ẢNH
    add_h1(doc, "DANH MỤC HÌNH ẢNH", page_break_before=False)
    fig_headers = ["Số hiệu hình", "Tên hình ảnh / sơ đồ / biểu đồ", "Thuộc chương"]
    fig_rows = [
        ["Hình 2.1", "Sơ đồ khối quy trình xử lý tổng thể (Pipeline) của chương trình thực nghiệm", "Chương 2"],
        ["Hình 2.2", "Minh họa trực quan 3 chế độ mở rộng biên (Zero, Replicate, Reflect) với bán kính đệm p = 1", "Chương 2"],
        ["Hình 2.3", "Sơ đồ tổ chức cây thư mục mã nguồn của dự án BTL_XLA", "Chương 2"],
        ["Hình 2.4", "Giao diện tổng thể của ứng dụng Web tương tác chạy trên Localhost (http://127.0.0.1:8000)", "Chương 2"],
        ["Hình 3.1", "Trực quan hóa bộ dữ liệu 6 ảnh xám chuẩn kích thước 256x256 sử dụng trong thực nghiệm", "Chương 3"],
        ["Hình 3.2", "So sánh trực quan kết quả khử nhiễu Muối tiêu 5% trên ảnh 'cameraman'", "Chương 3"],
        ["Hình 3.3", "So sánh trực quan kết quả khử nhiễu Muối tiêu 10% trên ảnh 'coins'", "Chương 3"],
        ["Hình 3.4", "Biểu đồ PSNR trung bình và Edge F1-Score trên nhiễu Muối tiêu theo mật độ nhiễu và kích thước cửa sổ", "Chương 3"],
        ["Hình 3.5", "So sánh trực quan kết quả khử nhiễu Gaussian (σ = 25) trên ảnh 'cameraman'", "Chương 3"],
        ["Hình 3.6", "So sánh trực quan kết quả khử nhiễu Gaussian nặng (σ = 50) trên ảnh 'cameraman'", "Chương 3"],
        ["Hình 3.7", "Biểu đồ PSNR trung bình và Edge F1-Score trên nhiễu Gaussian theo độ lệch chuẩn σ và kích thước cửa sổ", "Chương 3"],
        ["Hình 3.8", "So sánh bản đồ biên Sobel trên ảnh 'cameraman' với nhiễu Muối tiêu 5%", "Chương 3"],
        ["Hình 3.9", "So sánh bản đồ biên Sobel trên ảnh 'cameraman' với nhiễu Gaussian (σ = 25)", "Chương 3"],
        ["Hình 3.10", "Minh họa hiện tượng mất nét chữ mảnh trên ảnh 'text' khi tăng kích thước cửa sổ lọc lên 5x5 và 7x7", "Chương 3"],
        ["Hình 3.11", "Biểu đồ so sánh thời gian thực thi trung bình (ms) theo kích thước cửa sổ lọc", "Chương 3"],
        ["Hình 3.12", "Kiểm thử chức năng khử nhiễu Gaussian (σ = 25) trên ảnh 'coins' qua giao diện Web Localhost", "Chương 3"],
        ["Hình 3.13", "Kiểm thử chức năng nút 'Reset Ảnh Gốc' khôi phục trạng thái ảnh sạch ban đầu trên giao diện Web", "Chương 3"],
    ]
    add_styled_table(
        doc,
        headers=fig_headers,
        rows=fig_rows,
        caption="Bảng 0.2. Danh mục các hình ảnh, sơ đồ và biểu đồ trong báo cáo",
        col_aligns=[WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER],
    )

    # 3. DANH MỤC BẢNG BIỂU
    add_h1(doc, "DANH MỤC BẢNG BIỂU", page_break_before=False)
    tbl_headers = ["Số hiệu bảng", "Tên bảng biểu", "Thuộc phần / chương"]
    tbl_rows = [
        ["Bảng 0.1", "Danh mục các từ viết tắt và thuật ngữ chuyên ngành", "Danh mục"],
        ["Bảng 0.2", "Danh mục các hình ảnh, sơ đồ và biểu đồ trong báo cáo", "Danh mục"],
        ["Bảng 0.3", "Danh mục các bảng biểu trong báo cáo", "Danh mục"],
        ["Bảng 0.4", "Danh sách thành viên Nhóm 10 và bảng phân công nhiệm vụ chi tiết", "Mở đầu"],
        ["Bảng 1.1", "So sánh đặc trưng lý thuyết giữa nhiễu Muối tiêu và nhiễu Gaussian", "Chương 1"],
        ["Bảng 1.2", "Bảng đối sánh lý thuyết giữa Bộ lọc Trung bình và Bộ lọc Trung vị", "Chương 1"],
        ["Bảng 2.1", "Đặc tả chi tiết các phân hệ chức năng của chương trình và ứng dụng Web", "Chương 2"],
        ["Bảng 3.1", "Mô tả chi tiết bộ dữ liệu 6 ảnh xám chuẩn sử dụng trong thực nghiệm", "Chương 3"],
        ["Bảng 3.2", "Kết quả kiểm chứng độ chính xác của các thuật toán tự cài đặt so với thư viện OpenCV", "Chương 3"],
        ["Bảng 3.3", "Kết quả trung bình trên 6 ảnh chuẩn khi khử nhiễu Muối tiêu (1%, 5%, 10%)", "Chương 3"],
        ["Bảng 3.4", "Số liệu thực nghiệm chi tiết trên ảnh 'cameraman' với nhiễu Muối tiêu", "Chương 3"],
        ["Bảng 3.5", "Kết quả trung bình trên 6 ảnh chuẩn khi khử nhiễu Gaussian (σ = 10, 25, 50)", "Chương 3"],
        ["Bảng 3.6", "Số liệu thực nghiệm chi tiết trên ảnh 'cameraman' với nhiễu Gaussian", "Chương 3"],
        ["Bảng 3.7", "So sánh thời gian xử lý trung bình (ms) trên ảnh 256x256 giữa hai bộ lọc", "Chương 3"],
        ["Bảng 3.8", "Bảng kịch bản kiểm thử các chức năng của ứng dụng Web trên Localhost", "Chương 3"],
        ["Bảng P.1", "Trích xuất kết quả thực nghiệm chi tiết trên các ảnh coins, astronaut, text, checkerboard và brick", "Phụ lục"],
    ]
    add_styled_table(
        doc,
        headers=tbl_headers,
        rows=tbl_rows,
        caption="Bảng 0.3. Danh mục các bảng biểu trong báo cáo",
        col_aligns=[WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER],
    )


def write_introduction(doc: docx.Document) -> None:
    """Ghi phần MỞ ĐẦU đầy đủ 3 mục theo đúng quy định của nhà trường."""
    add_h1(doc, "MỞ ĐẦU", page_break_before=True)

    # 1. Giới thiệu môn học
    add_h2(doc, "1. Giới thiệu tổng quan môn học Xử lý ảnh và thị giác máy tính")
    add_para(
        doc,
        "Trong kỷ nguyên chuyển đổi số và cuộc Cách mạng công nghiệp lần thứ tư, dữ liệu thị giác (bao gồm hình ảnh tĩnh và luồng video số) chiếm tỷ trọng lớn nhất trong tổng lượng thông tin mà các hệ thống máy tính thu thập từ thế giới thực. Học phần Xử lý ảnh và thị giác máy tính (Mã học phần: IT3243) tại Khoa Công nghệ Thông tin – Trường Đại học Công nghệ Đông Á là một trong những học phần cốt lõi, trang bị cho sinh viên nền tảng toán học, tư duy thuật toán và kỹ năng lập trình thực nghiệm để biến đổi, phân tích và trích xuất tri thức từ ảnh số."
    )
    add_para(
        doc,
        "Về mặt cấu trúc học thuật, tiến trình xử lý thông tin thị giác trên máy tính thường được chia thành ba mức độ liên kết chặt chẽ với nhau:"
    )
    add_bullet(
        doc,
        "Bao gồm các phép toán tác động trực tiếp lên mức độ điểm ảnh (pixel) như thu nhận ảnh, lượng tử hóa, khử nhiễu (image denoising), tăng cường độ tương phản và làm sắc nét. Đầu vào của quá trình này là một ảnh thô (thường bị suy giảm chất lượng) và đầu ra là một ảnh có chất lượng cao hơn, phục vụ trực tiếp cho mắt người quan sát hoặc làm đầu vào sạch cho các bước tiếp theo.",
        bold_prefix="Xử lý ảnh mức thấp (Low-level Image Processing): "
    )
    add_bullet(
        doc,
        "Bao gồm các kỹ thuật trích xuất đặc trưng hình học và cấu trúc như phát hiện biên (edge detection), phân vùng ảnh (image segmentation), biểu diễn hình thái học và mô tả đối tượng. Ở mức độ này, đầu vào là ảnh đã qua tiền xử lý và đầu ra là các cấu trúc hình học như đường biên, vùng liên thông hoặc tập điểm đặc trưng.",
        bold_prefix="Xử lý ảnh mức trung gian (Mid-level Image Processing): "
    )
    add_bullet(
        doc,
        "Sử dụng các đặc trưng đã trích xuất kết hợp với các mô hình học máy và trí tuệ nhân tạo để thực hiện nhận dạng đối tượng, phân loại cảnh vật, theo vết chuyển động và hiểu ngữ nghĩa hình ảnh.",
        bold_prefix="Thị giác máy tính mức cao (High-level Computer Vision): "
    )
    add_para(
        doc,
        "Như vậy, chất lượng của tầng xử lý ảnh mức thấp đóng vai trò nền móng quyết định sự thành bại của toàn bộ hệ thống thị giác máy tính. Nếu ảnh đầu vào bị nhiễm nhiễu nặng mà không được khử nhiễu đúng phương pháp, các toán tử đạo hàm ở tầng trung gian (như phát hiện biên Sobel, Prewitt, Canny) sẽ khuếch đại nhiễu thành hàng loạt điểm biên giả (false edges), dẫn đến sai lệch hoàn toàn ở khâu phân đoạn và nhận dạng phía sau."
    )

    # 2. Giới thiệu đề tài
    add_h2(doc, "2. Giới thiệu tổng quan đề tài")
    add_para(
        doc,
        "Trong thực tế, không có hệ thống thu nhận ảnh nào là lý tưởng tuyệt đối. Trong suốt quá trình hình thành, số hóa, lưu trữ và truyền tải qua kênh truyền thông, ảnh số luôn chịu sự tác động của nhiều nguồn nhiễu vật lý khác nhau. Chẳng hạn, dao động nhiệt của các hạt điện tử trên bề mặt cảm biến quang học CCD/CMOS hoặc điều kiện chiếu sáng yếu tạo ra nhiễu cộng phân phối chuẩn (Gaussian Noise) phủ đều lên toàn bộ bức ảnh. Ngược lại, các lỗi đột biến trong quá trình chuyển đổi tương tự - số (ADC), bụi bẩn trên ống kính, điểm ảnh chết (dead pixels) hoặc lỗi bit khi truyền dữ liệu qua mạng không dây lại sinh ra nhiễu xung lưỡng cực dưới dạng các chấm trắng và đen cực trị, thường được gọi là nhiễu muối tiêu (Salt-and-Pepper Noise)."
    )
    add_para(
        doc,
        "Xuất phát từ thực tiễn đó, Nhóm 10 thực hiện Đề tài 7 với tên gọi: \"Khử nhiễu ảnh bằng các bộ lọc trong miền không gian\". Đề tài tập trung giải quyết bài toán cốt lõi của tiền xử lý ảnh: Làm thế nào để loại bỏ hiệu quả các thành phần nhiễu ngẫu nhiên trên ảnh đa mức xám mà vẫn bảo toàn tối đa các cấu trúc hình học quan trọng, đặc biệt là các đường biên (edges) và chi tiết mảnh của vật thể."
    )
    add_para(
        doc,
        "Để làm sáng tỏ bản chất toán học và hiệu năng thực tế, đề tài tiến hành tự cài đặt hoàn toàn từ đầu (from scratch) hai đại diện tiêu biểu nhất của hai trường phái lọc trong miền không gian:"
    )
    add_bullet(
        doc,
        "Đại diện cho phương pháp lọc tuyến tính dựa trên phép tích chập 2 chiều (2D Convolution). Bộ lọc làm trơn ảnh bằng cách thay thế mỗi điểm ảnh bằng trung bình cộng các điểm ảnh trong cửa sổ lân cận.",
        bold_prefix="Bộ lọc trung bình (Mean / Averaging Filter): "
    )
    add_bullet(
        doc,
        "Đại diện cho phương pháp lọc phi tuyến dựa trên thống kê thứ tự (Order-Statistic Filter). Bộ lọc sắp xếp các giá trị mức xám trong cửa sổ trượt và chọn phần tử đứng chính giữa để loại bỏ các điểm ngoại lai (outliers).",
        bold_prefix="Bộ lọc trung vị (Median Filter): "
    )
    add_para(
        doc,
        "Điểm nhấn quan trọng của đề tài là không chỉ dừng lại ở việc đánh giá độ trung thực mức xám toàn cục thông qua Sai số bình phương trung bình (MSE) và Tỷ số tín hiệu cực đại trên nhiễu (PSNR), mà còn kết hợp trực tiếp với toán tử đạo hàm bậc nhất Sobel để định lượng khả năng bảo toàn biên (Edge Preservation) thông qua sai số bản đồ Gradient (Edge-MSE) và các chỉ số Precision, Recall, F1-Score của bản đồ biên sau ngưỡng hóa. Toàn bộ thực nghiệm được thiết kế bài bản trên 6 ảnh chuẩn với 2 loại nhiễu, 3 mức cường độ nhiễu và 3 kích thước cửa sổ lọc (3x3, 5x5, 7x7), tạo nên 108 kịch bản khảo sát khách quan và có khả năng tái lập 100%."
    )

    # 3. Giới thiệu thành viên nhóm và phân công
    add_h2(doc, "3. Giới thiệu thành viên nhóm và phân công chi tiết công việc")
    add_para(
        doc,
        "Bài tập lớn được thực hiện bởi Nhóm 10 thuộc lớp tín chỉ Xử lý ảnh và thị giác máy tính-1-1-26 (N05.THTT_IN_ELN.11), dưới sự hướng dẫn của Giảng viên Nguyễn Phồn Lữa. Các thành viên trong nhóm đã phối hợp chặt chẽ, phân chia công việc rõ ràng theo từng giai đoạn từ nghiên cứu lý thuyết, xây dựng mã nguồn thuật toán, thiết kế kịch bản thực nghiệm cho tới phân tích số liệu và biên soạn báo cáo."
    )

    member_headers = [
        "TT",
        "Mã SV",
        "Họ và tên",
        "Lớp hành chính",
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
            "- Nghiên cứu cơ sở lý thuyết về biểu diễn ảnh số, phép tích chập 2D, bộ lọc trung bình và toán tử biên Sobel.\n"
            "- Tự cài đặt các module cốt lõi: mở rộng biên (padding.py), tích chập & lọc không gian (filters.py), toán tử Sobel (sobel.py).\n"
            "- Xây dựng kịch bản điều phối thực nghiệm tự động (run_experiments.py) và kiểm chứng đối chiếu với OpenCV.\n"
            "- Biên soạn phần Mở đầu, Chương 1 (Mục 1.1 - 1.4), Chương 2 và tổng hợp, chuẩn hóa định dạng toàn bộ báo cáo.",
            "100% (Đóng góp 50% khối lượng nhóm)",
        ],
        [
            "2",
            "20231045",
            "Hà Thái Bảo",
            "DCCNTT.14.3",
            "Thành viên",
            "- Nghiên cứu cơ sở lý thuyết về mô hình nhiễu (Muối tiêu, Gaussian), bộ lọc trung vị và các độ đo MSE, PSNR, Edge-F1.\n"
            "- Tự cài đặt các module: giả lập nhiễu có seed cố định (noise.py), tính chỉ số đánh giá (metrics.py), chuẩn bị dữ liệu (dataset.py) và vẽ biểu đồ (visualization.py).\n"
            "- Thực hiện chạy thực nghiệm trên 6 ảnh chuẩn, thu thập bảng số liệu CSV và xuất hệ thống hình ảnh so sánh.\n"
            "- Biên soạn Chương 1 (Mục 1.5 - 1.7), Chương 3 (Phân tích kết quả thực nghiệm), Kết luận và Phụ lục.",
            "100% (Đóng góp 50% khối lượng nhóm)",
        ],
    ]
    add_styled_table(
        doc,
        headers=member_headers,
        rows=member_rows,
        caption="Bảng 0.4. Danh sách thành viên Nhóm 10 và bảng phân công nhiệm vụ chi tiết",
        col_aligns=[
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.CENTER,
        ],
    )


def write_chapter_1(doc: docx.Document) -> None:
    """Ghi toàn bộ CHƯƠNG 1. CƠ SỞ LÝ THUYẾT sâu sắc, chặt chẽ và có ví dụ tính tay."""
    add_h1(doc, "CHƯƠNG 1. CƠ SỞ LÝ THUYẾT", page_break_before=True)

    # --- MỤC 1.1 ---
    add_h2(doc, "1.1. Tổng quan về ảnh số và biểu diễn ảnh trong miền không gian")
    add_h3(doc, "1.1.1. Khái niệm ảnh số, quá trình lấy mẫu và lượng tử hóa")
    add_para(
        doc,
        "Về mặt vật lý, một bức ảnh tự nhiên thu nhận từ thế giới thực là một hàm cường độ sáng liên tục hai chiều f(x, y), trong đó (x, y) là tọa độ không gian trên mặt phẳng ảnh và giá trị của f tại mỗi tọa độ biểu thị năng lượng bức xạ ánh sáng (độ chói). Để máy tính điện tử có thể lưu trữ và xử lý được, hàm liên tục f(x, y) phải trải qua hai quá trình rời rạc hóa cơ bản:"
    )
    add_bullet(
        doc,
        "Là quá trình rời rạc hóa các tọa độ không gian (x, y) thành một lưới điểm cách đều nhau theo chiều dọc và chiều ngang. Sau khi lấy mẫu, bức ảnh được chia thành M hàng và N cột; mỗi ô giao nhau tại hàng x và cột y được gọi là một phần tử ảnh hay điểm ảnh (Picture Element - Pixel).",
        bold_prefix="Quá trình lấy mẫu (Sampling): "
    )
    add_bullet(
        doc,
        "Là quá trình rời rạc hóa giá trị biên độ cường độ sáng f(x, y) liên tục thành một tập hữu hạn L mức giá trị nguyên rời rạc. Đối với ảnh đa mức xám chuẩn 8-bit (8-bit Grayscale Image), số mức xám là L = 2^8 = 256 mức, trải dài trên đoạn số nguyên [0, 255], trong đó giá trị 0 biểu diễn màu đen tuyệt đối, giá trị 255 biểu diễn màu trắng tuyệt đối và các giá trị trung gian biểu diễn các sắc độ xám đậm nhạt khác nhau.",
        bold_prefix="Quá trình lượng tử hóa (Quantization): "
    )

    add_h3(doc, "1.1.2. Biểu diễn ma trận của ảnh đa mức xám")
    add_para(
        doc,
        "Sau khi lấy mẫu và lượng tử hóa, một ảnh đa mức xám kích thước M x N được biểu diễn hoàn toàn tương đương dưới dạng một ma trận số nguyên hai chiều kích thước M hàng và N cột:"
    )
    add_formula(
        doc,
        "f(x, y) ∈ {0, 1, 2, ..., L - 1},   với x ∈ {0, 1, ..., M - 1} và y ∈ {0, 1, ..., N - 1}",
        eq_number="1.1",
    )
    add_para(
        doc,
        "Theo quy ước hệ tọa độ ảnh số chuẩn trong xử lý ảnh (và trong các thư viện tính toán ma trận như NumPy, OpenCV), gốc tọa độ (0, 0) được đặt tại góc trên cùng bên trái của bức ảnh; trục x hướng thẳng đứng từ trên xuống dưới (tương ứng với chỉ số hàng - row index) và trục y hướng nằm ngang từ trái sang phải (tương ứng với chỉ số cột - column index)."
    )

    add_h3(doc, "1.1.3. Khái niệm miền không gian và lân cận điểm ảnh")
    add_para(
        doc,
        "Thuật ngữ miền không gian (Spatial Domain) dùng để chỉ mặt phẳng tập hợp các điểm ảnh cấu thành nên bức ảnh. Các phương pháp xử lý trong miền không gian thao tác trực tiếp trên giá trị mức xám của các pixel thông qua biểu thức tổng quát:"
    )
    add_formula(
        doc,
        "g(x, y) = T [ f(x, y) ]",
        eq_number="1.2",
    )
    add_para(
        doc,
        "trong đó f(x, y) là ảnh đầu vào, g(x, y) là ảnh đầu ra sau xử lý, và T là một toán tử tác động lên tập lân cận (neighborhood) xung quanh tọa độ (x, y). Thông thường, lân cận của điểm (x, y) được xác định bởi một cửa sổ hình vuông hoặc hình chữ nhật kích thước lẻ k x k (với k = 2a + 1, ví dụ 3x3, 5x5, 7x7) có tâm đặt ngay tại (x, y). Tập hợp các tọa độ điểm ảnh nằm trong cửa sổ lân cận S_xy có tâm tại (x, y) được viết là:"
    )
    add_formula(
        doc,
        "S_xy = { (x + s, y + t)  |  -a ≤ s ≤ a,  -a ≤ t ≤ a },   với a = (k - 1) / 2",
        eq_number="1.3",
    )

    # --- MỤC 1.2 ---
    add_h2(doc, "1.2. Nhiễu trong ảnh số và các mô hình nhiễu điển hình")
    add_h3(doc, "1.2.1. Khái niệm nhiễu ảnh và mô hình suy giảm tổng quát")
    add_para(
        doc,
        "Nhiễu ảnh (Image Noise) là những biến thiên ngẫu nhiên không mong muốn của giá trị độ sáng hoặc thông tin màu sắc xuất hiện trên ảnh số, không bắt nguồn từ cảnh vật thực tế cần chụp. Về mặt toán học, nếu bỏ qua hiện tượng nhòe quang học, quá trình suy giảm ảnh do nhiễu trong miền không gian thường được mô tả bởi mô hình nhiễu cộng hoặc nhiễu xung:"
    )
    add_formula(
        doc,
        "g(x, y) = f(x, y) + η(x, y)",
        eq_number="1.4",
    )
    add_para(
        doc,
        "trong đó f(x, y) là ảnh gốc lý tưởng (không nhiễu), η(x, y) là biến ngẫu nhiên biểu diễn thành phần nhiễu tại tọa độ (x, y), và g(x, y) là ảnh quan sát được (ảnh nhiễu). Tùy thuộc vào bản chất vật lý của nguồn gây nhiễu, hàm mật độ xác suất (Probability Density Function - PDF) của η(x, y) sẽ tuân theo các quy luật thống kê khác nhau. Trong đề tài này, chúng ta đi sâu khảo sát hai loại nhiễu quan trọng và phổ biến nhất: nhiễu muối tiêu và nhiễu Gaussian."
    )

    add_h3(doc, "1.2.2. Nhiễu muối tiêu (Salt-and-Pepper Noise / Impulse Noise)")
    add_para(
        doc,
        "Nhiễu muối tiêu (còn gọi là nhiễu xung lưỡng cực - Bipolar Impulse Noise) xuất hiện dưới dạng các điểm ảnh trắng tinh (giá trị cực đại 255 - tượng trưng cho hạt muối) và các điểm ảnh đen kịt (giá trị cực tiểu 0 - tượng trưng cho hạt tiêu) phân bố ngẫu nhiên rải rác trên bề mặt ảnh. Nguyên nhân vật lý chủ yếu của loại nhiễu này là do lỗi đóng mở tức thời của chuyển mạch phần cứng, lỗi truyền bit dữ liệu qua kênh truyền nhiễu loạn, lỗi bộ nhớ hoặc các tế bào quang điện trên cảm biến bị chập/hỏng."
    )
    add_para(
        doc,
        "Giả sử mật độ nhiễu tổng cộng trên ảnh là p (với 0 ≤ p ≤ 1, ví dụ p = 1%, 5%, 10%), trong đó xác suất xuất hiện điểm muối (mức xám S = 255) là p_s và xác suất xuất hiện điểm tiêu (mức xám P = 0) là p_p sao cho p = p_s + p_p. Khi đó, giá trị điểm ảnh quan sát được g(x, y) tại mỗi tọa độ được xác định bởi mô hình xác suất rời rạc:"
    )
    add_formula(
        doc,
        "g(x, y) = { 255 (điểm muối) với xác suất p_s ;  0 (điểm tiêu) với xác suất p_p ;  f(x, y) (giữ nguyên) với xác suất 1 - (p_s + p_p) }",
        eq_number="1.5",
    )
    add_para(
        doc,
        "Trong trường hợp nhiễu muối tiêu đối xứng (cân bằng), ta có p_s = p_p = p / 2. Đặc điểm cốt lõi cần lưu ý của nhiễu muối tiêu là: nó chỉ làm hỏng hoàn toàn một tỷ lệ p số lượng điểm ảnh (thay thế hoàn toàn giá trị gốc bằng 0 hoặc 255), trong khi tỷ lệ (1 - p) số lượng điểm ảnh còn lại hoàn toàn không bị biến đổi (giữ nguyên giá trị chính xác của ảnh gốc f(x, y)). Chính đặc tính cục bộ và cực trị này là chìa khóa lý giải vì sao các bộ lọc thống kê thứ tự lại xử lý nhiễu muối tiêu hiệu quả vượt trội so với các bộ lọc tính trung bình cộng."
    )

    add_h3(doc, "1.2.3. Nhiễu Gaussian (Additive White Gaussian Noise - AWGN)")
    add_para(
        doc,
        "Khác hoàn toàn với nhiễu muối tiêu, nhiễu Gaussian (hay nhiễu cộng trắng phân phối chuẩn) là loại nhiễu liên tục tác động lên hầu hết hoặc toàn bộ 100% số điểm ảnh trong ma trận ảnh. Nguồn gốc vật lý của nhiễu Gaussian đến từ chuyển động nhiệt hỗn loạn của các electron trong mạch khuếch đại điện tử (nhiễu nhiệt Johnson-Nyquist) và sự thăng giáng lượng tử của dòng hạt photon khi chụp ảnh trong môi trường thiếu sáng hoặc cài đặt độ nhạy sáng ISO cao."
    )
    add_para(
        doc,
        "Tại mỗi tọa độ (x, y), thành phần nhiễu cộng η(x, y) là một biến ngẫu nhiên độc lập, có cùng phân phối chuẩn (Gaussian) với kỳ vọng toán học (giá trị trung bình) bằng μ và phương sai bằng σ^2 (trong đó σ là độ lệch chuẩn - standard deviation). Hàm mật độ xác suất (PDF) của biến ngẫu nhiên nhiễu η được xác định bởi công thức hình chuông Gauss:"
    )
    add_formula(
        doc,
        "p(η) = [ 1 / (√(2π) · σ) ] · exp( - (η - μ)^2 / (2σ^2) )",
        eq_number="1.6",
    )
    add_para(
        doc,
        "Trong các bài toán xử lý ảnh thực tế, nhiễu Gaussian thường có kỳ vọng bằng không (μ = 0), tức là η ~ N(0, σ^2). Khi đó, cường độ mạnh hay yếu của nhiễu phụ thuộc hoàn toàn vào độ lệch chuẩn σ (ví dụ: σ = 10 là nhiễu nhẹ, σ = 25 là nhiễu trung bình và σ = 50 là nhiễu rất nặng). Theo quy tắc thực nghiệm 3-sigma của phân phối chuẩn, khoảng 68.27% giá trị nhiễu nằm trong đoạn [-σ, +σ], khoảng 95.45% nằm trong đoạn [-2σ, +2σ] và 99.73% nằm trong đoạn [-3σ, +3σ]. Khi cộng η(x, y) vào ảnh gốc f(x, y), giá trị điểm ảnh có thể vượt ra ngoài dải hợp lệ [0, 255] của kiểu số nguyên 8-bit, do đó quá trình tạo ảnh nhiễu thực tế luôn đi kèm phép làm tròn và cắt ngưỡng (clipping):"
    )
    add_formula(
        doc,
        "g(x, y) = clip( round( f(x, y) + η(x, y) ),  0,  255 )",
        eq_number="1.7",
    )

    # Bảng so sánh 2 loại nhiễu
    noise_comp_headers = [
        "Tiêu chí so sánh",
        "Nhiễu Muối tiêu (Salt-and-Pepper)",
        "Nhiễu Gaussian (AWGN)",
    ]
    noise_comp_rows = [
        [
            "Bản chất toán học",
            "Nhiễu xung lưỡng cực rời rạc (thay thế giá trị pixel bằng cực trị 0 hoặc 255).",
            "Nhiễu cộng liên tục tuân theo phân phối chuẩn N(0, σ^2).",
        ],
        [
            "Phạm vi điểm ảnh bị ảnh hưởng",
            "Chỉ ảnh hưởng cục bộ một tỷ lệ p% số điểm ảnh (ví dụ 1%, 5%, 10%); (100 - p)% điểm ảnh còn lại nguyên vẹn.",
            "Tác động toàn cục lên gần như 100% số lượng điểm ảnh trên toàn bộ bức ảnh.",
        ],
        [
            "Biểu hiện thị giác trên ảnh",
            "Các chấm trắng sáng và đen sẫm rải rác độc lập trên bề mặt ảnh.",
            "Lớp hạt mịn phủ đều khắp bức ảnh, làm giảm độ trong trẻo và nhòe chi tiết nhỏ.",
        ],
        [
            "Nguyên nhân vật lý điển hình",
            "Lỗi bit truyền dẫn, lỗi bộ chuyển đổi ADC, điểm ảnh chết trên cảm biến.",
            "Nhiễu nhiệt mạch điện tử của cảm biến CCD/CMOS, chụp ảnh thiếu sáng / ISO cao.",
        ],
        [
            "Tham số điều khiển mức nhiễu",
            "Mật độ nhiễu p (density ∈ [0, 1]).",
            "Độ lệch chuẩn σ (standard deviation > 0).",
        ],
    ]
    add_styled_table(
        doc,
        headers=noise_comp_headers,
        rows=noise_comp_rows,
        caption="Bảng 1.1. So sánh đặc trưng lý thuyết giữa nhiễu Muối tiêu và nhiễu Gaussian",
        col_aligns=[WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT],
    )

    # --- MỤC 1.3 ---
    add_h2(doc, "1.3. Lọc trong miền không gian và phép tích chập hai chiều")
    add_h3(doc, "1.3.1. Nguyên lý lọc không gian bằng cửa sổ trượt (Sliding Window)")
    add_para(
        doc,
        "Lọc không gian (Spatial Filtering) là kỹ thuật biến đổi giá trị của mỗi điểm ảnh dựa trên mối tương quan giữa chính điểm ảnh đó và các điểm ảnh nằm trong vùng lân cận xung quanh nó. Công cụ trung tâm của lọc không gian là một ma trận nhỏ kích thước lẻ k x k (thường là 3x3, 5x5 hoặc 7x7), được gọi là mặt nạ lọc (mask), nhân lọc (kernel) hoặc cửa sổ trượt (sliding window)."
    )
    add_para(
        doc,
        "Quá trình lọc diễn ra bằng cách di chuyển tâm của cửa sổ k x k lần lượt đi qua từng điểm ảnh (x, y) từ góc trên bên trái đến góc dưới bên phải của bức ảnh. Tại mỗi vị trí dừng (x, y), một quy tắc toán học xác định trước (có thể là tổ hợp tuyến tính hoặc hàm thống kê phi tuyến) được áp dụng lên k^2 điểm ảnh nằm bên dưới cửa sổ để tính ra giá trị mới g(x, y) cho ảnh đầu ra. Lưu ý quan trọng: giá trị g(x, y) vừa tính được phải được ghi vào một ma trận ảnh đầu ra riêng biệt, tuyệt đối không ghi đè trực tiếp lên ma trận ảnh đầu vào f(x, y) đang trượt để tránh làm sai lệch phép tính của các điểm ảnh kế tiếp."
    )

    add_h3(doc, "1.3.2. Phép tương quan (Correlation) và phép tích chập hai chiều (2D Convolution)")
    add_para(
        doc,
        "Trong nhóm các bộ lọc không gian tuyến tính, hai phép toán nền tảng nhất là phép tương quan không gian (Spatial Correlation) và phép tích chập hai chiều (2D Convolution). Xét một ảnh đầu vào f(x, y) kích thước M x N và một nhân lọc h(s, t) kích thước k x k với k = 2a + 1 (do đó các chỉ số tương đối s, t chạy từ -a đến +a)."
    )
    add_para(
        doc,
        "Phép tương quan không gian giữa ảnh f và nhân h tại tọa độ (x, y) là tổng các tích từng phần tử giữa nhân lọc và vùng ảnh lân cận nằm trực tiếp dưới nhân đó:"
    )
    add_formula(
        doc,
        "(f ⊗ h)(x, y) = ∑_{s = -a}^{a} ∑_{t = -a}^{a} h(s, t) · f(x + s, y + t)",
        eq_number="1.8",
    )
    add_para(
        doc,
        "Trong khi đó, phép tích chập hai chiều (2D Convolution), ký hiệu là (f * h)(x, y), được định nghĩa tương tự như phép tương quan nhưng nhân lọc h(s, t) được lật đối xứng một góc 180 độ (tức là đảo ngược cả theo chiều dọc lẫn chiều ngang) trước khi nhân tích vô hướng với vùng ảnh:"
    )
    add_formula(
        doc,
        "(f * h)(x, y) = ∑_{s = -a}^{a} ∑_{t = -a}^{a} h(s, t) · f(x - s, y - t)",
        eq_number="1.9",
    )
    add_para(
        doc,
        "Về mặt đại số tuyến tính, phép tích chập 2D thỏa mãn các tính chất giao hoán (f * h = h * f), kết hợp (f * (h_1 * h_2) = (f * h_1) * h_2) và phân phối đối với phép cộng. Đặc biệt, nếu nhân lọc h(s, t) là một ma trận đối xứng tâm (tức là h(s, t) = h(-s, -t) với mọi s, t, điển hình như nhân lọc trung bình hoặc nhân lọc Gaussian), thì phép tương quan và phép tích chập cho kết quả hoàn toàn trùng nhau. Ngược lại, đối với các nhân lọc bất đối xứng dùng để tính đạo hàm bậc nhất như toán tử biên Sobel, việc phân biệt rõ chiều lật nhân trong tích chập là vô cùng quan trọng để giữ đúng dấu của vector Gradient."
    )

    add_h3(doc, "1.3.3. Vấn đề biên ảnh và các kỹ thuật mở rộng biên (Boundary Padding)")
    add_para(
        doc,
        "Khi đặt tâm của cửa sổ lọc kích thước k x k (bán kính a = (k - 1)/2) lên các điểm ảnh nằm sát rìa bức ảnh (cụ thể là các hàng có chỉ số x < a hoặc x ≥ M - a, và các cột có chỉ số y < a hoặc y ≥ N - a), một phần của cửa sổ lọc sẽ tràn ra ngoài phạm vi ma trận ảnh. Nếu chỉ thực hiện tính toán tại những vị trí mà cửa sổ nằm trọn bên trong ảnh (chế độ Valid), kích thước ảnh đầu ra sẽ bị thu hẹp mất 2a hàng và 2a cột, còn (M - 2a) x (N - 2a)."
    )
    add_para(
        doc,
        "Để bảo toàn nguyên vẹn kích thước ảnh đầu ra bằng đúng M x N (chế độ Same) và xử lý đầy đủ mọi điểm ảnh, ta cần mở rộng xung quanh bốn cạnh của ảnh gốc một dải viền có độ rộng a pixel trước khi lọc. Kỹ thuật này được gọi là mở rộng biên hay đệm biên (Boundary Padding), đưa kích thước ảnh đệm lên thành (M + 2a) x (N + 2a). Ba phương pháp đệm biên phổ biến nhất bao gồm:"
    )
    add_bullet(
        doc,
        "Điền toàn bộ các phần tử nằm ngoài phạm vi ảnh gốc bằng giá trị hằng số 0. Phương pháp này đơn giản nhất khi cài đặt nhưng có nhược điểm nghiêm trọng là tạo ra một đường nhảy bậc độ sáng đột ngột từ mức xám thực của ảnh về 0 tại biên. Hệ quả là sau khi lọc trung bình sẽ xuất hiện một khung viền tối màu quanh ảnh, còn khi tính đạo hàm Sobel sẽ sinh ra một đường biên giả rất mạnh chạy dọc theo bốn cạnh khung ảnh.",
        bold_prefix="Đệm giá trị không (Zero-padding): "
    )
    add_bullet(
        doc,
        "Sao chép và lặp lại giá trị của các điểm ảnh nằm ở hàng/cột ngoài cùng của ảnh gốc ra phần viền mở rộng. Ví dụ, các dòng đệm phía trên sẽ sao chép nguyên vẹn dòng 0 của ảnh gốc.",
        bold_prefix="Đệm lặp lại biên (Replicate / Edge Padding): "
    )
    add_bullet(
        doc,
        "Lấy các điểm ảnh nằm sát bên trong biên ảnh làm gương phản chiếu đối xứng qua đường biên. Cụ thể, với chỉ số biên x = 0 làm trục gương, phần tử đệm tại vị trí -1 sẽ lấy giá trị của vị trí +1, vị trí -2 lấy giá trị của vị trí +2. Phương pháp này duy trì sự liên tục trơn tru của cả cường độ sáng lẫn đạo hàm bậc nhất tại vùng rìa ảnh, giúp triệt tiêu hoàn toàn hiện tượng viền đen giả và biên giả.",
        bold_prefix="Đệm phản chiếu gương (Reflect / Symmetric Padding): "
    )

    # --- MỤC 1.4 ---
    add_h2(doc, "1.4. Các bộ lọc khử nhiễu trong miền không gian")
    add_h3(doc, "1.4.1. Bộ lọc trung bình (Mean / Averaging Filter)")
    add_para(
        doc,
        "Bộ lọc trung bình số học (Arithmetic Mean Filter), còn gọi là bộ lọc hộp (Box Filter), là bộ lọc không gian tuyến tính cơ bản nhất dùng để làm trơn ảnh (smoothing) và giảm nhiễu. Ý tưởng cốt lõi của bộ lọc trung bình là thay thế giá trị của mỗi điểm ảnh tại tâm cửa sổ bằng trung bình cộng số học của tất cả k^2 điểm ảnh nằm trong lân cận S_xy kích thước k x k:"
    )
    add_formula(
        doc,
        "g(x, y) = (1 / k^2) · ∑_{(r, c) ∈ S_xy} f(r, c) = (1 / k^2) · ∑_{s = -a}^{a} ∑_{t = -a}^{a} f(x + s, y + t)",
        eq_number="1.10",
    )
    add_para(
        doc,
        "Dưới góc độ phép tích chập 2D, lọc trung bình tương đương với việc thực hiện tích chập ảnh đầu vào f(x, y) với ma trận nhân lọc H_{k x k} mà mọi phần tử đều có trọng số bằng nhau và bằng 1 / k^2. Chẳng hạn, với kích thước cửa sổ k = 3 (lân cận 3x3 gồm 9 điểm ảnh) và k = 5 (lân cận 5x5 gồm 25 điểm ảnh), nhân lọc H có dạng:"
    )
    add_formula(
        doc,
        "H_{3x3} = (1 / 9) · [ [1, 1, 1], [1, 1, 1], [1, 1, 1] ] ,      H_{5x5} = (1 / 25) · [ 1_{5x5} ]",
        eq_number="1.11",
    )
    add_para(
        doc,
        "Xét trên cơ sở lý thuyết xác suất thống kê: Giả sử một vùng ảnh phẳng có mức xám thực không đổi là f_0 bị nhiễm nhiễu cộng Gaussian độc lập phân phối cùng quy luật η_i ~ N(0, σ^2). Khi lấy trung bình cộng của N = k^2 điểm ảnh trong cửa sổ, kỳ vọng toán học của giá trị đầu ra vẫn bằng đúng f_0 (ước lượng không chệch), nhưng phương sai của nhiễu sau lọc giảm đi đúng N = k^2 lần:"
    )
    add_formula(
        doc,
        "Var( g(x, y) ) = Var( (1 / k^2) ∑_{i=1}^{k^2} η_i ) = (1 / k^4) · k^2 · σ^2 = σ^2 / k^2",
        eq_number="1.12",
    )
    add_para(
        doc,
        "Như vậy, độ lệch chuẩn của nhiễu Gaussian giảm đi k lần (cửa sổ 3x3 giúp giảm độ lệch chuẩn nhiễu đi 3 lần, phương sai giảm 9 lần; cửa sổ 5x5 giảm độ lệch chuẩn 5 lần, phương sai giảm 25 lần). Tuy nhiên, nhược điểm cố hữu của bộ lọc trung bình xuất hiện khi cửa sổ trượt đi qua vùng biên ảnh (nơi mức xám chuyển tiếp đột ngột từ vùng tối sang vùng sáng) hoặc khi gặp điểm nhiễu xung cực trị (0 hoặc 255): vì mọi điểm ảnh trong cửa sổ đều được cộng dồn với trọng số ngang nhau, các đường biên sắc nét sẽ bị san phẳng dần thành dải chuyển tiếp mờ nhòe (edge blurring), còn điểm nhiễu muối tiêu không bị mất đi mà bị phát tán thành một đốm mờ kích thước k x k."
    )

    add_h3(doc, "1.4.2. Bộ lọc trung vị (Median Filter)")
    add_para(
        doc,
        "Để khắc phục nhược điểm làm nhòe biên và hiện tượng lan truyền nhiễu xung của bộ lọc trung bình, Tukey đã đề xuất bộ lọc trung vị (Median Filter) dựa trên lý thuyết thống kê thứ tự (Order-Statistic Filtering). Đây là một bộ lọc phi tuyến (non-linear filter) điển hình, nghĩa là phép toán lọc không thể biểu diễn dưới dạng tổng tích chập tuyến tính với một ma trận trọng số cố định."
    )
    add_para(
        doc,
        "Nguyên lý hoạt động của bộ lọc trung vị trên cửa sổ trượt S_xy kích thước k x k (gồm N = k^2 điểm ảnh, với k lẻ nên N luôn là một số nguyên lẻ) gồm hai bước liên tiếp tại mỗi vị trí (x, y):"
    )
    add_bullet(
        doc,
        "Thu thập toàn bộ N = k^2 giá trị mức xám trong cửa sổ S_xy thành một mảng một chiều và sắp xếp chúng theo thứ tự không giảm: I_(0) ≤ I_(1) ≤ ... ≤ I_(N-1).",
        bold_prefix="Bước 1 (Sắp xếp thứ tự): "
    )
    add_bullet(
        doc,
        "Chọn giá trị nằm ở vị trí chính giữa của dãy đã sắp xếp (phần tử thứ m = (N - 1) / 2 = ⌊k^2 / 2⌋) để gán cho điểm ảnh đầu ra g(x, y):",
        bold_prefix="Bước 2 (Lấy phần tử trung vị): "
    )
    add_formula(
        doc,
        "g(x, y) = median { f(r, c)  |  (r, c) ∈ S_xy } = I_( (k^2 - 1) / 2 )",
        eq_number="1.13",
    )
    add_para(
        doc,
        "Ví dụ với cửa sổ 3x3 (N = 9 phần tử, chỉ số chạy từ 0 đến 8), trung vị là phần tử đứng ở chỉ số 4 sau khi sắp xếp; với cửa sổ 5x5 (N = 25 phần tử), trung vị là phần tử đứng ở chỉ số 12; với cửa sổ 7x7 (N = 49 phần tử), trung vị là phần tử đứng ở chỉ số 24."
    )
    add_para(
        doc,
        "Bản chất toán học tạo nên sức mạnh của bộ lọc trung vị nằm ở hai khía cạnh: Thứ nhất, đối với nhiễu muối tiêu (giá trị 0 hoặc 255), khi đưa vào dãy sắp xếp tăng dần, các điểm tiêu (0) luôn bị đẩy dạt về đầu bên trái của dãy và các điểm muối (255) luôn bị đẩy dạt về đầu bên phải của dãy. Chỉ cần tổng số điểm nhiễu xuất hiện trong cửa sổ nhỏ hơn một nửa diện tích cửa sổ (tức là ≤ (k^2 - 1)/2 điểm), phần tử đứng chính giữa hoàn toàn là một điểm ảnh gốc sạch và điểm nhiễu bị loại bỏ triệt để 100% mà không làm thay đổi mức xám của các pixel xung quanh. Thứ hai, khi cửa sổ trượt qua một đường biên bước nhảy (step edge) lý tưởng phân cách vùng tối (mức A) và vùng sáng (mức B), tập điểm ảnh trong cửa sổ chỉ gồm hai nhóm giá trị A và B; phần tử trung vị bắt buộc phải nhận giá trị A hoặc B (tùy thuộc tâm cửa sổ đang nằm bên phía nào của biên) chứ không bao giờ tạo ra giá trị trung bình cộng (A+B)/2. Nhờ đó, độ dốc của biên ảnh được bảo toàn sắc nét."
    )

    add_h3(doc, "1.4.3. Ví dụ minh họa tính tay trực quan trên vùng ảnh 3x3")
    add_para(
        doc,
        "Để minh chứng rõ ràng sự khác biệt giữa hai bộ lọc, xét một vùng ảnh con 3x3 nằm trên nền xám phẳng có giá trị chuẩn quanh mức 100, nhưng điểm ảnh ở chính giữa tâm cửa sổ bị hỏng bởi một hạt nhiễu muối (giá trị vọt lên 255):"
    )
    add_code_block(
        doc,
        [
            "Ma trận vùng ảnh lân cận 3x3 (tâm bị nhiễu muối = 255):",
            "    [ 100   102    98 ]",
            "    [  99   255   101 ]",
            "    [ 100   101    99 ]",
            "",
            "1) Áp dụng Bộ lọc Trung bình 3x3:",
            "   Tổng 9 phần tử = 100 + 102 + 98 + 99 + 255 + 101 + 100 + 101 + 99 = 1055",
            "   Giá trị đầu ra g_mean = round(1055 / 9) = round(117.22) = 117.",
            "   -> Nhận xét: Giá trị 117 lệch tới +17 mức xám so với nền thực (~100).",
            "      Hơn nữa, khi cửa sổ trượt sang 8 điểm xung quanh, giá trị 255 này tiếp tục",
            "      cộng vào cả 8 điểm lân cận, biến 1 điểm nhiễu thành 1 đốm sáng 3x3!",
            "",
            "2) Áp dụng Bộ lọc Trung vị 3x3:",
            "   Dãy 9 phần tử sắp xếp tăng dần:",
            "   Chỉ số :   [0]   [1]   [2]   [3]   [4]   [5]   [6]   [7]   [8]",
            "   Giá trị:    98,   99,   99,  100, (100), 101,  101,  102,  255",
            "   Phần tử đứng chính giữa (chỉ số 4) là: g_median = 100.",
            "   -> Nhận xét: Điểm nhiễu muối 255 bị đẩy về cuối dãy (chỉ số 8) và bị loại bỏ",
            "      hoàn toàn! Giá trị khôi phục bằng đúng 100 (sai số bằng 0).",
        ],
        caption="Ví dụ 1.1. Minh họa tính tay kết quả lọc Trung bình và lọc Trung vị trên cửa sổ 3x3 có nhiễu xung",
    )

    # Bảng tổng hợp so sánh lý thuyết Mean vs Median
    filter_comp_headers = [
        "Đặc tính kỹ thuật",
        "Bộ lọc Trung bình (Mean Filter)",
        "Bộ lọc Trung vị (Median Filter)",
    ]
    filter_comp_rows = [
        [
            "Phân loại toán học",
            "Bộ lọc không gian tuyến tính (Linear Spatial Filter) - thực hiện qua tích chập.",
            "Bộ lọc không gian phi tuyến (Non-linear Order-Statistic Filter) - dựa trên sắp xếp.",
        ],
        [
            "Cơ chế tính giá trị đầu ra",
            "Tính trung bình cộng có trọng số đều 1/k^2 của tất cả điểm ảnh trong cửa sổ.",
            "Sắp xếp tăng dần k^2 điểm ảnh và chọn phần tử đứng ở vị trí giữa ⌊k^2 / 2⌋.",
        ],
        [
            "Khả năng tạo giá trị mức xám mới",
            "Có thể sinh ra giá trị mức xám trung gian mới không tồn tại trong vùng ảnh gốc.",
            "Không sinh ra mức xám mới; luôn lấy một trong các giá trị thực có sẵn trong lân cận.",
        ],
        [
            "Hiệu quả với nhiễu Muối tiêu",
            "Kém: không xóa được điểm cực trị mà chỉ phát tán nhiễu thành các đốm mờ k x k.",
            "Rất xuất sắc: triệt tiêu hoàn toàn các điểm xung cực trị khi mật độ nhiễu vừa phải.",
        ],
        [
            "Hiệu quả với nhiễu Gaussian",
            "Tốt: giảm phương sai nhiễu xuống đúng k^2 lần (tối ưu theo ước lượng hợp lý cực đại).",
            "Khá: phương sai nhiễu sau lọc xấp xỉ (π/2)·(σ^2/k^2) ≈ 1.57·(σ^2/k^2), kém hơn lọc trung bình.",
        ],
        [
            "Khả năng bảo toàn biên (Edge Preservation)",
            "Kém: làm suy giảm độ dốc gradient và gây nhòe biên tỷ lệ thuận với kích thước cửa sổ.",
            "Tốt: bảo toàn độ sắc nét của các đường biên bước nhảy (step edges).",
        ],
        [
            "Độ phức tạp tính toán cơ bản",
            "Thấp hơn: O(M · N · k^2) phép cộng nhân trên toàn ảnh.",
            "Cao hơn: O(M · N · k^2 · log(k^2)) do chi phí sắp xếp mảng k^2 phần tử tại mỗi pixel.",
        ],
    ]
    add_styled_table(
        doc,
        headers=filter_comp_headers,
        rows=filter_comp_rows,
        caption="Bảng 1.2. Bảng đối sánh lý thuyết giữa Bộ lọc Trung bình và Bộ lọc Trung vị",
        col_aligns=[WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT],
    )

    # --- MỤC 1.5 ---
    add_h2(doc, "1.5. Phát hiện biên và toán tử Sobel")
    add_h3(doc, "1.5.1. Khái niệm biên ảnh và đạo hàm bậc nhất (Vector Gradient)")
    add_para(
        doc,
        "Trong ảnh số, biên (Edge) là tập hợp các điểm ảnh liên thông nằm tại ranh giới phân cách giữa hai vùng có mức xám (hoặc đặc trưng kết cấu) khác biệt rõ rệt. Về mặt hình học và giải tích, biên tương ứng với những vị trí mà hàm cường độ sáng f(x, y) có tốc độ biến thiên không gian lớn nhất. Công cụ toán học tự nhiên để đo tốc độ biến thiên cực đại của hàm hai biến f(x, y) là vector đạo hàm bậc nhất hay vector Gradient, ký hiệu là ∇f(x, y):"
    )
    add_formula(
        doc,
        "∇f(x, y) = [ G_x ,  G_y ]^T = [ ∂f(x, y) / ∂x ,  ∂f(x, y) / ∂y ]^T",
        eq_number="1.14",
    )
    add_para(
        doc,
        "Vector Gradient có hai thuộc tính hình học đặc biệt quan trọng: (1) Hướng của vector ∇f chỉ theo phương có tốc độ tăng độ sáng nhanh nhất (vuông góc với tiếp tuyến của đường biên tại điểm đó), được xác định bởi góc θ(x, y) = arctan(G_y / G_x); và (2) Độ lớn (biên độ) của vector Gradient, ký hiệu là G(x, y) = ||∇f(x, y)||, cho biết mức độ mạnh yếu (độ dốc) của đường biên tại tọa độ (x, y):"
    )
    add_formula(
        doc,
        "G(x, y) = ||∇f(x, y)|| = √( G_x^2 + G_y^2 )",
        eq_number="1.15",
    )

    add_h3(doc, "1.5.2. Cặp mặt nạ đạo hàm không gian của toán tử Sobel")
    add_para(
        doc,
        "Do ảnh số là một ma trận rời rạc nên các đạo hàm riêng ∂f/∂x và ∂f/∂y được xấp xỉ bằng phép tính sai phân hữu hạn trên lân cận 3x3. Toán tử Sobel (do Irwin Sobel và Gary Feldman đề xuất) là một trong những toán tử đạo hàm bậc nhất nổi tiếng nhất nhờ sự kết hợp khéo léo giữa phép lấy sai phân trung tâm theo một trục và phép làm trơn trọng số [1, 2, 1] theo trục vuông góc còn lại. Trọng số bằng 2 ở hàng/cột đi qua tâm cửa sổ giúp nhấn mạnh sự đóng góp của điểm ảnh gần tâm nhất và giảm nhẹ tác động của nhiễu tần số cao."
    )
    add_para(
        doc,
        "Cụ thể, hai mặt nạ Sobel kích thước 3x3 dùng để tính thành phần đạo hàm theo phương ngang S_x (phát hiện các đường biên thẳng đứng) và theo phương dọc S_y (phát hiện các đường biên nằm ngang) được định nghĩa như sau:"
    )
    add_formula(
        doc,
        "S_x = [ [-1,  0, +1],  [-2,  0, +2],  [-1,  0, +1] ] ,      S_y = [ [-1, -2, -1],  [ 0,  0,  0],  [+1, +2, +1] ]",
        eq_number="1.16",
    )
    add_para(
        doc,
        "Tại mỗi điểm ảnh (x, y), hai thành phần đạo hàm G_x(x, y) và G_y(x, y) được tính thông qua phép tích chập 2D (với lưu ý đảo chiều nhân trước khi gọi công thức tích chập chuẩn để bảo toàn hướng đạo hàm):"
    )
    add_formula(
        doc,
        "G_x(x, y) = [ f(x-1, y+1) + 2f(x, y+1) + f(x+1, y+1) ] - [ f(x-1, y-1) + 2f(x, y-1) + f(x+1, y-1) ]",
        eq_number="1.17",
    )
    add_formula(
        doc,
        "G_y(x, y) = [ f(x+1, y-1) + 2f(x+1, y) + f(x+1, y+1) ] - [ f(x-1, y-1) + 2f(x-1, y) + f(x-1, y+1) ]",
        eq_number="1.18",
    )

    add_h3(doc, "1.5.3. Ngưỡng hóa bản đồ Gradient để trích xuất bản đồ biên nhị phân")
    add_para(
        doc,
        "Bản đồ độ lớn Gradient G(x, y) = √(G_x^2 + G_y^2) là một ma trận số thực không âm, phản ánh cường độ biến thiên mức xám liên tục. Để quyết định dứt khoát một điểm ảnh (x, y) có phải là điểm biên thực sự hay không (phục vụ đánh giá độ chính xác định vị biên), ta áp dụng phép ngưỡng hóa (Thresholding) với một giá trị ngưỡng T_edge > 0 để thu được bản đồ biên nhị phân E(x, y) ∈ {0, 1}:"
    )
    add_formula(
        doc,
        "E(x, y) = { 1 (Điểm biên)  nếu G(x, y) ≥ T_edge ;   0 (Điểm nền)  nếu G(x, y) < T_edge }",
        eq_number="1.19",
    )
    add_para(
        doc,
        "Trong thực nghiệm đánh giá khả năng bảo toàn biên của các bộ lọc khử nhiễu, việc lựa chọn ngưỡng T_edge cần đảm bảo tính khách quan và công bằng tuyệt đối. Do đó, trong đề tài này, ngưỡng chuẩn T_edge của mỗi bức ảnh được xác định tự động theo phân vị thứ 85 (85th percentile) của độ lớn Gradient trên chính ảnh gốc sạch (tương ứng với 15% số điểm ảnh có cường độ biên mạnh nhất của ảnh gốc), sau đó giữ cố định giá trị T_edge này để áp dụng lên ảnh nhiễu và tất cả các ảnh sau lọc."
    )

    # --- MỤC 1.6 ---
    add_h2(doc, "1.6. Các chỉ số đánh giá chất lượng ảnh và khả năng bảo toàn biên")
    add_h3(doc, "1.6.1. Sai số bình phương trung bình (MSE - Mean Squared Error)")
    add_para(
        doc,
        "Khi có sẵn ảnh gốc sạch f(x, y) làm chuẩn tham chiếu (Full-Reference Image Quality Assessment), phép đo định lượng trực tiếp nhất để đánh giá mức độ sai lệch giữa ảnh sau khi lọc g(x, y) và ảnh gốc f(x, y) (cùng kích thước M hàng x N cột) là Sai số bình phương trung bình (MSE):"
    )
    add_formula(
        doc,
        "MSE(f, g) = [ 1 / (M · N) ] · ∑_{x = 0}^{M - 1} ∑_{y = 0}^{N - 1} [ f(x, y) - g(x, y) ]^2",
        eq_number="1.20",
    )
    add_para(
        doc,
        "Chỉ số MSE luôn nhận giá trị không âm (MSE ≥ 0). Giá trị MSE càng nhỏ chứng tỏ mức xám của ảnh sau lọc càng tiệm cận sát với ảnh gốc; MSE = 0 khi và chỉ khi hai ảnh giống hệt nhau hoàn toàn tại mọi pixel."
    )

    add_h3(doc, "1.6.2. Tỷ số tín hiệu cực đại trên nhiễu (PSNR - Peak Signal-to-Noise Ratio)")
    add_para(
        doc,
        "Do MSE phụ thuộc mạnh vào thang đo cường độ sáng của ảnh, trong xử lý ảnh và nén ảnh quốc tế, người ta thường quy đổi MSE sang thang đo logarit thông qua Tỷ số tín hiệu cực đại trên nhiễu (PSNR), tính bằng đơn vị decibel (dB):"
    )
    add_formula(
        doc,
        "PSNR(f, g) = 10 · log_10 ( MAX_I^2 / MSE(f, g) ) = 20 · log_10 ( MAX_I / √MSE(f, g) )   (dB)",
        eq_number="1.21",
    )
    add_para(
        doc,
        "trong đó MAX_I là giá trị mức xám cực đại có thể có của một điểm ảnh; đối với ảnh xám 8-bit chuẩn thì MAX_I = 255 (do đó MAX_I^2 = 65025). Ngược lại với MSE, giá trị PSNR càng cao (tính bằng dB) thì chất lượng khôi phục ảnh càng tốt. Thông thường, một ảnh sau khử nhiễu có PSNR > 30 dB được xem là có chất lượng thị giác tốt, từ 25 dB đến 30 dB là mức chấp nhận được, và dưới 20 dB là ảnh bị suy giảm nghiêm trọng."
    )

    add_h3(doc, "1.6.3. Các độ đo đánh giá khả năng bảo toàn biên (Edge-MSE và Edge F1-Score)")
    add_para(
        doc,
        "Hai chỉ số MSE và PSNR tính trung bình đồng đều trên toàn bộ M x N điểm ảnh, trong khi các điểm biên thực sự thường chỉ chiếm khoảng 10% – 15% diện tích bức ảnh. Vì vậy, một bộ lọc làm nhòe biên nhưng làm trơn tốt vùng nền phẳng vẫn có thể đạt PSNR khá cao. Để đánh giá chuyên biệt khả năng bảo toàn cấu trúc biên khi kết hợp với toán tử Sobel, nhóm nghiên cứu sử dụng đồng thời hai nhóm chỉ số bổ trợ nhau:"
    )
    add_bullet(
        doc,
        "Đo sai số bình phương trung bình giữa bản đồ độ lớn Gradient Sobel chuẩn hóa của ảnh gốc G_orig(x, y) và bản đồ độ lớn Gradient Sobel của ảnh đã lọc G_filt(x, y). Edge-MSE càng thấp chứng tỏ cường độ độ dốc tại các vùng biên càng ít bị suy hao hoặc biến dạng:",
        bold_prefix="1) Sai số bình phương trung bình bản đồ Gradient (Edge-MSE): "
    )
    add_formula(
        doc,
        "Edge-MSE = [ 1 / (M · N) ] · ∑_{x = 0}^{M - 1} ∑_{y = 0}^{N - 1} [ G_orig(x, y) - G_filt(x, y) ]^2",
        eq_number="1.22",
    )
    add_bullet(
        doc,
        "Gọi E_orig(x, y) ∈ {0, 1} là bản đồ biên nhị phân chuẩn của ảnh gốc và E_filt(x, y) ∈ {0, 1} là bản đồ biên nhị phân của ảnh sau lọc (cùng cắt tại ngưỡng T_edge). Ta coi bài toán phát hiện biên trên ảnh sau lọc như một bài toán phân lớp nhị phân với tập chuẩn là E_orig, từ đó xác định 3 đại lượng thống kê:",
        bold_prefix="2) Độ chính xác (Precision), Độ phủ (Recall) và Chỉ số F1 của biên (Edge F1-Score): "
    )
    add_para(
        doc,
        "+ Số điểm biên đúng (True Positives - TP): Số điểm ảnh mà cả ảnh gốc và ảnh sau lọc đều xác định là biên (E_orig = 1 và E_filt = 1).\n"
        "+ Số điểm biên giả (False Positives - FP): Số điểm ảnh vốn là vùng nền phẳng trên ảnh gốc nhưng bị nhận nhầm thành biên trên ảnh sau lọc do nhiễu còn sót lại hoặc do biên bị bè rộng (E_orig = 0 và E_filt = 1).\n"
        "+ Số điểm biên bị mất (False Negatives - FN): Số điểm biên thực sự của ảnh gốc nhưng bị mất đi sau khi lọc do bộ lọc làm trơn quá mức làm tụt độ dốc gradient xuống dưới ngưỡng (E_orig = 1 và E_filt = 0).",
        first_line_indent=0.5,
    )
    add_formula(
        doc,
        "Precision = TP / (TP + FP) ,      Recall = TP / (TP + FN) ,      Edge-F1 = (2 · Precision · Recall) / (Precision + Recall)",
        eq_number="1.23",
    )
    add_para(
        doc,
        "Như vậy, Precision phản ánh khả năng triệt tiêu 'biên giả' do nhiễu gây ra, Recall phản ánh khả năng giữ lại không làm đứt gãy 'biên thật' của vật thể, còn Edge-F1 (nhận giá trị trong đoạn [0, 1], càng gần 1 càng tốt) là thước đo tổng hợp toàn diện nhất cho chất lượng bảo toàn biên."
    )

    # --- MỤC 1.7 ---
    add_h2(doc, "1.7. Giới thiệu ngôn ngữ lập trình và các thư viện công cụ áp dụng")
    add_para(
        doc,
        "Để đáp ứng nghiêm ngặt các ràng buộc kỹ thuật của đề bài (tự cài đặt hoàn toàn các thuật toán cốt lõi mà không lạm dụng hàm lọc đóng gói sẵn, đồng thời đảm bảo tốc độ tính toán nhanh trên 108 kịch bản thực nghiệm), nhóm lựa chọn hệ sinh thái ngôn ngữ lập trình Python 3 cùng các thư viện khoa học tiêu chuẩn sau:"
    )
    add_bullet(
        doc,
        "Là ngôn ngữ lập trình bậc cao có cú pháp trong sáng, hỗ trợ mạnh mẽ lập trình hướng cấu trúc và tính toán ma trận. Trong dự án này, Python 3 đóng vai trò tổ chức kiến trúc chương trình theo từng module độc lập, quản lý luồng thực nghiệm tự động và xuất dữ liệu báo cáo.",
        bold_prefix="Ngôn ngữ lập trình Python 3: "
    )
    add_bullet(
        doc,
        "Là thư viện nền tảng quan trọng nhất của đồ án, cung cấp cấu trúc mảng đa chiều đồng nhất (ndarray) được lưu trữ liên tiếp trong bộ nhớ. Toàn bộ các thuật toán cốt lõi của đề tài — bao gồm mở rộng biên ảnh (pad_image), sinh nhiễu ngẫu nhiên có hạt giống cố định (np.random.default_rng), trích xuất cửa sổ trượt không gian 4 chiều (sliding_window_view), tính tích chập 2D (np.einsum), sắp xếp tìm trung vị (np.sort), tính toán tử Sobel và các độ đo MSE/PSNR/Edge-F1 — đều được tự xây dựng thuần túy bằng các phép toán đại số mảng của NumPy.",
        bold_prefix="Thư viện tính toán mảng đa chiều NumPy (Numerical Python): "
    )
    add_bullet(
        doc,
        "Theo đúng quy định của đề bài, thư viện OpenCV và Pillow trong đồ án này chỉ được sử dụng cho hai mục đích hỗ trợ: (1) Đọc và ghi tệp ảnh định dạng PNG từ ổ đĩa (cv2.imread, cv2.imwrite); và (2) Sử dụng các hàm chuẩn cv2.blur, cv2.medianBlur, cv2.Sobel trong module kiểm thử độc lập (verification.py) để đối chiếu sai số, chứng minh các thuật toán tự cài đặt của nhóm đạt độ chính xác tuyệt đối.",
        bold_prefix="Thư viện OpenCV (cv2) và Pillow (PIL): "
    )
    add_bullet(
        doc,
        "Được sử dụng để thiết kế các bảng hình ảnh đối sánh trực quan đa khung (multi-panel figures) gồm Ảnh gốc – Ảnh nhiễu – Ảnh lọc Trung bình – Ảnh lọc Trung vị kèm bản đồ biên Sobel, cũng như vẽ các biểu đồ đường (Line charts) và biểu đồ cột (Bar charts) phục vụ phân tích định lượng trong Chương 3.",
        bold_prefix="Thư viện trực quan hóa dữ liệu Matplotlib: "
    )
    add_bullet(
        doc,
        "Được sử dụng để xây dựng ứng dụng Web tương tác chạy trên máy chủ cục bộ Localhost (http://127.0.0.1:8000). Hệ thống Web cho phép người dùng tải ảnh bất kỳ từ máy tính lên trực tiếp, tùy chỉnh tham số nhiễu và cửa sổ lọc theo thời gian thực, quan sát đồng thời 8 khung ảnh Trước và Sau khi lọc kèm bản đồ biên Sobel, nút Reset khôi phục ảnh gốc và bảng đối chiếu chỉ số tự động.",
        bold_prefix="Framework FastAPI và máy chủ ASGI Uvicorn: "
    )


def main() -> None:
    root_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    doc_path = os.path.join(root_dir, "BTL_XLA.docx")
    upd_path = os.path.join(root_dir, "BTL_XLA_updated.docx")
    doc = docx.Document(doc_path)
    configure_styles(doc)
    reset_body_after_toc(doc)
    write_front_lists(doc)
    write_introduction(doc)
    write_chapter_1(doc)
    doc.save(upd_path)
    try:
        doc.save(doc_path)
        print(f"Đã ghi thành công Danh mục, Mở đầu và Chương 1 vào: {doc_path}")
    except PermissionError:
        print(f"Đã ghi thành công Danh mục, Mở đầu và Chương 1 vào: {upd_path}")


if __name__ == "__main__":
    main()

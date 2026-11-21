# -*- coding: utf-8 -*-
"""
Script: scripts/write_step4_ch2.py
Mô tả: Ghi trực tiếp CHƯƠNG 2. XÂY DỰNG CHƯƠNG TRÌNH (ỨNG DỤNG) vào file BTL_XLA.docx.
       Bao gồm: mô tả bài toán, đặc tả các phân hệ chức năng, sơ đồ quy trình tổng thể,
       giải pháp kỹ thuật mở rộng biên ảnh (Padding), thuật toán và mã giả chi tiết
       của từng bộ lọc / toán tử Sobel / chỉ số đánh giá, phân tích độ phức tạp tính toán
       và cấu trúc tổ chức mã nguồn.
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
    add_figure,
)


def remove_existing_chapter2_onwards(doc: docx.Document) -> None:
    """Nếu trong file đã có Chương 2 trở đi thì xóa để ghi mới không bị trùng lặp."""
    body = doc.element.body
    children = list(body)
    ch2_idx = None
    for idx, child in enumerate(children):
        if child.tag.endswith("p"):
            txt = "".join(child.itertext()).strip()
            if txt.startswith("CHƯƠNG 2."):
                ch2_idx = idx
                break
    if ch2_idx is not None:
        for child in children[ch2_idx:]:
            if not child.tag.endswith("sectPr"):
                body.remove(child)


def write_chapter_2(doc: docx.Document, fig_dir: str) -> None:
    """Ghi toàn bộ nội dung Chương 2 vào tài liệu Word."""
    add_h1(doc, "CHƯƠNG 2. XÂY DỰNG CHƯƠNG TRÌNH", page_break_before=True)

    # --- MỤC 2.1 ---
    add_h2(doc, "2.1. Mô tả chi tiết bài toán và các chức năng của ứng dụng cần xây dựng")
    add_h3(doc, "2.1.1. Phát biểu bài toán và các ràng buộc kỹ thuật")
    add_para(
        doc,
        "Bài toán đặt ra cho Nhóm 10 trong Đề tài 7 là thiết kế và xây dựng một hệ thống phần mềm thực nghiệm hoàn chỉnh bằng ngôn ngữ Python 3 nhằm thực hiện khử nhiễu ảnh đa mức xám (8-bit Grayscale) trong miền không gian bằng hai phương pháp: Bộ lọc trung bình (Mean Filter) và Bộ lọc trung vị (Median Filter), đồng thời kết hợp toán tử đạo hàm bậc nhất Sobel để đánh giá định lượng khả năng bảo toàn cấu trúc biên ảnh sau khi lọc."
    )
    add_para(
        doc,
        "Để đảm bảo tính học thuật nghiêm túc và đánh giá đúng năng lực nắm bắt bản chất thuật toán xử lý ảnh của sinh viên, hệ thống phần mềm được xây dựng dưới các ràng buộc kỹ thuật bắt buộc sau đây:"
    )
    add_bullet(
        doc,
        "Toàn bộ các thuật toán cốt lõi của đề tài — bao gồm: hàm mở rộng biên ảnh 2 chiều (Padding), hàm sinh nhiễu muối tiêu và nhiễu cộng Gaussian, phép tích chập 2 chiều (2D Convolution), bộ lọc trung bình, bộ lọc trung vị, toán tử phát hiện biên Sobel (tính G_x, G_y, độ lớn Gradient và ngưỡng hóa), cùng các độ đo MSE, PSNR, Edge-MSE, Edge-F1 — đều phải được tự cài đặt từ đầu (from scratch) chỉ sử dụng cấu trúc mảng của thư viện NumPy.",
        bold_prefix="Ràng buộc 1 (Tự cài đặt thuật toán cốt lõi): "
    )
    add_bullet(
        doc,
        "Tuyệt đối không sử dụng các hàm lọc đóng gói sẵn của thư viện OpenCV hoặc SciPy (như cv2.blur, cv2.boxFilter, cv2.medianBlur, cv2.Sobel, scipy.ndimage) để tạo ra kết quả thực nghiệm chính. Thư viện OpenCV/Pillow chỉ được phép dùng để đọc/ghi tệp ảnh trên đĩa và dùng trong một module kiểm thử độc lập (verification.py) nhằm đối chiếu sai số, chứng minh tính đúng đắn của các hàm tự viết.",
        bold_prefix="Ràng buộc 2 (Giới hạn phạm vi sử dụng thư viện): "
    )
    add_bullet(
        doc,
        "Quá trình sinh nhiễu ngẫu nhiên (cả nhiễu muối tiêu lẫn nhiễu Gaussian) phải được cố định hạt giống bộ sinh số giả ngẫu nhiên (Random Seed = 42). Điều này đảm bảo mỗi khi chạy lại chương trình trên bất kỳ máy tính nào, ma trận nhiễu sinh ra đều giống hệt nhau 100%, giúp việc đối sánh giữa các bộ lọc và các kích thước cửa sổ hoàn toàn công bằng và có tính tái lập khoa học.",
        bold_prefix="Ràng buộc 3 (Tính tái lập thực nghiệm - Reproducibility): "
    )

    add_h3(doc, "2.1.2. Đặc tả các phân hệ chức năng của chương trình ứng dụng")
    add_para(
        doc,
        "Dựa trên yêu cầu của bài toán, chương trình ứng dụng được thiết kế theo kiến trúc hướng module (Modular Architecture) gồm 7 phân hệ chức năng liên kết chặt chẽ:"
    )

    func_headers = ["TT", "Tên phân hệ chức năng", "Module đảm nhiệm", "Mô tả chi tiết đầu vào, xử lý và đầu ra"]
    func_rows = [
        [
            "1",
            "Phân hệ Quản lý và Chuẩn hóa Bộ dữ liệu",
            "src/dataset.py",
            "- Đầu vào: Các ảnh chuẩn kinh điển trong xử lý ảnh (cameraman, coins, astronaut, text, checkerboard, brick).\n"
            "- Xử lý: Chuyển đổi về ảnh đa mức xám 8-bit (uint8), chuẩn hóa kích thước về đúng 256 x 256 pixel và lưu trữ tại thư mục data/original/.\n"
            "- Đầu ra: Từ điển (Dictionary) chứa 6 ma trận ảnh xám gốc chuẩn.",
        ],
        [
            "2",
            "Phân hệ Mở rộng biên ảnh 2 chiều (Boundary Padding)",
            "src/padding.py",
            "- Đầu vào: Ma trận ảnh 2D kích thước (H, W), số dòng/cột cần đệm (pad_h, pad_w) và chế độ đệm ('zero', 'replicate', 'reflect').\n"
            "- Xử lý: Khởi tạo ma trận mở rộng (H + 2*pad_h, W + 2*pad_w) và điền giá trị viền theo đúng quy tắc toán học.\n"
            "- Đầu ra: Ma trận ảnh đã mở rộng biên, phục vụ tích chập và lọc không gian.",
        ],
        [
            "3",
            "Phân hệ Giả lập Nhiễu có kiểm soát",
            "src/noise.py",
            "- Đầu vào: Ảnh xám gốc, tham số mức nhiễu và hạt giống ngẫu nhiên seed = 42.\n"
            "- Xử lý: Tự cài đặt hàm sinh nhiễu Muối tiêu ở 3 mức mật độ (1%, 5%, 10%) và nhiễu cộng Gaussian ở 3 mức độ lệch chuẩn (σ = 10, 25, 50).\n"
            "- Đầu ra: Các ma trận ảnh nhiễu kiểu uint8 trong dải [0, 255].",
        ],
        [
            "4",
            "Phân hệ Khử nhiễu trong miền không gian",
            "src/filters.py",
            "- Đầu vào: Ảnh nhiễu, loại bộ lọc (Trung bình / Trung vị) và kích thước cửa sổ k ∈ {3, 5, 7}.\n"
            "- Xử lý: Tự cài đặt phép tích chập 2D (convolve2d), bộ lọc trung bình (mean_filter) và bộ lọc trung vị (median_filter).\n"
            "- Đầu ra: Ảnh đã khử nhiễu có cùng kích thước (H, W) và kiểu uint8.",
        ],
        [
            "5",
            "Phân hệ Phát hiện biên Sobel",
            "src/sobel.py",
            "- Đầu vào: Ảnh gốc, ảnh nhiễu hoặc ảnh đã qua bộ lọc khử nhiễu; ngưỡng cắt biên T_edge.\n"
            "- Xử lý: Tính đạo hàm ngang G_x, dọc G_y bằng tích chập tự viết với mặt nạ Sobel 3x3; tính độ lớn Gradient G = √(G_x^2 + G_y^2) và bản đồ biên nhị phân.\n"
            "- Đầu ra: Ma trận độ lớn Gradient thực, ảnh biên mức xám [0, 255] và bản đồ biên nhị phân {0, 1}.",
        ],
        [
            "6",
            "Phân hệ Đo lường Chất lượng & Kiểm chứng",
            "src/metrics.py &\nsrc/verification.py",
            "- Đầu vào: Cặp ảnh (Ảnh gốc, Ảnh sau lọc) và cặp bản đồ biên Sobel (Biên gốc, Biên sau lọc).\n"
            "- Xử lý: Tính MSE, PSNR (dB), Edge-MSE, Precision, Recall, Edge-F1; đo thời gian thực thi (ms) và đối chiếu sai số với OpenCV.\n"
            "- Đầu ra: Các chỉ số đánh giá định lượng chính xác.",
        ],
        [
            "7",
            "Phân hệ Điều phối Thực nghiệm & Trực quan hóa",
            "run_experiments.py &\nsrc/visualization.py",
            "- Đầu vào: Toàn bộ 108 kịch bản tổ hợp thực nghiệm.\n"
            "- Xử lý: Tự động chạy vòng lặp thực nghiệm, xuất 3 bảng CSV vào results/tables/ và vẽ hệ thống hình ảnh đối sánh, biểu đồ vào results/figures/.",
        ],
        [
            "8",
            "Phân hệ Giao diện Web Tương tác trên Localhost",
            "web_app.py\n(FastAPI + HTML5)",
            "- Đầu vào: Ảnh chọn từ bộ 6 ảnh mẫu hoặc ảnh người dùng tải trực tiếp từ máy tính (Upload); các tham số thanh trượt nhiễu, kích thước cửa sổ và chế độ Padding.\n"
            "- Xử lý: Gọi trực tiếp các hàm tự cài đặt trong src/, mã hóa ảnh kết quả sang Base64 PNG thời gian thực.\n"
            "- Đầu ra: Giao diện Web tại http://127.0.0.1:8000 hiển thị 8 khung ảnh Trước/Sau lọc & Biên Sobel, nút Reset ảnh gốc và bảng đối sánh MSE, PSNR, Edge-F1.",
        ],
    ]
    add_styled_table(
        doc,
        headers=func_headers,
        rows=func_rows,
        caption="Bảng 2.1. Đặc tả chi tiết các phân hệ chức năng của chương trình ứng dụng",
        col_aligns=[
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.LEFT,
        ],
    )

    add_h3(doc, "2.1.3. Quy trình xử lý tổng thể (Pipeline) của chương trình")
    add_para(
        doc,
        "Luồng xử lý dữ liệu từ đầu đến cuối của chương trình được tự động hóa hoàn toàn thông qua kịch bản điều phối trung tâm. Với mỗi ảnh chuẩn trong bộ dữ liệu 6 ảnh, trước hết chương trình trích xuất bản đồ biên Sobel trên ảnh gốc sạch và xác định ngưỡng biên tham chiếu T_edge (theo phân vị thứ 85 của độ lớn Gradient ảnh gốc). Tiếp đó, ảnh gốc lần lượt được đưa qua bộ tạo nhiễu với 6 cấu hình nhiễu (Muối tiêu 1%, 5%, 10% và Gaussian σ = 10, 25, 50). Mỗi ảnh nhiễu lại được khử nhiễu song song bằng Bộ lọc Trung bình và Bộ lọc Trung vị ở 3 kích thước cửa sổ (3x3, 5x5, 7x7). Cuối cùng, ảnh sau lọc được đưa qua toán tử Sobel (áp dụng cùng ngưỡng chuẩn T_edge) để đo lường đồng thời hai nhóm chỉ số: chất lượng khôi phục mức xám (MSE, PSNR) và chất lượng bảo toàn biên (Edge-MSE, Edge-Precision, Edge-Recall, Edge-F1) cùng thời gian thực thi (ms)."
    )

    pipeline_img = os.path.join(fig_dir, "ch2_pipeline_diagram.png")
    if os.path.exists(pipeline_img):
        add_figure(
            doc,
            pipeline_img,
            caption="Hình 2.1. Sơ đồ khối quy trình xử lý tổng thể (Pipeline) của chương trình thực nghiệm",
            width_inches=6.0,
        )

    # --- MỤC 2.2 ---
    add_h2(doc, "2.2. Thiết kế chi tiết phương pháp, thuật toán và giải pháp kỹ thuật")
    add_h3(doc, "2.2.1. Giải pháp kỹ thuật xử lý biên ảnh (Boundary Padding tự cài đặt)")
    add_para(
        doc,
        "Khi thực hiện phép lọc không gian với cửa sổ trượt vuông kích thước lẻ k x k (trong đó k ∈ {3, 5, 7}), bán kính mở rộng biên ở mỗi phía (trên, dưới, trái, phải) được xác định bởi công thức p = ⌊k / 2⌋ = (k - 1) / 2. Cụ thể: với cửa sổ 3x3 thì p = 1 pixel; với cửa sổ 5x5 thì p = 2 pixel; với cửa sổ 7x7 thì p = 3 pixel. Từ một ảnh đầu vào I có kích thước H hàng x W cột, hàm tự cài đặt pad_image(I, p, p, mode) sẽ cấp phát một ma trận mới P có kích thước (H + 2p) x (W + 2p) và chép nguyên vẹn ảnh gốc I vào vùng lõi trung tâm P[p : p+H, p : p+W]."
    )
    add_para(
        doc,
        "Để xử lý phần dải viền xung quanh có độ dày p pixel, nhóm đã cài đặt cả ba chế độ: 'zero' (đệm số 0), 'replicate' (lặp lại hàng/cột biên ngoài cùng) và 'reflect' (phản chiếu gương qua biên). Trong toàn bộ các phép lọc thực nghiệm chính và phép tính đạo hàm Sobel, nhóm lựa chọn chế độ phản chiếu gương ('reflect') làm mặc định. Lý do kỹ thuật là: nếu dùng 'zero', tại viền ảnh sẽ xuất hiện sự sụt giảm cường độ sáng đột ngột từ mức xám nền (ví dụ ~180) xuống 0, khiến toán tử Sobel tính ra đạo hàm cực lớn dọc theo 4 cạnh khung ảnh và làm sai lệch chỉ số Edge-MSE. Ngược lại, chế độ 'reflect' lấy đối xứng gương P[p - 1 - i, :] = I[1 + i, :], đảm bảo đạo hàm bậc nhất vuông góc với viền ảnh bằng 0 và giữ cho vùng biên ảnh tự nhiên nhất."
    )

    pad_img = os.path.join(fig_dir, "ch2_padding_modes.png")
    if os.path.exists(pad_img):
        add_figure(
            doc,
            pad_img,
            caption="Hình 2.2. Minh họa trực quan 3 chế độ mở rộng biên (Zero, Replicate, Reflect) với bán kính đệm p = 1",
            width_inches=5.9,
        )

    add_code_block(
        doc,
        [
            "Thuật toán 2.1: Mở rộng biên ảnh 2 chiều (pad_image)",
            "-------------------------------------------------------------------------",
            "Đầu vào : Ma trận ảnh xám I kích thước H x W, bán kính đệm p, chế độ mode.",
            "Đầu ra  : Ma trận ảnh đã mở rộng biên P kích thước (H + 2p) x (W + 2p).",
            "Bước 1  : Khởi tạo ma trận P kích thước (H + 2p) x (W + 2p) gồm toàn số 0.",
            "Bước 2  : Gán vùng tâm: P[p : p + H, p : p + W] <- I[0 : H, 0 : W].",
            "Bước 3  : Nếu mode == 'zero' thì trả về P.",
            "Bước 4  : Nếu mode == 'replicate' thì:",
            "            Với mỗi i từ 0 đến p - 1:",
            "                P[i, p : p + W] <- I[0, :]            (Đệm biên trên)",
            "                P[p + H + i, p : p + W] <- I[H - 1, :] (Đệm biên dưới)",
            "            Với mỗi j từ 0 đến p - 1:",
            "                P[:, j] <- P[:, p]                    (Đệm biên trái & góc)",
            "                P[:, p + W + j] <- P[:, p + W - 1]    (Đệm biên phải & góc)",
            "Bước 5  : Nếu mode == 'reflect' thì:",
            "            Với mỗi i từ 0 đến p - 1:",
            "                src_top <- min(p - i, H - 1); src_bot <- max(H - 2 - i, 0)",
            "                P[i, p : p + W] <- I[src_top, :]",
            "                P[p + H + i, p : p + W] <- I[src_bot, :]",
            "            Với mỗi j từ 0 đến p - 1:",
            "                src_left <- min(p - j, W - 1); src_right <- max(W - 2 - j, 0)",
            "                P[:, j] <- P[:, p + src_left]",
            "                P[:, p + W + j] <- P[:, p + src_right]",
            "Bước 6  : Trả về ma trận P.",
        ],
        caption="Mã giả 2.1. Thuật toán mở rộng biên ảnh 2 chiều tự cài đặt trong module src/padding.py",
    )

    add_h3(doc, "2.2.2. Thuật toán và mã giả sinh nhiễu Muối tiêu và nhiễu Gaussian")
    add_para(
        doc,
        "Cả hai thuật toán giả lập nhiễu trong module src/noise.py đều khởi tạo bộ sinh số ngẫu nhiên độc lập rng = np.random.default_rng(seed=42) ngay đầu hàm. Đối với nhiễu muối tiêu mật độ d (với d ∈ {0.01, 0.05, 0.10}) và tỷ lệ cân bằng muối/tiêu bằng 0.5, thuật toán sinh một ma trận số thực ngẫu nhiên R cùng kích thước H x W tuân theo phân phối đều U(0, 1). Những tọa độ (x, y) có R(x, y) < d/2 được gán mức xám 255 (hạt muối trắng); những tọa độ có d/2 ≤ R(x, y) < d được gán mức xám 0 (hạt tiêu đen); các tọa độ còn lại giữ nguyên giá trị của ảnh gốc."
    )
    add_para(
        doc,
        "Đối với nhiễu cộng Gaussian có độ lệch chuẩn σ ∈ {10, 25, 50} và kỳ vọng μ = 0, thuật toán sinh ma trận nhiễu N ~ Normal(0, σ^2) kích thước H x W. Để tránh lỗi tràn số nguyên không dấu 8-bit (uint8 overflow/underflow, ví dụ 250 + 20 bị quay vòng thành 14 nếu tính trực tiếp trên uint8), ảnh gốc bắt buộc phải được ép kiểu sang số thực dấu phẩy động 64-bit (float64) trước khi cộng ma trận nhiễu N, sau đó làm tròn tới số nguyên gần nhất (np.rint), cắt ngưỡng vào đoạn [0, 255] bằng hàm np.clip rồi mới chuyển ngược về kiểu uint8."
    )

    add_code_block(
        doc,
        [
            "Thuật toán 2.2: Sinh nhiễu Muối tiêu và nhiễu Gaussian có kiểm soát",
            "-------------------------------------------------------------------------",
            "A. HÀM ADD_SALT_AND_PEPPER_NOISE(I, density d, seed = 42):",
            "   1. Khởi tạo bộ sinh ngẫu nhiên rng <- default_rng(seed).",
            "   2. Sao chép ảnh đầu ra: G <- copy(I).",
            "   3. Sinh ma trận xác suất đều: R <- rng.random(shape=(H, W)) ∈ [0, 1).",
            "   4. Tính ngưỡng xác suất: p_salt <- d / 2 ;  p_pepper <- d / 2.",
            "   5. Với mọi tọa độ (x, y):",
            "        Nếu R[x, y] < p_salt thì: G[x, y] <- 255  (Điểm muối)",
            "        Ngược lại nếu R[x, y] < p_salt + p_pepper thì: G[x, y] <- 0 (Điểm tiêu)",
            "   6. Trả về ảnh nhiễu G (kiểu uint8).",
            "",
            "B. HÀM ADD_GAUSSIAN_NOISE(I, sigma σ, mean μ = 0, seed = 42):",
            "   1. Khởi tạo bộ sinh ngẫu nhiên rng <- default_rng(seed).",
            "   2. Sinh ma trận nhiễu chuẩn: Eta <- rng.normal(loc=μ, scale=σ, size=(H, W)).",
            "   3. Cộng nhiễu trên miền số thực: G_float <- float64(I) + Eta.",
            "   4. Làm tròn và cắt ngưỡng: G <- uint8( clip( round(G_float), 0, 255 ) ).",
            "   5. Trả về ảnh nhiễu G (kiểu uint8).",
        ],
        caption="Mã giả 2.2. Thuật toán sinh nhiễu Muối tiêu và nhiễu Gaussian trong module src/noise.py",
    )

    add_h3(doc, "2.2.3. Thuật toán phép tích chập 2D và Bộ lọc trung bình (Mean Filter)")
    add_para(
        doc,
        "Trong module src/filters.py, phép tích chập 2 chiều convolve2d(image, kernel, pad_mode) được thiết kế tổng quát cho mọi nhân lọc kích thước lẻ k_h x k_w. Đúng theo định nghĩa toán học của tích chập, bước đầu tiên là lật ma trận nhân lọc 180 độ qua cả trục dọc và trục ngang. Tiếp theo, ảnh đầu vào được mở rộng biên với bán kính p_h = ⌊k_h / 2⌋ và p_w = ⌊k_w / 2⌋ bằng hàm pad_image."
    )
    add_para(
        doc,
        "Về mặt kỹ thuật lập trình trong Python: nếu cài đặt phép tích chập bằng 4 vòng lặp for lồng nhau thuần Python (2 vòng lặp duyệt qua H x W điểm ảnh và 2 vòng lặp duyệt qua k x k phần tử của cửa sổ), tốc độ thực thi sẽ rất chậm do chi phí thông dịch vòng lặp của Python. Để tối ưu hóa tốc độ mà vẫn giữ nguyên bản chất tự cài đặt thuật toán trên mảng NumPy, nhóm sử dụng cấu trúc cửa sổ trượt bộ nhớ (memory-view) thông qua hàm sliding_window_view(padded, (k_h, k_w)) để ánh xạ ảnh đệm 2 chiều thành một tensor 4 chiều kích thước (H, W, k_h, k_w) chứa toàn bộ các vùng lân cận, sau đó thực hiện tổng tích vô hướng rút gọn chỉ số Einstein (np.einsum('ijkl,kl->ij', windows, flipped_kernel)). Kỹ thuật này hoàn toàn tương đương toán học với 4 vòng lặp tường minh nhưng tận dụng được tập lệnh vector hóa cấp thấp của CPU."
    )
    add_para(
        doc,
        "Trên cơ sở hàm convolve2d đã xây dựng, Bộ lọc trung bình mean_filter(image, kernel_size=k) được cài đặt ngắn gọn và chuẩn xác bằng cách khởi tạo ma trận nhân lọc đều H_{k x k} gồm toàn các phần tử có giá trị bằng 1 / k^2, gọi hàm convolve2d, rồi làm tròn và giới hạn kết quả về kiểu số nguyên uint8."
    )

    add_code_block(
        doc,
        [
            "Thuật toán 2.3: Tích chập 2D tổng quát và Bộ lọc Trung bình (Mean Filter)",
            "-------------------------------------------------------------------------",
            "A. HÀM CONVOLVE2D(I, Kernel K kích thước k x k, pad_mode = 'reflect'):",
            "   1. Tính bán kính đệm: p <- (k - 1) // 2.",
            "   2. Lật nhân lọc 180 độ: K_flip[s, t] <- K[k - 1 - s, k - 1 - t].",
            "   3. Mở rộng biên ảnh: P <- pad_image(float64(I), p, p, mode=pad_mode).",
            "   4. Khởi tạo ma trận đầu ra O kích thước H x W bằng 0.",
            "   5. Với mọi hàng x từ 0 đến H - 1:",
            "        Với mọi cột y từ 0 đến W - 1:",
            "            Trích xuất vùng lân cận: Window <- P[x : x + k, y : y + k]",
            "            O[x, y] <- ∑_{u=0}^{k-1} ∑_{v=0}^{k-1} Window[u, v] * K_flip[u, v]",
            "   6. Trả về ma trận số thực O.",
            "",
            "B. HÀM MEAN_FILTER(I, kernel_size k, pad_mode = 'reflect'):",
            "   1. Tạo nhân lọc trung bình: K_mean <- Ma trận (k x k) với mọi phần tử = 1.0 / (k * k).",
            "   2. Tính tích chập: O_float <- CONVOLVE2D(I, K_mean, pad_mode).",
            "   3. Chuẩn hóa kiểu dữ liệu: I_filtered <- uint8( clip( round(O_float), 0, 255 ) ).",
            "   4. Trả về ảnh đã lọc trung bình I_filtered.",
        ],
        caption="Mã giả 2.3. Thuật toán tích chập 2 chiều và Bộ lọc trung bình trong module src/filters.py",
    )

    add_h3(doc, "2.2.4. Thuật toán và mã giả Bộ lọc trung vị (Median Filter)")
    add_para(
        doc,
        "Đối với Bộ lọc trung vị median_filter(image, kernel_size=k, pad_mode='reflect'), vì đây là bộ lọc phi tuyến không thể dùng tích chập, quy trình xử lý tại mỗi điểm ảnh (x, y) đòi hỏi phải trích xuất cửa sổ lân cận kích thước k x k, duỗi phẳng (flatten) cửa sổ 2 chiều đó thành một mảng 1 chiều gồm N = k^2 phần tử, thực hiện thuật toán sắp xếp tăng dần (Sorting) trên mảng N phần tử này, và cuối cùng truy xuất phần tử nằm tại chỉ số trung tâm median_idx = (k^2 - 1) // 2."
    )
    add_para(
        doc,
        "Trong cài đặt thực tế tại src/filters.py, sau khi đệm biên ảnh bằng pad_image(image, p, p, mode=pad_mode), chương trình tạo khung nhìn cửa sổ trượt 4 chiều kích thước (H, W, k, k), biến đổi hình dạng (reshape) về mảng 3 chiều kích thước (H, W, k^2), sau đó áp dụng phép sắp xếp np.sort(flat_windows, axis=-1) dọc theo chiều cuối cùng (chiều chứa k^2 phần tử lân cận) và lấy lát cắt tại chỉ số median_idx. Cách cài đặt này vừa thể hiện minh bạch bản chất sắp xếp tìm trung vị, vừa tránh được việc gọi hàm thư viện có sẵn."
    )

    add_code_block(
        doc,
        [
            "Thuật toán 2.4: Bộ lọc Trung vị (Median Filter)",
            "-------------------------------------------------------------------------",
            "HÀM MEDIAN_FILTER(I, kernel_size k, pad_mode = 'reflect'):",
            "   1. Tính bán kính đệm: p <- (k - 1) // 2.",
            "   2. Tính số phần tử trong cửa sổ: N <- k * k ; Chỉ số trung vị: m_idx <- (N - 1) // 2.",
            "   3. Mở rộng biên ảnh: P <- pad_image(I, p, p, mode=pad_mode).",
            "   4. Khởi tạo ảnh đầu ra I_filtered kích thước H x W (kiểu uint8).",
            "   5. Với mọi hàng x từ 0 đến H - 1:",
            "        Với mọi cột y từ 0 đến W - 1:",
            "            a) Lấy cửa sổ lân cận k x k: W_xy <- P[x : x + k, y : y + k].",
            "            b) Trải phẳng W_xy thành mảng 1 chiều A gồm N = k^2 phần tử.",
            "            c) Sắp xếp mảng A theo thứ tự tăng dần: A_sorted <- SORT(A).",
            "            d) Chọn phần tử chính giữa: I_filtered[x, y] <- A_sorted[m_idx].",
            "   6. Trả về ảnh đã lọc trung vị I_filtered.",
        ],
        caption="Mã giả 2.4. Thuật toán Bộ lọc trung vị dựa trên sắp xếp thứ tự trong module src/filters.py",
    )

    add_h3(doc, "2.2.5. Phân tích và so sánh độ phức tạp tính toán của hai bộ lọc")
    add_para(
        doc,
        "Giả sử ảnh đầu vào có kích thước H hàng và W cột (tổng số điểm ảnh là M = H · W) và cửa sổ lọc có kích thước k x k (số phần tử trong một cửa sổ là N = k^2):"
    )
    add_bullet(
        doc,
        "Tại mỗi điểm ảnh trong số H · W điểm ảnh, phép tích chập thực hiện k^2 phép nhân và (k^2 - 1) phép cộng. Do đó, độ phức tạp thời gian lý thuyết là O(H · W · k^2) và độ phức tạp không gian bộ nhớ cho ảnh đệm là O((H + k) · (W + k)). Khi k tăng từ 3 lên 5 và 7, số phép tính tại mỗi pixel tăng từ 9 lên 25 và 49.",
        bold_prefix="Độ phức tạp của Bộ lọc Trung bình: "
    )
    add_bullet(
        doc,
        "Tại mỗi điểm ảnh trong số H · W điểm ảnh, thuật toán phải thực hiện sắp xếp một mảng gồm N = k^2 phần tử. Với các thuật toán sắp xếp so sánh tối ưu (như Quicksort / Introsort được cài đặt trong np.sort), chi phí sắp xếp một cửa sổ là O(k^2 · log_2(k^2)). Do đó, tổng độ phức tạp thời gian của bộ lọc trung vị là O(H · W · k^2 · log(k)). Vì có thêm nhân tử log(k^2) và chi phí hoán đổi bộ nhớ khi sắp xếp, bộ lọc trung vị luôn tiêu tốn nhiều thời gian CPU hơn bộ lọc trung bình ở cùng một kích thước cửa sổ k, và khoảng cách chênh lệch thời gian này sẽ càng nới rộng khi k tăng từ 3x3 (9 phần tử) lên 5x5 (25 phần tử) và 7x7 (49 phần tử).",
        bold_prefix="Độ phức tạp của Bộ lọc Trung vị: "
    )

    add_h3(doc, "2.2.6. Thuật toán phát hiện biên Sobel và đánh giá bảo toàn biên")
    add_para(
        doc,
        "Module src/sobel.py cài đặt hàm sobel_edge_detection(image, threshold, pad_mode) bằng cách tái sử dụng trực tiếp hàm tích chập tự viết convolve2d. Một chi tiết kỹ thuật tinh tế đã được xử lý chính xác trong code: Vì công thức toán học của đạo hàm Sobel định nghĩa theo phép tương quan không gian với nhân S_x = [[-1, 0, 1], [-2, 0, 2], [-1, 0, 1]], trong khi hàm convolve2d có bước lật nhân 180 độ, nên trước khi truyền vào convolve2d, chương trình thực hiện đảo chiều sẵn S_x và S_y. Nhờ đó, kết quả G_x và G_y thu được khớp tuyệt đối 100% (sai số 0.0) cả về độ lớn lẫn dấu so với hàm chuẩn cv2.Sobel của OpenCV."
    )

    add_code_block(
        doc,
        [
            "Thuật toán 2.5: Phát hiện biên Sobel và Đánh giá chỉ số bảo toàn biên",
            "-------------------------------------------------------------------------",
            "A. HÀM SOBEL_EDGE_DETECTION(I, threshold = None, pad_mode = 'reflect'):",
            "   1. Định nghĩa 2 nhân đạo hàm Sobel 3x3 chuẩn S_x và S_y.",
            "   2. Tính đạo hàm ngang và dọc qua hàm tích chập tự viết:",
            "        G_x <- CONVOLVE2D(I, flip180(S_x), pad_mode)",
            "        G_y <- CONVOLVE2D(I, flip180(S_y), pad_mode)",
            "   3. Tính biên độ Gradient thực: G_mag <- sqrt( G_x^2 + G_y^2 ).",
            "   4. Tạo ảnh biên mức xám hiển thị: G_vis <- uint8( clip( round(G_mag), 0, 255 ) ).",
            "   5. Xác định ngưỡng cắt biên T_edge:",
            "        Nếu threshold là None thì: T_edge <- Percentile_85(G_mag)  (Lấy trên ảnh gốc)",
            "        Ngược lại: T_edge <- threshold                             (Áp dụng cho ảnh lọc)",
            "   6. Tạo bản đồ biên nhị phân: E_bin[x, y] <- 1 nếu G_mag[x, y] >= T_edge, ngược lại 0.",
            "   7. Trả về (G_mag, G_vis, E_bin, (G_x, G_y), T_edge).",
            "",
            "B. HÀM EVALUATE_EDGE_PRESERVATION(G_orig, E_orig, G_filt, E_filt):",
            "   1. Chuẩn hóa G_orig và G_filt về đoạn [0, 255].",
            "   2. Tính Edge-MSE <- mean( (G_orig_norm - G_filt_norm)^2 ).",
            "   3. Đếm số điểm ảnh trùng khớp trên bản đồ biên nhị phân:",
            "        TP <- count( E_orig == 1 VÀ E_filt == 1 )",
            "        FP <- count( E_orig == 0 VÀ E_filt == 1 )",
            "        FN <- count( E_orig == 1 VÀ E_filt == 0 )",
            "   4. Tính Precision <- TP / (TP + FP) ;  Recall <- TP / (TP + FN).",
            "   5. Tính Edge-F1 <- (2 * Precision * Recall) / (Precision + Recall).",
            "   6. Trả về (Edge-MSE, Precision, Recall, Edge-F1).",
        ],
        caption="Mã giả 2.5. Thuật toán phát hiện biên Sobel và đánh giá bảo toàn biên trong src/sobel.py và src/metrics.py",
    )

    # --- MỤC 2.3 ---
    add_h2(doc, "2.3. Tổ chức cấu trúc thư mục mã nguồn của chương trình")
    add_para(
        doc,
        "Toàn bộ dự án được tổ chức khoa học, phân tách rõ ràng giữa mã nguồn thuật toán (src/), dữ liệu ảnh gốc (data/), kịch bản chạy thực nghiệm (run_experiments.py) và thư mục lưu trữ kết quả đầu ra tự động (results/). Cấu trúc cây thư mục của dự án như sau:"
    )

    add_code_block(
        doc,
        [
            "BTL_XLA/",
            "├── BTL_XLA.docx                        # Báo cáo Bài tập lớn hoàn chỉnh của Nhóm 10",
            "├── requirements.txt                    # Danh sách các gói thư viện Python cần thiết",
            "├── run_experiments.py                  # Script chính điều phối chạy 108 kịch bản thực nghiệm",
            "├── web_app.py                          # Ứng dụng Web tương tác chạy trên Localhost (cổng 8000)",
            "├── src/                                # Package chứa toàn bộ các thuật toán tự cài đặt",
            "│   ├── __init__.py                     # Khởi tạo package src",
            "│   ├── padding.py                      # Tự cài đặt mở rộng biên ảnh (zero, replicate, reflect)",
            "│   ├── noise.py                        # Tự cài đặt sinh nhiễu Muối tiêu & Gaussian (seed=42)",
            "│   ├── filters.py                      # Tự cài đặt tích chập 2D, Lọc Trung bình, Lọc Trung vị",
            "│   ├── sobel.py                        # Tự cài đặt toán tử Sobel 3x3 và ngưỡng hóa bản đồ biên",
            "│   ├── metrics.py                      # Tự cài đặt MSE, PSNR, Edge-MSE, Precision, Recall, F1",
            "│   ├── dataset.py                      # Tự động chuẩn bị 6 ảnh xám chuẩn kích thước 256x256",
            "│   ├── verification.py                 # Kiểm chứng độ chính xác code tự cài với OpenCV",
            "│   └── visualization.py                # Xuất hình ảnh so sánh trực quan và vẽ biểu đồ",
            "├── data/",
            "│   └── original/                       # Lưu 6 ảnh gốc chuẩn (cameraman, coins, astronaut, text...)",
            "└── results/",
            "    ├── tables/                         # Chứa các bảng số liệu thực nghiệm xuất ra định dạng .CSV",
            "    │   ├── verification_results.csv    # Bảng đối chiếu sai số giữa code tự cài và OpenCV",
            "    │   ├── detailed_results.csv        # Bảng số liệu chi tiết của toàn bộ 108 kịch bản lọc",
            "    │   └── summary_by_noise_and_filter.csv # Bảng tổng hợp trung bình trên 6 ảnh (36 dòng)",
            "    └── figures/                        # Chứa toàn bộ hình ảnh so sánh trực quan và biểu đồ (.PNG)",
        ],
        caption="Hình 2.3. Sơ đồ tổ chức cây thư mục mã nguồn của dự án BTL_XLA",
    )

    # --- MỤC 2.4 ---
    add_h2(doc, "2.4. Xây dựng giao diện ứng dụng Web tương tác chạy trên Localhost")
    add_para(
        doc,
        "Bên cạnh kịch bản chạy thực nghiệm hàng loạt (run_experiments.py), Nhóm 10 đã phát triển thêm một ứng dụng Web tương tác hoàn chỉnh đặt tại tệp web_app.py chạy trên máy chủ cục bộ Localhost (địa chỉ http://127.0.0.1:8000). Mục tiêu của ứng dụng Web là cung cấp một môi trường trực quan sinh động, cho phép giảng viên và người dùng kiểm thử trực tiếp các thuật toán tự cài đặt trên bất kỳ bức ảnh nào theo thời gian thực."
    )
    add_bullet(
        doc,
        "Sử dụng framework FastAPI kết hợp máy chủ ASGI Uvicorn. Backend cung cấp 5 điểm cuối RESTful API chính: GET '/' (trả về giao diện HTML5/CSS3/JS), POST '/api/select_sample' (chọn 1 trong 6 ảnh chuẩn), POST '/api/upload_image' (nhận tệp ảnh bất kỳ tải lên từ máy tính, chuyển sang ảnh xám 8-bit và chuẩn hóa kích thước 256x256 hoặc 512x512), POST '/api/process' (thực thi quy trình tạo nhiễu -> lọc Trung bình & Trung vị -> trích xuất biên Sobel -> tính chỉ số MSE, PSNR, Edge-MSE, Edge-F1), và POST '/api/reset' (khôi phục trạng thái ảnh về ảnh gốc sạch ban đầu). Toàn bộ ma trận ảnh NumPy uint8 sau khi xử lý được mã hóa trực tiếp trong bộ nhớ RAM sang định dạng ảnh PNG chuỗi Base64 (Data URL) để truyền tức thì về trình duyệt mà không cần ghi tệp tạm ra đĩa cứng.",
        bold_prefix="Kiến trúc Backend (FastAPI & NumPy Pipeline): "
    )
    add_bullet(
        doc,
        "Được thiết kế theo bố cục bảng điều khiển (Dashboard) gồm 2 vùng: Cột bên trái là bảng điều khiển cho phép chọn nguồn ảnh, tải ảnh trực tiếp, kéo thanh trượt (Slider) điều chỉnh mật độ nhiễu Muối tiêu (0.5% – 30%) hoặc độ lệch chuẩn nhiễu Gaussian (σ = 2 – 80), chọn kích thước cửa sổ lọc (3x3, 5x5, 7x7), chọn phương pháp đệm biên (Reflect, Replicate, Zero) và nút 'Reset Ảnh Gốc'. Khu vực chính bên phải hiển thị đồng thời 8 khung ảnh trực quan (Hàng 1: Ảnh gốc, Ảnh nhiễu, Ảnh sau Lọc Trung bình, Ảnh sau Lọc Trung vị; Hàng 2: 4 bản đồ biên Sobel tương ứng) kèm nút tải ảnh về máy và bảng đối sánh số liệu trực tiếp ở cuối trang.",
        bold_prefix="Kiến trúc Frontend (Interactive Dashboard): "
    )

    web_fig1 = os.path.join(fig_dir, "web_ui_salt_pepper.png")
    if os.path.exists(web_fig1):
        add_figure(
            doc,
            web_fig1,
            caption="Hình 2.4. Giao diện tổng thể của ứng dụng Web tương tác chạy trên Localhost (http://127.0.0.1:8000)",
            width_inches=6.1,
        )


def main() -> None:
    root_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    doc_path = os.path.join(root_dir, "BTL_XLA.docx")
    upd_path = os.path.join(root_dir, "BTL_XLA_updated.docx")
    fig_dir = os.path.join(root_dir, "results", "figures")

    src_path = upd_path if os.path.exists(upd_path) else doc_path
    doc = docx.Document(src_path)
    configure_styles(doc)
    remove_existing_chapter2_onwards(doc)
    write_chapter_2(doc, fig_dir)
    doc.save(upd_path)
    try:
        doc.save(doc_path)
        print(f"Đã ghi thành công Chương 2 vào: {doc_path}")
    except PermissionError:
        print(f"Đã ghi thành công Chương 2 vào: {upd_path}")


if __name__ == "__main__":
    main()

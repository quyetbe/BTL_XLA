# BÀI TẬP LỚN: XỬ LÝ ẢNH VÀ THỊ GIÁC MÁY TÍNH (IT3243)

**Đề tài 7: Khử nhiễu ảnh bằng các bộ lọc trong miền không gian**

- **Trường:** Đại học Công nghệ Đông Á (EAUT) — Khoa Công nghệ Thông tin
- **Học phần:** Xử lý ảnh và thị giác máy tính (`IT3243`)
- **Lớp tín chỉ:** `XỬ LÝ ẢNH VÀ THỊ GIÁC MÁY TÍNH-1-1-26(N05.THTT_IN_ELN.11)`
- **Giảng viên hướng dẫn:** Thầy Nguyễn Phồn Lữa
- **Nhóm thực hiện:** Nhóm 10

| TT | Mã sinh viên | Họ và tên | Lớp hành chính | Vai trò |
| :-: | :-: | :--- | :-: | :-: |
| 1 | `20231053` | **Nguyễn Mạnh Quyết** | `DCCNTT.14.3` | Nhóm trưởng |
| 2 | `20231045` | **Hà Thái Bảo** | `DCCNTT.14.3` | Thành viên |

---

## 1. Giới thiệu đề tài & Ràng buộc kỹ thuật

Dự án nghiên cứu, tự cài đặt và đánh giá đối sánh hai bộ lọc khử nhiễu kinh điển trong miền không gian (**Bộ lọc Trung bình** và **Bộ lọc Trung vị**) kết hợp với **toán tử phát hiện biên Sobel** để khảo sát khả năng bảo toàn cấu trúc biên ảnh.

**Tuân thủ ràng buộc kỹ thuật:**
- Ngôn ngữ **Python 3**, tự cài đặt hoàn toàn bằng **NumPy** đối với tất cả thuật toán cốt lõi:
  - Mở rộng biên ảnh 2D (`zero`, `replicate`, `reflect`) trong `src/padding.py`.
  - Giả lập nhiễu Muối tiêu (`1%`, `5%`, `10%`) và nhiễu cộng Gaussian (`sigma = 10, 25, 50`) với hạt giống cố định `seed = 42` trong `src/noise.py`.
  - Phép tích chập 2D (`convolve2d`), Bộ lọc Trung bình (`mean_filter`), Bộ lọc Trung vị (`median_filter`) với cửa sổ `3x3`, `5x5`, `7x7` trong `src/filters.py`.
  - Toán tử phát hiện biên Sobel 3x3 (`Gx`, `Gy`, độ lớn Gradient và ngưỡng hóa phân vị thứ 85) trong `src/sobel.py`.
  - Các chỉ số đánh giá `MSE`, `PSNR`, `Edge-MSE`, `Precision`, `Recall`, `Edge F1-Score` trong `src/metrics.py`.
- **OpenCV / Pillow** chỉ được dùng để đọc/ghi ảnh và đối chiếu kiểm chứng tính đúng đắn (`src/verification.py`) với `cv2.blur`, `cv2.medianBlur`, `cv2.Sobel` (đạt sai số tuyệt đối cực đại bằng `0.0`).

---

## 2. Cấu trúc thư mục dự án

```text
BTL_XLA/
├── README.md                              # Tài liệu hướng dẫn dự án
├── requirements.txt                       # Danh sách thư viện Python cần thiết
├── run_experiments.py                     # Script chạy toàn bộ 108 kịch bản thực nghiệm & xuất CSV/PNG
├── web_app.py                             # Ứng dụng Web tương tác trên Localhost (FastAPI + Uvicorn)
├── BTL_XLA.docx                           # Báo cáo Bài tập lớn hoàn chỉnh của Nhóm 10
├── .vscode/                               # Cấu hình sẵn Interpreter & Run/Debug cho VS Code
├── src/                                   # Package chứa các module thuật toán tự cài đặt
│   ├── __init__.py
│   ├── padding.py                         # Hàm pad_image ('zero', 'replicate', 'reflect')
│   ├── noise.py                           # Hàm add_salt_and_pepper_noise & add_gaussian_noise
│   ├── filters.py                         # Hàm convolve2d, mean_filter, median_filter
│   ├── sobel.py                           # Hàm sobel_edge_detection
│   ├── metrics.py                         # Hàm calculate_mse, calculate_psnr, evaluate_edge_preservation
│   ├── dataset.py                         # Chuẩn bị bộ 6 ảnh xám chuẩn 256x256
│   ├── verification.py                    # Đối chiếu kiểm thử hàm tự cài đặt với OpenCV
│   └── visualization.py                   # Xuất hình so sánh trực quan và biểu đồ Matplotlib
├── data/
│   └── original/                          # 6 ảnh gốc chuẩn: cameraman, coins, astronaut, text, checkerboard, brick
├── results/
│   ├── tables/                            # Kết quả xuất CSV (verification, detailed_results, summary)
│   └── figures/                           # Hình minh họa, bản đồ biên Sobel, biểu đồ và ảnh chụp Web UI
└── scripts/                               # Các script hỗ trợ tạo sơ đồ và biên soạn báo cáo .docx
```

---

## 3. Cài đặt môi trường & Chạy chương trình

### 3.1. Cài đặt thư viện
Yêu cầu Python 3.10+ (khuyến nghị tạo môi trường ảo `.venv`):

```powershell
# (Tùy chọn) Tạo và kích hoạt môi trường ảo
python -m venv .venv
.\.venv\Scripts\Activate.ps1

# Cài đặt các gói phụ thuộc
pip install -r requirements.txt
```

> **Lưu ý trên Windows 11 (Smart App Control):** Nếu gặp cảnh báo bảo mật liên quan đến `cu2qu*.pyd` của gói `fonttools` khi chạy Matplotlib, hãy cài lại `fonttools` bản thuần Python:
> ```powershell
> pip install --force-reinstall --no-binary fonttools fonttools
> ```

### 3.2. Chạy kịch bản thực nghiệm hàng loạt (`run_experiments.py`)
```powershell
.\.venv\Scripts\python.exe run_experiments.py
```
Kết quả tự động xuất ra:
- `results/tables/verification_results.csv`: Bảng kiểm chứng đối chiếu với OpenCV.
- `results/tables/detailed_results.csv`: Số liệu chi tiết cho toàn bộ `6 ảnh × 2 loại nhiễu × 3 mức × 2 bộ lọc × 3 kích thước cửa sổ` (`108` kịch bản lọc, `216` dòng dữ liệu).
- `results/tables/summary_by_noise_and_filter.csv`: Bảng tổng hợp trung bình trên 6 ảnh (`36` cấu hình).
- `results/figures/*.png`: Các hình ảnh so sánh trực quan, bản đồ biên Sobel và biểu đồ PSNR / Edge-F1 / Thời gian xử lý.

### 3.3. Chạy Giao diện Web Demo trên Localhost (`web_app.py`)
```powershell
.\.venv\Scripts\python.exe web_app.py
```
Sau khi chạy, trình duyệt sẽ tự động mở địa chỉ **`http://127.0.0.1:8000`** (nhấn `Ctrl + C` trong Terminal để dừng server).

**Tính năng chính trên Web UI:**
- Chọn nhanh 1 trong 6 ảnh chuẩn hoặc **Upload trực tiếp ảnh bất kỳ** từ máy tính (hỗ trợ chuẩn hóa `256x256`, `512x512` hoặc giữ tỷ lệ gốc).
- Tùy chỉnh loại nhiễu (**Muối tiêu**, **Gaussian**, **Hỗn hợp**, hoặc **Không nhiễu**), điều chỉnh cường độ theo thời gian thực qua thanh trượt và cố định `seed`.
- Chuyển đổi kích thước cửa sổ (`3x3`, `5x5`, `7x7`) và chế độ xử lý biên (`reflect`, `replicate`, `zero`).
- Quan sát song song 8 khung ảnh **Trước / Sau lọc** và **Bản đồ biên Sobel**, bảng đối chiếu chỉ số (`MSE`, `PSNR`, `Edge-MSE`, `Precision`, `Recall`, `Edge-F1`, `Thời gian`), nút **Reset Ảnh Gốc** và liên kết tải từng ảnh kết quả.

### 3.4. Chạy trên VS Code hoặc PyCharm
- **VS Code:** Mở thư mục `BTL_XLA`, cấu hình sẵn trong `.vscode/settings.json` và `.vscode/launch.json` cho phép nhấn `F5` (hoặc chọn *Run Python File*) để chạy ngay `web_app.py` hoặc `run_experiments.py`.
- **PyCharm:** Mở thư mục `BTL_XLA`, trỏ Python Interpreter tới `.venv\Scripts\python.exe`, chuột phải vào `web_app.py` hoặc `run_experiments.py` và chọn **Run**.

---

## 4. Tóm tắt kết quả thực nghiệm tiêu biểu

1. **Kiểm chứng độ chính xác (`verification_results.csv`):**
   - Lọc Trung bình (`3x3`, `5x5`, `7x7`) vs. `cv2.blur`: Sai số cực đại = **`0.0`** (Đạt).
   - Lọc Trung vị (`3x3`, `5x5`, `7x7`) vs. `cv2.medianBlur`: Sai số cực đại = **`0.0`** (Đạt).
   - Toán tử Sobel `3x3` (`Gx`, `Gy`, Gradient Magnitude) vs. `cv2.Sobel`: Sai số cực đại = **`0.0`** (Đạt).

2. **Tổng hợp hiệu năng trung bình trên 6 ảnh chuẩn (`summary_by_noise_and_filter.csv`):**
   - **Nhiễu Muối tiêu (`1%`, `5%`, `10%`):** Bộ lọc **Trung vị `3x3`** vượt trội hoàn toàn, đạt PSNR trung bình lần lượt là **`31.49 dB`**, **`29.43 dB`**, **`28.04 dB`** (so với `26.08 dB`, `23.66 dB`, `21.69 dB` của Lọc Trung bình `3x3`) và bảo toàn biên Sobel tốt hơn hẳn (`Edge-F1` đạt **`0.7971`**, **`0.7899`**, **`0.7771`**).
   - **Nhiễu Gaussian (`sigma = 10, 25, 50`):** Ở mức nhiễu nặng `sigma = 50`, Bộ lọc **Trung bình `3x3`** đạt PSNR trung bình **`21.64 dB`** (`Edge-F1 = 0.5722`), vượt trội so với Lọc Trung vị `3x3` (`21.15 dB`, `Edge-F1 = 0.5297`), đồng thời có tốc độ xử lý nhanh gấp **2.85 – 10.95 lần**.

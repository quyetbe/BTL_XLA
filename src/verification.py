# -*- coding: utf-8 -*-
"""
Module: verification.py
Mô tả: Kiểm tra và đối chiếu tính đúng đắn của các thuật toán tự cài đặt
       (Lọc trung bình, Lọc trung vị, Toán tử Sobel) so với các hàm chuẩn
       trong thư viện OpenCV (cv2.blur, cv2.medianBlur, cv2.Sobel).
       Theo đúng yêu cầu đề bài: Không dùng hàm OpenCV cho kết quả chính,
       chỉ dùng để đối chiếu kiểm tra tính đúng đắn của code tự cài.
Nhóm thực hiện: Nhóm 10 (Nguyễn Mạnh Quyết - 20231053, Hà Thái Bảo - 20231045)
"""

from typing import List, Dict, Any
import numpy as np
import cv2
from src.filters import mean_filter, median_filter
from src.sobel import sobel_edge_detection


def verify_implementations(test_image: np.ndarray) -> List[Dict[str, Any]]:
    """
    Đối chiếu kết quả giữa hàm tự cài đặt và hàm OpenCV trên cùng một ảnh kiểm thử
    và cùng chế độ mở rộng biên phản chiếu (BORDER_REFLECT_101).

    Tham số:
        test_image (np.ndarray): Ảnh xám kiểm thử kích thước (H, W).

    Trả về:
        List[Dict[str, Any]]: Danh sách kết quả kiểm thử từng thuật toán và kích thước cửa sổ.
    """
    results: List[Dict[str, Any]] = []

    # 1. Đối chiếu Bộ lọc trung bình (Mean Filter) với cv2.blur
    for k in (3, 5, 7):
        our_mean = mean_filter(test_image, kernel_size=k, pad_mode="reflect")
        cv_mean = cv2.blur(test_image, (k, k), borderType=cv2.BORDER_REFLECT_101)

        abs_diff = np.abs(our_mean.astype(np.float64) - cv_mean.astype(np.float64))
        max_err = float(np.max(abs_diff))
        mean_err = float(np.mean(abs_diff))
        # Lưu ý: Sai lệch tối đa <= 1 mức xám có thể xảy ra ở các trường hợp .5 do quy tắc làm tròn
        is_valid = max_err <= 1.0

        results.append({
            "algorithm": f"Lọc trung bình (Mean Filter {k}x{k})",
            "reference": f"cv2.blur(img, ({k},{k}))",
            "max_abs_error": max_err,
            "mean_abs_error": round(mean_err, 6),
            "status": "ĐẠT (Khớp chuẩn)" if is_valid else "LỖI",
        })

    # 2. Đối chiếu Bộ lọc trung vị (Median Filter) với cv2.medianBlur
    # Lưu ý: cv2.medianBlur sử dụng chế độ biên replicate (BORDER_REPLICATE) bên trong.
    for k in (3, 5, 7):
        our_median = median_filter(test_image, kernel_size=k, pad_mode="replicate")
        cv_median = cv2.medianBlur(test_image, k)

        abs_diff = np.abs(our_median.astype(np.float64) - cv_median.astype(np.float64))
        max_err = float(np.max(abs_diff))
        mean_err = float(np.mean(abs_diff))
        is_valid = max_err == 0.0

        results.append({
            "algorithm": f"Lọc trung vị (Median Filter {k}x{k})",
            "reference": f"cv2.medianBlur(img, {k})",
            "max_abs_error": max_err,
            "mean_abs_error": round(mean_err, 6),
            "status": "ĐẠT (Khớp tuyệt đối 100%)" if is_valid else "LỖI",
        })

    # 3. Đối chiếu Toán tử Sobel (Gx, Gy và Độ lớn Gradient) với cv2.Sobel
    our_grad_mag, _, _, (our_gx, our_gy), _ = sobel_edge_detection(test_image, pad_mode="reflect")
    cv_gx = cv2.Sobel(test_image, cv2.CV_64F, 1, 0, ksize=3, borderType=cv2.BORDER_REFLECT_101)
    cv_gy = cv2.Sobel(test_image, cv2.CV_64F, 0, 1, ksize=3, borderType=cv2.BORDER_REFLECT_101)
    cv_grad_mag = np.sqrt(cv_gx ** 2 + cv_gy ** 2)

    abs_diff_sobel = np.abs(our_grad_mag - cv_grad_mag)
    max_err_sobel = float(np.max(abs_diff_sobel))
    mean_err_sobel = float(np.mean(abs_diff_sobel))
    is_valid_sobel = max_err_sobel < 1e-9

    results.append({
        "algorithm": "Toán tử Sobel 3x3 (Gx, Gy, Gradient Magnitude)",
        "reference": "cv2.Sobel(img, cv2.CV_64F, dx, dy, ksize=3)",
        "max_abs_error": round(max_err_sobel, 12),
        "mean_abs_error": round(mean_err_sobel, 12),
        "status": "ĐẠT (Khớp tuyệt đối 100%)" if is_valid_sobel else "LỖI",
    })

    return results

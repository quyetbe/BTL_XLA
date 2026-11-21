# -*- coding: utf-8 -*-
"""
Module: filters.py
Mô tả: Tự cài đặt phép tích chập 2D (2D Convolution), Bộ lọc trung bình (Mean Filter)
       và Bộ lọc trung vị (Median Filter) trong miền không gian.
       Tuyệt đối không sử dụng cv2.blur hay cv2.medianBlur cho kết quả chính.
Nhóm thực hiện: Nhóm 10 (Nguyễn Mạnh Quyết - 20231053, Hà Thái Bảo - 20231045)
"""

import numpy as np
from numpy.lib.stride_tricks import sliding_window_view
from src.padding import pad_image


def convolve2d(image: np.ndarray, kernel: np.ndarray, pad_mode: str = "reflect") -> np.ndarray:
    """
    Thực hiện phép tích chập 2 chiều (2D Convolution) tự cài đặt trên ảnh xám.
    Công thức toán học:
        (f * h)(x, y) = sum_{s=-a}^{a} sum_{t=-b}^{b} h(s, t) * f(x - s, y - t)

    Tham số:
        image (np.ndarray): Ảnh đầu vào 2 chiều (H, W).
        kernel (np.ndarray): Ma trận nhân lọc kích thước lẻ (k_h, k_w).
        pad_mode (str): Phương pháp mở rộng biên ('reflect', 'replicate', 'zero').

    Trả về:
        np.ndarray: Ma trận kết quả tích chập kiểu float64 kích thước (H, W).
    """
    if image.ndim != 2 or kernel.ndim != 2:
        raise ValueError("Ảnh và nhân lọc (kernel) phải là ma trận 2 chiều.")

    k_h, k_w = kernel.shape
    if k_h % 2 == 0 or k_w % 2 == 0:
        raise ValueError("Kích thước nhân lọc (kernel) phải là các số lẻ (ví dụ: 3x3, 5x5, 7x7).")

    # Lật nhân lọc 180 độ (theo cả trục dọc và trục ngang) theo đúng định nghĩa tích chập
    flipped_kernel = np.flip(np.flip(kernel, axis=0), axis=1).astype(np.float64)

    pad_h = k_h // 2
    pad_w = k_w // 2

    # Mở rộng biên ảnh bằng hàm tự cài đặt
    padded = pad_image(image.astype(np.float64), pad_h, pad_w, mode=pad_mode)

    # Trích xuất các cửa sổ trượt kích thước (H, W, k_h, k_w) và tính tổng tích vô hướng
    windows = sliding_window_view(padded, window_shape=(k_h, k_w))
    output = np.einsum("ijkl,kl->ij", windows, flipped_kernel)

    return output


def mean_filter(image: np.ndarray, kernel_size: int = 3, pad_mode: str = "reflect") -> np.ndarray:
    """
    Bộ lọc trung bình số học (Arithmetic Mean Filter) tự cài đặt.
    Mỗi điểm ảnh đầu ra là trung bình cộng của k*k điểm ảnh trong cửa sổ lân cận.

    Tham số:
        image (np.ndarray): Ảnh xám đầu vào kích thước (H, W), kiểu uint8.
        kernel_size (int): Kích thước cửa sổ lọc k (k = 3, 5, 7...).
        pad_mode (str): Phương pháp mở rộng biên ('reflect', 'replicate', 'zero').

    Trả về:
        np.ndarray: Ảnh sau khi lọc trung bình, kiểu uint8.
    """
    if kernel_size % 2 == 0 or kernel_size < 1:
        raise ValueError("Kích thước cửa sổ lọc (kernel_size) phải là số nguyên dương lẻ.")

    # Khởi tạo mặt nạ trung bình H = (1 / k^2) * J_{k x k}
    kernel = np.ones((kernel_size, kernel_size), dtype=np.float64) / (kernel_size * kernel_size)

    # Thực hiện tích chập 2D tự viết
    filtered_float = convolve2d(image, kernel, pad_mode=pad_mode)

    # Làm tròn và giới hạn trong đoạn [0, 255]
    return np.clip(np.rint(filtered_float), 0, 255).astype(np.uint8)


def median_filter(image: np.ndarray, kernel_size: int = 3, pad_mode: str = "reflect") -> np.ndarray:
    """
    Bộ lọc trung vị (Median Filter) tự cài đặt.
    Mỗi điểm ảnh đầu ra được thay bằng giá trị đứng chính giữa (phần tử thứ (k^2 - 1)//2)
    sau khi sắp xếp tăng dần toàn bộ k*k điểm ảnh trong cửa sổ lân cận.

    Tham số:
        image (np.ndarray): Ảnh xám đầu vào kích thước (H, W), kiểu uint8.
        kernel_size (int): Kích thước cửa sổ lọc k (k = 3, 5, 7...).
        pad_mode (str): Phương pháp mở rộng biên ('reflect', 'replicate', 'zero').

    Trả về:
        np.ndarray: Ảnh sau khi lọc trung vị, kiểu uint8.
    """
    if kernel_size % 2 == 0 or kernel_size < 1:
        raise ValueError("Kích thước cửa sổ lọc (kernel_size) phải là số nguyên dương lẻ.")

    pad = kernel_size // 2
    H, W = image.shape

    # Mở rộng biên ảnh bằng hàm tự cài đặt
    padded = pad_image(image, pad, pad, mode=pad_mode)

    # Tạo tập cửa sổ trượt 4 chiều (H, W, k, k) rồi trải phẳng lân cận thành (H, W, k*k)
    windows = sliding_window_view(padded, window_shape=(kernel_size, kernel_size))
    flat_windows = windows.reshape(H, W, kernel_size * kernel_size)

    # Sắp xếp các phần tử trong từng cửa sổ và lấy phần tử đứng chính giữa (chỉ số trung vị)
    median_idx = (kernel_size * kernel_size) // 2
    sorted_windows = np.sort(flat_windows, axis=-1)
    median_result = sorted_windows[:, :, median_idx]

    return median_result.astype(np.uint8)

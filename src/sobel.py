# -*- coding: utf-8 -*-
"""
Module: sobel.py
Mô tả: Tự cài đặt toán tử phát hiện biên Sobel bằng tích chập 2D tự viết.
       Tính gradient theo phương ngang Gx, phương dọc Gy, độ lớn gradient G
       và bản đồ biên nhị phân phục vụ đánh giá khả năng bảo toàn biên sau lọc.
Nhóm thực hiện: Nhóm 10 (Nguyễn Mạnh Quyết - 20231053, Hà Thái Bảo - 20231045)
"""

from typing import Tuple, Optional
import numpy as np
from src.filters import convolve2d


# Định nghĩa hai mặt nạ đạo hàm Sobel 3x3 chuẩn (dạng tương quan / đạo hàm không gian)
SOBEL_KERNEL_X = np.array([
    [-1.0,  0.0,  1.0],
    [-2.0,  0.0,  2.0],
    [-1.0,  0.0,  1.0]
], dtype=np.float64)

SOBEL_KERNEL_Y = np.array([
    [-1.0, -2.0, -1.0],
    [ 0.0,  0.0,  0.0],
    [ 1.0,  2.0,  1.0]
], dtype=np.float64)


def sobel_edge_detection(
    image: np.ndarray,
    threshold: Optional[float] = None,
    pad_mode: str = "reflect"
) -> Tuple[np.ndarray, np.ndarray, np.ndarray, np.ndarray, float]:
    """
    Thực hiện phát hiện biên Sobel trên ảnh xám bằng phép tích chập tự cài đặt.

    Lưu ý toán học: Vì toán tử Sobel định nghĩa theo phép tương quan không gian (correlation)
    với mặt nạ bất đối xứng dấu [-1, 0, 1], khi gọi qua hàm tích chập `convolve2d` (có bước
    lật mặt nạ 180 độ), ta truyền mặt nạ đã đảo chiều sẵn để thu được đúng đạo hàm Gx, Gy.

    Tham số:
        image (np.ndarray): Ảnh xám đầu vào (H, W).
        threshold (float, optional): Ngưỡng nhị phân hóa biên. Nếu bằng None, tự động chọn
            ngưỡng theo phương pháp phân vị thứ 85 (85th percentile) của độ lớn gradient.
        pad_mode (str): Chế độ mở rộng biên ('reflect', 'replicate', 'zero').

    Trả về:
        Tuple gồm 5 phần tử:
            - grad_mag (np.ndarray): Độ lớn gradient thực G = sqrt(Gx^2 + Gy^2) (kiểu float64).
            - grad_mag_uint8 (np.ndarray): Bản đồ biên mức xám chuẩn hóa về [0, 255] để hiển thị.
            - binary_edge (np.ndarray): Bản đồ biên nhị phân (0 hoặc 1) sau khi cắt ngưỡng.
            - (gx, gy) (Tuple[np.ndarray, np.ndarray]): Đạo hàm theo phương X và phương Y.
            - used_threshold (float): Giá trị ngưỡng thực tế đã áp dụng.
    """
    # Lật sẵn mặt nạ để khi convolve2d lật lại sẽ khớp chính xác với công thức đạo hàm Sobel
    kernel_x_for_conv = np.flip(np.flip(SOBEL_KERNEL_X, axis=0), axis=1)
    kernel_y_for_conv = np.flip(np.flip(SOBEL_KERNEL_Y, axis=0), axis=1)

    gx = convolve2d(image, kernel_x_for_conv, pad_mode=pad_mode)
    gy = convolve2d(image, kernel_y_for_conv, pad_mode=pad_mode)

    # Tính độ lớn vector Gradient: G = sqrt(Gx^2 + Gy^2)
    grad_mag = np.sqrt(gx ** 2 + gy ** 2)

    # Chuẩn hóa bản đồ gradient để hiển thị trực quan (giới hạn ở 255)
    grad_mag_uint8 = np.clip(np.rint(grad_mag), 0, 255).astype(np.uint8)

    # Xác định ngưỡng tạo bản đồ biên nhị phân
    if threshold is None:
        # Chọn phân vị thứ 85 (top 15% điểm có độ dốc biên mạnh nhất trên ảnh chuẩn)
        used_threshold = float(np.percentile(grad_mag, 85.0))
    else:
        used_threshold = float(threshold)

    binary_edge = (grad_mag >= used_threshold).astype(np.uint8)

    return grad_mag, grad_mag_uint8, binary_edge, (gx, gy), used_threshold

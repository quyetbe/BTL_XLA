# -*- coding: utf-8 -*-
"""
Module: padding.py
Mô tả: Tự cài đặt các phương pháp mở rộng biên ảnh (Image Boundary Padding)
       phục vụ phép tích chập 2D và các bộ lọc không gian.
Nhóm thực hiện: Nhóm 10 (Nguyễn Mạnh Quyết - 20231053, Hà Thái Bảo - 20231045)
"""

import numpy as np


def pad_image(image: np.ndarray, pad_h: int, pad_w: int, mode: str = "reflect") -> np.ndarray:
    """
    Mở rộng biên ảnh 2D (ảnh xám) thủ công không dùng hàm lọc có sẵn của OpenCV.

    Tham số:
        image (np.ndarray): Ma trận ảnh xám 2 chiều kích thước (H, W).
        pad_h (int): Số hàng cần đệm thêm ở mỗi phía (trên và dưới).
        pad_w (int): Số cột cần đệm thêm ở mỗi phía (trái và phải).
        mode (str): Chế độ đệm biên:
            - 'zero': Đệm giá trị 0 xung quanh biên.
            - 'replicate' (hoặc 'edge'): Lặp lại điểm ảnh ngoài cùng của biên.
            - 'reflect': Phản chiếu gương qua đường biên (mặc định, giúp tránh nhiễu giả ở biên).

    Trả về:
        np.ndarray: Ảnh sau khi mở rộng biên có kích thước (H + 2*pad_h, W + 2*pad_w).
    """
    if image.ndim != 2:
        raise ValueError("Hàm pad_image chỉ nhận ảnh xám 2 chiều (H, W).")
    if pad_h == 0 and pad_w == 0:
        return image.copy()

    H, W = image.shape
    new_H = H + 2 * pad_h
    new_W = W + 2 * pad_w

    # Khởi tạo ma trận kết quả bằng 0 với cùng kiểu dữ liệu
    padded = np.zeros((new_H, new_W), dtype=image.dtype)

    # Đặt vùng ảnh gốc vào chính giữa ma trận mở rộng
    padded[pad_h:pad_h + H, pad_w:pad_w + W] = image

    if mode == "zero":
        return padded

    elif mode in ("replicate", "edge"):
        # Đệm phía trên và phía dưới bằng cách lặp lại dòng đầu và dòng cuối
        for i in range(pad_h):
            padded[i, pad_w:pad_w + W] = image[0, :]
            padded[pad_h + H + i, pad_w:pad_w + W] = image[-1, :]

        # Đệm phía trái và phía phải (bao gồm cả 4 góc)
        for j in range(pad_w):
            padded[:, j] = padded[:, pad_w]
            padded[:, pad_w + W + j] = padded[:, pad_w + W - 1]

        return padded

    elif mode == "reflect":
        # Đệm phía trên và phía dưới theo nguyên tắc phản chiếu gương qua trục biên
        for i in range(pad_h):
            src_top = min(pad_h - i, H - 1)
            src_bottom = max(H - 2 - i, 0)
            padded[i, pad_w:pad_w + W] = image[src_top, :]
            padded[pad_h + H + i, pad_w:pad_w + W] = image[src_bottom, :]

        # Đệm phía trái và phía phải (bao gồm cả 4 góc phản chiếu)
        for j in range(pad_w):
            src_left = min(pad_w - j, W - 1)
            src_right = max(W - 2 - j, 0)
            padded[:, j] = padded[:, pad_w + src_left]
            padded[:, pad_w + W + j] = padded[:, pad_w + src_right]

        return padded

    else:
        raise ValueError(f"Chế độ padding '{mode}' không được hỗ trợ. Chọn 'zero', 'replicate' hoặc 'reflect'.")

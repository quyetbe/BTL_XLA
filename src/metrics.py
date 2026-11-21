# -*- coding: utf-8 -*-
"""
Module: metrics.py
Mô tả: Tự cài đặt các chỉ số đánh giá chất lượng ảnh và khả năng bảo toàn biên:
       1. MSE (Mean Squared Error) - Sai số bình phương trung bình.
       2. PSNR (Peak Signal-to-Noise Ratio) - Tỷ số tín hiệu cực đại trên nhiễu (dB).
       3. Edge-MSE - Sai số bình phương trung bình giữa hai bản đồ độ lớn Gradient Sobel.
       4. Edge Precision, Recall và F1-Score - Đánh giá độ trùng khớp vị trí biên nhị phân.
Nhóm thực hiện: Nhóm 10 (Nguyễn Mạnh Quyết - 20231053, Hà Thái Bảo - 20231045)
"""

from typing import Dict
import numpy as np


def calculate_mse(original: np.ndarray, processed: np.ndarray) -> float:
    """
    Tính sai số bình phương trung bình (MSE) giữa ảnh gốc và ảnh đã xử lý.
    Công thức: MSE = (1 / (M * N)) * sum_{i=0}^{M-1} sum_{j=0}^{N-1} [I(i,j) - K(i,j)]^2
    """
    if original.shape != processed.shape:
        raise ValueError("Hai ảnh phải có cùng kích thước khi tính MSE.")

    diff = original.astype(np.float64) - processed.astype(np.float64)
    mse = float(np.mean(diff ** 2))
    return mse


def calculate_psnr(original: np.ndarray, processed: np.ndarray, max_pixel: float = 255.0) -> float:
    """
    Tính tỷ số tín hiệu cực đại trên nhiễu (PSNR) theo đơn vị decibel (dB).
    Công thức: PSNR = 10 * log10(MAX_I^2 / MSE)
    """
    mse = calculate_mse(original, processed)
    if mse == 0.0:
        return float("inf")
    psnr = 10.0 * np.log10((max_pixel ** 2) / mse)
    return float(psnr)


def evaluate_edge_preservation(
    ref_grad_mag: np.ndarray,
    ref_binary_edge: np.ndarray,
    test_grad_mag: np.ndarray,
    test_binary_edge: np.ndarray
) -> Dict[str, float]:
    """
    Đánh giá khả năng bảo toàn biên bằng cách so sánh bản đồ biên Sobel của ảnh sau lọc
    với bản đồ biên Sobel của ảnh gốc (làm chuẩn Ground-Truth).

    Tham số:
        ref_grad_mag (np.ndarray): Độ lớn gradient Sobel của ảnh gốc.
        ref_binary_edge (np.ndarray): Bản đồ biên nhị phân (0/1) của ảnh gốc.
        test_grad_mag (np.ndarray): Độ lớn gradient Sobel của ảnh cần đánh giá.
        test_binary_edge (np.ndarray): Bản đồ biên nhị phân (0/1) của ảnh cần đánh giá
                                       (được cắt bằng cùng ngưỡng của ảnh gốc).

    Trả về:
        Dict[str, float]: Từ điển chứa 'edge_mse', 'precision', 'recall', 'f1'.
    """
    # 1. Tính MSE giữa hai bản đồ độ lớn Gradient chuẩn hóa về [0, 255]
    ref_norm = np.clip(ref_grad_mag, 0.0, 255.0)
    test_norm = np.clip(test_grad_mag, 0.0, 255.0)
    edge_mse = calculate_mse(ref_norm, test_norm)

    # 2. Tính Precision, Recall, F1-Score của các điểm biên nhị phân
    ref_bool = ref_binary_edge.astype(bool)
    test_bool = test_binary_edge.astype(bool)

    tp = float(np.sum(ref_bool & test_bool))
    fp = float(np.sum((~ref_bool) & test_bool))
    fn = float(np.sum(ref_bool & (~test_bool)))

    precision = tp / (tp + fp) if (tp + fp) > 0 else 0.0
    recall = tp / (tp + fn) if (tp + fn) > 0 else 0.0
    f1 = (2.0 * precision * recall / (precision + recall)) if (precision + recall) > 0 else 0.0

    return {
        "edge_mse": edge_mse,
        "precision": precision,
        "recall": recall,
        "f1": f1,
    }

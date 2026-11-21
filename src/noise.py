# -*- coding: utf-8 -*-
"""
Module: noise.py
Mô tả: Tự cài đặt các hàm giả lập nhiễu ảnh trong miền không gian:
       1. Nhiễu muối tiêu (Salt-and-Pepper Noise) với mật độ p (1%, 5%, 10%).
       2. Nhiễu cộng Gaussian (Additive White Gaussian Noise) với độ lệch chuẩn sigma (10, 25, 50).
       Sử dụng seed cố định để đảm bảo tính tái lập hoàn toàn của thực nghiệm.
Nhóm thực hiện: Nhóm 10 (Nguyễn Mạnh Quyết - 20231053, Hà Thái Bảo - 20231045)
"""

import numpy as np


def add_salt_and_pepper_noise(
    image: np.ndarray,
    density: float,
    salt_vs_pepper: float = 0.5,
    seed: int = 42
) -> np.ndarray:
    """
    Thêm nhiễu muối tiêu (Salt-and-Pepper) vào ảnh xám 8-bit.

    Tham số:
        image (np.ndarray): Ảnh xám gốc kích thước (H, W), giá trị [0, 255].
        density (float): Mật độ nhiễu tổng cộng (ví dụ: 0.01 cho 1%, 0.05 cho 5%, 0.10 cho 10%).
        salt_vs_pepper (float): Tỷ lệ phân bố giữa điểm muối (255) và điểm tiêu (0), mặc định 0.5.
        seed (int): Hạt giống ngẫu nhiên cố định để tái lập kết quả.

    Trả về:
        np.ndarray: Ảnh đã bị nhiễu muối tiêu, kiểu uint8.
    """
    if not (0.0 <= density <= 1.0):
        raise ValueError("Mật độ nhiễu (density) phải nằm trong đoạn [0.0, 1.0].")

    rng = np.random.default_rng(seed)
    noisy = image.copy().astype(np.uint8)

    # Sinh ma trận xác suất phân phối đều U(0, 1) có cùng kích thước với ảnh
    random_matrix = rng.random(image.shape)

    # Xác suất điểm muối (trắng = 255) và điểm tiêu (đen = 0)
    prob_salt = density * salt_vs_pepper
    prob_pepper = density * (1.0 - salt_vs_pepper)

    # Gán giá trị 255 (muối) cho các vị trí có xác suất < prob_salt
    noisy[random_matrix < prob_salt] = 255

    # Gán giá trị 0 (tiêu) cho các vị trí nằm trong khoảng [prob_salt, prob_salt + prob_pepper)
    noisy[(random_matrix >= prob_salt) & (random_matrix < prob_salt + prob_pepper)] = 0

    return noisy


def add_gaussian_noise(
    image: np.ndarray,
    sigma: float,
    mean: float = 0.0,
    seed: int = 42
) -> np.ndarray:
    """
    Thêm nhiễu cộng Gaussian trắng (AWGN) vào ảnh xám 8-bit.
    Mô hình toán học: g(x, y) = clip(round(f(x, y) + eta(x, y)), 0, 255), với eta ~ N(mean, sigma^2).

    Tham số:
        image (np.ndarray): Ảnh xám gốc kích thước (H, W), giá trị [0, 255].
        sigma (float): Độ lệch chuẩn của phân phối chuẩn (ví dụ: 10, 25, 50).
        mean (float): Giá trị trung bình (kỳ vọng) của nhiễu, mặc định bằng 0.0.
        seed (int): Hạt giống ngẫu nhiên cố định để tái lập kết quả.

    Trả về:
        np.ndarray: Ảnh đã bị nhiễu Gaussian, kiểu uint8.
    """
    if sigma < 0:
        raise ValueError("Độ lệch chuẩn sigma phải lớn hơn hoặc bằng 0.")

    rng = np.random.default_rng(seed)

    # Sinh nhiễu phân phối chuẩn N(mean, sigma^2)
    gaussian_eta = rng.normal(loc=mean, scale=sigma, size=image.shape)

    # Cộng nhiễu vào ảnh gốc (chuyển sang float64 để không bị tràn số uint8)
    noisy_float = image.astype(np.float64) + gaussian_eta

    # Cắt ngưỡng về dải hợp lệ [0, 255] và làm tròn về kiểu uint8
    noisy_clipped = np.clip(np.rint(noisy_float), 0, 255).astype(np.uint8)

    return noisy_clipped

# -*- coding: utf-8 -*-
"""
Module: dataset.py
Mô tả: Chuẩn bị bộ dữ liệu 6 ảnh xám chuẩn kích thước 256x256 đa dạng đặc trưng
       (nhiều biên sắc nét, vùng phẳng, chi tiết mịn, văn bản, kết cấu lặp).
       Hỗ trợ trích xuất từ bộ ảnh chuẩn kinh điển (scikit-image / scipy / tải chuẩn)
       và lưu vào thư mục data/original/ dưới định dạng PNG.
Nhóm thực hiện: Nhóm 10 (Nguyễn Mạnh Quyết - 20231053, Hà Thái Bảo - 20231045)
"""

import os
from typing import Dict
import numpy as np
import cv2


def _create_fallback_standard_images(size: int = 256) -> Dict[str, np.ndarray]:
    """
    Khởi tạo dự phòng nếu máy không cài sẵn scikit-image, đảm bảo luôn có đủ 6 ảnh chuẩn.
    """
    rng = np.random.default_rng(42)
    images = {}

    # 1. Checkerboard (Bàn cờ - nhiều góc nhọn và biên thẳng độ tương phản cao)
    coords = np.indices((size, size))
    checker = ((coords[0] // 32 + coords[1] // 32) % 2) * 200 + 28
    images["checkerboard"] = checker.astype(np.uint8)

    # 2. Coins / Circles (Các đồng xu tròn trên nền phẳng mịn)
    coins = np.full((size, size), 45, dtype=np.uint8)
    centers = [(60, 60, 35, 210), (180, 65, 40, 195), (120, 140, 45, 225),
               (55, 200, 30, 180), (200, 195, 38, 215)]
    for cx, cy, r, val in centers:
        cv2.circle(coins, (cx, cy), r, int(val), -1, lineType=cv2.LINE_AA)
        cv2.circle(coins, (cx, cy), r - 5, int(val - 20), 2, lineType=cv2.LINE_AA)
    images["coins"] = coins

    # 3. Cameraman-style geometric + smooth sky scene
    cam = np.zeros((size, size), dtype=np.uint8)
    for y in range(size):
        cam[y, :] = int(180 - 0.25 * y) if y < 160 else int(70 + 0.1 * (y - 160))
    cv2.rectangle(cam, (170, 80), (215, 160), 130, -1)
    cv2.fillPoly(cam, [np.array([[60, 255], [95, 110], [145, 110], [160, 255]])], 30)
    cv2.circle(cam, (120, 85), 25, 40, -1, lineType=cv2.LINE_AA)
    cv2.line(cam, (135, 115), (205, 70), 230, 3, lineType=cv2.LINE_AA)
    images["cameraman"] = cam

    # 4. Text / Page (Trang văn bản với các nét chữ mảnh, kích thước khác nhau)
    page = np.full((size, size), 235, dtype=np.uint8)
    cv2.putText(page, "EAUT - IT3243", (18, 45), cv2.FONT_HERSHEY_COMPLEX, 0.85, 20, 2, cv2.LINE_AA)
    cv2.putText(page, "IMAGE PROCESSING", (18, 85), cv2.FONT_HERSHEY_SIMPLEX, 0.7, 30, 2, cv2.LINE_AA)
    cv2.putText(page, "Spatial Domain Filtering", (18, 125), cv2.FONT_HERSHEY_SIMPLEX, 0.55, 25, 1, cv2.LINE_AA)
    cv2.putText(page, "Mean vs Median Filter", (18, 160), cv2.FONT_HERSHEY_SIMPLEX, 0.55, 25, 1, cv2.LINE_AA)
    cv2.putText(page, "Sobel Edge Preservation", (18, 195), cv2.FONT_HERSHEY_SIMPLEX, 0.55, 25, 1, cv2.LINE_AA)
    cv2.rectangle(page, (12, 12), (244, 244), 40, 2)
    images["text"] = page

    # 5. Astronaut / Portrait-like multi-scale structure
    astro = np.zeros((size, size), dtype=np.uint8)
    Y, X = np.ogrid[:size, :size]
    radial = np.sqrt((X - 128) ** 2 + (Y - 128) ** 2)
    astro[:] = np.clip(140 + 60 * np.cos(radial / 12.0), 0, 255).astype(np.uint8)
    cv2.circle(astro, (128, 128), 65, 210, -1, lineType=cv2.LINE_AA)
    cv2.circle(astro, (128, 128), 48, 55, -1, lineType=cv2.LINE_AA)
    images["astronaut"] = astro

    # 6. Brick / Texture (Tường gạch nhiều chi tiết mịn và rãnh vữa)
    brick = np.full((size, size), 210, dtype=np.uint8)
    for row in range(0, size, 32):
        offset = 0 if (row // 32) % 2 == 0 else 32
        for col in range(-32, size, 64):
            x1 = max(0, col + offset + 3)
            y1 = max(0, row + 3)
            x2 = min(size, col + offset + 61)
            y2 = min(size, row + 29)
            if x2 > x1 and y2 > y1:
                shade = int(rng.integers(110, 155))
                brick[y1:y2, x1:x2] = shade
    images["brick"] = brick

    return images


def load_or_prepare_dataset(data_dir: str = "data/original", size: int = 256) -> Dict[str, np.ndarray]:
    """
    Chuẩn bị 6 ảnh xám chuẩn kích thước 256x256 đa dạng đặc trưng:
      1. cameraman    : Ảnh kinh điển có vùng bầu trời phẳng lớn và đường biên áo/chân máy sắc nét.
      2. coins        : Ảnh các đồng xu kim loại trên nền tối đồng nhất, nhiều biên cong.
      3. astronaut    : Ảnh chân dung phi hành gia có sự chuyển tiếp độ sáng mượt và chi tiết mũ/cờ.
      4. text         : Ảnh văn bản có chữ viết nét mảnh, độ tương phản cao (nhạy với kích thước cửa sổ).
      5. checkerboard : Ảnh bàn cờ hình học có các cạnh thẳng và góc vuông sắc nét.
      6. brick        : Ảnh kết cấu bề mặt tường gạch có tính chu kỳ và chi tiết rãnh.

    Tham số:
        data_dir (str): Thư mục lưu và đọc ảnh gốc.
        size (int): Kích thước chuẩn hóa (size x size), mặc định 256x256.

    Trả về:
        Dict[str, np.ndarray]: Từ điển ánh xạ tên ảnh -> ma trận ảnh xám uint8 (256, 256).
    """
    os.makedirs(data_dir, exist_ok=True)
    target_names = ["cameraman", "coins", "astronaut", "text", "checkerboard", "brick"]
    dataset: Dict[str, np.ndarray] = {}

    # Kiểm tra nếu cả 6 ảnh đã tồn tại trên đĩa thì đọc trực tiếp bằng OpenCV
    all_exist = all(os.path.exists(os.path.join(data_dir, f"{name}.png")) for name in target_names)
    if all_exist:
        for name in target_names:
            path = os.path.join(data_dir, f"{name}.png")
            img = cv2.imread(path, cv2.IMREAD_GRAYSCALE)
            if img is not None:
                if img.shape != (size, size):
                    img = cv2.resize(img, (size, size), interpolation=cv2.INTER_AREA)
                dataset[name] = img
        if len(dataset) == len(target_names):
            return dataset

    # Nếu chưa có đủ trên đĩa, trích xuất từ kho ảnh chuẩn của scikit-image (nếu có)
    try:
        from skimage import data as skdata

        raw_sources = {
            "cameraman": skdata.camera(),
            "coins": skdata.coins(),
            "astronaut": cv2.cvtColor(skdata.astronaut(), cv2.COLOR_RGB2GRAY),
            "text": skdata.page(),
            "checkerboard": skdata.checkerboard(),
            "brick": skdata.brick(),
        }
        for name, img in raw_sources.items():
            resized = cv2.resize(img.astype(np.uint8), (size, size), interpolation=cv2.INTER_AREA)
            dataset[name] = resized
            cv2.imwrite(os.path.join(data_dir, f"{name}.png"), resized)
        return dataset

    except Exception:
        # Nếu môi trường chưa có scikit-image, sử dụng bộ tạo ảnh chuẩn dự phòng
        fallback_imgs = _create_fallback_standard_images(size=size)
        for name, img in fallback_imgs.items():
            dataset[name] = img
            cv2.imwrite(os.path.join(data_dir, f"{name}.png"), img)
        return dataset

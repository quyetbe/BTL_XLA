# -*- coding: utf-8 -*-
"""
Module: web_app.py
Mô tả: Ứng dụng Web tương tác chạy trên Localhost (http://127.0.0.1:8000)
       phục vụ trình diễn trực tiếp (Live Demo) khi bảo vệ Bài tập lớn Đề tài 7:
       "Khử nhiễu ảnh bằng các bộ lọc trong miền không gian".
       - Cho phép tải ảnh bất kỳ từ máy tính (Upload) hoặc chọn 1 trong 6 ảnh chuẩn.
       - Thêm nhiễu trực tiếp (Muối tiêu với mật độ tùy chỉnh; Gaussian với sigma tùy chỉnh).
       - Thực thi bộ lọc Trung bình / Trung vị (3x3, 5x5, 7x7) và chế độ Padding tự cài đặt.
       - Hiển thị trực quan Ảnh trước - sau (Ảnh gốc / Ảnh nhiễu / Ảnh sau lọc / Bản đồ biên Sobel).
       - Nút Reset khôi phục ảnh gốc, tải ảnh kết quả về máy và bảng đối chiếu chỉ số MSE, PSNR, Edge-F1.
Nhóm thực hiện: Nhóm 10 (Nguyễn Mạnh Quyết - 20231053, Hà Thái Bảo - 20231045)
"""

import os
import io
import time
import base64
import webbrowser
from threading import Timer
from contextlib import asynccontextmanager
from typing import Dict, Any

import numpy as np
import cv2
from PIL import Image
from fastapi import FastAPI, UploadFile, File, Form, Request
from fastapi.responses import HTMLResponse, JSONResponse
import uvicorn

from src.dataset import load_or_prepare_dataset
from src.noise import add_salt_and_pepper_noise, add_gaussian_noise
from src.filters import mean_filter, median_filter
from src.sobel import sobel_edge_detection
from src.metrics import calculate_mse, calculate_psnr, evaluate_edge_preservation

# Bộ nhớ đệm trạng thái phiên làm việc hiện tại trên giao diện Web
STATE: Dict[str, Any] = {
    "dataset": {},
    "current_name": "cameraman",
    "original_img": None,      # Ảnh xám gốc chuẩn (H, W) uint8
    "working_noisy_img": None, # Ảnh hiện tại sau khi thêm nhiễu (hoặc đang bằng ảnh gốc)
    "noise_info": "Chưa thêm nhiễu (Ảnh gốc sạch)",
}


def init_dataset() -> None:
    ds = load_or_prepare_dataset(data_dir=os.path.join("data", "original"), size=256)
    STATE["dataset"] = ds
    STATE["current_name"] = "cameraman"
    STATE["original_img"] = ds["cameraman"].copy()
    STATE["working_noisy_img"] = add_salt_and_pepper_noise(STATE["original_img"], density=0.05, seed=42)
    STATE["noise_info"] = "Nhiễu Muối tiêu (Mật độ: 5.0%, Seed=42)"


@asynccontextmanager
async def lifespan(app_instance: FastAPI):
    init_dataset()
    yield


app = FastAPI(title="BTL Xử lý ảnh - Đề tài 7 - Nhóm 10", lifespan=lifespan)


def img_to_base64(img: np.ndarray) -> str:
    """Chuyển ma trận ảnh xám NumPy uint8 sang chuỗi Data URL Base64 PNG để hiển thị trên Web."""
    pil_img = Image.fromarray(img.astype(np.uint8), mode="L")
    buf = io.BytesIO()
    pil_img.save(buf, format="PNG")
    b64 = base64.b64encode(buf.getvalue()).decode("utf-8")
    return f"data:image/png;base64,{b64}"


def compute_full_comparison(
    orig_img: np.ndarray,
    noisy_img: np.ndarray,
    kernel_size: int = 3,
    pad_mode: str = "reflect",
) -> Dict[str, Any]:
    """
    Thực hiện chạy Lọc Trung bình, Lọc Trung vị và Toán tử Sobel tự cài đặt
    trên ảnh hiện tại, trả về toàn bộ ảnh Base64 và các chỉ số MSE, PSNR, Edge-MSE, Edge-F1.
    """
    # 1. Biên Sobel trên ảnh gốc làm chuẩn tham chiếu (Ground-Truth)
    orig_grad, orig_edge_vis, orig_bin, _, ref_thresh = sobel_edge_detection(
        orig_img, threshold=None, pad_mode=pad_mode
    )

    # 2. Đánh giá ảnh nhiễu (trước khi lọc)
    noisy_mse = calculate_mse(orig_img, noisy_img)
    noisy_psnr = calculate_psnr(orig_img, noisy_img)
    noisy_grad, noisy_edge_vis, noisy_bin, _, _ = sobel_edge_detection(
        noisy_img, threshold=ref_thresh, pad_mode=pad_mode
    )
    noisy_edge_eval = evaluate_edge_preservation(orig_grad, orig_bin, noisy_grad, noisy_bin)

    # 3. Chạy Bộ lọc Trung bình (Mean Filter) tự cài đặt
    t0 = time.perf_counter()
    mean_img = mean_filter(noisy_img, kernel_size=kernel_size, pad_mode=pad_mode)
    mean_time_ms = (time.perf_counter() - t0) * 1000.0
    mean_mse = calculate_mse(orig_img, mean_img)
    mean_psnr = calculate_psnr(orig_img, mean_img)
    mean_grad, mean_edge_vis, mean_bin, _, _ = sobel_edge_detection(
        mean_img, threshold=ref_thresh, pad_mode=pad_mode
    )
    mean_edge_eval = evaluate_edge_preservation(orig_grad, orig_bin, mean_grad, mean_bin)

    # 4. Chạy Bộ lọc Trung vị (Median Filter) tự cài đặt
    t0 = time.perf_counter()
    median_img = median_filter(noisy_img, kernel_size=kernel_size, pad_mode=pad_mode)
    median_time_ms = (time.perf_counter() - t0) * 1000.0
    median_mse = calculate_mse(orig_img, median_img)
    median_psnr = calculate_psnr(orig_img, median_img)
    median_grad, median_edge_vis, median_bin, _, _ = sobel_edge_detection(
        median_img, threshold=ref_thresh, pad_mode=pad_mode
    )
    median_edge_eval = evaluate_edge_preservation(orig_grad, orig_bin, median_grad, median_bin)

    return {
        "shape": f"{orig_img.shape[1]} x {orig_img.shape[0]} px",
        "noise_info": STATE["noise_info"],
        "kernel_size": kernel_size,
        "pad_mode": pad_mode,
        "sobel_threshold": round(ref_thresh, 2),
        "images": {
            "original": img_to_base64(orig_img),
            "noisy": img_to_base64(noisy_img),
            "mean": img_to_base64(mean_img),
            "median": img_to_base64(median_img),
            "orig_edge": img_to_base64(orig_edge_vis),
            "noisy_edge": img_to_base64(noisy_edge_vis),
            "mean_edge": img_to_base64(mean_edge_vis),
            "median_edge": img_to_base64(median_edge_vis),
        },
        "metrics": {
            "noisy": {
                "mse": round(noisy_mse, 2),
                "psnr": "∞ (Ảnh gốc)" if noisy_mse == 0 else f"{noisy_psnr:.2f} dB",
                "edge_mse": round(noisy_edge_eval["edge_mse"], 2),
                "precision": round(noisy_edge_eval["precision"], 4),
                "recall": round(noisy_edge_eval["recall"], 4),
                "f1": round(noisy_edge_eval["f1"], 4),
                "time_ms": "-",
            },
            "mean": {
                "mse": round(mean_mse, 2),
                "psnr": "∞" if mean_mse == 0 else f"{mean_psnr:.2f} dB",
                "psnr_val": 99.0 if mean_mse == 0 else round(mean_psnr, 2),
                "edge_mse": round(mean_edge_eval["edge_mse"], 2),
                "precision": round(mean_edge_eval["precision"], 4),
                "recall": round(mean_edge_eval["recall"], 4),
                "f1": round(mean_edge_eval["f1"], 4),
                "time_ms": f"{mean_time_ms:.2f} ms",
            },
            "median": {
                "mse": round(median_mse, 2),
                "psnr": "∞" if median_mse == 0 else f"{median_psnr:.2f} dB",
                "psnr_val": 99.0 if median_mse == 0 else round(median_psnr, 2),
                "edge_mse": round(median_edge_eval["edge_mse"], 2),
                "precision": round(median_edge_eval["precision"], 4),
                "recall": round(median_edge_eval["recall"], 4),
                "f1": round(median_edge_eval["f1"], 4),
                "time_ms": f"{median_time_ms:.2f} ms",
            },
        },
    }


@app.post("/api/select_sample")
async def select_sample(
    sample_name: str = Form("cameraman"),
    noise_type: str = Form("salt_pepper"),
    sp_density: float = Form(5.0),
    gauss_sigma: float = Form(25.0),
    kernel_size: int = Form(3),
    pad_mode: str = Form("reflect"),
    seed: int = Form(42),
):
    if sample_name in STATE["dataset"]:
        STATE["current_name"] = sample_name
        STATE["original_img"] = STATE["dataset"][sample_name].copy()
    return await process_pipeline(noise_type, sp_density, gauss_sigma, kernel_size, pad_mode, seed)


@app.post("/api/upload_image")
async def upload_image(
    file: UploadFile = File(...),
    resize_mode: str = Form("256"),
    noise_type: str = Form("salt_pepper"),
    sp_density: float = Form(5.0),
    gauss_sigma: float = Form(25.0),
    kernel_size: int = Form(3),
    pad_mode: str = Form("reflect"),
    seed: int = Form(42),
):
    raw_bytes = await file.read()
    pil_img = Image.open(io.BytesIO(raw_bytes)).convert("L")
    img_np = np.array(pil_img, dtype=np.uint8)

    if resize_mode == "256":
        img_np = cv2.resize(img_np, (256, 256), interpolation=cv2.INTER_AREA)
    elif resize_mode == "512":
        img_np = cv2.resize(img_np, (512, 512), interpolation=cv2.INTER_AREA)
    else:
        # Giữ nguyên tỷ lệ nhưng giới hạn cạnh lớn nhất <= 512 để xử lý mượt
        h, w = img_np.shape
        if max(h, w) > 512:
            scale = 512.0 / max(h, w)
            img_np = cv2.resize(img_np, (int(w * scale), int(h * scale)), interpolation=cv2.INTER_AREA)

    STATE["current_name"] = file.filename or "uploaded_image.png"
    STATE["original_img"] = img_np.copy()
    return await process_pipeline(noise_type, sp_density, gauss_sigma, kernel_size, pad_mode, seed)


@app.post("/api/process")
async def process_pipeline(
    noise_type: str = Form("salt_pepper"),
    sp_density: float = Form(5.0),
    gauss_sigma: float = Form(25.0),
    kernel_size: int = Form(3),
    pad_mode: str = Form("reflect"),
    seed: int = Form(42),
):
    orig = STATE["original_img"]
    if noise_type == "none":
        STATE["working_noisy_img"] = orig.copy()
        STATE["noise_info"] = "Không thêm nhiễu (Ảnh gốc sạch)"
    elif noise_type == "salt_pepper":
        density_val = max(0.0, min(100.0, float(sp_density))) / 100.0
        STATE["working_noisy_img"] = add_salt_and_pepper_noise(orig, density=density_val, seed=int(seed))
        STATE["noise_info"] = f"Nhiễu Muối tiêu (Mật độ: {sp_density:.1f}%, Seed={seed})"
    elif noise_type == "gaussian":
        sigma_val = max(0.0, float(gauss_sigma))
        STATE["working_noisy_img"] = add_gaussian_noise(orig, sigma=sigma_val, seed=int(seed))
        STATE["noise_info"] = f"Nhiễu Gaussian (Độ lệch chuẩn σ = {sigma_val:.1f}, Seed={seed})"
    elif noise_type == "both":
        density_val = max(0.0, min(100.0, float(sp_density))) / 100.0
        sigma_val = max(0.0, float(gauss_sigma))
        temp = add_gaussian_noise(orig, sigma=sigma_val, seed=int(seed))
        STATE["working_noisy_img"] = add_salt_and_pepper_noise(temp, density=density_val, seed=int(seed) + 1)
        STATE["noise_info"] = f"Nhiễu Hỗn hợp (Gaussian σ={sigma_val:.1f} + Muối tiêu {sp_density:.1f}%)"

    data = compute_full_comparison(
        STATE["original_img"],
        STATE["working_noisy_img"],
        kernel_size=int(kernel_size),
        pad_mode=pad_mode,
    )
    data["image_name"] = STATE["current_name"]
    return JSONResponse(data)


@app.post("/api/reset")
async def reset_image():
    """Reset khôi phục trạng thái ảnh về ảnh gốc sạch ban đầu."""
    orig = STATE["original_img"]
    STATE["working_noisy_img"] = orig.copy()
    STATE["noise_info"] = "Đã Reset về Ảnh gốc sạch (0% nhiễu)"
    data = compute_full_comparison(orig, orig, kernel_size=3, pad_mode="reflect")
    data["image_name"] = STATE["current_name"]
    return JSONResponse(data)


HTML_PAGE = """<!DOCTYPE html>
<html lang="vi">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>BTL Xử Lý Ảnh - Đề tài 7 - Nhóm 10 (EAUT)</title>
  <style>
    :root {
      --bg: #0f172a;
      --panel: #1e293b;
      --card: #334155;
      --border: #475569;
      --primary: #38bdf8;
      --primary-hover: #0284c7;
      --success: #22c55e;
      --warning: #f59e0b;
      --danger: #ef4444;
      --text: #f8fafc;
      --muted: #cbd5e1;
    }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: 'Segoe UI', Tahoma, Arial, sans-serif;
      background: var(--bg);
      color: var(--text);
      line-height: 1.5;
    }
    header {
      background: linear-gradient(90deg, #1e3a8a, #0f172a);
      border-bottom: 1px solid var(--border);
      padding: 14px 24px;
      display: flex;
      justify-content: space-between;
      align-items: center;
      flex-wrap: wrap;
      gap: 10px;
    }
    .header-title h1 {
      font-size: 1.25rem;
      color: #fff;
      letter-spacing: 0.3px;
    }
    .header-title p {
      font-size: 0.85rem;
      color: var(--primary);
    }
    .badge-group {
      display: flex;
      gap: 8px;
      flex-wrap: wrap;
    }
    .badge {
      background: rgba(56, 189, 248, 0.15);
      border: 1px solid var(--primary);
      color: var(--primary);
      padding: 4px 10px;
      border-radius: 999px;
      font-size: 0.8rem;
      font-weight: 600;
    }
    .container {
      display: grid;
      grid-template-columns: 340px 1fr;
      gap: 18px;
      padding: 18px;
      max-width: 1650px;
      margin: 0 auto;
    }
    @media (max-width: 1100px) {
      .container { grid-template-columns: 1fr; }
    }
    .sidebar {
      background: var(--panel);
      border: 1px solid var(--border);
      border-radius: 12px;
      padding: 16px;
      display: flex;
      flex-direction: column;
      gap: 16px;
      height: fit-content;
    }
    .section-box {
      background: rgba(15, 23, 42, 0.55);
      border: 1px solid var(--border);
      border-radius: 8px;
      padding: 12px;
    }
    .section-box h3 {
      font-size: 0.95rem;
      color: var(--primary);
      margin-bottom: 10px;
      border-bottom: 1px solid rgba(71, 85, 105, 0.6);
      padding-bottom: 5px;
    }
    label {
      display: block;
      font-size: 0.84rem;
      color: var(--muted);
      margin-bottom: 4px;
      margin-top: 8px;
    }
    select, input[type="number"], input[type="file"] {
      width: 100%;
      padding: 8px 10px;
      background: #0f172a;
      border: 1px solid var(--border);
      color: #fff;
      border-radius: 6px;
      font-size: 0.88rem;
    }
    input[type="range"] {
      width: 100%;
      accent-color: var(--primary);
      cursor: pointer;
    }
    .range-row {
      display: flex;
      justify-content: space-between;
      font-size: 0.82rem;
      color: var(--primary);
      font-weight: 600;
    }
    .btn-group {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 8px;
      margin-top: 10px;
    }
    button {
      padding: 10px 12px;
      border: none;
      border-radius: 6px;
      font-weight: 600;
      font-size: 0.88rem;
      cursor: pointer;
      transition: all 0.15s ease;
    }
    .btn-primary {
      background: var(--primary);
      color: #0f172a;
    }
    .btn-primary:hover { background: #7dd3fc; }
    .btn-danger {
      background: var(--danger);
      color: #fff;
    }
    .btn-danger:hover { background: #dc2626; }
    .btn-outline {
      background: transparent;
      border: 1px solid var(--border);
      color: var(--text);
    }
    .btn-outline:hover {
      border-color: var(--primary);
      color: var(--primary);
    }
    .preset-pills {
      display: flex;
      flex-wrap: wrap;
      gap: 6px;
      margin-top: 6px;
    }
    .pill {
      padding: 5px 9px;
      font-size: 0.78rem;
      background: #1e293b;
      border: 1px solid var(--border);
      color: var(--muted);
      border-radius: 5px;
      cursor: pointer;
    }
    .pill:hover, .pill.active {
      background: var(--primary);
      color: #0f172a;
      border-color: var(--primary);
      font-weight: 700;
    }
    .main-area {
      display: flex;
      flex-direction: column;
      gap: 16px;
    }
    .status-bar {
      background: var(--panel);
      border: 1px solid var(--border);
      border-radius: 10px;
      padding: 12px 16px;
      display: flex;
      justify-content: space-between;
      align-items: center;
      flex-wrap: wrap;
      gap: 12px;
    }
    .status-bar .info-item {
      font-size: 0.88rem;
    }
    .status-bar .info-item strong {
      color: var(--primary);
    }
    .tabs {
      display: flex;
      gap: 8px;
    }
    .tab-btn {
      padding: 7px 14px;
      border-radius: 6px;
      background: #0f172a;
      border: 1px solid var(--border);
      color: var(--muted);
      font-size: 0.85rem;
    }
    .tab-btn.active {
      background: var(--primary);
      color: #0f172a;
      border-color: var(--primary);
    }
    .grid-4 {
      display: grid;
      grid-template-columns: repeat(4, 1fr);
      gap: 14px;
    }
    @media (max-width: 1350px) {
      .grid-4 { grid-template-columns: repeat(2, 1fr); }
    }
    @media (max-width: 650px) {
      .grid-4 { grid-template-columns: 1fr; }
    }
    .img-card {
      background: var(--panel);
      border: 1px solid var(--border);
      border-radius: 10px;
      padding: 10px;
      display: flex;
      flex-direction: column;
      align-items: center;
      position: relative;
    }
    .img-card.highlight {
      border: 2px solid var(--success);
      box-shadow: 0 0 12px rgba(34, 197, 94, 0.2);
    }
    .img-card h4 {
      font-size: 0.92rem;
      margin-bottom: 6px;
      text-align: center;
    }
    .img-wrapper {
      width: 100%;
      aspect-ratio: 1 / 1;
      background: #090d16;
      border-radius: 6px;
      overflow: hidden;
      display: flex;
      align-items: center;
      justify-content: center;
      border: 1px solid #334155;
    }
    .img-wrapper img {
      width: 100%;
      height: 100%;
      object-fit: contain;
      image-rendering: pixelated;
    }
    .metric-strip {
      width: 100%;
      margin-top: 8px;
      padding: 6px 8px;
      background: rgba(15, 23, 42, 0.7);
      border-radius: 6px;
      font-size: 0.8rem;
      display: flex;
      justify-content: space-between;
      color: var(--muted);
    }
    .metric-strip strong { color: #fff; }
    .download-link {
      margin-top: 6px;
      font-size: 0.76rem;
      color: var(--primary);
      text-decoration: none;
    }
    .download-link:hover { text-decoration: underline; }
    .table-card {
      background: var(--panel);
      border: 1px solid var(--border);
      border-radius: 10px;
      padding: 14px;
      overflow-x: auto;
    }
    .table-card h3 {
      font-size: 1rem;
      color: var(--primary);
      margin-bottom: 10px;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      font-size: 0.88rem;
    }
    th, td {
      border: 1px solid var(--border);
      padding: 8px 10px;
      text-align: center;
    }
    th {
      background: #0f172a;
      color: var(--primary);
    }
    tr:nth-child(even) { background: rgba(15, 23, 42, 0.4); }
    .winner-badge {
      background: rgba(34, 197, 94, 0.2);
      color: #4ade80;
      padding: 2px 8px;
      border-radius: 4px;
      font-weight: 700;
      font-size: 0.78rem;
    }
  </style>
</head>
<body>
  <header>
    <div class="header-title">
      <h1>HỆ THỐNG KHỬ NHIỄU ẢNH TRONG MIỀN KHÔNG GIAN & PHÁT HIỆN BIÊN SOBEL</h1>
      <p>Học phần: Xử lý ảnh và thị giác máy tính (IT3243) - Đề tài 7 | GVHD: Thầy Nguyễn Phồn Lữa</p>
    </div>
    <div class="badge-group">
      <span class="badge">Nhóm 10 - DCCNTT.14.3</span>
      <span class="badge">Nguyễn Mạnh Quyết (20231053)</span>
      <span class="badge">Hà Thái Bảo (20231045)</span>
    </div>
  </header>

  <div class="container">
    <!-- CỘT ĐIỀU KHIỂN BÊN TRÁI -->
    <aside class="sidebar">
      <!-- 1. Chọn hoặc Tải ảnh -->
      <div class="section-box">
        <h3>1. Nguồn ảnh đầu vào (Input Image)</h3>
        <label for="sampleSelect">Chọn từ 6 ảnh chuẩn thực nghiệm:</label>
        <select id="sampleSelect" onchange="onSelectSample()">
          <option value="cameraman">1. Cameraman (Biên sắc & nền trời phẳng)</option>
          <option value="coins">2. Coins (Đồng xu tròn trên nền tối)</option>
          <option value="astronaut">3. Astronaut (Chân dung đa chi tiết)</option>
          <option value="text">4. Text / Page (Văn bản nét chữ mảnh)</option>
          <option value="checkerboard">5. Checkerboard (Bàn cờ góc vuông)</option>
          <option value="brick">6. Brick (Kết cấu tường gạch lặp)</option>
        </select>

        <label for="fileUpload" style="margin-top:12px;">Hoặc tải ảnh trực tiếp từ máy tính:</label>
        <input type="file" id="fileUpload" accept="image/*" onchange="onUploadFile()" />
        <label for="resizeMode">Chuẩn hóa kích thước khi tải ảnh:</label>
        <select id="resizeMode">
          <option value="256" selected>Chuẩn 256 x 256 px (Nhanh & Chuẩn BTL)</option>
          <option value="512">Chuẩn 512 x 512 px (Độ phân giải cao)</option>
          <option value="keep">Giữ tỷ lệ gốc (Tối đa 512 px)</option>
        </select>
      </div>

      <!-- 2. Cấu hình Nhiễu -->
      <div class="section-box">
        <h3>2. Giả lập Nhiễu (Noise Generator)</h3>
        <label for="noiseType">Loại nhiễu bổ sung vào ảnh:</label>
        <select id="noiseType" onchange="updateNoiseControls()">
          <option value="salt_pepper" selected>Nhiễu Muối tiêu (Salt & Pepper)</option>
          <option value="gaussian">Nhiễu Cộng Gaussian (AWGN)</option>
          <option value="both">Nhiễu Hỗn hợp (Gaussian + Muối tiêu)</option>
          <option value="none">Không thêm nhiễu (Giữ nguyên ảnh gốc)</option>
        </select>

        <div id="spControl">
          <label>
            <div class="range-row">
              <span>Mật độ Muối tiêu (%):</span>
              <span id="spValLabel">5.0%</span>
            </div>
          </label>
          <input type="range" id="spDensity" min="0.5" max="30" step="0.5" value="5.0"
                 oninput="document.getElementById('spValLabel').innerText=this.value+'%'; runProcess();" />
          <div class="preset-pills">
            <span class="pill" onclick="setSp(1.0)">Mức đề bài: 1%</span>
            <span class="pill active" onclick="setSp(5.0)">Mức đề bài: 5%</span>
            <span class="pill" onclick="setSp(10.0)">Mức đề bài: 10%</span>
            <span class="pill" onclick="setSp(20.0)">Nặng: 20%</span>
          </div>
        </div>

        <div id="gaussControl" style="display:none;">
          <label>
            <div class="range-row">
              <span>Độ lệch chuẩn Gaussian (σ):</span>
              <span id="gaussValLabel">σ = 25</span>
            </div>
          </label>
          <input type="range" id="gaussSigma" min="2" max="80" step="1" value="25"
                 oninput="document.getElementById('gaussValLabel').innerText='σ = '+this.value; runProcess();" />
          <div class="preset-pills">
            <span class="pill" onclick="setGauss(10)">Mức đề bài: σ=10</span>
            <span class="pill active" onclick="setGauss(25)">Mức đề bài: σ=25</span>
            <span class="pill" onclick="setGauss(50)">Mức đề bài: σ=50</span>
          </div>
        </div>

        <label for="seedInput">Random Seed (Cố định để tái lập):</label>
        <input type="number" id="seedInput" value="42" onchange="runProcess()" />
      </div>

      <!-- 3. Cấu hình Bộ lọc & Padding -->
      <div class="section-box">
        <h3>3. Tham số Bộ lọc Không gian (Tự cài đặt)</h3>
        <label>Kích thước cửa sổ trượt (Kernel Size k × k):</label>
        <div class="preset-pills" id="kernelPills">
          <span class="pill active" data-k="3" onclick="setKernel(3)">Cửa sổ 3 × 3</span>
          <span class="pill" data-k="5" onclick="setKernel(5)">Cửa sổ 5 × 5</span>
          <span class="pill" data-k="7" onclick="setKernel(7)">Cửa sổ 7 × 7</span>
        </div>
        <input type="hidden" id="kernelSize" value="3" />

        <label for="padMode" style="margin-top:10px;">Phương pháp xử lý biên (Padding):</label>
        <select id="padMode" onchange="runProcess()">
          <option value="reflect" selected>Reflect (Phản chiếu gương - Khuyên dùng)</option>
          <option value="replicate">Replicate (Lặp lại điểm ảnh mép biên)</option>
          <option value="zero">Zero-Padding (Đệm giá trị 0 quanh viền)</option>
        </select>

        <div class="btn-group">
          <button class="btn-primary" onclick="runProcess()">⚡ Chạy Lọc & Sobel</button>
          <button class="btn-danger" onclick="resetImage()">🔄 Reset Ảnh Gốc</button>
        </div>
      </div>
    </aside>

    <!-- KHU VỰC HIỂN THỊ TRỰC QUAN BÊN PHẢI -->
    <main class="main-area">
      <!-- Thanh trạng thái -->
      <div class="status-bar">
        <div class="info-item">Ảnh đang xử lý: <strong id="lblImgName">cameraman</strong> (<span id="lblShape">256 x 256 px</span>)</div>
        <div class="info-item">Trạng thái nhiễu: <strong id="lblNoiseInfo">...</strong></div>
        <div class="info-item">Ngưỡng biên Sobel chuẩn (P85): <strong id="lblThreshold">...</strong></div>
      </div>

      <!-- Hàng 1: So sánh trực quan Ảnh Trước & Sau Lọc -->
      <div class="grid-4">
        <div class="img-card">
          <h4>1. Ảnh Gốc (Original)</h4>
          <div class="img-wrapper"><img id="imgOrig" src="" alt="Original" /></div>
          <div class="metric-strip"><span>Chuẩn tham chiếu</span><strong>Ground-Truth</strong></div>
          <a id="dlOrig" class="download-link" download="1_original.png">⬇ Tải ảnh gốc</a>
        </div>

        <div class="img-card">
          <h4 style="color:#f87171;">2. Ảnh Bị Nhiễu (Trước khi lọc)</h4>
          <div class="img-wrapper"><img id="imgNoisy" src="" alt="Noisy" /></div>
          <div class="metric-strip">
            <span>MSE: <strong id="mNoisyMse">-</strong></span>
            <span>PSNR: <strong id="mNoisyPsnr">-</strong></span>
          </div>
          <a id="dlNoisy" class="download-link" download="2_noisy.png">⬇ Tải ảnh nhiễu</a>
        </div>

        <div class="img-card" id="cardMean">
          <h4 style="color:#60a5fa;">3. Sau Lọc Trung Bình (<span class="lblK">3x3</span>)</h4>
          <div class="img-wrapper"><img id="imgMean" src="" alt="Mean Filtered" /></div>
          <div class="metric-strip">
            <span>MSE: <strong id="mMeanMse">-</strong></span>
            <span>PSNR: <strong id="mMeanPsnr">-</strong></span>
          </div>
          <a id="dlMean" class="download-link" download="3_mean_filtered.png">⬇ Tải ảnh lọc Trung bình</a>
        </div>

        <div class="img-card" id="cardMedian">
          <h4 style="color:#4ade80;">4. Sau Lọc Trung Vị (<span class="lblK">3x3</span>)</h4>
          <div class="img-wrapper"><img id="imgMedian" src="" alt="Median Filtered" /></div>
          <div class="metric-strip">
            <span>MSE: <strong id="mMedianMse">-</strong></span>
            <span>PSNR: <strong id="mMedianPsnr">-</strong></span>
          </div>
          <a id="dlMedian" class="download-link" download="4_median_filtered.png">⬇ Tải ảnh lọc Trung vị</a>
        </div>
      </div>

      <!-- Hàng 2: So sánh trực quan Bản đồ Biên Sobel -->
      <div class="grid-4">
        <div class="img-card">
          <h4>5. Biên Sobel Ảnh Gốc</h4>
          <div class="img-wrapper"><img id="imgOrigEdge" src="" alt="Original Sobel" /></div>
          <div class="metric-strip"><span>Bản đồ Gradient</span><strong>|G| = √(Gx²+Gy²)</strong></div>
          <a id="dlOrigEdge" class="download-link" download="5_sobel_original.png">⬇ Tải biên gốc</a>
        </div>

        <div class="img-card">
          <h4 style="color:#f87171;">6. Biên Sobel Ảnh Nhiễu</h4>
          <div class="img-wrapper"><img id="imgNoisyEdge" src="" alt="Noisy Sobel" /></div>
          <div class="metric-strip">
            <span>E-MSE: <strong id="mNoisyEmse">-</strong></span>
            <span>Edge-F1: <strong id="mNoisyF1">-</strong></span>
          </div>
          <a id="dlNoisyEdge" class="download-link" download="6_sobel_noisy.png">⬇ Tải biên nhiễu</a>
        </div>

        <div class="img-card">
          <h4 style="color:#60a5fa;">7. Biên Sobel sau Lọc Trung Bình</h4>
          <div class="img-wrapper"><img id="imgMeanEdge" src="" alt="Mean Sobel" /></div>
          <div class="metric-strip">
            <span>E-MSE: <strong id="mMeanEmse">-</strong></span>
            <span>Edge-F1: <strong id="mMeanF1">-</strong></span>
          </div>
          <a id="dlMeanEdge" class="download-link" download="7_sobel_mean.png">⬇ Tải biên lọc Trung bình</a>
        </div>

        <div class="img-card">
          <h4 style="color:#4ade80;">8. Biên Sobel sau Lọc Trung Vị</h4>
          <div class="img-wrapper"><img id="imgMedianEdge" src="" alt="Median Sobel" /></div>
          <div class="metric-strip">
            <span>E-MSE: <strong id="mMedianEmse">-</strong></span>
            <span>Edge-F1: <strong id="mMedianF1">-</strong></span>
          </div>
          <a id="dlMedianEdge" class="download-link" download="8_sobel_median.png">⬇ Tải biên lọc Trung vị</a>
        </div>
      </div>

      <!-- Bảng tổng hợp số liệu trực tiếp -->
      <div class="table-card">
        <h3>BẢNG ĐỐI SÁNH ĐỊNH LƯỢNG TRỰC TIẾP (MSE, PSNR, BẢO TOÀN BIÊN SOBEL & THỜI GIAN CHẠY)</h3>
        <table>
          <thead>
            <tr>
              <th>Trạng thái / Phương pháp</th>
              <th>Cửa sổ lọc</th>
              <th>MSE (Càng nhỏ càng tốt)</th>
              <th>PSNR (Càng cao càng tốt)</th>
              <th>Edge-MSE (Sai số Gradient)</th>
              <th>Edge Precision (Độ chính xác biên)</th>
              <th>Edge Recall (Độ phủ biên)</th>
              <th>Edge F1-Score (Bảo toàn biên)</th>
              <th>Thời gian xử lý</th>
              <th>Đánh giá tự động</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td><strong>Ảnh nhiễu (Trước khi lọc)</strong></td>
              <td>-</td>
              <td id="tNoisyMse">-</td>
              <td id="tNoisyPsnr">-</td>
              <td id="tNoisyEmse">-</td>
              <td id="tNoisyPrec">-</td>
              <td id="tNoisyRec">-</td>
              <td id="tNoisyF1">-</td>
              <td>-</td>
              <td>Ảnh suy giảm</td>
            </tr>
            <tr>
              <td><strong style="color:#60a5fa;">Bộ lọc Trung bình (Mean Filter)</strong></td>
              <td class="lblK">3x3</td>
              <td id="tMeanMse">-</td>
              <td id="tMeanPsnr">-</td>
              <td id="tMeanEmse">-</td>
              <td id="tMeanPrec">-</td>
              <td id="tMeanRec">-</td>
              <td id="tMeanF1">-</td>
              <td id="tMeanTime">-</td>
              <td id="tMeanEval">-</td>
            </tr>
            <tr>
              <td><strong style="color:#4ade80;">Bộ lọc Trung vị (Median Filter)</strong></td>
              <td class="lblK">3x3</td>
              <td id="tMedianMse">-</td>
              <td id="tMedianPsnr">-</td>
              <td id="tMedianEmse">-</td>
              <td id="tMedianPrec">-</td>
              <td id="tMedianRec">-</td>
              <td id="tMedianF1">-</td>
              <td id="tMedianTime">-</td>
              <td id="tMedianEval">-</td>
            </tr>
          </tbody>
        </table>
      </div>
    </main>
  </div>

  <script>
    function updateNoiseControls() {
      const nt = document.getElementById('noiseType').value;
      document.getElementById('spControl').style.display = (nt === 'salt_pepper' || nt === 'both') ? 'block' : 'none';
      document.getElementById('gaussControl').style.display = (nt === 'gaussian' || nt === 'both') ? 'block' : 'none';
      runProcess();
    }

    function setSp(val) {
      document.getElementById('spDensity').value = val;
      document.getElementById('spValLabel').innerText = val.toFixed(1) + '%';
      runProcess();
    }

    function setGauss(val) {
      document.getElementById('gaussSigma').value = val;
      document.getElementById('gaussValLabel').innerText = 'σ = ' + val;
      runProcess();
    }

    function setKernel(k) {
      document.getElementById('kernelSize').value = k;
      document.querySelectorAll('#kernelPills .pill').forEach(el => {
        el.classList.toggle('active', parseInt(el.dataset.k) === k);
      });
      document.querySelectorAll('.lblK').forEach(el => el.innerText = k + 'x' + k);
      runProcess();
    }

    function buildFormData() {
      const fd = new FormData();
      fd.append('sample_name', document.getElementById('sampleSelect').value);
      fd.append('noise_type', document.getElementById('noiseType').value);
      fd.append('sp_density', document.getElementById('spDensity').value);
      fd.append('gauss_sigma', document.getElementById('gaussSigma').value);
      fd.append('kernel_size', document.getElementById('kernelSize').value);
      fd.append('pad_mode', document.getElementById('padMode').value);
      fd.append('seed', document.getElementById('seedInput').value);
      return fd;
    }

    async function onSelectSample() {
      const fd = buildFormData();
      const res = await fetch('/api/select_sample', { method: 'POST', body: fd });
      const data = await res.json();
      renderData(data);
    }

    async function onUploadFile() {
      const fileInput = document.getElementById('fileUpload');
      if (!fileInput.files.length) return;
      const fd = buildFormData();
      fd.append('file', fileInput.files[0]);
      fd.append('resize_mode', document.getElementById('resizeMode').value);
      const res = await fetch('/api/upload_image', { method: 'POST', body: fd });
      const data = await res.json();
      renderData(data);
    }

    async function runProcess() {
      const fd = buildFormData();
      const res = await fetch('/api/process', { method: 'POST', body: fd });
      const data = await res.json();
      renderData(data);
    }

    async function resetImage() {
      document.getElementById('noiseType').value = 'none';
      document.getElementById('spControl').style.display = 'none';
      document.getElementById('gaussControl').style.display = 'none';
      const res = await fetch('/api/reset', { method: 'POST' });
      const data = await res.json();
      renderData(data);
    }

    function renderData(data) {
      document.getElementById('lblImgName').innerText = data.image_name;
      document.getElementById('lblShape').innerText = data.shape;
      document.getElementById('lblNoiseInfo').innerText = data.noise_info;
      document.getElementById('lblThreshold').innerText = data.sobel_threshold;

      // Cập nhật 8 khung hình và nút tải về
      const mapImgs = [
        ['imgOrig', 'dlOrig', data.images.original],
        ['imgNoisy', 'dlNoisy', data.images.noisy],
        ['imgMean', 'dlMean', data.images.mean],
        ['imgMedian', 'dlMedian', data.images.median],
        ['imgOrigEdge', 'dlOrigEdge', data.images.orig_edge],
        ['imgNoisyEdge', 'dlNoisyEdge', data.images.noisy_edge],
        ['imgMeanEdge', 'dlMeanEdge', data.images.mean_edge],
        ['imgMedianEdge', 'dlMedianEdge', data.images.median_edge],
      ];
      mapImgs.forEach(([imgId, dlId, b64]) => {
        document.getElementById(imgId).src = b64;
        document.getElementById(dlId).href = b64;
      });

      const m = data.metrics;
      // Strip metrics
      document.getElementById('mNoisyMse').innerText = m.noisy.mse;
      document.getElementById('mNoisyPsnr').innerText = m.noisy.psnr;
      document.getElementById('mMeanMse').innerText = m.mean.mse;
      document.getElementById('mMeanPsnr').innerText = m.mean.psnr;
      document.getElementById('mMedianMse').innerText = m.median.mse;
      document.getElementById('mMedianPsnr').innerText = m.median.psnr;

      document.getElementById('mNoisyEmse').innerText = m.noisy.edge_mse;
      document.getElementById('mNoisyF1').innerText = m.noisy.f1;
      document.getElementById('mMeanEmse').innerText = m.mean.edge_mse;
      document.getElementById('mMeanF1').innerText = m.mean.f1;
      document.getElementById('mMedianEmse').innerText = m.median.edge_mse;
      document.getElementById('mMedianF1').innerText = m.median.f1;

      // Table metrics
      document.getElementById('tNoisyMse').innerText = m.noisy.mse;
      document.getElementById('tNoisyPsnr').innerText = m.noisy.psnr;
      document.getElementById('tNoisyEmse').innerText = m.noisy.edge_mse;
      document.getElementById('tNoisyPrec').innerText = m.noisy.precision;
      document.getElementById('tNoisyRec').innerText = m.noisy.recall;
      document.getElementById('tNoisyF1').innerText = m.noisy.f1;

      document.getElementById('tMeanMse').innerText = m.mean.mse;
      document.getElementById('tMeanPsnr').innerText = m.mean.psnr;
      document.getElementById('tMeanEmse').innerText = m.mean.edge_mse;
      document.getElementById('tMeanPrec').innerText = m.mean.precision;
      document.getElementById('tMeanRec').innerText = m.mean.recall;
      document.getElementById('tMeanF1').innerText = m.mean.f1;
      document.getElementById('tMeanTime').innerText = m.mean.time_ms;

      document.getElementById('tMedianMse').innerText = m.median.mse;
      document.getElementById('tMedianPsnr').innerText = m.median.psnr;
      document.getElementById('tMedianEmse').innerText = m.median.edge_mse;
      document.getElementById('tMedianPrec').innerText = m.median.precision;
      document.getElementById('tMedianRec').innerText = m.median.recall;
      document.getElementById('tMedianF1').innerText = m.median.f1;
      document.getElementById('tMedianTime').innerText = m.median.time_ms;

      // Highlight bộ lọc tốt hơn theo PSNR
      const meanWin = m.mean.psnr_val > m.median.psnr_val;
      const medWin = m.median.psnr_val > m.mean.psnr_val;
      document.getElementById('cardMean').classList.toggle('highlight', meanWin);
      document.getElementById('cardMedian').classList.toggle('highlight', medWin);
      document.getElementById('tMeanEval').innerHTML = meanWin ? '<span class="winner-badge">★ Tối ưu PSNR hơn</span>' : 'Nhòe biên hơn';
      document.getElementById('tMedianEval').innerHTML = medWin ? '<span class="winner-badge">★ Tối ưu PSNR & Biên</span>' : 'Chậm hơn';
    }

    // Khởi chạy lần đầu khi mở trang (hỗ trợ query params cho chụp ảnh tự động)
    window.addEventListener('DOMContentLoaded', () => {
      const params = new URLSearchParams(window.location.search);
      if (params.get('sample')) {
        document.getElementById('sampleSelect').value = params.get('sample');
      }
      if (params.get('noise')) {
        document.getElementById('noiseType').value = params.get('noise');
      }
      if (params.get('sp')) {
        document.getElementById('spDensity').value = params.get('sp');
        document.getElementById('spValLabel').innerText = parseFloat(params.get('sp')).toFixed(1) + '%';
      }
      if (params.get('sigma')) {
        document.getElementById('gaussSigma').value = params.get('sigma');
        document.getElementById('gaussValLabel').innerText = 'σ = ' + params.get('sigma');
      }
      if (params.get('k')) {
        const kVal = parseInt(params.get('k'));
        document.getElementById('kernelSize').value = kVal;
        document.querySelectorAll('#kernelPills .pill').forEach(el => {
          el.classList.toggle('active', parseInt(el.dataset.k) === kVal);
        });
        document.querySelectorAll('.lblK').forEach(el => el.innerText = kVal + 'x' + kVal);
      }
      if (params.get('reset') === '1') {
        resetImage();
      } else {
        const nt = document.getElementById('noiseType').value;
        document.getElementById('spControl').style.display = (nt === 'salt_pepper' || nt === 'both') ? 'block' : 'none';
        document.getElementById('gaussControl').style.display = (nt === 'gaussian' || nt === 'both') ? 'block' : 'none';
        onSelectSample();
      }
    });
  </script>
</body>
</html>
"""


@app.get("/", response_class=HTMLResponse)
async def index():
    return HTML_PAGE


def open_browser():
    webbrowser.open("http://127.0.0.1:8000")


if __name__ == "__main__":
    print("=" * 75)
    print("KHỞI ĐỘNG WEB APP DEMO KHỬ NHIỄU ẢNH & BIÊN SOBEL - NHÓM 10")
    print("Địa chỉ truy cập Localhost: http://127.0.0.1:8000")
    print("=" * 75)
    Timer(1.2, open_browser).start()
    uvicorn.run(app, host="127.0.0.1", port=8000, log_level="info")

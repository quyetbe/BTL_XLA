# -*- coding: utf-8 -*-
"""
Script chính: run_experiments.py
Mô tả: Chạy toàn bộ kịch bản thực nghiệm cho Đề tài 7:
       "Khử nhiễu ảnh bằng các bộ lọc trong miền không gian"
       - Kiểm chứng độ chính xác của code tự cài đặt với OpenCV.
       - Thực nghiệm trên 6 ảnh chuẩn x 2 loại nhiễu (Muối tiêu 1%, 5%, 10%; Gaussian sigma 10, 25, 50)
         x 2 bộ lọc (Trung bình, Trung vị) x 3 kích thước cửa sổ (3x3, 5x5, 7x7).
       - Đánh giá MSE, PSNR, khả năng bảo toàn biên với Sobel (Edge-MSE, Edge-Precision, Recall, F1)
         và thời gian xử lý trung bình (ms).
       - Xuất bảng số liệu ra file CSV và lưu các hình ảnh trực quan, biểu đồ.
Nhóm thực hiện: Nhóm 10 (Nguyễn Mạnh Quyết - 20231053, Hà Thái Bảo - 20231045)
"""

import os
import csv
import time
from typing import List, Dict, Any
import numpy as np

from src.dataset import load_or_prepare_dataset
from src.noise import add_salt_and_pepper_noise, add_gaussian_noise
from src.filters import mean_filter, median_filter
from src.sobel import sobel_edge_detection
from src.metrics import calculate_mse, calculate_psnr, evaluate_edge_preservation
from src.verification import verify_implementations
from src.visualization import (
    plot_dataset_overview,
    plot_denoising_comparison,
    plot_sobel_edge_comparison,
    plot_psnr_and_edge_charts,
)


def main() -> None:
    print("=" * 80)
    print("BẮT ĐẦU CHẠY THỰC NGHIỆM ĐỀ TÀI 7: KHỬ NHIỄU ẢNH TRONG MIỀN KHÔNG GIAN")
    print("=" * 80)

    # Cố định seed ngẫu nhiên toàn cục
    SEED = 42
    np.random.seed(SEED)

    # Khởi tạo các thư mục lưu kết quả
    tables_dir = os.path.join("results", "tables")
    figures_dir = os.path.join("results", "figures")
    os.makedirs(tables_dir, exist_ok=True)
    os.makedirs(figures_dir, exist_ok=True)

    # 1. Tải và chuẩn bị 6 ảnh xám chuẩn (256x256)
    print("[1/5] Đang chuẩn bị bộ dữ liệu 6 ảnh xám chuẩn kích thước 256x256...")
    dataset = load_or_prepare_dataset(data_dir=os.path.join("data", "original"), size=256)
    plot_dataset_overview(dataset, os.path.join(figures_dir, "dataset_overview.png"))
    print(f"      -> Đã tải thành công {len(dataset)} ảnh: {', '.join(dataset.keys())}")

    # 2. Kiểm chứng tính đúng đắn của các thuật toán tự cài đặt so với OpenCV
    print("[2/5] Đang kiểm chứng đối chiếu thuật toán tự cài đặt với thư viện OpenCV...")
    sample_img = dataset["cameraman"]
    verification_rows = verify_implementations(sample_img)
    ver_csv_path = os.path.join(tables_dir, "verification_results.csv")
    with open(ver_csv_path, mode="w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=["algorithm", "reference", "max_abs_error", "mean_abs_error", "status"])
        writer.writeheader()
        writer.writerows(verification_rows)
    for r in verification_rows:
        print(f"      - {r['algorithm']}: Sai số cực đại = {r['max_abs_error']} -> {r['status']}")

    # 3. Cấu hình các kịch bản nhiễu và bộ lọc
    noise_configs = [
        {"type": "Salt & Pepper", "level_label": "1%", "param": 0.01},
        {"type": "Salt & Pepper", "level_label": "5%", "param": 0.05},
        {"type": "Salt & Pepper", "level_label": "10%", "param": 0.10},
        {"type": "Gaussian", "level_label": "sigma=10", "param": 10.0},
        {"type": "Gaussian", "level_label": "sigma=25", "param": 25.0},
        {"type": "Gaussian", "level_label": "sigma=50", "param": 50.0},
    ]
    kernel_sizes = [3, 5, 7]

    detailed_rows: List[Dict[str, Any]] = []

    print("[3/5] Đang chạy thực nghiệm khử nhiễu và trích xuất biên Sobel (108 kịch bản)...")
    for img_name, orig_img in dataset.items():
        # Tính bản đồ biên Sobel chuẩn (Ground-Truth) trên ảnh gốc
        # Chọn ngưỡng phân vị thứ 85 của ảnh gốc làm chuẩn cố định để so sánh công bằng
        orig_grad_mag, orig_edge_vis, orig_bin_edge, _, ref_threshold = sobel_edge_detection(
            orig_img, threshold=None, pad_mode="reflect"
        )

        for n_cfg in noise_configs:
            n_type = n_cfg["type"]
            n_label = n_cfg["level_label"]
            n_param = n_cfg["param"]

            # Tạo ảnh nhiễu với seed cố định
            if n_type == "Salt & Pepper":
                noisy_img = add_salt_and_pepper_noise(orig_img, density=n_param, seed=SEED)
            else:
                noisy_img = add_gaussian_noise(orig_img, sigma=n_param, seed=SEED)

            # Tính MSE, PSNR và biên Sobel của ảnh nhiễu (trước khi lọc)
            noisy_mse = calculate_mse(orig_img, noisy_img)
            noisy_psnr = calculate_psnr(orig_img, noisy_img)
            noisy_grad_mag, noisy_edge_vis, noisy_bin_edge, _, _ = sobel_edge_detection(
                noisy_img, threshold=ref_threshold, pad_mode="reflect"
            )
            noisy_edge_metrics = evaluate_edge_preservation(
                orig_grad_mag, orig_bin_edge, noisy_grad_mag, noisy_bin_edge
            )

            results_mean_vis: Dict[int, Dict[str, Any]] = {}
            results_median_vis: Dict[int, Dict[str, Any]] = {}

            for k in kernel_sizes:
                # --- Chạy Bộ lọc Trung bình (Mean Filter) ---
                t0 = time.perf_counter()
                mean_img = mean_filter(noisy_img, kernel_size=k, pad_mode="reflect")
                t_mean_ms = (time.perf_counter() - t0) * 1000.0

                mean_mse = calculate_mse(orig_img, mean_img)
                mean_psnr = calculate_psnr(orig_img, mean_img)
                mean_grad, mean_edge_vis, mean_bin_edge, _, _ = sobel_edge_detection(
                    mean_img, threshold=ref_threshold, pad_mode="reflect"
                )
                mean_edge_eval = evaluate_edge_preservation(
                    orig_grad_mag, orig_bin_edge, mean_grad, mean_bin_edge
                )

                detailed_rows.append({
                    "image": img_name,
                    "noise_type": n_type,
                    "noise_level": n_label,
                    "noisy_mse": round(noisy_mse, 2),
                    "noisy_psnr": round(noisy_psnr, 2),
                    "filter_type": "Mean",
                    "kernel_size": k,
                    "mse": round(mean_mse, 2),
                    "psnr": round(mean_psnr, 2),
                    "edge_mse": round(mean_edge_eval["edge_mse"], 2),
                    "edge_precision": round(mean_edge_eval["precision"], 4),
                    "edge_recall": round(mean_edge_eval["recall"], 4),
                    "edge_f1": round(mean_edge_eval["f1"], 4),
                    "time_ms": round(t_mean_ms, 3),
                })

                results_mean_vis[k] = {
                    "filtered": mean_img,
                    "mse": mean_mse,
                    "psnr": mean_psnr,
                    "edge_vis": mean_edge_vis,
                    "edge_mse": mean_edge_eval["edge_mse"],
                    "edge_f1": mean_edge_eval["f1"],
                    "noisy_mse": noisy_mse,
                    "noisy_psnr": noisy_psnr,
                    "noisy_edge_mse": noisy_edge_metrics["edge_mse"],
                    "noisy_edge_f1": noisy_edge_metrics["f1"],
                }

                # --- Chạy Bộ lọc Trung vị (Median Filter) ---
                t0 = time.perf_counter()
                med_img = median_filter(noisy_img, kernel_size=k, pad_mode="reflect")
                t_med_ms = (time.perf_counter() - t0) * 1000.0

                med_mse = calculate_mse(orig_img, med_img)
                med_psnr = calculate_psnr(orig_img, med_img)
                med_grad, med_edge_vis, med_bin_edge, _, _ = sobel_edge_detection(
                    med_img, threshold=ref_threshold, pad_mode="reflect"
                )
                med_edge_eval = evaluate_edge_preservation(
                    orig_grad_mag, orig_bin_edge, med_grad, med_bin_edge
                )

                detailed_rows.append({
                    "image": img_name,
                    "noise_type": n_type,
                    "noise_level": n_label,
                    "noisy_mse": round(noisy_mse, 2),
                    "noisy_psnr": round(noisy_psnr, 2),
                    "filter_type": "Median",
                    "kernel_size": k,
                    "mse": round(med_mse, 2),
                    "psnr": round(med_psnr, 2),
                    "edge_mse": round(med_edge_eval["edge_mse"], 2),
                    "edge_precision": round(med_edge_eval["precision"], 4),
                    "edge_recall": round(med_edge_eval["recall"], 4),
                    "edge_f1": round(med_edge_eval["f1"], 4),
                    "time_ms": round(t_med_ms, 3),
                })

                results_median_vis[k] = {
                    "filtered": med_img,
                    "mse": med_mse,
                    "psnr": med_psnr,
                    "edge_vis": med_edge_vis,
                    "edge_mse": med_edge_eval["edge_mse"],
                    "edge_f1": med_edge_eval["f1"],
                    "noisy_mse": noisy_mse,
                    "noisy_psnr": noisy_psnr,
                    "noisy_edge_mse": noisy_edge_metrics["edge_mse"],
                    "noisy_edge_f1": noisy_edge_metrics["f1"],
                }

            # Lưu hình minh họa trực quan cho các kịch bản tiêu biểu
            safe_label = n_label.replace("%", "pct").replace("=", "")
            if img_name in ("cameraman", "coins", "text"):
                denoise_fig_path = os.path.join(
                    figures_dir, f"compare_{img_name}_{n_type[:2].lower()}_{safe_label}.png"
                )
                plot_denoising_comparison(
                    orig_img, noisy_img, results_mean_vis, results_median_vis,
                    title=f"SO SÁNH KHỬ NHIỄU TRÊN ẢNH '{img_name.upper()}' - NHIỄU {n_type.upper()} ({n_label})",
                    save_path=denoise_fig_path,
                )

                sobel_fig_path = os.path.join(
                    figures_dir, f"sobel_{img_name}_{n_type[:2].lower()}_{safe_label}.png"
                )
                plot_sobel_edge_comparison(
                    orig_edge_vis, noisy_edge_vis, results_mean_vis, results_median_vis,
                    title=f"BẢN ĐỒ BIÊN SOBEL TRÊN ẢNH '{img_name.upper()}' - NHIỄU {n_type.upper()} ({n_label})",
                    save_path=sobel_fig_path,
                )

    # 4. Lưu bảng kết quả chi tiết (108 dòng) và bảng tổng hợp trung bình trên 6 ảnh (36 dòng)
    print("[4/5] Đang tổng hợp số liệu và xuất các bảng CSV...")
    detail_csv_path = os.path.join(tables_dir, "detailed_results.csv")
    with open(detail_csv_path, mode="w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=list(detailed_rows[0].keys()))
        writer.writeheader()
        writer.writerows(detailed_rows)

    summary_rows: List[Dict[str, Any]] = []
    for n_cfg in noise_configs:
        n_type = n_cfg["type"]
        n_label = n_cfg["level_label"]
        for f_type in ("Mean", "Median"):
            for k in kernel_sizes:
                subset = [
                    r for r in detailed_rows
                    if r["noise_type"] == n_type and r["noise_level"] == n_label
                    and r["filter_type"] == f_type and r["kernel_size"] == k
                ]
                summary_rows.append({
                    "noise_type": n_type,
                    "noise_level": n_label,
                    "noisy_psnr_mean": round(float(np.mean([x["noisy_psnr"] for x in subset])), 2),
                    "filter_type": f_type,
                    "kernel_size": k,
                    "mse_mean": round(float(np.mean([x["mse"] for x in subset])), 2),
                    "psnr_mean": round(float(np.mean([x["psnr"] for x in subset])), 2),
                    "edge_mse_mean": round(float(np.mean([x["edge_mse"] for x in subset])), 2),
                    "edge_precision_mean": round(float(np.mean([x["edge_precision"] for x in subset])), 4),
                    "edge_recall_mean": round(float(np.mean([x["edge_recall"] for x in subset])), 4),
                    "edge_f1_mean": round(float(np.mean([x["edge_f1"] for x in subset])), 4),
                    "time_ms_mean": round(float(np.mean([x["time_ms"] for x in subset])), 3),
                })

    summary_csv_path = os.path.join(tables_dir, "summary_by_noise_and_filter.csv")
    with open(summary_csv_path, mode="w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=list(summary_rows[0].keys()))
        writer.writeheader()
        writer.writerows(summary_rows)

    # 5. Vẽ biểu đồ tổng hợp
    print("[5/5] Đang vẽ biểu đồ tổng hợp PSNR, Edge-F1 và Thời gian xử lý...")
    plot_psnr_and_edge_charts(summary_rows, figures_dir)

    print("=" * 80)
    print("HOÀN TẤT THỰC NGHIỆM!")
    print(f"- Bảng kiểm chứng OpenCV : {ver_csv_path}")
    print(f"- Bảng chi tiết 108 mẫu  : {detail_csv_path}")
    print(f"- Bảng trung bình 36 mẫu : {summary_csv_path}")
    print(f"- Thư mục hình & biểu đồ : {figures_dir}")
    print("=" * 80)


if __name__ == "__main__":
    main()

# -*- coding: utf-8 -*-
"""
Module: visualization.py
Mô tả: Trực quan hóa dữ liệu thực nghiệm bằng thư viện Matplotlib:
       1. Lưu hình ảnh tổng quan bộ dữ liệu 6 ảnh xám chuẩn.
       2. Lưu bảng hình so sánh trực quan (Ảnh gốc / Ảnh nhiễu / Lọc Trung bình / Lọc Trung vị).
       3. Lưu bảng hình so sánh bản đồ biên Sobel (Biên gốc / Biên nhiễu / Biên sau lọc).
       4. Vẽ biểu đồ PSNR và Edge-F1 theo mức độ nhiễu và theo kích thước cửa sổ (3x3, 5x5, 7x7).
       5. Vẽ biểu đồ so sánh thời gian xử lý trung bình của hai bộ lọc.
Nhóm thực hiện: Nhóm 10 (Nguyễn Mạnh Quyết - 20231053, Hà Thái Bảo - 20231045)
"""

import os
from typing import Dict, List, Any
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

# Cấu hình font hiển thị tiếng Việt rõ nét trên Matplotlib
plt.rcParams["font.sans-serif"] = ["Arial", "Tahoma", "DejaVu Sans"]
plt.rcParams["axes.unicode_minus"] = False


def plot_dataset_overview(dataset: Dict[str, np.ndarray], save_path: str) -> None:
    """
    Vẽ lưới 2x3 giới thiệu 6 ảnh xám chuẩn trong bộ dữ liệu thực nghiệm.
    """
    os.makedirs(os.path.dirname(save_path), exist_ok=True)
    fig, axes = plt.subplots(2, 3, figsize=(12, 8), dpi=180)
    axes = axes.flatten()

    titles_vi = {
        "cameraman": "1. Cameraman (256x256)\n(Biên sắc nét & nền trời phẳng)",
        "coins": "2. Coins (256x256)\n(Vật thể tròn tách biệt trên nền tối)",
        "astronaut": "3. Astronaut (256x256)\n(Chân dung & chi tiết đa tỷ lệ)",
        "text": "4. Text / Page (256x256)\n(Chữ viết nét mảnh, tương phản cao)",
        "checkerboard": "5. Checkerboard (256x256)\n(Cấu trúc bàn cờ, góc vuông)",
        "brick": "6. Brick (256x256)\n(Kết cấu tường gạch lặp tuần hoàn)",
    }

    for idx, (name, img) in enumerate(dataset.items()):
        ax = axes[idx]
        ax.imshow(img, cmap="gray", vmin=0, vmax=255)
        ax.set_title(titles_vi.get(name, name), fontsize=10, fontweight="bold")
        ax.axis("off")

    fig.suptitle("BỘ DỮ LIỆU 6 ẢNH XÁM CHUẨN SỬ DỤNG TRONG THỰC NGHIỆM", fontsize=13, fontweight="bold", y=0.98)
    plt.tight_layout()
    plt.savefig(save_path, bbox_inches="tight")
    plt.close(fig)


def plot_denoising_comparison(
    original: np.ndarray,
    noisy: np.ndarray,
    results_mean: Dict[int, Dict[str, Any]],
    results_median: Dict[int, Dict[str, Any]],
    title: str,
    save_path: str
) -> None:
    """
    Vẽ hình so sánh trực quan kết quả khử nhiễu:
    Hàng 1: Ảnh gốc | Ảnh nhiễu | Lọc Trung bình 3x3 | Lọc Trung bình 5x5 | Lọc Trung bình 7x7
    Hàng 2: Bản đồ biên gốc | Bản đồ biên nhiễu | Lọc Trung vị 3x3 | Lọc Trung vị 5x5 | Lọc Trung vị 7x7
    """
    os.makedirs(os.path.dirname(save_path), exist_ok=True)
    fig, axes = plt.subplots(2, 4, figsize=(15, 7.6), dpi=180)

    # Hàng 1: Ảnh gốc + Lọc trung bình (3x3, 5x5, 7x7)
    axes[0, 0].imshow(original, cmap="gray", vmin=0, vmax=255)
    axes[0, 0].set_title("Ảnh gốc (Original)", fontsize=10, fontweight="bold")
    axes[0, 0].axis("off")

    for idx, k in enumerate((3, 5, 7), start=1):
        res = results_mean[k]
        axes[0, idx].imshow(res["filtered"], cmap="gray", vmin=0, vmax=255)
        axes[0, idx].set_title(
            f"Lọc Trung bình {k}x{k}\nPSNR: {res['psnr']:.2f} dB | MSE: {res['mse']:.1f}",
            fontsize=9.5
        )
        axes[0, idx].axis("off")

    # Hàng 2: Ảnh nhiễu + Lọc trung vị (3x3, 5x5, 7x7)
    noisy_psnr = results_mean[3]["noisy_psnr"]
    noisy_mse = results_mean[3]["noisy_mse"]
    axes[1, 0].imshow(noisy, cmap="gray", vmin=0, vmax=255)
    axes[1, 0].set_title(
        f"Ảnh nhiễu (Noisy)\nPSNR: {noisy_psnr:.2f} dB | MSE: {noisy_mse:.1f}",
        fontsize=10, fontweight="bold", color="#b22222"
    )
    axes[1, 0].axis("off")

    for idx, k in enumerate((3, 5, 7), start=1):
        res = results_median[k]
        axes[1, idx].imshow(res["filtered"], cmap="gray", vmin=0, vmax=255)
        axes[1, idx].set_title(
            f"Lọc Trung vị {k}x{k}\nPSNR: {res['psnr']:.2f} dB | MSE: {res['mse']:.1f}",
            fontsize=9.5
        )
        axes[1, idx].axis("off")

    fig.suptitle(title, fontsize=12.5, fontweight="bold", y=0.98)
    plt.tight_layout()
    plt.savefig(save_path, bbox_inches="tight")
    plt.close(fig)


def plot_sobel_edge_comparison(
    orig_edge_vis: np.ndarray,
    noisy_edge_vis: np.ndarray,
    results_mean: Dict[int, Dict[str, Any]],
    results_median: Dict[int, Dict[str, Any]],
    title: str,
    save_path: str
) -> None:
    """
    Vẽ hình so sánh trực quan khả năng bảo toàn biên bằng toán tử Sobel:
    Hàng 1: Biên Sobel ảnh gốc + Biên Sobel sau Lọc Trung bình (3x3, 5x5, 7x7)
    Hàng 2: Biên Sobel ảnh nhiễu + Biên Sobel sau Lọc Trung vị (3x3, 5x5, 7x7)
    """
    os.makedirs(os.path.dirname(save_path), exist_ok=True)
    fig, axes = plt.subplots(2, 4, figsize=(15, 7.6), dpi=180)

    axes[0, 0].imshow(orig_edge_vis, cmap="gray", vmin=0, vmax=255)
    axes[0, 0].set_title("Biên Sobel Ảnh gốc\n(Chuẩn tham chiếu)", fontsize=10, fontweight="bold")
    axes[0, 0].axis("off")

    for idx, k in enumerate((3, 5, 7), start=1):
        res = results_mean[k]
        axes[0, idx].imshow(res["edge_vis"], cmap="gray", vmin=0, vmax=255)
        axes[0, idx].set_title(
            f"Biên sau Lọc Trung bình {k}x{k}\nEdge-F1: {res['edge_f1']:.3f} | E-MSE: {res['edge_mse']:.1f}",
            fontsize=9.5
        )
        axes[0, idx].axis("off")

    noisy_f1 = results_mean[3]["noisy_edge_f1"]
    noisy_emse = results_mean[3]["noisy_edge_mse"]
    axes[1, 0].imshow(noisy_edge_vis, cmap="gray", vmin=0, vmax=255)
    axes[1, 0].set_title(
        f"Biên Sobel Ảnh nhiễu\nEdge-F1: {noisy_f1:.3f} | E-MSE: {noisy_emse:.1f}",
        fontsize=10, fontweight="bold", color="#b22222"
    )
    axes[1, 0].axis("off")

    for idx, k in enumerate((3, 5, 7), start=1):
        res = results_median[k]
        axes[1, idx].imshow(res["edge_vis"], cmap="gray", vmin=0, vmax=255)
        axes[1, idx].set_title(
            f"Biên sau Lọc Trung vị {k}x{k}\nEdge-F1: {res['edge_f1']:.3f} | E-MSE: {res['edge_mse']:.1f}",
            fontsize=9.5
        )
        axes[1, idx].axis("off")

    fig.suptitle(title, fontsize=12.5, fontweight="bold", y=0.98)
    plt.tight_layout()
    plt.savefig(save_path, bbox_inches="tight")
    plt.close(fig)


def plot_psnr_and_edge_charts(summary_rows: List[Dict[str, Any]], out_dir: str) -> None:
    """
    Vẽ các biểu đồ phân tích định lượng tổng hợp trên toàn bộ bộ dữ liệu:
      1. Biểu đồ PSNR trung bình theo mức nhiễu Muối tiêu và kích thước cửa sổ.
      2. Biểu đồ PSNR trung bình theo mức nhiễu Gaussian và kích thước cửa sổ.
      3. Biểu đồ F1-Score bảo toàn biên Sobel trên hai loại nhiễu.
      4. Biểu đồ thời gian thực thi trung bình (ms) theo kích thước cửa sổ.
    """
    os.makedirs(out_dir, exist_ok=True)

    # --- 1. Biểu đồ PSNR và Edge-F1 cho Nhiễu Muối Tiêu ---
    sp_levels = ["1%", "5%", "10%"]
    fig, axes = plt.subplots(1, 2, figsize=(13, 5.2), dpi=180)

    markers = {3: "o", 5: "s", 7: "^"}
    for k in (3, 5, 7):
        mean_psnr = [
            row["psnr_mean"] for row in summary_rows
            if row["noise_type"] == "Salt & Pepper" and row["noise_level"] in sp_levels
            and row["filter_type"] == "Mean" and row["kernel_size"] == k
        ]
        median_psnr = [
            row["psnr_mean"] for row in summary_rows
            if row["noise_type"] == "Salt & Pepper" and row["noise_level"] in sp_levels
            and row["filter_type"] == "Median" and row["kernel_size"] == k
        ]
        axes[0].plot(sp_levels, mean_psnr, linestyle="--", marker=markers[k], label=f"Lọc Trung bình {k}x{k}")
        axes[0].plot(sp_levels, median_psnr, linestyle="-", marker=markers[k], linewidth=2, label=f"Lọc Trung vị {k}x{k}")

        mean_f1 = [
            row["edge_f1_mean"] for row in summary_rows
            if row["noise_type"] == "Salt & Pepper" and row["noise_level"] in sp_levels
            and row["filter_type"] == "Mean" and row["kernel_size"] == k
        ]
        median_f1 = [
            row["edge_f1_mean"] for row in summary_rows
            if row["noise_type"] == "Salt & Pepper" and row["noise_level"] in sp_levels
            and row["filter_type"] == "Median" and row["kernel_size"] == k
        ]
        axes[1].plot(sp_levels, mean_f1, linestyle="--", marker=markers[k], label=f"Lọc Trung bình {k}x{k}")
        axes[1].plot(sp_levels, median_f1, linestyle="-", marker=markers[k], linewidth=2, label=f"Lọc Trung vị {k}x{k}")

    axes[0].set_title("PSNR trung bình trên Nhiễu Muối tiêu", fontsize=11, fontweight="bold")
    axes[0].set_xlabel("Mật độ nhiễu Muối tiêu (%)")
    axes[0].set_ylabel("PSNR (dB) - Càng cao càng tốt")
    axes[0].grid(True, linestyle=":", alpha=0.7)
    axes[0].legend(fontsize=8.5)

    axes[1].set_title("Chỉ số bảo toàn biên Sobel (Edge F1-Score) - Muối tiêu", fontsize=11, fontweight="bold")
    axes[1].set_xlabel("Mật độ nhiễu Muối tiêu (%)")
    axes[1].set_ylabel("Edge F1-Score - Càng cao càng tốt")
    axes[1].grid(True, linestyle=":", alpha=0.7)
    axes[1].legend(fontsize=8.5)

    fig.suptitle("ĐÁNH GIÁ HIỆU NĂNG KHỬ NHIỄU VÀ BẢO TOÀN BIÊN TRÊN NHIỄU MUỐI TIÊU", fontsize=12.5, fontweight="bold")
    plt.tight_layout()
    plt.savefig(os.path.join(out_dir, "chart_salt_pepper_psnr_f1.png"), bbox_inches="tight")
    plt.close(fig)

    # --- 2. Biểu đồ PSNR và Edge-F1 cho Nhiễu Gaussian ---
    gauss_levels = ["sigma=10", "sigma=25", "sigma=50"]
    fig, axes = plt.subplots(1, 2, figsize=(13, 5.2), dpi=180)

    for k in (3, 5, 7):
        mean_psnr = [
            row["psnr_mean"] for row in summary_rows
            if row["noise_type"] == "Gaussian" and row["noise_level"] in gauss_levels
            and row["filter_type"] == "Mean" and row["kernel_size"] == k
        ]
        median_psnr = [
            row["psnr_mean"] for row in summary_rows
            if row["noise_type"] == "Gaussian" and row["noise_level"] in gauss_levels
            and row["filter_type"] == "Median" and row["kernel_size"] == k
        ]
        axes[0].plot(gauss_levels, mean_psnr, linestyle="-", marker=markers[k], linewidth=2, label=f"Lọc Trung bình {k}x{k}")
        axes[0].plot(gauss_levels, median_psnr, linestyle="--", marker=markers[k], label=f"Lọc Trung vị {k}x{k}")

        mean_f1 = [
            row["edge_f1_mean"] for row in summary_rows
            if row["noise_type"] == "Gaussian" and row["noise_level"] in gauss_levels
            and row["filter_type"] == "Mean" and row["kernel_size"] == k
        ]
        median_f1 = [
            row["edge_f1_mean"] for row in summary_rows
            if row["noise_type"] == "Gaussian" and row["noise_level"] in gauss_levels
            and row["filter_type"] == "Median" and row["kernel_size"] == k
        ]
        axes[1].plot(gauss_levels, mean_f1, linestyle="-", marker=markers[k], linewidth=2, label=f"Lọc Trung bình {k}x{k}")
        axes[1].plot(gauss_levels, median_f1, linestyle="--", marker=markers[k], label=f"Lọc Trung vị {k}x{k}")

    axes[0].set_title("PSNR trung bình trên Nhiễu Gaussian", fontsize=11, fontweight="bold")
    axes[0].set_xlabel("Độ lệch chuẩn nhiễu Gaussian (sigma)")
    axes[0].set_ylabel("PSNR (dB) - Càng cao càng tốt")
    axes[0].grid(True, linestyle=":", alpha=0.7)
    axes[0].legend(fontsize=8.5)

    axes[1].set_title("Chỉ số bảo toàn biên Sobel (Edge F1-Score) - Gaussian", fontsize=11, fontweight="bold")
    axes[1].set_xlabel("Độ lệch chuẩn nhiễu Gaussian (sigma)")
    axes[1].set_ylabel("Edge F1-Score - Càng cao càng tốt")
    axes[1].grid(True, linestyle=":", alpha=0.7)
    axes[1].legend(fontsize=8.5)

    fig.suptitle("ĐÁNH GIÁ HIỆU NĂNG KHỬ NHIỄU VÀ BẢO TOÀN BIÊN TRÊN NHIỄU GAUSSIAN", fontsize=12.5, fontweight="bold")
    plt.tight_layout()
    plt.savefig(os.path.join(out_dir, "chart_gaussian_psnr_f1.png"), bbox_inches="tight")
    plt.close(fig)

    # --- 3. Biểu đồ Thời gian xử lý trung bình theo kích thước cửa sổ ---
    fig, ax = plt.subplots(figsize=(8.5, 5.0), dpi=180)
    kernels = [3, 5, 7]
    mean_times = []
    median_times = []

    for k in kernels:
        m_t = [r["time_ms_mean"] for r in summary_rows if r["filter_type"] == "Mean" and r["kernel_size"] == k]
        med_t = [r["time_ms_mean"] for r in summary_rows if r["filter_type"] == "Median" and r["kernel_size"] == k]
        mean_times.append(float(np.mean(m_t)))
        median_times.append(float(np.mean(med_t)))

    x = np.arange(len(kernels))
    width = 0.35
    rects1 = ax.bar(x - width / 2, mean_times, width, label="Bộ lọc Trung bình (Mean Filter)", color="#2b6cb0")
    rects2 = ax.bar(x + width / 2, median_times, width, label="Bộ lọc Trung vị (Median Filter)", color="#dd6b20")

    ax.set_title("SO SÁNH THỜI GIAN XỬ LÝ TRUNG BÌNH THEO KÍCH THƯỚC CỬA SỔ (ẢNH 256x256)", fontsize=11.5, fontweight="bold")
    ax.set_xticks(x)
    ax.set_xticklabels([f"Cửa sổ {k}x{k}" for k in kernels])
    ax.set_ylabel("Thời gian thực thi trung bình (ms)")
    ax.grid(True, axis="y", linestyle=":", alpha=0.7)
    ax.legend()

    for rect in rects1:
        h = rect.get_height()
        ax.annotate(f"{h:.2f} ms", xy=(rect.get_x() + rect.get_width() / 2, h),
                    xytext=(0, 3), textcoords="offset points", ha="center", va="bottom", fontsize=9)
    for rect in rects2:
        h = rect.get_height()
        ax.annotate(f"{h:.2f} ms", xy=(rect.get_x() + rect.get_width() / 2, h),
                    xytext=(0, 3), textcoords="offset points", ha="center", va="bottom", fontsize=9)

    plt.tight_layout()
    plt.savefig(os.path.join(out_dir, "chart_execution_time.png"), bbox_inches="tight")
    plt.close(fig)

# -*- coding: utf-8 -*-
"""
Script: scripts/generate_ch2_diagrams.py
Mô tả: Vẽ các sơ đồ kỹ thuật học thuật cho Chương 2 (Sơ đồ quy trình xử lý tổng thể
       và Sơ đồ minh họa kỹ thuật mở rộng biên ảnh Padding) lưu vào results/figures/.
"""

import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as patches

plt.rcParams["font.sans-serif"] = ["Arial", "Tahoma", "DejaVu Sans"]
plt.rcParams["axes.unicode_minus"] = False


def draw_pipeline_diagram(save_path: str) -> None:
    """Vẽ sơ đồ khối quy trình xử lý tổng thể (Pipeline) của ứng dụng."""
    os.makedirs(os.path.dirname(save_path), exist_ok=True)
    fig, ax = plt.subplots(figsize=(11, 5.5), dpi=180)
    ax.set_xlim(0, 11)
    ax.set_ylim(0, 5.5)
    ax.axis("off")

    boxes = [
        (0.4, 3.6, 2.1, 1.2, "#EBF8FF", "#2B6CB0",
         "1. CHUẨN BỊ DỮ LIỆU\n-----------------\n6 Ảnh xám chuẩn\n(256 x 256, uint8)"),
        (3.1, 3.6, 2.2, 1.2, "#FFF5F5", "#C53030",
         "2. GIẢ LẬP NHIỄU\n(seed = 42 cố định)\n-----------------\n- Muối tiêu: 1%, 5%, 10%\n- Gaussian: σ=10, 25, 50"),
        (6.0, 3.6, 2.2, 1.2, "#F0FFF4", "#2F855A",
         "3. LỌC KHÔNG GIAN\n(Tự cài đặt Padding)\n-----------------\n- Lọc Trung bình (3,5,7)\n- Lọc Trung vị (3,5,7)"),
        (8.8, 3.6, 1.9, 1.2, "#FAF5FF", "#6B46C1",
         "4. TRÍCH XUẤT BIÊN\n-----------------\nToán tử Sobel 3x3\n(Gx, Gy, |G|, Ngưỡng)"),
        (6.0, 1.0, 2.2, 1.3, "#FFFAF0", "#C05621",
         "5. ĐÁNH GIÁ CHỈ SỐ\n-----------------\n- MSE & PSNR (dB)\n- Edge-MSE & Edge-F1\n- Thời gian chạy (ms)"),
        (2.4, 1.0, 2.7, 1.3, "#EDF2F7", "#2D3748",
         "6. XUẤT KẾT QUẢ & BÁO CÁO\n-------------------------\n- 3 Bảng số liệu (.CSV)\n- Bộ ảnh so sánh & Biên\n- Biểu đồ phân tích (.PNG)"),
    ]

    for x, y, w, h, fc, ec, text in boxes:
        rect = patches.FancyBboxPatch(
            (x, y), w, h,
            boxstyle="round,pad=0.08,rounding_size=0.12",
            facecolor=fc, edgecolor=ec, linewidth=2
        )
        ax.add_patch(rect)
        ax.text(x + w / 2, y + h / 2, text, ha="center", va="center", fontsize=9.2, fontweight="bold", color="#1A202C")

    arrow_style = dict(arrowstyle="->,head_width=0.35,head_length=0.35", color="#2D3748", lw=2)
    # 1 -> 2
    ax.annotate("", xy=(3.08, 4.2), xytext=(2.52, 4.2), arrowprops=arrow_style)
    # 2 -> 3
    ax.annotate("", xy=(5.98, 4.2), xytext=(5.32, 4.2), arrowprops=arrow_style)
    # 3 -> 4
    ax.annotate("", xy=(8.78, 4.2), xytext=(8.22, 4.2), arrowprops=arrow_style)
    # 1 -> 4 (Ground truth reference curve)
    ax.annotate(
        "Ảnh gốc chuẩn tham chiếu (Ground-Truth)",
        xy=(7.1, 2.32), xytext=(1.45, 3.55),
        ha="center", va="center", fontsize=8.5, fontstyle="italic", color="#2B6CB0",
        arrowprops=dict(arrowstyle="->", connectionstyle="arc3,rad=0.22", color="#2B6CB0", lw=1.5, ls="--")
    )
    # 3 -> 5
    ax.annotate("", xy=(7.1, 2.35), xytext=(7.1, 3.55), arrowprops=arrow_style)
    # 4 -> 5
    ax.annotate("", xy=(8.25, 1.65), xytext=(9.75, 3.55),
                arrowprops=dict(arrowstyle="->,head_width=0.35,head_length=0.35",
                                connectionstyle="angle,angleA=-90,angleB=0,rad=5", color="#2D3748", lw=2))
    # 5 -> 6
    ax.annotate("", xy=(5.15, 1.65), xytext=(5.95, 1.65), arrowprops=arrow_style)

    ax.set_title("SƠ ĐỒ KHỐI QUY TRÌNH XỬ LÝ TỔNG THỂ (PIPELINE) CỦA CHƯƠNG TRÌNH THỰC NGHIỆM",
                 fontsize=12, fontweight="bold", pad=10)
    plt.tight_layout()
    plt.savefig(save_path, bbox_inches="tight")
    plt.close(fig)


def draw_padding_illustration(save_path: str) -> None:
    """Vẽ hình minh họa trực quan 3 chế độ mở rộng biên (Zero, Replicate, Reflect) trên ma trận 3x3."""
    os.makedirs(os.path.dirname(save_path), exist_ok=True)
    fig, axes = plt.subplots(1, 3, figsize=(12, 4.0), dpi=180)

    core = [
        [10, 20, 30],
        [40, 50, 60],
        [70, 80, 90],
    ]
    zero_pad = [
        [0,  0,  0,  0, 0],
        [0, 10, 20, 30, 0],
        [0, 40, 50, 60, 0],
        [0, 70, 80, 90, 0],
        [0,  0,  0,  0, 0],
    ]
    rep_pad = [
        [10, 10, 20, 30, 30],
        [10, 10, 20, 30, 30],
        [40, 40, 50, 60, 60],
        [70, 70, 80, 90, 90],
        [70, 70, 80, 90, 90],
    ]
    ref_pad = [
        [50, 40, 50, 60, 50],
        [20, 10, 20, 30, 20],
        [50, 40, 50, 60, 50],
        [80, 70, 80, 90, 80],
        [50, 40, 50, 60, 50],
    ]

    configs = [
        (axes[0], zero_pad, "a) Zero-Padding\n(Đệm hằng số 0 ở viền)"),
        (axes[1], rep_pad, "b) Replicate-Padding\n(Lặp lại điểm ảnh sát mép)"),
        (axes[2], ref_pad, "c) Reflect-Padding\n(Phản chiếu gương qua biên)"),
    ]

    for ax, mat, title in configs:
        ax.set_xlim(0, 5)
        ax.set_ylim(0, 5)
        ax.invert_yaxis()
        ax.set_xticks([])
        ax.set_yticks([])
        ax.set_title(title, fontsize=10.5, fontweight="bold")

        for r in range(5):
            for c in range(5):
                is_core = (1 <= r <= 3) and (1 <= c <= 3)
                fc = "#BEE3F8" if is_core else "#FEFCBF"
                ec = "#2B6CB0" if is_core else "#B7791F"
                lw = 1.8 if is_core else 1.0
                rect = patches.Rectangle((c, r), 1, 1, facecolor=fc, edgecolor=ec, linewidth=lw)
                ax.add_patch(rect)
                ax.text(c + 0.5, r + 0.5, str(mat[r][c]),
                        ha="center", va="center",
                        fontsize=10.5, fontweight="bold" if is_core else "normal",
                        color="#1A365D" if is_core else "#744210")

    fig.suptitle("MINH HỌA CÁC PHƯƠNG PHÁP MỞ RỘNG BIÊN (PADDING p=1 CHO CỬA SỔ 3x3)\n"
                 "(Ô màu xanh: Ma trận ảnh gốc 3x3 | Ô màu vàng: Dải viền đệm thêm)",
                 fontsize=11.5, fontweight="bold", y=1.02)
    plt.tight_layout()
    plt.savefig(save_path, bbox_inches="tight")
    plt.close(fig)


if __name__ == "__main__":
    fig_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "results", "figures"))
    draw_pipeline_diagram(os.path.join(fig_dir, "ch2_pipeline_diagram.png"))
    draw_padding_illustration(os.path.join(fig_dir, "ch2_padding_modes.png"))
    print("Đã tạo thành công các sơ đồ minh họa Chương 2.")

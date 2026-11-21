# -*- coding: utf-8 -*-
"""
Script: scripts/capture_web_screenshots.py
Mô tả: Khởi chạy ngầm web_app.py và dùng trình duyệt Microsoft Edge (có chữ ký số chính thức
       của Microsoft, không bị chặn bởi Windows Application Control) chụp lại 3 bức ảnh
       giao diện Web thực tế sắc nét để chèn vào Báo cáo BTL_XLA.docx.
"""

import os
import time
import threading
import subprocess
import uvicorn

import sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import web_app


def start_server():
    uvicorn.run(web_app.app, host="127.0.0.1", port=8055, log_level="error")


def capture_with_edge(url: str, out_path: str):
    edge_exe = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    if not os.path.exists(edge_exe):
        edge_exe = r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"

    cmd = [
        edge_exe,
        "--headless",
        "--disable-gpu",
        "--hide-scrollbars",
        "--window-size=1600,1140",
        "--virtual-time-budget=3000",
        f"--screenshot={out_path}",
        url,
    ]
    subprocess.run(cmd, check=True)


def main():
    fig_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "results", "figures"))
    os.makedirs(fig_dir, exist_ok=True)

    # Xóa ảnh test tạm nếu có
    test_img = os.path.join(fig_dir, "test_edge.png")
    if os.path.exists(test_img):
        os.remove(test_img)

    srv_thread = threading.Thread(target=start_server, daemon=True)
    srv_thread.start()
    time.sleep(2.0)

    shots = [
        (
            "http://127.0.0.1:8055/?sample=cameraman&noise=salt_pepper&sp=5.0&k=3",
            os.path.join(fig_dir, "web_ui_salt_pepper.png"),
        ),
    ]

    for url, out_path in shots:
        capture_with_edge(url, out_path)
        print("Đã chụp thành công:", out_path)


if __name__ == "__main__":
    main()

# -*- coding: utf-8 -*-
"""Dựng lại bản cuối vào thư mục Ban_cuoi: báo cáo tổng kết gộp ba chương và bài báo.

    python3 scripts/ban_cuoi/dung_tat_ca.py

Từng chương vẫn có thể dựng riêng bằng chuong1.py, scripts/chuong2/build_gon.py, chuong3.py.
"""
import os
import subprocess
import sys

GOC = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
for s in ["scripts/ra_soat/ra_soat_chuong3.py",   # bản rà soát có theo dõi thay đổi, đầu vào của Chương 3
          "scripts/ban_cuoi/bao_cao.py",
          "scripts/ban_cuoi/bai_bao_khung.py",       # bài báo chính: khung đánh giá hiệu quả
          "scripts/ban_cuoi/bai_bao_chinh_sach.py",  # phương án dự phòng: phân tích chính sách
          "scripts/ban_cuoi/bai_bao.py"]:            # phương án dùng số liệu nội bộ, khi Nhà trường đồng ý
    if os.path.exists(os.path.join(GOC, s)):
        subprocess.run([sys.executable, os.path.join(GOC, s)], check=True, cwd=GOC)

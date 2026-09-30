# -*- coding: utf-8 -*-
"""Dựng lại ba chương bản cuối vào thư mục Ban_cuoi.

    python3 scripts/ban_cuoi/dung_tat_ca.py
"""
import os
import subprocess
import sys

GOC = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
for s in ["scripts/ra_soat/ra_soat_chuong3.py",   # bản rà soát có theo dõi thay đổi, đầu vào của Chương 3
          "scripts/ban_cuoi/chuong1.py",
          "scripts/chuong2/build_gon.py",
          "scripts/ban_cuoi/chuong3.py"]:
    subprocess.run([sys.executable, os.path.join(GOC, s)], check=True, cwd=GOC)

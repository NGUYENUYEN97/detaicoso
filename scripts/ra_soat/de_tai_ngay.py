# -*- coding: utf-8 -*-
"""Trích mã số, quyết định giao, ngày nghiệm thu của đề tài cấp cơ sở từ danh mục PDF của Phòng KHCN."""
import os
import re

import pymupdf

GOC = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
F = os.path.join(GOC, "Tai lieu thanh do", "Tổng hợp đề tài KHCN cấp cơ sở của GV 2021-2025.pdf")


def trich():
    kq = {}
    for trang in pymupdf.open(F):
        for bang in trang.find_tables().tables:
            for h in bang.extract():
                o = " ".join((c or "").replace("\n", " ") for c in h)
                m = re.search(r"(\d{2})\s*-\s*(20\d{2})\s*/\s*KHC\s*N", o)
                if not m:
                    continue
                ma = f"{m.group(1)}-{m.group(2)}"
                ngay = re.findall(r"\d{1,2}/\d{1,2}/20\d{2}", o)
                kq[ma] = ngay
    return kq


if __name__ == "__main__":
    for ma, ngay in sorted(trich().items(), key=lambda x: (x[0][3:], x[0][:2])):
        print(ma, ngay)

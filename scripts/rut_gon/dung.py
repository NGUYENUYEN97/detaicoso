# -*- coding: utf-8 -*-
"""Dựng bản rút gọn của báo cáo tổng kết từ bản v2 của nhóm tác giả.

Giữ nguyên trang bìa, bảng, hình, tài liệu tham khảo và định dạng của bản v2; thay phần thân bằng nội dung rút gọn
trong nd_modau_c1.py, nd_c2.py, nd_c3_kl.py; dựng lại mục lục, danh mục bảng, danh mục hình kèm số trang.
"""
import copy
import os
import re
import subprocess
import sys
import tempfile

import pymupdf as fitz
from docx import Document
from docx.oxml.ns import qn

sys.path.insert(0, os.path.dirname(__file__))
from nd_modau_c1 import MO_DAU, CHUONG_1  # noqa: E402
from nd_c2 import CHUONG_2  # noqa: E402
from nd_c3_kl import CHUONG_3, KET_LUAN  # noqa: E402

GOC = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
VAO = os.path.join(GOC, "Ban_cuoi", "Ban_chinh_sua_02-10-2026", "Bao_cao_tong_ket_de_tai_sua_v2.docx")
RA = os.path.join(GOC, "Ban_cuoi", "Ban_chinh_sua_02-10-2026", "Bao_cao_tong_ket_de_tai_rut_gon.docx")

# Chỉ số khối của bản v2.
MAU = {"H0": 119, "H1": 160, "H2": 161, "H3": 509, "P": 163, "TK": 321}
CAP = {"H0": 0, "H1": 1, "H2": 2, "TK": 1}
TOC_DAU, TOC_MAU = 11, {0: 19, 1: 12, 2: 21}
TOC_KHOANG = (11, 85)          # các mục mục lục, khối 85 là đoạn kết trường
BANG_KHOANG, BANG_DAU, BANG_MAU = (87, 99), 87, 88
HINH_KHOANG, HINH_DAU, HINH_MAU = (101, 118), 101, 102
THAN = (119, 646)
TLTK = (646, 700)
# Tài liệu không còn được trích dẫn trong bản rút gọn.
BO_TLTK = ["Rialti, R.", "Chính phủ. (2023). Nghị định số 17/2023"]

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def chu(el):
    return "".join(t.text or "" for t in el.iter(qn("w:t")))


def tach_markup(s):
    """Tách chuỗi có **đậm** và *nghiêng* thành các đoạn (chữ, đậm, nghiêng)."""
    ra = []
    for m in re.finditer(r"\*\*(.+?)\*\*|\*(.+?)\*|([^*]+)", s):
        if m.group(1):
            ra.append((m.group(1), True, False))
        elif m.group(2):
            ra.append((m.group(2), False, True))
        else:
            ra.append((m.group(3), False, False))
    return ra


class Dung:
    def __init__(self):
        self.doc = Document(VAO)
        self.els = list(self.doc.element.body)
        self.bm = 9000
        self.muc = []   # (cấp, chữ, tên bookmark) cho mục lục

    def bookmark(self, p, ten=None):
        ten = ten or f"_TocRG{self.bm}"
        bs = p.makeelement(qn("w:bookmarkStart"), {qn("w:id"): str(self.bm), qn("w:name"): ten})
        be = p.makeelement(qn("w:bookmarkEnd"), {qn("w:id"): str(self.bm)})
        self.bm += 1
        ppr = p.find(qn("w:pPr"))
        vt = list(p).index(ppr) + 1 if ppr is not None else 0
        p.insert(vt, bs)
        p.append(be)
        return ten

    def doan(self, loai, s):
        mau = self.els[MAU[loai]]
        p = copy.deepcopy(mau)
        rpr = mau.find(qn("w:r")).find(qn("w:rPr"))
        for c in list(p):
            if c.tag != qn("w:pPr"):
                p.remove(c)
        for t, dam, nghieng in tach_markup(s):
            r = p.makeelement(qn("w:r"), {})
            rp = copy.deepcopy(rpr) if rpr is not None else r.makeelement(qn("w:rPr"), {})
            if dam and rp.find(qn("w:b")) is None:
                rp.insert(0, rp.makeelement(qn("w:b"), {}))
            if nghieng and rp.find(qn("w:i")) is None:
                rp.insert(0, rp.makeelement(qn("w:i"), {}))
            r.append(rp)
            te = r.makeelement(qn("w:t"), {})
            te.text = t
            te.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            r.append(te)
            p.append(r)
        if loai in CAP:
            self.muc.append((CAP[loai], s, self.bookmark(p)))
        return p

    def giu(self, a, b, thay):
        ra = []
        for i in range(a, b + 1):
            el = copy.deepcopy(self.els[i])
            for p in ([el] if el.tag == qn("w:p") else el.iter(qn("w:p"))):
                s = chu(p)
                moi = s
                for cu, m in thay.items():
                    moi = moi.replace(cu, m)
                if moi != s:
                    ts = list(p.iter(qn("w:t")))
                    ts[0].text = moi
                    ts[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    for t in ts[1:]:
                        t.text = ""
            ra.append(el)
        return ra

    def chuong(self, idx):
        p = copy.deepcopy(self.els[idx])
        bs = p.find(qn("w:bookmarkStart"))
        ten = bs.get(qn("w:name")) if bs is not None else self.bookmark(p)
        parts = []
        for c in p.iter():
            if c.tag == qn("w:t"):
                parts.append(c.text or "")
            elif c.tag == qn("w:br"):
                parts.append(" ")
        self.muc.append((0, re.sub(r"\s+", " ", "".join(parts)).strip(), ten))
        return p

    def than(self):
        ra = []
        for muc in MO_DAU + CHUONG_1 + CHUONG_2 + CHUONG_3 + KET_LUAN:
            loai = muc[0]
            if loai == "CH":
                ra.append(self.chuong(muc[1]))
            elif loai == "K":
                ra.extend(self.giu(muc[1], muc[2], muc[3]))
            else:
                ra.append(self.doan(loai, muc[1]))
        return ra

    def tltk(self):
        ra = []
        a, b = TLTK
        for i in range(a, b):
            el = self.els[i]
            s = chu(el)
            if any(s.startswith(k) for k in BO_TLTK):
                continue
            el = copy.deepcopy(el)
            if i == a:
                bs = el.find(qn("w:bookmarkStart"))
                self.muc.append((0, s, bs.get(qn("w:name")) if bs is not None else self.bookmark(el)))
            ra.append(el)
        return ra

    def lap(self):
        body = self.doc.element.body
        moi = self.than() + self.tltk()
        cuoi = self.els[-1]
        for el in self.els[THAN[0]:]:
            if el is not cuoi:
                body.remove(el)
        for el in moi:
            cuoi.addprevious(el)
        self.chu_thich = [p for p in moi if p.tag == qn("w:p") and p.find(qn("w:pPr")) is not None
                          and p.find(qn("w:pPr")).find(qn("w:pStyle")) is not None
                          and p.find(qn("w:pPr")).find(qn("w:pStyle")).get(qn("w:val")) in ("Chuthichbang", "Chuthichhinh")]

    @staticmethod
    def muc_toc(mau, chu_muc, ten, trang):
        p = copy.deepcopy(mau)
        h = p.find(qn("w:hyperlink"))
        h.set(qn("w:anchor"), ten)
        rs = h.findall(qn("w:r"))
        rs[0].find(qn("w:t")).text = chu_muc
        for it in h.iter(qn("w:instrText")):
            it.text = f" PAGEREF {ten} \\h "
        ts = [r.find(qn("w:t")) for r in rs[1:] if r.find(qn("w:t")) is not None]
        ts[-1].text = str(trang)
        return p

    def dung_danh_muc(self, khoang, dau, mau_khac, cac_muc):
        """Thay các khối từ khoang[0] đến trước khoang[1] bằng danh sách (mẫu, chữ, bookmark, trang)."""
        a, b = khoang
        cu = self.els[a:b]
        neo = self.els[b]
        mau_dau = copy.deepcopy(self.els[dau])
        mau_k = {k: copy.deepcopy(self.els[v]) for k, v in mau_khac.items()} if isinstance(mau_khac, dict) \
            else copy.deepcopy(self.els[mau_khac])
        # Giữ lại các run mở trường của mục đầu.
        h_dau = mau_dau.find(qn("w:hyperlink"))
        for el in cu:
            el.getparent().remove(el)
        for j, (cap, s, ten, trang) in enumerate(cac_muc):
            if j == 0:
                p = self.muc_toc(mau_dau, s, ten, trang)
            else:
                m = mau_k[cap] if isinstance(mau_k, dict) else mau_k
                p = self.muc_toc(m, s, ten, trang)
            neo.addprevious(p)
        return h_dau

    def danh_muc(self, trang):
        """trang: dict bookmark -> số trang."""
        self.dung_danh_muc(TOC_KHOANG, TOC_DAU, TOC_MAU,
                           [(c, s, t, trang.get(t, "")) for c, s, t in self.muc])
        bang, hinh = [], []
        for p in self.chu_thich:
            bs = p.find(qn("w:bookmarkStart"))
            ten = bs.get(qn("w:name")) if bs is not None else self.bookmark(p)
            s = chu(p).strip()
            (bang if s.startswith("Bảng") else hinh).append((0, s, ten, trang.get(ten, "")))
        self.dung_danh_muc(BANG_KHOANG, BANG_DAU, BANG_MAU, bang)
        self.dung_danh_muc(HINH_KHOANG, HINH_DAU, HINH_MAU, hinh)

    def cac_dich(self):
        """(bookmark, chữ cần tìm) theo thứ tự xuất hiện trong thân."""
        ra = [(t, s) for _, s, t in self.muc] + [("_DAT_LAI_", "")]
        for p in self.chu_thich:
            bs = p.find(qn("w:bookmarkStart"))
            if bs is None:
                self.bookmark(p)
                bs = p.find(qn("w:bookmarkStart"))
            ra.append((bs.get(qn("w:name")), chu(p).strip()))
        return ra


def ve_pdf(duong, thu_muc):
    subprocess.run(["soffice", "--headless", "-env:UserInstallation=file:///tmp/lo_prof3", "--convert-to", "pdf",
                    "--outdir", thu_muc, duong], check=True, capture_output=True, timeout=600)
    return os.path.join(thu_muc, os.path.splitext(os.path.basename(duong))[0] + ".pdf")


def chuan(s):
    return re.sub(r"\s+", " ", s).strip()


def tim_trang(pdf, dich):
    d = fitz.open(pdf)
    trang = [chuan(p.get_text()) for p in d]
    # Trang Mở đầu: trang đầu tiên sau phần danh mục có tiêu đề MỞ ĐẦU ở đầu trang.
    md = next(i for i, s in enumerate(trang) if s.startswith("MỞ ĐẦU") or s.startswith("MỞ ĐẦU ".strip()) and i > 3)
    ra, vt = {}, md
    for ten, s in dich:
        if ten == "_DAT_LAI_":
            vt = md
            continue
        mau = chuan(s)[:45]
        for i in range(vt, len(trang)):
            if mau in trang[i]:
                ra[ten] = i - md + 1
                vt = i
                break
        else:
            print("Không tìm thấy trang:", mau)
    return ra, len(d), len(d) - md


def main():
    tmp = tempfile.mkdtemp()
    b = Dung()
    b.lap()
    dich = b.cac_dich()
    b.danh_muc({})
    nhap = os.path.join(tmp, "nhap.docx")
    b.doc.save(nhap)
    trang, tong, than = tim_trang(ve_pdf(nhap, tmp), dich)

    b = Dung()
    b.lap()
    b.cac_dich()
    b.danh_muc(trang)
    b.doc.save(RA)
    _, tong, than = tim_trang(ve_pdf(RA, tmp), dich)
    print(f"Đã lưu {RA}: {tong} trang PDF, phần thân từ Mở đầu {than} trang")


if __name__ == "__main__":
    main()

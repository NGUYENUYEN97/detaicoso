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
BO_TLTK = ["Rialti, R.", "Chính phủ. (2023). Nghị định số 17/2023", "Quốc hội. (2025c). Luật số 123/2025"]

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# Tên gọi tắt thống nhất, áp dụng cho các bảng, hình giữ lại từ bản v2 (thứ tự thay có ý nghĩa).
TEN_GON = [
    ("theo quyền tại điểm c khoản 1 Điều 86 Luật Sở hữu trí tuệ", "theo quyền đăng ký của tổ chức chủ trì"),
    ("Dựa vào Điều 28, Điều 73 Luật số 93/2025/QH15 và Điều 135 Luật Sở hữu trí tuệ để",
     "Dựa vào Luật KH,CN&ĐMST và Luật Sở hữu trí tuệ để"),
    ("theo Điều 28, Điều 73 Luật số 93/2025/QH15, Điều 32, Điều 34 Nghị định số 267/2025/NĐ-CP và Điều 135 Luật Sở "
     "hữu trí tuệ", "theo Luật KH,CN&ĐMST, Nghị định 267 và Luật Sở hữu trí tuệ"),
    ("điểm a Điều 36 Quyết định 213 cần rà soát theo Điều 28, Điều 73 Luật số 93/2025/QH15",
     "mức trần tại Quyết định 213 cần rà soát theo Luật KH,CN&ĐMST"),
    ("điểm c khoản 1 Điều 86 do Luật số 131/2025/QH15 bổ sung trao quyền đăng ký cho tổ chức chủ trì; Điều 28 Luật Giáo "
     "dục đại học số 125/2025/QH15 cho phép",
     "Luật Sở hữu trí tuệ trao quyền đăng ký cho tổ chức chủ trì; Luật Giáo dục đại học cho phép"),
    ("Điều 35, Điều 38 Quyết định 213 có căn cứ chi", "Quyết định 213 có căn cứ chi"),
    ("Điều 35 Quyết định 213 giao Phòng Khoa học Công nghệ nộp đơn, lệ phí; Điều 38 cho phép chi thuê ngoài, chi khác "
     "liên quan trực tiếp", "Quyết định 213 giao Phòng Khoa học Công nghệ nộp đơn, lệ phí và cho phép chi thuê ngoài"),
    (" theo Điều 119 Luật Sở hữu trí tuệ", ""),
    (" Bảng tổng hợp cuối Phụ lục II Thông tư số 83/2026/TT-BGDĐT chưa nêu giải pháp hữu ích dù công thức đã tính.", ""),
    ("các văn bản nêu tại Mục 3.1.1", "các văn bản tại Bảng 1.1"),
    ("theo Bảng 1 Phụ lục I Thông tư số 83/2026/TT-BGDĐT", "theo Thông tư 83"),
    ("trên cơ sở Điều 11 Quyết định 217", "trên cơ sở Quyết định 217"),
    ("khoản 4 Điều 36 Quy chế ban hành kèm Quyết định số 213/QĐ-ĐHTĐ", "Điều 36 Quyết định 213"),
    ("điểm a và điểm b khoản 4 Điều 36 Quyết định 213", "điểm a và điểm b Điều 36 Quyết định 213"),
    ("Luật số 93/2025/QH15 và Văn bản hợp nhất số 67/VBHN-VPQH", "Luật KH,CN&ĐMST và Luật Sở hữu trí tuệ"),
    ("Luật Giáo dục đại học số 125/2025/QH15", "Luật Giáo dục đại học"),
    ("Luật số 93/2025/QH15", "Luật KH,CN&ĐMST"),
    ("Nghị định số 267/2025/NĐ-CP", "Nghị định 267"),
    ("Nghị định số 134/2026/NĐ-CP", "Nghị định 134"),
    ("Thông tư số 83/2026/TT-BGDĐT", "Thông tư 83"),
    ("Quyết định số 1624/QĐ-TTg", "Quyết định 1624"),
    ("Kết luận số 51-KL/TW", "Kết luận 51"),
    ("Chỉ thị số 02/CT-TTg", "Chỉ thị 02"),
    ("Kế hoạch số 07/KH-ĐHTĐ", "Kế hoạch 07"),
    ("Kế hoạch 07/KH-ĐHTĐ", "Kế hoạch 07"),
]


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
                for cu, m in TEN_GON + list(thay.items()):
                    moi = moi.replace(cu, m)
                if moi != s:
                    ts = list(p.iter(qn("w:t")))
                    ts[0].text = moi
                    ts[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    for t in ts[1:]:
                        t.text = ""
            ra.append(el)
        return ra

    def o(self, mau_tc, cac_doan, rong):
        """Một ô bảng từ ô mẫu, mỗi phần tử của cac_doan là một đoạn, nhận **đậm**."""
        tc = copy.deepcopy(mau_tc)
        tc.find(qn("w:tcPr")).find(qn("w:tcW")).set(qn("w:w"), str(rong))
        mau_p = tc.find(qn("w:p"))
        rpr = mau_p.find(qn("w:r")).find(qn("w:rPr"))
        for p in tc.findall(qn("w:p")):
            tc.remove(p)
        for s in cac_doan:
            p = copy.deepcopy(mau_p)
            for c in list(p):
                if c.tag != qn("w:pPr"):
                    p.remove(c)
            for chu_, dam, _ in tach_markup(s):
                r = p.makeelement(qn("w:r"), {})
                rp = copy.deepcopy(rpr)
                b = rp.find(qn("w:b"))
                if dam and b is None:
                    rp.insert(0, rp.makeelement(qn("w:b"), {}))
                r.append(rp)
                te = r.makeelement(qn("w:t"), {})
                te.text = chu_
                te.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                r.append(te)
                p.append(r)
            tc.append(p)
        return tc

    def bang(self, chu_thich, tieu_de, hang, nguon, rong):
        """Dựng bảng mới theo mẫu Bảng 3.2 của bản v2 (khối 495 - 497)."""
        cap = copy.deepcopy(self.els[495])
        for bm in cap.findall(qn("w:bookmarkStart")) + cap.findall(qn("w:bookmarkEnd")):
            cap.remove(bm)
        ts = list(cap.iter(qn("w:t")))
        ts[0].text = chu_thich
        for x in ts[1:]:
            x.text = ""
        tbl = copy.deepcopy(self.els[496])
        grid = tbl.find(qn("w:tblGrid"))
        for g, w in zip(grid.findall(qn("w:gridCol")), rong):
            g.set(qn("w:w"), str(w))
        trs = tbl.findall(qn("w:tr"))
        tr_dau, tr_mau = trs[0], trs[1]
        for tr in trs:
            tbl.remove(tr)
        tr = copy.deepcopy(tr_dau)
        tcs = tr.findall(qn("w:tc"))
        for tc in tcs:
            tr.remove(tc)
        for j, s in enumerate(tieu_de):
            tr.append(self.o(tcs[j], ["**" + s + "**"], rong[j]))
        tbl.append(tr)
        for h in hang:
            tr = copy.deepcopy(tr_mau)
            trpr = tr.find(qn("w:trPr"))
            if trpr is None:
                trpr = tr.makeelement(qn("w:trPr"), {})
                tr.insert(0, trpr)
            if trpr.find(qn("w:cantSplit")) is None:
                trpr.insert(0, trpr.makeelement(qn("w:cantSplit"), {}))
            tcs = tr.findall(qn("w:tc"))
            for tc in tcs:
                tr.remove(tc)
            for j, cac_doan in enumerate(h):
                # Phần mở đầu trước dấu hai chấm ở cột giữa được in đậm cho dễ dò.
                cac_doan = [re.sub(r"^([^:*]{1,40}):", r"**\1:**", s) if j == 1 else s for s in cac_doan]
                tc = self.o(tcs[j], cac_doan, rong[j])
                for p in tc.findall(qn("w:p")):
                    jc = p.find(qn("w:pPr")).find(qn("w:jc"))
                    if jc is not None:
                        jc.set(qn("w:val"), "left")
                tr.append(tc)
            tbl.append(tr)
        ng = copy.deepcopy(self.els[497])
        ts = list(ng.iter(qn("w:t")))
        ts[0].text = nguon
        for x in ts[1:]:
            x.text = ""
        return [cap, tbl, ng]

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
            elif loai == "BANG":
                ra.extend(self.bang(*muc[1:]))
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
            if s.startswith("Quốc hội. (2025d)"):
                for x in el.iter(qn("w:t")):
                    x.text = (x.text or "").replace("2025d", "2025c")
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

# -*- coding: utf-8 -*-
"""Dựng báo cáo toàn văn theo khung đã chốt (cả ba chương theo chu trình 4 khâu).

Nguồn: bản toàn văn trước (Bao_cao_toan_van_de_tai.docx), giữ nguyên trang bìa, cam đoan, các bảng, hình và các đoạn
đã có; nội dung mới lấy từ nd_khung_moi.py.
"""
import copy
import os
import re
import sys
import tempfile

from docx.oxml.ns import qn

sys.path.insert(0, os.path.dirname(__file__))
import toan_van  # noqa: E402
from toan_van import Sua, chu, bo_id, XML_SPACE  # noqa: E402
from dung import tach_markup, ve_pdf, tim_trang  # noqa: E402
from nd_khung_moi import VIET_TAT, MO_DAU, CHUONG_1, CHUONG_2, CHUONG_3, KET_LUAN  # noqa: E402

THU_MUC = toan_van.THU_MUC
NGUON = os.path.join(THU_MUC, "Bao_cao_toan_van_de_tai.docx")
RA = os.path.join(THU_MUC, "Bao_cao_toan_van_de_tai_khung_moi.docx")


class KhungMoi(Sua):
    def __init__(self):
        toan_van.VAO = NGUON
        super().__init__()
        self.mau["TK"] = copy.deepcopy(self.tim("TIỂU KẾT CHƯƠNG 1"))
        self.mau_ch = copy.deepcopy(self.tim("CHƯƠNG 1", kieu="H0"))
        self.goc = list(self.body)
        ids = [int(b.get(qn("w:id"))) for b in self.body.iter(qn("w:bookmarkStart"))]
        self.bm = max(ids + [0]) + 1
        self.md = self.goc.index(self.tim("MỞ ĐẦU", kieu="H0"))
        self.mau_bang = copy.deepcopy(self.o_bang("Bảng 3.3.", "3.3"))
        self.mau_cap = copy.deepcopy(self.tim("Bảng 3.3."))
        self.mau_nguon = copy.deepcopy(self.tim("Nguồn: Nhóm nghiên cứu đề xuất trên cơ sở Mục 2.5.2"))

    # ---- tìm trong bản gốc ----
    def goc_tim(self, dau, tu=None):
        bat_dau = self.goc.index(tu) if tu is not None else self.md
        for el in self.goc[bat_dau:]:
            if el.tag == qn("w:p") and chu(el).strip().startswith(dau):
                ppr = el.find(qn("w:pPr"))
                st = ppr.find(qn("w:pStyle")) if ppr is not None else None
                if st is not None and st.get(qn("w:val")).startswith("TOC"):
                    continue
                return el
        raise KeyError(dau)

    def sua_chu(self, el, thay):
        for p in ([el] if el.tag == qn("w:p") else el.iter(qn("w:p"))):
            s = chu(p)
            moi = s
            for cu, m in thay.items():
                moi = moi.replace(cu, m)
            if moi != s:
                self.dat_chu(p, moi)
        return el

    def gp(self, dau, thay=None):
        el = copy.deepcopy(self.goc_tim(dau))
        if thay:
            truoc = chu(el)
            self.sua_chu(el, thay)
            assert chu(el) != truoc, ("không thay được", dau)
        return [el]

    def gk(self, cap, thay):
        dau = self.goc_tim(cap)
        i = self.goc.index(dau)
        ra = []
        for el in self.goc[i:]:
            ra.append(self.sua_chu(copy.deepcopy(el), thay))
            if el is not dau and el.tag == qn("w:p") and chu(el).strip().startswith("Nguồn:"):
                break
        return ra

    def gkr(self, dau, cuoi, thay):
        a = self.goc_tim(dau)
        b = self.goc_tim(cuoi, tu=a)
        i, j = self.goc.index(a), self.goc.index(b)
        return [self.sua_chu(copy.deepcopy(el), thay) for el in self.goc[i:j + 1]]

    # ---- đoạn mới ----
    def doan(self, loai, s, mau=None):
        if loai == "TK":
            p = super().doan("H1", s, mau=self.mau["TK"])
            return p
        return super().doan(loai, s, mau)

    def chuong(self, so, dong):
        p = bo_id(copy.deepcopy(self.mau_ch))
        r0 = p.find(qn("w:r"))
        rpr = copy.deepcopy(r0.find(qn("w:rPr")))
        for c in list(p):
            if c.tag != qn("w:pPr"):
                p.remove(c)
        for k, s in enumerate([f"CHƯƠNG {so}"] + dong):
            if k:
                r = p.makeelement(qn("w:r"), {})
                r.append(r.makeelement(qn("w:br"), {}))
                p.append(r)
            r = p.makeelement(qn("w:r"), {})
            r.append(copy.deepcopy(rpr))
            t = r.makeelement(qn("w:t"), {})
            t.text = s
            r.append(t)
            p.append(r)
        self.bookmark(p)
        return p

    def o(self, tc_mau, s, rong, dam=False, can=None):
        tc = bo_id(copy.deepcopy(tc_mau))
        tc.find(qn("w:tcPr")).find(qn("w:tcW")).set(qn("w:w"), str(rong))
        for x in tc.findall(qn("w:tcPr")):
            for g in x.findall(qn("w:gridSpan")):
                x.remove(g)
        p = tc.findall(qn("w:p"))[0]
        for x in tc.findall(qn("w:p"))[1:]:
            tc.remove(x)
        r0 = p.find(".//" + qn("w:r"))
        rpr = copy.deepcopy(r0.find(qn("w:rPr"))) if r0 is not None and r0.find(qn("w:rPr")) is not None else None
        for c in list(p):
            if c.tag != qn("w:pPr"):
                p.remove(c)
        if can:
            jc = p.find(qn("w:pPr")).find(qn("w:jc"))
            if jc is not None:
                jc.set(qn("w:val"), can)
        for t, b, _ in tach_markup(("**" + s + "**") if dam else s):
            r = p.makeelement(qn("w:r"), {})
            rp = copy.deepcopy(rpr) if rpr is not None else r.makeelement(qn("w:rPr"), {})
            if rp.find(qn("w:b")) is not None and not b:
                rp.remove(rp.find(qn("w:b")))
            if b and rp.find(qn("w:b")) is None:
                rp.insert(0, rp.makeelement(qn("w:b"), {}))
            r.append(rp)
            te = r.makeelement(qn("w:t"), {})
            te.text = t
            te.set(XML_SPACE, "preserve")
            r.append(te)
            p.append(r)
        return tc

    def bang_moi(self, tieu_de, hang, rong):
        tbl = bo_id(copy.deepcopy(self.mau_bang))
        grid = tbl.find(qn("w:tblGrid"))
        for g in grid.findall(qn("w:gridCol")):
            grid.remove(g)
        for w in rong:
            grid.append(grid.makeelement(qn("w:gridCol"), {qn("w:w"): str(w)}))
        trs = tbl.findall(qn("w:tr"))
        for tr in trs:
            tbl.remove(tr)
        for k, cac_o in enumerate([tieu_de] + hang):
            tr = copy.deepcopy(trs[0] if k == 0 else trs[1])
            tcs = tr.findall(qn("w:tc"))
            for tc in tcs:
                tr.remove(tc)
            for j, s in enumerate(cac_o):
                tr.append(self.o(tcs[min(j, len(tcs) - 1)], s, rong[j], dam=(k == 0),
                                 can="center" if k == 0 else "left"))
            tbl.append(tr)
        return tbl

    def bang(self, cap, tieu_de, hang, nguon, rong):
        c = bo_id(copy.deepcopy(self.mau_cap))
        for bm in c.findall(qn("w:bookmarkStart")) + c.findall(qn("w:bookmarkEnd")):
            c.remove(bm)
        self.dat_chu(c, cap)
        n = bo_id(copy.deepcopy(self.mau_nguon))
        self.dat_chu(n, nguon)
        return [c, self.bang_moi(tieu_de, hang, rong), n]

    def dung(self, ds):
        ra = []
        for m in ds:
            k = m[0]
            if k == "CH":
                ra.append(self.chuong(m[1], m[2]))
            elif k == "GP":
                ra += self.gp(m[1])
            elif k == "GPS":
                ra += self.gp(m[1], m[2])
            elif k == "GK":
                ra += self.gk(m[1], m[2])
            elif k == "GKR":
                ra += self.gkr(m[1], m[2], m[3])
            elif k == "BANG":
                ra += self.bang(*m[1:])
            else:
                ra.append(self.doan(k, m[1]))
        return ra

    # ---- phần đầu: danh mục chữ viết tắt ----
    def viet_tat(self):
        ket = self.goc[self.md - 1]
        ppr = ket.find(qn("w:pPr"))
        sect = ppr.find(qn("w:sectPr"))
        assert sect is not None
        td = bo_id(copy.deepcopy(self.tim("DANH MỤC HÌNH")))
        self.dat_chu(td, "DANH MỤC CHỮ VIẾT TẮT")
        tbl = self.bang_moi(["Chữ viết tắt", "Nội dung đầy đủ"], [list(x) for x in VIET_TAT], [2500, 6631])
        p_sect = bo_id(copy.deepcopy(self.mau["P"]))
        for c in list(p_sect):
            if c.tag != qn("w:pPr"):
                p_sect.remove(c)
        ppr.remove(sect)
        p_sect.find(qn("w:pPr")).append(sect)
        self.chen_sau(ket, [td, tbl, p_sect])

    # ---- tài liệu tham khảo: bỏ mục không còn được trích ----
    def tltk(self, chu_than):
        a = self.goc_tim("TÀI LIỆU THAM KHẢO")
        i = self.goc.index(a)
        ra, bo = [], []
        for el in self.goc[i:]:
            if el.tag == qn("w:sectPr"):
                break
            s = chu(el).strip()
            giu = True
            if el is not a and s and not re.match(r"^[ABC]\. ", s):
                so = re.search(r"số (\d+[\w/\-–ĐđÐ]*)", s)
                ten = s.split("). ", 1)[1] if "). " in s else ""
                ten_dau = " ".join(ten.split()[:5])
                if s.startswith("Cục Sở hữu trí tuệ. (n.d.)"):
                    giu = "n.d." in chu_than
                elif so or s.startswith(("Trường Đại học Thành Đô", "Quốc hội", "Chính phủ", "Bộ ", "Thủ tướng",
                                         "Văn phòng Quốc hội")):
                    giu = bool(so and so.group(1) in chu_than) or bool(ten_dau and ten_dau in chu_than)
                else:
                    tg = re.match(r"^([^,(]+)", s).group(1).strip().rstrip(".")
                    nam = re.search(r"\((\d{4})", s)
                    giu = tg in chu_than and (nam is None or nam.group(1) in chu_than)
            if giu:
                ra.append(copy.deepcopy(el))
            else:
                bo.append(s[:60])
        return ra, bo

    def lam(self):
        than = self.dung(MO_DAU + CHUONG_1 + CHUONG_2 + CHUONG_3 + KET_LUAN)
        chu_than = " ".join(chu(e) for e in than) + " " + " ".join(n for v in VIET_TAT for n in v)
        tl, self.tltk_bo = self.tltk(chu_than)
        cuoi = self.goc[-1]
        for el in self.goc[self.md:]:
            if el is not cuoi:
                self.body.remove(el)
        for el in than + tl:
            cuoi.addprevious(el)
        self.viet_tat()


def main():
    tmp = tempfile.mkdtemp()
    s = KhungMoi()
    s.lam()
    dich = s.danh_muc({})
    nhap = os.path.join(tmp, "nhap.docx")
    s.doc.save(nhap)
    trang, _, _ = tim_trang(ve_pdf(nhap, tmp), dich)
    s = KhungMoi()
    s.lam()
    s.danh_muc(trang)
    s.doc.save(RA)
    _, tong, than = tim_trang(ve_pdf(RA, tmp), dich)
    print("Bỏ khỏi tài liệu tham khảo:", s.tltk_bo)
    print(f"Đã lưu {RA}: {tong} trang PDF, phần thân từ Mở đầu {than} trang")


if __name__ == "__main__":
    main()

# -*- coding: utf-8 -*-
"""Dựng báo cáo tóm tắt tổng kết đề tài trên file mẫu của Trường (Mau_bao_cao_tom_tat.docx).

Các ô thông tin hành chính (mã số, thời gian thực hiện thực tế, kinh phí thực hiện, tình trạng bài báo) để trống
cho nhóm tác giả tự điền.
"""
import copy
import os
import sys

from docx import Document
from docx.oxml.ns import qn

sys.path.insert(0, os.path.dirname(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "rut_gon"))
from dung import tach_markup, ve_pdf  # noqa: E402
from nd_tom_tat import TEN, THAY_DOI, PHAN_II, KIEN_NGHI  # noqa: E402

GOC = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
THU_MUC = os.path.join(GOC, "Ban_cuoi", "Ban_hoan_thien_03-10-2026")
VAO = os.path.join(THU_MUC, "Mau_bao_cao_tom_tat.docx")
RA = os.path.join(THU_MUC, "Bao_cao_tom_tat_de_tai.docx")
W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def chu(el):
    return "".join(t.text or "" for t in el.iter(qn("w:t")))


def bo_id(el):
    for x in el.iter():
        for k in (W14 + "paraId", W14 + "textId"):
            if k in x.attrib:
                del x.attrib[k]
    return el


def them_run(p, mau_rpr, s, bo_dam=False):
    for t, dam, nghieng in tach_markup(s):
        r = p.makeelement(qn("w:r"), {})
        rp = copy.deepcopy(mau_rpr) if mau_rpr is not None else r.makeelement(qn("w:rPr"), {})
        for th in ("w:b", "w:i"):
            if bo_dam and rp.find(qn(th)) is not None:
                rp.remove(rp.find(qn(th)))
        if dam and rp.find(qn("w:b")) is None:
            rp.insert(0, rp.makeelement(qn("w:b"), {}))
        if nghieng and rp.find(qn("w:i")) is None:
            rp.insert(0, rp.makeelement(qn("w:i"), {}))
        r.append(rp)
        te = r.makeelement(qn("w:t"), {})
        te.text = t
        te.set(XML_SPACE, "preserve")
        r.append(te)
        p.append(r)


class TomTat:
    def __init__(self):
        self.doc = Document(VAO)
        self.body = self.doc.element.body
        self.mau_td = copy.deepcopy(self.tim("1. Đặt vấn đề"))
        self.mau_p = copy.deepcopy(self.tim("Viết theo cấu trúc một bài báo"))
        ppr = self.mau_p.find(qn("w:pPr"))
        ppr.find(qn("w:ind")).set(qn("w:firstLine"), "567")
        jc = ppr.makeelement(qn("w:jc"), {qn("w:val"): "both"})
        ppr.find(qn("w:rPr")).addprevious(jc)

    def tim(self, dau):
        for el in self.body:
            if el.tag == qn("w:p") and chu(el).strip().startswith(dau):
                return el
        raise KeyError(dau)

    def doan(self, loai, s, can_giua=False):
        mau = self.mau_td if loai == "H" else self.mau_p
        p = bo_id(copy.deepcopy(mau))
        rpr = mau.find(qn("w:r")).find(qn("w:rPr"))
        for c in list(p):
            if c.tag != qn("w:pPr"):
                p.remove(c)
        if can_giua:
            ppr = p.find(qn("w:pPr"))
            ppr.find(qn("w:ind")).set(qn("w:firstLine"), "0")
            jc = ppr.find(qn("w:jc"))
            if jc is None:
                jc = ppr.makeelement(qn("w:jc"), {})
                ppr.find(qn("w:rPr")).addprevious(jc)
            jc.set(qn("w:val"), "center")
        them_run(p, rpr, s, bo_dam=(loai == "P"))
        return p

    @staticmethod
    def dat_o(tc, s, can="left"):
        """Ghi chữ vào ô bảng, giữ định dạng ô; nhận **đậm**."""
        ps = tc.findall(qn("w:p"))
        p = ps[0]
        for x in ps[1:]:
            tc.remove(x)
        r0 = p.find(".//" + qn("w:r"))
        rpr = copy.deepcopy(r0.find(qn("w:rPr"))) if r0 is not None and r0.find(qn("w:rPr")) is not None else None
        if rpr is None:
            ppr = p.find(qn("w:pPr"))
            prp = ppr.find(qn("w:rPr")) if ppr is not None else None
            rpr = copy.deepcopy(prp) if prp is not None else None
        for c in list(p):
            if c.tag != qn("w:pPr"):
                p.remove(c)
        if can:
            ppr = p.find(qn("w:pPr"))
            if ppr is None:
                ppr = p.makeelement(qn("w:pPr"), {})
                p.insert(0, ppr)
            jc = ppr.find(qn("w:jc"))
            if jc is None:
                jc = ppr.makeelement(qn("w:jc"), {})
                if ppr.find(qn("w:rPr")) is not None:
                    ppr.find(qn("w:rPr")).addprevious(jc)
                else:
                    ppr.append(jc)
            jc.set(qn("w:val"), can)
        them_run(p, rpr, s, bo_dam=True)

    def bang_moi(self, tieu_de, hang, rong):
        """Bảng hai cột theo mẫu bảng danh sách thành viên của file mẫu."""
        mau = self.body.findall(qn("w:tbl"))[0]
        tbl = bo_id(copy.deepcopy(mau))
        grid = tbl.find(qn("w:tblGrid"))
        for g in grid.findall(qn("w:gridCol")):
            grid.remove(g)
        for w in rong:
            grid.append(grid.makeelement(qn("w:gridCol"), {qn("w:w"): str(w)}))
        trs = tbl.findall(qn("w:tr"))
        for tr in trs:
            tbl.remove(tr)

        def hang_moi(tr_mau, cac_o, dam):
            tr = copy.deepcopy(tr_mau)
            tcs = tr.findall(qn("w:tc"))
            for tc in tcs[len(cac_o):]:
                tr.remove(tc)
            for tc, s, w in zip(tcs, cac_o, rong):
                tc.find(qn("w:tcPr")).find(qn("w:tcW")).set(qn("w:w"), str(w))
                if isinstance(s, list):
                    self.dat_o(tc, s[0])
                    for x in s[1:]:
                        p2 = copy.deepcopy(tc.findall(qn("w:p"))[0])
                        for c in list(p2):
                            if c.tag != qn("w:pPr"):
                                p2.remove(c)
                        r0 = tc.findall(qn("w:p"))[0].find(".//" + qn("w:r"))
                        them_run(p2, copy.deepcopy(r0.find(qn("w:rPr"))) if r0.find(qn("w:rPr")) is not None else None,
                                 x, bo_dam=True)
                        tc.append(p2)
                else:
                    self.dat_o(tc, ("**" + s + "**") if dam else s, can="center" if dam else "left")
            return tr

        # Hàng 0 và 1 của mẫu là tiêu đề gộp dọc; hàng 2 là hàng dữ liệu.
        for x in list(trs[0].iter(qn("w:vMerge"))):
            x.getparent().remove(x)
        tbl.append(hang_moi(trs[0], tieu_de, True))
        for h in hang:
            tbl.append(hang_moi(trs[2], h, False))
        return tbl

    def chen_sau(self, neo, ds):
        for el in ds:
            neo.addnext(el)
            neo = el
        return neo

    def noi_dung(self, ds):
        ra = []
        for muc in ds:
            if muc[0] == "BANG":
                _, cap, td, hang, nguon, rong = muc
                ra.append(self.doan("H", cap, can_giua=True))
                ra.append(self.bang_moi(td, hang, rong))
                ng = self.doan("P", "*" + nguon + "*")
                ng.find(qn("w:pPr")).find(qn("w:ind")).set(qn("w:firstLine"), "0")
                ra.append(ng)
            else:
                ra.append(self.doan(*muc))
        return ra

    # ---- các phần ----
    def bia(self):
        doi = {"ĐỀ TÀI KHOA HỌC CÔNG NGHỆ CẤP CƠ SỞ NĂM...": "ĐỀ TÀI KHOA HỌC CÔNG NGHỆ CẤP CƠ SỞ NĂM 2026",
               "TÊN ĐỀ TÀI": TEN.upper(),
               "Chủ nhiệm đề tài: ………………………………………": "Chủ nhiệm đề tài: Nguyễn Thị Tố Uyên",
               "Hà Nội, ........…": "Hà Nội, 2026"}
        for t in self.body.iter(qn("w:t")):
            if t.text in doi:
                t.text = doi[t.text]

    def phan_1(self):
        p = self.tim("1.1. Tên đề tài:")
        them_run(p, p.find(qn("w:r")).find(qn("w:rPr")), TEN, bo_dam=True)
        tbl = self.body.findall(qn("w:tbl"))[0]
        trs = tbl.findall(qn("w:tr"))
        for tr, o in zip(trs[2:4], [["1", "ThS. Nguyễn Thị Tố Uyên", "Viện Quản trị và Công nghệ", "Chủ nhiệm đề tài"],
                                    ["2", "Trần Đăng Bộ", "Viện Ngôn ngữ - Văn hóa - Quốc tế", "Thành viên"]]):
            for tc, s in zip(tr.findall(qn("w:tc")), o):
                self.dat_o(tc, s, can="center" if s in ("1", "2") else "left")
        for tr in trs[4:]:
            tbl.remove(tr)
        p = self.tim("1.4. Đơn vị chủ trì:")
        them_run(p, p.find(qn("w:r")).find(qn("w:rPr")), "Trường Đại học Thành Đô", bo_dam=True)
        p = self.tim("1.5.1. Theo đề xuất:")
        p.findall(qn("w:r"))[-1].find(qn("w:t")).text = "từ tháng 3 năm 2026 đến tháng 11 năm 2026"
        hd = self.tim("(Về mục tiêu, nội dung, phương pháp")
        self.chen_sau(hd, self.noi_dung(THAY_DOI))
        self.body.remove(hd)
        p = self.tim("1.7. Tổng kinh phí")
        t0 = p.findall(qn("w:r"))[0].find(qn("w:t"))
        t0.text = "1.7. Tổng kinh phí được phê duyệt của đề tài: "
        t0.set(XML_SPACE, "preserve")
        for r in p.findall(qn("w:r"))[1:]:
            p.remove(r)
        them_run(p, p.find(qn("w:r")).find(qn("w:rPr")), "5,0 triệu đồng.", bo_dam=True)

    def phan_2(self):
        self.body.remove(self.tim("Viết theo cấu trúc một bài báo"))
        for td, nd in PHAN_II.items():
            self.chen_sau(self.tim(td), self.noi_dung(nd))

    def phan_3_4_5(self):
        tbls = self.body.findall(qn("w:tbl"))
        # Bảng 3.1: sản phẩm
        sp = [t for t in tbls if chu(t).startswith("TTSản phẩm Tình trạng")][0]
        trs = sp.findall(qn("w:tr"))
        dien = {
            "1.1": ["Không", "", "", ""], "2.1": ["Không", "", "", ""], "3.1": ["Không", "", "", ""],
            "4.1": ["Không", "", "", ""], "6.1": ["Không", "", "", ""],
            "5.1": ["Nguyễn Thị Tố Uyên, Trần Đăng Bộ. Quản lý quyền sở hữu trí tuệ trong trường đại học trước yêu cầu mới "
                    "của chính sách. Tạp chí Nghiên cứu Khoa học và Phát triển, Trường Đại học Thành Đô, số ……, năm ……, "
                    "tr. ……", "……", "Có", ""],
            "7.1": ["Báo cáo toàn văn đề tài và hệ thống giải pháp, đề xuất Phòng Khoa học Công nghệ dùng làm căn cứ sửa "
                    "Quy chế quản trị tài sản trí tuệ", "Đã hoàn thành báo cáo; dự kiến ứng dụng", "Có", ""],
            "7.2": ["Quy trình 8 khâu, phiếu khai báo, phiếu rà soát, danh mục số và bộ chỉ số theo dõi; đề xuất thí điểm "
                    "tại Viện Y - Dược từ quý IV/2026 đến quý II/2027", "Dự kiến ứng dụng", "Có", ""],
        }
        da_co = set()
        for tr in trs[1:]:
            tcs = tr.findall(qn("w:tc"))
            ma = chu(tcs[0]).strip()
            if len(tcs) == 5 and ma in dien and ma not in da_co:
                da_co.add(ma)
                for tc, s in zip(tcs[1:], dien[ma]):
                    self.dat_o(tc, s)
            elif len(tcs) == 5:
                sp.remove(tr)
        # Bảng 3.2: đào tạo
        dt = [t for t in tbls if chu(t).startswith("TTHọ và tên")][0]
        for tr in dt.findall(qn("w:tr")):
            tcs = tr.findall(qn("w:tc"))
            if len(tcs) == 5 and chu(tcs[0]).strip() == "1":
                self.dat_o(tcs[1], "Không")
        # Phần IV
        th = [t for t in tbls if chu(t).startswith("TTSản phẩmSố lượng đăng ký")][0]
        for tr in th.findall(qn("w:tr"))[1:]:
            tcs = tr.findall(qn("w:tc"))
            so = chu(tcs[0]).strip()
            dk, ht = {"5": ("01", ""), "7": ("01", "01")}.get(so, ("0", "0"))
            self.dat_o(tcs[2], dk, can="center")
            self.dat_o(tcs[3], ht, can="center")
        # Phần V
        kp = [t for t in tbls if chu(t).startswith("TTNội dung chi")][0]
        duyet = {"Chi phí trực tiếp": "5,0", "Thuê khoán chuyên môn": "4,5", "Hội nghị, Hội thảo": "0,3",
                 "In ấn, Văn phòng phẩm": "0,2", "Tổng số": "5,0"}
        for tr in kp.findall(qn("w:tr"))[1:]:
            tcs = tr.findall(qn("w:tc"))
            ten = chu(tcs[1]).strip()
            gt = next((v for k, v in duyet.items() if ten.startswith(k)), "0")
            self.dat_o(tcs[2], ("**" + gt + "**") if ten in ("Chi phí trực tiếp", "Tổng số") else gt, can="center")
        gc = self.doan("P", "*Ghi chú: thuê khoán chuyên môn gồm xây dựng thuyết minh 0,5 triệu đồng, thù lao thành viên "
                            "2,0 triệu đồng, viết báo cáo tổng kết 0,5 triệu đồng và 3 chuyên đề 1,5 triệu đồng; hội nghị "
                            "là chi nghiệm thu 0,3 triệu đồng; theo thuyết minh đề tài. Cột kinh phí thực hiện do chủ nhiệm "
                            "đề tài điền theo chứng từ.*")
        gc.find(qn("w:pPr")).find(qn("w:ind")).set(qn("w:firstLine"), "0")
        kp.addnext(gc)

    def phan_6_7(self):
        self.chen_sau(self.tim("PHẦN VI. KIẾN NGHỊ"), self.noi_dung(KIEN_NGHI))
        self.chen_sau(self.tim("PHẦN VII. PHỤ LỤC"), self.noi_dung([
            ("P", "1. Bản sao bài báo là sản phẩm của đề tài, kèm trang bìa và mục lục số tạp chí."),
            ("P", "2. Báo cáo toàn văn đề tài."),
        ]))

    def lam(self):
        self.bia()
        self.phan_1()
        self.phan_2()
        self.phan_3_4_5()
        self.phan_6_7()


def main():
    t = TomTat()
    t.lam()
    t.doc.save(RA)
    import tempfile
    import pymupdf
    tmp = tempfile.mkdtemp()
    d = pymupdf.open(ve_pdf(RA, tmp))
    print(f"Đã lưu {RA}: {len(d)} trang")


if __name__ == "__main__":
    main()

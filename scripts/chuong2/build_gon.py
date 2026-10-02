# -*- coding: utf-8 -*-
"""Dựng Chương 2 bản rút gọn, có biểu đồ Excel gốc nhúng trong Word.

Chạy từ thư mục gốc của kho:
    python3 scripts/chuong2/build_gon.py

Đầu vào : Chuong_2_Thuc_trang_hoan_chinh mới.docx (chỉ lấy định dạng trang, kiểu chữ)
          scripts/chuong2/du_lieu.py, bieu_do.py (số liệu và cấu hình biểu đồ đã chuẩn hóa)
Đầu ra  : Ban_cuoi/Chuong_2_Thuc_trang_quan_ly_quyen_SHTT.docx
          Du_lieu_bieu_do_Chuong_2.xlsx

Nội dung chương được viết trực tiếp trong tệp này (hàm noi_dung). Mọi con số
trích từ du_lieu.py hoặc bieu_do.py được kiểm tra bằng assert trước khi ghi.
"""
import copy
import os
import re
import sys

import docx
from docx.oxml.ns import qn
from lxml import etree

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import bieu_do as B  # noqa: E402
import build as BD  # noqa: E402
import du_lieu as D  # noqa: E402

GOC = BD.GOC
VAO = BD.VAO
RA_DOCX = os.path.join(GOC, "Ban_cuoi", "Chuong_2_Thuc_trang_quan_ly_quyen_SHTT.docx")
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

so, pt = B.so, B.pt

# ---------------------------------------------------------------------------
# Số liệu dẫn xuất dùng trong văn bản
# ---------------------------------------------------------------------------
NL = D.NHAN_LUC
tong_nl = sum(r[2] for r in NL)
tong_gv = sum(r[9] for r in NL)
ts_tong = sum(r[5] for r in NL)
ths_tong = sum(r[6] for r in NL)
hoc_ham = sum(r[3] + r[4] for r in NL)
assert (tong_nl, tong_gv, ts_tong, ths_tong, hoc_ham) == (252, 145, 94, 94, 25)
assert (B.nl_ba_vien, B.ts_ba_vien, B.hoc_ham_ba_vien, B.nl_khoi_qt, B.dh_khac_khoi_qt) == (146, 81, 22, 29, 23)

ghph = B.dem_q[0] + B.dem_q[1] + B.dem_q[4]
assert ghph == 8

kh_dat = sum(1 for r in B.kh_rows if r[4] >= 1)
assert kh_dat == 6 and len(B.kh_rows) == 9 and len(B.kh_chua_xd) == 1
kh = B.kh
tong_nam = B.tong_nam
assert tong_nam == [56, 100, 85, 128, 192] and B.tong_ban_ghi == 582
assert B.bb[0] == 23 and B.bb[4] == 161
q_tong = sum(B.co_hang)
assert q_tong == 71 and sum(B.co_hang[3:]) == 68
assert (B.tong_T, B.tong_M, B.tong_C) == (4, 5, 7)
assert (len(B.DE_TAI), len(B.dt_du_dk), len(B.nop_don), len(B.du_dk_den_2024)) == (38, 11, 1, 8)
assert (len(B.dt_shcn), len(B.dt_shcn_den_2024), len(B.dt_qtg), len(B.dt_den_2024)) == (9, 6, 2, 31)
dt_ydd = [d for d in B.dt_du_dk if d[2] == "Viện Y - Dược"]
assert len(dt_ydd) == 9
assert (D.NHAN_SU_CO_TEN, D.NHAN_SU_CO_BAI, D.BAI_KHOP) == (246, 87, 256)
assert len(D.TSTT_KY) == 11 and sum(1 for t in D.TSTT_KY if t[4] == D.DA_CAP) == 4
assert len(B.dt_co_tien) == 19 and B.kp_tong == 424.75 and B.kp_du_dk == 396.75
xep_loai = {k: sum(1 for d in D.DE_TAI if d[3] == k) for k in ("Xuất sắc", "Tốt", "Khá", "Đạt")}
assert xep_loai == {"Xuất sắc": 1, "Tốt": 19, "Khá": 1, "Đạt": 17}, xep_loai


# ---------------------------------------------------------------------------
# Khung văn bản: lấy định dạng từ bản gốc rồi xóa toàn bộ nội dung
# ---------------------------------------------------------------------------
class VanBan:
    def __init__(self):
        self.doc = docx.Document(VAO)
        d = self.doc
        ps = d.paragraphs

        def mau(dau):
            ds = [p for p in ps if p.text.strip().startswith(dau)]
            assert ds, dau
            return copy.deepcopy(ds[0]._p)

        self.mau = {
            "chuong1": copy.deepcopy(ps[0]._p), "chuong2": copy.deepcopy(ps[1]._p),
            "chuong3": copy.deepcopy(ps[2]._p),
            "h1": mau("2.1. "), "h2": mau("2.1.1. "), "h3": mau("a) Nguyên nhân khách quan") if any(
                p.text.strip().startswith("a) Nguyên nhân khách quan") for p in ps) else mau("2.1.1. "),
            "than": mau("Trường Đại học Thành Đô là cơ sở"),
            "tieu_de": mau("Bảng 2.1. "), "nguon": mau("Nguồn: Tổng hợp của tác giả từ danh sách"),
        }
        self.tblPr = copy.deepcopy(d.tables[0]._tbl.tblPr)
        body = d.element.body
        for el in list(body):
            if el.tag != qn("w:sectPr"):
                body.remove(el)
        # gỡ quan hệ tới ảnh cũ không còn dùng
        for rid, rel in list(d.part.rels.items()):
            if rel.reltype.endswith("/image"):
                del d.part.rels[rid]
        self.body = body
        self.sect = body.find(qn("w:sectPr"))
        self.so_bang = 0
        self.so_hinh = 0
        self.dem_bd = 0  # đếm biểu đồ nhúng trong toàn văn bản, dùng đặt tên phần chart
        self.hinh = []
        self.ket_qua = {}

    def _them(self, el):
        self.sect.addprevious(el)
        return el

    def doan(self, loai, text, bold=None, italic=None, giu=False):
        el = copy.deepcopy(self.mau[loai])
        runs = el.findall(qn("w:r"))
        mau_r = copy.deepcopy(runs[0]) if runs else etree.Element(qn("w:r"))
        for x in list(el):
            if x.tag != qn("w:pPr"):
                el.remove(x)
        for t in mau_r.findall(qn("w:t")) + mau_r.findall(qn("w:br")) + mau_r.findall(qn("w:drawing")):
            mau_r.remove(t)
        el.append(mau_r)
        p = docx.text.paragraph.Paragraph(el, self.doc._body)
        r = p.runs[0]
        r.text = text
        if bold is not None:
            r.bold = bold
        if italic is not None:
            r.italic = italic
        if giu:
            p.paragraph_format.keep_with_next = True
        self._them(el)
        return p

    def than(self, *ds):
        for t in ds:
            self.doan("than", t)

    # --- bảng -------------------------------------------------------------
    def bang(self, tieu_de, cot, dong, nguon, rong, can=None, dong_tong=False):
        self.so_bang += 1
        self.doan("tieu_de", f"Bảng 2.{self.so_bang}. {tieu_de}", bold=True, giu=True)
        tbl = etree.SubElement(etree.Element("x"), qn("w:tbl"))
        tbl.append(copy.deepcopy(self.tblPr))
        grid = etree.SubElement(tbl, qn("w:tblGrid"))
        tong = sum(rong)
        tw = [int(9070 * r / tong) for r in rong]  # 16 cm vùng chữ
        for w in tw:
            etree.SubElement(grid, qn("w:gridCol")).set(qn("w:w"), str(w))
        can = can or (["left"] + ["center"] * (len(cot) - 1))
        for i, hang in enumerate([cot] + dong):
            tr = etree.SubElement(tbl, qn("w:tr"))
            trpr = etree.SubElement(tr, qn("w:trPr"))
            etree.SubElement(trpr, qn("w:cantSplit"))
            if i == 0:
                etree.SubElement(trpr, qn("w:tblHeader"))
            dau = i == 0
            cuoi = dong_tong and i == len(dong)
            for j, v in enumerate(hang):
                tc = etree.SubElement(tr, qn("w:tc"))
                tcpr = etree.SubElement(tc, qn("w:tcPr"))
                x = etree.SubElement(tcpr, qn("w:tcW"))
                x.set(qn("w:w"), str(tw[j]))
                x.set(qn("w:type"), "dxa")
                if dau:
                    shd = etree.SubElement(tcpr, qn("w:shd"))
                    shd.set(qn("w:val"), "clear")
                    shd.set(qn("w:color"), "auto")
                    shd.set(qn("w:fill"), "DCE8F7")
                etree.SubElement(tcpr, qn("w:vAlign")).set(qn("w:val"), "center")
                p = etree.SubElement(tc, qn("w:p"))
                ppr = etree.SubElement(p, qn("w:pPr"))
                sp_ = etree.SubElement(ppr, qn("w:spacing"))
                sp_.set(qn("w:before"), "20")
                sp_.set(qn("w:after"), "20")
                sp_.set(qn("w:line"), "252")
                sp_.set(qn("w:lineRule"), "auto")
                ind = etree.SubElement(ppr, qn("w:ind"))
                ind.set(qn("w:firstLine"), "0")
                etree.SubElement(ppr, qn("w:jc")).set(qn("w:val"), "center" if dau else
                                                      {"left": "left", "center": "center"}[can[j]])
                r = etree.SubElement(p, qn("w:r"))
                rpr = etree.SubElement(r, qn("w:rPr"))
                if dau or cuoi:
                    etree.SubElement(rpr, qn("w:b"))
                sz = etree.SubElement(rpr, qn("w:sz"))
                sz.set(qn("w:val"), "24")
                t = etree.SubElement(r, qn("w:t"))
                t.text = str(v)
                t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        self._them(tbl)
        self.doan("nguon", nguon, italic=True)
        return self.so_bang

    # --- hình -------------------------------------------------------------
    def _ten_hinh(self, tien_to, so):
        return f"Hình {tien_to}.{so}" if tien_to else f"Hình {so}"

    def hinh_bd(self, ma, nguon=None, tieu_de=None, tien_to="2"):
        h = next(x for x in B.HINH if x["id_cu"] == ma)
        self.so_hinh += 1
        self.dem_bd += 1
        h["so"] = self.so_hinh
        h["id"] = f"H{tien_to or 'BB'}.{self.so_hinh}"
        if nguon:
            h["nguon"] = nguon
        if tieu_de:
            h["tieu_de"] = tieu_de
        self.hinh.append(h)
        self.doan("tieu_de", f"{self._ten_hinh(tien_to, h['so'])}. {h['tieu_de']}", bold=True, giu=True)
        p = self.doan("tieu_de", "", giu=True)
        for r in p._p.findall(qn("w:r")):
            p._p.remove(r)
        pf = p.paragraph_format
        pf.space_before = 0
        pf.space_after = 0
        pf.line_spacing = 1.0
        p._p.append(BD.nhung_bieu_do(self.doc, h, self.dem_bd))
        self.doan("nguon", h["nguon"], italic=True)
        return self.so_hinh

    def so_do(self, tieu_de, anh, nguon, tien_to="2", rong_cm=15.5):
        """Chèn sơ đồ khái niệm dạng ảnh, đánh số chung với hình."""
        from docx.shared import Cm
        self.so_hinh += 1
        self.doan("tieu_de", f"{self._ten_hinh(tien_to, self.so_hinh)}. {tieu_de}", bold=True, giu=True)
        p = self.doan("tieu_de", "", giu=True)
        for r in p._p.findall(qn("w:r")):
            p._p.remove(r)
        pf = p.paragraph_format
        pf.space_before = 0
        pf.space_after = 0
        pf.line_spacing = 1.0
        p.add_run().add_picture(anh, width=Cm(rong_cm))
        self.doan("nguon", nguon, italic=True)
        return self.so_hinh


for _h in B.HINH:
    _h["id_cu"] = _h["id"]


# ---------------------------------------------------------------------------
# Nội dung chương
# ---------------------------------------------------------------------------
def noi_dung(v):
    v.doan("chuong1", "CHƯƠNG 2")
    v.doan("chuong2", "THỰC TRẠNG QUẢN LÝ QUYỀN SỞ HỮU TRÍ TUỆ")
    v.doan("chuong3", "TẠI TRƯỜNG ĐẠI HỌC THÀNH ĐÔ")
    v.than(
        "Chương 2 phân tích thực trạng quản lý quyền sở hữu trí tuệ tại Trường Đại học Thành Đô giai đoạn 2021 - 2025 "
        "theo chu trình tạo lập, xác lập, bảo vệ và khai thác đã xác định tại Chương 1, đặt trong khung thể chế, tổ "
        "chức và nguồn lực của Nhà trường. Dữ liệu được tổng hợp từ các danh mục thống kê của Phòng Khoa học Công "
        "nghệ, danh sách nhân sự năm 2026, danh mục tài sản trí tuệ, các quy chế nội bộ và Kế hoạch số 07/KH-ĐHTĐ; "
        "văn bản pháp luật được đối chiếu theo bản hiện hành đến tháng 9 năm 2026. Văn bản ban hành sau năm 2025 chỉ "
        "được dùng để mô tả bối cảnh và yêu cầu hiện hành, không dùng để giải thích kết quả giai đoạn 2021 - 2025. Năm "
        "của mỗi số liệu được ghi rõ tại từng bảng, hình: năm theo mã số hoặc năm phê duyệt đối với đề tài, năm nghiệm "
        "thu, năm nộp đơn hoặc năm cấp văn bằng đối với tài sản trí tuệ.")

    # ===================================================================== 2.1
    v.doan("h1", "2.1. Khái quát về Nhà trường và nguồn hình thành tài sản trí tuệ")
    v.doan("h2", "2.1.1. Tổ chức và nhân lực")
    v.than(
        "Trường Đại học Thành Đô là cơ sở giáo dục đại học tư thục định hướng ứng dụng, đào tạo mười tám ngành do ba "
        "viện quản lý: Viện Quản trị và Công nghệ chín ngành, trong đó có ngành Quản lý kinh tế ở trình độ thạc sĩ và "
        "tiến sĩ; Viện Ngôn ngữ - Văn hóa - Quốc tế sáu ngành; Viện Y - Dược ba ngành. Bộ máy gồm ba khối dưới Ban Giám "
        "hiệu. Khối Quản trị và Dịch vụ có Phòng Khoa học Công nghệ, Phòng Tài chính - Kế toán, Trung tâm Tuyển sinh và "
        "Quản trị thương hiệu, Trung tâm Dịch vụ và Quản trị hành chính tổng hợp cùng một số đơn vị khác. Khối Đào tạo "
        "và Nghiên cứu gồm ba viện đào tạo, Viện Nghiên cứu giáo dục và Chuyển giao tri thức, Tạp chí Nghiên cứu Khoa "
        "học và Phát triển và các trung tâm. Khối Doanh nghiệp có Công ty Thadotek. Trường Tiểu học và Trung học cơ sở "
        "UNIGO trong hệ sinh thái giáo dục là pháp nhân riêng, nên tài sản trí tuệ của trường này không thuộc phạm vi "
        "phân tích.",
        f"Năm 2026, Nhà trường có {tong_nl} nhân sự, trong đó {tong_gv} giảng viên. Cơ cấu theo đơn vị và trình độ "
        "được trình bày tại Bảng 2.1.")
    dong = []
    ten_bang = {"Trung tâm Tuyển sinh và Quản trị thương hiệu": "Trung tâm Tuyển sinh và Quản trị thương hiệu",
                "Trung tâm Dịch vụ và Quản trị hành chính tổng hợp": "Trung tâm Dịch vụ và Quản trị hành chính "
                                                                      "tổng hợp"}
    for r in NL:
        dong.append([ten_bang.get(r[0], r[0]), r[2], r[9], r[5], r[3] + r[4], r[6], r[7] + r[8]])
    dong.append(["Tổng cộng", tong_nl, tong_gv, ts_tong, hoc_ham, ths_tong, sum(r[7] + r[8] for r in NL)])
    v.bang("Nhân lực Trường Đại học Thành Đô theo đơn vị và trình độ, năm 2026",
           ["Đơn vị", "Tổng số", "Giảng viên", "Tiến sĩ và tương đương", "Trong đó GS, PGS",
            "Thạc sĩ và tương đương", "Đại học và khác"], dong,
           "Nguồn: Nhóm nghiên cứu tổng hợp từ danh sách nhân sự năm 2026 của Trường Đại học Thành Đô. Tiến sĩ và "
           "tương đương gồm người có học hàm giáo sư, phó giáo sư và bác sĩ chuyên khoa II; thạc sĩ và tương đương gồm "
           "dược sĩ chuyên khoa I.",
           [5.2, 1.4, 1.6, 1.8, 1.5, 1.8, 1.6], dong_tong=True)
    v.than(
        f"Bảng 2.1 cho thấy nguồn nhân lực trình độ cao tập trung ở các viện đào tạo: ba viện chiếm "
        f"{pt(B.nl_ba_vien / tong_nl)} nhân sự nhưng có {B.ts_ba_vien} trên {ts_tong} người trình độ tiến sĩ, tức "
        f"{pt(B.ts_ba_vien / ts_tong)}, và {B.hoc_ham_ba_vien} trên {hoc_ham} người có học hàm. Đây là lực lượng có khả "
        "năng nhận diện kết quả nghiên cứu có thể bảo hộ. Các thủ tục xác lập quyền, quản trị thương hiệu và pháp chế "
        "được giao cho các đơn vị thuộc khối Quản trị và Dịch vụ, nơi đội ngũ được bố trí theo yêu cầu nghiệp vụ hành "
        "chính và dịch vụ. Phòng Khoa học Công nghệ, đơn vị đầu mối theo quy chế, có 2 nhân sự trình độ thạc sĩ đồng "
        "thời đảm nhiệm nhiều mảng quản lý khoa học, nên chưa có vị trí chuyên trách về sở hữu trí tuệ. Cách phân công "
        "này hợp lý về chức năng, song đòi hỏi một cơ chế phối hợp chặt chẽ giữa hai khối, như phân tích tại Mục 2.2.2.")

    # ---------------------------------------------------------------- 2.1.2
    v.doan("h2", "2.1.2. Sản phẩm khoa học giai đoạn 2021 - 2025")
    v.than("Các danh mục thống kê của Phòng Khoa học Công nghệ ghi nhận 582 bản ghi sản phẩm khoa học trong giai đoạn "
           f"2021 - 2025, được trình bày tại Bảng 2.{v.so_bang + 1}.")
    dong = []
    for ten, vals in D.SAN_PHAM:
        dong.append([ten] + [x if x else "-" for x in vals] + [sum(vals)])
    dong.insert(6, ["Tham luận hội thảo quốc gia", "-", "-", "-", "-", "-", D.THAM_LUAN_QUOC_GIA])
    dong.append(["Tổng cộng"] + tong_nam + [B.tong_ban_ghi])
    b_sp = v.bang("Sản phẩm khoa học của Trường Đại học Thành Đô, giai đoạn 2021 - 2025",
                  ["Loại sản phẩm", "2021", "2022", "2023", "2024", "2025", "Tổng"], dong,
                  "Nguồn: Nhóm nghiên cứu tổng hợp từ các danh mục thống kê của Phòng Khoa học Công nghệ. Danh mục tham "
                  "luận hội thảo quốc gia không có trường thời gian nên chỉ có số tổng. Đề tài cấp cơ sở xếp theo năm ghi "
                  "trong mã số; đề tài cấp quốc gia xếp theo năm phê duyệt kinh phí. Các dòng có đơn vị thống kê khác "
                  "nhau nên tổng cộng là tổng số bản ghi, không phải số sản phẩm độc lập.",
                  [5.4, 1.2, 1.2, 1.2, 1.2, 1.2, 1.3], dong_tong=True)
    v.than(
        f"Con số 582 là tổng số bản ghi của các danh mục, không phải số sản phẩm độc lập, vì các danh mục dùng đơn vị "
        "thống kê khác nhau: bài báo, sách, giáo trình là ấn phẩm, còn đề tài là nhiệm vụ, và kết quả của một đề tài có "
        "thể đồng thời xuất hiện dưới dạng bài báo hoặc tham luận. Hồ sơ nghiệm thu cho thấy 26 trên 38 đề tài khai tổng "
        "cộng 32 công bố, nhưng danh mục bài báo không ghi mã đề tài nên chưa xác lập được bản ghi nào trùng với bản ghi "
        "nào. Đề tài vì vậy giữ con số 582 như tổng số bản ghi và không quy đổi thành số sản phẩm độc lập.")
    h_sp = v.hinh_bd("H2.2")
    v.than(
        f"Hình 2.{h_sp} cho thấy số bản ghi tăng từ {tong_nam[0]} năm 2021 lên {tong_nam[4]} năm 2025, bình quân "
        f"{pt(B.cagr(tong_nam[0], tong_nam[4], 4))} một năm, và mức tăng này chủ yếu đến từ bài báo. Bài báo tăng từ "
        f"{B.bb[0]} lên {B.bb[4]} bài, gấp {so(B.bb[4] / B.bb[0], 1)} lần, nâng tỷ trọng từ {pt(B.bb[0] / tong_nam[0])} "
        f"lên {pt(B.bb[4] / tong_nam[4])} số bản ghi; giáo trình giảm từ {B.gt[0]} xuống {B.gt[4]} tài liệu, tỷ trọng từ "
        f"{pt(B.gt[0] / tong_nam[0])} còn {pt(B.gt[4] / tong_nam[4])}. Chất lượng công bố quốc tế cũng tăng: "
        f"{q_tong} trên 115 bài quốc tế đăng trên tạp chí có phân hạng Q, trong đó {sum(B.co_hang[3:])} bài thuộc hai "
        "năm 2024 - 2025.",
        "Đối với quản lý quyền sở hữu trí tuệ, cơ cấu này có hai hàm ý. Thứ nhất, bài báo chỉ phát sinh quyền tác giả "
        "và hầu như không có khả năng khai thác thương mại, nên tăng công bố không tự động tạo ra tài sản trí tuệ có "
        "giá trị khai thác. Thứ hai, trong cùng giai đoạn, số công bố tăng mạnh nhưng số đơn đăng ký sở hữu công nghiệp "
        "không tăng tương ứng. Dữ liệu này chưa đủ để xác định nguyên nhân, song cho thấy cần xem xét các yếu tố về quy "
        "trình và cơ chế bên cạnh năng lực nghiên cứu, như phân tích tại Mục 2.2.3 và Mục 2.3.3.",
        "Trong hai năm cuối kỳ, Nhà trường chủ trì ba đề tài cấp quốc gia do Quỹ Phát triển khoa học và công nghệ quốc "
        "gia tài trợ, tổng kinh phí 4,67 tỷ đồng, được phê duyệt ngày 26 tháng 3 năm 2024, ngày 30 tháng 9 năm 2025 và "
        "ngày 15 tháng 12 năm 2025, đều đang thực hiện. Theo điểm c khoản 1 Điều 86 Luật Sở hữu trí tuệ được bổ sung "
        "bởi Luật số 131/2025/QH15, tổ chức được giao quyền quản lý, sử dụng, quyền sở hữu kết quả của nhiệm vụ sử dụng "
        "ngân sách nhà nước có quyền đăng ký sáng chế, kiểu dáng công nghiệp, thiết kế bố trí là kết quả của nhiệm vụ "
        "đó. Về lợi ích của tác giả khi thương mại hóa, ba đề tài không chịu cùng một chế độ, như phân tích tại Mục "
        "2.2.1. Đây là nguồn tài sản trí tuệ tiềm năng của giai đoạn tới, với điều kiện việc sàng lọc khả năng bảo hộ "
        "được thực hiện trước khi công bố kết quả.",
        f"Tính trên toàn bộ nhân sự, gồm cả khối hành chính, {D.NHAN_SU_TONG} người trong danh sách năm 2026 tương ứng "
        f"{D.NHAN_SU_CO_TEN} tên khác nhau do có {D.NHAN_SU_TONG - D.NHAN_SU_CO_TEN} trường hợp trùng họ tên. Theo quy "
        f"tắc so khớp giữ dấu và đúng thứ tự họ tên, {D.NHAN_SU_CO_BAI} trên {D.NHAN_SU_CO_TEN} tên, tức "
        f"{pt(D.NHAN_SU_CO_BAI / D.NHAN_SU_CO_TEN)}, đứng tên ít nhất một bài báo trong năm năm; hệ số Gini về số bài "
        f"là {so(D.GINI_BAI, 3)}, và trong nhóm có công bố, 10% người dẫn đầu chiếm {pt(D.TOP10_BAI)} số lượt đứng "
        f"tên. Nếu chấp nhận cả cách viết không dấu trong bài quốc tế, số tên có bài là {D.DO_NHAY_B['co_bai']}, tức "
        f"{pt(D.DO_NHAY_B['co_bai'] / D.NHAN_SU_CO_TEN)}, và hệ số Gini là {so(D.DO_NHAY_B['gini'], 3)}; nhận định về "
        "mức độ tập trung không đổi giữa hai quy tắc. Lực lượng nghiên cứu nòng cốt vì vậy còn tương đối mỏng so với quy "
        f"mô Nhà trường. Viện Y - Dược là nơi phát sinh {len(dt_ydd)} trên 11 sản phẩm đề tài có tiềm năng tạo lập tài "
        "sản trí tuệ và đơn sáng chế duy nhất trong kỳ; số lượng công bố và khả năng hình thành tài sản trí tuệ là hai "
        "thước đo khác nhau, cần được theo dõi riêng.")

    # ===================================================================== 2.2
    v.doan("h1", "2.2. Thực trạng thể chế, tổ chức và nguồn lực quản lý quyền sở hữu trí tuệ")
    v.doan("h2", "2.2.1. Hệ thống quy định nội bộ")
    v.than(
        "Nhà trường có bốn văn bản chứa quy định về quyền sở hữu trí tuệ. Quy chế hoạt động khoa học công nghệ ban hành "
        "kèm Quyết định số 213/QĐ-ĐHTĐ ngày 28 tháng 12 năm 2021, sau đây gọi là Quyết định 213, dành Chương VI cho sở "
        "hữu trí tuệ và chuyển giao công nghệ: Điều 34 liệt kê phạm vi tài sản khá đầy đủ, từ tên trường, nhãn hiệu, "
        "sáng chế, giải pháp hữu ích, kiểu dáng công nghiệp đến giáo trình, ngân hàng đề thi, phần mềm và quy trình công "
        "nghệ; Điều 35 quy định quy trình đăng ký một cửa, trong đó Phòng Khoa học Công nghệ kiểm tra đơn, trình Hiệu "
        "trưởng ký và thực hiện việc nộp đơn, lệ phí tại cơ quan nhà nước; Điều 36 quy định phân chia nguồn thu khi "
        "chuyển giao. Ngày 21 tháng 11 năm 2024, Nhà trường ban hành Quy chế quản trị tài sản trí tuệ kèm Quyết định số "
        "217/QĐ-ĐHTĐ, sau đây gọi là Quyết định 217, mở rộng phạm vi tới cơ sở dữ liệu, giáo trình điện tử, bí quyết và "
        "tên miền. Điều 10 Quyết định 217 xác định quyền công bố kết quả nghiên cứu thuộc về Trường, yêu cầu tác giả xin "
        "ý kiến Phòng Khoa học Công nghệ trước khi bộc lộ công khai tài sản có thể bảo hộ và đặt nghĩa vụ bảo mật; Điều "
        "11 giao Phòng xây dựng quy trình, biểu mẫu ghi nhận, khai báo, nhận diện, lập hồ sơ theo dõi, xúc tiến thương "
        "mại hóa và giao Bộ phận Pháp chế thực hiện thủ tục xác lập quyền; Điều 14 quy định các hành vi xâm phạm quyền "
        "tác giả như mạo danh, công bố trái phép, trích dẫn không đầy đủ. Hai văn bản còn lại là Quy chế chi tiêu nội bộ "
        "ban hành ngày 01 tháng 8 năm 2026 và Điều lệ Quỹ Học bổng sau tiến sĩ Ngô Xuân Độ năm 2025. Quy chế chi tiêu "
        "nội bộ năm 2026 ban hành sau kỳ đánh giá nên chỉ được dùng để mô tả cơ chế hiện hành; quy chế chi tiêu áp dụng "
        "trong giai đoạn 2021 - 2025 và văn bản Điều lệ Quỹ chưa có trong hồ sơ Đề tài tiếp cận, nên các nội dung về Quỹ "
        "được ghi theo thông tin được cung cấp.",
        "Các văn bản này được ban hành ở những thời điểm khác nhau, cho những kênh tài trợ khác nhau và phù hợp với khung "
        "pháp luật tại thời điểm ban hành. Việc Nhà trường ban hành quy chế về sở hữu trí tuệ từ năm 2021 và quy chế "
        "chuyên biệt từ năm 2024, trước khi Luật Khoa học, công nghệ và đổi mới sáng tạo số 93/2025/QH15 và Luật số "
        "131/2025/QH15 sửa đổi Luật Sở hữu trí tuệ ra đời, cho thấy tầm nhìn sớm của Nhà trường. Do Quyết định 217 không "
        "dẫn chiếu và không thay thế Chương VI Quyết định 213, hai văn bản cùng hiệu lực. Bảng "
        f"2.{v.so_bang + 1} tập hợp các quy định nội bộ và quy định của Luật liên quan đến lợi ích của tác giả và phân "
        "chia nguồn thu, theo phạm vi, đối tượng hưởng, cơ sở tính và điều kiện áp dụng.")
    b_ll = v.bang(
        "Các quy định về lợi ích của tác giả và phân chia nguồn thu từ tài sản trí tuệ",
        ["Văn bản, điều khoản", "Phạm vi, nguồn kinh phí", "Đối tượng hưởng và tỷ lệ", "Cơ sở tính",
         "Điều kiện, mức trần"],
        [list(r) for r in D.QUY_DINH_LOI_ICH],
        "Nguồn: Nhóm nghiên cứu tổng hợp từ Quyết định 213, Quyết định 217, Quy chế chi tiêu nội bộ năm 2026, Luật số "
        "93/2025/QH15 và Văn bản hợp nhất số 67/VBHN-VPQH. Nội dung về Quỹ Học bổng sau tiến sĩ Ngô Xuân Độ ghi theo "
        "thông tin được cung cấp, chưa đối chiếu được văn bản gốc.",
        [3.0, 3.4, 3.4, 3.2, 3.4], can=["left", "left", "left", "left", "left"])
    v.than(
        f"Bảng 2.{b_ll} cho thấy các tỷ lệ trong văn bản nội bộ và trong Luật không thể so sánh trực tiếp với nhau, vì "
        "chúng khác nhau về đối tượng hưởng, cơ sở tính và phạm vi. Phần 30% tại điểm a là khen thưởng tập thể tác giả "
        "tính trên nguồn thu sau chi phí; mức tối thiểu 30% tại Điều 28 Luật số 93/2025/QH15 là thưởng tính trên lợi "
        "nhuận sau thuế; mức 15% tại Điều 135 Luật Sở hữu trí tuệ là thù lao tính trên tổng số tiền nhận được trước thuế. "
        "Chỉ điểm a và điểm b khoản 4 Điều 36 Quyết định 213 có cùng cơ sở tính, nên được mô phỏng chung tại hình dưới "
        "đây.")
    h_mp = v.hinh_bd("H2.8")
    v.than(
        f"Hình 2.{h_mp} cho thấy, trên cùng nguồn thu sau chi phí, phần dành cho tác giả theo hai điểm bằng nhau cho đến "
        "khoảng 333 triệu đồng; từ mức này, phần theo điểm a không tăng thêm do mức trần 100 triệu đồng một đề tài. Vì "
        "hai điểm áp dụng cho hai phạm vi khác nhau, khác biệt này không phải là xung đột giữa hai quy định mà thể hiện "
        "tác động của mức trần đối với đề tài sử dụng ngân sách nhà nước.",
        "Đối chiếu theo từng tình huống, chưa xác định được trường hợp nào mà cùng một tài sản đồng thời chịu hai yêu cầu "
        "nội bộ không tương thích: điểm a áp dụng cho đề tài sử dụng ngân sách nhà nước, điểm b cho tài sản thuộc sở hữu "
        "của Trường, còn khoản 3 Điều 13 Quyết định 217 tự xác định thứ tự áp dụng khi chỉ được dùng nếu không có thỏa "
        "thuận, Quy chế chi tiêu nội bộ hoặc pháp luật không quy định. Các điểm cần thống nhất cách hiểu và thứ tự áp "
        "dụng gồm: quan hệ giữa Điều 13 Quyết định 217 với khoản 4 Điều 36 Quyết định 213, vì văn bản sau không dẫn chiếu "
        "văn bản trước; quan hệ giữa quy định trích 50% kinh phí chuyển giao công nghệ về Nhà trường tại Quy chế chi tiêu "
        "nội bộ năm 2026 với phần 30% của tác giả tại điểm b; và thuật ngữ, khi điểm a gọi phần của tác giả là khen "
        "thưởng còn điểm b ghi tác giả hưởng 30% theo Luật Sở hữu trí tuệ, trong khi pháp luật phân biệt thưởng theo "
        "Điều 28 Luật số 93/2025/QH15 với thù lao theo Điều 135 Luật Sở hữu trí tuệ.",
        "Đối chiếu với Luật số 93/2025/QH15, có hiệu lực từ ngày 01 tháng 10 năm 2025, và Luật Sở hữu trí tuệ hiện hành "
        "cho thấy ba điểm cần rà soát. Thứ nhất, với nhiệm vụ sử dụng ngân sách nhà nước được giao từ ngày 01 tháng 10 "
        "năm 2025, và với sáng chế, kiểu dáng công nghiệp, thiết kế bố trí đã được cấp văn bằng là kết quả của nhiệm vụ "
        "giao từ ngày 01 tháng 01 năm 2023 theo khoản 7 Điều 73, phần thưởng cho tác giả phải đạt tối thiểu 30% lợi "
        "nhuận sau thuế; mức trần 100 triệu đồng tại điểm a chỉ dẫn đến phần thưởng thấp hơn mức này khi 30% lợi nhuận "
        "sau thuế vượt 100 triệu đồng, còn khoản nộp ngân sách nhà nước 40% không thuộc các mục đích sử dụng lợi nhuận "
        "tại khoản 3 Điều 28. Thứ hai, ba đề tài cấp quốc gia không chịu cùng một chế độ: đề tài phê duyệt ngày 15 tháng "
        "12 năm 2025 thuộc phạm vi Điều 28; hai đề tài phê duyệt ngày 26 tháng 3 năm 2024 và ngày 30 tháng 9 năm 2025 "
        "tiếp tục theo pháp luật và văn bản có hiệu lực tại thời điểm phê duyệt theo khoản 3 Điều 73, trừ trường hợp kết "
        "quả là sáng chế, kiểu dáng công nghiệp, thiết kế bố trí được cấp văn bằng theo khoản 7 Điều 73. Thứ ba, với "
        "sáng chế, kiểu dáng công nghiệp, thiết kế bố trí, điểm b tính phần 30% của tác giả trên nguồn thu sau chi phí, "
        "trong khi mức mặc định tại khoản 1 Điều 135 tính trên tổng số tiền nhận được trước thuế; khi chi phí vượt một "
        "nửa nguồn thu, 30% sau chi phí thấp hơn 15% tổng số tiền. Do đó cần xác định quy chế có được coi là thỏa thuận "
        "về thù lao theo khoản 1 Điều 135 hay không. Việc cập nhật quy chế vì vậy cần căn cứ vào phạm vi từng trường hợp, "
        "không thể kết luận mức trần trái luật trong mọi trường hợp, cũng không thể thay bằng một tỷ lệ chung.")

    v.doan("h2", "2.2.2. Tổ chức bộ máy và đầu mối quản lý")
    v.than(
        "Theo Điều 35 Quyết định 213 và Điều 11 Quyết định 217, Phòng Khoa học Công nghệ là đầu mối tiếp nhận hồ sơ, "
        "nhận diện, lập hồ sơ theo dõi và xúc tiến thương mại hóa tài sản trí tuệ; Bộ phận Pháp chế thực hiện thủ tục "
        "xác lập quyền và hỗ trợ đăng ký; các đơn vị có trách nhiệm phối hợp ghi nhận tài sản mới phát sinh. Trong hồ "
        "sơ hiện có, công việc được phân theo bốn đơn vị, như thể hiện tại hình dưới đây. Phòng Khoa học Công nghệ, với "
        "2 nhân sự, tiếp nhận hồ sơ đề tài. Danh mục theo dõi nhãn hiệu và logo do bộ phận quản trị thương hiệu thuộc "
        "Trung tâm Tuyển sinh và Quản trị thương hiệu lập. Bộ phận Pháp chế thuộc Trung tâm Dịch vụ và Quản trị hành "
        "chính tổng hợp; theo thông tin được cung cấp, bộ phận này là đầu mối của Mạng lưới Trung tâm Hỗ trợ công nghệ "
        "và đổi mới sáng tạo mà Nhà trường tham gia từ năm 2023, nhưng hồ sơ Đề tài tiếp cận chưa có văn bản xác nhận. "
        "Viện Nghiên cứu giáo dục và Chuyển giao tri thức là đơn vị có thống kê hợp đồng khai thác quyền đối với sách.")
    h_bdm = v.so_do("Phân công đầu mối quản lý quyền sở hữu trí tuệ tại Trường Đại học Thành Đô",
                    os.path.join(GOC, "Ban_cuoi", "so_do", "bon_dau_moi.png"),
                    "Nguồn: Nhóm nghiên cứu tổng hợp từ Quyết định 213, Quyết định 217 và danh sách nhân sự năm 2026.")
    v.than(
        f"Hình 2.{h_bdm} cho thấy ba trong bốn đầu mối thuộc khối Quản trị và Dịch vụ, còn kết quả nghiên cứu phát sinh "
        "ở khối Đào tạo và Nghiên cứu. Quyết định 217 đã phân công trách nhiệm, trong đó Điều 11 giao Phòng Khoa học "
        "Công nghệ xây dựng quy trình, biểu mẫu ghi nhận, khai báo và lập hồ sơ theo dõi tài sản trí tuệ. Trong hồ sơ Đề "
        "tài tiếp cận chưa thấy biểu mẫu khai báo hay hồ sơ theo dõi dùng chung giữa các đơn vị; mức độ thực hiện các "
        "nhiệm vụ tại Điều 10 và Điều 11 vì vậy là vấn đề cần được đánh giá, chưa thể kết luận là chưa được thực hiện.")

    v.doan("h2", "2.2.3. Nguồn lực tài chính và cơ chế khuyến khích")
    v.than(f"Nhà trường có ba kênh tài trợ nghiên cứu với chế độ sở hữu trí tuệ khác nhau, được trình bày tại Bảng "
           f"2.{v.so_bang + 1}.")
    b_kenh = v.bang(
        "Ba kênh tài trợ nghiên cứu và cơ chế chi cho xác lập quyền",
        ["Kênh tài trợ", "Quy mô", "Cơ chế chi cho xác lập quyền", "Kết quả tài sản trí tuệ trong kỳ"],
        [["Đề tài cấp cơ sở, ngân sách của Trường",
          f"38 đề tài giai đoạn 2021 - 2025; 19 đề tài được cấp tổng {so(B.kp_tong, 2)} triệu đồng, 16 đề tài chỉ "
          "quy đổi giờ, 3 đề tài tự tìm tài trợ",
          "Điều 35 Quyết định 213 giao Phòng Khoa học Công nghệ nộp đơn, lệ phí; Điều 38 cho phép chi thuê ngoài, chi "
          "khác liên quan trực tiếp; chưa thấy dự toán riêng",
          f"11 đề tài có sản phẩm tiềm năng, trong đó {len(B.dt_shcn)} thuộc sở hữu công nghiệp; 1 đơn sáng chế"],
         ["Quỹ Học bổng sau tiến sĩ Ngô Xuân Độ",
          "Ngân sách 5 tỷ đồng giai đoạn 2025 - 2029; 10 đến 15 học bổng mỗi năm, mức 50 đến 500 triệu đồng",
          "Khoản chi phí khác tối đa bằng 50% chi trực tiếp", "Quỹ hoạt động từ năm 2025, chưa có kết quả"],
         ["Đề tài cấp quốc gia, Quỹ Phát triển khoa học và công nghệ quốc gia",
          "3 đề tài, tổng 4,67 tỷ đồng, phê duyệt năm 2024 và 2025", "Theo quy định của cơ quan tài trợ",
          "Đang thực hiện, chưa nghiệm thu"]],
        "Nguồn: Nhóm nghiên cứu tổng hợp từ danh mục đề tài cấp cơ sở, danh mục đề tài cấp quốc gia và Quyết định 213. "
        "Thông tin về Quỹ Học bổng sau tiến sĩ Ngô Xuân Độ ghi theo tài liệu được cung cấp, chưa đối chiếu được Điều lệ.",
        [3.4, 4.6, 4.2, 3.4], can=["left", "left", "left", "left"])
    v.than(
        f"Bảng 2.{b_kenh} cho thấy ba kênh khác nhau về quy mô và cơ chế. Đề tài cấp cơ sở, kênh duy nhất trong kỳ đã "
        "tạo ra sản phẩm có tiềm năng tạo lập tài sản trí tuệ, được Nhà trường cấp "
        f"{so(B.kp_tong, 2)} triệu đồng trong năm năm. Quy chế đã có căn cứ chi cho bước xác lập quyền: Điều 35 Quyết "
        "định 213 giao Phòng Khoa học Công nghệ thực hiện việc nộp đơn và lệ phí tại cơ quan nhà nước, Điều 38 cho phép "
        "chi thuê ngoài và chi khác liên quan trực tiếp đến đề tài, Điều 13 Quyết định 217 coi lệ phí xác lập quyền là "
        "khoản được trừ khi chia lợi ích. Điều chưa thấy trong hồ sơ là dự toán riêng cho phí nộp đơn, phí đại diện, "
        "phí duy trì hiệu lực; cách tạm ứng và thanh toán khi đơn được nộp sau khi đề tài đã nghiệm thu; thời hạn và "
        "người chịu trách nhiệm đề xuất chi. Số liệu chi thực tế cho xác lập quyền trong kỳ cũng chưa được thống kê. Theo "
        "thông tin được cung cấp, Quỹ Học bổng sau tiến sĩ Ngô Xuân Độ có khoản chi phí khác có thể trang trải chi phí "
        "đăng ký nhưng chỉ dành cho đối tượng học bổng và hoạt động từ năm 2025; chưa có quy định chuyển tiếp giữa hai "
        "kênh.")
    h_kp = v.hinh_bd("H2.10")
    v.than(
        f"Hình 2.{h_kp} cho thấy cơ chế cấp kinh phí đề tài cấp cơ sở thay đổi sau năm 2022: toàn bộ {B.dt_gio[0]} đề "
        f"tài mã số 2021 và {B.dt_gio[1]} trên {len(B.dt_nam['2022'])} đề tài mã số 2022 chỉ được quy đổi giờ nghiên "
        f"cứu; từ năm 2023, phần lớn đề tài được cấp tiền và tổng kinh phí đạt {so(B.kp_nam[4], 0)} triệu đồng năm "
        f"2025. Mức cấp phân tán mạnh, trung vị {so(B.kp_trung_vi, 1)} triệu đồng, cao nhất 80 triệu đồng. Cả 11 đề tài "
        f"có sản phẩm tiềm năng đều thuộc nhóm {len(B.dt_co_tien)} đề tài được cấp tiền và dùng {so(B.kp_du_dk, 2)} "
        f"triệu đồng, tức {pt(B.kp_du_dk / B.kp_tong)} tổng kinh phí. Mối liên hệ này phù hợp với đặc điểm của các sản "
        "phẩm thực nghiệm, chủ yếu ở lĩnh vực dược, cần kinh phí vật tư; dữ liệu không đủ để kết luận kinh phí bằng tiền "
        "là nguyên nhân tạo ra sản phẩm. Hồ sơ hiện có cũng không cho biết dự toán của các đề tài này có khoản dành cho "
        "bước đăng ký hay không.",
        "Về khuyến khích, Quy chế chi tiêu nội bộ ban hành ngày 01 tháng 8 năm 2026 quy định kinh phí xét duyệt không "
        "quá 50 triệu đồng cho đề tài có đăng ký sở hữu trí tuệ và quy đổi giờ nghiên cứu cho văn bằng, với điều kiện "
        "đã có chứng nhận đăng ký thành công. Vì quy chế này ban hành sau kỳ đánh giá, phân tích dưới đây mô tả cơ chế "
        "hiện hành và có ý nghĩa đối với giai đoạn tới; định mức được đặt cạnh chế độ dành cho công bố tại hình dưới "
        "đây.")
    h_kk = v.hinh_bd("H2.11")
    v.than(
        f"Hình 2.{h_kk} cho thấy, xét riêng giờ quy đổi, văn bằng được định giá cao: sáng chế chuẩn Việt Nam được tính "
        "360 giờ, cao hơn bài WoS hạng Q1 với 300 giờ. Nhưng chỉ bài báo quốc tế được thưởng tiền, từ 10 đến 20 triệu "
        "đồng một bài, và được ghi nhận ngay khi đăng; văn bằng không có tiền thưởng và chỉ được ghi nhận khi được cấp, "
        "tức sau các giai đoạn thẩm định hình thức, công bố đơn và thẩm định nội dung theo Điều 119 Luật Sở hữu trí tuệ. "
        "Với cơ chế này, công bố trước là lựa chọn có lợi hơn về thời gian ghi nhận đối với giảng viên. Theo khoản 3 và "
        "khoản 4 Điều 60 Luật Sở hữu trí tuệ, đơn phải được nộp trong thời hạn mười hai tháng kể từ ngày bộc lộ thì sáng "
        "chế mới không bị coi là mất tính mới. Cấu trúc khuyến khích này vì vậy là một rủi ro cần xử lý cho giai đoạn "
        "tới; do quy chế chi tiêu áp dụng giai đoạn 2021 - 2025 chưa có trong hồ sơ, Đề tài không dùng cơ chế năm 2026 "
        "để giải thích kết quả của kỳ đánh giá.")

    v.doan("h2", "2.2.4. Hệ thống dữ liệu phục vụ quản lý")
    v.than(
        "Trong các danh mục thống kê của Phòng Khoa học Công nghệ mà Đề tài tiếp cận, gồm bài báo, sách, giáo trình, đề "
        "tài và tham luận, chưa có danh mục theo dõi tài sản trí tuệ đã đăng ký hoặc được cấp văn bằng, trong khi Điều "
        "11 Quyết định 217 giao Phòng lập hồ sơ thống kê, theo dõi. Danh mục tài sản trí tuệ hiện có là tệp theo dõi đơn "
        "và văn bằng gồm hai bảng chưa thống nhất về trạng thái của 5 kiểu dáng công nghiệp, và không liên kết với danh "
        "mục đề tài. Hồ sơ nghiệm thu cho thấy 26 trên 38 đề tài khai tổng cộng 32 công bố; do danh mục bài báo không ghi "
        "mã đề tài, chưa thể xác định số bài báo thực sự phát sinh từ đề tài để đánh giá mức độ khai báo. Dữ liệu cũng "
        "chưa chuẩn hóa: tên đơn vị ghi nhiều cách, danh mục bài báo không có trường đơn vị, biểu mẫu giáo trình năm 2021 "
        "khác các năm sau, và năm của một đề tài có thể là năm ghi trong mã số, năm phê duyệt hoặc năm nghiệm thu.")
    h_tc = v.hinh_bd("H2.7")
    v.than(
        f"Hình 2.{h_tc} lượng hóa hệ quả: trong 16 tiêu chí đánh giá hiệu quả quản lý quyền sở hữu trí tuệ xây dựng "
        f"tại Mục 1.4.2, chỉ {B.tong_T} tiêu chí tính được đầy đủ, {B.tong_M} tiêu chí tính được một phần và "
        f"{B.tong_C} tiêu chí chưa tính được; nhóm kết quả không có tiêu chí nào tính được đầy đủ. Nhà trường đo được số "
        "đơn và giấy chứng nhận đã có nhưng chưa đo được giá trị mà các quyền mang lại. Từ năm 2026, đây là nhóm thông "
        "tin phải công khai theo điểm đ khoản 3 Điều 28 Luật Giáo dục đại học số 125/2025/QH15.")

    # ===================================================================== 2.3
    v.doan("h1", "2.3. Thực trạng tạo lập, xác lập và bảo vệ quyền sở hữu trí tuệ")
    v.doan("h2", "2.3.1. Đối sánh giữa quy định pháp luật, quy chế nội bộ và thực tế phát sinh")
    v.than(f"Bảng 2.{v.so_bang + 1} đối sánh ba lớp: đối tượng quyền theo Luật Sở hữu trí tuệ, tài sản được liệt kê tại "
           "Điều 34 Quyết định 213 và Điều 3 Quyết định 217, và thực tế phát sinh giai đoạn 2021 - 2025.")
    b_ds = v.bang(
        "Đối sánh đối tượng quyền theo Luật, theo quy chế nội bộ và theo thực tế phát sinh",
        ["Nhóm quyền", "Đối tượng", "Quy chế có liệt kê", "Thực tế phát sinh", "Tình trạng xác lập, đăng ký"],
        [["Quyền tác giả", "Bài báo, sách", "Có", "405 bài báo, 12 sách", "Quyền phát sinh tự động; chưa đăng ký"],
         ["Quyền tác giả", "Giáo trình, tài liệu giảng dạy", "Có", "87 tài liệu",
          "Quyền phát sinh tự động; chưa đăng ký"],
         ["Quyền tác giả", "Chương trình đào tạo, ngân hàng đề thi, phần mềm", "Có", "Chưa thống kê", "Chưa thống kê"],
         ["Quyền tác giả", "Sưu tập dữ liệu", "Có, tại Quyết định 217", "Bộ mẫu cây thuốc, bộ tiêu bản",
          "Chưa đăng ký"],
         ["Quyền tác giả", "Logo, bộ nhận diện", "Có", "Có", "2 giấy chứng nhận"],
         ["Sở hữu công nghiệp", "Sáng chế", "Có", "Có", "1 đơn năm 2025; 1 đơn năm 2026"],
         ["Sở hữu công nghiệp", "Giải pháp hữu ích", "Có", "Công thức, quy trình từ đề tài", "Chưa có đơn"],
         ["Sở hữu công nghiệp", "Kiểu dáng công nghiệp", "Có", "Có",
          "5 hồ sơ có số hiệu văn bằng, trạng thái chưa xác minh"],
         ["Sở hữu công nghiệp", "Nhãn hiệu", "Có", "Có", "2 văn bằng, 1 đơn"],
         ["Sở hữu công nghiệp", "Bí mật kinh doanh", "Có, tại Quyết định 217", "Công thức chưa công bố",
          "Không áp dụng đăng ký"]],
        "Nguồn: Nhóm nghiên cứu tổng hợp từ Luật Sở hữu trí tuệ, Quyết định 213, Quyết định 217 và các danh mục của "
        "Trường Đại học Thành Đô. Quyền tác giả phát sinh khi tác phẩm được định hình; chưa đăng ký không có nghĩa là "
        "chưa có quyền. Bảng không liệt kê quyền liên quan, thiết kế bố trí, chỉ dẫn địa lý và giống cây trồng do chưa "
        "phát sinh hoặc không thuộc phạm vi hoạt động của Nhà trường.",
        [2.6, 3.6, 2.6, 3.6, 3.2], can=["left", "left", "left", "left", "left"])
    v.than(
        f"Bảng 2.{b_ds} cho thấy Nhà trường phát sinh tài sản ở nhiều nhóm đối tượng nhưng mới có văn bằng, giấy chứng "
        "nhận hoặc đơn ở bốn nhóm. Giải pháp hữu ích và sưu tập dữ liệu, hai nhóm mà sản phẩm đề tài cấp cơ sở có tiềm "
        "năng nhiều nhất, đều đã có trong quy chế nhưng chưa có hồ sơ đăng ký. Giáo trình chưa được đăng ký quyền tác "
        "giả; quyền tác giả phát sinh từ khi tác phẩm được định hình, còn giấy chứng nhận đăng ký là chứng cứ phục vụ "
        "quản lý, chuyển giao và giải quyết tranh chấp. Khoảng trống vì vậy nằm ở khâu nhận diện và sàng lọc, không nằm "
        "ở phạm vi văn bản.")

    v.doan("h2", "2.3.2. Tài sản trí tuệ đã được xác lập quyền và đang xử lý")
    v.than("Tệp theo dõi đơn và văn bằng ghi nhận 11 hồ sơ tài sản trí tuệ thuộc sở hữu hoặc đồng sở hữu của Nhà trường "
           "phát sinh trong kỳ 2021 - 2025 và 1 đơn sáng chế nộp năm 2026, được liệt kê tại "
           f"Bảng 2.{v.so_bang + 1} theo bốn nhóm trạng thái thống nhất: đã nộp đơn, đã chấp nhận đơn hợp lệ, đã cấp văn "
           "bằng hoặc giấy chứng nhận, và chưa xác minh.")
    dong = [[i + 1, t[0], t[1], t[2], t[3], t[4], "Trường" if t[5] == "Trường" else "Đồng sở hữu"]
            for i, t in enumerate(D.TSTT)]
    b_ts = v.bang(
        "Hồ sơ tài sản trí tuệ thuộc sở hữu hoặc đồng sở hữu của Trường Đại học Thành Đô",
        ["TT", "Loại hình", "Tên tài sản", "Số đơn, số hiệu", "Năm", "Trạng thái", "Chủ sở hữu"], dong,
        "Nguồn: Nhóm nghiên cứu tổng hợp từ tệp theo dõi đơn nhãn hiệu, kiểu dáng công nghiệp và thống kê văn bằng của "
        "Trường Đại học Thành Đô. Năm là năm cấp văn bằng, giấy chứng nhận đối với hồ sơ đã cấp và năm nộp đơn đối với "
        "hồ sơ đã nộp đơn; năm của 5 kiểu dáng ghi theo bảng thống kê văn bằng. Năm kiểu dáng có số hiệu văn bằng tại "
        "bảng thống kê văn bằng nhưng được ghi chờ cấp bằng tại bảng theo dõi đơn nên được xếp nhóm chưa xác minh. Bảng "
        "theo dõi đơn có một dòng sáng chế ghi chấp nhận đơn hợp lệ nhưng không ghi số đơn nên chưa gán được cho hồ sơ "
        "nào. Hồ sơ số 12 nộp năm 2026, ngoài kỳ đánh giá. Không gồm tài sản của pháp nhân UNIGO và của cá nhân.",
        [0.9, 2.2, 4.0, 3.4, 1.2, 2.4, 2.0], can=["center", "left", "left", "left", "center", "left", "left"])
    h_ts = v.hinh_bd("H2.12")
    v.than(
        f"Bảng 2.{b_ts} và Hình 2.{h_ts} cho thấy 11 hồ sơ trong kỳ hình thành từ ba luồng. Luồng thương hiệu gồm 4 tài "
        "sản do Nhà trường đơn sở hữu, đều đã được cấp văn bằng hoặc giấy chứng nhận: hai giấy chứng nhận quyền tác giả "
        "đối với bộ logo và hai văn bằng nhãn hiệu. Luồng hợp tác doanh nghiệp gồm 5 kiểu dáng công nghiệp và 1 đơn nhãn "
        "hiệu đồng sở hữu với một doanh nghiệp đối tác; đây là thành công nổi bật trong chiến lược hợp tác đại học với "
        "doanh nghiệp mà Ban Giám hiệu đã dày công kết nối. Năm kiểu dáng đã có số hiệu văn bằng tại bảng thống kê, "
        "nhưng trạng thái cần được xác nhận trước khi tính là văn bằng đã cấp. Luồng nghiên cứu có 1 đơn sáng chế nộp năm "
        "2025 trong kỳ và 1 đơn nộp năm 2026. Như vậy, số văn bằng, giấy chứng nhận xác định được trong kỳ là 4; nếu 5 "
        "kiểu dáng được xác nhận, con số này là 9. Tổng số hồ sơ không được dùng như tổng số quyền đã được cấp.",
        "Giai đoạn tới, Nhà trường có thể tận dụng đà hợp tác này để phát triển thêm tài sản do chính Nhà trường đơn sở "
        "hữu từ kết quả nghiên cứu, nhất là công thức và quy trình, những đối tượng mà kiểu dáng công nghiệp không bảo "
        "hộ. Giải pháp hữu ích, loại hình phù hợp với nhiều sản phẩm đề tài cấp cơ sở, chưa có hồ sơ nào. Về bảo vệ "
        "quyền, hồ sơ Đề tài tiếp cận không ghi nhận tranh chấp, khiếu nại hay xử lý xâm phạm liên quan đến Nhà trường; "
        "điều này chưa đủ để kết luận không có hành vi xâm phạm, vì Nhà trường chưa có cơ chế theo dõi xâm phạm. Nếu 5 "
        "kiểu dáng được xác nhận cấp cùng năm, thời hạn gia hạn sẽ trùng nhau, nên cần được đưa vào sổ theo dõi thống "
        "nhất.")

    v.doan("h2", "2.3.3. Sản phẩm đề tài có tiềm năng tạo lập tài sản trí tuệ")
    v.than("Rà soát cột sản phẩm nghiệm thu của 38 đề tài cấp cơ sở cho thấy 11 đề tài có sản phẩm cụ thể ngoài báo cáo "
           "và bài báo có tiềm năng tạo lập tài sản trí tuệ, được liệt kê tại "
           f"Bảng 2.{v.so_bang + 1}. Đây là đánh giá sơ bộ dựa trên mô tả sản phẩm; khả năng bảo hộ thực tế còn phụ "
           "thuộc vào tính mới, trình độ sáng tạo, khả năng áp dụng và tình trạng bộc lộ của từng sản phẩm, chưa được "
           "thẩm định.")
    NT = D.NGAY_NGHIEM_THU
    b_dt = v.bang(
        "Đề tài cấp cơ sở có sản phẩm tiềm năng tạo lập tài sản trí tuệ, giai đoạn 2021 - 2025",
        ["Mã số", "Đơn vị chủ trì", "Sản phẩm nghiệm thu", "Nhóm quyền dự kiến", "Nghiệm thu", "Tình trạng"],
        [["37-2022", "Khoa Công nghệ kỹ thuật Ô tô", "Mô hình kiểm tra và sửa chữa hệ thống khởi động điện trên ô tô",
          "Kiểu dáng công nghiệp hoặc giải pháp hữu ích", NT["37-2022"], "Chưa có đơn"],
         ["39-2022", "Khoa Dược", "Sản phẩm cầm máu dạng màng", "Giải pháp hữu ích hoặc sáng chế", NT["39-2022"],
          "Chưa có đơn"],
         ["40-2022", "Khoa Dược", "Bộ mẫu cây thuốc", "Quyền tác giả, sưu tập dữ liệu; cần đánh giá thêm",
          NT["40-2022"], "Chưa đăng ký"],
         ["08-2023", "Viện Nghiên cứu giáo dục và Chuyển giao tri thức",
          "Tinh dầu Citrus grandis, hồ sơ ghi chuyển giao cho Trường thương mại hóa", "Giải pháp hữu ích",
          NT["08-2023"], "Chưa có đơn"],
         ["06-2024", "Khoa Dược", "Xà phòng hữu cơ tảo biển kèm quy trình và công thức", "Giải pháp hữu ích",
          NT["06-2024"], "Chưa có đơn"],
         ["07-2024", "Khoa Dược", "100 tiêu bản hiển vi và cẩm nang phát hiện giun sán",
          "Quyền tác giả, sưu tập dữ liệu", NT["07-2024"], "Chưa đăng ký"],
         ["08-2024", "Khoa Dược", "Rutin độ tinh khiết 90% kèm quy trình chiết xuất, tinh chế",
          "Giải pháp hữu ích hoặc sáng chế", NT["08-2024"], "Chưa có đơn"],
         ["09-2024", "Khoa Dược", "Thang thuốc 100 ml kèm công thức và quy trình sắc", "Giải pháp hữu ích",
          NT["09-2024"], "Chưa có đơn"],
         ["06-2025", "Viện Y - Dược", "Công thức cồn thuốc xoa bóp, hồ sơ ghi dự kiến đăng ký giải pháp hữu ích",
          "Giải pháp hữu ích", NT["06-2025"], "Chưa có đơn"],
         ["07-2025", "Viện Y - Dược", "Quy trình bào chế viên ngậm giảm ho, hồ sơ ghi dự kiến đăng ký giải pháp hữu ích",
          "Giải pháp hữu ích", NT["07-2025"], "Chưa có đơn"],
         ["09-2025", "Viện Y - Dược", "Chiết xuất lá Quế hoa, hồ sơ ghi giai đoạn 2 là sáng chế", "Sáng chế",
          NT["09-2025"], "Đã nộp đơn năm 2025"]],
        "Nguồn: Nhóm nghiên cứu tổng hợp từ danh mục đề tài khoa học công nghệ cấp cơ sở giai đoạn 2021 - 2025. Nhóm "
        "quyền dự kiến là đánh giá sơ bộ dựa trên mô tả sản phẩm trong hồ sơ nghiệm thu; tình trạng đối chiếu với tệp "
        "theo dõi đơn và văn bằng của Nhà trường.",
        [1.4, 2.8, 4.6, 3.2, 1.8, 1.9], can=["center", "left", "left", "left", "center", "left"])
    h_nq = v.hinh_bd("H2.14", nguon=f"Nguồn: Nhóm nghiên cứu tổng hợp từ Bảng 2.{b_dt}.")
    v.than(
        f"Hình 2.{h_nq} cho thấy {B.shcn} trên 11 sản phẩm thuộc nhóm sở hữu công nghiệp và giải pháp hữu ích là hình "
        f"thức dự kiến phù hợp với {ghph} sản phẩm. Loại hình này không đòi hỏi trình độ sáng tạo như sáng chế, có thời "
        "hạn bảo hộ ngắn hơn và phù hợp với quy mô đề tài cấp cơ sở, nhưng chưa có hồ sơ nào được nộp theo hình thức "
        "này. Mười trên 11 sản phẩm thuộc lĩnh vực dược, nên tiềm năng sở hữu công nghiệp của Nhà trường có thể được "
        "quản lý có trọng tâm.")
    assert len(B.dt_duoc) == 10
    h_ch = v.hinh_bd("H2.15", nguon=(
        f"Nguồn: Nhóm nghiên cứu tính toán từ danh mục đề tài cấp cơ sở và Bảng 2.{b_dt}. Chuỗi chỉ xét sở hữu công "
        "nghiệp, không gồm 2 sản phẩm thuộc quyền tác giả. Cột thứ hai gồm các đề tài mang mã số từ năm 2021 đến năm "
        "2024, đều nghiệm thu trước ngày 31 tháng 5 năm 2025."))
    v.than(
        f"Hình 2.{h_ch} cho thấy, trong chuỗi sở hữu công nghiệp, {len(B.dt_shcn)} trên {len(B.DE_TAI)} đề tài, tức "
        f"{pt(len(B.dt_shcn) / len(B.DE_TAI))}, có sản phẩm tiềm năng, và {len(B.nop_don)} trên {len(B.dt_shcn)} đề "
        f"tài, tức {pt(len(B.nop_don) / len(B.dt_shcn))}, đã có đơn. Với {len(B.dt_den_2024)} đề tài mang mã số 2021 - "
        f"2024, đều nghiệm thu trước ngày 31 tháng 5 năm 2025, {len(B.dt_shcn_den_2024)} đề tài có sản phẩm tiềm năng "
        "sở hữu công nghiệp nhưng chưa đề tài nào có đơn trong tệp theo dõi. Đây là điểm đứt gãy của chuỗi. Trong điều "
        "kiện hồ sơ hiện có, Đề tài nhận định điểm nghẽn nằm ở một khoảng trống kỹ thuật tại khâu nối giữa nghiệm thu và "
        "đăng ký: thiếu một biểu mẫu rà soát tại thời điểm nghiệm thu. Các biểu mẫu đề xuất, thuyết minh, hợp đồng và "
        "nghiệm thu tại Quyết định 213 đã có mục về đăng ký sở hữu trí tuệ ở dạng sản phẩm dự kiến, tiêu chí chấm điểm "
        "và tình trạng đăng ký, nhưng chưa yêu cầu sàng lọc khả năng bảo hộ và tình trạng bộc lộ của các sản phẩm khác. "
        "Với sản phẩm đã bộc lộ công khai quá mười hai tháng, khả năng đăng ký sáng chế, giải pháp hữu ích không còn "
        "theo khoản 3 Điều 60 Luật Sở hữu trí tuệ, nên khâu sàng lọc cần được thiết lập sớm.",
        "Đề tài chiết xuất lá Quế hoa mã số 09-2025, nghiệm thu ngày 25 tháng 12 năm 2025, là đề tài cấp cơ sở duy nhất "
        "trong kỳ có sản phẩm được nộp đơn sáng chế; đơn số 1-2025-07378 được nộp trong năm 2025. Trường hợp này cho "
        "thấy kênh chuyển hóa có thể vận hành trong điều kiện hiện có, nhưng mới có một trường hợp, do một chủ nhiệm "
        "đồng thời chủ trì đề tài cấp quốc gia, nên chưa đủ để coi là một quy trình. Hai đề tài 06-2025 và 07-2025, "
        "nghiệm thu ngày 25 tháng 12 năm 2025, ghi dự kiến đăng ký giải pháp hữu ích nhưng chưa có đơn trong tệp theo "
        "dõi; nếu kết quả đã được bộc lộ, thời hạn mười hai tháng kể từ ngày bộc lộ cần được kiểm tra ngay. Đề tài tinh "
        "dầu Citrus grandis mã số 08-2023 có liên hệ về tên gọi với kiểu dáng tinh dầu vỏ bưởi đào, nhưng hồ sơ chưa đủ "
        "để khẳng định; kể cả khi liên hệ được xác nhận, kiểu dáng chỉ bảo hộ hình dáng bên ngoài, còn công thức và quy "
        "trình chiết xuất chưa được đăng ký. Hai danh mục không có trường liên kết, nên chưa thể theo dõi một sản phẩm "
        "nghiên cứu đến khi xác lập quyền.")

    # ===================================================================== 2.4
    v.doan("h1", "2.4. Thực trạng sử dụng, chuyển giao và khai thác tài sản trí tuệ")
    v.than(
        "Khai thác tài sản trí tuệ tại Nhà trường được xem xét ở ba mức: sử dụng nội bộ, chuyển giao trong hệ thống và "
        "khai thác có thu phí.",
        "Ở mức sử dụng nội bộ, 61 giáo trình và tài liệu giai đoạn 2022 - 2025 đều ghi số tín chỉ, tức gắn với học phần "
        "cụ thể; 26 tài liệu năm 2021 không có thông tin này do biểu mẫu chưa có trường số tín chỉ. Hồ sơ nghiệm thu của "
        "sáu đề tài năm 2023 ghi địa chỉ ứng dụng tại Trường, trong khi đề tài các năm khác phần lớn chỉ ghi báo cáo và "
        "bài báo; khác biệt này có thể phụ thuộc vào biểu mẫu nghiệm thu hơn là bản chất đề tài.",
        "Ở mức chuyển giao trong hệ thống, đề tài tinh dầu Citrus grandis năm 2023 do Viện Nghiên cứu giáo dục và "
        "Chuyển giao tri thức chủ trì ghi sản phẩm là chuyển giao công nghệ cho Nhà trường thương mại hóa, nhưng hồ sơ "
        "Đề tài tiếp cận không có hợp đồng, biên bản chuyển giao hay kết quả thương mại hóa kèm theo.",
        "Ở mức khai thác có thu phí, theo thống kê do Viện Nghiên cứu giáo dục và Chuyển giao tri thức cung cấp, giai "
        "đoạn 2023 - 2024 có năm hợp đồng. Hai hợp đồng chuyển giao quyền sử dụng tác phẩm đối với sách “Giáo dục và "
        "khoa học mở: Cẩm nang hướng dẫn” và “Từng bước nhập môn Nghiên cứu định lượng”, với mức 10% trên tổng số bán "
        "được, là giao dịch khai thác quyền tác giả. Ba hợp đồng còn lại là dịch vụ đào tạo, giảng dạy, tổng trị giá 100 "
        "triệu đồng, thuộc khai thác tri thức chuyên môn chứ không phải khai thác một tài sản đã xác lập quyền. Đề tài "
        "chưa tiếp cận bản hợp đồng để đối chiếu.",
        "Như vậy, khai thác tài sản trí tuệ hiện chủ yếu ở dạng phi thương mại, phục vụ đào tạo của chính Nhà trường. "
        "Khai thác có thu phí ghi nhận được mới có hai hợp đồng, do một đơn vị thực hiện và đối với sách chuyên khảo; "
        "các văn bằng, giấy chứng nhận hiện có chưa ghi nhận giao dịch chuyển giao quyền trong hồ sơ. Đây cũng là nội "
        "dung có dữ liệu mỏng nhất, vì hệ thống thống kê chưa theo dõi việc sử dụng và khai thác tài sản trí tuệ; bản "
        "thân việc thiếu dữ liệu này là một phát hiện về thực trạng quản lý.")

    # ===================================================================== 2.5
    v.doan("h1", "2.5. Đánh giá chung về công tác quản lý quyền sở hữu trí tuệ tại Trường Đại học Thành Đô")
    v.than(
        "Kết quả hai năm 2024 - 2025 được đối chiếu trước hết với Kế hoạch hoạt động khoa học công nghệ giai đoạn 2024 - "
        "2028 ban hành kèm Kế hoạch số 07/KH-ĐHTĐ ngày 01 tháng 7 năm 2024. Đây là thước đo khách quan vì do chính Nhà "
        "trường đặt ra, và 2024 - 2025 là hai năm đầu của kế hoạch.")
    h_kh = v.hinh_bd("H2.16")
    k = kh
    v.than(
        f"Hình 2.{h_kh} cho thấy {kh_dat} trên {len(B.kh_rows)} chỉ tiêu xác định được đạt hoặc vượt kế hoạch. Nhóm "
        f"công bố vượt xa chỉ tiêu: bài báo tạp chí khoa học quốc tế đạt {k['Bài báo tạp chí khoa học quốc tế'][3]} bài "
        f"so với {k['Bài báo tạp chí khoa học quốc tế'][2]} bài, tức "
        f"{pt(k['Bài báo tạp chí khoa học quốc tế'][4], 0)}; sách có chỉ số ISBN "
        f"{pt(k['Sách xuất bản có chỉ số ISBN'][4], 0)}; bài báo kỷ yếu hội thảo cấp Trường "
        f"{pt(k['Bài báo đăng kỷ yếu Hội thảo cấp Trường'][4], 0)}; bài báo trong nước "
        f"{pt(k['Bài báo đăng tạp chí NCKH và PT của trường và tạp chí trong nước'][4], 0)}; riêng bài báo hội thảo "
        f"khoa học quốc tế đạt {pt(k['Bài báo Hội thảo khoa học quốc tế'][4], 0)}. Nhóm đề tài và học liệu dao động "
        f"quanh kế hoạch, từ {pt(k['Giáo trình/TLTK LHNB được nghiệm thu'][4], 0)} với giáo trình, tài liệu tham khảo "
        f"lưu hành nội bộ được nghiệm thu đến {pt(k['Đề tài KHCN cấp Bộ/Nhà nước'][4], 0)} với đề tài cấp Bộ, Nhà "
        "nước. Chỉ tiêu chuyển giao công nghệ năm 2025 chưa đạt.",
        "Chỉ tiêu 1.11 “Công nhận sáng chế, phát minh, KDCN, bản quyền tác giả”, tổng 4 cho hai năm, chưa xác định được "
        "tỷ lệ thực hiện. Trong phạm vi chỉ tiêu, không gồm nhãn hiệu, hai năm 2024 - 2025 chỉ có 5 kiểu dáng công "
        "nghiệp ghi năm 2024 nhưng trạng thái chưa thống nhất; giấy chứng nhận quyền tác giả năm 2025 thuộc pháp nhân "
        f"UNIGO. Nếu 5 kiểu dáng được xác nhận cấp năm 2024, kết quả là {D.KH_1_11['neu_xac_nhan']} trên "
        f"{D.KH_1_11['ke_hoach']}, tức {pt(D.KH_1_11['neu_xac_nhan'] / D.KH_1_11['ke_hoach'], 0)}, và cả 5 đều hình "
        "thành qua hợp tác doanh nghiệp, không từ đề tài. Do Kế hoạch gộp sáng chế với kiểu dáng công nghiệp và quyền "
        "tác giả trong một chỉ tiêu, chỉ tiêu này có thể được hoàn thành mà không cần kết quả nghiên cứu nào được bảo hộ. "
        "Kế hoạch 07/KH-ĐHTĐ từng tự đánh giá giai đoạn 2019 - 2023 là chưa có công trình được chuyển giao công nghệ và "
        "tài sản trí tuệ còn hạn chế; sau hai năm thực hiện, nhận định này về cơ bản vẫn giữ nguyên.")

    v.doan("h2", "2.5.1. Kết quả đạt được")
    v.than(
        "Thứ nhất, Nhà trường có tầm nhìn sớm về thể chế sở hữu trí tuệ. Quyết định 213 năm 2021 đã dành một chương riêng "
        "với phạm vi tài sản khá đầy đủ, quy trình đăng ký một cửa và căn cứ chi cho thủ tục xác lập quyền; Quyết định "
        "217 ngày 21 tháng 11 năm 2024 tiếp tục ban hành quy chế chuyên biệt, mở rộng phạm vi tài sản, bổ sung nguyên tắc "
        "công bố và bảo mật, phân công đầu mối và quy định các hành vi xâm phạm quyền tác giả. Cả hai văn bản đều ra đời "
        "trước khi Luật số 93/2025/QH15 và Luật số 131/2025/QH15 được ban hành.",
        "Thứ hai, trong kỳ 2021 - 2025, Nhà trường có 11 hồ sơ tài sản trí tuệ thuộc sở hữu hoặc đồng sở hữu, trong đó "
        "4 tài sản đã được cấp văn bằng, giấy chứng nhận và 5 kiểu dáng công nghiệp đã có số hiệu văn bằng cần xác nhận "
        "trạng thái. Nhóm thương hiệu gồm tên trường, bộ nhận diện và thương hiệu hệ sinh thái được bảo hộ từ năm 2021. "
        "5 kiểu dáng công nghiệp và 1 đơn nhãn hiệu đồng sở hữu là thành công nổi bật trong chiến lược hợp tác đại học "
        "với doanh nghiệp mà Ban Giám hiệu đã dày công kết nối.",
        f"Thứ ba, năng lực nghiên cứu tăng nhanh: số bài báo tăng bình quân {pt(B.cagr(B.bb[0], B.bb[4], 4))} một năm, "
        f"đạt {so(B.bb[4] / tong_gv, 2)} bài trên một giảng viên năm 2025; {kh_dat} trên {len(B.kh_rows)} chỉ tiêu "
        "khoa học công nghệ xác định được của Kế hoạch 07/KH-ĐHTĐ đạt hoặc vượt; ba đề tài cấp quốc gia với tổng kinh "
        "phí 4,67 tỷ đồng được phê duyệt.",
        f"Thứ tư, đội ngũ có tiềm năng tạo lập tài sản trí tuệ, nhất là ở khối ngành Y - Dược: 11 trên 38 đề tài cấp cơ "
        f"sở có sản phẩm tiềm năng, trong đó {len(B.dt_shcn)} thuộc sở hữu công nghiệp; kênh chuyển hóa từ đề tài sang "
        "đơn sáng chế đã có trường hợp đầu tiên với đề tài chiết xuất lá Quế hoa năm 2025, và năm 2026 có thêm một đơn "
        "sáng chế.",
        "Thứ năm, các biểu mẫu đề xuất, thuyết minh, hợp đồng và nghiệm thu tại Quyết định 213 đã có mục về đăng ký sở "
        "hữu trí tuệ, tạo nền để bổ sung nội dung sàng lọc; từ năm 2025, theo thông tin được cung cấp, Nhà trường có "
        "thêm Quỹ Học bổng sau tiến sĩ Ngô Xuân Độ với cơ chế đồng sở hữu văn bằng.")

    v.doan("h2", "2.5.2. Hạn chế")
    v.than(
        f"Thứ nhất, kết quả nghiên cứu chưa được chuyển hóa thành quyền tương xứng với tiềm năng: {len(B.dt_shcn)} đề "
        f"tài có sản phẩm tiềm năng sở hữu công nghiệp nhưng mới {len(B.nop_don)} đề tài có đơn; "
        f"{len(B.dt_shcn_den_2024)} đề tài mã số 2021 - 2024 có sản phẩm tiềm năng sở hữu công nghiệp chưa có đơn, và "
        f"{len(B.dt_qtg)} sản phẩm thuộc quyền tác giả chưa được đăng ký.",
        "Thứ hai, các quy định nội bộ ban hành ở những thời điểm khác nhau chưa làm rõ phạm vi và thứ tự áp dụng giữa "
        "các quy định về lợi ích của tác giả, chưa phân biệt thưởng, thù lao và phần chia nguồn thu; điểm a khoản 4 Điều "
        "36 Quyết định 213 cần được rà soát theo Điều 28 và Điều 73 Luật số 93/2025/QH15 đối với các trường hợp thuộc "
        "phạm vi áp dụng của Luật mới.",
        "Thứ ba, luồng thương hiệu và luồng nghiên cứu vận hành tách biệt: danh mục tài sản và danh mục sản phẩm đề tài "
        "do hai bộ phận lập, chưa có trường liên kết, nên sản phẩm nghiên cứu chưa được theo dõi liên tục đến khi xác lập "
        "quyền.",
        "Thứ tư, tài sản do Nhà trường đơn sở hữu hình thành từ kết quả nghiên cứu còn ít: ngoài 4 tài sản thương hiệu, "
        "trong kỳ mới có 1 đơn sáng chế; các hồ sơ sở hữu công nghiệp khác đều hình thành qua hợp tác doanh nghiệp.",
        f"Thứ năm, hệ thống dữ liệu chưa đáp ứng đầy đủ yêu cầu quản lý và nghĩa vụ công khai: chưa có danh mục tài sản "
        f"trí tuệ trong hệ thống thống kê, hai bảng theo dõi chưa thống nhất trạng thái 5 kiểu dáng, danh mục bài báo "
        f"không liên kết với đề tài, {B.tong_C} trên 16 tiêu chí đánh giá chưa tính được.",
        "Thứ sáu, hoạt động khai thác có thu phí còn ở quy mô nhỏ: hai hợp đồng chuyển giao quyền sử dụng tác phẩm trong "
        "năm năm; các văn bằng, giấy chứng nhận hiện có chưa ghi nhận giao dịch chuyển giao quyền.",
        "Thứ bảy, quy trình nghiệm thu đề tài trước đây chủ yếu tập trung đánh giá mức độ hoàn thành nhiệm vụ chuyên môn "
        "và bài báo công bố, chưa tích hợp tiêu chí sàng lọc và định hướng đăng ký bảo hộ quyền sở hữu trí tuệ.")

    v.doan("h2", "2.5.3. Nguyên nhân của hạn chế")
    v.doan("h3", "a) Nguyên nhân khách quan", bold=True)
    v.than(
        "Thứ nhất, độ trễ thể chế trước sự thay đổi dồn dập của pháp luật quốc gia giai đoạn 2025 - 2026. Luật Khoa học, "
        "công nghệ và đổi mới sáng tạo số 93/2025/QH15 có hiệu lực từ ngày 01 tháng 10 năm 2025, Luật Giáo dục đại học "
        "số 125/2025/QH15 có hiệu lực từ ngày 01 tháng 01 năm 2026, Luật số 131/2025/QH15 sửa đổi Luật Sở hữu trí tuệ có "
        "hiệu lực từ ngày 01 tháng 4 năm 2026; tiếp đó là Kết luận số 51-KL/TW ngày 17 tháng 6 năm 2026 và Quyết định số "
        "1624/QĐ-TTg ngày 21 tháng 8 năm 2026. Các quy chế ban hành năm 2021 và 2024 được xây dựng theo khung pháp luật "
        "tại thời điểm ban hành, nên từ cuối năm 2025 cần được rà soát để xác định phạm vi áp dụng theo luật mới. Nguyên "
        "nhân này giải thích nhu cầu cập nhật quy chế, không giải thích kết quả xác lập quyền giai đoạn 2021 - 2024, khi "
        "các luật mới chưa có hiệu lực.",
        "Thứ hai, thủ tục xác lập quyền sở hữu công nghiệp kéo dài và phát sinh chi phí tra cứu, soạn đơn, lệ phí, phí "
        "duy trì; là trường tư thục, Nhà trường tự cân đối các khoản này từ nguồn thu của mình.",
        "Thứ ba, cơ cấu ngành chủ yếu thuộc kinh tế, quản lý, ngôn ngữ, giáo dục và pháp luật, vốn chủ yếu tạo ra tác "
        "phẩm thuộc quyền tác giả; tiềm năng sở hữu công nghiệp tập trung ở lĩnh vực dược.",
        f"Các nguyên nhân khách quan giải thích vì sao quy mô tài sản trí tuệ còn nhỏ, song chưa giải thích đầy đủ vì sao "
        f"{len(B.dt_shcn_den_2024)} đề tài mã số 2021 - 2024 có sản phẩm tiềm năng sở hữu công nghiệp chưa có đơn.")
    v.doan("h3", "b) Nguyên nhân chủ quan", bold=True)
    v.than(
        "Các nguyên nhân chủ quan dưới đây là nhận định rút ra từ hồ sơ hiện có, cần được kiểm chứng thêm bằng hồ sơ "
        "thực hiện quy chế và ý kiến của chủ nhiệm đề tài.",
        "Về quy trình, đây là khoảng trống kỹ thuật cốt lõi: thiếu một biểu mẫu rà soát tại thời điểm nghiệm thu. Biểu "
        "mẫu hiện hành ghi nhận tình trạng đăng ký nhưng chưa yêu cầu sàng lọc khả năng bảo hộ và tình trạng bộc lộ, nên "
        "việc khởi động thủ tục phụ thuộc nhiều vào sự chủ động của chủ nhiệm đề tài khi nộp hồ sơ tại Phòng Khoa học "
        "Công nghệ theo Điều 35 Quyết định 213; mức độ thực hiện nhiệm vụ nhận diện của Phòng theo Điều 11 Quyết định "
        "217 chưa được đánh giá.",
        "Về thể chế, các văn bản được ban hành ở những thời điểm và cho những kênh tài trợ khác nhau, chưa xác định rõ "
        "phạm vi và thứ tự áp dụng giữa các quy định về lợi ích của tác giả.",
        "Về nguồn lực, quy chế đã có căn cứ chi lệ phí và chi thuê ngoài, nhưng chưa thấy dự toán riêng, cách tạm ứng, "
        "thanh toán khi đơn được nộp sau nghiệm thu và người chịu trách nhiệm đề xuất chi; đơn vị đầu mối chưa có vị trí "
        "chuyên trách về sở hữu trí tuệ.",
        "Về động lực, quy chế chi tiêu áp dụng giai đoạn 2021 - 2025 chưa có trong hồ sơ nên chưa đánh giá được vai trò "
        "của cơ chế khuyến khích đối với kết quả của kỳ. Quy chế chi tiêu nội bộ năm 2026 ghi nhận văn bằng khi được cấp "
        "và chưa có tiền thưởng cho văn bằng, trong khi bài báo quốc tế được thưởng ngay khi đăng; đây là rủi ro cho giai "
        "đoạn tới.",
        "Về tổ chức, việc phân công theo chức năng cho bốn đơn vị là hợp lý nhưng chưa có luồng hồ sơ liên thông giữa các "
        "đơn vị.",
        "Về dữ liệu, chưa có danh mục tài sản trí tuệ liên kết với danh mục đề tài nên đơn vị đầu mối chưa có công cụ phát "
        "hiện sản phẩm cần rà soát.",
        "Về con người, lực lượng nghiên cứu nòng cốt còn mỏng so với quy mô và chưa được bồi dưỡng chuyên sâu về nhận "
        "diện, bảo hộ tài sản trí tuệ.")
    h_cnn = v.so_do("Giả thuyết về chuỗi nguyên nhân dẫn đến việc sản phẩm có tiềm năng chưa được đăng ký bảo hộ",
                    os.path.join(GOC, "Ban_cuoi", "so_do", "chuoi_nguyen_nhan.png"),
                    "Nguồn: Nhóm nghiên cứu tổng hợp từ kết quả phân tích tại Mục 2.2 và Mục 2.3. Các mắt xích là giả "
                    "thuyết cần được kiểm chứng.")
    v.than(
        f"Hình 2.{h_cnn} trình bày các nguyên nhân chủ quan dưới dạng một chuỗi giả thuyết. Dữ liệu hiện có cho thấy sản "
        "phẩm tiềm năng đã được tạo ra, nên điểm nghẽn nhiều khả năng nằm ở khâu nối giữa nghiệm thu và đăng ký: khi chưa "
        "có bước sàng lọc tại nghiệm thu, việc nhận diện phụ thuộc vào sự chủ động của chủ nhiệm đề tài; khi chưa có dự "
        "toán và người chịu trách nhiệm đề xuất chi cho bước đăng ký, thủ tục khó được khởi động; nếu kết quả đã được "
        "công bố quá mười hai tháng, khả năng đăng ký sáng chế, giải pháp hữu ích không còn. Mỗi mắt xích cần được kiểm "
        "chứng bằng hồ sơ thực hiện Điều 10, Điều 11 Quyết định 217 và ý kiến của chủ nhiệm đề tài. Hệ thống giải pháp "
        "tại Chương 3 vì vậy ưu tiên bước sàng lọc tại thời điểm nghiệm thu và trước khi công bố, đồng thời thiết kế "
        "cách đo để kiểm chứng các giả thuyết này.")
    v.ket_qua["h_kh"] = h_kh
    v.doan("h1", "TIỂU KẾT CHƯƠNG 2")
    v.than(
        "Chương 2 đã phân tích thực trạng quản lý quyền sở hữu trí tuệ tại Trường Đại học Thành Đô giai đoạn 2021 - 2025 "
        "trên cơ sở các danh mục thống kê, văn bản nội bộ và kế hoạch của Nhà trường. Kết quả cho thấy Nhà trường có tầm "
        "nhìn thể chế sớm, năng lực công bố tăng nhanh, 11 hồ sơ tài sản trí tuệ trong kỳ, trong đó 4 tài sản đã được "
        "cấp văn bằng, giấy chứng nhận và 5 kiểu dáng cần xác nhận trạng thái, hợp tác doanh nghiệp hiệu quả và trường "
        "hợp đầu tiên chuyển đề tài thành đơn sáng chế. Tiềm năng tài sản trí tuệ, nhất là ở khối ngành Y - Dược, là rõ "
        f"ràng: {len(B.dt_du_dk)} trên 38 đề tài có sản phẩm tiềm năng tạo lập tài sản trí tuệ, trong đó "
        f"{len(B.dt_shcn)} thuộc sở hữu công nghiệp, trong khi mới {len(B.nop_don)} đề tài có đơn.",
        "Về khách quan, pháp luật thay đổi dồn dập trong giai đoạn 2025 - 2026 tạo ra độ trễ thể chế đối với các quy chế "
        "ban hành trước đó, làm phát sinh nhu cầu rà soát phạm vi áp dụng từ cuối năm 2025. Về chủ quan, Đề tài nhận định "
        "khoảng trống kỹ thuật cốt lõi là thiếu một biểu mẫu rà soát tại thời điểm nghiệm thu, cùng với việc chưa có dự "
        "toán riêng cho bước đăng ký, chưa thống nhất cách hiểu và thứ tự áp dụng các quy định về lợi ích của tác giả, và "
        "chưa có dữ liệu liên thông; các nhận định này cần được kiểm chứng trong quá trình triển khai. Đây là căn cứ để "
        "Chương 3 đề xuất hệ thống giải pháp, trong đó bước sàng lọc tại nghiệm thu và trước khi công bố, cùng việc hoàn "
        "thiện quy chế theo phạm vi áp dụng của luật mới, là hai điểm ưu tiên.")


# ---------------------------------------------------------------------------
NHAT_KY_LAN6 = [
    ("Lượt chỉnh sửa tháng 10 năm 2026", "", "Theo bản góp ý và yêu cầu chỉnh sửa ba chương; chi tiết tại bảng giải trình",
     "Bản góp ý của Hội đồng, chủ nhiệm đề tài"),
    ("Hình mô phỏng lợi ích của tác giả", "5 đường: điểm a, điểm b Điều 36 Quyết định 213, Quỹ năm đầu, Quỹ năm thứ hai, "
     "Điều 135", "2 đường cùng cơ sở tính: điểm a và điểm b Điều 36 Quyết định 213; các quy định khác cơ sở tính đưa vào "
     "Bảng 2.3 và sheet DL_Quy_dinh_loi_ich", "Không đặt tỷ lệ tính trên khoản thu sau chi phí cạnh tỷ lệ tính trên tổng "
     "tiền trước thuế"),
    ("Tài sản trí tuệ", "12 tài sản, 9 văn bằng", "11 hồ sơ trong kỳ 2021 - 2025: 4 đã cấp văn bằng, giấy chứng nhận; 5 "
     "kiểu dáng chưa xác minh; 2 đã nộp đơn. Đơn sáng chế 1-2026-07185 ghi riêng, ngoài kỳ",
     "Theo dõi đơn nhãn hiệu và KDCN.xlsx, Sheet1 và Sheet2"),
    ("Kế hoạch 07, mục 1.11", "6 trên 4, tức 150%, gồm nhãn hiệu Thado Edupark", "Chưa xác định; nếu 5 kiểu dáng được "
     "xác nhận cấp năm 2024 thì 5 trên 4, tức 125%. Biểu đồ còn 9 chỉ tiêu, 6 đạt", "Kế hoạch 07/KH-ĐHTĐ trang 7"),
    ("Chuỗi chuyển hóa", "38, 11, 1; giao đến 2024: 31, 8, 0", "Chỉ xét sở hữu công nghiệp: 38, 9, 1; mã số 2021 - "
     "2024: 31, 6, 0", "Không gộp quyền tác giả vào phễu sở hữu công nghiệp"),
    ("Bản ghi trùng", "107 bài gắn đề tài, 475 sản phẩm độc lập, độ phủ 29,9%", "Bỏ; giữ 582 bản ghi",
     "Chưa tái lập được phép so khớp và chưa xác lập được quan hệ trùng lặp"),
    ("Người có bài báo", "88 trên 247, Gini 0,829, 10% dẫn đầu 39,3%", "87 trên 246 tên, Gini 0,832, 39,7%; độ nhạy "
     "bỏ dấu: 103 tên, Gini 0,822", "scripts/ra_soat/so_khop_tac_gia.py"),
    ("Khả năng tính toán tiêu chí", "5 đầy đủ, 4 một phần, 7 chưa", "4 đầy đủ, 5 một phần, 7 chưa",
     "Kinh phí: Điều 35, 38 Quyết định 213; văn bằng: trạng thái kiểu dáng chưa thống nhất; tra cứu: chưa có số liệu"),
]

NHAT_KY_GON = [
    ("Bản rút gọn, cấu trúc", "43 trang, 16 hình, 10 bảng",
     "Khoảng 22 trang, 9 hình, 7 bảng, có tiểu kết; bỏ Mục 2.1.4 và 2.1.5 cũ, gộp Mục 2.3.4 vào 2.3.2, bỏ phân mục của 2.4; "
     "đánh số hình, bảng theo bản rút gọn", "Yêu cầu của chủ nhiệm đề tài: chỉ giữ nội dung trực tiếp liên quan"),
    ("Hình bỏ khỏi bản rút gọn", "Hình 2.1, 2.3, 2.4, 2.5, 2.6, 2.9, 2.13 cũ",
     "Số liệu chính giữ trong lời văn; dữ liệu vẫn có trong các sheet DL_ và tệp du_lieu.py", "Rút gọn"),
    ("Quy chế quản trị tài sản trí tuệ năm 2024", "Chưa rõ số hiệu",
     "Ban hành kèm Quyết định số 217/QĐ-ĐHTĐ, gọi tắt là Quyết định 217", "Xác nhận của chủ nhiệm đề tài"),
    ("Thông tư 83/2026/TT-BGDĐT về Chuẩn cơ sở giáo dục đại học", "Chưa có",
     "Không đưa vào Chương 2 vì Thông tư có hiệu lực từ ngày 15 tháng 11 năm 2026, chưa áp dụng cho giai đoạn đánh giá; "
     "được đưa vào Mục 3.1.2 Thời cơ và thách thức của Chương 3",
     "Ý kiến của chủ nhiệm đề tài"),
    ("Ngân sách Quỹ Ngô Xuân Độ", "5 tỷ đồng", "5 tỷ đồng giai đoạn 2025 - 2029", "Xác nhận của chủ nhiệm đề tài"),
    ("Biểu đồ", "Chú giải đặt dưới, bố cục tự động; nhãn trắng trên mọi nền",
     "Bố cục cố định: chú giải phía trên, vùng vẽ tách riêng; màu chữ nhãn chọn theo độ tương phản với nền",
     "Phản hồi của chủ nhiệm đề tài về chú giải đè biểu đồ và chữ khó đọc"),
]


def kiem_tra(doc):
    loi = []
    for p in doc.paragraphs:
        t = p.text
        if "—" in t or "–" in t:
            loi.append(("gạch dài", t[:80]))
        if re.search(r"\b(tôi|chúng tôi)\b", t, flags=re.I):
            loi.append(("ngôi thứ nhất", t[:80]))
        if "(" in t and not t.startswith("Nguồn"):
            loi.append(("ngoặc đơn", t[:80]))
    return loi


def main():
    v = VanBan()
    noi_dung(v)
    BD.dat_cap_de_muc(v.doc)
    for p in v.doc.paragraphs:
        if p.text.startswith("TIỂU KẾT"):
            ppr = p._p.get_or_add_pPr()
            ol = etree.Element(qn("w:outlineLvl"))
            ol.set(qn("w:val"), "1")
            rpr = ppr.find(qn("w:rPr"))
            rpr.addprevious(ol) if rpr is not None else ppr.append(ol)
    for p in v.doc.paragraphs:
        if re.match(r"^[ab]\) Nguyên nhân", p.text):
            p.paragraph_format.keep_with_next = True
    # workbook chỉ gồm các hình có trong bản rút gọn, đánh số mới
    B.HINH[:] = v.hinh
    BD.RA_XLSX = os.path.join(GOC, "Ban_cuoi", "Du_lieu_bieu_do_Chuong_2.xlsx")
    BD.NHAT_KY[:0] = NHAT_KY_GON
    BD.dung_workbook()
    BD.va_workbook(BD.RA_XLSX)
    v.doc.core_properties.title = "Chương 2. Thực trạng quản lý quyền sở hữu trí tuệ tại Trường Đại học Thành Đô"
    os.makedirs(os.path.dirname(RA_DOCX), exist_ok=True)
    v.doc.save(RA_DOCX)
    so_tu = sum(len(p.text.split()) for p in v.doc.paragraphs)
    print("Đã ghi:", RA_DOCX, f"({so_tu} từ ngoài bảng, {v.so_bang} bảng, {v.so_hinh} hình)")
    print("Đã ghi:", BD.RA_XLSX)
    for x in kiem_tra(v.doc):
        print("CẢNH BÁO:", x)


if __name__ == "__main__":
    main()

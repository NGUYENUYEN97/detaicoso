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
assert kh_dat == 7 and len(B.kh_rows) == 10
kh = B.kh
tong_nam = B.tong_nam
assert tong_nam == [56, 100, 85, 128, 192] and B.tong_ban_ghi == 582 and B.doc_lap == 475
assert B.bb[0] == 23 and B.bb[4] == 161
q_tong = sum(B.co_hang)
assert q_tong == 71 and sum(B.co_hang[3:]) == 68
assert (B.tong_T, B.tong_M, B.tong_C) == (5, 4, 7)
assert (len(B.DE_TAI), len(B.dt_du_dk), len(B.nop_don), len(B.du_dk_den_2024)) == (38, 11, 1, 8)
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
        self.hinh = []

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
    def hinh_bd(self, ma, nguon=None, tieu_de=None):
        h = next(x for x in B.HINH if x["id_cu"] == ma)
        self.so_hinh += 1
        h["so"] = self.so_hinh
        h["id"] = f"H2.{self.so_hinh}"
        if nguon:
            h["nguon"] = nguon
        if tieu_de:
            h["tieu_de"] = tieu_de
        self.hinh.append(h)
        self.doan("tieu_de", f"Hình 2.{h['so']}. {h['tieu_de']}", bold=True, giu=True)
        p = self.doan("tieu_de", "", giu=True)
        for r in p._p.findall(qn("w:r")):
            p._p.remove(r)
        pf = p.paragraph_format
        pf.space_before = 0
        pf.space_after = 0
        pf.line_spacing = 1.0
        p._p.append(BD.nhung_bieu_do(self.doc, h, self.so_hinh))
        self.doan("nguon", h["nguon"], italic=True)
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
        "văn bản pháp luật được đối chiếu theo bản hiện hành đến tháng 9 năm 2026.")

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
        f"Bảng 2.1 cho thấy năng lực chuyên môn và thẩm quyền xử lý hồ sơ sở hữu trí tuệ nằm ở hai khối khác nhau. Ba "
        f"viện đào tạo chiếm {pt(B.nl_ba_vien / tong_nl)} nhân sự nhưng có {B.ts_ba_vien} trên {ts_tong} người trình "
        f"độ tiến sĩ, tức {pt(B.ts_ba_vien / ts_tong)}, và {B.hoc_ham_ba_vien} trên {hoc_ham} người có học hàm. Ngược "
        f"lại, hai trung tâm thuộc khối Quản trị và Dịch vụ, nơi đang thực hiện việc xác lập quyền đối với nhãn hiệu và "
        f"giữ đầu mối pháp chế, có {B.nl_khoi_qt} nhân sự, không có tiến sĩ và "
        f"{pt(B.dh_khac_khoi_qt / B.nl_khoi_qt)} có trình độ đại học trở xuống. Phòng Khoa học Công nghệ, đơn vị đầu "
        "mối theo quy chế, có 2 nhân sự trình độ thạc sĩ, chức danh trưởng phòng do Hiệu trưởng kiêm nhiệm và không có "
        "người được đào tạo về sở hữu trí tuệ. Người có khả năng nhận diện kết quả có thể bảo hộ ở khối Đào tạo và "
        "Nghiên cứu, còn người xử lý thủ tục ở khối có năng lực chuyên môn khoa học mỏng hơn; khoảng cách này là tiền "
        "đề của các điểm nghẽn về tổ chức phân tích tại Mục 2.2.2.")

    # ---------------------------------------------------------------- 2.1.2
    v.doan("h2", "2.1.2. Sản phẩm khoa học giai đoạn 2021 - 2025")
    v.than("Các danh mục thống kê của Phòng Khoa học Công nghệ ghi nhận 582 bản ghi sản phẩm khoa học trong giai đoạn "
           "2021 - 2025, được trình bày tại Bảng 2.2.")
    dong = []
    for ten, vals in D.SAN_PHAM:
        dong.append([ten] + [x if x else "-" for x in vals] + [sum(vals)])
    dong.insert(6, ["Tham luận hội thảo quốc gia", "-", "-", "-", "-", "-", D.THAM_LUAN_QUOC_GIA])
    dong.append(["Tổng cộng"] + tong_nam + [B.tong_ban_ghi])
    v.bang("Sản phẩm khoa học của Trường Đại học Thành Đô, giai đoạn 2021 - 2025",
           ["Loại sản phẩm", "2021", "2022", "2023", "2024", "2025", "Tổng"], dong,
           "Nguồn: Nhóm nghiên cứu tổng hợp từ các danh mục thống kê của Phòng Khoa học Công nghệ. Danh mục tham luận "
           "hội thảo quốc gia không có trường thời gian nên chỉ có số tổng. Đề tài cấp cơ sở xếp theo năm ghi trong mã "
           "số; đề tài cấp quốc gia xếp theo năm phê duyệt kinh phí. Số liệu bài báo, giáo trình và đề tài cấp cơ sở "
           "giai đoạn 2021 - 2023 khớp với số liệu tổng kết tại Kế hoạch số 07/KH-ĐHTĐ ngày 01 tháng 7 năm 2024.",
           [5.4, 1.2, 1.2, 1.2, 1.2, 1.2, 1.3], dong_tong=True)
    v.than(
        f"Con số 582 là số bản ghi, chưa phải số sản phẩm độc lập. So khớp tên chủ nhiệm đề tài với danh sách tác giả "
        f"trong ba năm kể từ năm nghiệm thu cho thấy {D.BAI_BAO_GAN_DE_TAI} trên 405 bài báo, tức "
        f"{pt(D.BAI_BAO_GAN_DE_TAI / 405)}, gắn với ít nhất một đề tài và đã được tính hai lần; quy mô thực tế vào "
        f"khoảng {B.doc_lap} sản phẩm độc lập.")
    h_sp = v.hinh_bd("H2.2")
    v.than(
        f"Hình 2.{h_sp} cho thấy sản lượng tăng từ {tong_nam[0]} sản phẩm năm 2021 lên {tong_nam[4]} sản phẩm năm "
        f"2025, bình quân {pt(B.cagr(tong_nam[0], tong_nam[4], 4))} một năm, và mức tăng này gần như hoàn toàn đến từ "
        f"bài báo. Bài báo tăng từ {B.bb[0]} lên {B.bb[4]} bài, gấp {so(B.bb[4] / B.bb[0], 1)} lần, nâng tỷ trọng từ "
        f"{pt(B.bb[0] / tong_nam[0])} lên {pt(B.bb[4] / tong_nam[4])} sản lượng; giáo trình giảm từ {B.gt[0]} xuống "
        f"{B.gt[4]} tài liệu, tỷ trọng từ {pt(B.gt[0] / tong_nam[0])} còn {pt(B.gt[4] / tong_nam[4])}. Chất lượng "
        f"công bố quốc tế cũng tăng: {q_tong} trên 115 bài quốc tế đăng trên tạp chí có phân hạng Q, trong đó "
        f"{sum(B.co_hang[3:])} bài thuộc hai năm 2024 - 2025.",
        f"Đối với quản lý quyền sở hữu trí tuệ, cơ cấu này có hai hàm ý. Thứ nhất, bài báo chỉ phát sinh quyền tác giả "
        "và hầu như không có khả năng khai thác thương mại, nên tăng công bố không tự động tạo ra tài sản trí tuệ có "
        "giá trị khai thác. Thứ hai, cùng một đội ngũ trong cùng giai đoạn đã tạo được bước nhảy về công bố nhưng không "
        "tạo được bước nhảy tương ứng về đơn đăng ký sở hữu công nghiệp; khoảng trống vì vậy không nằm chủ yếu ở năng "
        "lực nghiên cứu mà ở cách thiết kế động lực và quy trình, được phân tích tại Mục 2.2.3 và Mục 2.3.3.",
        "Trong hai năm cuối kỳ, Nhà trường chủ trì ba đề tài cấp quốc gia do Quỹ Phát triển khoa học và công nghệ quốc "
        "gia tài trợ, tổng kinh phí 4,67 tỷ đồng, đều đang thực hiện. Theo điểm c khoản 1 Điều 86 Luật Sở hữu trí tuệ "
        "được bổ sung bởi Luật số 131/2025/QH15, tổ chức được giao quyền quản lý, sử dụng, quyền sở hữu kết quả của "
        "nhiệm vụ sử dụng ngân sách nhà nước có quyền đăng ký sáng chế, kiểu dáng công nghiệp, thiết kế bố trí là kết "
        "quả của nhiệm vụ đó. Đây là nguồn tài sản trí tuệ tiềm năng lớn nhất của giai đoạn tới, với điều kiện quy "
        "trình rà soát được chuẩn bị trước thời điểm nghiệm thu.",
        f"Năng lực công bố phân bố không đều. So khớp họ tên tác giả với danh sách nhân sự cho thấy chỉ "
        f"{D.NHAN_SU_CO_BAI} trên {D.NHAN_SU_CO_TEN} nhân sự có họ tên đầy đủ, tức "
        f"{pt(D.NHAN_SU_CO_BAI / D.NHAN_SU_CO_TEN)}, đứng tên ít nhất một bài báo trong năm năm; hệ số Gini về số bài "
        "trên toàn bộ nhân sự là 0,829, và trong nhóm có công bố, 10% người dẫn đầu chiếm 39,3% số lượt đứng tên. "
        "Viện Y - Dược có số bài bình quân thấp nhất trong ba viện, 0,79 bài một nhân sự, nhưng lại là nơi phát sinh "
        "phần lớn sản phẩm đề tài đủ điều kiện xác lập quyền và cả hai đơn sáng chế. Số lượng công bố và khả năng hình "
        "thành tài sản trí tuệ vì vậy là hai đại lượng khác nhau, cần được theo dõi bằng hai thước đo riêng.")

    # ===================================================================== 2.2
    v.doan("h1", "2.2. Thực trạng thể chế, tổ chức và nguồn lực quản lý quyền sở hữu trí tuệ")
    v.doan("h2", "2.2.1. Hệ thống quy định nội bộ")
    v.than(
        "Nhà trường có bốn văn bản chứa quy định về quyền sở hữu trí tuệ. Quy chế hoạt động khoa học công nghệ ban hành "
        "kèm Quyết định số 213/QĐ-ĐHTĐ ngày 28 tháng 12 năm 2021, sau đây gọi là Quyết định 213, dành Chương VI cho sở "
        "hữu trí tuệ và chuyển giao công nghệ: Điều 34 liệt kê phạm vi tài sản khá đầy đủ, từ tên trường, nhãn hiệu, "
        "sáng chế, giải pháp hữu ích, kiểu dáng công nghiệp đến giáo trình, ngân hàng đề thi, phần mềm và quy trình công "
        "nghệ; Điều 35 quy định quy trình đăng ký một cửa tại Phòng Khoa học Công nghệ; Điều 36 quy định hai công thức "
        "chia lợi ích. Năm 2024, Nhà trường ban hành Quy chế quản trị tài sản trí tuệ kèm Quyết định số 217/QĐ-ĐHTĐ, "
        "sau đây gọi là Quyết định 217, mở rộng phạm vi tới cơ sở dữ liệu, giáo trình điện tử, bí quyết và tên miền. "
        "Điều 10 Quyết định 217 yêu cầu tác giả xin ý kiến Phòng Khoa học Công nghệ trước khi bộc lộ công khai tài sản "
        "có thể bảo hộ; Điều 11 giao Phòng nhận diện, lập hồ sơ theo dõi, xúc tiến thương mại hóa và giao bộ phận pháp "
        "chế thực hiện thủ tục xác lập quyền. Hai văn bản còn lại là Quy chế chi tiêu nội bộ ban hành ngày 01 tháng 8 "
        "năm 2026 và Điều lệ Quỹ Học bổng sau tiến sĩ Ngô Xuân Độ năm 2025.",
        "Quyết định 217 không dẫn chiếu và không thay thế Chương VI Quyết định 213, nên hai văn bản cùng hiệu lực. Bốn "
        f"văn bản chứa năm quy định khác nhau về phân chia lợi ích, được trình bày tại Bảng 2.{v.so_bang + 1}.")
    b_ll = v.bang(
        "Các quy định về phân chia lợi ích từ tài sản trí tuệ trong nội bộ Trường Đại học Thành Đô",
        ["Văn bản", "Năm", "Phạm vi áp dụng", "Công thức chia lợi ích", "Mức trần"],
        [["Quyết định 213, điểm a khoản 4 Điều 36", "2021", "Đề tài sử dụng ngân sách nhà nước do Trường chủ trì",
          "40% nộp ngân sách nhà nước, 30% Trường, 30% khen thưởng tập thể tác giả", "100 triệu đồng"],
         ["Quyết định 213, điểm b khoản 4 Điều 36", "2021", "Tài sản trí tuệ thuộc sở hữu của Trường được chuyển giao",
          "30% tác giả, 20% đơn vị có tác giả, 50% Quỹ nghiên cứu khoa học của Trường", "Không đặt trần"],
         ["Quyết định 217, Điều 13", "2024", "Tài sản trí tuệ của Trường khi các bên không có thỏa thuận",
          "Sau khi trừ chi phí, Hiệu trưởng quyết định tỷ lệ", "Không quy định"],
         ["Quy chế chi tiêu nội bộ", "2026", "Đề tài có đăng ký sở hữu trí tuệ",
          "Trích 50% kinh phí chuyển giao công nghệ về Trường, không nêu phần của tác giả",
          "50 triệu đồng kinh phí đề tài"],
         ["Điều lệ Quỹ Ngô Xuân Độ, Điều 9", "2025", "Sản phẩm hình thành từ kinh phí của Quỹ",
          "Năm đầu: 50% tác giả, 50% Trường; từ năm thứ hai: 20% tác giả, 80% Trường", "Không đặt trần"]],
        "Nguồn: Nhóm nghiên cứu tổng hợp từ Quyết định 213, Quyết định 217, Quy chế chi tiêu nội bộ năm 2026 và Điều "
        "lệ Quỹ Học bổng sau tiến sĩ Ngô Xuân Độ.",
        [3.2, 1.4, 3.4, 4.6, 2.0], can=["left", "center", "left", "left", "left"])
    h_mp = v.hinh_bd("H2.8", nguon=(
        "Nguồn: Nhóm nghiên cứu mô phỏng từ Điều 36 Quyết định 213, Điều 9 Điều lệ Quỹ Học bổng sau tiến sĩ Ngô Xuân "
        "Độ và điểm b khoản 1 Điều 135 Luật Sở hữu trí tuệ. Khoản thu là số tiền nhận được từ một hợp đồng chuyển giao "
        "sau khi trừ chi phí hợp lệ; đơn vị: triệu đồng."))
    v.than(
        f"Hình 2.{h_mp} mô phỏng phần của tác giả từ một hợp đồng chuyển giao theo từng quy định. Với khoản thu 500 "
        "triệu đồng, phần của tác giả dao động từ 100 triệu đồng theo điểm a Điều 36 Quyết định 213 hoặc theo Quỹ từ "
        "năm thứ hai, 150 triệu đồng theo điểm b, đến 250 triệu đồng theo Quỹ trong năm đầu, chênh lệch 2,5 lần. Đường "
        "điểm a gãy tại khoảng 333 triệu đồng do trần 100 triệu đồng; vượt khoảng 667 triệu đồng, đường này nằm dưới "
        "mức mặc định 15% của Luật. Điều 13 Quyết định 217 giao Hiệu trưởng quyết định tỷ lệ nên không mô phỏng được.",
        "Về tỷ lệ, quy chế nội bộ không kém hào phóng so với mức mặc định của Luật. Vấn đề là tác giả không biết trước "
        "mình thuộc đường nào, trong khi mức thưởng cho công bố được niêm yết theo từng hạng tạp chí. Mức trần 100 triệu "
        "đồng được xây dựng theo khung pháp luật trước đây; khoản 2 Điều 135 Luật Sở hữu trí tuệ về khung thù lao cho "
        "nhiệm vụ sử dụng ngân sách nhà nước đã bị bãi bỏ, Điều 135 hiện hành chỉ quy định mức áp dụng khi không có "
        "thỏa thuận, gồm 10% lợi nhuận trước thuế khi chủ sở hữu tự sử dụng và 15% số tiền nhận được mỗi lần chuyển "
        "giao quyền sử dụng, và không đặt trần. Trong bốn văn bản, Điều lệ Quỹ là văn bản duy nhất vừa không đặt trần "
        "vừa cho phép các bên thỏa thuận tỷ lệ khác.")

    v.doan("h2", "2.2.2. Tổ chức bộ máy và đầu mối quản lý")
    v.than(
        "Theo Điều 35 Quyết định 213 và Điều 11 Quyết định 217, Phòng Khoa học Công nghệ là đầu mối tiếp nhận hồ sơ, "
        "nhận diện và theo dõi tài sản trí tuệ; bộ phận pháp chế thực hiện thủ tục xác lập quyền. Trên thực tế, chức "
        "năng này phân tán ở bốn điểm. Phòng Khoa học Công nghệ giữ khâu tiếp nhận hồ sơ nhưng chỉ có 2 nhân sự. Trung "
        "tâm Tuyển sinh và Quản trị thương hiệu thực hiện xác lập quyền đối với nhãn hiệu và logo. Bộ phận pháp chế "
        "thuộc Trung tâm Dịch vụ và Quản trị hành chính tổng hợp là đầu mối của Mạng lưới Trung tâm Hỗ trợ công nghệ và "
        "đổi mới sáng tạo mà Nhà trường tham gia từ năm 2023. Viện Nghiên cứu giáo dục và Chuyển giao tri thức là đơn "
        "vị duy nhất đã đăng ký hoạt động khoa học và công nghệ, có con dấu riêng và có hợp đồng khai thác quyền.",
        "Ba trong bốn điểm thuộc khối Quản trị và Dịch vụ, còn kết quả nghiên cứu phát sinh ở khối Đào tạo và Nghiên "
        "cứu. Vấn đề vì vậy không phải thiếu đầu mối hay thiếu phân công, vì Quyết định 217 đã phân công rõ, mà nằm ở "
        "khâu thực thi: bộ hồ sơ thu thập được không có biểu mẫu khai báo hay hồ sơ theo dõi tài sản trí tuệ do Phòng "
        "lập; mỗi đơn vị nắm một đoạn của chu trình và giữa các đoạn không có cơ chế chuyển hồ sơ.")

    v.doan("h2", "2.2.3. Nguồn lực tài chính và cơ chế khuyến khích")
    v.than(f"Nhà trường có ba kênh tài trợ nghiên cứu với chế độ sở hữu trí tuệ khác nhau, được trình bày tại Bảng 2.{v.so_bang + 1}.")
    b_kenh = v.bang(
        "Ba kênh tài trợ nghiên cứu và chế độ sở hữu trí tuệ tương ứng",
        ["Kênh tài trợ", "Quy mô", "Dòng chi cho phí nộp đơn", "Kết quả tài sản trí tuệ"],
        [["Đề tài cấp cơ sở, ngân sách của Trường",
          f"38 đề tài giai đoạn 2021 - 2025; 19 đề tài được cấp tổng {so(B.kp_tong, 2)} triệu đồng, 16 đề tài chỉ "
          "quy đổi giờ, 3 đề tài tự tìm tài trợ", "Không có", "11 đề tài có sản phẩm đủ điều kiện, 1 đơn sáng chế"],
         ["Quỹ Học bổng sau tiến sĩ Ngô Xuân Độ",
          "Ngân sách 5 tỷ đồng giai đoạn 2025 - 2029; 10 đến 15 học bổng mỗi năm, mức 50 đến 500 triệu đồng",
          "Có, qua khoản chi phí khác tối đa bằng 50% chi trực tiếp", "Quỹ hoạt động từ năm 2025, chưa có kết quả"],
         ["Đề tài cấp quốc gia, Quỹ Phát triển khoa học và công nghệ quốc gia",
          "3 đề tài, tổng 4,67 tỷ đồng, giao năm 2024 và 2025", "Theo quy định của cơ quan tài trợ",
          "Đang thực hiện, chưa nghiệm thu"]],
        "Nguồn: Nhóm nghiên cứu tổng hợp từ danh mục đề tài cấp cơ sở, danh mục đề tài cấp quốc gia và Điều lệ Quỹ "
        "Học bổng sau tiến sĩ Ngô Xuân Độ.",
        [3.6, 5.0, 3.2, 3.4], can=["left", "left", "left", "left"])
    v.than(
        f"Bảng 2.{b_kenh} cho thấy nguồn lực không thiếu nhưng đặt chưa đúng chỗ. Quỹ Học bổng sau tiến sĩ Ngô Xuân Độ "
        "có ngân sách 5 tỷ đồng, có khoản chi phí khác đủ trang trải chi phí đăng ký và quy định rõ về đồng sở hữu văn "
        "bằng, nhưng chỉ dành cho nhà khoa học theo diện học bổng sau tiến sĩ. Đề tài cấp cơ sở, kênh duy nhất đến nay "
        f"tạo ra sản phẩm đủ điều kiện xác lập quyền, chỉ được Nhà trường cấp {so(B.kp_tong, 2)} triệu đồng trong năm "
        f"năm, bằng {pt(B.kp_tong / 5000)} ngân sách Quỹ, và không có dòng chi cho phí nộp đơn, phí duy trì hiệu lực. "
        "Điều 38 Quyết định 213 không có mục chi riêng cho lệ phí đăng ký; Điều 13 Quyết định 217 chỉ coi lệ phí xác "
        "lập quyền là khoản được trừ khi chia lợi ích, tức sau khi đã có doanh thu. Giữa hai kênh không có cơ chế "
        "chuyển tiếp.")
    h_kp = v.hinh_bd("H2.10")
    v.than(
        f"Hình 2.{h_kp} cho thấy cơ chế cấp kinh phí đề tài cấp cơ sở thay đổi sau năm 2022: toàn bộ {B.dt_gio[0]} đề "
        f"tài năm 2021 và {B.dt_gio[1]} trên {len(B.dt_nam['2022'])} đề tài năm 2022 chỉ được quy đổi giờ nghiên cứu; "
        f"từ năm 2023, phần lớn đề tài được cấp tiền và tổng kinh phí đạt {so(B.kp_nam[4], 0)} triệu đồng năm 2025. Mức "
        f"cấp phân tán mạnh, trung vị {so(B.kp_trung_vi, 1)} triệu đồng, cao nhất 80 triệu đồng. Đáng chú ý, cả 11 đề "
        f"tài có sản phẩm đủ điều kiện xác lập quyền đều thuộc nhóm {len(B.dt_co_tien)} đề tài được cấp tiền và dùng "
        f"{so(B.kp_du_dk, 2)} triệu đồng, tức {pt(B.kp_du_dk / B.kp_tong)} tổng kinh phí. Kinh phí bằng tiền là điều "
        "kiện cần để tạo ra sản phẩm có thể bảo hộ, nhưng chỉ đủ tạo ra sản phẩm, không có phần cho bước đăng ký.",
        "Về khuyến khích, Quy chế chi tiêu nội bộ năm 2026 quy định kinh phí xét duyệt không quá 50 triệu đồng cho đề "
        "tài có đăng ký sở hữu trí tuệ và quy đổi giờ nghiên cứu cho văn bằng, với điều kiện đã có chứng nhận đăng ký "
        "thành công. Định mức này được đặt cạnh chế độ dành cho công bố tại hình dưới đây.")
    h_kk = v.hinh_bd("H2.11")
    v.than(
        f"Hình 2.{h_kk} cho thấy, xét riêng giờ quy đổi, văn bằng được định giá cao: sáng chế chuẩn Việt Nam được tính "
        "360 giờ, cao hơn bài WoS hạng Q1 với 300 giờ. Nhưng chỉ bài báo quốc tế được thưởng tiền, từ 10 đến 20 triệu "
        "đồng một bài, và được ghi nhận ngay khi đăng; văn bằng không có tiền thưởng và chỉ được ghi nhận khi được cấp, "
        "thường từ hai đến ba năm sau khi nộp đơn. Với cùng một kết quả nghiên cứu, lựa chọn hợp lý của giảng viên là "
        "công bố trước. Theo khoản 3 và khoản 4 Điều 60 Luật Sở hữu trí tuệ, đơn phải được nộp trong thời hạn mười hai "
        "tháng kể từ ngày bộc lộ thì sáng chế mới không bị coi là mất tính mới. Cấu trúc khuyến khích này phù hợp với "
        "việc công bố tăng mạnh trong khi đơn đăng ký hầu như vắng mặt.")

    v.doan("h2", "2.2.4. Hệ thống dữ liệu phục vụ quản lý")
    v.than(
        "Hệ thống thống kê của Phòng Khoa học Công nghệ chỉ theo dõi bài báo, sách, giáo trình, đề tài và tham luận; "
        "không có danh mục nào theo dõi tài sản trí tuệ đã đăng ký hoặc được cấp văn bằng, dù Điều 11 Quyết định 217 "
        "giao Phòng nhiệm vụ này. Danh mục tài sản trí tuệ hiện có do bộ phận quản trị thương hiệu lập và không liên "
        "kết với các danh mục trên. Mức độ khai báo trong hồ sơ nghiệm thu thấp: 26 trên 38 đề tài khai tổng cộng 32 "
        f"công bố, trong khi phép so khớp gắn được {D.BAI_BAO_GAN_DE_TAI} bài báo với các đề tài, tức độ phủ khoảng "
        "29,9%. Nếu bài báo, loại sản phẩm dễ nhận diện nhất, còn khai thiếu thì công thức, quy trình càng ít khả năng "
        "được ghi nhận. Dữ liệu cũng chưa chuẩn hóa: tên đơn vị ghi nhiều cách, danh mục bài báo không có trường đơn vị, "
        "biểu mẫu giáo trình năm 2021 khác các năm sau.")
    h_tc = v.hinh_bd("H2.7")
    v.than(
        f"Hình 2.{h_tc} lượng hóa hệ quả: trong 16 tiêu chí đánh giá hiệu quả quản lý quyền sở hữu trí tuệ xây dựng "
        f"tại Mục 1.4.2, chỉ {B.tong_T} tiêu chí tính được đầy đủ, {B.tong_M} tiêu chí tính được một phần và "
        f"{B.tong_C} tiêu chí chưa tính được; nhóm kết quả không có tiêu chí nào tính được đầy đủ. Nhà trường đo được "
        "đã xác lập bao nhiêu quyền nhưng chưa đo được quyền mang lại giá trị gì. Đây lại là nhóm thông tin phải công "
        "khai theo điểm đ khoản 3 Điều 28 Luật Giáo dục đại học số 125/2025/QH15.")

    # ===================================================================== 2.3
    v.doan("h1", "2.3. Thực trạng tạo lập, xác lập và bảo vệ quyền sở hữu trí tuệ")
    v.doan("h2", "2.3.1. Đối sánh giữa quy định pháp luật, quy chế nội bộ và thực tế phát sinh")
    v.than(f"Bảng 2.{v.so_bang + 1} đối sánh ba lớp: đối tượng quyền theo Luật Sở hữu trí tuệ, tài sản được liệt kê tại Điều 34 Quyết "
           "định 213 và Điều 3 Quyết định 217, và thực tế phát sinh giai đoạn 2021 - 2025.")
    b_ds = v.bang(
        "Đối sánh đối tượng quyền theo Luật, theo quy chế nội bộ và theo thực tế phát sinh",
        ["Nhóm quyền", "Đối tượng", "Quy chế có liệt kê", "Thực tế phát sinh", "Đã xác lập quyền"],
        [["Quyền tác giả", "Bài báo, sách", "Có", "405 bài báo, 12 sách", "Chưa"],
         ["Quyền tác giả", "Giáo trình, tài liệu giảng dạy", "Có", "87 tài liệu", "Chưa"],
         ["Quyền tác giả", "Chương trình đào tạo, ngân hàng đề thi, phần mềm", "Có", "Chưa thống kê", "Chưa"],
         ["Quyền tác giả", "Sưu tập dữ liệu", "Có, tại Quyết định 217", "Bộ mẫu cây thuốc, bộ tiêu bản", "Chưa"],
         ["Quyền tác giả", "Logo, bộ nhận diện", "Có", "Có", "2 văn bằng"],
         ["Sở hữu công nghiệp", "Sáng chế", "Có", "Có", "2 đơn đang xử lý"],
         ["Sở hữu công nghiệp", "Giải pháp hữu ích", "Có", "Công thức, quy trình từ đề tài", "Chưa"],
         ["Sở hữu công nghiệp", "Kiểu dáng công nghiệp", "Có", "Có", "5 văn bằng"],
         ["Sở hữu công nghiệp", "Nhãn hiệu", "Có", "Có", "2 văn bằng, 1 đơn"],
         ["Sở hữu công nghiệp", "Bí mật kinh doanh", "Có, tại Quyết định 217", "Công thức chưa công bố",
          "Không áp dụng đăng ký"]],
        "Nguồn: Nhóm nghiên cứu tổng hợp từ Luật Sở hữu trí tuệ, Quyết định 213, Quyết định 217 và các danh mục của "
        "Trường Đại học Thành Đô. Bảng không liệt kê quyền liên quan, thiết kế bố trí, chỉ dẫn địa lý và giống cây trồng "
        "do chưa phát sinh hoặc không thuộc phạm vi hoạt động của Nhà trường.",
        [2.6, 3.8, 2.6, 3.8, 2.6], can=["left", "left", "left", "left", "left"])
    v.than(
        f"Bảng 2.{b_ds} cho thấy Nhà trường phát sinh tài sản ở nhiều nhóm đối tượng nhưng mới xác lập quyền ở bốn "
        "nhóm, trong đó ba nhóm đã có văn bằng. Giải pháp hữu ích và sưu tập dữ liệu, hai nhóm mà sản phẩm đề tài cấp "
        "cơ sở rơi vào nhiều nhất, đều đã có trong quy chế nhưng chưa sản phẩm nào được xác lập quyền; giáo trình, nhóm "
        "có khả năng khai thác trong đào tạo, chưa được đăng ký quyền tác giả. Khoảng trống vì vậy nằm ở khâu nhận "
        "diện, không nằm ở phạm vi văn bản.")

    v.doan("h2", "2.3.2. Tài sản trí tuệ đã được xác lập quyền")
    v.than("Đến thời điểm nghiên cứu, Nhà trường sở hữu 12 tài sản trí tuệ đã được xác lập quyền hoặc đang xử lý đơn, "
           f"được liệt kê tại Bảng 2.{v.so_bang + 1}.")
    dong = [[i + 1, t[0], t[1], t[2], t[3], t[4], t[5]] for i, t in enumerate(D.TSTT)]
    b_ts = v.bang(
        "Tài sản trí tuệ thuộc sở hữu của Trường Đại học Thành Đô",
        ["TT", "Loại hình", "Tên tài sản", "Năm", "Trạng thái", "Cơ cấu sở hữu", "Nguồn hình thành"], dong,
        "Nguồn: Nhóm nghiên cứu tổng hợp từ danh mục theo dõi đơn nhãn hiệu, kiểu dáng công nghiệp và thống kê văn "
        "bằng của Trường Đại học Thành Đô. Năm là năm cấp văn bằng hoặc năm nộp đơn đối với hồ sơ đang xử lý; danh mục "
        "chưa ghi số đơn của hai sáng chế.",
        [1.0, 2.4, 4.2, 1.4, 2.1, 2.5, 2.4], can=["center", "left", "left", "center", "left", "left", "left"])
    h_ts = v.hinh_bd("H2.12", nguon=f"Nguồn: Nhóm nghiên cứu tổng hợp từ Bảng 2.{b_ts}.")
    v.than(
        f"Bảng 2.{b_ts} và Hình 2.{h_ts} cho thấy danh mục hình thành từ ba luồng với kết quả rất khác nhau. Luồng "
        "thương hiệu gồm 4 tài sản, đều do Nhà trường đơn sở hữu và đều đã có văn bằng. Luồng hợp tác doanh nghiệp gồm "
        "5 kiểu dáng công nghiệp và 1 nhãn hiệu, đều đồng sở hữu với cùng một doanh nghiệp; riêng năm 2024 có 5 kiểu "
        "dáng được cấp, nâng số lũy kế từ 4 lên 9 tài sản. Luồng nghiên cứu chỉ có 2 đơn sáng chế nộp năm 2025 và 2026, "
        f"chưa có văn bằng. Như vậy, khoảng {B.doc_lap} sản phẩm khoa học chỉ đóng góp 2 trên 12 tài sản; năng lực xác "
        "lập quyền sở hữu công nghiệp phụ thuộc đáng kể vào một đối tác; và kiểu dáng công nghiệp chỉ bảo hộ hình dáng "
        "bên ngoài, không bảo hộ công thức hay quy trình. Giải pháp hữu ích, loại hình phù hợp nhất với sản phẩm đề tài "
        "cấp cơ sở, vắng mặt hoàn toàn.",
        "Về bảo vệ quyền, giai đoạn nghiên cứu không ghi nhận tranh chấp, khiếu nại hay xử lý xâm phạm liên quan đến "
        "Nhà trường, phù hợp với quy mô tài sản nhỏ và mức khai thác hạn chế. Thời hạn văn bằng được theo dõi trong "
        "bảng do bộ phận quản trị thương hiệu lập; do 5 kiểu dáng được cấp cùng năm 2024, việc gia hạn sẽ dồn vào cùng "
        "một thời điểm nên cần được đưa vào sổ theo dõi thống nhất.")

    v.doan("h2", "2.3.3. Sản phẩm đề tài đủ điều kiện xác lập quyền")
    v.than("Rà soát cột sản phẩm nghiệm thu của 38 đề tài cấp cơ sở cho thấy 11 đề tài có sản phẩm cụ thể đủ điều kiện "
           f"xác lập quyền ngoài báo cáo và bài báo, được liệt kê tại Bảng 2.{v.so_bang + 1}.")
    b_dt = v.bang(
        "Đề tài cấp cơ sở có sản phẩm đủ điều kiện xác lập quyền, giai đoạn 2021 - 2025",
        ["Năm", "Đơn vị chủ trì", "Sản phẩm nghiệm thu", "Nhóm quyền có thể xác lập", "Tình trạng"],
        [["2022", "Khoa Công nghệ kỹ thuật Ô tô", "Mô hình kiểm tra và sửa chữa hệ thống khởi động điện trên ô tô",
          "Kiểu dáng công nghiệp hoặc giải pháp hữu ích", "Chưa nộp"],
         ["2022", "Khoa Dược", "Sản phẩm cầm máu dạng màng", "Giải pháp hữu ích hoặc sáng chế", "Chưa nộp"],
         ["2022", "Khoa Dược", "Bộ mẫu cây thuốc", "Sưu tập dữ liệu", "Chưa nộp"],
         ["2023", "Viện Nghiên cứu giáo dục và Chuyển giao tri thức",
          "Tinh dầu Citrus grandis, hồ sơ ghi chuyển giao cho Trường thương mại hóa", "Giải pháp hữu ích",
          "Chưa nộp"],
         ["2024", "Khoa Dược", "Xà phòng hữu cơ tảo biển kèm quy trình và công thức", "Giải pháp hữu ích",
          "Chưa nộp"],
         ["2024", "Khoa Dược", "100 tiêu bản hiển vi và cẩm nang phát hiện giun sán",
          "Sưu tập dữ liệu, quyền tác giả", "Chưa nộp"],
         ["2024", "Khoa Dược", "Rutin độ tinh khiết 90% kèm quy trình chiết xuất, tinh chế",
          "Giải pháp hữu ích hoặc sáng chế", "Chưa nộp"],
         ["2024", "Khoa Dược", "Thang thuốc 100 ml kèm công thức và quy trình sắc", "Giải pháp hữu ích",
          "Chưa nộp"],
         ["2025", "Viện Y - Dược", "Công thức cồn thuốc xoa bóp, hồ sơ ghi dự kiến đăng ký giải pháp hữu ích",
          "Giải pháp hữu ích", "Chưa nộp"],
         ["2025", "Viện Y - Dược", "Quy trình bào chế viên ngậm giảm ho, hồ sơ ghi dự kiến đăng ký giải pháp hữu ích",
          "Giải pháp hữu ích", "Chưa nộp"],
         ["2025", "Viện Y - Dược", "Chiết xuất lá Quế hoa, hồ sơ ghi giai đoạn 2 là sáng chế", "Sáng chế",
          "Đã nộp đơn năm 2025"]],
        "Nguồn: Nhóm nghiên cứu tổng hợp từ danh mục đề tài khoa học công nghệ cấp cơ sở giai đoạn 2021 - 2025. Nhóm "
        "quyền có thể xác lập là đánh giá sơ bộ dựa trên mô tả sản phẩm trong hồ sơ nghiệm thu.",
        [1.4, 3.0, 5.2, 3.3, 2.1], can=["center", "left", "left", "left", "left"])
    h_nq = v.hinh_bd("H2.14", nguon=f"Nguồn: Nhóm nghiên cứu tổng hợp từ Bảng 2.{b_dt}.")
    v.than(
        f"Hình 2.{h_nq} cho thấy {B.shcn} trên 11 sản phẩm thuộc nhóm sở hữu công nghiệp và giải pháp hữu ích phù hợp "
        f"với {ghph} sản phẩm. Loại hình này không đòi hỏi trình độ sáng tạo như sáng chế, có thời hạn bảo hộ ngắn hơn "
        "và phù hợp với quy mô đề tài cấp cơ sở, nhưng chưa sản "
        f"phẩm nào được nộp đơn theo hình thức này. Mười trên 11 sản phẩm thuộc lĩnh vực dược, nên tiềm năng sở hữu "
        "công nghiệp của Nhà trường có thể được quản lý có trọng tâm.")
    assert len(B.dt_duoc) == 10
    h_ch = v.hinh_bd("H2.15", nguon=f"Nguồn: Nhóm nghiên cứu tính toán từ danh mục đề tài cấp cơ sở và Bảng 2.{b_dt}.")
    v.than(
        f"Hình 2.{h_ch} cho thấy điểm đứt gãy nằm ở bậc thứ hai của chuỗi. Có {pt(len(B.dt_du_dk) / len(B.DE_TAI))} "
        "số đề tài tạo ra sản phẩm có thể xác lập quyền, tỷ lệ đáng ghi nhận với một trường có phần lớn ngành thuộc "
        f"kinh tế, xã hội và ngôn ngữ; nhưng chỉ {pt(len(B.nop_don) / len(B.dt_du_dk))} số đề tài đủ điều kiện được "
        f"nộp đơn. Nếu chỉ xét đề tài giao đến năm 2024, đã đủ thời gian để nộp đơn, có {len(B.du_dk_den_2024)} đề tài "
        "đủ điều kiện và không đề tài nào nộp đơn. Vấn đề không nằm ở đầu vào mà ở khâu nối giữa nghiệm thu và đăng ký.",
        "Đề tài chiết xuất lá Quế hoa năm 2025 là đề tài cấp cơ sở duy nhất trong năm năm chuyển thành đơn sáng chế, và "
        "đơn được nộp ngay trong năm nghiệm thu. Trường hợp này cho thấy kênh chuyển hóa vận hành được trong điều kiện "
        "hiện có, nhưng mới vận hành một lần, do một chủ nhiệm đồng thời chủ trì đề tài cấp quốc gia. Hai đề tài năm "
        "2025 khác ghi dự kiến đăng ký giải pháp hữu ích nhưng chưa có đơn; hai sản phẩm này vẫn có thể được nộp đơn "
        "nếu còn trong thời hạn mười hai tháng kể từ ngày bộc lộ. Đề tài tinh dầu Citrus grandis năm 2023 có liên hệ về "
        "tên gọi với kiểu dáng tinh dầu vỏ bưởi đào cấp năm 2024, nhưng hồ sơ chưa đủ để khẳng định; kể cả khi liên hệ "
        "được xác nhận, công thức và quy trình chiết xuất vẫn chưa được đăng ký. Hai danh mục do hai bộ phận lập và "
        "không có trường liên kết, là bằng chứng trực tiếp cho sự tách rời giữa luồng thương hiệu và luồng nghiên cứu.")

    # ===================================================================== 2.4
    v.doan("h1", "2.4. Thực trạng sử dụng, chuyển giao và khai thác tài sản trí tuệ")
    v.than(
        "Khai thác tài sản trí tuệ tại Nhà trường được xem xét ở ba mức: sử dụng nội bộ, chuyển giao trong hệ thống và "
        "khai thác có thu phí.",
        "Ở mức sử dụng nội bộ, 61 giáo trình và tài liệu giai đoạn 2022 - 2025 đều ghi số tín chỉ, tức gắn với học phần "
        "cụ thể; 26 tài liệu năm 2021 không có thông tin này do biểu mẫu chưa có trường số tín chỉ. Hồ sơ nghiệm thu của "
        "sáu đề tài năm 2023 ghi địa chỉ ứng dụng tại Trường, trong khi đề tài các năm khác phần lớn chỉ ghi báo cáo và "
        "bài báo; khác biệt này phụ thuộc vào biểu mẫu nghiệm thu hơn là bản chất đề tài.",
        "Ở mức chuyển giao trong hệ thống, đề tài tinh dầu Citrus grandis năm 2023 do Viện Nghiên cứu giáo dục và "
        "Chuyển giao tri thức chủ trì ghi sản phẩm là chuyển giao công nghệ cho Nhà trường thương mại hóa, nhưng không "
        "có hợp đồng, biên bản chuyển giao hay kết quả thương mại hóa kèm theo. Ý định chuyển giao được ghi vào hồ sơ "
        "nhưng không có bước nào biến ý định thành giao dịch.",
        "Ở mức khai thác có thu phí, thống kê của Viện Nghiên cứu giáo dục và Chuyển giao tri thức ghi nhận năm hợp "
        "đồng giai đoạn 2023 - 2024. Hai hợp đồng chuyển giao quyền sử dụng tác phẩm đối với sách “Giáo dục và khoa học "
        "mở: Cẩm nang hướng dẫn” và “Từng bước nhập môn Nghiên cứu định lượng”, với mức 10% trên tổng số bán được, là "
        "giao dịch khai thác quyền đúng nghĩa. Ba hợp đồng còn lại là dịch vụ đào tạo, giảng dạy, tổng trị giá 100 triệu "
        "đồng, thuộc khai thác tri thức chuyên môn chứ không phải khai thác một tài sản đã xác lập quyền.",
        "Như vậy, khai thác tài sản trí tuệ hiện chủ yếu ở dạng phi thương mại, phục vụ đào tạo của chính Nhà trường. "
        "Khai thác có thu phí mới có hai hợp đồng, do một đơn vị thực hiện và đối với sách chuyên khảo; chín văn bằng đã "
        "được cấp chưa phát sinh giao dịch chuyển giao quyền nào. Năng lực khai thác gắn với tư cách pháp lý và chức "
        "năng được giao của Viện Nghiên cứu giáo dục và Chuyển giao tri thức, không phân bố theo năng lực chuyên môn. "
        "Đây cũng là nội dung có dữ liệu mỏng nhất, vì hệ thống thống kê không theo dõi việc sử dụng và khai thác tài "
        "sản trí tuệ; bản thân việc thiếu dữ liệu này là một phát hiện về thực trạng quản lý.")

    # ===================================================================== 2.5
    v.doan("h1", "2.5. Đánh giá chung về công tác quản lý quyền sở hữu trí tuệ tại Trường Đại học Thành Đô")
    v.than(
        "Kết quả hai năm 2024 - 2025 được đối chiếu trước hết với Kế hoạch hoạt động khoa học công nghệ giai đoạn 2024 - "
        "2028 ban hành kèm Kế hoạch số 07/KH-ĐHTĐ ngày 01 tháng 7 năm 2024. Đây là thước đo khách quan vì do chính Nhà "
        "trường đặt ra, và 2024 - 2025 là hai năm đầu của kế hoạch.")
    h_kh = v.hinh_bd("H2.16")
    k = kh
    v.than(
        f"Hình 2.{h_kh} cho thấy {kh_dat} trên 10 chỉ tiêu đạt hoặc vượt kế hoạch, nhưng kết quả phân hóa rõ theo nhóm. "
        f"Nhóm công bố vượt xa chỉ tiêu: bài báo quốc tế đạt {k['Bài báo tạp chí quốc tế'][3]} bài so với "
        f"{k['Bài báo tạp chí quốc tế'][2]} bài, tức {pt(k['Bài báo tạp chí quốc tế'][4], 0)}; sách "
        f"{pt(k['Sách có chỉ số ISBN'][4], 0)}; tham luận cấp trường {pt(k['Tham luận hội thảo cấp trường'][4], 0)}; "
        f"bài báo trong nước {pt(k['Bài báo tạp chí trong nước, gồm tạp chí của Trường'][4], 0)}. Nhóm đề tài và học "
        f"liệu dao động quanh kế hoạch, từ {pt(k['Giáo trình, tài liệu tham khảo'][4], 0)} với giáo trình đến "
        f"{pt(k['Đề tài cấp Bộ, Nhà nước'][4], 0)} với đề tài cấp Bộ, Nhà nước. Ở nhóm tài sản trí tuệ, chỉ tiêu văn "
        f"bằng đạt {pt(k['Công nhận sáng chế, kiểu dáng, quyền tác giả'][4], 0)} với "
        f"{k['Công nhận sáng chế, kiểu dáng, quyền tác giả'][3]} văn bằng, nhưng đó là 5 kiểu dáng đồng sở hữu với "
        "doanh nghiệp và nhãn hiệu Thado Edupark, không văn bằng nào hình thành từ đề tài; chỉ tiêu chuyển giao công "
        "nghệ năm 2025 không đạt.",
        "Do Kế hoạch gộp sáng chế với kiểu dáng công nghiệp và quyền tác giả trong một chỉ tiêu, chỉ tiêu này có thể "
        "hoàn thành mà không cần kết quả nghiên cứu nào được bảo hộ. Kế hoạch 07/KH-ĐHTĐ từng tự đánh giá giai đoạn "
        "2019 - 2023 là chưa có công trình được chuyển giao công nghệ và tài sản trí tuệ còn hạn chế; sau hai năm thực "
        "hiện, nhận định này về cơ bản vẫn giữ nguyên.")

    v.doan("h2", "2.5.1. Kết quả đạt được")
    v.than(
        "Thứ nhất, Nhà trường có hệ thống quy định về sở hữu trí tuệ từ sớm. Quyết định 213 năm 2021 đã có một chương "
        "riêng với phạm vi tài sản khá đầy đủ và quy trình đăng ký một cửa; Quyết định 217 năm 2024 mở rộng phạm vi tài "
        "sản, bổ sung nguyên tắc công bố và bảo mật, phân công đầu mối.",
        "Thứ hai, Nhà trường đã xác lập 12 tài sản trí tuệ, trong đó 9 tài sản có văn bằng; nhóm thương hiệu gồm tên "
        "trường, bộ nhận diện và thương hiệu hệ sinh thái được bảo hộ liên tục từ năm 2021, có ý nghĩa trong cạnh tranh "
        "tuyển sinh.",
        f"Thứ ba, năng lực nghiên cứu tăng nhanh: số bài báo tăng bình quân {pt(B.cagr(B.bb[0], B.bb[4], 4))} một năm, "
        f"đạt {so(B.bb[4] / tong_gv, 2)} bài trên một giảng viên năm 2025; {kh_dat} trên 10 chỉ tiêu khoa học công "
        "nghệ của Kế hoạch 07/KH-ĐHTĐ đạt hoặc vượt; ba đề tài cấp quốc gia với tổng kinh phí 4,67 tỷ đồng được giao.",
        "Thứ tư, kênh chuyển hóa từ đề tài sang quyền sở hữu công nghiệp đã vận hành được, với đơn sáng chế từ đề tài "
        "chiết xuất lá Quế hoa năm 2025 và đơn sáng chế thứ hai năm 2026.",
        "Thứ năm, Nhà trường đã có nguồn lực tài chính gắn cơ chế sở hữu trí tuệ tương đối hoàn chỉnh là Quỹ Học bổng "
        "sau tiến sĩ Ngô Xuân Độ với ngân sách 5 tỷ đồng, và có kết nối với Mạng lưới Trung tâm Hỗ trợ công nghệ và đổi "
        "mới sáng tạo từ năm 2023.")

    v.doan("h2", "2.5.2. Hạn chế")
    v.than(
        f"Thứ nhất, kết quả nghiên cứu hầu như không được chuyển hóa thành quyền: {len(B.dt_du_dk)} trên 38 đề tài có "
        f"sản phẩm đủ điều kiện nhưng chỉ {len(B.nop_don)} đề tài nộp đơn; {len(B.du_dk_den_2024)} đề tài đủ điều kiện "
        "giai đoạn 2021 - 2024 đều không nộp đơn; hoạt động nghiên cứu chỉ đóng góp 2 trên 12 tài sản trí tuệ.",
        "Thứ hai, hệ thống quy định chồng lấn và chưa theo kịp pháp luật: bốn văn bản chứa năm quy định chia lợi ích, "
        "và mức trần 100 triệu đồng không còn tương thích với Điều 135 Luật Sở hữu trí tuệ hiện hành.",
        "Thứ ba, quản lý quyền diễn ra theo hai luồng tách rời: luồng thương hiệu vận hành đều và có kết quả, luồng "
        "nghiên cứu gần như không; danh mục tài sản và danh mục sản phẩm đề tài gần như không có điểm giao.",
        "Thứ tư, năng lực xác lập quyền sở hữu công nghiệp phụ thuộc đối tác: 6 trên 12 tài sản, gồm toàn bộ kiểu dáng "
        "công nghiệp, đồng sở hữu với một doanh nghiệp.",
        f"Thứ năm, hệ thống dữ liệu chưa đáp ứng yêu cầu quản lý và nghĩa vụ công khai: không có danh mục tài "
        f"sản trí tuệ trong hệ thống thống kê, độ phủ khai báo khoảng 29,9%, {B.tong_C} trên 16 tiêu chí đánh giá chưa "
        "tính được.",
        "Thứ sáu, khai thác có thu phí rất hẹp: hai hợp đồng chuyển giao quyền sử dụng tác phẩm trong năm năm, chín văn "
        "bằng chưa phát sinh giao dịch chuyển giao quyền.",
        f"Thứ bảy, nghiệm thu chưa gắn với sản phẩm có thể bảo hộ: {xep_loai['Xuất sắc'] + xep_loai['Tốt']} trên 38 đề "
        "tài xếp loại Tốt trở lên, 7 đề tài chỉ có báo cáo tổng kết nhưng 5 trong số đó vẫn xếp loại Tốt, và quy trình "
        "nghiệm thu không có khâu đánh giá khả năng bảo hộ.")

    v.doan("h2", "2.5.3. Nguyên nhân của hạn chế")
    v.doan("h3", "a) Nguyên nhân khách quan", bold=True)
    v.than(
        "Thứ nhất, khung pháp luật thay đổi dồn dập trong giai đoạn 2025 - 2026, gồm Luật Khoa học, công nghệ và đổi mới "
        "sáng tạo số 93/2025/QH15 có hiệu lực từ ngày 01 tháng 10 năm 2025, Luật Giáo dục đại học số 125/2025/QH15, Luật "
        "số 131/2025/QH15 sửa đổi Luật Sở hữu trí tuệ có hiệu lực từ ngày 01 tháng 4 năm 2026, Kết luận số 51-KL/TW "
        "ngày 17 tháng 6 năm 2026 và Quyết định số 1624/QĐ-TTg ngày 21 tháng 8 năm 2026. Quy chế ban hành năm 2021 được xây dựng trên khung cũ nên một số quy định như mức "
        "trần thù lao không còn tương thích.",
        "Thứ hai, thủ tục xác lập quyền sở hữu công nghiệp kéo dài và phát sinh chi phí tra cứu, soạn đơn, lệ phí, phí "
        "duy trì; là trường tư thục, Nhà trường phải tự cân đối các khoản này từ nguồn thu của mình.",
        "Thứ ba, cơ cấu ngành chủ yếu thuộc kinh tế, quản lý, ngôn ngữ, giáo dục và pháp luật, vốn chủ yếu tạo ra tác "
        "phẩm thuộc quyền tác giả; tiềm năng sở hữu công nghiệp tập trung ở lĩnh vực dược.",
        f"Các nguyên nhân khách quan giải thích vì sao quy mô tài sản trí tuệ còn nhỏ, nhưng không giải thích được vì sao "
        f"{len(B.du_dk_den_2024)} đề tài đủ điều kiện giai đoạn 2021 - 2024 đều không được nộp đơn.")
    v.doan("h3", "b) Nguyên nhân chủ quan", bold=True)
    v.than(
        "Về thể chế, bốn văn bản chứa năm quy định chia lợi ích với phạm vi không rõ ranh giới; Quyết định 217 không "
        "thay thế Chương VI Quyết định 213, nên người có sản phẩm không xác định được quy định nào áp dụng.",
        "Về quy trình, Điều 35 Quyết định 213 thiết kế theo cơ chế tác giả chủ động nộp đơn, không có khâu rà soát bắt "
        "buộc khả năng bảo hộ khi nghiệm thu; Điều 10 Quyết định 217 tiếp tục đặt trách nhiệm tự nhận diện lên tác giả. "
        "Quy trình có trên văn bản nhưng không được kích hoạt.",
        "Về nguồn lực, kênh sinh ra sản phẩm có thể bảo hộ không có dòng chi cho phí nộp đơn và duy trì hiệu lực; kênh "
        "có đủ cơ chế là Quỹ Học bổng sau tiến sĩ chỉ dành cho đối tượng hẹp; đơn vị đầu mối chỉ có 2 nhân sự và không "
        "có người chuyên trách về sở hữu trí tuệ.",
        "Về động lực, văn bằng được quy đổi từ 180 đến 600 giờ nghiên cứu nhưng không có tiền thưởng và chỉ được ghi "
        "nhận sau khi được cấp, trong khi bài báo quốc tế được thưởng từ 10 đến 20 triệu đồng ngay khi đăng.",
        "Về tổ chức, chức năng quản lý bị chia cho bốn đơn vị, ba đơn vị thuộc khối Quản trị và Dịch vụ, và không đơn "
        "vị nào thực hiện trọn chu trình từ nhận diện đến khai thác.",
        "Về dữ liệu, không có danh mục tài sản trí tuệ và độ phủ khai báo thấp nên đơn vị đầu mối không có công cụ phát "
        "hiện sản phẩm cần rà soát.",
        "Về con người, năng lực nghiên cứu tập trung ở khoảng một phần ba nhân sự, và trường hợp chuyển hóa thành công "
        "duy nhất thuộc về một chủ nhiệm đồng thời chủ trì đề tài cấp quốc gia, nên hoạt động xác lập quyền mang tính "
        "đơn lẻ.",
        "Bảy nguyên nhân chủ quan liên kết thành một chuỗi: thiếu dữ liệu nên đầu mối không phát hiện sản phẩm; thiếu "
        "khâu rà soát nên trách nhiệm dồn về tác giả; tác giả đối diện quy định không thống nhất, không có kinh phí nộp "
        "đơn và phần thưởng đến chậm nên chọn công bố; sau mười hai tháng, sản phẩm mất khả năng xác lập quyền sáng chế "
        "và giải pháp hữu ích. Việc Nhà trường tham gia Mạng lưới Trung tâm Hỗ trợ công nghệ và đổi mới sáng tạo từ năm "
        "2023 mà hai năm 2023 - 2024 không có đơn nào xuất phát từ đề tài cho thấy điểm nghẽn không nằm ở khả năng tiếp "
        "cận thông tin mà ở khâu nhận diện và khởi động quy trình. Hệ thống giải pháp tại Chương 3 vì vậy cần tác động "
        "đồng thời vào cả chuỗi, trong đó khâu rà soát khả năng bảo hộ tại thời điểm nghiệm thu là điểm can thiệp ưu "
        "tiên.")


    v.doan("h1", "TIỂU KẾT CHƯƠNG 2")
    v.than(
        "Chương 2 đã phân tích thực trạng quản lý quyền sở hữu trí tuệ tại Trường Đại học Thành Đô giai đoạn 2021 - 2025 "
        "trên cơ sở các danh mục thống kê, văn bản nội bộ và kế hoạch của Nhà trường. Kết quả cho thấy Nhà trường có nền "
        "tảng thể chế từ sớm, năng lực công bố tăng nhanh, 12 tài sản trí tuệ đã được xác lập hoặc đang xử lý đơn và kênh "
        "chuyển hóa từ đề tài sang đơn sáng chế đã vận hành. Tuy nhiên, điểm nghẽn cốt lõi nằm ở khâu nối giữa nghiệm thu và "
        f"đăng ký: {len(B.dt_du_dk)} trên 38 đề tài có sản phẩm đủ điều kiện xác lập quyền nhưng chỉ {len(B.nop_don)} đề "
        "tài được nộp đơn, hoạt động nghiên cứu chỉ đóng góp 2 trên 12 tài sản, và khai thác có thu phí mới dừng ở hai hợp "
        "đồng chuyển giao quyền sử dụng tác phẩm.",
        "Nguyên nhân chủ yếu thuộc về chủ quan: bốn văn bản với năm quy định chia lợi ích chưa thống nhất, quy trình thiếu "
        "khâu rà soát bắt buộc, kênh đề tài cấp cơ sở không có kinh phí nộp đơn, cơ chế khuyến khích nghiêng về công bố, "
        "chức năng quản lý phân tán ở bốn đơn vị và hệ thống dữ liệu chưa theo dõi tài sản trí tuệ. Các nguyên nhân này liên "
        "kết thành một chuỗi, nên hệ thống giải pháp tại Chương 3 cần tác động đồng thời, trong đó khâu rà soát khả năng "
        "bảo hộ tại thời điểm nghiệm thu là điểm can thiệp ưu tiên.")

# ---------------------------------------------------------------------------
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

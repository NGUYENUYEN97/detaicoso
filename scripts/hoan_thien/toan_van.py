# -*- coding: utf-8 -*-
"""Hoàn thiện báo cáo toàn văn theo mẫu của Trường và thuyết minh đề tài, trên bản BAN_CUOI của nhóm tác giả.

Các việc: thêm Mã số và Lời cam đoan; Mở đầu đúng 7 mục của mẫu; Chương 2 bổ sung nhận thức sở hữu trí tuệ;
Chương 3 xếp thành 3 nhóm giải pháp như thuyết minh, thêm Giải pháp 4 về số hóa; dựng lại mục lục, danh mục bảng, hình.
"""
import copy
import os
import re
import sys
import tempfile

from docx import Document
from docx.oxml.ns import qn

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "rut_gon"))
from dung import tach_markup, ve_pdf, tim_trang  # noqa: E402

GOC = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
THU_MUC = os.path.join(GOC, "Ban_cuoi", "Ban_hoan_thien_03-10-2026")
VAO = os.path.join(THU_MUC, "Bao_cao_tong_ket_de_tai_BAN_CUOI.docx")
RA = os.path.join(THU_MUC, "Bao_cao_toan_van_de_tai.docx")
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


class Sua:
    def __init__(self):
        self.doc = Document(VAO)
        self.body = self.doc.element.body
        self.bm = 20000
        els = self.ds()
        self.mau = {
            "H0": copy.deepcopy(self.tim("MỞ ĐẦU", kieu="H0")),
            "H1": copy.deepcopy(self.tim("1. Tính cấp thiết của đề tài")),
            "H2": copy.deepcopy(self.tim("2.1. Tình hình nghiên cứu trong nước")),
            "H3": copy.deepcopy(self.tim("a) Mục tiêu")),
            "P": copy.deepcopy(self.tim("Tài sản trí tuệ là nguồn lực quan trọng")),
        }
        del els

    # ---- tiện ích ----
    def ds(self):
        return list(self.body)

    def tim(self, dau, sau=None, kieu=None):
        """Đoạn đầu tiên (sau phần danh mục) có chữ bắt đầu bằng `dau`."""
        els = self.ds()
        bat_dau = els.index(sau) + 1 if sau is not None else 0
        for el in els[bat_dau:]:
            if el.tag != qn("w:p"):
                continue
            s = chu(el)
            if not s.startswith(dau):
                continue
            ppr = el.find(qn("w:pPr"))
            st = ppr.find(qn("w:pStyle")) if ppr is not None else None
            if st is not None and st.get(qn("w:val")).startswith("TOC"):
                continue
            if kieu == "H0" and (ppr is None or ppr.find(qn("w:outlineLvl")) is None):
                continue
            return el
        raise KeyError(dau)

    def bookmark(self, p):
        ten = f"_TocHT{self.bm}"
        bs = p.makeelement(qn("w:bookmarkStart"), {qn("w:id"): str(self.bm), qn("w:name"): ten})
        be = p.makeelement(qn("w:bookmarkEnd"), {qn("w:id"): str(self.bm)})
        self.bm += 1
        ppr = p.find(qn("w:pPr"))
        p.insert(list(p).index(ppr) + 1 if ppr is not None else 0, bs)
        p.append(be)

    def doan(self, loai, s, mau=None):
        mau = mau if mau is not None else self.mau[loai]
        p = bo_id(copy.deepcopy(mau))
        r0 = mau.find(".//" + qn("w:r"))
        rpr = r0.find(qn("w:rPr")) if r0 is not None else None
        for c in list(p):
            if c.tag != qn("w:pPr"):
                p.remove(c)
        for t, dam, nghieng in tach_markup(s):
            r = p.makeelement(qn("w:r"), {})
            rp = copy.deepcopy(rpr) if rpr is not None else r.makeelement(qn("w:rPr"), {})
            for th in ("w:b", "w:i"):
                if rp.find(qn(th)) is not None and loai == "P":
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
        if loai in ("H0", "H1", "H2"):
            self.bookmark(p)
        return p

    def cac_doan(self, ds):
        return [self.doan(loai, s) for loai, s in ds]

    @staticmethod
    def dat_chu(p, s):
        ts = list(p.iter(qn("w:t")))
        ts[0].text = s
        ts[0].set(XML_SPACE, "preserve")
        for t in ts[1:]:
            t.text = ""

    def thay(self, dau, cu, moi):
        p = self.tim(dau)
        s = chu(p)
        assert cu in s, (dau, cu)
        self.dat_chu(p, s.replace(cu, moi))
        return p

    @staticmethod
    def chen_sau(neo, moi):
        for el in moi:
            neo.addnext(el)
            neo = el
        return neo

    @staticmethod
    def chen_truoc(neo, moi):
        for el in moi:
            neo.addprevious(el)

    def xoa_khoang(self, dau, cuoi):
        """Xóa từ đoạn `dau` đến trước đoạn `cuoi`."""
        els = self.ds()
        a, b = els.index(dau), els.index(cuoi)
        for el in els[a:b]:
            self.body.remove(el)

    # ---- các phần sửa ----
    def bia_va_cam_doan(self):
        ten = self.tim("NGHIÊN CỨU GIẢI PHÁP NÂNG CAO")
        cn = self.tim("Chủ nhiệm đề tài:")
        ma = bo_id(copy.deepcopy(cn))
        self.dat_chu(ma, "Mã số: ……………………………………")
        cn.addprevious(ma)

        ml = self.tim("MỤC LỤC")
        td = bo_id(copy.deepcopy(ml))
        for x in td.iter(qn("w:lastRenderedPageBreak")):
            x.getparent().remove(x)
        self.dat_chu(td, "LỜI CAM ĐOAN")
        than = self.cac_doan([
            ("P", "Nhóm thực hiện đề tài “Nghiên cứu giải pháp nâng cao hiệu quả quản lý quyền sở hữu trí tuệ tại Trường "
                  "Đại học Thành Đô” xin cam đoan đây là kết quả nghiên cứu của nhóm. Số liệu, kết quả trong báo cáo được "
                  "tổng hợp từ hồ sơ, văn bản của Trường Đại học Thành Đô và các nguồn công khai, có ghi nguồn đầy đủ. Dữ "
                  "liệu cá nhân chỉ được sử dụng ở dạng tổng hợp."),
            ("P", "Kết quả nghiên cứu chưa được công bố trong công trình nào khác, trừ bài báo là sản phẩm của đề tài. Nhóm "
                  "thực hiện đề tài chịu trách nhiệm về tính trung thực của báo cáo."),
        ])
        ky = []
        for s, dam in (("Hà Nội, ngày …… tháng …… năm 2026", False), ("Chủ nhiệm đề tài", True), ("", False),
                       ("", False), ("Nguyễn Thị Tố Uyên", True)):
            p = self.doan("P", ("**" + s + "**") if dam and s else (s or " "))
            ppr = p.find(qn("w:pPr"))
            for th in ("w:jc", "w:ind"):
                if ppr.find(qn(th)) is not None:
                    ppr.remove(ppr.find(qn(th)))
            vt = ppr.find(qn("w:rPr"))
            ind = ppr.makeelement(qn("w:ind"), {qn("w:left"): "4500", qn("w:firstLine"): "0"})
            jc = ppr.makeelement(qn("w:jc"), {qn("w:val"): "center"})
            if vt is not None:
                vt.addprevious(ind)
                vt.addprevious(jc)
            else:
                ppr.append(ind)
                ppr.append(jc)
            ky.append(p)
        self.chen_truoc(ml, [td] + than + ky)

    def mo_dau(self):
        dau = self.tim("3. Mục tiêu và câu hỏi nghiên cứu")
        tq = chu(self.tim("Mục tiêu tổng quát:"))
        ct = chu(self.tim("Mục tiêu cụ thể:"))
        ch = chu(self.tim("Câu hỏi nghiên cứu:"))
        du_lieu = chu(self.tim("Phương pháp phân tích văn bản được dùng"))
        lam_sach = chu(self.tim("Dữ liệu được làm sạch"))
        gioi_han = chu(self.tim("Giới hạn phương pháp."))
        quy_uoc = chu(self.tim("Quy ước tên gọi văn bản."))
        ch1 = self.tim("CHƯƠNG 1", kieu="H0")
        self.xoa_khoang(dau, ch1)
        du_lieu = du_lieu.replace("Phương pháp phân tích văn bản được dùng để đối chiếu quy chế, kế hoạch của Nhà trường "
                                  "với pháp luật hiện hành. ", "")
        moi = [
            ("H1", "3. Mục tiêu nghiên cứu"),
            ("P", "**Mục tiêu tổng quát:**" + tq[len("Mục tiêu tổng quát:"):]),
            ("P", "**Mục tiêu cụ thể:**" + ct[len("Mục tiêu cụ thể:"):]),
            ("P", "**Câu hỏi nghiên cứu:**" + ch[len("Câu hỏi nghiên cứu:"):]),
            ("H1", "4. Đối tượng nghiên cứu"),
            ("P", "Đối tượng nghiên cứu là hoạt động quản lý quyền sở hữu trí tuệ tại Trường Đại học Thành Đô, gồm thể chế, "
                  "tổ chức, nguồn lực, dữ liệu và chu trình tạo lập, xác lập, bảo vệ, khai thác tài sản trí tuệ, cùng các "
                  "giải pháp nâng cao hiệu quả quản lý."),
            ("H1", "5. Phạm vi nghiên cứu"),
            ("P", "- **Về nội dung:** toàn bộ chu trình quản lý quyền sở hữu trí tuệ, gồm xác lập quyền, bảo hộ, khai thác, "
                  "thương mại hóa và quản lý nội bộ; tài sản trí tuệ phát sinh từ hoạt động khoa học công nghệ và đào tạo "
                  "của giảng viên, người lao động và người học."),
            ("P", "- **Về không gian:** Trường Đại học Thành Đô và các đơn vị trực thuộc. Thuyết minh xác định phạm vi gồm cả "
                  "các đơn vị trong hệ sinh thái giáo dục. Khi thực hiện, Trường Tiểu học và Trung học cơ sở UNIGO được xác "
                  "định là pháp nhân riêng, tài sản trí tuệ thuộc sở hữu của pháp nhân này, nên không đưa vào số liệu thực "
                  "trạng; nhãn hiệu của hệ sinh thái vẫn được xem xét trong giải pháp quản lý nhãn hiệu."),
            ("P", "- **Về thời gian:** số liệu thực trạng giai đoạn 2021 - 2025. Thuyết minh dự kiến thu thập số liệu từ năm "
                  "2020; tuy nhiên dữ liệu thống kê đầy đủ mà đề tài thu thập được bắt đầu từ năm 2021, trùng với năm ban "
                  "hành Quyết định 213, nên năm 2021 được chọn làm mốc. Danh sách nhân sự lấy năm 2026; văn bản pháp luật "
                  "cập nhật đến hết tháng 9 năm 2026; giải pháp hướng tới giai đoạn 2026 - 2030."),
            ("H1", "6. Phương pháp nghiên cứu"),
            ("P", "Đề tài kết hợp ba cách tiếp cận đã xác định trong thuyết minh. Tiếp cận hệ thống xem quản lý quyền sở hữu "
                  "trí tuệ trong mối quan hệ với thể chế, tổ chức, nguồn lực và dữ liệu của nhà trường. Tiếp cận đa bên xem "
                  "xét quyền lợi của nhà trường, giảng viên, người học và doanh nghiệp. Tiếp cận nghiên cứu trường hợp phân "
                  "tích cụ thể tại Trường Đại học Thành Đô. Khung phân tích dựa trên chu trình tạo lập, xác lập, khai thác và "
                  "bảo vệ, kết hợp bộ tiêu chí theo chuỗi đầu vào, quá trình, đầu ra và kết quả tại Chương 1. Các phương "
                  "pháp cụ thể gồm:"),
            ("P", "- **Phân tích, tổng hợp tài liệu:** hệ thống hóa lý luận, các nghiên cứu trong và ngoài nước và văn bản "
                  "pháp luật hiện hành về quản lý quyền sở hữu trí tuệ."),
            ("P", "- **So sánh:** đối chiếu quy chế nội bộ của Nhà trường với pháp luật hiện hành; đối chiếu kinh nghiệm trong "
                  "và ngoài nước để rút bài học; đối chiếu kết quả thực hiện với chỉ tiêu của Kế hoạch 07."),
            ("P", "- **Nghiên cứu trường hợp dựa trên dữ liệu hành chính:** " + du_lieu),
            ("P", lam_sach),
            ("P", "- **Phương pháp chuyên gia:** thuyết minh dự kiến tham vấn ý kiến chuyên gia, khảo sát và phỏng vấn sâu. "
                  "Trong thời gian thực hiện, đề tài chưa tổ chức tham vấn chuyên gia, khảo sát hay phỏng vấn chính thức; "
                  "các nhận định được kiểm tra bằng đối chiếu với văn bản gốc và dữ liệu hành chính. Việc lấy ý kiến của chủ "
                  "nhiệm đề tài, Hội đồng nghiệm thu và khảo sát nhận thức của giảng viên được đưa vào kế hoạch thí điểm tại "
                  "Mục 3.5.2."),
            ("P", "**Giới hạn phương pháp.**" + gioi_han[len("Giới hạn phương pháp."):]),
            ("P", "**Quy ước tên gọi văn bản.**" + quy_uoc[len("Quy ước tên gọi văn bản."):]),
            ("H1", "7. Bố cục của đề tài"),
            ("P", "Ngoài Mở đầu, Kết luận và kiến nghị, Tài liệu tham khảo, báo cáo gồm ba chương: Chương 1. Cơ sở lý luận về "
                  "quản lý quyền sở hữu trí tuệ trong cơ sở giáo dục đại học; Chương 2. Thực trạng quản lý quyền sở hữu trí "
                  "tuệ tại Trường Đại học Thành Đô; Chương 3. Hệ thống giải pháp nâng cao hiệu quả quản lý quyền sở hữu trí "
                  "tuệ tại Trường Đại học Thành Đô đáp ứng khung pháp lý mới."),
        ]
        self.chen_truoc(ch1, self.cac_doan(moi))

    def chuong_2(self):
        self.thay("2.4. Thực trạng sử dụng, chuyển giao và khai thác tài sản trí tuệ",
                  "2.4. Thực trạng sử dụng, chuyển giao và khai thác tài sản trí tuệ",
                  "2.4. Thực trạng khai thác tài sản trí tuệ và nhận thức về sở hữu trí tuệ")
        neo = self.tim("Như vậy, khai thác chủ yếu ở dạng phi thương mại")
        self.chen_sau(neo, self.cac_doan([
            ("P", "Về nhận thức, đề tài không khảo sát trực tiếp. Một số dữ kiện hành vi cho thấy kỹ năng nhận diện và bảo "
                  "hộ tài sản trí tuệ của đội ngũ còn hạn chế: 6 đề tài mã số 2021 - 2024 có sản phẩm tiềm năng sở hữu công "
                  "nghiệp nhưng chưa có đơn; hai đề tài năm 2025 ghi dự kiến đăng ký giải pháp hữu ích nhưng chưa có đơn; 87 "
                  "giáo trình chưa được đăng ký quyền tác giả; Kế hoạch 07 thống kê 12 lượt tập huấn chung giai đoạn 2019 - "
                  "2023 nhưng không tách riêng tập huấn về sở hữu trí tuệ."),
            ("P", "Ở chiều ngược lại, đơn sáng chế từ đề tài chiết xuất lá Quế hoa và 5 kiểu dáng công nghiệp hợp tác với "
                  "doanh nghiệp cho thấy một nhóm giảng viên đã có ý thức bảo hộ. Mức độ hiểu biết của đội ngũ cần được đo "
                  "bằng khảo sát trong kế hoạch thí điểm tại Mục 3.5.2."),
        ]))

    def chuong_3(self):
        self.thay("3.1. Căn cứ, định hướng", "3.1. Căn cứ, định hướng và nguyên tắc xây dựng giải pháp",
                  "3.1. Quan điểm, mục tiêu và căn cứ xây dựng giải pháp")
        h313 = self.thay("3.1.3. Nguyên tắc xây dựng giải pháp", "3.1.3. Nguyên tắc xây dựng giải pháp",
                         "3.1.3. Quan điểm, nguyên tắc và mục tiêu")
        self.chen_sau(h313, self.cac_doan([
            ("P", "Hệ thống giải pháp được xây dựng theo quan điểm chuyển từ quản lý hành chính sang đồng hành và kiến tạo, "
                  "bám căn cứ pháp lý và điều kiện thực tế của Nhà trường, với năm nguyên tắc:"),
        ]))
        cuoi_nt = self.tim("- Hài hòa lợi ích.")
        self.chen_sau(cuoi_nt, self.cac_doan([
            ("P", "**Mục tiêu đến năm 2030:** ban hành Quy chế quản trị tài sản trí tuệ sửa đổi trước ngày 31 tháng 5 năm "
                  "2027; từ năm 2027, mọi đề tài nghiệm thu có phiếu rà soát; số đơn sở hữu công nghiệp từ kết quả nghiên "
                  "cứu đạt từ 2 đơn năm 2027 và 3 đến 5 đơn mỗi năm từ năm 2028; có ít nhất 1 hợp đồng chuyển giao hoặc cấp "
                  "phép mỗi năm; dữ liệu tài sản trí tuệ đủ để báo cáo trên HEMIS. Các mức này là mục tiêu tham khảo, được "
                  "điều chỉnh sau thí điểm."),
            ("P", "Từ các căn cứ trên, Đề tài đề xuất sáu giải pháp, xếp thành ba nhóm theo nội dung đã đăng ký trong thuyết "
                  "minh: nhóm hoàn thiện thể chế, quy chế nội bộ và bộ máy (Mục 3.2); nhóm chuẩn hóa quy trình và số hóa "
                  "công tác quản lý (Mục 3.3); nhóm thúc đẩy thương mại hóa và nâng cao nhận thức (Mục 3.4). Mỗi giải pháp "
                  "trình bày theo bốn phần: mục tiêu, nội dung, chủ thể thực hiện và điều kiện bảo đảm. Tổ chức thực hiện, "
                  "thí điểm, lộ trình và chỉ số theo dõi được trình bày tại Mục 3.5."),
        ]))

        # Nhóm 1
        h32 = self.tim("3.2. Hệ thống giải pháp cụ thể")
        self.dat_chu(h32, "3.2. Nhóm giải pháp hoàn thiện thể chế, quy chế nội bộ và bộ máy quản lý")
        self.dat_chu(self.tim("Đề tài đề xuất năm nhóm giải pháp."),
                     "Nhóm này xử lý hạn chế thứ hai và thứ ba tại Mục 2.5.2, gồm Giải pháp 1 và Giải pháp 2.")
        self.thay("Giai đoạn 1 không tăng biên chế.", "dòng dự toán tại Giải pháp 4", "dòng dự toán tại Giải pháp 5")

        # Nhóm 2
        h3 = self.tim("3.2.3. Giải pháp 3:")
        self.dat_chu(h3, "3.3.1. Giải pháp 3: Chuẩn hóa quy trình 8 khâu, sàng lọc trước khi công bố và nộp đơn sớm")
        self.chen_truoc(h3, self.cac_doan([
            ("H1", "3.3. Nhóm giải pháp chuẩn hóa quy trình và số hóa công tác quản lý"),
            ("P", "Nhóm này xử lý hạn chế thứ nhất, thứ năm và thứ bảy tại Mục 2.5.2, gồm Giải pháp 3 và Giải pháp 4."),
        ]))
        dk3 = self.tim("Phiếu khai báo, phiếu rà soát ban hành kèm quy trình;")
        self.dat_chu(dk3, "Phiếu khai báo, phiếu rà soát ban hành kèm quy trình; tài khoản công cụ tra cứu sáng chế; danh "
                          "mục số tài sản trí tuệ theo Giải pháp 4; nhân sự đầu mối theo Giải pháp 2.")
        self.chen_sau(dk3, self.cac_doan([
            ("H2", "3.3.2. Giải pháp 4: Số hóa dữ liệu và lập danh mục số tài sản trí tuệ"),
            ("H3", "a) Mục tiêu"),
            ("P", "Xử lý hạn chế thứ năm. Chương 2 cho thấy chưa có danh mục tài sản trí tuệ trong hệ thống thống kê; hai "
                  "bảng theo dõi chưa thống nhất trạng thái 5 kiểu dáng công nghiệp; danh mục bài báo không ghi mã đề tài; 7 "
                  "trên 16 tiêu chí đánh giá chưa tính được. Giải pháp nhằm có một nguồn dữ liệu duy nhất để theo dõi tài sản "
                  "từ khai báo đến khai thác, đáp ứng yêu cầu theo dõi riêng kết quả từ ngân sách của Nghị định 267, nghĩa vụ "
                  "công khai của Luật Giáo dục đại học và báo cáo trên HEMIS theo Thông tư 83."),
            ("H3", "b) Nội dung"),
            ("P", "- **Danh mục số dùng chung.** Mỗi tài sản có một mã riêng và các trường: mã đề tài, loại đối tượng, chủ "
                  "thể quyền, đồng tác giả và mức đóng góp, nguồn kinh phí, thời điểm giao nhiệm vụ, ngày bộc lộ, số đơn, "
                  "trạng thái theo bốn nhóm thống nhất, hạn nộp phí duy trì, hợp đồng khai thác và nguồn thu. Kết quả hình "
                  "thành từ ngân sách nhà nước được đánh dấu để theo dõi riêng."),
            ("P", "- **Làm sạch dữ liệu hiện có.** Hợp nhất hai bảng theo dõi đơn và văn bằng; xác minh trạng thái 5 kiểu "
                  "dáng công nghiệp; bổ sung mã đề tài vào danh mục bài báo từ năm 2026; thống nhất tên đơn vị và cách ghi "
                  "năm."),
            ("P", "- **Phân quyền và công khai.** Hồ sơ chưa nộp đơn, bí mật kinh doanh và thông tin bị ràng buộc bởi hợp "
                  "đồng chỉ người được giao xử lý mới truy cập được. Dữ liệu được phép công bố được trích riêng cho báo cáo "
                  "trên HEMIS, Nền tảng số quốc gia và chuyên trang tại Giải pháp 6."),
            ("P", "- **Cảnh báo theo hạn.** Danh mục nhắc hạn nộp đơn trước ngày công bố dự kiến của luồng sớm, hạn 12 "
                  "tháng kể từ ngày bộc lộ và hạn nộp phí duy trì trước 03 tháng."),
            ("P", "- **Báo cáo định kỳ.** Hằng năm, Phòng Khoa học Công nghệ tính 16 tiêu chí tại Bảng 1.2 và các chỉ số "
                  "tại Bảng 3.4 từ danh mục."),
            ("P", "- **Làm theo giai đoạn.** Giai đoạn đầu dùng bảng tính dùng chung có cấu trúc thống nhất, không cần mua "
                  "phần mềm; khi khối lượng hồ sơ tăng mới xem xét phần mềm quản lý vòng đời tài sản trí tuệ."),
            ("H3", "c) Chủ thể thực hiện"),
            ("P", "Phòng Khoa học Công nghệ chủ trì, quản lý danh mục; Bộ phận Pháp chế cập nhật thông tin đơn, văn bằng; bộ "
                  "phận quản trị thương hiệu cập nhật nhãn hiệu, biểu trưng; Viện Nghiên cứu giáo dục và Chuyển giao tri "
                  "thức cập nhật hợp đồng khai thác; Phòng Tài chính - Kế toán cập nhật nguồn thu và khoản chi trả cho tác "
                  "giả."),
            ("H3", "d) Điều kiện bảo đảm"),
            ("P", "Giai đoạn đầu không phát sinh chi phí đáng kể. Trách nhiệm cập nhật được quy định trong quy chế phối hợp "
                  "tại Giải pháp 2. Danh mục hoàn thành trước ngày 31 tháng 5 năm 2027 để phục vụ kỳ tự đánh giá đầu tiên "
                  "theo Thông tư 83."),
        ]))

        # Nhóm 3
        h5 = self.tim("3.2.4. Giải pháp 4:")
        self.dat_chu(h5, "3.4.1. Giải pháp 5: Hoàn thiện cơ chế tài chính và chuẩn bị thương mại hóa theo giai đoạn")
        self.chen_truoc(h5, self.cac_doan([
            ("H1", "3.4. Nhóm giải pháp thúc đẩy thương mại hóa và nâng cao nhận thức"),
            ("P", "Nhóm này xử lý hạn chế thứ sáu và các nguyên nhân về nguồn lực, động lực, con người tại Mục 2.5.3, gồm "
                  "Giải pháp 5 và Giải pháp 6."),
        ]))
        self.dat_chu(self.tim("3.2.5. Giải pháp 5:"), "3.4.2. Giải pháp 6: Đào tạo và phát triển văn hóa sở hữu trí tuệ")

        # Mục 3.5
        h351 = self.tim("3.2.6. Tổ chức thực hiện các giải pháp")
        self.dat_chu(h351, "3.5.1. Tổ chức thực hiện và điều kiện bảo đảm")
        self.chen_truoc(h351, self.cac_doan([("H1", "3.5. Lộ trình thực thi và điều kiện bảo đảm")]))
        cu = self.tim("3.3. Kế hoạch thí điểm quy trình sàng lọc và nộp đơn sớm")
        cu.addprevious(self.doan("H2", "3.5.2. Kế hoạch thí điểm quy trình sàng lọc và nộp đơn sớm"))
        self.body.remove(cu)
        self.thay("- Nội dung đo:", "và ý kiến của chủ nhiệm đề tài, Hội đồng nghiệm thu.",
                  "; ý kiến của chủ nhiệm đề tài, Hội đồng nghiệm thu; mức hiểu biết về sở hữu trí tuệ của giảng viên tham "
                  "gia, đo bằng phiếu khảo sát ngắn trước và sau thí điểm.")
        self.body.remove(self.tim("3.4. Lộ trình triển khai và bộ chỉ số theo dõi"))
        self.thay("3.4.1. Lộ trình triển khai", "3.4.1.", "3.5.3.")
        self.thay("Năm nhóm giải pháp được triển khai", "Năm nhóm giải pháp", "Sáu giải pháp")
        self.thay("- Đánh giá điều kiện chuyển bước;", "tại Giải pháp 4", "tại Giải pháp 5")
        self.thay("3.4.2. Bộ chỉ số theo dõi", "3.4.2.", "3.5.4.")
        self.thay("Chương 3 đề xuất năm nhóm giải pháp:", "Chương 3 đề xuất năm nhóm giải pháp: hoàn thiện quy chế với lợi "
                  "ích của tác giả theo ba lớp; giao đầu mối và cơ chế phối hợp; quy trình 8 khâu với sàng lọc trước khi "
                  "công bố và luồng nộp đơn sớm; dòng dự toán cho xác lập quyền và khai thác theo giai đoạn; đào tạo và văn "
                  "hóa sở hữu trí tuệ.",
                  "Chương 3 đề xuất sáu giải pháp thuộc ba nhóm: hoàn thiện quy chế với lợi ích của tác giả theo ba lớp và "
                  "kiện toàn đầu mối, cơ chế phối hợp; chuẩn hóa quy trình 8 khâu với sàng lọc trước khi công bố, luồng nộp "
                  "đơn sớm và lập danh mục số tài sản trí tuệ; hoàn thiện cơ chế tài chính, chuẩn bị thương mại hóa và phát "
                  "triển đào tạo, văn hóa sở hữu trí tuệ.")
        self.bang_3_2()
        self.bang_3_3()

    def o_bang(self, cap, so_bang):
        cap_p = self.tim(cap)
        tbl = cap_p.getnext()
        assert tbl.tag == qn("w:tbl"), so_bang
        return tbl

    def dat_o(self, tc, s):
        ps = tc.findall(qn("w:p"))
        self.dat_chu(ps[0], s)
        for p in ps[1:]:
            tc.remove(p)

    def bang_3_2(self):
        tbl = self.o_bang("Bảng 3.2.", "3.2")
        doi = {"Giải pháp 2, 3, 4": "Giải pháp 2, 3, 5", "Giải pháp 1, 4": "Giải pháp 1, 5",
               "Giải pháp 2, 4, 5": "Giải pháp 2, 4, 5, 6"}
        for tc in tbl.iter(qn("w:tc")):
            s = chu(tc).strip()
            if s in doi:
                self.dat_o(tc, doi[s])

    def bang_3_3(self):
        tbl = self.o_bang("Bảng 3.3.", "3.3")
        trs = tbl.findall(qn("w:tr"))
        tc3 = trs[3].findall(qn("w:tc"))
        assert chu(tc3[0]).startswith("Giải pháp 3.")
        moi = bo_id(copy.deepcopy(trs[3]))
        s = chu(tc3[0])
        s = s.replace("Hạn chế thứ nhất, thứ tư, thứ năm, thứ bảy: chuyển hóa thấp, thiếu sàng lọc và dữ liệu còn rời "
                      "rạc.", "Hạn chế thứ nhất, thứ tư, thứ bảy: chuyển hóa thấp, thiếu sàng lọc.")
        s = s.replace("; trạng thái một số hồ sơ chưa được xác minh đầy đủ.", ".")
        self.dat_o(tc3[0], s)
        self.dat_o(tc3[2], "Biểu mẫu hiện có được hoàn thiện; công cụ tra cứu")
        self.dat_o(tc3[4], "Phiếu khai báo, phiếu rà soát; quy trình 8 khâu")
        self.dat_o(tc3[5], "Chỉ số 4 đến 8 tại Bảng 3.4")
        cot = [
            "Giải pháp 4. Hạn chế thứ năm: dữ liệu chưa đáp ứng yêu cầu quản lý và công khai. Bằng chứng: chưa có danh "
            "mục tài sản trí tuệ trong hệ thống thống kê; hai bảng theo dõi chưa thống nhất trạng thái 5 kiểu dáng; 7 trên "
            "16 tiêu chí chưa tính được.",
            "Phòng Khoa học Công nghệ; Bộ phận Pháp chế, bộ phận quản trị thương hiệu, Phòng Tài chính - Kế toán",
            "Bảng tính dùng chung giai đoạn đầu; phần mềm khi đủ khối lượng",
            "Từ quý IV/2026; hoàn thành trước 31/5/2027",
            "Danh mục số có phân quyền; bộ dữ liệu được phép công bố",
            "Chỉ số 11 tại Bảng 3.4; số tiêu chí tính được trong 16 tiêu chí",
        ]
        for tc, s in zip(moi.findall(qn("w:tc")), cot):
            self.dat_o(tc, s)
        trs[3].addnext(moi)
        for tr, cu, m in ((trs[4], "Giải pháp 4.", "Giải pháp 5."), (trs[5], "Giải pháp 5.", "Giải pháp 6.")):
            tc = tr.findall(qn("w:tc"))[0]
            s = chu(tc)
            assert s.startswith(cu)
            self.dat_o(tc, m + s[len(cu):])

    def ket_luan(self):
        self.thay("Về giải pháp, Đề tài đề xuất năm nhóm giải pháp", "năm nhóm giải pháp", "sáu giải pháp thuộc ba nhóm")
        neo = self.tim("Về giải pháp, Đề tài đề xuất")
        self.chen_sau(neo, self.cac_doan([
            ("P", "**Về đóng góp**, đề tài cung cấp khung phân tích và bộ tiêu chí có thể tính từ dữ liệu hành chính của nhà "
                  "trường, cùng căn cứ cụ thể để Nhà trường hoàn thiện quy chế theo luật mới, chuẩn hóa quy trình, số hóa dữ "
                  "liệu và theo dõi bằng bộ chỉ số. Kết quả có thể tham khảo cho các trường đại học tư thục có điều kiện "
                  "tương tự."),
        ]))

    # ---- mục lục và danh mục ----
    def tieu_de(self):
        els = self.ds()
        md = els.index(self.tim("MỞ ĐẦU", kieu="H0"))
        muc, cap = [], []
        for el in els[md:]:
            if el.tag != qn("w:p"):
                continue
            ppr = el.find(qn("w:pPr"))
            if ppr is None:
                continue
            ol = ppr.find(qn("w:outlineLvl"))
            st = ppr.find(qn("w:pStyle"))
            if ol is not None and int(ol.get(qn("w:val"))) <= 2:
                if el.find(qn("w:bookmarkStart")) is None:
                    self.bookmark(el)
                parts = []
                for c in el.iter():
                    if c.tag == qn("w:t"):
                        parts.append(c.text or "")
                    elif c.tag == qn("w:br"):
                        parts.append(" ")
                muc.append((int(ol.get(qn("w:val"))), re.sub(r"\s+", " ", "".join(parts)).strip(),
                            el.find(qn("w:bookmarkStart")).get(qn("w:name"))))
            elif st is not None and st.get(qn("w:val")) in ("Chuthichbang", "Chuthichhinh"):
                if el.find(qn("w:bookmarkStart")) is None:
                    self.bookmark(el)
                cap.append((chu(el).strip(), el.find(qn("w:bookmarkStart")).get(qn("w:name"))))
        return muc, cap

    def khoang_truong(self, lenh):
        """Các đoạn mục của một trường TOC: (đoạn đầu chứa lệnh, các đoạn mục, đoạn kết)."""
        els = self.ds()
        dau = next(el for el in els if any(lenh in (it.text or "") for it in el.iter(qn("w:instrText"))))
        i = els.index(dau)
        j = i + 1
        while not any(fc.get(qn("w:fldCharType")) == "end" and fc.getparent().getparent() is els[j]
                      for fc in els[j].iter(qn("w:fldChar"))) or els[j].find(qn("w:hyperlink")) is not None:
            j += 1
        return els[i:j], els[j]

    @staticmethod
    def muc_toc(mau, chu_muc, ten, trang):
        p = bo_id(copy.deepcopy(mau))
        h = p.find(qn("w:hyperlink"))
        h.set(qn("w:anchor"), ten)
        rs = h.findall(qn("w:r"))
        rs[0].find(qn("w:t")).text = chu_muc
        for t in rs[0].findall(qn("w:t"))[1:]:
            t.text = ""
        for it in h.iter(qn("w:instrText")):
            it.text = f" PAGEREF {ten} \\h "
        ts = [r.find(qn("w:t")) for r in rs[1:] if r.find(qn("w:t")) is not None]
        ts[-1].text = str(trang)
        return p

    def dung_truong(self, lenh, cac_muc, cap_mau=None):
        cu, ket = self.khoang_truong(lenh)
        dau = cu[0]
        mau_cap = {}
        for el in cu[1:]:
            st = el.find(qn("w:pPr")).find(qn("w:pStyle")).get(qn("w:val"))
            mau_cap.setdefault(st, el)
        for el in cu:
            self.body.remove(el)
        for j, (cap, s, ten, trang) in enumerate(cac_muc):
            if j == 0:
                p = self.muc_toc(dau, s, ten, trang)
            else:
                p = self.muc_toc(mau_cap.get(f"TOC{cap + 1}", next(iter(mau_cap.values()))), s, ten, trang)
            ket.addprevious(p)

    def danh_muc(self, trang):
        muc, cap = self.tieu_de()
        self.dung_truong('TOC \\o "1-3"', [(c, s, t, trang.get(t, "")) for c, s, t in muc])
        self.dung_truong("Chu thich bang", [(0, s, t, trang.get(t, "")) for s, t in cap if s.startswith("Bảng")])
        self.dung_truong("Chu thich hinh", [(0, s, t, trang.get(t, "")) for s, t in cap if s.startswith("Hình")])
        dich = [(t, s) for _, s, t in muc] + [("_DAT_LAI_", "")] + [(t, s) for s, t in cap]
        return dich

    def sua_dinh_dang_bang(self):
        """Đoạn trong ô bảng thiếu định dạng (do sửa tay) lấy định dạng của ô bên cạnh cùng hàng."""
        so = 0
        for tbl in self.body.iter(qn("w:tbl")):
            for tr in tbl.findall(qn("w:tr")):
                mau = None
                for tc in tr.findall(qn("w:tc")):
                    for p in tc.findall(qn("w:p")):
                        if p.find(qn("w:pPr")) is not None and p.find(qn("w:r")) is not None:
                            mau = p
                            break
                    if mau is not None:
                        break
                if mau is None:
                    continue
                r_mau = mau.find(qn("w:r")).find(qn("w:rPr"))
                for tc in tr.findall(qn("w:tc")):
                    for p in tc.findall(qn("w:p")):
                        if p.find(qn("w:pPr")) is None:
                            ppr = copy.deepcopy(mau.find(qn("w:pPr")))
                            jc = ppr.find(qn("w:jc"))
                            if jc is not None:
                                jc.set(qn("w:val"), "left")
                            p.insert(0, ppr)
                            for r in p.findall(qn("w:r")):
                                if r.find(qn("w:rPr")) is None and r_mau is not None:
                                    r.insert(0, copy.deepcopy(r_mau))
                            so += 1
        # Bảng 3.3: nới cột đầu vì chứa cả hạn chế và bằng chứng.
        tbl = self.o_bang("Bảng 3.3.", "3.3")
        rong = [2700, 1450, 1250, 1100, 1350, 1281]
        for g, w in zip(tbl.find(qn("w:tblGrid")).findall(qn("w:gridCol")), rong):
            g.set(qn("w:w"), str(w))
        for tr in tbl.findall(qn("w:tr")):
            for tc, w in zip(tr.findall(qn("w:tc")), rong):
                tc.find(qn("w:tcPr")).find(qn("w:tcW")).set(qn("w:w"), str(w))
        return so

    def lam(self):
        self.sua_dinh_dang_bang()
        self.bia_va_cam_doan()
        self.mo_dau()
        self.chuong_2()
        self.chuong_3()
        self.ket_luan()


def main():
    tmp = tempfile.mkdtemp()
    s = Sua()
    s.lam()
    dich = s.danh_muc({})
    nhap = os.path.join(tmp, "nhap.docx")
    s.doc.save(nhap)
    trang, _, _ = tim_trang(ve_pdf(nhap, tmp), dich)
    s = Sua()
    s.lam()
    s.danh_muc(trang)
    s.doc.save(RA)
    _, tong, than = tim_trang(ve_pdf(RA, tmp), dich)
    print(f"Đã lưu {RA}: {tong} trang PDF, phần thân từ Mở đầu {than} trang")


if __name__ == "__main__":
    main()

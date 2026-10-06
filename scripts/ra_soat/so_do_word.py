# -*- coding: utf-8 -*-
"""Sơ đồ dựng bằng hình vẽ gốc của Word (nhóm hình DrawingML: khung chữ, mũi tên).

Sơ đồ chèn vào Word không phải ảnh: người dùng nhấp vào từng khung để sửa chữ, đổi màu,
kéo vị trí. Tọa độ tính bằng xăng-ti-mét, gốc ở góc trên bên trái của sơ đồ.
"""
from xml.sax.saxutils import escape

EMU = 360000
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '
      'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
      'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"')

CHU = "262522"
XANH_DAM, XANH, XANH_NHAT = "1F5FB0", "2A78D6", "DCE8F7"
CAM, CAM_NHAT = "C2501B", "FCE3D7"
NGOC, NGOC_NHAT = "138A60", "D6F2E7"
XAM, XAM_NHAT = "8C8B86", "F2F1ED"


def e(v):
    return int(round(v * EMU))


class SoDo:
    def __init__(self, rong, cao):
        self.rong, self.cao = rong, cao
        self.hinh = []
        self.n = 1

    def _id(self):
        self.n += 1
        return self.n

    @staticmethod
    def _doan(dong, size, dam, mau, can):
        """dong: chuỗi hoặc danh sách (chuỗi, đậm?)"""
        manh = dong if isinstance(dong, list) else [(dong, dam)]
        runs = []
        for text, d in manh:
            runs.append(
                '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
                + ("<w:b/>" if d else "") + f'<w:color w:val="{mau}"/><w:sz w:val="{int(size * 2)}"/>'
                f'<w:szCs w:val="{int(size * 2)}"/></w:rPr><w:t xml:space="preserve">{escape(text)}</w:t></w:r>')
        return (f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="245" w:lineRule="auto"/>'
                f'<w:ind w:firstLine="0"/><w:jc w:val="{can}"/></w:pPr>{"".join(runs)}</w:p>')

    def hop(self, x, y, w, h, dong, nen=XANH_NHAT, vien=XANH, size=10.5, dam_dong_dau=True, mau_chu=CHU,
            can="center", hinh="roundRect", net=1.0, gach=False, le=0.15):
        """Khung chữ. dong: danh sách dòng; dòng đầu in đậm nếu dam_dong_dau."""
        ps = "".join(self._doan(d, size, dam_dong_dau and i == 0, mau_chu, can) for i, d in enumerate(dong))
        ln = (f'<a:ln w="{int(net * 12700)}"><a:solidFill><a:srgbClr val="{vien}"/></a:solidFill>'
              + ('<a:prstDash val="dash"/>' if gach else "") + "</a:ln>") if vien else "<a:ln><a:noFill/></a:ln>"
        fill = f'<a:solidFill><a:srgbClr val="{nen}"/></a:solidFill>' if nen else "<a:noFill/>"
        geom = f'<a:prstGeom prst="{hinh}"><a:avLst>' + (
            '<a:gd name="adj" fmla="val 10000"/>' if hinh == "roundRect" else "") + "</a:avLst></a:prstGeom>"
        i = self._id()
        self.hinh.append(
            f'<wps:wsp><wps:cNvPr id="{i}" name="Khung {i}"/><wps:cNvSpPr/>'
            f'<wps:spPr><a:xfrm><a:off x="{e(x)}" y="{e(y)}"/><a:ext cx="{e(w)}" cy="{e(h)}"/></a:xfrm>'
            f'{geom}{fill}{ln}</wps:spPr>'
            f'<wps:txbx><w:txbxContent>{ps}</w:txbxContent></wps:txbx>'
            f'<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="{e(le)}" tIns="{e(0.08)}" rIns="{e(le)}" '
            f'bIns="{e(0.08)}" anchor="ctr" anchorCtr="0"><a:noAutofit/></wps:bodyPr></wps:wsp>')

    def mui_ten(self, x1, y1, x2, y2, mau=XAM, net=1.25, dau=True, gach=False):
        x, y = min(x1, x2), min(y1, y2)
        w, h = abs(x2 - x1), abs(y2 - y1)
        flip = (' flipH="1"' if x2 < x1 else "") + (' flipV="1"' if y2 < y1 else "")
        i = self._id()
        self.hinh.append(
            f'<wps:wsp><wps:cNvPr id="{i}" name="Mũi tên {i}"/><wps:cNvCnPr/>'
            f'<wps:spPr><a:xfrm{flip}><a:off x="{e(x)}" y="{e(y)}"/><a:ext cx="{e(w)}" cy="{e(h)}"/></a:xfrm>'
            f'<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>'
            f'<a:ln w="{int(net * 12700)}"><a:solidFill><a:srgbClr val="{mau}"/></a:solidFill>'
            + ('<a:prstDash val="dash"/>' if gach else "")
            + ('<a:tailEnd type="triangle" w="med" len="med"/>' if dau else "")
            + '</a:ln></wps:spPr><wps:bodyPr/></wps:wsp>')

    def run_xml(self, doc_id, ten):
        cx, cy = e(self.rong), e(self.cao)
        return (
            f'<w:r {NS}><w:rPr><w:noProof/></w:rPr><mc:AlternateContent><mc:Choice Requires="wpg"><w:drawing>'
            f'<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{cx}" cy="{cy}"/>'
            f'<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="{doc_id}" name="{escape(ten)}"/>'
            f'<wp:cNvGraphicFramePr/><a:graphic><a:graphicData '
            f'uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><wpg:wgp><wpg:cNvGrpSpPr/>'
            f'<wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/><a:chOff x="0" y="0"/>'
            f'<a:chExt cx="{cx}" cy="{cy}"/></a:xfrm></wpg:grpSpPr>{"".join(self.hinh)}</wpg:wgp>'
            f'</a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice></mc:AlternateContent></w:r>')


# ---------------------------------------------------------------------------
# Ba sơ đồ của Chương 2
# ---------------------------------------------------------------------------
def dau_moi():
    """Hình 2.1. Phân công đầu mối quản lý quyền sở hữu trí tuệ (Điều 11, 13 QĐ 217; Điều 35 QĐ 213)."""
    s = SoDo(15.5, 9.2)
    s.hop(3.5, 0.05, 8.5, 1.3, ["Ban Giám hiệu", "Ký đơn đăng ký; quyết định tỷ lệ phân chia lợi ích"],
          nen=XANH_DAM, vien=None, mau_chu="FFFFFF", size=10)
    y2, h2 = 2.0, 2.6
    hop2 = [
        (0.0, ["Phòng Khoa học công nghệ", "Quản lý, giám sát; quy trình, biểu mẫu khai báo; tiếp nhận, nộp đơn; "
                                           "hồ sơ theo dõi; xúc tiến thương mại hóa"]),
        (5.35, ["Bộ phận Pháp chế", "Tư vấn pháp lý; thực hiện thủ tục xác lập quyền; phối hợp tham mưu "
                                    "hợp đồng chuyển giao, li-xăng"]),
        (10.7, ["Phòng Tài chính - Kế toán", "Tham mưu tỷ lệ phân chia lợi ích từ khai thác tài sản trí tuệ"]),
    ]
    for x, dong in hop2:
        s.hop(x, y2, 4.8, h2, dong, size=9.5)
        s.mui_ten(7.75, 1.35, x + 2.4, y2)
    y3, h3 = 5.2, 1.7
    s.hop(0.0, y3, 15.5, h3, [
        "Các viện đào tạo, nghiên cứu",
        "Mỗi viện có Phòng Học vụ và Hợp tác đối ngoại, Phòng Công nghệ, Đổi mới sáng tạo và Khởi nghiệp. "
        "Yêu cầu ghi nhận tài sản trí tuệ mới phát sinh; phối hợp đăng ký bảo hộ và khai thác"],
        nen=NGOC_NHAT, vien=NGOC, size=9.5)
    for x, _ in hop2:
        s.mui_ten(x + 2.4, y2 + h2, x + 2.4, y3, dau=True)
    y4 = 7.5
    s.hop(1.5, y4, 12.5, 1.65, [
        "Tác giả: giảng viên, người học, cộng tác viên",
        "Khai báo kết quả; xin ý kiến Phòng Khoa học công nghệ trước khi bộc lộ công khai; giữ bí mật thông tin"],
        nen=XAM_NHAT, vien=XAM, size=9.5)
    s.mui_ten(7.75, y3 + h3, 7.75, y4)
    return s


def pheu():
    """Hình 2.7. Chuỗi chuyển hóa từ đề tài cấp cơ sở sang đơn đăng ký sở hữu công nghiệp."""
    s = SoDo(15.5, 8.6)
    bac = [
        (10.0, "38 đề tài cấp cơ sở đã nghiệm thu", "100% số đề tài", XANH_NHAT, CHU),
        (8.8, "11 đề tài có sản phẩm có thể bảo hộ", "28,9% số đề tài", "B5D2F4", CHU),
        (7.6, "9 đề tài có sản phẩm sở hữu công nghiệp", "23,7% số đề tài", "6FA8E8", CHU),
        (6.4, "1 đơn đăng ký sáng chế", "Đề tài 09-2025; 2,6% số đề tài", XANH_DAM, "FFFFFF"),
    ]
    ghi = [None, "2 sản phẩm thuộc quyền tác giả, chưa đăng ký",
           "8 đề tài đã nghiệm thu, chưa nộp đơn", None]
    cx, y, h, k = 5.2, 0.15, 1.55, 0.55
    for i, (w, chu, tl, nen, mau) in enumerate(bac):
        s.hop(cx - w / 2, y, w, h, [chu, tl], nen=nen, vien=None, mau_chu=mau, size=10.5,
              hinh="rect")
        if i < len(bac) - 1:
            s.mui_ten(cx, y + h, cx, y + h + k, mau=XAM)
        if ghi[i]:
            s.mui_ten(cx + w / 2, y + h / 2, 10.55, y + h / 2, mau=CAM, gach=True)
            s.hop(10.6, y + 0.15, 4.85, h - 0.3, [ghi[i]], nen=CAM_NHAT, vien=CAM, size=10, dam_dong_dau=False)
        y += h + k
    return s


def nguyen_nhan():
    """Hình 2.9. Nguyên nhân khách quan, chủ quan và các hạn chế."""
    s = SoDo(15.5, 12.4)
    w = 7.5
    s.hop(0.0, 0.0, w, 0.9, ["Nguyên nhân khách quan"], nen=XAM, vien=None, mau_chu="FFFFFF", size=11)
    s.hop(8.0, 0.0, w, 0.9, ["Nguyên nhân chủ quan"], nen=XANH_DAM, vien=None, mau_chu="FFFFFF", size=11)
    kq = ["Thủ tục xác lập quyền đối với sáng chế kéo dài (Điều 110, Điều 119 Luật Sở hữu trí tuệ)",
          "Chi phí bảo hộ phát sinh ở nhiều khâu và trong suốt thời hạn bảo hộ",
          "Thị trường chuyển giao công nghệ, quyền sở hữu trí tuệ còn hạn chế"]
    cq = ["Chưa có quy trình sàng lọc trước công bố",
          "Chưa có đầu mối chuyên trách về sở hữu trí tuệ",
          "Cơ chế hỗ trợ xác lập quyền còn hạn chế",
          "Dữ liệu quản lý phân tán"]
    h, k = 1.3, 0.25
    for i, t in enumerate(kq):
        s.hop(0.0, 1.15 + i * (h + k), w, h, [t], nen=XAM_NHAT, vien=XAM, size=10, dam_dong_dau=False)
    for i, t in enumerate(cq):
        s.hop(8.0, 1.15 + i * (h + k), w, h, [t], nen=XANH_NHAT, vien=XANH, size=10, dam_dong_dau=False)
    day = 1.15 + 4 * (h + k)
    y_hc = day + 0.9
    s.mui_ten(w / 2, 1.15 + 3 * (h + k) - k, w / 2, y_hc, mau=XAM)
    s.mui_ten(8.0 + w / 2, day - k, 8.0 + w / 2, y_hc, mau=XANH)
    s.hop(0.0, y_hc, 15.5, 12.4 - y_hc - 0.05, [
        "Hạn chế trong quản lý quyền sở hữu trí tuệ",
        "1. Kết quả có thể bảo hộ chưa được chuyển thành đơn đăng ký",
        "2. Quy chế nội bộ chưa đồng bộ với khung pháp lý mới",
        "3. Chưa có hoạt động khai thác thương mại phát sinh doanh thu",
        "4. Mô hình quản lý phân tán, chưa có cơ sở dữ liệu theo dõi tài sản"],
        nen=CAM_NHAT, vien=CAM, size=10.5, can="left", le=0.4)
    return s

# -*- coding: utf-8 -*-
"""Vẽ các sơ đồ khái niệm của báo cáo và bài báo (ảnh PNG độ phân giải cao, phông Liberation Serif,
cùng kích thước chữ với Times New Roman). Tọa độ tính bằng xăng-ti-mét trên khổ rộng 15,5 cm.

    python3 scripts/ban_cuoi/so_do.py      # ghi ảnh vào Ban_cuoi/so_do/
"""
import math
import os

from PIL import Image, ImageDraw, ImageFont

GOC = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
THU_MUC = os.path.join(GOC, "Ban_cuoi", "so_do")
PX = 150  # điểm ảnh trên 1 cm, tương đương khoảng 380 dpi
PT = PX * 0.03528
FONT = "/usr/share/fonts/truetype/liberation/LiberationSerif-Regular.ttf"
FONT_DAM = "/usr/share/fonts/truetype/liberation/LiberationSerif-Bold.ttf"
FONT_NGHIENG = "/usr/share/fonts/truetype/liberation/LiberationSerif-Italic.ttf"

CHU = "#262522"
XANH_DAM, XANH, XANH_NHAT = "#1F5FB0", "#2A78D6", "#DCE8F7"
CAM, CAM_NHAT = "#D9571F", "#FCE3D7"
NGOC, NGOC_NHAT = "#138A60", "#D6F2E7"
VANG, VANG_NHAT = "#B87D00", "#FFF0CC"
XAM, XAM_NHAT = "#8C8B86", "#F2F1ED"


def p(v):
    return int(round(v * PX))


def font(size, dam=False, nghieng=False):
    return ImageFont.truetype(FONT_DAM if dam else (FONT_NGHIENG if nghieng else FONT), int(round(size * PT)))


class SoDo:
    def __init__(self, rong, cao):
        self.im = Image.new("RGB", (p(rong), p(cao)), "white")
        self.d = ImageDraw.Draw(self.im)

    # --- chữ ----------------------------------------------------------------
    def _ngat(self, text, f, rong_px):
        dong, cur = [], ""
        for tu in text.split():
            thu = (cur + " " + tu).strip()
            if f.getlength(thu) <= rong_px or not cur:
                cur = thu
            else:
                dong.append(cur)
                cur = tu
        if cur:
            dong.append(cur)
        return dong

    def khoi_chu(self, x, y, w, h, doan, can="center", mau=CHU, dem=0.18, doc="center"):
        """doan: danh sách (chữ, cỡ, đậm, nghiêng). Tự ngắt dòng trong khung (x, y, w, h)."""
        dong = []
        for text, size, dam, nghieng in doan:
            f = font(size, dam, nghieng)
            for d in self._ngat(text, f, p(w - 2 * dem)):
                dong.append((d, f, int(f.size * 1.22)))
        tong = sum(k for _, _, k in dong)
        yy = p(y) + (p(h) - tong) // 2 if doc == "center" else p(y + dem)
        for d, f, k in dong:
            if can == "center":
                xx = p(x) + (p(w) - f.getlength(d)) / 2
            else:
                xx = p(x + dem)
            self.d.text((xx, yy), d, font=f, fill=mau)
            yy += k

    # --- hình ---------------------------------------------------------------
    def hop(self, x, y, w, h, doan, nen=XANH_NHAT, vien=XANH, day=0.035, bo=0.18, can="center", mau_chu=CHU,
            gach=False, doc="center"):
        box = [p(x), p(y), p(x + w), p(y + h)]
        self.d.rounded_rectangle(box, radius=p(bo), fill=nen, outline=None if gach else vien, width=max(1, p(day)))
        if gach:
            self._vien_gach(x, y, w, h, vien, day)
        self.khoi_chu(x, y, w, h, doan, can=can, mau=mau_chu, doc=doc)

    def _vien_gach(self, x, y, w, h, mau, day):
        for a, b in [((x, y), (x + w, y)), ((x + w, y), (x + w, y + h)), ((x + w, y + h), (x, y + h)), ((x, y + h), (x, y))]:
            self._gach(a, b, mau, day)

    def _gach(self, a, b, mau, day, dai=0.18, ho=0.12):
        (x1, y1), (x2, y2) = a, b
        L = math.hypot(x2 - x1, y2 - y1)
        if L == 0:
            return
        ux, uy = (x2 - x1) / L, (y2 - y1) / L
        t = 0.0
        while t < L:
            t2 = min(L, t + dai)
            self.d.line([p(x1 + ux * t), p(y1 + uy * t), p(x1 + ux * t2), p(y1 + uy * t2)], fill=mau, width=max(1, p(day)))
            t += dai + ho

    def mui_ten(self, diem, mau=XAM, day=0.045, mui=0.28, hai_dau=False, gach=False):
        for a, b in zip(diem[:-1], diem[1:]):
            if gach:
                self._gach(a, b, mau, day)
            else:
                self.d.line([p(a[0]), p(a[1]), p(b[0]), p(b[1])], fill=mau, width=max(1, p(day)))
        self._dau(diem[-2], diem[-1], mau, mui)
        if hai_dau:
            self._dau(diem[1], diem[0], mau, mui)

    def _dau(self, a, b, mau, mui):
        (x1, y1), (x2, y2) = a, b
        g = math.atan2(y2 - y1, x2 - x1)
        c = math.radians(24)
        pts = [(x2, y2), (x2 - mui * math.cos(g - c), y2 - mui * math.sin(g - c)),
               (x2 - mui * math.cos(g + c), y2 - mui * math.sin(g + c))]
        self.d.polygon([(p(a_), p(b_)) for a_, b_ in pts], fill=mau)

    def nhan(self, x, y, text, size=9, dam=False, nghieng=True, mau=XAM, can="center"):
        f = font(size, dam, nghieng)
        w = f.getlength(text)
        xx = p(x) - w / 2 if can == "center" else p(x)
        self.d.rectangle([xx - 6, p(y) - 4, xx + w + 6, p(y) + f.size * 1.2], fill="white")
        self.d.text((xx, p(y)), text, font=f, fill=mau)

    def mui_ten_khoi(self, x, y, w, h, doan, nen, mau_chu="white"):
        """Khối hình mũi tên chỉ sang phải (dùng cho lộ trình)."""
        n = h * 0.45
        pts = [(x, y), (x + w - n, y), (x + w, y + h / 2), (x + w - n, y + h), (x, y + h), (x + n * 0.6, y + h / 2)]
        self.d.polygon([(p(a), p(b)) for a, b in pts], fill=nen)
        self.khoi_chu(x + n * 0.6, y, w - n * 1.6, h, doan, mau=mau_chu)

    def luu(self, ten):
        os.makedirs(THU_MUC, exist_ok=True)
        duong = os.path.join(THU_MUC, ten)
        self.im.save(duong, dpi=(round(PX * 2.54), round(PX * 2.54)))
        return duong


def T(text, size=10.5, dam=False, nghieng=False):
    return (text, size, dam, nghieng)


# ---------------------------------------------------------------------------
def chu_trinh():
    s = SoDo(15.5, 11.0)
    s.hop(5.0, 4.15, 5.5, 2.7, [T("BẢO VỆ QUYỀN", 11, True), T("xuyên suốt vòng đời tài sản trí tuệ", 10, nghieng=True)],
          nen=VANG_NHAT, vien=VANG, bo=1.2)
    s.hop(4.25, 0.2, 7.0, 2.6, [T("1. Tạo lập và nhận diện quyền", 11, True),
                                 T("Tiếp nhận ý tưởng, khai báo phát minh nội bộ, thẩm định tính mới trước khi công bố", 10)])
    s.hop(11.0, 3.9, 4.3, 3.2, [T("2. Xác lập quyền", 11, True),
                                 T("Chuẩn bị hồ sơ, bố trí kinh phí, nộp đơn, nhận văn bằng hoặc giấy chứng nhận", 10)])
    s.hop(4.25, 8.2, 7.0, 2.6, [T("3. Khai thác và thương mại hóa", 11, True),
                                 T("Sử dụng trong giảng dạy, chuyển giao quyền sử dụng, góp vốn, thành lập doanh nghiệp", 10)])
    s.hop(0.2, 3.9, 4.3, 3.2, [T("4. Bảo vệ và phân chia lợi ích", 11, True),
                                T("Giám sát xâm phạm, chi trả thù lao, trích quỹ tái đầu tư", 10)])
    m = XANH
    s.mui_ten([(11.25, 1.5), (13.15, 1.5), (13.15, 3.85)], m)
    s.mui_ten([(13.15, 7.1), (13.15, 9.5), (11.3, 9.5)], m)
    s.mui_ten([(4.25, 9.5), (2.35, 9.5), (2.35, 7.15)], m)
    s.mui_ten([(2.35, 3.9), (2.35, 1.5), (4.2, 1.5)], m)
    return s.luu("chu_trinh.png")


def yeu_to():
    s = SoDo(15.5, 8.6)
    s.hop(5.4, 3.2, 4.7, 2.2, [T("HIỆU QUẢ QUẢN LÝ", 11, True), T("QUYỀN SỞ HỮU TRÍ TUỆ", 11, True)],
          nen=XANH, vien=XANH_DAM, mau_chu="white")
    trai = [("1. Thể chế", "Pháp luật quốc gia và quy chế nội bộ"),
            ("2. Tổ chức bộ máy", "Đơn vị chuyên trách, thẩm quyền, quy trình"),
            ("3. Nguồn lực tài chính", "Kinh phí nộp đơn, duy trì hiệu lực, thử nghiệm")]
    phai = [("4. Dữ liệu và công nghệ", "Cơ sở dữ liệu theo dõi vòng đời tài sản"),
            ("5. Con người", "Năng lực cán bộ quản lý và nhà khoa học"),
            ("6. Động lực và văn hóa", "Phân chia lợi ích, tôn vinh, liêm chính học thuật")]
    for i, (a, b) in enumerate(trai):
        y = 0.3 + i * 2.8
        s.hop(0.2, y, 4.5, 2.2, [T(a, 10.5, True), T(b, 10)], nen=XAM_NHAT, vien=XAM)
        s.mui_ten([(4.7, y + 1.1), (5.35, 3.2 + 1.1 + (i - 1) * 0.6)], XAM)
    for i, (a, b) in enumerate(phai):
        y = 0.3 + i * 2.8
        s.hop(10.8, y, 4.5, 2.2, [T(a, 10.5, True), T(b, 10)], nen=XAM_NHAT, vien=XAM)
        s.mui_ten([(10.8, y + 1.1), (10.15, 3.2 + 1.1 + (i - 1) * 0.6)], XAM)
    return s.luu("yeu_to.png")


def khung_phan_tich():
    s = SoDo(15.5, 13.6)
    tang = [
        ("Bối cảnh và căn cứ pháp lý", 0.2, 2.2, XAM_NHAT, XAM,
         [["Luật Sở hữu trí tuệ hợp nhất 2026, Luật Khoa học, công nghệ và đổi mới sáng tạo 2025, Luật Giáo dục đại "
           "học 2025, Chiến lược sở hữu trí tuệ đến năm 2030, Chuẩn cơ sở giáo dục đại học"]]),
        ("Khung lý thuyết, Chương 1", 2.95, 3.3, XANH_NHAT, XANH,
         [["Bốn chức năng quản lý", "Hoạch định, tổ chức, thực hiện, kiểm tra"],
          ["Chu trình bốn giai đoạn", "Tạo lập, xác lập, khai thác, bảo vệ"],
          ["Mười sáu tiêu chí", "Đầu vào, quá trình, đầu ra, kết quả"],
          ["Sáu nhóm yếu tố", "Thể chế, bộ máy, tài chính, dữ liệu, con người, văn hóa"]]),
        ("Thực trạng, Chương 2", 6.8, 3.2, NGOC_NHAT, NGOC,
         [["Nguồn hình thành", "582 bản ghi, 38 đề tài, 12 tài sản trí tuệ"],
          ["Thể chế, tổ chức, nguồn lực", "Quy chế, bộ máy, kinh phí, dữ liệu"],
          ["Chu trình quyền", "Tạo lập, xác lập, bảo vệ, khai thác"],
          ["Đánh giá chung", "Kết quả, hạn chế, nguyên nhân"]]),
        ("Giải pháp, Chương 3", 10.55, 2.85, VANG_NHAT, VANG,
         [["Phân tích SWOT", "và năm nguyên tắc"], ["Năm nhóm giải pháp", "Bốn trụ cột mỗi giải pháp"],
          ["Thí điểm", "Rà soát tại Viện Y - Dược"], ["Lộ trình", "và bộ chỉ số theo dõi"]]),
    ]
    for ten, y, h, nen, vien, o in tang:
        s.hop(0.2, y, 3.3, h, [T(ten, 10.5, True)], nen=vien, vien=vien, mau_chu="white")
        n = len(o)
        w = (11.6 - 0.2 * (n - 1)) / n
        for i, noi in enumerate(o):
            x = 3.7 + i * (w + 0.2)
            doan = [T(noi[0], 10.5 if n > 1 else 10.5, n > 1)] + ([T(noi[1], 9.5)] if len(noi) > 1 else [])
            s.hop(x, y, w, h, doan, nen=nen, vien=vien)
    for y1, y2 in [(2.4, 2.95), (6.25, 6.8), (10.0, 10.55)]:
        s.mui_ten([(9.5, y1), (9.5, y2)], XAM, mui=0.25)
    return s.luu("khung_phan_tich.png")


def bon_dau_moi():
    s = SoDo(15.5, 8.9)
    s.hop(0.2, 0.2, 7.8, 8.5, [T("KHỐI QUẢN TRỊ VÀ DỊCH VỤ", 8, True)], nen=XAM_NHAT, vien=XAM, doc="top")
    s.hop(9.6, 0.2, 5.7, 8.5, [T("KHỐI ĐÀO TẠO VÀ NGHIÊN CỨU", 8, True)], nen="#EEF4FB", vien=XANH, doc="top")
    s.hop(0.5, 1.0, 7.2, 2.3, [T("Phòng Khoa học Công nghệ", 10.5, True),
                               T("Tiếp nhận hồ sơ; nhận diện, theo dõi tài sản trí tuệ theo Điều 11 Quyết định 217", 10)],
          nen="white", vien=XAM)
    s.hop(0.5, 3.55, 7.2, 2.3, [T("Bộ phận Pháp chế, Trung tâm Dịch vụ và Quản trị hành chính tổng hợp", 10.5, True),
                               T("Thủ tục xác lập quyền; đầu mối mạng lưới hỗ trợ công nghệ và đổi mới sáng tạo", 10)],
          nen="white", vien=XAM)
    s.hop(0.5, 6.1, 7.2, 2.3, [T("Trung tâm Tuyển sinh và Quản trị thương hiệu", 10.5, True),
                               T("Xác lập quyền đối với nhãn hiệu, logo, bộ nhận diện", 10)], nen="white", vien=XAM)
    s.hop(9.85, 1.0, 5.2, 3.4, [T("Ba viện đào tạo", 10.5, True),
                                T("Tạo ra kết quả nghiên cứu; nơi phát sinh các sản phẩm đủ điều kiện xác lập quyền", 10)],
          nen="white", vien=XANH)
    s.hop(9.85, 5.0, 5.2, 3.4, [T("Viện Nghiên cứu giáo dục và Chuyển giao tri thức", 10.5, True),
                                T("Tổ chức khoa học và công nghệ; thực hiện hợp đồng khai thác quyền", 10)],
          nen="white", vien=XANH)
    s.mui_ten([(9.85, 2.05), (7.75, 2.05)], XANH)
    s.nhan(8.8, 1.25, "hồ sơ", 8.5)
    s.nhan(8.8, 1.6, "nghiệm thu", 8.5)
    s.mui_ten([(9.85, 4.0), (7.75, 4.7)], CAM, gach=True, hai_dau=True)
    s.nhan(8.8, 4.85, "cần luồng", 8.5, mau=CAM)
    s.nhan(8.8, 5.2, "hồ sơ chung", 8.5, mau=CAM)
    return s.luu("bon_dau_moi.png")


def chuoi_nguyen_nhan():
    s = SoDo(15.5, 7.7)
    o = [(0.2, 0.4, "Chưa có danh mục tài sản trí tuệ trong hệ thống thống kê", XAM_NHAT, XAM, False),
         (5.4, 0.4, "Khoảng trống kỹ thuật: chưa có biểu mẫu rà soát khả năng bảo hộ khi nghiệm thu", CAM_NHAT, CAM, True),
         (10.6, 0.4, "Tác giả tự nhận diện và tự khởi động thủ tục đăng ký", XAM_NHAT, XAM, False),
         (10.6, 4.6, "Chưa có dòng kinh phí nộp đơn; phần thưởng cho văn bằng đến chậm", XAM_NHAT, XAM, False),
         (5.4, 4.6, "Kết quả nghiên cứu được ưu tiên công bố trước", XAM_NHAT, XAM, False),
         (0.2, 4.6, "Quá 12 tháng kể từ ngày bộc lộ, sáng chế và giải pháp hữu ích mất tính mới", VANG_NHAT, VANG, False)]
    for i, (x, y, t, nen, vien, dam) in enumerate(o, 1):
        s.hop(x, y, 4.7, 2.7, [T(f"{i}", 10.5, True), T(t, 10.5, dam)], nen=nen, vien=vien, day=0.05 if dam else 0.035)
    s.nhan(7.75, 3.25, "điểm nghẽn cốt lõi", 9.5, dam=True, mau=CAM)
    s.mui_ten([(4.9, 1.75), (5.35, 1.75)], XAM)
    s.mui_ten([(10.1, 1.75), (10.55, 1.75)], XAM)
    s.mui_ten([(12.95, 3.1), (12.95, 4.55)], XAM)
    s.mui_ten([(10.6, 5.95), (10.15, 5.95)], XAM)
    s.mui_ten([(5.4, 5.95), (4.95, 5.95)], XAM)
    return s.luu("chuoi_nguyen_nhan.png")


def phoi_hop():
    s = SoDo(15.5, 11.5)
    s.hop(4.9, 0.2, 5.7, 3.0, [T("Bộ phận Pháp chế", 10.5, True), T("Đầu mối thể chế và thủ tục", 10, nghieng=True),
                               T("Quy chế, thẩm định hồ sơ, làm việc với Cục Sở hữu trí tuệ, sổ theo dõi văn bằng", 10)])
    s.hop(0.2, 5.1, 5.7, 3.1, [T("Phòng Khoa học Công nghệ", 10.5, True), T("Đầu mối chuyên môn và ngân sách", 10, nghieng=True),
                               T("Rà soát khả năng bảo hộ tại nghiệm thu, kinh phí nghiên cứu, danh mục số tài sản trí tuệ", 10)],
          nen=NGOC_NHAT, vien=NGOC)
    s.hop(9.6, 5.1, 5.7, 3.1, [T("Viện Nghiên cứu giáo dục và Chuyển giao tri thức, các viện đào tạo", 10.5, True),
                               T("Đầu mối tạo sinh tài sản và thương mại hóa", 10, nghieng=True),
                               T("Sáng chế, không gian thử nghiệm, định giá, chuyển giao", 10)], nen=VANG_NHAT, vien=VANG)
    s.hop(6.1, 3.65, 3.3, 2.2, [T("Giao ban liên phòng ban hằng quý", 9.5, True),
                                T("Phó Hiệu trưởng phụ trách khoa học chủ trì", 9.5)], nen="white", vien=XAM, gach=True)
    s.mui_ten([(5.4, 3.2), (3.3, 5.05)], XANH, hai_dau=True)
    s.mui_ten([(10.1, 3.2), (12.2, 5.05)], XANH, hai_dau=True)
    s.mui_ten([(5.95, 6.9), (9.55, 6.9)], XANH, hai_dau=True)
    s.hop(1.3, 9.1, 5.9, 2.2, [T("Phòng Tài chính - Kế toán", 10.5, True), T("Kinh phí nộp đơn, chi trả thù lao, giám sát quỹ", 10)],
          nen=XAM_NHAT, vien=XAM, gach=True)
    s.hop(8.3, 9.1, 5.9, 2.2, [T("Trung tâm Tuyển sinh và Quản trị thương hiệu", 10.5, True),
                               T("Nhãn hiệu, bộ nhận diện, cấp phép sử dụng", 10)], nen=XAM_NHAT, vien=XAM, gach=True)
    s.mui_ten([(4.25, 9.05), (3.05, 8.25)], XAM, gach=True)
    s.mui_ten([(11.25, 9.05), (12.45, 8.25)], XAM, gach=True)
    return s.luu("phoi_hop.png")


def tam_khau():
    s = SoDo(15.5, 9.4)
    k = [("Xác định trước đối tượng quyền từ thuyết minh đề tài", "Chủ nhiệm đề tài, Phòng Khoa học Công nghệ"),
         ("Ươm tạo, thử nghiệm tại Không gian sáng tạo mở thử nghiệm", "Viện Nghiên cứu giáo dục và Chuyển giao tri thức"),
         ("Tra cứu thông tin sáng chế tiền kiểm", "Bộ phận Pháp chế"),
         ("Rà soát bắt buộc khả năng bảo hộ tại nghiệm thu", "Hội đồng nghiệm thu"),
         ("Thẩm định nội bộ trong 15 ngày làm việc", "Bộ phận Pháp chế"),
         ("Soạn và nộp đơn trước hoặc cùng lúc công bố", "Bộ phận Pháp chế"),
         ("Duy trì hiệu lực, cảnh báo trước 3 tháng", "Bộ phận Pháp chế"),
         ("Khai thác thương mại và phân bổ lợi ích", "Viện Nghiên cứu giáo dục và Chuyển giao tri thức")]
    xs = [0.2, 4.05, 7.9, 11.75]
    w, h = 3.55, 3.6
    vt = [(xs[i], 0.3) for i in range(4)] + [(xs[3 - i], 5.5) for i in range(4)]
    for i, ((x, y), (ten, chu_tri)) in enumerate(zip(vt, k), 1):
        noi = i == 4
        s.hop(x, y, w, h, [T(f"Khâu {i}", 10.5, True), T(ten, 10, noi), T(chu_tri, 9, nghieng=True)],
              nen=CAM_NHAT if noi else XANH_NHAT, vien=CAM if noi else XANH, day=0.05 if noi else 0.035)
    for i in range(3):
        s.mui_ten([(xs[i] + w, 2.1), (xs[i + 1] - 0.05, 2.1)], XAM, mui=0.24)
        s.mui_ten([(xs[3 - i], 7.3), (xs[2 - i] + w + 0.05, 7.3)], XAM, mui=0.24)
    s.mui_ten([(xs[3] + w / 2, 3.9), (xs[3] + w / 2, 5.45)], XAM, mui=0.24)
    s.nhan(xs[3] + w / 2 - 1.55, 4.45, "mắt xích quyết định", 9, dam=True, mau=CAM)
    return s.luu("tam_khau.png")


def lo_trinh():
    s = SoDo(15.5, 7.4)
    gd = [("Giai đoạn 1", "Quý IV/2026 - quý II/2027", XANH, XANH_NHAT,
           ["Ban hành quy chế hợp nhất trước 31/5/2027", "Quy trình 8 khâu, biểu mẫu nghiệm thu mới",
            "Thí điểm rà soát tại Viện Y - Dược", "Nhân sự chuyên trách, dòng kinh phí nộp đơn", "Danh mục số tài sản trí tuệ"]),
          ("Giai đoạn 2", "Quý III/2027 - năm 2028", NGOC, NGOC_NHAT,
           ["Đề án doanh nghiệp quản lý tài sản trí tuệ", "Học phần sở hữu trí tuệ từ năm học 2027 - 2028",
            "Chuyên trang dữ liệu số tài sản trí tuệ", "Ít nhất 01 hợp đồng chuyển giao"]),
          ("Giai đoạn 3", "Năm 2029 - 2030", VANG, VANG_NHAT,
           ["Doanh nghiệp quản lý tài sản trí tuệ vận hành", "Định giá, góp vốn bằng tài sản trí tuệ",
            "Thương mại hóa thường xuyên", "Công khai dữ liệu trên Nền tảng số quốc gia"])]
    w = 5.1
    for i, (ten, tg, dam, nhat, ds) in enumerate(gd):
        x = 0.2 + i * (w + 0.05)
        s.mui_ten_khoi(x, 0.2, w + 0.3, 1.5, [T(ten, 10.5, True), T(tg, 9.5)], dam)
        s.hop(x + 0.1, 1.95, w - 0.25, 5.25, [T("- " + d, 10) for d in ds], nen=nhat, vien=dam, can="left", doc="top")
    return s.luu("lo_trinh.png")


def cong_ra_soat():
    """Sơ đồ chung cho bài báo phân tích chính sách: vị trí cổng rà soát tại nghiệm thu."""
    s = SoDo(15.5, 8.6)
    w, h1, h2, y2 = 3.55, 2.9, 3.3, 5.0
    xs = [0.2, 4.05, 7.9, 11.75]
    td, nd, cc = 9.5, 9, 8

    def o(x, y, h, ten, noi, can_cu, **kw):
        s.hop(x, y, w, h, [T(ten, td, True), T(noi, nd)] + ([T(can_cu, cc, nghieng=True)] if can_cu else []), **kw)

    o(xs[0], 0.3, h1, "Nhiệm vụ nghiên cứu", "Tổ chức chủ trì được giao quyền", "khoản 2 Điều 25 Luật số 93/2025/QH15")
    o(xs[1], 0.3, h1, "Sản phẩm nghiệm thu", "Công thức, quy trình, mẫu, dữ liệu, phần mềm", "hồ sơ nghiệm thu")
    o(xs[2], 0.3, h1, "Cổng rà soát khả năng bảo hộ", "Phiếu rà soát bắt buộc tại hội đồng nghiệm thu", "quy chế nội bộ",
      nen=CAM_NHAT, vien=CAM, day=0.05)
    o(xs[3], 0.3, h1, "Không đủ điều kiện bảo hộ", "Công bố, dùng trong đào tạo, đăng ký quyền tác giả", None,
      nen=XAM_NHAT, vien=XAM, gach=True)
    o(xs[2], y2, h2, "Xác lập quyền", "Tra cứu, nộp đơn trong 12 tháng kể từ ngày bộc lộ",
      "khoản 3 Điều 60 Luật Sở hữu trí tuệ; điểm b khoản 2 Điều 66 Luật số 93/2025/QH15", nen=NGOC_NHAT, vien=NGOC)
    o(xs[1], y2, h2, "Văn bằng bảo hộ", "Được tính vào chỉ số sản phẩm khoa học quy đổi",
      "Thông tư số 83/2026/TT-BGDĐT", nen=NGOC_NHAT, vien=NGOC)
    o(xs[0], y2, h2, "Khai thác và chia lợi ích", "Tự quyết thương mại hóa; tác giả tối thiểu 30%",
      "Điều 27, Điều 28 Luật số 93/2025/QH15", nen=NGOC_NHAT, vien=NGOC)
    ym, yd = 0.3 + h1 / 2, y2 + h2 / 2
    s.mui_ten([(xs[0] + w, ym), (xs[1] - 0.05, ym)], XAM, mui=0.22)
    s.mui_ten([(xs[1] + w, ym), (xs[2] - 0.05, ym)], XAM, mui=0.22)
    s.mui_ten([(xs[2] + w, ym), (xs[3] - 0.05, ym)], XAM, mui=0.22, gach=True)
    s.mui_ten([(xs[2] + w / 2, 0.3 + h1), (xs[2] + w / 2, y2 - 0.05)], NGOC, mui=0.24)
    s.mui_ten([(xs[2], yd), (xs[1] + w + 0.05, yd)], NGOC, mui=0.22)
    s.mui_ten([(xs[1], yd), (xs[0] + w + 0.05, yd)], NGOC, mui=0.22)
    s.nhan(xs[2] + w / 2 + 1.1, 3.85, "đủ điều kiện", 9, dam=True, mau=NGOC)
    s.nhan(3.9, 3.85, "thiếu cổng này, sản phẩm đi thẳng sang công bố", 9, mau=CAM)
    return s.luu("cong_ra_soat.png")

TAT_CA = {
    "chu_trinh": chu_trinh, "yeu_to": yeu_to, "khung_phan_tich": khung_phan_tich, "bon_dau_moi": bon_dau_moi,
    "chuoi_nguyen_nhan": chuoi_nguyen_nhan, "phoi_hop": phoi_hop, "tam_khau": tam_khau, "lo_trinh": lo_trinh,
    "cong_ra_soat": cong_ra_soat,
}


def ve_tat_ca():
    return {k: f() for k, f in TAT_CA.items()}


if __name__ == "__main__":
    for k, v in ve_tat_ca().items():
        print(k, v)

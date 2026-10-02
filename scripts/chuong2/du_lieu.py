# -*- coding: utf-8 -*-
"""Dữ liệu đã chuẩn hóa cho Chương 2.

Mọi con số dưới đây đã được đối chiếu với tài liệu gốc trong thư mục
"Tai lieu thanh do" (các danh mục PDF của Phòng Khoa học Công nghệ, danh sách
nhân sự năm 2026, Quy chế chi tiêu nội bộ 2026, Quyết định 213/QĐ-ĐHTĐ) và với
tệp IP_Master_Dataset_ThanhDo_2021_2025_v08.xlsx. Ghi chú nguồn đặt ngay cạnh
từng khối dữ liệu.
"""

NAM = ["2021", "2022", "2023", "2024", "2025"]

# ---------------------------------------------------------------------------
# 1. Nhân lực năm 2026 - nguồn: Tai lieu thanh do/2026 DS.xlsx (252 dòng).
#    Tiến sĩ và tương đương gồm cả Bác sĩ CKII và người có học hàm GS, PGS.
#    Thạc sĩ và tương đương gồm cả Dược sĩ CKI.
# ---------------------------------------------------------------------------
NHAN_LUC = [
    # đơn vị (tên đầy đủ), nhãn ngắn, tổng, GS, PGS, TS, ThS, ĐH, Khác, giảng viên
    ("Viện Y - Dược", "Viện Y - Dược", 70, 0, 8, 42, 26, 2, 0, 64),
    ("Viện Quản trị và Công nghệ", "Viện QT và CN", 47, 2, 7, 29, 17, 1, 0, 42),
    ("Viện Ngôn ngữ - Văn hóa - Quốc tế", "Viện NN - VH - QT", 29, 0, 5, 10, 18, 1, 0, 25),
    ("Trung tâm Tuyển sinh và Quản trị thương hiệu", "TT Tuyển sinh và QTTH", 15, 0, 0, 0, 2, 10, 3, 0),
    ("Trung tâm Dịch vụ và Quản trị hành chính tổng hợp", "TT Dịch vụ và QTHC", 14, 0, 0, 0, 4, 5, 5, 0),
    ("Ban Giáo dục cơ bản", "Ban GDCB", 12, 0, 1, 6, 4, 1, 1, 10),
    ("Viện Nghiên cứu giáo dục và Chuyển giao tri thức", "Viện NCGD và CGTT", 6, 0, 0, 1, 1, 4, 0, 0),
    ("Các đơn vị khác", "Đơn vị khác", 59, 0, 2, 6, 22, 15, 16, 4),
]

# ---------------------------------------------------------------------------
# 2. Sản phẩm khoa học theo năm - nguồn: các danh mục PDF của Phòng KHCN.
#    Đếm theo số thứ tự trong từng khối năm của danh mục.
#    Đề tài cấp cơ sở xếp theo năm trong mã số đề tài (38 mã số, 2021: 6 ...).
#    Tham luận quốc tế xếp theo ngày tổ chức (16 lượt).
#    Đề tài cấp quốc gia xếp theo năm phê duyệt kinh phí.
#    Tham luận hội thảo quốc gia (21) không có trường thời gian.
# ---------------------------------------------------------------------------
SAN_PHAM = [
    ("Bài báo đăng tạp chí trong nước", [17, 42, 50, 72, 109]),
    ("Bài báo đăng tạp chí quốc tế", [6, 13, 11, 33, 52]),
    ("Giáo trình, tài liệu giảng dạy", [26, 31, 12, 7, 11]),
    ("Đề tài khoa học công nghệ cấp cơ sở", [6, 10, 7, 8, 7]),
    ("Sách xuất bản", [0, 2, 2, 2, 6]),
    ("Tham luận hội thảo quốc tế", [1, 2, 3, 5, 5]),
    ("Đề tài khoa học công nghệ cấp quốc gia", [0, 0, 0, 1, 2]),
]
THAM_LUAN_QUOC_GIA = 21
THAM_LUAN_CAP_TRUONG = [49, 50, 45, 60, 74]  # 278 lượt có ngày hợp lệ + 1 lượt ghi năm 1905

# Phân hạng bài báo quốc tế - nguồn: cột ISI/SCOPUS của danh mục PDF.
PHAN_HANG_QT = {
    "Q1": [0, 1, 1, 9, 8],
    "Q2": [0, 1, 0, 9, 16],
    "Q3": [0, 0, 0, 10, 12],
    "Q4": [0, 0, 0, 0, 4],
    "Chưa xếp hạng hoặc tạp chí quốc tế khác": [6, 11, 10, 5, 12],
}

# ---------------------------------------------------------------------------
# 3. Sản lượng theo đơn vị (Bảng 2.3). Đề tài: 37 trên 38 đề tài được gán
#    (Khoa Sau đại học quy về Viện Quản trị và Công nghệ như danh mục giáo
#    trình; 1 đề tài của Phòng KHCN không gán). Bài báo: so khớp họ tên tác
#    giả với danh sách nhân sự (256 trên 405 bài), giữ nguyên từ bản trước.
# ---------------------------------------------------------------------------
DON_VI_SAN_LUONG = [
    # nhãn ngắn, nhân lực, đề tài, giáo trình, bài báo
    ("Viện NN - VH - QT", 29, 12, 21, 27),
    ("Viện QT và CN", 47, 10, 37, 57),
    ("Viện Y - Dược", 70, 9, 23, 55),
    ("Viện NCGD và CGTT", 6, 3, 0, 17),
    ("Ban GDCB", 12, 3, 6, 34),
]
# Người có bài báo - tái lập bằng scripts/ra_soat/so_khop_tac_gia.py (chỉ xuất số tổng hợp). Phân tích thăm dò trên
# danh sách nhân sự năm 2026. Đơn vị đếm là người: 252 người, trong đó 240 người có họ tên duy nhất; 12 người thuộc 6
# nhóm trùng họ tên được đếm riêng, không gán bài báo vì danh mục bài báo không có thông tin định danh để phân biệt
# (đối chiếu số CCCD, ngày sinh cho thấy đây là những người khác nhau; chỉ dùng ở dạng tổng hợp).
# Quy tắc chính A: giữ dấu, đúng thứ tự họ tên.
NHAN_SU_TONG = 252
NHAN_SU_TEN_DUY_NHAT = 240
NHAN_SU_TRUNG_TEN = 12
NHOM_TRUNG_TEN = 6
NHAN_SU_TRUNG_TEN_CO_TRONG_BAI = 8
NHAN_SU_CO_BAI = 83
BAI_KHOP = 256
GINI_BAI = 0.836
TOP10_BAI = 0.383
# Độ nhạy, quy tắc B: bỏ dấu, đúng thứ tự (nhận cả tên tác giả viết không dấu trong bài quốc tế).
DO_NHAY_B = dict(bai_khop=340, co_bai=99, gini=0.827)

# ---------------------------------------------------------------------------
# 4. Đề tài cấp cơ sở - nguồn: "Tổng hợp đề tài KHCN cấp cơ sở của GV
#    2021-2025.pdf" (38 mã số), khớp với sheet 01_PROJECT_ALL của IP dataset.
#    Kinh phí: triệu đồng, chỉ tính khoản Trường cấp bằng tiền.
# ---------------------------------------------------------------------------
DE_TAI = [
    # mã, năm, đơn vị quy đổi, xếp loại, hình thức kinh phí, kinh phí Trường cấp, sản phẩm đủ điều kiện, nhóm quyền, tình trạng
    ("20-2021", 2021, "Viện NN - VH - QT", "Đạt", "Quy đổi giờ", 0, "", "", ""),
    ("22-2021", 2021, "Viện NN - VH - QT", "Đạt", "Quy đổi giờ", 0, "", "", ""),
    ("24-2021", 2021, "Viện NN - VH - QT", "Tốt", "Quy đổi giờ", 0, "", "", ""),
    ("27-2021", 2021, "Viện NN - VH - QT", "Đạt", "Quy đổi giờ", 0, "", "", ""),
    ("28-2021", 2021, "Viện NN - VH - QT", "Đạt", "Quy đổi giờ", 0, "", "", ""),
    ("37-2021", 2021, "Viện QT và CN", "Đạt", "Quy đổi giờ", 0, "", "", ""),
    ("01-2022", 2022, "Phòng KHCN", "Tốt", "Quy đổi giờ", 0, "", "", ""),
    ("02-2022", 2022, "Viện QT và CN", "Đạt", "Quy đổi giờ", 0, "", "", ""),
    ("12-2022", 2022, "Viện NN - VH - QT", "Tốt", "Quy đổi giờ", 0, "", "", ""),
    ("13-2022", 2022, "Viện NN - VH - QT", "Tốt", "Quy đổi giờ", 0, "", "", ""),
    ("14-2022", 2022, "Viện NN - VH - QT", "Tốt", "Quy đổi giờ", 0, "", "", ""),
    ("15-2022", 2022, "Viện NN - VH - QT", "Tốt", "Quy đổi giờ", 0, "", "", ""),
    ("37-2022", 2022, "Viện QT và CN", "Đạt", "Tiền mặt", 5, "Mô hình kiểm tra và sửa chữa hệ thống khởi động điện trên ô tô", "Kiểu dáng công nghiệp hoặc giải pháp hữu ích", "Chưa nộp"),
    ("39-2022", 2022, "Viện Y - Dược", "Đạt", "Tiền mặt", 29, "Sản phẩm cầm máu ở dạng màng", "Giải pháp hữu ích hoặc sáng chế", "Chưa nộp"),
    ("40-2022", 2022, "Viện Y - Dược", "Đạt", "Tiền mặt", 50, "Bộ mẫu cây thuốc", "Sưu tập dữ liệu", "Chưa nộp"),
    ("51-2022", 2022, "Viện QT và CN", "Đạt", "Quy đổi giờ", 0, "", "", ""),
    ("01-2023", 2023, "Ban GDCB", "Khá", "Tiền mặt", 3, "", "", ""),
    ("02-2023", 2023, "Viện NN - VH - QT", "Tốt", "Tiền mặt", 5, "", "", ""),
    ("03-2023", 2023, "Viện NN - VH - QT", "Tốt", "Tiền mặt", 3, "", "", ""),
    ("04-2023", 2023, "Viện NN - VH - QT", "Đạt", "Tiền mặt", 3, "", "", ""),
    ("05-2023", 2023, "Viện QT và CN", "Tốt", "Tiền mặt", 5, "", "", ""),
    ("07-2023", 2023, "Viện QT và CN", "Đạt", "Tiền mặt", 1, "", "", ""),
    ("08-2023", 2023, "Viện NCGD và CGTT", "Tốt", "Tiền mặt", 50, "Tinh dầu Citrus grandis, hồ sơ ghi chuyển giao cho Trường thương mại hóa", "Giải pháp hữu ích", "Chưa nộp đơn cho công thức"),
    ("06-2024", 2024, "Viện Y - Dược", "Tốt", "Tiền mặt", 58, "Bánh xà phòng hữu cơ tảo biển kèm quy trình sản xuất và công thức", "Giải pháp hữu ích", "Chưa nộp"),
    ("07-2024", 2024, "Viện Y - Dược", "Tốt", "Tiền mặt", 43, "100 tiêu bản hiển vi cố định và cẩm nang phát hiện giun sán", "Sưu tập dữ liệu", "Chưa nộp"),
    ("08-2024", 2024, "Viện Y - Dược", "Tốt", "Tiền mặt", 13.25, "Rutin độ tinh khiết 90% kèm quy trình chiết xuất và tinh chế", "Giải pháp hữu ích hoặc sáng chế", "Chưa nộp"),
    ("09-2024", 2024, "Viện Y - Dược", "Đạt", "Tiền mặt", 8.5, "Thang thuốc 100ml kèm công thức và quy trình sắc", "Giải pháp hữu ích", "Chưa nộp"),
    ("10-2024", 2024, "Ban GDCB", "Đạt", "Tiền mặt", 3, "", "", ""),
    ("12-2024", 2024, "Viện NCGD và CGTT", "Đạt", "Tự tìm tài trợ", 0, "", "", ""),
    ("13-2024", 2024, "Ban GDCB", "Đạt", "Tự tìm tài trợ", 0, "", "", ""),
    ("14-2024", 2024, "Viện NCGD và CGTT", "Đạt", "Tự tìm tài trợ", 0, "", "", ""),
    ("01-2025", 2025, "Viện QT và CN", "Tốt", "Tiền mặt", 5, "", "", ""),
    ("02-2025", 2025, "Viện QT và CN", "Tốt", "Quy đổi giờ", 0, "", "", ""),
    ("03-2025", 2025, "Viện QT và CN", "Tốt", "Quy đổi giờ", 0, "", "", ""),
    ("04-2025", 2025, "Viện QT và CN", "Tốt", "Quy đổi giờ", 0, "", "", ""),
    ("06-2025", 2025, "Viện Y - Dược", "Tốt", "Tiền mặt", 30, "Công thức bào chế cồn thuốc xoa bóp, hồ sơ ghi dự kiến đăng ký giải pháp hữu ích", "Giải pháp hữu ích", "Chưa nộp"),
    ("07-2025", 2025, "Viện Y - Dược", "Tốt", "Tiền mặt", 30, "Quy trình bào chế viên ngậm giảm ho, hồ sơ ghi dự kiến đăng ký giải pháp hữu ích", "Giải pháp hữu ích", "Chưa nộp"),
    ("09-2025", 2025, "Viện Y - Dược", "Xuất sắc", "Tiền mặt", 80, "Chiết xuất lá Quế hoa, hồ sơ ghi giai đoạn 2 là sáng chế", "Sáng chế", "Đã nộp đơn trong năm 2025"),
]
TU_TIM_TAI_TRO_CO_KINH_PHI = 19.5
# Ngày nghiệm thu của 11 đề tài có sản phẩm tiềm năng - nguồn: cùng danh mục, cột thời gian nghiệm thu
# (tái lập bằng scripts/ra_soat/de_tai_ngay.py). 31 đề tài mã số 2021 - 2024 đều nghiệm thu trước 31/5/2025.
NGAY_NGHIEM_THU = {
    "37-2022": "09/11/2022", "39-2022": "10/11/2022", "40-2022": "10/11/2022", "08-2023": "13/11/2023",
    "06-2024": "15/11/2024", "07-2024": "15/11/2024", "08-2024": "15/11/2024", "09-2024": "15/11/2024",
    "06-2025": "25/12/2025", "07-2025": "25/12/2025", "09-2025": "25/12/2025",
}  # đề tài 14-2024 tự tìm nguồn 19,5 triệu đồng

# ---------------------------------------------------------------------------
# 5. Tài sản trí tuệ thuộc sở hữu hoặc đồng sở hữu của Trường - nguồn: Tai lieu thanh do/Theo dõi đơn nhãn hiệu
#    và KDCN.xlsx, Sheet1 "Đơn đang theo dõi" và Sheet2 "Thống kê văn bằng SHTT". Loại khỏi danh mục tài sản
#    của pháp nhân UNIGO và của cá nhân. Bốn nhóm trạng thái thống nhất:
#    "Đã nộp đơn", "Đã chấp nhận đơn hợp lệ", "Đã cấp văn bằng, giấy chứng nhận", "Chưa xác minh".
#    Năm kiểu dáng: Sheet2 ghi số hiệu 3-00396xx-000, năm 2024; Sheet1 ghi trạng thái "Chờ cấp bằng"
#    -> xếp nhóm "Chưa xác minh" cho đến khi đối chiếu được văn bằng.
#    Sheet1 có một dòng sáng chế "Chấp nhận đơn hợp lệ" không ghi số đơn -> không gán được cho đơn nào.
# ---------------------------------------------------------------------------
DA_CAP, DA_NOP, CHUA_XM = "Đã cấp văn bằng, giấy chứng nhận", "Đã nộp đơn", "Chưa xác minh"
TSTT = [
    # loại hình, tên, số đơn hoặc số văn bằng, năm, trạng thái, sở hữu, nguồn hình thành, trong kỳ 2021 - 2025
    ("Quyền tác giả", "Bộ Logo Trường Đại học Thành Đô", "Giấy chứng nhận 6534/2021/QTG", 2021, DA_CAP,
     "Trường", "Thương hiệu", True),
    ("Quyền tác giả", "Bộ Logo Hệ sinh thái giáo dục Trường Đại học Thành Đô", "Giấy chứng nhận 5790/2023/QTG", 2023,
     DA_CAP, "Trường", "Thương hiệu", True),
    ("Nhãn hiệu", "Thanh do University", "Đơn 4-2021-34771; văn bằng 468353", 2023, DA_CAP, "Trường",
     "Thương hiệu", True),
    ("Nhãn hiệu", "Thado Edupark", "Đơn 4-2023-22141; văn bằng 545780", 2025, DA_CAP, "Trường", "Thương hiệu",
     True),
    ("Nhãn hiệu", "Double2n", "Đơn 4-2023-42540", 2023, DA_NOP, "Đồng sở hữu với doanh nghiệp",
     "Hợp tác doanh nghiệp", True),
    ("Kiểu dáng công nghiệp", "Dung dịch vệ sinh phụ nữ", "Đơn 3-2023-02841; số hiệu 3-0039693-000", 2024, CHUA_XM,
     "Đồng sở hữu với doanh nghiệp", "Hợp tác doanh nghiệp", True),
    ("Kiểu dáng công nghiệp", "Tinh dầu vỏ bưởi đào", "Đơn 3-2023-02842; số hiệu 3-0039694-000", 2024, CHUA_XM,
     "Đồng sở hữu với doanh nghiệp", "Hợp tác doanh nghiệp", True),
    ("Kiểu dáng công nghiệp", "Xịt miệng - họng thảo dược", "Đơn 3-2023-02843; số hiệu 3-0039902-000", 2024, CHUA_XM,
     "Đồng sở hữu với doanh nghiệp", "Hợp tác doanh nghiệp", True),
    ("Kiểu dáng công nghiệp", "Dầu gội nha đam - bưởi đào", "Đơn 3-2023-02844; số hiệu 3-0039695-000", 2024, CHUA_XM,
     "Đồng sở hữu với doanh nghiệp", "Hợp tác doanh nghiệp", True),
    ("Kiểu dáng công nghiệp", "Kem ủ xả tóc chiết xuất từ Bưởi - Bơ - Dừa", "Đơn 3-2023-02845; số hiệu 3-0039696-000",
     2024, CHUA_XM, "Đồng sở hữu với doanh nghiệp", "Hợp tác doanh nghiệp", True),
    ("Sáng chế", "Hợp chất furofuran lignan glucoside và phương pháp chiết từ lá cây quế hoa", "Đơn 1-2025-07378",
     2025, DA_NOP, "Trường", "Nghiên cứu", True),
    ("Sáng chế", "Phương pháp chiết tách và phân lập hợp chất Isoembigenin từ cây Piper aduncum L",
     "Đơn 1-2026-07185", 2026, DA_NOP, "Trường", "Nghiên cứu", False),
]
TSTT_KY = [t for t in TSTT if t[7]]

# ---------------------------------------------------------------------------
# 6. Cơ chế khuyến khích - nguồn: Quy chế chi tiêu nội bộ ban hành 01/8/2026,
#    Bảng 6 (giờ NCKH quy đổi) và Bảng 7 (thưởng bài báo quốc tế).
# ---------------------------------------------------------------------------
KHUYEN_KHICH = [
    # loại sản phẩm, giờ NCKH, thưởng tiền mặt (triệu đồng)
    ("Sáng chế chuẩn quốc tế", 600, 0),
    ("Sáng chế chuẩn Việt Nam", 360, 0),
    ("Giải pháp hữu ích", 180, 0),
    ("Bài WoS Q1", 300, 20),
    ("Bài WoS Q2", 270, 15),
    ("Bài WoS Q3", 240, 12),
    ("Bài WoS Q4", 240, 10),
    ("Bài trong nước 1,75 điểm", 150, 0),
]

# ---------------------------------------------------------------------------
# 7. Ba kênh tài trợ (triệu đồng).
# ---------------------------------------------------------------------------
KENH_TAI_TRO = [
    ("Đề tài cấp cơ sở, 2021 - 2025", 424.75),
    ("Đề tài cấp quốc gia, giao 2024 - 2025", 4670),
    ("Quỹ Học bổng sau tiến sĩ Ngô Xuân Độ, 2025 - 2029", 5000),
]

# ---------------------------------------------------------------------------
# 8. Mức độ tính được bộ tiêu chí tại Mục 1.4.2 (đánh giá của nhóm nghiên cứu
#    dựa trên dữ liệu hiện có). 16 tiêu chí.
# ---------------------------------------------------------------------------
TIEU_CHI = [
    # nhóm, tiêu chí, mức (T: tính được, M: một phần, C: chưa tính được), căn cứ
    ("Đầu vào", "Kinh phí nộp đơn, duy trì văn bằng, khen thưởng sáng tạo", "M", "Điều 35, Điều 38 Quyết định 213 cho phép nộp lệ phí, chi thuê ngoài; chưa có dự toán và số liệu chi riêng"),
    ("Đầu vào", "Nhân lực quản lý sở hữu trí tuệ", "T", "Danh sách nhân sự năm 2026"),
    ("Đầu vào", "Cơ sở dữ liệu tra cứu, hạ tầng kỹ thuật", "C", "Chưa có số liệu về công cụ tra cứu sáng chế được sử dụng"),
    ("Quá trình", "Mức độ tương thích của quy chế nội bộ với pháp luật", "T", "Đối chiếu văn bản tại Mục 2.2.1"),
    ("Quá trình", "Thời gian xử lý một hồ sơ đề xuất bảo hộ", "C", "Không lưu ngày tiếp nhận, ngày nộp đơn"),
    ("Quá trình", "Mức độ chuẩn hóa biểu mẫu và quy trình phối hợp", "M", "Biểu mẫu đề xuất, thuyết minh, hợp đồng, nghiệm thu đã có mục sở hữu trí tuệ; chưa có nội dung sàng lọc khả năng bảo hộ và luồng hồ sơ liên thông"),
    ("Quá trình", "Số lớp tập huấn sở hữu trí tuệ", "C", "Kế hoạch 07/KH-ĐHTĐ chỉ thống kê lượt tập huấn chung, không tách riêng sở hữu trí tuệ"),
    ("Đầu ra", "Số đơn đăng ký sở hữu công nghiệp", "T", "Danh mục theo dõi đơn"),
    ("Đầu ra", "Số văn bằng bảo hộ được cấp", "M", "Trạng thái 5 kiểu dáng chưa thống nhất giữa hai bảng theo dõi"),
    ("Đầu ra", "Số giấy chứng nhận quyền tác giả cho giáo trình, phần mềm", "T", "Danh mục theo dõi đơn"),
    ("Đầu ra", "Số công trình có sản phẩm chuyển hóa được thành tài sản trí tuệ", "M", "Suy ra từ mô tả sản phẩm nghiệm thu, chưa có phiếu sàng lọc"),
    ("Kết quả", "Số hợp đồng chuyển giao, cấp phép", "M", "Chỉ có số liệu của một đơn vị"),
    ("Kết quả", "Nguồn thu từ khai thác tài sản trí tuệ", "C", "Phí bản quyền theo doanh số không được theo dõi"),
    ("Kết quả", "Tỷ trọng nguồn thu khoa học công nghệ trên tổng thu", "C", "Không có số liệu tài chính tách riêng"),
    ("Kết quả", "Tác động đến kiểm định và xếp hạng", "C", "Chưa có đối sánh chính thức"),
    ("Kết quả", "Mức độ hài lòng và động lực của giảng viên", "C", "Chưa có khảo sát"),
]

# ---------------------------------------------------------------------------
# 9. Kế hoạch và thực hiện 2024 - 2025 - nguồn chỉ tiêu: Kế hoạch số 07/KH-ĐHTĐ
#    ngày 01/7/2024, mục 2.2.4 (bản quét, đã nhận dạng chữ và đối chiếu ảnh gốc).
#    Thực hiện: các danh mục PDF của Phòng KHCN và danh mục tài sản trí tuệ.
#    Tham luận hội thảo quốc gia không đưa vào vì danh mục không có năm.
# ---------------------------------------------------------------------------
KE_HOACH = [
    # mục, chỉ tiêu theo nguyên văn Kế hoạch, nhóm, kế hoạch 2024, kế hoạch 2025, thực hiện 2024, thực hiện 2025
    # Thực hiện None: chưa xác định được (xem ghi chú KH_1_11).
    ("1.1", "Bài báo tạp chí khoa học quốc tế", "Công bố", 15, 15, 33, 52),
    ("3.6", "Sách xuất bản có chỉ số ISBN", "Công bố", 1, 2, 2, 6),
    ("1.6", "Bài báo đăng kỷ yếu Hội thảo cấp Trường", "Công bố", 30, 30, 60, 74),
    ("1.3, 1.4", "Bài báo đăng tạp chí NCKH và PT của trường và tạp chí trong nước", "Công bố", 55, 55, 72, 109),
    ("1.9", "Đề tài KHCN cấp Bộ/Nhà nước", "Đề tài và học liệu", 1, 1, 1, 2),
    ("1.10", "Đề tài KHCN cấp cơ sở", "Đề tài và học liệu", 7, 7, 8, 7),
    ("1.2", "Bài báo Hội thảo khoa học quốc tế", "Công bố", 5, 6, 5, 5),
    ("3.4", "Giáo trình/TLTK LHNB được nghiệm thu", "Đề tài và học liệu", 10, 10, 7, 11),
    ("1.12", "Chuyển giao công nghệ", "Tài sản trí tuệ", 0, 1, 0, 0),
    ("1.11", "Công nhận sáng chế, phát minh, KDCN, bản quyền tác giả", "Tài sản trí tuệ", 2, 2, None, None),
]
# Chỉ tiêu 1.11 không gồm nhãn hiệu. Trong phạm vi chỉ tiêu, năm 2024 - 2025 chỉ có 5 kiểu dáng ghi năm 2024 tại
# Sheet2 nhưng trạng thái chưa thống nhất; Logo Thần đồng năm 2025 thuộc Trường UNIGO. Nếu 5 kiểu dáng được xác nhận
# cấp năm 2024 thì kết quả là 5 trên 4, tức 125%; bản trước tính 6 trên 4 do gộp cả nhãn hiệu Thado Edupark.
KH_1_11 = dict(neu_xac_nhan=5, ke_hoach=4)

# ---------------------------------------------------------------------------
# 10. Quy định về lợi ích của tác giả và phân chia nguồn thu (Bảng 2.3) - nguồn: Quyết định 213 (Điều 36),
#     Quy chế quản trị tài sản trí tuệ kèm Quyết định 217 (Điều 13), Quy chế chi tiêu nội bộ ban hành 01/8/2026
#     (Bảng 7 mục 3), Luật số 93/2025/QH15 (Điều 28, Điều 73), Văn bản hợp nhất 67/VBHN-VPQH (Điều 135).
#     Điều lệ Quỹ Ngô Xuân Độ chưa có văn bản gốc trong hồ sơ: ghi theo thông tin được cung cấp.
# ---------------------------------------------------------------------------
QUY_DINH_LOI_ICH = [
    # văn bản, điều khoản | phạm vi và nguồn kinh phí | đối tượng hưởng và tỷ lệ | cơ sở tính | điều kiện, mức trần
    ("Quyết định 213, điểm a khoản 4 Điều 36, năm 2021",
     "Sản phẩm đề tài sử dụng ngân sách nhà nước do Trường chủ trì, đã nghiệm thu, được thương mại hóa",
     "40% nộp ngân sách nhà nước; 30% Trường; 30% khen thưởng tập thể tác giả",
     "Nguồn thu sau khi trừ các khoản chi phí cần thiết, hợp lệ",
     "Khen thưởng tối đa 100 triệu đồng một đề tài; phần vượt chuyển vào quỹ khen thưởng, phúc lợi"),
    ("Quyết định 213, điểm b khoản 4 Điều 36, năm 2021",
     "Tài sản trí tuệ thuộc sở hữu của Trường: sáng chế, giải pháp hữu ích, quy trình công nghệ, giải pháp kỹ thuật",
     "Tác giả 30%, ghi là theo Luật Sở hữu trí tuệ; khoa, đơn vị có tác giả 20%; Quỹ nghiên cứu khoa học 50%",
     "Kinh phí chuyển giao sau khi trừ các khoản chi phí cần thiết, hợp lệ",
     "Không đặt trần"),
    ("Quyết định 217, Điều 13, năm 2024",
     "Tài sản trí tuệ của Trường; khoản 1 dẫn theo nguồn kinh phí hoặc thỏa thuận; khoản 2 nhuận bút sách, giáo trình "
     "theo Quy chế chi tiêu nội bộ",
     "Khoản 3: Hiệu trưởng quyết định tỷ lệ sau khi có tham mưu của Phòng Khoa học Công nghệ, Pháp chế, Tài chính - "
     "Kế toán",
     "Khoản thu sau khi trừ thuế, phí, lệ phí xác lập quyền, phần trích nộp cơ quan cấp kinh phí, quỹ phát triển khoa "
     "học công nghệ",
     "Khoản 3 chỉ áp dụng khi không có thỏa thuận, Quy chế chi tiêu nội bộ hoặc pháp luật không quy định"),
    ("Quy chế chi tiêu nội bộ, Bảng 7 mục 3, năm 2026",
     "Đề tài cấp cơ sở có đăng ký sở hữu trí tuệ",
     "Trích 50% kinh phí chuyển giao công nghệ về Nhà trường, nếu có; không nêu phần của tác giả",
     "Kinh phí chuyển giao công nghệ",
     "Áp dụng từ ngày 01 tháng 8 năm 2026; kinh phí đề tài không quá 50 triệu đồng; nghiệm thu phải có chứng nhận đăng "
     "ký sở hữu trí tuệ"),
    ("Điều lệ Quỹ Học bổng sau tiến sĩ Ngô Xuân Độ, Điều 9, năm 2025",
     "Sản phẩm hình thành từ kinh phí của Quỹ",
     "Năm đầu 50% tác giả, 50% Trường; từ năm thứ hai 20% tác giả, 80% Trường",
     "Chưa đối chiếu được văn bản gốc",
     "Ghi theo thông tin được cung cấp, cần bổ sung văn bản"),
    ("Luật số 93/2025/QH15, khoản 2 Điều 28",
     "Phần kết quả không sử dụng ngân sách nhà nước",
     "Chủ sở hữu tự quyết định, kể cả thưởng cho tác giả",
     "Lợi nhuận thu được tương ứng phần kết quả",
     "Không quy định mức"),
    ("Luật số 93/2025/QH15, khoản 3, khoản 4 Điều 28",
     "Phần kết quả nhiệm vụ sử dụng ngân sách nhà nước",
     "Thưởng tác giả tối thiểu 30%; thưởng người tổ chức thương mại hóa; tái đầu tư; mục đích khác",
     "Lợi nhuận sau thuế; hoặc giá trị kết quả khi góp vốn, liên doanh, thành lập doanh nghiệp",
     "Nhiệm vụ giao từ ngày 01 tháng 10 năm 2025; nhiệm vụ giao từ ngày 01 tháng 01 năm 2023 nếu là sáng chế, kiểu dáng "
     "công nghiệp, thiết kế bố trí đã được cấp văn bằng, theo khoản 7 Điều 73"),
    ("Luật Sở hữu trí tuệ, khoản 1 Điều 135",
     "Sáng chế, kiểu dáng công nghiệp, thiết kế bố trí, không phân biệt nguồn kinh phí",
     "Thù lao cho tác giả theo thỏa thuận; mặc định 10% hoặc 15%",
     "10% lợi nhuận trước thuế khi tự sử dụng; 15% tổng số tiền nhận được mỗi lần trước thuế khi chuyển giao quyền sử "
     "dụng",
     "Mức mặc định chỉ áp dụng khi không có thỏa thuận; trả trong suốt thời hạn bảo hộ; không đặt trần"),
]

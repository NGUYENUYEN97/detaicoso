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
BAI_BAO_GAN_DE_TAI = 107  # kết quả so khớp chủ nhiệm đề tài - tác giả bài báo (giữ nguyên từ bản trước)

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
NHAN_SU_CO_TEN = 247
NHAN_SU_CO_BAI = 88

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
TU_TIM_TAI_TRO_CO_KINH_PHI = 19.5  # đề tài 14-2024 tự tìm nguồn 19,5 triệu đồng

# ---------------------------------------------------------------------------
# 5. Tài sản trí tuệ thuộc sở hữu của Trường - nguồn: sheet 13B_PROTECTION_FULL.
#    Loại khỏi danh mục các tài sản của pháp nhân UNIGO và của cá nhân.
# ---------------------------------------------------------------------------
TSTT = [
    # loại hình, tên, năm, trạng thái, sở hữu, nguồn hình thành
    ("Quyền tác giả", "Bộ Logo Trường Đại học Thành Đô", 2021, "Đã cấp văn bằng", "Trường đơn sở hữu", "Thương hiệu"),
    ("Quyền tác giả", "Bộ Logo Hệ sinh thái giáo dục", 2023, "Đã cấp văn bằng", "Trường đơn sở hữu", "Thương hiệu"),
    ("Nhãn hiệu", "Thanh do University", 2023, "Đã cấp văn bằng", "Trường đơn sở hữu", "Thương hiệu"),
    ("Nhãn hiệu", "Thado Edupark", 2025, "Đã cấp văn bằng", "Trường đơn sở hữu", "Thương hiệu"),
    ("Nhãn hiệu", "Double2n", 2023, "Đang xử lý", "Đồng sở hữu với doanh nghiệp", "Hợp tác doanh nghiệp"),
    ("Kiểu dáng công nghiệp", "Dung dịch vệ sinh phụ nữ", 2024, "Đã cấp văn bằng", "Đồng sở hữu với doanh nghiệp", "Hợp tác doanh nghiệp"),
    ("Kiểu dáng công nghiệp", "Tinh dầu vỏ bưởi đào", 2024, "Đã cấp văn bằng", "Đồng sở hữu với doanh nghiệp", "Hợp tác doanh nghiệp"),
    ("Kiểu dáng công nghiệp", "Xịt miệng - họng thảo dược", 2024, "Đã cấp văn bằng", "Đồng sở hữu với doanh nghiệp", "Hợp tác doanh nghiệp"),
    ("Kiểu dáng công nghiệp", "Dầu gội nha đam - bưởi đào", 2024, "Đã cấp văn bằng", "Đồng sở hữu với doanh nghiệp", "Hợp tác doanh nghiệp"),
    ("Kiểu dáng công nghiệp", "Kem ủ xả tóc từ Bưởi, Bơ, Dừa", 2024, "Đã cấp văn bằng", "Đồng sở hữu với doanh nghiệp", "Hợp tác doanh nghiệp"),
    ("Sáng chế", "Hợp chất furofuran lignan glucoside từ lá cây quế hoa", 2025, "Đang xử lý", "Trường đơn sở hữu", "Nghiên cứu"),
    ("Sáng chế", "Phương pháp chiết tách hợp chất Isoembigenin từ cây Piper aduncum L", 2026, "Đang xử lý", "Trường đơn sở hữu", "Nghiên cứu"),
]

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
    ("Đầu vào", "Kinh phí nộp đơn, duy trì văn bằng, khen thưởng sáng tạo", "C", "Không có dòng chi riêng cho sở hữu trí tuệ"),
    ("Đầu vào", "Nhân lực quản lý sở hữu trí tuệ", "T", "Danh sách nhân sự năm 2026"),
    ("Đầu vào", "Cơ sở dữ liệu tra cứu, hạ tầng kỹ thuật", "M", "Có quyền truy cập qua Mạng lưới Trung tâm Hỗ trợ công nghệ và đổi mới sáng tạo, chưa có số liệu sử dụng"),
    ("Quá trình", "Mức độ tương thích của quy chế nội bộ với pháp luật", "T", "Đối chiếu văn bản tại Mục 2.2.1"),
    ("Quá trình", "Thời gian xử lý một hồ sơ đề xuất bảo hộ", "C", "Không lưu ngày tiếp nhận, ngày nộp đơn"),
    ("Quá trình", "Mức độ chuẩn hóa biểu mẫu và quy trình phối hợp", "M", "Có quy trình tại Điều 35, chưa có biểu mẫu nhận diện"),
    ("Quá trình", "Số lớp tập huấn sở hữu trí tuệ", "C", "Không có thống kê"),
    ("Đầu ra", "Số đơn đăng ký sở hữu công nghiệp", "T", "Danh mục theo dõi đơn"),
    ("Đầu ra", "Số văn bằng bảo hộ được cấp", "T", "Danh mục theo dõi đơn"),
    ("Đầu ra", "Số giấy chứng nhận quyền tác giả cho giáo trình, phần mềm", "T", "Danh mục theo dõi đơn"),
    ("Đầu ra", "Số công trình có sản phẩm chuyển hóa được thành tài sản trí tuệ", "M", "Suy ra từ mô tả sản phẩm nghiệm thu, chưa có phiếu nhận diện"),
    ("Kết quả", "Số hợp đồng chuyển giao, cấp phép", "M", "Chỉ có số liệu của một đơn vị"),
    ("Kết quả", "Nguồn thu từ khai thác tài sản trí tuệ", "C", "Phí bản quyền theo doanh số không được theo dõi"),
    ("Kết quả", "Tỷ trọng nguồn thu khoa học công nghệ trên tổng thu", "C", "Không có số liệu tài chính tách riêng"),
    ("Kết quả", "Tác động đến kiểm định và xếp hạng", "C", "Chưa có đối sánh chính thức"),
    ("Kết quả", "Mức độ hài lòng và động lực của giảng viên", "C", "Chưa có khảo sát"),
]

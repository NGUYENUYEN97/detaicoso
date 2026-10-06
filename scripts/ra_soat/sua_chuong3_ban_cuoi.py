# -*- coding: utf-8 -*-
"""Sửa Chương 3 bản cuối theo góp ý, nối tiếp các sửa của Chương 1, Chương 2 (một tệp, theo dõi thay đổi).

Đầu vào : Bao_cao_toan_van_ban_cuoi_sua_Chuong1.docx (như sua_chuong2_ban_cuoi.py)
Đầu ra  : Bao_cao_toan_van_ban_cuoi_sua_Chuong1_2_3.docx

Góp ý: rút tên chương; 3.1.1 rút khoảng 40%; SWOT nêu điểm yếu, không đưa tỷ lệ chi tiết; 3.2.1 tra cứu sơ bộ
có hỗ trợ; Tổ tư vấn thay Hội đồng; doanh nghiệp khởi nguồn chỉ nghiên cứu khi đủ điều kiện; Bảng 3.3 rút
còn ý chính; KPI ghi "phấn đấu"; bớt cụm từ khẩu hiệu.
Độ chính xác (đối chiếu VBPL và quy chế của Trường): khoản 2 Điều 25, điểm a khoản 3 Điều 28, điểm b khoản 2
Điều 66 Luật 93; khoản 1 Điều 28 Luật 125; Điều 9a Nghị định 65 (bổ sung bởi Nghị định 100); Tiêu chí 1.1,
6.2 Thông tư 83; Điều 10, 11, 13 Quyết định 217; Quyết định 1624 (thí điểm hỗ trợ định giá); bỏ Viện REK,
Nghị quyết 45 (không có trong danh mục tài liệu); Phòng Khoa học công nghệ; Viện Y - Dược 9 trên 11 đề tài.
"""
import os
import sys

DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, DIR)
import sua_chuong2_ban_cuoi as S2  # noqa: E402

S2.RA = os.path.join(S2.THU_MUC, "Bao_cao_toan_van_ban_cuoi_sua_Chuong1_2_3.docx")

# Các sửa từng phần trong Chương 3 của script Chương 2 nay được thay bằng viết lại cả ô, cả bảng
S2.MOT_PHAN = [m for m in S2.MOT_PHAN
               if not m[0].startswith(("W1.", "S2.", "Xử lý Hạn chế"))]

S2.THAY_TAT_CA += [
    ("CHƯƠNG 3. HỆ THỐNG GIẢI PHÁP", "CHƯƠNG 3. GIẢI PHÁP"),
    (" ĐÁP ỨNG KHUNG PHÁP LÝ MỚI", ""),
    ("Định hướng chiến lược và yêu cầu mới từ khung pháp lý giai đoạn 2025 - 2026", "Căn cứ xây dựng hệ thống giải pháp"),
]
# Bỏ Bảng 3.2 (lặp với Mục 3.2): Bảng 3.3, 3.4 thành 3.2, 3.3
S2.DOI_SO.update({"Bảng 3.3": "Bảng 3.2", "Bảng 3.4": "Bảng 3.3"})
S2.BANG_XOA.append("Nhóm phương án")

TPL = "Phòng Khoa học công nghệ"
S2.DOAN_NGOAI += [
    ("Chương 3: Hệ thống giải pháp nâng cao hiệu quả",
     "Chương 3: Giải pháp nâng cao hiệu quả quản lý quyền sở hữu trí tuệ tại Trường Đại học Thành Đô."),
    ("Từ những kết quả đạt được, các điểm nghẽn thể chế",
     "Trên cơ sở các hạn chế và nguyên nhân đã xác định ở Chương 2, chương này đề xuất các giải pháp nâng cao hiệu quả "
     "quản lý quyền sở hữu trí tuệ tại Trường Đại học Thành Đô giai đoạn 2026 - 2030, gắn với các quy định mới có hiệu "
     "lực từ năm 2025, 2026."),
    # 3.1.1: rút gọn, dẫn đúng điều khoản
    ("Việc xây dựng hệ thống giải pháp quản lý quyền sở hữu trí tuệ tại Trường",
     "Các giải pháp được xây dựng trên ba căn cứ: chính sách, pháp lý và thực tiễn rút ra từ Chương 2."),
    ("Về chủ trương của Đảng và Nhà nước: Nghị quyết số 45-NQ/TW",
     "Về căn cứ chính sách, Chiến lược sở hữu trí tuệ đến năm 2030 (Quyết định số 1068/QĐ-TTg, được sửa đổi, bổ sung "
     "bởi Quyết định số 1624/QĐ-TTg), Chiến lược phát triển khoa học, công nghệ và đổi mới sáng tạo đến năm 2030 (Quyết "
     "định số 569/QĐ-TTg) và Kết luận số 51-KL/TW đều yêu cầu đẩy mạnh khai thác, thương mại hóa tài sản trí tuệ hình "
     "thành từ nghiên cứu."),
    ("Về khung pháp lý chuyên ngành giai đoạn 2025 - 2026:",
     "Về căn cứ pháp lý, các quy định mới đã phân tích ở Chương 1 tạo điều kiện trực tiếp cho các giải pháp: Luật số "
     "93/2025/QH15 giao tổ chức chủ trì quyền sở hữu phần kết quả sử dụng ngân sách nhà nước và cho phép dùng quỹ phát "
     "triển khoa học và công nghệ cho đăng ký, bảo hộ quyền sở hữu trí tuệ; Luật Giáo dục đại học số 125/2025/QH15 cho "
     "phép thành lập doanh nghiệp khoa học và công nghệ; Nghị định số 100/2026/NĐ-CP yêu cầu lập Danh mục quyền sở hữu "
     "trí tuệ; Thông tư số 83/2026/TT-BGDĐT yêu cầu có quy định về sở hữu trí tuệ và tính sáng chế, giải pháp hữu ích "
     "với hệ số quy đổi cao."),
    ("Thứ nhất, Luật Khoa học, Công nghệ và Đổi mới sáng tạo số 93/2025/QH15 mang lại",
     "Về căn cứ thực tiễn, Chương 2 xác định điểm nghẽn ở bước chuyển từ kết quả nghiên cứu sang đơn đăng ký: trong 9 "
     "đề tài có sản phẩm thuộc nhóm sở hữu công nghiệp, mới có 1 đơn sáng chế. Bốn hạn chế đi kèm là kết quả có thể bảo "
     "hộ chưa được chuyển thành đơn, quy chế nội bộ chưa đồng bộ, chưa có khai thác thương mại và quản lý còn phân tán. "
     "Các giải pháp được thiết kế để xử lý trực tiếp các hạn chế này."),
    ("Thứ hai, Luật Giáo dục đại học số 125/2025/QH15 tại Khoản 1 Điều 28", None),
    ("Thứ ba, Nghị định số 100/2026/NĐ-CP tại Điều 9a", None),
    ("Thứ tư, Thông tư số 83/2026/TT-BGDĐT (ban hành ngày 30/9/2026", None),
    ("Nhằm bảo đảm tính khoa học và thực tiễn, nhóm nghiên cứu tiến hành",
     "Để xác định hướng giải pháp, nhóm nghiên cứu tổng hợp các yếu tố bên trong và bên ngoài thành ma trận điểm mạnh, "
     "điểm yếu, cơ hội và thách thức tại Bảng 3.1."),
    ("Từ ma trận phân tích tại Bảng 3.1, nhóm nghiên cứu xây dựng", None),
    ("Bảng 3.2. Các phương án kết hợp từ ma trận", None),
    ("(Nguồn: Nhóm nghiên cứu xây dựng từ Bảng 3.1)", None),
    # 3.1.3: bớt cụm từ khẩu hiệu
    ("Hệ thống giải pháp quản lý quyền sở hữu trí tuệ tại Trường Đại học Thành Đô được xây dựng trên bốn nguyên tắc",
     "Các giải pháp được xây dựng theo bốn nguyên tắc:"),
    ("Nguyên tắc đồng bộ và liên thông khép kín:",
     "Nguyên tắc đồng bộ: Các giải pháp gắn với bốn khâu sáng tạo, xác lập, khai thác và bảo vệ; hạn chế ở một khâu "
     "sẽ ảnh hưởng đến kết quả của các khâu sau."),
    ("Nguyên tắc hài hòa lợi ích và lấy nhà khoa học làm trung tâm:",
     "Nguyên tắc hài hòa lợi ích: Cơ chế phân chia lợi ích phải công bằng, minh bạch giữa Nhà trường, tác giả và đối "
     "tác, đủ để tạo động lực cho nhà khoa học."),
    ("Hệ thống giải pháp nâng cao hiệu quả quản lý quyền sở hữu trí tuệ tại Trường Đại học Thành Đô được thiết kế",
     "Các giải pháp được thiết kế theo bốn khâu của quản lý quyền sở hữu trí tuệ:"),
    # 3.2.1
    ("Mục tiêu giải pháp: Chuyển dịch mạnh mẽ định hướng nghiên cứu",
     "Mục tiêu giải pháp: Tăng số kết quả nghiên cứu có khả năng bảo hộ và nhận diện tài sản trí tuệ ngay từ khi phê "
     "duyệt thuyết minh đề tài; phấn đấu nâng tỷ lệ đề tài cấp cơ sở có sản phẩm thuộc nhóm sở hữu công nghiệp từ 23,7% "
     "lên 40% vào năm 2028."),
    ("Hai là, cải cách biểu mẫu thuyết minh đề tài",
     "Hai là, bổ sung vào biểu mẫu thuyết minh đề tài mục \"Dự kiến kết quả tạo lập tài sản trí tuệ và phương án bảo "
     "hộ\". Tác giả thực hiện tra cứu sơ bộ, hoặc được " + TPL + " hỗ trợ tra cứu sơ bộ, trước khi nộp thuyết minh."),
    ("Ba là, ban hành quy chế khai báo sáng kiến nội bộ",
     "Ba là, ban hành biểu mẫu khai báo kết quả mới và mẫu nhật ký phòng thí nghiệm, thực hiện nhiệm vụ xây dựng quy "
     "trình, biểu mẫu khai báo của " + TPL + " tại Điều 11 Quyết định số 217/QĐ-ĐHTĐ. Giảng viên, nhất là tại Viện Y - "
     "Dược, ghi chép nhật ký nghiên cứu và khai báo ngay khi có giải pháp kỹ thuật mới có khả năng ứng dụng."),
    ("Bốn là, tổ chức các khóa tập huấn thường niên",
     "Bốn là, tổ chức tập huấn hằng năm về nhận diện tài sản trí tuệ và tra cứu sáng chế trên PATENTSCOPE, Google "
     "Patents và cơ sở dữ liệu của Cục Sở hữu trí tuệ cho giảng viên tham gia nghiên cứu."),
    ("Chủ thể chủ trì và phối hợp: Phòng Quản lý Khoa học và Công nghệ chủ trì; phối hợp với Viện Quản trị",
     "Chủ thể chủ trì và phối hợp: " + TPL + " chủ trì; phối hợp các viện thông qua Phòng Công nghệ, Đổi mới sáng tạo "
     "và Khởi nghiệp."),
    # 3.2.2
    ("Mục tiêu giải pháp: Từng bước khắc phục điểm nghẽn",
     "Mục tiêu giải pháp: Khắc phục điểm nghẽn ở bước chuyển từ kết quả nghiên cứu sang đơn đăng ký; nâng tỷ "
     "lệ nộp đơn trong nhóm đề tài có sản phẩm sở hữu công nghiệp, phấn đấu đạt từ 2 đến 3 đơn sáng chế, giải pháp hữu "
     "ích mỗi năm trong giai đoạn 2028 - 2030; không để kết quả có thể bảo hộ bị công bố trước khi nộp đơn."),
    ("Một là, ban hành Quy trình sàng lọc bảo hộ bắt buộc",
     "Một là, ban hành quy trình sàng lọc khả năng bảo hộ trước khi nghiệm thu và công bố, cụ thể hóa yêu cầu xin ý "
     "kiến trước khi bộc lộ tại Điều 10 Quyết định số 217/QĐ-ĐHTĐ. Báo cáo nghiệm thu, bản thảo bài báo có giải pháp "
     "kỹ thuật hoặc công thức chế phẩm được chuyển Tổ tư vấn sở hữu trí tuệ xem xét trước khi gửi đăng."),
    ("Hai là, thành lập Hội đồng thẩm định và tư vấn sở hữu trí tuệ",
     "Hai là, thành lập Tổ tư vấn sở hữu trí tuệ gồm đại diện " + TPL + ", Bộ phận Pháp chế và chuyên gia theo từng "
     "lĩnh vực, làm việc kiêm nhiệm theo từng hồ sơ, không tổ chức thành hội đồng cố định."),
    ("Ba là, thiết lập dòng dự toán ngân sách chuyên biệt",
     "Ba là, lập dòng dự toán riêng hỗ trợ xác lập quyền sở hữu công nghiệp. Với mỗi đề tài được Tổ tư vấn đánh giá có "
     "khả năng bảo hộ, Nhà trường hỗ trợ từ 15 đến 30 triệu đồng cho tra cứu chuyên sâu, phí, lệ phí và thuê tổ chức "
     "đại diện sở hữu công nghiệp hoàn thiện bản mô tả, yêu cầu bảo hộ."),
    ("Bốn là, ký kết thỏa thuận hợp tác chiến lược với các tổ chức đại diện",
     "Bốn là, ký thỏa thuận hợp tác với tổ chức đại diện sở hữu công nghiệp để được tư vấn, rút ngắn thời gian chuẩn bị "
     "và bảo đảm chất lượng hồ sơ nộp tại Cục Sở hữu trí tuệ."),
    ("Chủ thể chủ trì và phối hợp: Phòng Quản lý Khoa học và Công nghệ chủ trì; phối hợp với Bộ phận Pháp chế",
     "Chủ thể chủ trì và phối hợp: " + TPL + " chủ trì; phối hợp Bộ phận Pháp chế, Phòng Tài chính - Kế toán và tổ "
     "chức đại diện sở hữu công nghiệp."),
    # 3.2.3
    ("Mục tiêu giải pháp: Xóa bỏ xung đột thể chế nội bộ",
     "Mục tiêu giải pháp: Thống nhất cơ chế phân chia lợi ích giữa các quy chế nội bộ và với Luật số 93/2025/QH15; "
     "phấn đấu đến năm 2028 có hợp đồng chuyển giao hoặc chuyển quyền sử dụng đầu tiên phát sinh doanh thu; nghiên cứu "
     "khả năng hình thành doanh nghiệp khởi nguồn công nghệ khi đủ điều kiện."),
    ("Một là, sửa đổi, bổ sung Điều 36 Quyết định số 213/QĐ-ĐHTĐ",
     "Một là, sửa đổi khoản 4 Điều 36 Quyết định số 213/QĐ-ĐHTĐ: rà soát khoản nộp ngân sách nhà nước và mức trần 100 "
     "triệu đồng tại điểm a cho phù hợp với khoản 2 Điều 25 và điểm a khoản 3 Điều 28 Luật số 93/2025/QH15; thống nhất "
     "tỷ lệ tại điểm b với Điều 13 Quyết định số 217/QĐ-ĐHTĐ. Đối với phần kết quả sử dụng ngân sách nhà nước, sau khi "
     "bảo đảm tác giả được thưởng tối thiểu 30% lợi nhuận thu được, nhóm nghiên cứu đề xuất phương án tham khảo phân bổ "
     "phần còn lại: 40% bổ sung Quỹ Phát triển khoa học và công nghệ của Nhà trường; 20% cho đơn vị trực tiếp nghiên "
     "cứu; 10% thưởng cho đơn vị, cá nhân trực tiếp kết nối, đàm phán hợp đồng thương mại hóa."),
    ("Hai là, giao quyền tự chủ cho Viện Quản trị và Công nghệ (Viện REK)",
     "Hai là, " + TPL + " phối hợp Phòng Công nghệ, Đổi mới sáng tạo và Khởi nghiệp của các viện làm đầu mối xúc tiến "
     "thương mại hóa theo Điều 11 Quyết định số 217/QĐ-ĐHTĐ: lập danh mục kết quả có thể chuyển giao, tìm kiếm đối tác, "
     "phối hợp Bộ phận Pháp chế đàm phán hợp đồng chuyển giao, chuyển quyền sử dụng; thuê tổ chức thẩm định giá khi cần "
     "định giá tài sản trí tuệ, tận dụng chương trình thí điểm hỗ trợ xác định giá trị quyền sở hữu trí tuệ theo Quyết "
     "định số 1624/QĐ-TTg."),
    ("Ba là, triển khai đề án thành lập doanh nghiệp khởi nguồn công nghệ",
     "Ba là, nghiên cứu khả năng hình thành doanh nghiệp khởi nguồn công nghệ khi đủ điều kiện, tức là khi đã có văn "
     "bằng bảo hộ và đối tác thị trường cho sản phẩm dược liệu của Viện Y - Dược; khoản 1 Điều 28 Luật Giáo dục đại học "
     "số 125/2025/QH15 là căn cứ để Nhà trường thành lập hoặc góp vốn vào doanh nghiệp khoa học và công nghệ."),
    ("Chủ thể chủ trì và phối hợp: Viện Quản trị và Công nghệ (Viện REK) chủ trì",
     "Chủ thể chủ trì và phối hợp: " + TPL + " chủ trì xúc tiến thương mại hóa; " + TPL + " và Bộ phận Pháp chế chủ "
     "trì sửa đổi quy chế phân chia lợi ích; Phòng Tài chính - Kế toán phối hợp."),
    ("Điều kiện bảo đảm nguồn lực: Nhà trường phê duyệt Quy chế tài chính đặc thù",
     "Điều kiện bảo đảm nguồn lực: Nhà trường bổ sung quy định tài chính cho hoạt động thương mại hóa, gồm chi phí định "
     "giá, đàm phán và phân chia nguồn thu."),
    # 3.2.4
    ("Mục tiêu giải pháp: Hướng tới bảo vệ tối đa",
     "Mục tiêu giải pháp: Bảo vệ tài sản trí tuệ và thông tin chưa công bố của Nhà trường; theo dõi đầy đủ thời hạn, "
     "nghĩa vụ duy trì văn bằng; đến năm 2027 hoàn thành Danh mục quyền sở hữu trí tuệ và cơ sở dữ liệu tập trung về "
     "tài sản trí tuệ."),
    ("Một là, xây dựng và đưa vào vận hành Hệ thống thông tin quản lý tài sản trí tuệ",
     "Một là, xây dựng cơ sở dữ liệu tập trung về tài sản trí tuệ trên nền tảng quản trị số của Nhà trường, đồng thời "
     "là Danh mục quyền sở hữu trí tuệ theo Điều 9a Nghị định số 65/2023/NĐ-CP: lưu hồ sơ thuyết minh, kết quả đề tài, "
     "tình trạng đơn, cảnh báo thời hạn nộp phí duy trì và dữ liệu hợp đồng chuyển giao."),
    ("Hai là, ban hành Quy định về bảo mật thông tin và quản trị bí mật kinh doanh",
     "Hai là, cụ thể hóa Điều 10 Quyết định số 217/QĐ-ĐHTĐ thành quy trình bảo mật tại phòng thí nghiệm và trong hợp "
     "đồng hợp tác nghiên cứu; áp dụng mẫu cam kết bảo mật đối với sinh viên, học viên và chuyên gia tham gia các đề "
     "tài có kết quả có thể bảo hộ."),
    ("Ba là, hoàn thiện quy định về liêm chính học thuật, duy trì bắt buộc",
     "Ba là, hoàn thiện quy định về liêm chính học thuật theo Tiêu chí 1.1 Thông tư số 83/2026/TT-BGDĐT, kiểm tra "
     "trùng lặp trước khi nghiệm thu đề tài; theo dõi hành vi xâm phạm nhãn hiệu Thanh do University và Thado Edupark "
     "trên môi trường số."),
    ("Chủ thể chủ trì và phối hợp: Bộ phận Pháp chế chủ trì công tác bảo vệ quyền",
     "Chủ thể chủ trì và phối hợp: Bộ phận Pháp chế chủ trì bảo vệ quyền và xử lý vi phạm; " + TPL + " chủ trì xây "
     "dựng cơ sở dữ liệu, phối hợp đơn vị phụ trách công nghệ thông tin của Nhà trường."),
    # 3.3
    ("Để các giải pháp không dừng lại ở định hướng lý thuyết",
     "Bảng 3.2 tóm tắt đơn vị chủ trì và kết quả cần đạt của từng giải pháp, gắn với hạn chế đã xác định tại Mục 2.3.2."),
    ("Lý do lựa chọn đơn vị thí điểm: Viện Y - Dược là đơn vị nghiên cứu nòng cốt",
     "Lý do lựa chọn đơn vị thí điểm: Viện Y - Dược có 42 tiến sĩ và chủ trì 9 trên 11 đề tài cấp cơ sở có sản phẩm "
     "có thể bảo hộ (Bảng 2.4), chủ yếu về chiết xuất dược liệu và bào chế thuốc."),
    ("Lộ trình triển khai hệ thống giải pháp nâng cao hiệu quả quản lý",
     "Lộ trình triển khai được chia thành ba giai đoạn:"),
    ("Nội dung trọng tâm: Sửa đổi, bổ sung Quyết định số 213/QĐ-ĐHTĐ",
     "Nội dung trọng tâm: Sửa đổi Quyết định số 213/QĐ-ĐHTĐ (điều chỉnh khoản 4 Điều 36); ban hành quy trình sàng lọc "
     "trước công bố; bố trí dòng dự toán hỗ trợ nộp đơn từ Quỹ Phát triển khoa học và công nghệ; thí điểm tại Viện Y - "
     "Dược; xây dựng cơ sở dữ liệu tài sản trí tuệ."),
    ("Giai đoạn 2 (Năm 2028 - 2029): Mở rộng toàn diện",
     "Giai đoạn 2 (Năm 2028 - 2029): Mở rộng áp dụng, tăng số đơn đăng ký và xúc tiến thương mại hóa."),
    ("Nội dung trọng tâm: Áp dụng bắt buộc quy trình sàng lọc",
     "Nội dung trọng tâm: Phấn đấu áp dụng quy trình sàng lọc cho 100% đề tài cấp cơ sở; nộp từ 2 đến 3 đơn sáng chế, "
     "giải pháp hữu ích mỗi năm; " + TPL + " xúc tiến đàm phán từ 1 đến 2 hợp đồng chuyển quyền sử dụng với doanh "
     "nghiệp."),
    ("Giai đoạn 3 (Năm 2030): Vận hành hệ sinh thái",
     "Giai đoạn 3 (Năm 2030): Đánh giá kết quả và nghiên cứu khả năng hình thành doanh nghiệp khởi nguồn công nghệ."),
    ("Nội dung trọng tâm: Thành lập và đưa vào vận hành ít nhất 1 doanh nghiệp",
     "Nội dung trọng tâm: Đánh giá kết quả thực hiện các giải pháp theo Bảng 3.3; nghiên cứu khả năng hình thành doanh "
     "nghiệp khởi nguồn công nghệ dựa trên tài sản trí tuệ của Viện Y - Dược khi đã có văn bằng bảo hộ và đối tác thị "
     "trường; dùng nguồn thu từ chuyển giao (nếu có) để tái đầu tư cho nghiên cứu."),
    ("Để đo lường khách quan tiến độ và hiệu quả",
     "Để theo dõi tiến độ và hiệu quả, nhóm nghiên cứu đề xuất bộ chỉ số phân tầng từ thể chế đến khai thác kinh tế "
     "tại Bảng 3.3."),
    ("Bộ chỉ số tại Bảng 3.4 là công cụ điều hành then chốt",
     "Bảng 3.3 giúp Ban Giám hiệu và Hội đồng trường theo dõi định kỳ tiến độ triển khai các giải pháp và điều chỉnh "
     "việc phân bổ nguồn lực theo từng giai đoạn."),
    ("Chương 3 đã xây dựng một hệ thống giải pháp toàn diện",
     "Chương 3 đề xuất bốn nhóm giải pháp tương ứng với bốn hạn chế đã xác định ở Chương 2: nhận diện tài sản trí tuệ "
     "từ khâu thuyết minh đề tài; sàng lọc khả năng bảo hộ trước công bố, có Tổ tư vấn và kinh phí hỗ trợ nộp đơn; thống "
     "nhất cơ chế phân chia lợi ích với Luật số 93/2025/QH15 và xúc tiến thương mại hóa; cụ thể hóa quy định bảo mật và "
     "xây dựng cơ sở dữ liệu tài sản trí tuệ. Các giải pháp dựa trên ma trận SWOT và các quy định mới của Luật số "
     "93/2025/QH15, Luật số 125/2025/QH15, Nghị định số 100/2026/NĐ-CP và Thông tư số 83/2026/TT-BGDĐT."),
    ("Kế hoạch tổ chức thực hiện với bước đi thí điểm thận trọng",
     "Việc tổ chức thực hiện bắt đầu bằng thí điểm tại Viện Y - Dược, triển khai theo ba giai đoạn 2026 - 2030 và được "
     "theo dõi bằng bộ chỉ số tại Bảng 3.3. Kết quả thí điểm là căn cứ để điều chỉnh quy trình trước khi áp dụng cho "
     "toàn trường."),
]

S2.MOT_PHAN += [
    ("Nguyên tắc khả thi và phù hợp điều kiện thực tiễn:", "tránh áp đặt các mô hình cồng kềnh, quan liêu hoặc vượt quá "
     "khả năng cân đối ngân sách của Nhà trường", "tránh các mô hình vượt quá khả năng tài chính và nhân lực của Nhà trường"),
    ("Đánh giá rút kinh nghiệm: Sau 12 tháng", "Phòng Quản lý Khoa học và Công nghệ", TPL),
    ("Quy mô và chỉ tiêu thí điểm: Lựa chọn từ 3 đến 5", "Quỹ Khoa học và Công nghệ", "Quỹ Phát triển khoa học và công nghệ"),
    # Bảng 3.4: "phấn đấu"
    ("Ban hành trước ngày 31/5/2027", "Đạt 100% giảng viên mới hàng năm", "Phấn đấu đạt 100% giảng viên mới hằng năm"),
    ("Đạt 100% đề tài cấp cơ sở từ năm 2027",
     "Đạt 100% đề tài cấp cơ sở từ năm 2027<br>Đạt 100% đề tài khối Y - Dược và Kỹ thuật<br>Tối đa không quá 15 ngày làm việc",
     "Phấn đấu đạt 100% đề tài cấp cơ sở từ năm 2027\nPhấn đấu đạt 100% đề tài khối Y - Dược và Kỹ thuật\n"
     "Không quá 15 ngày làm việc"),
    ("Từ 02 đơn năm 2027", "Từ 02 đơn năm 2027; đạt 03 - 05 đơn/năm từ 2028",
     "Phấn đấu 02 - 03 đơn/năm giai đoạn 2028 - 2030"),
    ("Từ 02 đơn năm 2027", "Đăng ký bảo hộ cho 100% giáo trình trọng điểm",
     "Phấn đấu đăng ký quyền tác giả cho 100% giáo trình trọng điểm"),
    ("Tối thiểu 02 - 03 văn bằng/năm theo Kế hoạch 07", "Tối thiểu 02 - 03 văn bằng/năm theo Kế hoạch 07",
     "Phấn đấu có từ 02 - 03 văn bằng bảo hộ sở hữu công nghiệp được cấp trong giai đoạn đến năm 2030"),
    ("Một là, tái cơ cấu định hướng danh mục", "tái cơ cấu", "điều chỉnh"),
    ("Ký tối thiểu 01 - 02 hợp đồng/năm từ 2028", "Bảo đảm tối thiểu 30% lợi nhuận sau thuế",
     "Tối thiểu 30% lợi nhuận đối với phần kết quả sử dụng ngân sách nhà nước"),
]

# Bảng 3.1: viết lại bốn ô, nêu điểm mạnh, điểm yếu, không đưa tỷ lệ chi tiết
S2.DOAN_CHUA += [
    ("S1. Ban hành quy định nội bộ từ sớm",
     "S1. Đã có quy chế về hoạt động khoa học công nghệ (Quyết định 213) và quản trị tài sản trí tuệ (Quyết định 217).\n"
     "S2. Năng lực nghiên cứu tăng nhanh: số bài báo năm 2025 gấp 7 lần năm 2021.\n"
     "S3. Đã có kết quả xác lập quyền: 9 văn bằng (quyền tác giả, nhãn hiệu, kiểu dáng công nghiệp), 1 đơn sáng chế; "
     "11 đề tài có sản phẩm có thể bảo hộ."),
    ("W1. Tỷ lệ chuyển hóa đề tài sang đơn bảo hộ còn thấp",
     "W1. Tỷ lệ chuyển hóa kết quả nghiên cứu thành đơn đăng ký sở hữu công nghiệp còn thấp.\n"
     "W2. Chưa có quy trình sàng lọc trước công bố; quy chế nội bộ chưa thống nhất.\n"
     "W3. Chưa có bộ phận chuyên trách về sở hữu trí tuệ; dữ liệu quản lý phân tán.\n"
     "W4. Chưa ghi nhận hợp đồng thương mại hóa phát sinh doanh thu."),
    ("O1. Khung pháp luật mới trao quyền tự chủ",
     "O1. Luật số 93/2025/QH15 giao tổ chức chủ trì quyền sở hữu phần kết quả sử dụng ngân sách nhà nước; Luật số "
     "125/2025/QH15 cho phép thành lập doanh nghiệp khoa học và công nghệ.\n"
     "O2. Thông tư 83/2026/TT-BGDĐT tính sáng chế, giải pháp hữu ích với hệ số quy đổi cao.\n"
     "O3. Chương trình thí điểm hỗ trợ xác định giá trị quyền sở hữu trí tuệ theo Quyết định 1624/QĐ-TTg."),
    ("T1. Áp lực tuân thủ kiểm định chất lượng",
     "T1. Áp lực công bố có thể dẫn đến bộc lộ kết quả trước khi nộp đơn, làm mất tính mới.\n"
     "T2. Thủ tục xác lập quyền đối với sáng chế kéo dài; chi phí bảo hộ phát sinh trong suốt thời hạn bảo hộ.\n"
     "T3. Nguy cơ tranh chấp quyền giữa giảng viên, người học và doanh nghiệp cùng tham gia nghiên cứu."),
]

# Bảng 3.3 rút gọn: mỗi giải pháp một hạn chế được xử lý, đơn vị chủ trì và một đến hai kết quả chính
S2.BANG_THAY.append(dict(
    dau="Tên giải pháp",
    rong=[2300, 2100, 2300, 2370],
    cot=["Giải pháp", "Hạn chế được xử lý (Mục 2.3.2)", "Chủ trì; phối hợp", "Kết quả cần đạt"],
    dong=[
        ["3.2.1. Sáng tạo và nhận diện tài sản trí tuệ sớm", "Hạn chế thứ nhất",
         TPL + "; các viện (Phòng Công nghệ, Đổi mới sáng tạo và Khởi nghiệp)",
         "Thuyết minh đề tài có mục dự kiến tài sản trí tuệ; tăng tỷ lệ đề tài có sản phẩm có thể bảo hộ"],
        ["3.2.2. Sàng lọc trước công bố và hỗ trợ nộp đơn", "Hạn chế thứ nhất, thứ hai",
         TPL + "; Bộ phận Pháp chế, Phòng Tài chính - Kế toán",
         "Quy trình sàng lọc được ban hành; 2 đến 3 đơn sở hữu công nghiệp mỗi năm giai đoạn 2028 - 2030"],
        ["3.2.3. Phân chia lợi ích và thương mại hóa", "Hạn chế thứ hai, thứ ba",
         TPL + ", Bộ phận Pháp chế; Phòng Tài chính - Kế toán",
         "Quy chế phân chia lợi ích sửa đổi; hợp đồng chuyển giao đầu tiên"],
        ["3.2.4. Bảo vệ quyền và quản trị dữ liệu", "Hạn chế thứ tư", "Bộ phận Pháp chế; " + TPL,
         "Danh mục quyền sở hữu trí tuệ và cơ sở dữ liệu tập trung; không để quá hạn nghĩa vụ duy trì văn bằng"],
    ],
))

if __name__ == "__main__":
    S2.main()

# -*- coding: utf-8 -*-
"""Rà soát bài báo tạp chí theo văn bản gốc và Chương 2 đã chuẩn hóa, xuất tệp có theo dõi thay đổi."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from track_changes import ap_dung  # noqa: E402
from ra_soat_chuong3 import MERGE  # noqa: E402

GOC = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
VAO = os.path.join(GOC, "Bai_bao_khoa_hoc_Tap_chi_DHTD.docx")
RA = os.path.join(GOC, "Bai_bao_khoa_hoc_Tap_chi_DHTD_ra_soat.docx")

SUA = [
    # --- Đặt vấn đề: Luật 125 (Điều 28), Luật 93 (Điều 135), Luật 131 (Điều 86)
    ("Luật Giáo dục đại học số 125/2025/QH15 tại Điều 27 và Điều 28 bắt buộc công khai kết quả khoa học công nghệ hằng "
     "năm trên Nền tảng số quốc gia, đồng thời khẳng định quyền của trường đại học được thành lập doanh nghiệp quản lý tài "
     "sản trí tuệ, định giá và góp vốn.",
     "Luật Giáo dục đại học số 125/2025/QH15 tại điểm đ khoản 3 Điều 28 yêu cầu công khai năng lực, kết quả hoạt động "
     "khoa học, công nghệ và đổi mới sáng tạo, cập nhật hằng năm trên Nền tảng số quốc gia; khoản 1 và điểm d khoản 2 Điều "
     "28 khẳng định quyền của trường đại học được thành lập doanh nghiệp quản lý tài sản trí tuệ, định giá, góp vốn và "
     "phân chia lợi ích từ tài sản trí tuệ."),
    ("tại khoản 7 Điều 71 sửa đổi Điều 135 Luật Sở hữu trí tuệ (thể hiện tại Văn bản hợp nhất số 67/VBHN-VPQH năm 2026) "
     "quy định thù lao tối thiểu 30% lợi nhuận thuần cho tác giả sáng chế và bãi bỏ mức trần thù lao.",
     "tại điểm b và điểm h khoản 7 Điều 71 sửa đổi Điều 135 Luật Sở hữu trí tuệ, thể hiện tại Văn bản hợp nhất số "
     "67/VBHN-VPQH ngày 23 tháng 3 năm 2026, theo đó thù lao cho tác giả được xác định theo thỏa thuận, trường hợp không có "
     "thỏa thuận là 10% lợi nhuận trước thuế khi chủ sở hữu tự sử dụng hoặc 15% số tiền nhận được mỗi lần chuyển giao "
     "quyền sử dụng, đồng thời bãi bỏ khung thù lao đối với nhiệm vụ sử dụng ngân sách nhà nước."),
    ("tại điểm c khoản 1 Điều 86 (có hiệu lực từ ngày 01/4/2026) giao quyền chủ động nộp đơn đăng ký sở hữu đối với "
     "nhiệm vụ sử dụng ngân sách nhà nước.",
     "có hiệu lực từ ngày 01/4/2026 đã bổ sung điểm c khoản 1 Điều 86, theo đó tổ chức được giao quyền quản lý, sử dụng, "
     "quyền sở hữu kết quả của nhiệm vụ khoa học, công nghệ và đổi mới sáng tạo sử dụng ngân sách nhà nước có quyền đăng "
     "ký sáng chế, kiểu dáng công nghiệp, thiết kế bố trí."),
    # --- Phương pháp: số liệu đã chuẩn hóa; văn bản nội bộ theo tài liệu gốc
    ("với 580 bản ghi sản phẩm khoa học. Tác giả tiến hành khử trùng đếm để xác định quy mô thực tế 470 sản phẩm độc lập;",
     "với 582 bản ghi sản phẩm khoa học. Nhóm nghiên cứu tiến hành khử trùng đếm để xác định quy mô thực tế khoảng 475 sản "
     "phẩm độc lập;"),
    ("(Quyết định số 213/QĐ-ĐHTĐ, Quyết định số 217/QĐ-ĐHTĐ, Quy chế chi tiêu nội bộ).",
     "(Quyết định số 213/QĐ-ĐHTĐ, Quyết định số 217/QĐ-ĐHTĐ, Quy chế chi tiêu nội bộ năm 2026, Kế hoạch số 07/KH-ĐHTĐ)."),
    # --- Kết quả 3.1
    ("Tuy nhiên, mức độ chuyển hóa sang quyền sở hữu công nghiệp lại rất thấp (chỉ 2 đơn đăng ký sáng chế từ nghiên cứu "
     "trong 5 năm).",
     "Tuy nhiên, mức độ chuyển hóa sang quyền sở hữu công nghiệp lại rất thấp: trong 38 đề tài cấp cơ sở, 11 đề tài có "
     "sản phẩm đủ điều kiện xác lập quyền nhưng chỉ 1 đề tài được nộp đơn; nếu chỉ xét 8 đề tài đủ điều kiện giai đoạn "
     "2021 - 2024 thì không đề tài nào nộp đơn."),
    ("(0,237 đối với đề tài, 0,314 đối với giáo trình)", "(0,238 đối với đề tài, 0,314 đối với giáo trình)"),
    ("Con số này phản ánh thực trạng là phần lớn cán bộ giảng viên hiện chưa chú trọng và chưa có thói quen tạo lập tài "
     "sản trí tuệ.",
     "Khả năng hình thành tài sản trí tuệ vì vậy phụ thuộc vào một số ít cá nhân thay vì một quy trình vận hành ổn định. "
     "Cơ chế khuyến khích cũng nghiêng về công bố: bài báo WoS hạng Q1 được quy đổi 300 giờ và thưởng 20 triệu đồng ngay "
     "khi đăng, trong khi bằng sáng chế chuẩn Việt Nam được quy đổi 360 giờ nhưng không có tiền thưởng và chỉ được ghi "
     "nhận sau khi có văn bằng."),
    # --- Kết quả 3.2
    ("Ba văn bản nội bộ (Quyết định số 213/QĐ-ĐHTĐ, Quyết định số 217/QĐ-ĐHTĐ, Quy chế chi tiêu nội bộ và Điều lệ Quỹ Ngô "
     "Xuân Độ) quy định ba công thức chia lợi ích khác nhau và giữ trần thù lao 100 triệu đồng trái với tinh thần Luật "
     "93/2025/QH15.",
     "Bốn văn bản nội bộ (Quyết định số 213/QĐ-ĐHTĐ, Quyết định số 217/QĐ-ĐHTĐ, Quy chế chi tiêu nội bộ và Điều lệ "
     "Quỹ Học bổng sau tiến sĩ Ngô Xuân Độ) chứa năm quy định khác nhau về phân chia lợi ích; điểm a khoản 4 Điều 36 Quyết "
     "định 213 vẫn giữ trần thù lao 100 triệu đồng, không còn tương thích với việc Luật 93/2025/QH15 bãi bỏ khung thù "
     "lao."),
    ("thiếu khâu rà soát bắt buộc tại biên bản nghiệm thu đề tài khiến nhiều sản phẩm bị bộc lộ quá 12 tháng và vĩnh viễn "
     "mất khả năng xác lập quyền.",
     "thiếu khâu rà soát bắt buộc tại biên bản nghiệm thu đề tài; Điều 10 Quy chế ban hành kèm Quyết định 217 vẫn đặt "
     "trách nhiệm tự nhận diện tài sản có thể bảo hộ lên tác giả, trong khi các sản phẩm đã công bố quá 12 tháng sẽ mất "
     "khả năng xác lập quyền sáng chế và giải pháp hữu ích."),
    ("Đồng thời, nhóm giáo trình và nhãn hiệu các đơn vị thành viên hệ sinh thái chưa được đưa vào quy trình bảo hộ bài "
     "bản.",
     "Đồng thời, 87 giáo trình, tài liệu giảng dạy biên soạn giai đoạn 2021 - 2025 chưa được đăng ký quyền tác giả. Đối "
     "chiếu với Kế hoạch số 07/KH-ĐHTĐ cho thấy trong hai năm 2024 - 2025, bài báo quốc tế đạt 283% chỉ tiêu, trong khi "
     "chỉ tiêu chuyển giao công nghệ năm 2025 không đạt và 6 văn bằng được cấp đều thuộc tài sản thương hiệu và hợp tác "
     "doanh nghiệp."),
    # --- Giải pháp
    ("Đưa sở hữu trí tuệ thành môn học bắt buộc cho sinh viên theo Quyết định 1624/QĐ-TTg;",
     "Đón đầu định hướng nghiên cứu đưa sở hữu trí tuệ thành nội dung học bắt buộc tại các cơ sở giáo dục đại học theo "
     "Quyết định 1624/QĐ-TTg;"),
    ("Ban hành Quy chế hợp nhất sửa đổi Quyết định 213 và Quyết định 217; bổ sung giải pháp hữu ích, sưu tập dữ liệu và "
     "giáo trình số vào phạm vi bảo hộ; bãi bỏ mức trần thù lao 100 triệu đồng theo Luật 93/2025/QH15;",
     "Ban hành Quy chế hợp nhất sửa đổi Quyết định 213 và Quyết định 217, thống nhất một danh mục tài sản và một cơ chế "
     "phân chia lợi ích, bổ sung nội dung liêm chính khoa học theo Thông tư số 83/2026/TT-BGDĐT; bãi bỏ mức trần thù lao "
     "100 triệu đồng phù hợp với Luật 93/2025/QH15;"),
    ("theo Điều 28 Luật Giáo dục đại học 2025 để thực hiện định giá và góp vốn thương mại hóa.",
     "theo khoản 1 Điều 28 Luật Giáo dục đại học 2025 để thực hiện định giá và góp vốn thương mại hóa."),
    ("phối hợp với Viện Nghiên cứu thành lập Trung tâm tư vấn định giá tài sản trí tuệ (theo Quyết định 1624/QĐ-TTg);",
     "phối hợp với Viện Nghiên cứu giáo dục và Chuyển giao tri thức thành lập Trung tâm tư vấn định giá tài sản trí tuệ "
     "theo định hướng phát triển các trung tâm tư vấn, hỗ trợ định giá trong cơ sở giáo dục đại học tại Quyết định "
     "1624/QĐ-TTg;"),
    ("kết nối tự động với Nền tảng số quốc gia theo Điều 28 Luật Giáo dục đại học 2025.",
     "kết nối tự động với Nền tảng số quốc gia theo điểm đ khoản 3 Điều 28 Luật Giáo dục đại học 2025."),
    # --- Tài liệu tham khảo: đối chiếu số, ngày, cơ quan ban hành
    ("Quốc hội (2025), Luật Giáo dục đại học số 125/2025/QH15.",
     "Quốc hội (2025), Luật Giáo dục đại học số 125/2025/QH15 ngày 10 tháng 12 năm 2025."),
    ("Quốc hội (2025), Luật Khoa học, Công nghệ và Đổi mới sáng tạo số 93/2025/QH15.",
     "Quốc hội (2025), Luật Khoa học, công nghệ và đổi mới sáng tạo số 93/2025/QH15 ngày 27 tháng 6 năm 2025."),
    ("Quốc hội (2025), Luật sửa đổi, bổ sung một số điều của Luật Sở hữu trí tuệ số 131/2025/QH15.",
     "Quốc hội (2025), Luật số 131/2025/QH15 ngày 10 tháng 12 năm 2025 sửa đổi, bổ sung một số điều của Luật Sở hữu trí "
     "tuệ."),
    ("Ủy ban Thường vụ Quốc hội (2026), Văn bản hợp nhất số 67/VBHN-VPQH ngày 08 tháng 01 năm 2026 hợp nhất Luật Sở hữu "
     "trí tuệ.",
     "Văn phòng Quốc hội (2026), Văn bản hợp nhất số 67/VBHN-VPQH ngày 23 tháng 3 năm 2026 Luật Sở hữu trí tuệ."),
    ("Quyết định số 1624/QĐ-TTg ngày 18 tháng 8 năm 2026", "Quyết định số 1624/QĐ-TTg ngày 21 tháng 8 năm 2026"),
    ("Thủ tướng Chính phủ (2026), Chỉ thị số 02/CT-TTg về tăng cường thực thi quyền sở hữu trí tuệ.",
     "Thủ tướng Chính phủ (2026), Chỉ thị số 02/CT-TTg ngày 30 tháng 01 năm 2026 về tăng cường thực thi quyền sở hữu trí "
     "tuệ."),
    ("Trường Đại học Thành Đô (2021), Quyết định số 213/QĐ-ĐHTĐ về việc ban hành Quy chế hoạt động khoa học và công nghệ.",
     "Trường Đại học Thành Đô (2021), Quy chế hoạt động khoa học công nghệ ban hành kèm văn bản số 213/QĐ-ĐHTĐ ngày 28 "
     "tháng 12 năm 2021."),
    ("Trường Đại học Thành Đô (2026), Quyết định số 217/QĐ-ĐHTĐ về việc sửa đổi, bổ sung Quy chế hoạt động khoa học và "
     "công nghệ.",
     "Trường Đại học Thành Đô (2024), Quyết định số 217/QĐ-ĐHTĐ ban hành Quy chế quản trị tài sản trí tuệ tại Trường Đại "
     "học Thành Đô.\n"
     "10. Trường Đại học Thành Đô (2024), Kế hoạch số 07/KH-ĐHTĐ ngày 01 tháng 7 năm 2024 về hoạt động khoa học công nghệ "
     "giai đoạn 2024 - 2028, tầm nhìn 2035.\n"
     "11. Trường Đại học Thành Đô (2026), Quy chế chi tiêu nội bộ ban hành ngày 01 tháng 8 năm 2026.\n"
     "12. Bộ Giáo dục và Đào tạo (2026), Thông tư số 83/2026/TT-BGDĐT ngày 30 tháng 9 năm 2026 quy định Chuẩn cơ sở "
     "giáo dục đại học."),
]

if __name__ == "__main__":
    n = ap_dung(VAO, RA, SUA, author="Claude", merge_runs=MERGE if os.path.exists(MERGE) else None)
    print("Đã ghi:", RA, "|", n, "chỉnh sửa")

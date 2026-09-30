# -*- coding: utf-8 -*-
"""Rà soát Chương 3 theo văn bản gốc, xuất tệp có theo dõi thay đổi.

Căn cứ đối chiếu: thư mục VBPL (Luật 125/2025/QH15, Văn bản hợp nhất 67/VBHN-VPQH,
Quyết định 1624/QĐ-TTg, Chỉ thị 02/CT-TTg) và thư mục "Tai lieu thanh do"
(Quyết định 213, Quyết định 217 ban hành Quy chế quản trị tài sản trí tuệ năm 2024,
Kế hoạch 07/KH-ĐHTĐ, Thông tư 83/2026/TT-BGDĐT,
danh sách nhân sự 2026, danh mục đề tài cấp cơ sở).
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from track_changes import ap_dung  # noqa: E402

GOC = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
VAO = os.path.join(GOC, "Chuong_3_He_thong_giai_phap.docx")
RA = os.path.join(GOC, "Chuong_3_He_thong_giai_phap_ra_soat.docx")
MERGE = "/root/.claude/skills/synced/511a21b4-da62-4b21-9cb6-da109b6fbd84_b74c1dbc-fffb-4c9d-9a90-c57409abedcc/docx/scripts/merge_runs.py"

QD1068 = "Mục III Điều 1 Quyết định số 1068/QĐ-TTg đã được sửa đổi tại Quyết định số 1624/QĐ-TTg"

SUA = [
    # --- 3.1.1, Luật Giáo dục đại học 125/2025/QH15 (đối chiếu Điều 27, 28)
    ("Điều 27 quy định hoạt động khoa học, công nghệ, đổi mới sáng tạo và phát triển tài sản trí tuệ là một trong các "
     "chức năng cốt lõi của trường đại học;",
     "điểm e khoản 3 Điều 27 xác định đăng ký bản quyền hoặc bảo hộ, khai thác và phát triển tài sản trí tuệ là một nội "
     "dung của hoạt động khoa học, công nghệ và đổi mới sáng tạo trong cơ sở giáo dục đại học;"),
    ("Điều 28 khoản 2 điểm đ quy định nghĩa vụ công khai", "Điều 28 khoản 3 điểm đ quy định nghĩa vụ công khai"),
    ("Nền tảng số quản lý khoa học, công nghệ và đổi mới sáng tạo quốc gia. Những quy định này",
     "Nền tảng số quản lý khoa học, công nghệ và đổi mới sáng tạo quốc gia; điểm d khoản 3 Điều 28 giao trách nhiệm "
     "thành lập và vận hành Quỹ phát triển khoa học và công nghệ. Những quy định này"),
    # --- Luật 93/2025/QH15 sửa Điều 135 Luật Sở hữu trí tuệ (Văn bản hợp nhất 67/VBHN-VPQH, chú thích 254, 255)
    ("tại khoản 7 Điều 71 đã sửa đổi Điều 135", "tại điểm b và điểm h khoản 7 Điều 71 đã sửa đổi Điều 135"),
    ("khẳng định chủ sở hữu có nghĩa vụ trả thù lao cho tác giả theo thỏa thuận; trường hợp không có thỏa thuận thì mức "
     "thù lao trả cho tác giả sáng chế tối thiểu là 30% lợi nhuận thuần mà chủ sở hữu thu được từ việc sử dụng sáng chế. "
     "Luật khẳng định ưu tiên cơ chế thỏa thuận và không đặt trần thù lao.",
     "quy định chủ sở hữu có nghĩa vụ trả thù lao cho tác giả theo thỏa thuận; trường hợp không có thỏa thuận thì mức thù "
     "lao là 10% lợi nhuận trước thuế thu được tương ứng với giá trị mà sáng chế, kiểu dáng công nghiệp, thiết kế bố trí "
     "đóng góp khi chủ sở hữu tự sử dụng, hoặc 15% tổng số tiền chủ sở hữu nhận được trong mỗi lần thanh toán khi chuyển "
     "giao quyền sử dụng. Đồng thời, khoản 2 Điều 135 về khung thù lao đối với nhiệm vụ sử dụng ngân sách nhà nước đã bị "
     "bãi bỏ. Như vậy, pháp luật ưu tiên cơ chế thỏa thuận và không đặt trần thù lao."),
    ("mức giới hạn trần 100 triệu đồng tại Điều 36 Quyết định số 213/QĐ-ĐHTĐ",
     "mức giới hạn trần 100 triệu đồng tại điểm a khoản 4 Điều 36 Quyết định số 213/QĐ-ĐHTĐ"),
    # --- Luật 131/2025/QH15 bổ sung điểm c khoản 1 Điều 86 (chú thích 160)
    ("tại điểm c khoản 1 Điều 86 quy định tổ chức được giao quyền quản lý nhiệm vụ khoa học công nghệ sử dụng ngân sách "
     "nhà nước có quyền chủ động nộp đơn đăng ký xác lập quyền sở hữu độc quyền đối với sáng chế, kiểu dáng công nghiệp và "
     "thiết kế bố trí. Quy định này cho phép Nhà trường - với tư cách là tổ chức chủ trì nhiều đề tài có sử dụng ngân sách "
     "- có thể chủ động tiến hành đăng ký bảo hộ mà không cần chờ phê duyệt từ cơ quan giao nhiệm vụ, rút ngắn đáng kể "
     "thời gian từ khi phát sinh kết quả nghiên cứu đến khi nộp đơn.",
     "tại khoản 22 Điều 1 đã bổ sung điểm c khoản 1 Điều 86, theo đó tổ chức được giao quyền quản lý, sử dụng, quyền sở "
     "hữu kết quả của nhiệm vụ khoa học, công nghệ và đổi mới sáng tạo sử dụng ngân sách nhà nước có quyền đăng ký sáng "
     "chế, kiểu dáng công nghiệp, thiết kế bố trí là kết quả của nhiệm vụ đó. Quy định này tạo cơ sở pháp lý để Nhà "
     "trường, với tư cách tổ chức chủ trì ba đề tài cấp quốc gia do Quỹ Phát triển khoa học và công nghệ quốc gia tài "
     "trợ, chủ động xác lập quyền đối với kết quả của các đề tài này khi được nghiệm thu."),
    # --- Quyết định 1624/QĐ-TTg (đối chiếu khoản 15, 18, 20 Điều 1 và Phụ lục)
    ("Quyết định số 1624/QĐ-TTg của Thủ tướng Chính phủ phê duyệt sửa đổi",
     "Quyết định số 1624/QĐ-TTg ngày 21 tháng 8 năm 2026 của Thủ tướng Chính phủ về việc sửa đổi"),
    ("Quyết định đặt ra 5 yêu cầu trực tiếp đối với cơ sở giáo dục đại học mà Nhà trường cần đáp ứng:",
     "Quyết định đặt ra 5 yêu cầu trực tiếp đối với cơ sở giáo dục đại học mà Nhà trường cần đáp ứng, được viện dẫn theo "
     "khoản, điểm của " + QD1068 + ":"),
    ("Xác định trước các đối tượng quyền sở hữu trí tuệ cần đặt đối với kết quả nghiên cứu sử dụng ngân sách nhà nước ngay "
     "từ khâu đề xuất thuyết minh đề tài (điểm b khoản 4);",
     "Xác định các đối tượng quyền sở hữu trí tuệ cần đạt được đối với các kết quả nghiên cứu sử dụng ngân sách nhà nước; "
     "đối với cơ sở giáo dục khối kỹ thuật, công nghệ, tiến hành thủ tục đăng ký bảo hộ đồng thời với việc công bố bài báo "
     "khoa học về các kết quả nghiên cứu có tính ứng dụng cao (điểm b khoản 4);"),
    ("Thúc đẩy hình thành các trung tâm đổi mới sáng tạo, trung tâm tư vấn định giá và ươm tạo tài sản trí tuệ từ khâu "
     "hình thành ý tưởng (điểm d khoản 4);",
     "Thúc đẩy hình thành các trung tâm đổi mới sáng tạo ươm tạo các đối tượng quyền sở hữu trí tuệ từ khâu hình thành ý "
     "tưởng (điểm d khoản 4) và phát triển các trung tâm tư vấn, hỗ trợ định giá, khai thác thương mại quyền sở hữu trí "
     "tuệ trong cơ sở giáo dục đại học (điểm a khoản 6);"),
    ("Đưa sở hữu trí tuệ thành nội dung học bắt buộc trong chương trình đào tạo sinh viên khối kỹ thuật, công nghệ và các "
     "ngành liên quan (điểm h khoản 1);",
     "Xây dựng chương trình giáo dục, đào tạo về sở hữu trí tuệ và nghiên cứu đưa sở hữu trí tuệ, kỹ năng khai thác "
     "thương mại quyền sở hữu trí tuệ thành nội dung học bắt buộc tại các cơ sở giáo dục đại học (điểm b khoản 8);"),
    ("Phối hợp chuẩn bị cho Đề án gắn sở hữu trí tuệ với giáo dục đại học do Bộ Giáo dục và Đào tạo dự thảo vào quý 2 năm "
     "2027.",
     "Chuẩn bị tham gia Đề án tăng cường hoạt động sở hữu trí tuệ của cơ sở giáo dục đại học do Bộ Giáo dục và Đào tạo "
     "chủ trì, dự thảo trình phê duyệt trong quý II năm 2027 và thực hiện giai đoạn 2027 - 2030."),
    ("việc kết nối Mạng lưới Trung tâm Hỗ trợ công nghệ và đổi mới sáng tạo theo Quyết định số 1068/QĐ-TTg và Quyết định "
     "số 2205/QĐ-TTg giúp Nhà trường tiếp cận",
     "tư cách thành viên Mạng lưới Trung tâm Hỗ trợ công nghệ và đổi mới sáng tạo mà Nhà trường có từ năm 2023 giúp Nhà "
     "trường tiếp cận"),
    # --- Chỉ thị 02/CT-TTg (đối chiếu khoản 2, khoản 9)
    ("Chỉ thị số 02/CT-TTg của Thủ tướng Chính phủ về tăng cường thực thi quyền sở hữu trí tuệ yêu cầu ứng dụng công nghệ "
     "số, trí tuệ nhân tạo và kết nối cơ sở dữ liệu quốc gia về thực thi sở hữu trí tuệ đến năm 2026",
     "Chỉ thị số 02/CT-TTg ngày 30 tháng 01 năm 2026 của Thủ tướng Chính phủ về tăng cường thực thi quyền sở hữu trí tuệ "
     "giao Bộ Khoa học và Công nghệ xây dựng cơ sở dữ liệu quốc gia về thực thi quyền sở hữu trí tuệ, đưa vào vận hành "
     "trong năm 2026, tăng cường ứng dụng trí tuệ nhân tạo và công nghệ chuỗi khối trong bảo vệ quyền, đồng thời giao Bộ "
     "Giáo dục và Đào tạo nghiên cứu đưa chương trình giáo dục về sở hữu trí tuệ vào các hệ, cấp học phù hợp"),
    # --- 3.1.1, bổ sung căn cứ mới: Thông tư 83/2026/TT-BGDĐT (Điều 6, Phụ lục II)
    ("tạo thêm căn cứ cho việc xây dựng hạ tầng số hóa dữ liệu tài sản trí tuệ tại Nhà trường.",
     "Thứ bảy, Thông tư số 83/2026/TT-BGDĐT ngày 30 tháng 9 năm 2026 của Bộ trưởng Bộ Giáo dục và Đào tạo quy định Chuẩn "
     "cơ sở giáo dục đại học, có hiệu lực từ ngày 15 tháng 11 năm 2026 và thay thế Thông tư số 01/2024/TT-BGDĐT, lần đầu "
     "đưa quy định về sở hữu trí tuệ, liêm chính khoa học, liêm chính học thuật vào danh mục 21 nội dung quản trị nội bộ "
     "bắt buộc tại tiêu chí 1.1. Tiêu chí 6.2 chuyển sang tính sản phẩm khoa học, công nghệ và đổi mới sáng tạo quy đổi "
     "trên giảng viên quy đổi, trong đó một bằng độc quyền giải pháp hữu ích được tính 3 sản phẩm thay cho 1 sản phẩm như "
     "trước và một bằng độc quyền sáng chế được tính 5 sản phẩm; Bảng 6B Phụ lục II tách riêng khoản thu từ thương mại "
     "hóa kết quả nghiên cứu, sở hữu trí tuệ, spin-off, start-up. Dữ liệu về các kết quả này phải được cập nhật nhất "
     "quán trên HEMIS theo tiêu chí 1.3. Đây là căn cứ để hoàn thiện quy chế theo hướng bổ sung nội dung liêm chính và để "
     "Bộ chỉ số tại Mục 3.4.2 dùng chung định nghĩa sản phẩm với Chuẩn.", "doan_moi_sau"),
    # --- 3.1.2
    ("khoảng trống lớn về nhận thức và động lực", "khoảng trống lớn về quy trình và động lực"),
    # --- 3.2.1 (đối chiếu Quy chế quản trị tài sản trí tuệ 2024, Điều 3, 9, 11, 13)
    ("hiện nay Nhà trường đang vận hành đồng thời ba văn bản nội bộ (Quyết định số 213/QĐ-ĐHTĐ, Quyết định số "
     "217/QĐ-ĐHTĐ và Quy chế chi tiêu nội bộ) với ba công thức chia lợi ích khác nhau, dẫn đến sự chồng chéo, mâu thuẫn "
     "trong áp dụng thực tế. Đồng thời, phạm vi các đối tượng được bảo hộ còn thiếu nhiều loại hình quan trọng như giải "
     "pháp hữu ích, sưu tập dữ liệu, bí mật kinh doanh, giáo trình số và nhãn hiệu các pháp nhân thành viên.",
     "hiện nay Nhà trường đang vận hành đồng thời bốn văn bản nội bộ có quy định về sở hữu trí tuệ, gồm Quyết định số "
     "213/QĐ-ĐHTĐ, Quyết định số 217/QĐ-ĐHTĐ ban hành Quy chế quản trị tài sản trí tuệ năm 2024, Quy chế chi tiêu nội bộ "
     "và Điều lệ Quỹ Học bổng sau tiến sĩ Ngô Xuân Độ, với năm quy định khác nhau về phân chia lợi ích, dẫn đến sự chồng "
     "chéo trong áp dụng thực tế. Quy chế ban hành kèm Quyết định 217 "
     "đã mở rộng phạm vi tài sản tới giải pháp hữu ích, cơ sở dữ liệu, bí mật thương mại, giáo trình điện tử và nhãn hiệu "
     "của các đơn vị thuộc Trường, nhưng không dẫn chiếu và không thay thế Chương VI Quyết định 213, chưa có biểu mẫu "
     "để nhận diện các tài sản này trong thực tế và chưa có nội dung về liêm chính khoa học, liêm chính học thuật mà "
     "Thông tư số 83/2026/TT-BGDĐT xếp cùng nhóm nội dung quản trị bắt buộc về sở hữu trí tuệ."),
    ("Về phạm vi bảo hộ: Bổ sung đầy đủ các đối tượng bảo hộ vào phạm vi điều chỉnh, bao gồm:",
     "Về phạm vi bảo hộ: Thống nhất một danh mục đối tượng duy nhất trên cơ sở Điều 3 Quy chế ban hành kèm Quyết định "
     "217, bao gồm:"),
    ("trường hợp không có thỏa thuận thì áp dụng mức thù lao tối thiểu 30% lợi nhuận thuần theo quy định tại khoản 7 Điều "
     "71 Luật Khoa học, Công nghệ và Đổi mới sáng tạo số 93/2025/QH15.",
     "trường hợp không có thỏa thuận thì mức thù lao không thấp hơn mức quy định tại khoản 1 Điều 135 Luật Sở hữu trí tuệ "
     "đã được sửa đổi theo điểm b khoản 7 Điều 71 Luật số 93/2025/QH15, tức 10% lợi nhuận trước thuế khi Nhà trường tự sử "
     "dụng và 15% số tiền nhận được mỗi lần khi chuyển giao quyền sử dụng."),
    ("Bãi bỏ mức trần thù lao 100 triệu đồng tại Điều 36 Quyết định số 213/QĐ-ĐHTĐ.",
     "Bãi bỏ mức trần thù lao 100 triệu đồng tại điểm a khoản 4 Điều 36 Quyết định số 213/QĐ-ĐHTĐ."),
    ("Điều lệ Quỹ Ngô Xuân Đỗ", "Điều lệ Quỹ Học bổng sau tiến sĩ Ngô Xuân Độ"),
    # --- tên đơn vị theo danh sách nhân sự 2026 và Quyết định 217
    ("Phòng Quản lý Khoa học và Công nghệ", "Phòng Khoa học Công nghệ", "tat_ca"),
    ("Phòng Quản lýKhoa học và Công nghệ", "Phòng Khoa học\nCông nghệ"),
    ("Phòng Kế hoạch Tài chính", "Phòng Tài chính - Kế toán", "tat_ca"),
    ("Phòng Kế hoạchTài chính", "Phòng Tài chính -\nKế toán"),
    ("Phòng Tổ chức Cán bộ", "Bộ phận Hành chính - Nhân sự"),
    # --- 3.2.2
    ("Sự phân tán này dẫn đến thiếu đầu mối điều phối thống nhất",
     "Sự phân tán này, dù Điều 11 Quy chế ban hành kèm Quyết định số 217/QĐ-ĐHTĐ đã giao Phòng Khoa học Công nghệ vai "
     "trò quản lý chung, dẫn đến thiếu đầu mối điều phối thống nhất"),
    ("đáp ứng yêu cầu tại điểm d khoản 4 Quyết định số 1624/QĐ-TTg.",
     "đáp ứng định hướng tại điểm a khoản 6 " + QD1068 + "."),
    # --- 3.2.3
    ("theo đúng yêu cầu tại điểm b khoản 4 Quyết định số 1624/QĐ-TTg.",
     "phù hợp với yêu cầu tại điểm b khoản 4 " + QD1068 + "."),
    ("trước hoặc cùng thời điểm công bố bài báo khoa học để tránh mất tính mới.",
     "trước hoặc cùng thời điểm công bố bài báo khoa học để tránh mất tính mới, phù hợp với định hướng đăng ký bảo hộ "
     "đồng thời với công bố tại điểm b khoản 4 " + QD1068 + "."),
    # --- 3.2.4 (Quyết định 213, Điều 35, 38 và Chương VII; Luật 125, điểm d khoản 3 Điều 28)
    ("trong khi giảng viên phải tự bỏ kinh phí và chờ đợi",
     "trong khi Điều 38 Quyết định 213 không có mục chi riêng cho lệ phí đăng ký và giảng viên phải chờ đợi"),
    ("nằm trong Quỹ Phát triển khoa học công nghệ của Trường.",
     "nằm trong Quỹ Phát triển khoa học công nghệ của Trường, được hình thành trên cơ sở Quỹ nghiên cứu khoa học hiện có "
     "tại Chương VII Quyết định 213 và theo trách nhiệm thành lập Quỹ phát triển khoa học và công nghệ tại điểm d khoản 3 "
     "Điều 28 Luật Giáo dục đại học số 125/2025/QH15."),
    ("cho phép chi trả ngay 50% định mức thưởng giờ nghiên cứu khoa học khi",
     "cho phép ghi nhận ngay 50% định mức giờ nghiên cứu khoa học quy đổi khi"),
    ("Phần 50% còn lại chi trả khi văn bằng", "Phần 50% còn lại được ghi nhận khi văn bằng"),
    # --- 3.2.5: Chương 2 không có khảo sát nhận thức
    ("Kết quả khảo sát thực trạng tại Chương 2 cho thấy nhận thức về sở hữu trí tuệ trong cộng đồng giảng viên và sinh viên "
     "Trường Đại học Thành Đô còn rất hạn chế. Số lượng giảng viên quan tâm và chủ động tìm hiểu về đăng ký bảo hộ sở hữu "
     "trí tuệ chỉ chiếm tỷ lệ nhỏ. Phần lớn giảng viên chưa phân biệt được sự khác nhau giữa sáng chế, giải pháp hữu ích và "
     "quyền tác giả; chưa nhận thức được rằng giáo trình, bài giảng số, phần mềm ứng dụng và cơ sở dữ liệu sưu tập đều là "
     "các đối tượng có thể được bảo hộ.",
     "Chương 2 không sử dụng khảo sát nhận thức, nhưng các dữ kiện hành vi cho thấy năng lực nhận diện tài sản trí tuệ "
     "trong đội ngũ còn hạn chế: 8 đề tài có sản phẩm đủ điều kiện xác lập quyền trong giai đoạn 2021 - 2024 đều chưa "
     "được nộp đơn; hai đề tài năm 2025 ghi dự kiến đăng ký giải pháp hữu ích nhưng chưa có đơn; 87 giáo trình chưa được "
     "đăng ký quyền tác giả; Kế hoạch 07/KH-ĐHTĐ chỉ thống kê 12 lượt tập huấn chung trong giai đoạn 2019 - 2023 và không "
     "tách riêng tập huấn về sở hữu trí tuệ."),
    ("đưa sở hữu trí tuệ thành môn học bắt buộc cho sinh viên theo Quyết định số 1624/QĐ-TTg.",
     "chủ động đón đầu định hướng nghiên cứu đưa sở hữu trí tuệ thành nội dung học bắt buộc tại điểm b khoản 8 "
     + QD1068 + "."),
    # --- 3.3: thí điểm là đề xuất, chưa có kết quả thực tế
    ("áp dụng đối với các đề tài cấp cơ sở nghiệm thu trong kỳ 1 năm học 2025 - 2026 tại Viện Nghiên cứu giáo dục và "
     "Chuyển giao tri thức.",
     "áp dụng đối với các đề tài cấp cơ sở nghiệm thu năm 2026 tại Viện Y - Dược, đơn vị liên quan đến 10 trên 11 sản phẩm "
     "đủ điều kiện xác lập quyền giai đoạn 2021 - 2025 và là nơi phát sinh cả hai đơn sáng chế của Nhà trường; Viện "
     "Nghiên cứu giáo dục và Chuyển giao tri thức tham gia với vai trò hỗ trợ khai thác."),
    ("Kết quả mô phỏng trên đề tài chiết xuất dược liệu năm 2026 cho thấy khâu rà soát bắt buộc giúp kịp thời nhận diện "
     "quy trình chiết tách có tính mới, kích hoạt ngay bước tra cứu sáng chế tiền kiểm và hỗ trợ chủ nhiệm đề tài hoàn "
     "thiện hồ sơ nộp đơn đăng ký trước khi công bố bài báo khoa học. Nếu không có khâu rà soát bắt buộc này, chủ nhiệm "
     "đề tài đã công bố bài báo trước, dẫn đến bộc lộ công khai quy trình chiết tách và mất hoàn toàn khả năng đăng ký "
     "sáng chế.",
     "Hai trường hợp thực tế minh họa cho kịch bản này là đề tài chiết xuất lá Quế hoa năm 2025, đề tài cấp cơ sở duy "
     "nhất đã nộp đơn sáng chế ngay trong năm nghiệm thu, và đơn sáng chế năm 2026 về phương pháp chiết tách hợp chất từ "
     "cây Piper aduncum L. Hai trường hợp cho thấy khi chủ nhiệm đề tài chủ động, kết quả chiết tách có thể được nộp đơn "
     "trước khi công bố. Khâu rà soát bắt buộc nhằm biến cách làm của hai trường hợp đơn lẻ này thành quy trình áp dụng "
     "cho mọi đề tài, trong đó có 8 đề tài giai đoạn 2021 - 2024 đã không được nộp đơn."),
    ("Bài học kinh nghiệm rút ra từ việc thí điểm:", "Các giả định cần được kiểm chứng qua thí điểm:"),
    ("Thứ hai, thời gian bổ sung cho mỗi phiên nghiệm thu chỉ tăng thêm khoảng 15 đến 20 phút, hoàn toàn khả thi trong "
     "thực tế.",
     "Thứ hai, thời gian bổ sung cho mỗi phiên nghiệm thu được kỳ vọng không đáng kể và cần được đo lường trong thí "
     "điểm."),
    # --- 3.4
    ("Thực hiện thí điểm mô hình rà soát bộc lộ bắt buộc tại Viện Nghiên cứu giáo dục và Chuyển giao tri thức.",
     "Thực hiện thí điểm mô hình rà soát bộc lộ bắt buộc tại Viện Y - Dược."),
    ("Chuẩn bị nội dung tham gia Đề án gắn sở hữu trí tuệ với giáo dục đại học do Bộ Giáo dục và Đào tạo dự thảo vào quý "
     "2 năm 2027.",
     "Chuẩn bị nội dung tham gia Đề án tăng cường hoạt động sở hữu trí tuệ của cơ sở giáo dục đại học do Bộ Giáo dục và "
     "Đào tạo dự thảo trong quý II năm 2027."),
    ("Kết nối Nhà trường với Mạng lưới Trung tâm Hỗ trợ công nghệ và đổi mới sáng tạo.",
     "Khai thác thường xuyên tư cách thành viên Mạng lưới Trung tâm Hỗ trợ công nghệ và đổi mới sáng tạo mà Nhà trường có "
     "từ năm 2023."),
    ("theo yêu cầu tại điểm b khoản 4 Quyết định số 1624/QĐ-TTg.",
     "theo yêu cầu tại điểm b khoản 4 " + QD1068 + ". Chỉ tiêu văn bằng được đặt bằng hoặc cao hơn mức 3 văn bằng mỗi năm "
     "cho giai đoạn 2026 - 2028 tại Kế hoạch 07/KH-ĐHTĐ và cần được tính riêng cho văn bằng hình thành từ kết quả nghiên "
     "cứu, tránh lặp lại tình trạng chỉ tiêu gộp được hoàn thành bằng tài sản thương hiệu và hợp tác doanh nghiệp như "
     "phân tích tại Hình 2.9 Chương 2."),
    ("100% văn bản thống nhấtTối thiểu 30% lợi nhuậnthuần (theo Luật số93/2025/QH15)",
     "100% văn bản thống nhất\nKhông thấp hơn mức mặc định\ntại khoản 1 Điều 135\nLuật Sở hữu trí tuệ"),
]

if __name__ == "__main__":
    n = ap_dung(VAO, RA, SUA, author="Claude", merge_runs=MERGE if os.path.exists(MERGE) else None)
    print("Đã ghi:", RA, "|", n, "chỉnh sửa")

# -*- coding: utf-8 -*-
"""Lập báo cáo rà soát tính chính xác và liêm chính học thuật cho bản toàn văn tải lên ngày 04/10/2026.

Nguồn đối chiếu: thư mục VBPL, Tai lieu thanh do, PDF tài liệu tham khảo, bộ dữ liệu đã chuẩn hóa
Du_lieu_bieu_do_Chuong_2.xlsx. Không đưa dữ liệu cá nhân vào báo cáo.
"""
import os

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

GOC = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
RA = os.path.join(GOC, "Ban_cuoi", "Ban_hoan_thien_03-10-2026", "Bao_cao_ra_soat_liem_chinh_toan_van.docx")

MUC = {
    "A": "A - Phải sửa trước khi nộp",
    "B": "B - Cần sửa",
    "C": "C - Sai sót nhỏ",
}

# (vị trí, nội dung trong báo cáo, đối chiếu nguồn gốc, mức, đề xuất)
NHOM = [
    ("1. Khai báo phương pháp, dữ liệu và kết luận vượt bằng chứng", [
        ("Mở đầu, mục 7",
         "Nêu phương pháp chuyên gia: tham vấn chuyên gia quản lý khoa học, chuyên gia pháp lý và nhà khoa học đầu ngành.",
         "Đề tài không tổ chức tham vấn chuyên gia. Khai một phương pháp không thực hiện là lỗi liêm chính.",
         "A", "Bỏ phương pháp chuyên gia, hoặc ghi rõ phương pháp này chưa thực hiện và nêu thành giới hạn nghiên cứu."),
        ("Mở đầu, mục 5",
         "Phạm vi thời gian: số liệu thứ cấp và sơ cấp, giai đoạn từ năm 2020 đến năm 2025.",
         "Đề tài không thu thập dữ liệu sơ cấp. Toàn bộ phân tích ở Chương 2 dùng giai đoạn 2021 - 2025.",
         "A", "Sửa thành: số liệu thứ cấp giai đoạn 2021 - 2025. Thay các chữ khảo sát ở Chương 2 bằng rà soát hoặc tổng hợp dữ liệu hành chính."),
        ("Trang đầu",
         "Mã số đề tài: ĐTCS-2026-SHTT.",
         "Không có trong thuyết minh hay văn bản giao đề tài trong hồ sơ.",
         "B", "Đối chiếu với quyết định giao đề tài."),
        ("Mục 2.3.3, Hình 2.11, Tiểu kết Chương 2, Kết luận",
         "Khẳng định giảng viên đã vội vàng công bố bài báo, làm mất tính mới tuyệt đối của giải pháp; nhóm nghiên cứu đã làm sáng tỏ chuỗi nguyên nhân gốc rễ.",
         "Tên Hình 2.11 ghi là giả thuyết. Hồ sơ không có dữ liệu cho thấy sản phẩm nào trong 8 sản phẩm đã bị bộc lộ qua bài báo. Đây là giả thuyết chưa kiểm chứng nhưng được viết như phát hiện.",
         "A", "Viết thành giả thuyết hoặc nguy cơ cần kiểm tra; đưa việc rà soát tình trạng bộc lộ thành nhiệm vụ của giải pháp."),
        ("Mục 2.2.3; 2.3.2 (ý Sáu là); Bảng 3.3",
         "Nhà trường đã ký 02 hợp đồng chuyển giao quyền sử dụng sách chuyên khảo, mỗi hợp đồng 5 triệu đồng mỗi năm, theo số liệu của Viện Nghiên cứu giáo dục và Chuyển giao tri thức.",
         "Không tìm thấy hợp đồng, sổ theo dõi hay số liệu nào về nội dung này trong thư mục tài liệu của Trường và bộ dữ liệu đã chuẩn hóa.",
         "A", "Bổ sung văn bản gốc trước khi dùng; nếu không có thì bỏ."),
        ("Mục 2.2.3",
         "Đề tài 08-2023 đã bàn giao toàn bộ quy trình và sản phẩm mẫu cho Nhà trường để thương mại hóa phục vụ quà tặng.",
         "Hồ sơ đề tài chỉ ghi sản phẩm chuyển giao cho Trường thương mại hóa; không có hợp đồng, biên bản bàn giao hay kết quả thương mại hóa.",
         "B", "Viết đúng mức hồ sơ cho phép: hồ sơ ghi chuyển giao cho Trường, chưa có biên bản hay kết quả thương mại hóa."),
        ("Mục 2.2.3",
         "Toàn bộ 61 giáo trình nghiệm thu giai đoạn 2022 - 2025 đều được đưa vào giảng dạy theo đúng số tín chỉ.",
         "Số 61 khớp Bảng 2.2, nhưng không có nguồn nào xác nhận việc đưa vào giảng dạy.",
         "B", "Bỏ phần khẳng định, hoặc dẫn nguồn của Phòng Đào tạo."),
        ("Mục 2.2.4",
         "Tỷ lệ trùng lặp tối đa 20% đến 25% tùy bậc đào tạo; học liệu trên hệ thống quản lý học tập chưa gắn chữ ký số; chưa ghi nhận khiếu nại hay tranh chấp nào.",
         "Không có văn bản nguồn trong hồ sơ. Mức 20% chỉ xuất hiện như một đề xuất trong bản nháp cũ, không phải quy định đang áp dụng.",
         "B", "Dẫn văn bản quy định của Trường hoặc bỏ."),
        ("Bảng 2.4, Bảng 2.3",
         "Thông tin Quỹ Học bổng sau tiến sĩ Ngô Xuân Độ: ngân sách 5 tỷ đồng, Điều 9 chia 50/50 rồi 20/80.",
         "Văn bản Điều lệ Quỹ không có trong hồ sơ. Các bản trước đều ghi rõ: nội dung ghi theo thông tin được cung cấp, chưa đối chiếu văn bản gốc.",
         "B", "Giữ ghi chú chưa đối chiếu văn bản gốc, hoặc bổ sung Điều lệ."),
        ("Bảng 2.6, Mục 2.3.1",
         "Đơn sáng chế 1-2025-07378 đã được Cục Sở hữu trí tuệ chấp nhận đơn hợp lệ.",
         "Sổ theo dõi có một dòng ghi chấp nhận đơn hợp lệ nhưng không có số đơn, nên không gán được cho đơn này.",
         "B", "Ghi: đã nộp đơn năm 2025; trạng thái thẩm định cần xác minh."),
    ]),
    ("2. Dẫn chiếu pháp luật", [
        ("Mở đầu mục 1; 1.1.3; 1.2.3; Bảng 1.1",
         "Luật KH,CN&ĐMST quy định mức thù lao tối thiểu bắt buộc từ 30% đến 50% cho tác giả sáng chế.",
         "Điểm a khoản 3 Điều 28 Luật số 93/2025/QH15 chỉ quy định thưởng tối thiểu 30% lợi nhuận, cho phần kết quả dùng ngân sách nhà nước; luật không có mức 50%. Đây là khoản thưởng theo Luật KH,CN&ĐMST, khác với thù lao theo Điều 135 Luật Sở hữu trí tuệ.",
         "A", "Sửa thành: thưởng tối thiểu 30% lợi nhuận đối với phần kết quả sử dụng ngân sách nhà nước. Tách riêng thưởng và thù lao."),
        ("1.2.2 mục 1; Bảng 1.1; 3.1.1",
         "Căn cứ Điều 86a Luật Sở hữu trí tuệ, quyền đăng ký đối với kết quả dùng ngân sách nhà nước được giao tự động, không bồi hoàn.",
         "Điều 86a đã bị bãi bỏ từ ngày 01/10/2025 theo điểm h khoản 7 Điều 71 Luật số 93/2025/QH15 (Văn bản hợp nhất số 67/VBHN-VPQH, chú thích 161). Căn cứ hiện hành: khoản 2 Điều 25 Luật số 93/2025/QH15 và điểm c khoản 1 Điều 86 Luật Sở hữu trí tuệ, bổ sung bởi Luật số 131/2025/QH15, hiệu lực từ 01/4/2026.",
         "A", "Thay toàn bộ dẫn chiếu Điều 86a bằng hai căn cứ trên."),
        ("1.2.1 mục 2; Bảng 1.1",
         "Căn cứ Điều 73 Luật Sở hữu trí tuệ về nghĩa vụ của tác giả thông báo kết quả sáng tạo cho nhà trường.",
         "Điều 73 Luật Sở hữu trí tuệ quy định dấu hiệu không được bảo hộ với danh nghĩa nhãn hiệu, không liên quan.",
         "A", "Bỏ dẫn chiếu. Nghĩa vụ thông báo nên đặt trong hợp đồng giao nhiệm vụ và quy chế của Trường."),
        ("1.2.4 mục 2; Bảng 1.1",
         "Điều 126 Luật Sở hữu trí tuệ quy định các hành vi xâm phạm bí mật kinh doanh.",
         "Điều 126 quy định hành vi xâm phạm quyền đối với sáng chế, kiểu dáng công nghiệp, thiết kế bố trí. Hành vi xâm phạm bí mật kinh doanh thuộc Điều 127.",
         "A", "Sửa thành Điều 127."),
        ("3.1.1; Tài liệu tham khảo số 6",
         "Nghị định số 100/2026/NĐ-CP ngày 18/02/2026 về quản lý và sử dụng Quỹ phát triển KH&CN của cơ sở giáo dục đại học, cho phép Quỹ chi trực tiếp cho tra cứu, đăng ký, định giá.",
         "Nghị định số 100/2026/NĐ-CP ban hành ngày 31/3/2026, sửa đổi Nghị định số 65/2023/NĐ-CP về sở hữu công nghiệp (chính báo cáo cũng mô tả đúng như vậy ở Mục 1.2). Tên văn bản, ngày ban hành và nội dung gán cho nghị định này đều không đúng. Căn cứ cho phép quỹ chi đăng ký, bảo hộ, khai thác quyền sở hữu trí tuệ là điểm b khoản 2 Điều 66 Luật số 93/2025/QH15.",
         "A", "Sửa mục tài liệu tham khảo; thay căn cứ bằng điểm b khoản 2 Điều 66 Luật KH,CN&ĐMST."),
        ("Bảng 2.3; 2.2.3; 3.1.1; 3.2.3",
         "Luật KH,CN&ĐMST tuyệt đối không cho phép đặt mức trần; nghĩa vụ nộp 40% ngân sách tại Quyết định 213 đã hết hiệu lực; trích thưởng tối đa 10% cho người xúc tiến thương mại hóa.",
         "Luật không có quy định cấm đặt trần. Mức trần 100 triệu đồng chỉ trái khoản 3 Điều 28 khi làm phần thưởng thấp hơn 30% lợi nhuận, tức lợi nhuận trên khoảng 333 triệu đồng, với kết quả dùng ngân sách nhà nước như đề tài NAFOSTED. Quyết định 213 là văn bản nội bộ nên không hết hiệu lực theo luật, chỉ không còn phù hợp. Luật không quy định mức tối đa 10%; khoản 2 Điều 34 Nghị định 267 quy định tổ chức trung gian, môi giới hưởng tối thiểu 10% khi các bên không có thỏa thuận.",
         "A", "Diễn đạt lại theo đúng điều kiện áp dụng; bỏ các chữ tuyệt đối, hết hiệu lực, tối đa 10%."),
        ("1.2.3 mục 3",
         "Mức thù lao tối thiểu là 10% số tiền làm lợi do tự sử dụng.",
         "Khoản 1 Điều 135 hiện hành, sửa đổi bởi Luật số 93/2025/QH15: 10% lợi nhuận trước thuế tương ứng với giá trị mà sáng chế, kiểu dáng, thiết kế bố trí đóng góp. Cách diễn đạt số tiền làm lợi là của quy định cũ.",
         "B", "Sửa theo câu chữ hiện hành."),
        ("1.2 (hai chỗ); 1.2.4",
         "Kết luận số 51-KL/TW ngày 22/5/2026; Kết luận 51 yêu cầu đăng ký sở hữu trí tuệ song song với công bố, thí điểm định giá trong trường đại học.",
         "Kết luận 51 ban hành ngày 17/6/2026; ngày 22/5/2026 là ngày phiên họp. Yêu cầu đăng ký đồng thời với công bố và thí điểm hỗ trợ định giá ít nhất 100 quyền sở hữu trí tuệ nằm trong Quyết định 1624, không có trong Kết luận 51.",
         "B", "Sửa ngày; gán đúng nội dung cho Quyết định 1624."),
        ("1.2.4 (vai trò)",
         "Liêm chính học thuật theo khoản 4 Điều 2 Luật Giáo dục đại học.",
         "Điều 2 Luật số 125/2025/QH15 chỉ có hai khoản về đối tượng áp dụng; khái niệm liêm chính học thuật ở khoản 4 Điều 3.",
         "B", "Sửa thành khoản 4 Điều 3."),
        ("Kết luận, kiến nghị với Bộ GD&ĐT",
         "Cơ chế góp vốn, thành lập doanh nghiệp khởi nguồn theo Điều 30 Luật Giáo dục đại học.",
         "Điều 30 quy định quyền và nghĩa vụ của giảng viên. Quyền thành lập doanh nghiệp khoa học và công nghệ, doanh nghiệp quản lý tài sản trí tuệ ở khoản 1 Điều 28.",
         "B", "Sửa thành khoản 1 Điều 28."),
        ("1.2.3 mục 2",
         "Điều 58 Luật KH,CN&ĐMST cho phép giảng viên góp vốn, quản lý doanh nghiệp; Điều 42 Luật Giáo dục đại học quy định quản lý, định giá tài sản trí tuệ khi góp vốn.",
         "Điều 58 áp dụng cho viên chức tại tổ chức khoa học và công nghệ công lập, không áp dụng cho giảng viên trường tư thục. Điều 42 quy định tài sản hợp nhất không phân chia và tài sản của nhà đầu tư, không có quy định riêng về định giá tài sản trí tuệ.",
         "B", "Bỏ hai dẫn chiếu, hoặc nêu đúng phạm vi áp dụng."),
        ("1.2; 1.2.2; 2.2.4; 3.1.1",
         "Dùng Thông tư 01/2024/TT-BGDĐT như chuẩn hiện hành.",
         "Thông tư 83/2026/TT-BGDĐT có hiệu lực từ 15/11/2026 và thay thế Thông tư 01/2024; báo cáo nghiệm thu sau thời điểm này.",
         "B", "Thay bằng Thông tư 83."),
        ("Bảng 3.1 (T2, O3); 3.1.1",
         "Ngưỡng 0,3 bài trên giảng viên; sáng chế gấp 5 lần, giải pháp hữu ích gấp 3 lần bài báo.",
         "Thông tư 83: 0,3 sản phẩm quy đổi trên giảng viên mỗi năm; cơ sở có đào tạo tiến sĩ là 0,6, trong đó bài WoS hoặc Scopus không thấp hơn 0,3. Trong công thức P = P1 + 2P2 + 3P3 + 5P4, bài trong nước tính 1 và bài WoS, Scopus tính 2.",
         "B", "Nêu đúng đơn vị sản phẩm quy đổi và mốc so sánh."),
        ("Tài liệu tham khảo 13, 14; Bảng 3.1 (O4)",
         "Chỉ thị 02 về tăng cường thực thi quyền sở hữu trí tuệ và xây dựng cơ sở dữ liệu quốc gia về tài sản trí tuệ; Nền tảng số sở hữu trí tuệ quốc gia, Trung tâm TISC theo Chỉ thị 02; Quyết định 1624 về phát triển tài sản trí tuệ.",
         "Tên đúng của Chỉ thị 02 chỉ là: Về tăng cường thực thi quyền sở hữu trí tuệ; Chỉ thị nói đến cơ sở dữ liệu quốc gia về thực thi quyền sở hữu trí tuệ, không nhắc TISC. Tên đúng của Quyết định 1624: sửa đổi, bổ sung một số điều của Quyết định 1068 phê duyệt Chiến lược sở hữu trí tuệ đến năm 2030.",
         "B", "Sửa tên văn bản và nội dung dẫn."),
        ("1.2.3 (Điều 144, 148)",
         "Điều 144 cấm điều khoản ấn định giá bán; hợp đồng chuyển quyền sử dụng chỉ có giá trị với bên thứ ba khi đăng ký.",
         "Khoản 2 Điều 144 không liệt kê ấn định giá bán. Khoản 3 Điều 148 loại trừ hợp đồng sử dụng nhãn hiệu.",
         "C", "Sửa ví dụ cho khớp khoản 2 Điều 144; bổ sung ngoại lệ nhãn hiệu."),
    ]),
    ("3. Số liệu của Trường", [
        ("Hình 2.4 và đoạn phân tích",
         "Kinh phí theo năm: 2022 cấp 80 triệu đồng; 2023 cấp 89,75 triệu đồng cho 4 đề tài; 2024 cấp 110 triệu đồng cho 6 đề tài; 2025 cấp 145 triệu đồng cho 7 đề tài.",
         "Dữ liệu đã chuẩn hóa từ danh mục đề tài: 2022 có 3 đề tài, 84 triệu đồng; 2023 có 7 đề tài, 70 triệu đồng; 2024 có 5 đề tài và 3 đề tài tự tìm tài trợ, 125,75 triệu đồng; 2025 có 4 đề tài và 3 đề tài chỉ quy đổi giờ, 145 triệu đồng. Các số trong báo cáo cộng lại vẫn ra 424,75 triệu đồng nhưng không khớp danh mục, cũng không khớp cơ cấu 19/16/3 ở Bảng 2.4.",
         "A", "Thay bằng số liệu đã chuẩn hóa."),
        ("Hình 2.10; 2.3.1; Bảng 3.1 (S2)",
         "Tham luận cấp Trường đạt 67/30 bài; đề tài cấp cơ sở 15/16; đề tài cấp Tỉnh, Bộ, Nhà nước 3/4, tức 75%; giáo trình 18/12, tức 150%. Phần lời viết 6 trên 9 chỉ tiêu đạt.",
         "Theo mục 2.2.4 Kế hoạch 07 và dữ liệu đã chuẩn hóa: tham luận cấp Trường 134/60; đề tài cấp cơ sở 15/14, tức 107%; đề tài cấp Bộ, Nhà nước 3/2, tức 150%; giáo trình 18/20, tức 90%. Ngay trong Hình 2.10 của báo cáo chỉ có 5 chỉ tiêu đạt, mâu thuẫn với phần lời.",
         "A", "Thay bằng số liệu theo Kế hoạch 07."),
        ("Hình 2.5 và đoạn phân tích",
         "Giải pháp hữu ích 240 giờ, tương đương bài Q2; Q2 240 giờ; Q3, Q4 180 giờ và thưởng 10 triệu đồng; bài trong nước 60 giờ.",
         "Quy chế chi tiêu nội bộ 2026: giải pháp hữu ích 180 giờ; Q1 300, Q2 270, Q3 hoặc Q4 240 giờ; bài trong nước từ 30 đến 150 giờ tùy điểm tạp chí; thưởng WoS Q3 12 triệu đồng, Q4 10 triệu đồng. Sáng chế 600 giờ (chuẩn quốc tế), 360 giờ (chuẩn Việt Nam) và thưởng Q1 20 triệu đồng là đúng.",
         "A", "Sửa số liệu; bỏ câu giải pháp hữu ích tương đương bài Q2."),
        ("2.2.1; 2.3.1; Bảng 3.1 (S3); 3.3.2; Tiểu kết Chương 2",
         "10 trên 11 đề tài có sản phẩm tiềm năng thuộc Khoa Dược và Viện Y - Dược; Tiểu kết ghi Viện Y - Dược có 11 đề tài.",
         "Theo chính Bảng 2.7 và danh mục đề tài: 9 trên 11 thuộc Viện Y - Dược; 1 thuộc Viện Quản trị và Công nghệ (37-2022), 1 thuộc Viện Nghiên cứu giáo dục và Chuyển giao tri thức (08-2023).",
         "A", "Sửa thành 9 trên 11 ở mọi chỗ."),
        ("Bảng 2.7",
         "Đơn vị chủ trì ghi Khoa Công nghệ kỹ thuật Ô tô, Khoa Dược.",
         "Danh mục đề tài ghi đơn vị quy đổi là Viện Quản trị và Công nghệ (37-2022) và Viện Y - Dược (các đề tài dược).",
         "B", "Ghi theo danh mục, hoặc ghi rõ khoa thuộc viện nào."),
        ("Bảng 2.5",
         "Sáng chế hợp chất chiết xuất từ lá Quế hoa: đã nộp 01 đơn năm 2025 và 01 đơn năm 2026.",
         "Đơn năm 2026 (1-2026-07185) là phương pháp chiết tách hợp chất Isoembigenin từ cây Piper aduncum L., không phải lá Quế hoa.",
         "C", "Tách thành hai đối tượng."),
        ("Mục 2.1.1",
         "Trường Cao đẳng thành lập theo Quyết định số 762/QĐ-BGD&ĐT ngày 19/02/2004; triết lý Trí - Năng - Nhân - Hòa.",
         "Chiến lược của Trường ghi Trường Cao đẳng Công nghệ Thành Đô thành lập ngày 30/11/2004, nâng cấp thành Trường Đại học ngày 27/5/2009. Khung chiến lược Edupark ghi triết lý Trí - Năng - Hòa - Nhân. Số hiệu hai quyết định chưa có văn bản để đối chiếu.",
         "B", "Sửa ngày và thứ tự triết lý; kiểm tra số quyết định."),
        ("Hình 2.3; giải pháp 3.2.3, 3.2.4",
         "Đầu mối Trung tâm Đào tạo và Chuyển giao công nghệ, Viện Quản trị và Sáng tạo số; Trung tâm Công nghệ thông tin chủ trì hoặc phối hợp.",
         "Các đơn vị này không có trong cơ cấu tổ chức ở Bảng 2.1 và danh sách nhân sự năm 2026.",
         "B", "Kiểm tra tên đơn vị; dùng đúng tên trong quy chế tổ chức."),
    ]),
    ("4. Trích dẫn và tài liệu tham khảo", [
        ("Mở đầu (danh mục đầu) và Tài liệu tham khảo cuối",
         "Cùng một tài liệu được ghi hai thông tin xuất bản khác nhau, ví dụ Maresova et al. (2019): Administrative Sciences 9(3) ở đầu, Economies 7(4) ở cuối.",
         "Đối chiếu bản PDF và tệp Zotero: danh mục ở cuối sai thông tin xuất bản của 15 tài liệu (chi tiết ở Bảng 2). Thông tin xuất bản sai khiến người đọc không truy xuất được tài liệu, nên là lỗi liêm chính.",
         "A", "Sửa thông tin xuất bản theo Bảng 2."),
        ("Mở đầu 2.2; danh mục số 8",
         "Rialti, Marzi, Caputo & Mayah (2022), Establishing successful university-industry collaborations, DOI 10.1007/s10961-022-09932-2.",
         "Bản PDF có trong thư mục ghi tác giả là Michele O'Dwyer, Raffaele Filieri và Lisa O'Malley; The Journal of Technology Transfer (2023) 48: 900-931. Báo cáo gán bài cho sai tác giả, và danh mục cuối lại ghi sai tạp chí (Technovation).",
         "A", "Sửa thành O'Dwyer et al. (2023) ở mọi chỗ."),
        ("Chương 1",
         "Trích Milliken và Allen (2013), Tewari và Bhardwaj (2020); dẫn Nghị định 17/2023, Luật Giá, Bộ luật Hình sự.",
         "Các tài liệu này không có trong danh mục. Milliken và Allen là bản tóm tắt 5 trang của cơ quan nghiên cứu thuộc Quốc hội Bắc Ireland, nguồn yếu cho một luận điểm lý luận. Sách của Tewari và Bhardwaj là ấn bản năm 2021 của Đại học Panjab.",
         "B", "Bổ sung vào danh mục với đúng năm, hoặc thay bằng nguồn đã có; nên bỏ Milliken và Allen."),
        ("Tài liệu tham khảo",
         "Có Bradley (2013), Etzkowitz (2003), Goldfarb và Henrekson (2003), Shane (2004), Siegel et al. (2007), O'Dwyer et al. (2023) trong danh mục.",
         "Sáu tài liệu này không được trích trong thân bài.",
         "B", "Trích ở đúng chỗ hoặc bỏ khỏi danh mục."),
        ("1.2.1; 2.2 Mở đầu; 1.1.2; 1.2.2; 1.1.3",
         "Teece (2018) dùng để định nghĩa sáng tạo trong trường đại học và cho đề xuất về văn phòng chuyển giao ở nước đang phát triển; Perkmann et al. (2013) dùng cho việc mất tính mới khi công bố trước; Rocha et al. (2023) dùng để định nghĩa xác lập quyền; Thursby và Kemp (2002) dùng cho nguồn thu.",
         "Teece (2018) phân tích mô hình cấp phép công nghệ trong ngành viễn thông không dây; Perkmann et al. (2013) tổng quan về tương tác học thuật và thương mại hóa, không bàn về tính mới sáng chế; Rocha et al. (2023) đánh giá bản khai báo sáng chế; Thursby và Kemp (2002) đo hiệu quả cấp phép. Nội dung được gán không có trong các nguồn này.",
         "B", "Gán lại cho đúng nguồn (về tính mới: Điều 60 Luật Sở hữu trí tuệ) hoặc bỏ trích dẫn."),
        ("1.1.1",
         "Fisher (2001) và Guan (2014) chỉ ra tài sản trí tuệ là sự kết hợp hữu cơ giữa tính trừu tượng của ý tưởng và hình thức biểu hiện vật chất.",
         "Chương của Guan (2014) có trong thư mục mở đầu bằng nhận định không có định nghĩa thống nhất về sở hữu trí tuệ; không thấy luận điểm như báo cáo nêu.",
         "B", "Diễn đạt lại đúng luận điểm của nguồn."),
        ("Mở đầu 2.1, 2.2",
         "Maresova et al. tổng quan 22 công trình, phân loại theo bốn trụ cột; Öztürk khảo sát 96 chuyên gia; Phạm và Nguyễn (2018) nêu tỷ lệ 50% - 70% doanh thu cấp phép cho tác giả; Lê và Nguyễn (2019) đề xuất mô hình 5 bước.",
         "Thư mục không có bản gốc của các bài này nên chưa kiểm chứng được các con số. Danh mục cuối còn ghi Phạm và Nguyễn (2018) về Đại học Bách khoa Hà Nội, trái với chính Mở đầu ghi Đại học Thanh Hoa.",
         "B", "Kiểm tra bản gốc; chưa kiểm chứng được thì bỏ con số."),
        ("Mở đầu, Chương 1",
         "Viết đầy đủ tên tác giả Việt: Lê Thị Thu Hà và Nguyễn Thành Khang; Phạm Thúy Hằng; Lê Thị Thanh Tâm và Hoàng Đình Thái; Phạm Thị Thúy Hằng và Nguyễn Thanh Hùng.",
         "Danh mục chỉ có chữ viết tắt (Lê, T. T. H.; Phạm, T. H.); thư mục không có bản gốc. Tên đầy đủ có thể là suy đoán.",
         "B", "Dùng họ và năm theo APA, ví dụ (Lê & Nguyễn, 2019), trừ khi đã kiểm tra tên trên bài gốc."),
    ]),
    ("5. Dẫn chiếu nội bộ", [
        ("2.2.4; 3.4",
         "Dẫn chiếu bộ 16 chỉ tiêu đã xác lập tại Mục 1.4.1 ở Chương 1.",
         "Báo cáo không có Mục 1.4.1; Chương 1 chỉ nêu bốn nhóm tiêu chí, không có danh sách 16 chỉ tiêu được đánh số như Hình 2.6 sử dụng.",
         "B", "Bổ sung danh sách 16 chỉ tiêu ở Chương 1 hoặc sửa dẫn chiếu."),
    ]),
]

# Bảng 2: thông tin xuất bản sai trong danh mục cuối, so với bản đã kiểm chứng
TLTK = [
    ("Nguyễn, M. H. T. (2025)", "Tạp chí Kinh tế và Quản lý, 48(2), 62-71; tên bài đổi thành bối cảnh tự chủ",
     "Tạp chí Khoa học Trường Đại học Sư phạm TP Hồ Chí Minh, 22(1), 123-131; bối cảnh cuộc cách mạng công nghiệp 4.0 (bản PDF)"),
    ("Võ, N. H. P. (2025)", "Tạp chí Pháp luật và Thực tiễn, 14(4), 38-49 (danh mục đầu ghi Đại học Mở TP Hồ Chí Minh)",
     "Tạp chí Khoa học Trường Đại học Mở Hà Nội, 75-86, DOI 10.59266/houjs.2025.606 (bản PDF)"),
    ("O'Dwyer et al. (2023)", "Technovation, 120, 102627", "The Journal of Technology Transfer, 48(3), 900-931 (bản PDF)"),
    ("Guan, W. (2014)", "pp. 23-58", "Chương 1, pp. 1-10, DOI 10.1007/978-3-642-55265-6_1 (bản PDF)"),
    ("Lê, T. T. H., & Nguyễn, T. K. (2019)", "Tạp chí Khoa học ĐHQG Hà Nội: Nghiên cứu Chính sách và Quản lý, 35(3), 15-26",
     "Tạp chí Khoa học Đại học Văn Lang, 18, 27-36"),
    ("Lê, T. T. T., & Hoàng, Đ. T. (2021)", "29(1), 45-54", "Tạp chí Khoa học Quản lý Giáo dục, 2(30), 38-45"),
    ("Phạm, T. T. H., & Nguyễn, T. H. (2018)", "Bài học từ Đại học Bách khoa Hà Nội; Tạp chí Khoa học và Công nghệ Việt Nam, 60(9)",
     "Bài học từ Đại học Thanh Hoa, Trung Quốc; Tạp chí Khoa học và Giáo dục, Trường ĐHSP Huế, 3(47), 84-94"),
    ("Cục Sở hữu trí tuệ & Bộ KH&CN (2024)", "Kỷ yếu hội thảo khoa học quốc gia", "Nhà xuất bản Khoa học và Kỹ thuật"),
    ("Cục Sở hữu trí tuệ, tài liệu tập huấn", "Năm 2024, Trung tâm Đào tạo Sở hữu trí tuệ", "Không ghi năm (n.d.), kèm địa chỉ truy cập"),
    ("Bstieler et al. (2015)", "Tên bài rút gọn; trang 100-112", "... US biotechnology industry: IP policies, shared governance, and champions; 32(1), 111-121"),
    ("Bulsara & Vaghela (2025)", "International Journal of Innovation Studies, 9(2), 145-159", "Tech Monitor, WIPO, 38-43"),
    ("Holgersson (2021)", "European Journal of Innovation Management, 24(5), 1421-1440", "Báo cáo của European Patent Office / Chalmers University"),
    ("Holgersson & Aaboen (2019)", "The Journal of Technology Transfer, 44(3), 856-887",
     "Technology in Society, 59, Article 101132"),
    ("Maresova et al. (2019)", "Economies, 7(4), 105", "Administrative Sciences, 9(3), Article 67"),
    ("Öztürk (2026)", "Higher Education Policy, 39(1), 89-112", "Scientific Culture, 12(3.1), 287-303"),
    ("Pandey et al. (2025)", "Tên bài ghi Higher Education Institutions; 30(1), 45-56", "... in Academic Institutions; 30, 640-649"),
    ("Rocha et al. (2023)", "Technological Forecasting and Social Change, 188, 122289",
     "International Journal of Entrepreneurship and Innovation Management, 27(1/2), 119-136"),
]

DUNG = [
    "Bảng 2.1 khớp từng dòng với danh sách nhân sự ở dạng tổng hợp (252 nhân sự, 145 giảng viên, 94 tiến sĩ, 25 GS và PGS); các tỷ lệ 37,3%, 57,9%, 86,2%, 88,0%; Phòng KHCN có 02 nhân sự trình độ thạc sĩ; số ngành của ba viện.",
    "Bảng 2.2 khớp từng ô; các phép tính 3,43 lần, 36,1%/năm, 7 lần, 62,7%/năm, tỷ trọng 41,1% và 83,9%; tổng kinh phí 424,75 triệu đồng và cơ cấu 19/16/3; 396,75 triệu đồng (93,4%); 71 bài quốc tế có phân hạng; hệ số Gini 0,836 và tỷ lệ 38,3%; ba đề tài NAFOSTED 4,67 tỷ đồng và ngày phê duyệt.",
    "Bảng 2.7: mã số, sản phẩm, nhóm quyền và ngày nghiệm thu khớp danh mục đề tài (trừ cột đơn vị, xem phát hiện số 30).",
    "Bảng 2.6 về 12 hồ sơ tài sản trí tuệ (số đơn, số văn bằng, chủ sở hữu); chuỗi 38 - 9 - 1 và nhóm 31 đề tài 2021 - 2024 không có đơn ở Hình 2.9.",
    "Nội dung điểm a, điểm b khoản 4 Điều 36 Quyết định 213 (40/30/30, trần 100 triệu đồng; 30/20/50), Điều 13 Quyết định 217 và Bảng 7 Quy chế chi tiêu nội bộ ở Bảng 2.3; phép mô phỏng ngưỡng 333 triệu đồng.",
    "Khoản 1 Điều 135 ở Bảng 2.3 và Chương 3 (10% lợi nhuận trước thuế; 15% tổng số tiền nhận được trước thuế); Điều 6, Điều 60, Điều 90, Điều 93, khoản 1 và khoản 2 Điều 148 Luật Sở hữu trí tuệ.",
    "Khoản 1, điểm d và điểm đ khoản 2, điểm a và điểm d khoản 3 Điều 28 Luật Giáo dục đại học; khoản 2 Điều 34 Nghị định 267; Điều 5a Nghị định 134; mốc công khai trước 31/5 và yêu cầu cập nhật HEMIS của Thông tư 83; nội dung Quyết định 1624.",
]


def phong(run, co=12, dam=False, nghieng=False):
    run.font.name = "Times New Roman"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    run.font.size = Pt(co)
    run.bold = dam
    run.italic = nghieng


def doan(doc, s, co=13, dam=False, can=None, truoc=0, sau=6, nghieng=False):
    p = doc.add_paragraph()
    phong(p.add_run(s), co, dam, nghieng)
    pf = p.paragraph_format
    pf.space_before = Pt(truoc)
    pf.space_after = Pt(sau)
    pf.line_spacing = 1.15
    if can:
        p.alignment = can
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    return p


def to_nen(cell, mau):
    tcpr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), mau)
    tcpr.append(shd)


def lap_hang_tieu_de(row):
    trpr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    el.set(qn("w:val"), "true")
    trpr.append(el)


def o(cell, s, co=10.5, dam=False, can=WD_ALIGN_PARAGRAPH.LEFT, mau=None):
    cell.text = ""
    p = cell.paragraphs[0]
    r = p.add_run(s)
    phong(r, co, dam)
    if mau:
        r.font.color.rgb = RGBColor.from_string(mau)
    p.alignment = can
    p.paragraph_format.space_after = Pt(2)


def bang(doc, tieu_de, rong, hang, mau_muc=False):
    t = doc.add_table(rows=1, cols=len(tieu_de))
    t.style = "Table Grid"
    for j, s in enumerate(tieu_de):
        c = t.rows[0].cells[j]
        o(c, s, dam=True, can=WD_ALIGN_PARAGRAPH.CENTER)
        to_nen(c, "D9E2F3")
    lap_hang_tieu_de(t.rows[0])
    for h in hang:
        cells = t.add_row().cells
        for j, s in enumerate(h):
            mau = None
            if mau_muc and j == 4:
                mau = {"A": "C00000", "B": "B45F06", "C": "595959"}[s[0]]
            o(cells[j], s, dam=(mau_muc and j == 4), can=WD_ALIGN_PARAGRAPH.CENTER if j in (0, 4) and mau_muc else WD_ALIGN_PARAGRAPH.LEFT, mau=mau)
    t.autofit = False
    grid = t._tbl.tblGrid
    for j, gc in enumerate(grid.findall(qn("w:gridCol"))):
        gc.set(qn("w:w"), str(int(rong[j] * 567)))
    for row in t.rows:
        for j, w in enumerate(rong):
            row.cells[j].width = Cm(w)
    return t


def main():
    doc = Document()
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = Cm(29.7), Cm(21.0)
    for k in ("left_margin", "right_margin"):
        setattr(sec, k, Cm(2.0))
    sec.top_margin = sec.bottom_margin = Cm(2.0)

    doan(doc, "BÁO CÁO RÀ SOÁT TÍNH CHÍNH XÁC VÀ LIÊM CHÍNH HỌC THUẬT", 15, True, WD_ALIGN_PARAGRAPH.CENTER, sau=2)
    doan(doc, "Bản Báo cáo toàn văn tổng kết đề tài tải lên ngày 04/10/2026", 13, False, WD_ALIGN_PARAGRAPH.CENTER, sau=12, nghieng=True)

    doan(doc, "1. Phạm vi và cách rà soát", 13, True, truoc=6)
    doan(doc, "Bản báo cáo được đọc toàn văn, từ Mở đầu đến Tài liệu tham khảo (khoảng 36.800 chữ). Mỗi nhận định về pháp luật, số liệu "
              "và tài liệu được đối chiếu với nguồn gốc lưu trong thư mục đề tài: văn bản pháp luật trong thư mục VBPL (Văn bản hợp nhất "
              "số 67/VBHN-VPQH Luật Sở hữu trí tuệ, Luật số 93/2025/QH15, Luật số 125/2025/QH15, Nghị định 267, Nghị định 134, Thông tư 83, "
              "Kết luận 51, Quyết định 1624, Chỉ thị 02); tài liệu của Trường trong thư mục Tai lieu thanh do (Quy chế chi tiêu nội bộ 2026, "
              "Chiến lược của Trường, khung chiến lược Edupark); bộ dữ liệu đã chuẩn hóa Du_lieu_bieu_do_Chuong_2.xlsx; bản PDF của các tài "
              "liệu tham khảo hiện có. Nghị định 100/2026/NĐ-CP không có trong thư mục nên được tra cứu trên mạng. Dữ liệu cá nhân trong "
              "danh sách nhân sự chỉ được dùng ở dạng tổng hợp.")
    doan(doc, "Phạm vi: chỉ kiểm tra tính chính xác của nội dung, không đánh giá bố cục, cách trình bày hay văn phong. Mức độ: A là lỗi "
              "phải sửa trước khi nộp (sai quy định pháp luật, sai số liệu, tài liệu tham khảo sai, khai phương pháp không thực hiện); "
              "B là lỗi cần sửa; C là sai sót nhỏ.")

    tong = {m: sum(1 for _, ds in NHOM for x in ds if x[3] == m) for m in "ABC"}
    doan(doc, "2. Kết luận chung", 13, True, truoc=6)
    doan(doc, f"Có {tong['A'] + tong['B'] + tong['C']} phát hiện, trong đó {tong['A']} lỗi mức A, {tong['B']} lỗi mức B và {tong['C']} lỗi mức C. "
              "Báo cáo chưa đủ điều kiện nộp nghiệm thu nếu chưa sửa các lỗi mức A. Bốn vấn đề chính:")
    for s in [
        "Khai phương pháp chuyên gia và dữ liệu sơ cấp trong khi đề tài không thực hiện; viết giả thuyết về việc công bố sớm làm mất tính mới như một phát hiện đã chứng minh.",
        "Một số căn cứ pháp luật sai: Điều 86a đã bị bãi bỏ; Điều 73 và Điều 126 Luật Sở hữu trí tuệ dẫn sai nội dung; mức thưởng 30% đến 50% không có trong luật; tên, ngày và nội dung của Nghị định 100/2026 sai.",
        "Bốn nhóm số liệu của Trường lệch với nguồn gốc: kinh phí đề tài theo năm (Hình 2.4), chỉ tiêu Kế hoạch 07 (Hình 2.10), giờ quy đổi và mức thưởng (Hình 2.5), số đề tài tiềm năng thuộc Viện Y - Dược.",
        "Danh mục tài liệu tham khảo cuối sai thông tin xuất bản của 15 tài liệu và mâu thuẫn với danh mục ở Mở đầu; một bài báo bị gán cho sai tác giả (Rialti thay cho O'Dwyer).",
    ]:
        p = doan(doc, "- " + s, sau=3)
        p.paragraph_format.left_indent = Cm(0.6)

    doan(doc, "Những nội dung đã khớp nguồn gốc, có thể giữ:", 13, True, truoc=6)
    for s in DUNG:
        p = doan(doc, "- " + s, sau=3)
        p.paragraph_format.left_indent = Cm(0.6)

    doan(doc, "3. Bảng 1. Danh sách phát hiện", 13, True, truoc=8)
    stt = 0
    for ten, ds in NHOM:
        doan(doc, ten, 12, True, truoc=6, sau=4)
        hang = []
        for vt, nd, dc, m, dx in ds:
            stt += 1
            hang.append((str(stt), vt, nd, dc, MUC[m], dx))
        bang(doc, ["STT", "Vị trí", "Nội dung trong báo cáo", "Đối chiếu nguồn gốc", "Mức", "Đề xuất sửa"],
             [1.3, 3.0, 6.2, 8.0, 2.2, 5.0], hang, mau_muc=True)

    doan(doc, "4. Bảng 2. Thông tin xuất bản sai trong danh mục tài liệu tham khảo cuối", 13, True, truoc=10)
    doan(doc, "Cột thông tin đúng lấy từ bản PDF (khi có) và tệp Zotero trong thư mục đề tài.", 12, nghieng=True)
    bang(doc, ["Tài liệu", "Báo cáo ghi", "Thông tin đúng"], [5.0, 10.0, 10.7], TLTK)

    doan(doc, "5. Văn bản cần bổ sung", 13, True, truoc=10)
    doan(doc, "Các nội dung sau chưa có văn bản gốc trong hồ sơ; cần bổ sung để giữ lại, nếu không thì bỏ: hai hợp đồng chuyển giao quyền sử dụng "
              "sách; Điều lệ Quỹ Học bổng sau tiến sĩ Ngô Xuân Độ; quy định về tỷ lệ trùng lặp; quyết định thành lập Trường; tên các đơn vị ở "
              "Hình 2.3; bản gốc các bài của Maresova et al., Öztürk, Phạm và Nguyễn (2018), Lê và Nguyễn (2019) để kiểm tra các con số được dẫn.")

    doc.save(RA)
    print("Đã lưu", RA, tong)


if __name__ == "__main__":
    main()

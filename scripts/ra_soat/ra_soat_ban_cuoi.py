# -*- coding: utf-8 -*-
"""Rà soát tính chính xác nội dung bản cuối Báo cáo toàn văn (tải lên ngày 06/10/2026).

Chỉ kiểm tra độ chính xác của nội dung, không đánh giá bố cục, trình bày hay văn phong.
Dùng lại các hàm dựng bảng của ra_soat_toan_van.py. Không đưa dữ liệu cá nhân vào báo cáo.
"""
import os
import sys

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm

sys.path.insert(0, os.path.dirname(__file__))
from ra_soat_toan_van import doan, bang, MUC  # noqa: E402

GOC = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
RA = os.path.join(GOC, "Ban_cuoi", "Ban_hoan_thien_03-10-2026", "Bao_cao_ra_soat_ban_cuoi.docx")

DA_SUA = [
    "Bỏ phương pháp chuyên gia và dữ liệu sơ cấp; bổ sung phần giới hạn nghiên cứu, ghi rõ các quan hệ nguyên nhân là giả thuyết; bỏ mã số đề tài tự đặt.",
    "Thay Điều 86a đã bị bãi bỏ bằng khoản 2 Điều 25 Luật số 93/2025/QH15 và điểm c khoản 1 Điều 86 Luật Sở hữu trí tuệ; bỏ dẫn chiếu sai Điều 73, sửa Điều 126 thành Điều 127; bỏ mức thưởng 30% đến 50%.",
    "Sửa câu chữ Điều 135 (10% lợi nhuận trước thuế, 15% tổng số tiền nhận được); Kết luận 51 ghi đúng ngày 17/6/2026; tên Chỉ thị 02 đúng; mô tả Nghị định 100/2026 trong danh mục đúng; dùng Thông tư 83 thay Thông tư 01/2024; mức 5 và 3 sản phẩm quy đổi đúng.",
    "Bỏ Điều 30, Điều 58, Điều 42 dẫn sai; kiến nghị dẫn đúng khoản 1 Điều 28 Luật Giáo dục đại học; căn cứ chi quỹ cho đăng ký sở hữu trí tuệ ở Bảng 1.1 và Kết luận dẫn đúng điểm b khoản 2 Điều 66.",
    "Hình 2.4 (kinh phí theo năm), Hình 2.5 (giờ quy đổi, mức thưởng), Hình 2.10 (chỉ tiêu Kế hoạch 07) và cột đơn vị ở Bảng 2.7 nay khớp nguồn gốc; Kết luận và Bảng 3.1 ghi đúng 9 trên 11 đề tài thuộc Viện Y - Dược.",
    "Bỏ hai hợp đồng sách 5 triệu đồng, khẳng định về tinh dầu bưởi và 61 giáo trình; bỏ Trung tâm Đào tạo và Chuyển giao công nghệ, Viện Quản trị và Sáng tạo số, Trung tâm Tuyển sinh khỏi các nội dung quản lý tài sản trí tuệ; bổ sung Phòng Công nghệ, Đổi mới sáng tạo và Khởi nghiệp ở Bảng 3.3 và Kiến nghị.",
    "Sửa Rialti thành O'Dwyer et al. (2023); bỏ Milliken và Allen; Tewari và Bhardwaj ghi năm 2021; bỏ dẫn chiếu Mục 1.4.1.",
]

CU = [
    ("Mở đầu 2.1 (Thứ nhất); 1.1.1",
     "Phạm Thúy Hằng (2019) xác định bốn nội dung quản lý: thể chế nội bộ, tổ chức bộ máy, chu trình xác lập, khai thác và bảo vệ; định nghĩa có chủ thể quản lý là Hội đồng trường, Ban Giám hiệu, đơn vị chuyên trách.",
     "Bài gốc do Phạm Thị Thúy Hằng viết, nêu bốn nội dung: quản lý hành chính sở hữu trí tuệ (phát hiện, khai báo, ghi nhận); quản lý xác lập và bảo vệ quyền; quản lý khai thác thương mại; quản lý môi trường và điều kiện hỗ trợ. Định nghĩa gốc không liệt kê Hội đồng trường, Ban Giám hiệu.",
     "A", "Sửa tên tác giả, bốn nội dung và định nghĩa theo bài gốc."),
    ("Mở đầu 2.1 (Thứ ba)",
     "Phạm Thị Thúy Hằng và Nguyễn Thanh Hùng (2018): mô hình quản lý 3 cấp; chính sách dành 50% - 70% doanh thu cấp phép cho tác giả.",
     "Bài gốc ghi nhà phát minh được hưởng tối thiểu 25% doanh thu tạo ra từ sở hữu trí tuệ; không có tỷ lệ 50% - 70% và không mô tả mô hình 3 cấp.",
     "A", "Sửa thành tối thiểu 25% doanh thu; bỏ cụm mô hình 3 cấp."),
    ("Mở đầu 2.1; Tài liệu tham khảo số 31",
     "Lê Thị Thu Hà và Nguyễn Thành Khang (2019).",
     "Bài của Phạm Thị Thúy Hằng (2019) dẫn bài này với năm 2017. Thư mục chưa có bài gốc.",
     "B", "Kiểm tra năm xuất bản trên bài gốc."),
    ("Mở đầu 2.2 (Thứ nhất)",
     "Maresova et al. tổng quan 22 công trình, phân loại theo bốn trụ cột; Holgersson (2021) phân tích sự chuyển dịch từ quyền nhà sáng chế sang quyền sở hữu của đại học.",
     "Con số 22 và bốn trụ cột chưa kiểm chứng được vì thư mục chưa có bài gốc. Báo cáo của Holgersson là khảo sát 21 trường đại học châu Âu; quy định miễn trừ cho giảng viên chỉ được nêu như một khác biệt giữa các nước.",
     "B", "Kiểm tra bài Maresova (truy cập mở); diễn đạt lại phần Holgersson."),
    ("Mở đầu 2.2 (Thứ hai, Thứ ba)",
     "Öztürk (2026) chứng minh văn phòng chuyển giao giữ vai trò quyết định, rào cản lớn nhất là thiếu nhân sự có chứng chỉ; Pandey et al. (2025) khẳng định cần kết hợp với trung tâm ươm tạo, định giá và chia doanh thu linh hoạt (dẫn kèm Teece, 2018); Pandey et al. đề xuất hệ thống chỉ số; Holgersson và Öztürk chỉ ra phần mềm quản trị vòng đời tài sản trí tuệ.",
     "Öztürk viết vai trò quan trọng, các văn phòng gặp khó khăn về nhân sự có năng lực và thương mại hóa, không xếp hạng rào cản. Pandey et al. chỉ nêu rào cản thủ tục, kinh phí, hợp tác; nội dung ươm tạo, chia doanh thu, chỉ số có trong bài Bulsara và Vaghela. Teece (2018) phân tích cấp phép công nghệ viễn thông không dây. Holgersson và Öztürk không bàn về phần mềm quản lý vòng đời.",
     "B", "Gán lại nội dung cho đúng bài; bỏ Teece và câu về phần mềm."),
    ("1.1.2 (Hai là)",
     "Guan (2014) và Perkmann et al. (2013) nêu xung đột giữa công bố nhanh và điều kiện tính mới tuyệt đối.",
     "Perkmann et al. tổng quan về tương tác học thuật và thương mại hóa, không bàn về tính mới sáng chế; chương của Guan bàn về khái niệm sở hữu trí tuệ.",
     "B", "Dẫn Điều 60 Luật Sở hữu trí tuệ cho điều kiện tính mới."),
    ("Bảng 2.4, Bảng 2.3",
     "Thông tin Quỹ Học bổng sau tiến sĩ Ngô Xuân Độ (5 tỷ đồng; Điều 9 chia 50/50 rồi 20/80).",
     "Hồ sơ vẫn chưa có Điều lệ Quỹ.",
     "B", "Bổ sung Điều lệ hoặc ghi chú chưa đối chiếu văn bản gốc."),
    ("Mục 2.2.4",
     "Quy định bắt buộc kiểm tra trùng lặp bằng Turnitin với 100% luận văn, khóa luận, thuyết minh đề tài; ngưỡng dưới 20%.",
     "Không có văn bản nguồn trong hồ sơ.",
     "B", "Dẫn quyết định, quy định của Trường hoặc bỏ."),
    ("Hình 2.6; 2.3.2; Bảng 3.1 (W5); Bảng 3.3",
     "Bộ 16 tiêu chí quản trị tài sản trí tuệ tiêu chuẩn; 7 trên 16 tiêu chí chưa tính được.",
     "Báo cáo không trình bày danh sách 16 tiêu chí ở đâu; Bảng 1.2 là bảng đối sánh xếp hạng, không phải bộ 16 tiêu chí.",
     "B", "Bổ sung danh sách 16 tiêu chí và nguồn của bộ tiêu chí."),
    ("3.2.4; Bảng 3.3",
     "Trung tâm Công nghệ thông tin chủ trì xây dựng phần mềm.",
     "Đơn vị này không có trong cơ cấu ở Bảng 2.1.",
     "B", "Kiểm tra tên đơn vị."),
    ("2.3.3; Kiến nghị",
     "Thời gian thẩm định đơn sáng chế, giải pháp hữu ích thực tế 24 - 36 tháng, cá biệt 48 tháng.",
     "Chưa dẫn nguồn.",
     "C", "Dẫn nguồn số liệu."),
    ("2.1.1",
     "Triết lý giáo dục Trí - Năng - Nhân - Hòa.",
     "Khung chiến lược Edupark ghi Trí - Năng - Hòa - Nhân.",
     "C", "Sửa thứ tự."),
]

MOI = [
    ("Bảng 1.2 (dòng Thông tư 83)",
     "Tiêu chí 4.2: số sáng chế, giải pháp hữu ích, văn bằng trên giảng viên; Tiêu chí 4.3: tỷ trọng thu từ khoa học công nghệ; mốc tối thiểu 0,02 văn bằng/giảng viên/năm và 5% tổng thu.",
     "Trong Thông tư 83, Tiêu chuẩn 4 là Tài chính (Tiêu chí 4.2 là chỉ số tăng trưởng bền vững). Thu khoa học công nghệ không thấp hơn 5% là Tiêu chí 6.1 và chỉ áp dụng cho cơ sở có đào tạo tiến sĩ; Tiêu chí 6.2 là 0,3 sản phẩm quy đổi/giảng viên/năm (0,6 nếu có đào tạo tiến sĩ). Thông tư không có mốc 0,02 văn bằng/giảng viên.",
     "A", "Sửa theo Tiêu chí 6.1 và 6.2 của Thông tư 83."),
    ("Bảng 1.2 (dòng Nghị định 109/2022)",
     "Mốc chuẩn để thành lập viện nghiên cứu, trung tâm chuyển giao, đầu tư phòng thí nghiệm: tối thiểu 01 sáng chế hoặc 02 giải pháp hữu ích/năm; chuyển giao ≥ 05 công nghệ/5 năm; thương mại hóa ≥ 05 sản phẩm/5 năm.",
     "Các mốc này là tiêu chí công nhận nhóm nghiên cứu mạnh trong cơ sở giáo dục đại học; về chuyển giao, điều kiện là chuyển giao tối thiểu 05 công nghệ hoặc thương mại hóa tối thiểu 05 sản phẩm, hoặc có 01 sản phẩm quốc gia, trong 5 năm.",
     "B", "Sửa mục đích áp dụng và điều kiện hoặc."),
    ("Bảng 1.2 (dòng UPM); 2.1.1",
     "Trường đạt 4 sao UPM.",
     "Chiến lược của Trường (bản 2024 - 2025) ghi đạt 3 sao UPM năm 2020; hồ sơ chưa có chứng nhận 4 sao.",
     "B", "Bổ sung chứng nhận 4 sao hoặc sửa thành 3 sao (2020)."),
    ("Bảng 1.2 (các dòng UPM, THE, SCImago, Thông tư 20); Bảng 1.1",
     "Tỷ trọng 17%, 88,4%, 30%, 4/60 tiêu chí; nguồn Bảng 1.1 là Bộ công cụ chính sách sở hữu trí tuệ của WIPO năm 2020.",
     "Các con số khớp bảng đối sánh do nhóm lập (Doi-sanh-chi-so-SHTT-xep-hang-kiem-dinh.xlsx), nhưng phương pháp luận UPM, THE, SCImago và tài liệu WIPO được dùng làm nguồn không có trong danh mục tài liệu tham khảo; tài liệu WIPO trong danh mục (số 450, Sở hữu trí tuệ là gì?) là tờ giới thiệu, không phải bộ công cụ chính sách. Dòng Thông tư 20 nêu 4 trên 60 tiêu chí nhưng chỉ liệt kê 3 (bảng đối sánh có thêm Tiêu chí 7.4).",
     "B", "Bổ sung nguồn gốc vào danh mục; liệt kê đủ 4 tiêu chí."),
    ("Bảng 1.1 (dòng 6)",
     "Điểm c khoản 4 Điều 3 Luật số 125/2025/QH15.",
     "Khoản 4 Điều 3 là định nghĩa liêm chính học thuật, không có điểm.",
     "B", "Sửa thành khoản 4 Điều 3."),
    ("3.1.1 (Thứ ba); 3.2.3 (Hai là)",
     "Điều 9a Nghị định 100/2026 hướng dẫn cơ chế định giá tài sản trí tuệ hình thành từ ngân sách hoặc vốn hỗn hợp, tháo gỡ rủi ro pháp lý cho người đứng đầu; Viện REK định giá theo Nghị định 100.",
     "Điều 9a (được bổ sung vào Nghị định 65/2023 bởi Điều 8 Nghị định 100) quy định chủ sở hữu lập, cập nhật Danh mục quyền sở hữu trí tuệ để quản trị nội bộ; khoản 5 chỉ giao bộ, ngành, địa phương khuyến khích, hỗ trợ xác định giá trị. Không có cơ chế định giá cho vốn hỗn hợp hay quy định về rủi ro của người đứng đầu. Cách dùng Điều 9a ở Bảng 1.1, Mục 1.2.2 và Kiến nghị (lập Danh mục) là đúng.",
     "A", "Sửa mô tả Điều 9a; bỏ câu định giá theo Nghị định 100."),
    ("3.1.1 (Thứ nhất); 3.2.3 (Một là)",
     "Điểm b khoản 2 Điều 66 Luật số 93/2025/QH15 xác lập cơ chế phân bổ nguồn thu từ tài sản trí tuệ; phân chia lợi ích theo khoản 3 Điều 28 và khoản 2 Điều 66.",
     "Điểm b khoản 2 Điều 66 cho phép dùng quỹ phát triển khoa học và công nghệ để chi đăng ký, bảo hộ, quản lý, khai thác quyền sở hữu trí tuệ; không quy định phân bổ nguồn thu hay phân chia lợi ích.",
     "A", "Bỏ Điều 66 khỏi nội dung phân chia lợi ích."),
    ("3.1.1 (Thứ nhất)",
     "Khoản 2 Điều 25 cho cơ sở nghiên cứu toàn quyền định đoạt kết quả nhiệm vụ; điểm a khoản 3 Điều 28 cho tác giả sáng chế hưởng tối thiểu 30% lợi nhuận.",
     "Cả hai quy định chỉ áp dụng cho phần kết quả tương ứng với kinh phí ngân sách nhà nước; Điều 25 có các trường hợp loại trừ; khoản 3 Điều 28 áp dụng cho tác giả kết quả nhiệm vụ, không riêng sáng chế.",
     "B", "Bổ sung điều kiện ngân sách nhà nước."),
    ("Bảng 3.1 (O2); Tài liệu tham khảo số 21",
     "Quyết định 1624/QĐ-TTg ngày 26/12/2025.",
     "Văn bản gốc trong thư mục VBPL: ngày 21/8/2026.",
     "A", "Sửa ngày ban hành."),
    ("Tài liệu tham khảo số 1",
     "Kết luận 51-KL/TW về tiếp tục đổi mới căn bản, toàn diện giáo dục đại học, tạo đột phá phát triển nguồn nhân lực chất lượng cao.",
     "Tên đúng: về đẩy mạnh công tác sở hữu trí tuệ phục vụ phát triển kinh tế - xã hội trong tình hình mới.",
     "A", "Sửa tên văn bản."),
    ("Tài liệu tham khảo số 4",
     "Thông tư 20/2026/TT-BGDĐT ngày 20/7/2026 ban hành Quy định Chuẩn kiểm định chất lượng.",
     "Thông tư ban hành ngày 31/3/2026, hiệu lực 15/5/2026, quy định về kiểm định chất lượng cơ sở giáo dục đại học (tra cứu trang văn bản pháp luật; thư mục chưa có văn bản).",
     "B", "Sửa ngày và tên; bổ sung văn bản vào VBPL."),
    ("Mở đầu, mục 7 (kiểm chứng hồ sơ sở hữu trí tuệ)",
     "Toàn bộ 11 đối tượng sở hữu trí tuệ được đối soát trực tiếp với Sổ bộ đơn và Sổ bộ văn bằng tại Cục Sở hữu trí tuệ và Cục Bản quyền tác giả.",
     "Mâu thuẫn với Bảng 2.6, Hình 2.7: 5 kiểu dáng công nghiệp vẫn ghi chưa xác minh trạng thái. Hồ sơ chỉ có sổ theo dõi nội bộ của Trường.",
     "A", "Bỏ câu đối soát với Cục, hoặc bổ sung kết quả tra cứu và cập nhật trạng thái 5 kiểu dáng."),
    ("Bảng 2.2; Bảng 2.5",
     "Số liệu mới: bài trong nước 17, 62, 37, 72, 109 (297); bài quốc tế 6, 8, 16, 33, 52; giáo trình 26, 20, 17, 13, 11; sách 1, 0, 3, 4, 4; tham luận quốc tế 0, 0, 5, 1, 7 (13).",
     "Lệch bộ dữ liệu đã đối chiếu với các tệp thống kê gốc (bài trong nước 17, 42, 50, 72, 109 = 290; bài quốc tế 6, 13, 11, 33, 52; giáo trình 26, 31, 12, 7, 11; sách 0, 2, 2, 2, 6; tham luận quốc tế 1, 2, 3, 5, 5 = 16). Bảng mới tự cộng không khớp: cột 2024 ra 132 (ghi 128), cộng các dòng ra 586 (ghi 582). Bảng 2.5 vẫn ghi 405 bài báo, khác tổng 412 của bảng mới.",
     "A", "Trả về số liệu đã đối chiếu, hoặc nêu nguồn của số mới và sửa tổng."),
    ("2.1.3 (đoạn sau Bảng 2.2); 2.3.1",
     "161 bài báo (71 quốc tế, 90 trong nước); 87 giáo trình và sách chuyên khảo, trong đó 8 cuốn ISBN; 296 bài kỷ yếu hội thảo.",
     "161 là số bài báo năm 2025, không phải cả giai đoạn; 71 là số bài quốc tế có phân hạng, 90 không khớp số bài trong nước nào. Bảng 2.2 có 87 giáo trình và 12 sách riêng; tham luận chỉ có 34 (21 quốc gia, 13 quốc tế); 296 không có cơ sở trong bảng.",
     "A", "Viết lại đoạn theo đúng Bảng 2.2."),
    ("Mở đầu 1; 2.1.3",
     "Hoàn thành 38 đề tài cấp cơ sở được cấp tổng kinh phí 424,75 triệu đồng.",
     "424,75 triệu đồng chỉ cấp cho 19 đề tài; 16 đề tài chỉ quy đổi giờ, 3 đề tài tự tìm tài trợ (Bảng 2.4).",
     "B", "Ghi rõ 19 trên 38 đề tài được cấp kinh phí."),
    ("2.1.3; Tiểu kết Chương 2",
     "Từ năm 2022 ngân sách cấp trực tiếp tăng trưởng liên tục; kinh phí tăng trưởng đều đặn; 3 đề tài năm 2024 nhận thêm tài trợ doanh nghiệp.",
     "Kinh phí năm 2023 (70 triệu đồng) thấp hơn năm 2022 (84 triệu đồng). Danh mục đề tài ghi 3 đề tài tự tìm nguồn tài trợ, không ghi là doanh nghiệp.",
     "B", "Sửa nhận định xu hướng và nguồn tài trợ."),
    ("2.1.1; Tài liệu tham khảo số 18",
     "Quyết định 679/QĐ-TTg ngày 19/5/2009 thành lập Trường.",
     "Chiến lược của Trường ghi Trường được thành lập ngày 27/5/2009. Thư mục chưa có quyết định để đối chiếu số hiệu (bản trước ghi 676).",
     "B", "Đối chiếu quyết định gốc."),
    ("2.1.2; 3.2.1; 3.2.3; Bảng 3.3; Kiến nghị",
     "Viện Quản trị và Công nghệ (Viện REK) thành lập theo Quyết định 217 trên cơ sở kiện toàn Viện Nghiên cứu giáo dục và Chuyển giao tri thức; ở chỗ khác Viện REK lại là Viện Nghiên cứu giáo dục và Chuyển giao tri thức.",
     "Báo cáo tự mâu thuẫn về Viện REK là đơn vị nào; Bảng 2.1 vẫn ghi hai viện riêng. Quyết định 217 là Quy chế quản trị tài sản trí tuệ, không phải quyết định thành lập viện. Dữ liệu đề tài ghi Viện REK cho đề tài 08-2023 của Viện Nghiên cứu giáo dục và Chuyển giao tri thức.",
     "A", "Xác nhận Viện REK là đơn vị nào và thống nhất ở mọi chỗ; bỏ câu thành lập theo Quyết định 217."),
    ("2.1.2; Hình 2.3",
     "Cơ cấu gồm các khoa chuyên môn và hai viện nghiên cứu nòng cốt; Hình 2.3 có ba đầu mối.",
     "Theo xác nhận của chủ nhiệm đề tài, mỗi viện có Phòng Học vụ và Hợp tác đối ngoại và Phòng Công nghệ, Đổi mới sáng tạo và Khởi nghiệp. Hình 2.3 và Mục 2.1.2 chưa có các phòng này, dù Bảng 3.3 và Kiến nghị đã nhắc tới.",
     "B", "Bổ sung các phòng vào mô tả cơ cấu và Hình 2.3."),
    ("2.2.1; Hình 2.5",
     "Quyết định 213 ngày 15/4/2021, được sửa đổi, bổ sung bởi Quyết định 217; nguồn Hình 2.5 là Quyết định 213 và 217.",
     "Văn bản gốc: Quy chế ban hành kèm số 213/QĐ-ĐHTĐ ngày 28/12/2021 (danh mục cũng ghi ngày này). Quyết định 217 là quy chế riêng về tài sản trí tuệ. Giờ quy đổi và mức thưởng ở Hình 2.5 lấy từ Bảng 6, Bảng 7 Quy chế chi tiêu nội bộ năm 2026.",
     "B", "Sửa ngày, quan hệ giữa hai quyết định và nguồn của Hình 2.5."),
    ("2.2.1 (đoạn sau Bảng 2.4)",
     "Cả ba kênh (đề tài cơ sở, NAFOSTED, đề tài phối hợp doanh nghiệp) chưa có dự toán riêng; giảng viên phải tự ứng kinh phí cá nhân.",
     "Bảng 2.4 có ba kênh khác: đề tài cơ sở, Quỹ Ngô Xuân Độ, NAFOSTED. Chính Bảng 2.4 ghi Quyết định 213 giao Phòng KHCN nộp đơn, nộp lệ phí và cho phép chi thuê ngoài.",
     "B", "Sửa cho khớp Bảng 2.4."),
    ("2.2.2 (đoạn sau Bảng 2.5)",
     "Quy chế nội bộ bỏ trống hoàn toàn quy định về bí mật kinh doanh, kiểu dáng công nghiệp, thiết kế bố trí, giống cây trồng.",
     "Quyết định 217 có định nghĩa và liệt kê các đối tượng này (Điều 2 và danh mục tài sản trí tuệ của Trường, gồm bí mật thương mại, kiểu dáng công nghiệp, thiết kế bố trí). Chính Bảng 2.5 ghi quy chế có liệt kê.",
     "A", "Sửa nhận định theo Quyết định 217."),
    ("2.2.4; 2.3.2; 2.3.3; Mở đầu",
     "Chưa ban hành quy định bảo mật thông tin, bảo vệ bí mật kinh doanh trước khi nộp đơn; chưa thiết lập quy trình sàng lọc trước công bố.",
     "Điều 10 Quyết định 217 quy định nghĩa vụ bảo mật của mọi đơn vị, người lao động, người học và yêu cầu tác giả xin ý kiến Phòng KHCN trước khi bộc lộ công khai tài sản có thể bảo hộ.",
     "A", "Viết lại: đã có quy định nguyên tắc tại Điều 10 Quyết định 217 nhưng chưa có quy trình, biểu mẫu và bằng chứng áp dụng."),
    ("Bảng 2.6; 2.2.2; 2.3.2; Bảng 3.1 (S4)",
     "Đơn sáng chế 1-2025-07378: đang thẩm định hình thức; đang thẩm định nội dung; đang ở bước tiếp nhận đơn.",
     "Ba cách ghi khác nhau trong cùng báo cáo; sổ theo dõi chỉ xác định được là đã nộp đơn năm 2025.",
     "B", "Thống nhất một trạng thái có căn cứ."),
    ("2.3.2 (Thứ nhất)",
     "2 đề tài quyền tác giả (phần mềm và tài liệu giáo trình) đã được áp dụng nội bộ.",
     "Bảng 2.7: hai sản phẩm là bộ mẫu cây thuốc và 100 tiêu bản hiển vi (sưu tập dữ liệu), chưa đăng ký.",
     "A", "Sửa theo Bảng 2.7."),
    ("2.3.1; Bảng 3.1 (S2)",
     "Hoàn thành vượt mức 5 trên 9 chỉ tiêu Kế hoạch 07.",
     "Hình 2.10 có 6 chỉ tiêu đạt hoặc vượt (kể cả đề tài cấp cơ sở 107%).",
     "B", "Sửa thành 6 trên 9."),
    ("3.3.2",
     "Viện Y - Dược sở hữu 11 đề tài có sản phẩm tiềm năng.",
     "Bảng 2.7: 9 trên 11 thuộc Viện Y - Dược.",
     "A", "Sửa thành 9."),
    ("2.2.3 (mô tả Điều 36)",
     "Điểm a: tối đa 100 triệu đồng cho một hợp đồng; điểm b: tài sản hình thành từ kinh phí của Trường, tối đa 30% lợi nhuận; hai điểm xung đột nội tại.",
     "Quyết định 213: điểm a trần 100 triệu đồng một đề tài; điểm b áp dụng cho tài sản thuộc sở hữu của Trường, tác giả hưởng 30% kinh phí chuyển giao sau chi phí. Hai điểm áp dụng cho hai phạm vi khác nhau nên không cùng áp dụng cho một tài sản.",
     "B", "Sửa câu chữ; bỏ cụm xung đột nội tại."),
    ("Bảng 3.3; Kiến nghị",
     "Bảng 3.3 dẫn Hạn chế 4, 5, 6, 7 và Quy trình sàng lọc 3 bước.",
     "Mục 2.3.2 chỉ còn 4 hạn chế với nội dung khác; Mục 3.2.2 mô tả Hội đồng thẩm định, không có quy trình 3 bước.",
     "B", "Sửa dẫn chiếu cho khớp 2.3.2 và 3.2.2."),
    ("Trích dẫn trong bài và danh mục",
     "Trích Etzkowitz (2000), Siegel et al. (2003), Nghị quyết 45-NQ/TW.",
     "Danh mục có Etzkowitz (2003), Siegel et al. (2007), không có Nghị quyết 45. Danh mục có nhưng thân bài không trích: Thursby và Kemp, Rocha et al., Holgersson và Aaboen, Luật Giá, Luật Chuyển giao công nghệ, Nghị định 17/2023, Nghị định 65/2023, Thông tư 01/2024, Văn bản hợp nhất 67, Luật 131/2025, tài liệu tập huấn của Cục Sở hữu trí tuệ.",
     "B", "Thống nhất trích dẫn và danh mục hai chiều."),
    ("Toàn bài",
     "Phòng Quản lý Khoa học và Công nghệ; nhãn hiệu chữ Thanh do University, nhãn hiệu hình Thado Edupark; Rutin từ hoa Hòe.",
     "Các chỗ khác ghi Phòng Khoa học Công nghệ; Bảng 2.6 ghi cả hai nhãn hiệu là chữ và hình; danh mục đề tài không ghi nguồn hoa Hòe.",
     "C", "Thống nhất theo nguồn."),
]

TLTK = [
    ("Bstieler et al. (2015), số 39", "DOI 10.1111/jpim.12166", "DOI 10.1111/jpim.12242 (bản PDF)"),
    ("Holgersson & Aaboen (2019), số 46", "DOI 10.1016/j.techsoc.2019.04.002; tên bài thiếu phụ đề", "DOI 10.1016/j.techsoc.2019.04.008; ... From appropriation to utilization (bản PDF)"),
    ("Rocha et al. (2023), số 52", "DOI 10.1504/IJEIM.2023.129202", "DOI 10.1504/IJEIM.2023.129328 (trang nhà xuất bản)"),
    ("Holgersson (2021), số 45", "European Patent Office & Chalmers University of Technology", "European Commission, Directorate-General for Research and Innovation, 2022, doi:10.2777/969317 (bản PDF)"),
    ("Bulsara & Vaghela (2025), số 40", "Tên bài: ... intellectual property-driven technology transfer: Fostering innovation and sustainable growth; Asia-Pacific Tech Monitor, WIPO", "The role of universities in IP-driven technology transfer: Fostering innovation through policy and practice; Asia-Pacific Tech Monitor, 42(2), APCTT (ESCAP), không phải WIPO (bản PDF)"),
    ("Tewari & Bhardwaj (2021), số 56", "Bhardwaj, K.; Intellectual Property Rights: A Prerequisite for Knowledge Economy", "Bhardwaj, M. (Mamta); Intellectual Property: A Primer for Academia, Publication Bureau, Panjab University (bản PDF)"),
    ("Phạm, T. H. (2019), số 34", "Phạm, T. H.", "Phạm, T. T. H. (Phạm Thị Thúy Hằng) (bản PDF)"),
    ("Võ, N. H. P. (2025), số 37", "Tên bài: ...: Kinh nghiệm quốc tế và giải pháp cho Việt Nam", "..., kinh nghiệm quốc tế để hoàn thiện chính sách sở hữu trí tuệ của các trường đại học tại Việt Nam (bản PDF)"),
    ("Maresova et al. (2019), số 47", "Tên bài thiếu: management: A systematic review", "Models, processes, and roles of universities in technology transfer management: A systematic review"),
    ("Pandey et al. (2025), số 50", "30, 640-649", "30(5), 640-649"),
]

CAN_BO_SUNG = (
    "Điều lệ Quỹ Học bổng sau tiến sĩ Ngô Xuân Độ; quy định kiểm tra trùng lặp của Trường; quyết định thành lập Trường và Trường Cao đẳng; "
    "chứng nhận UPM 4 sao (nếu có); văn bản Thông tư 20/2026 và Nghị định 109/2022; tài liệu WIPO dùng cho Bảng 1.1 và phương pháp luận "
    "UPM, THE, SCImago dùng cho Bảng 1.2; bài gốc của Maresova et al. (2019) và của Lê và Nguyễn (Tạp chí Khoa học Đại học Văn Lang); "
    "kết quả tra cứu tình trạng 5 kiểu dáng công nghiệp và đơn sáng chế 1-2025-07378; xác nhận Viện REK là đơn vị nào."
)


def gach(doc, ds):
    for s in ds:
        p = doan(doc, "- " + s, sau=3)
        p.paragraph_format.left_indent = Cm(0.6)


def bang_phat_hien(doc, ds, bat_dau):
    hang = [(str(bat_dau + k), vt, nd, dc, MUC[m], dx) for k, (vt, nd, dc, m, dx) in enumerate(ds)]
    bang(doc, ["STT", "Vị trí", "Nội dung trong báo cáo", "Đối chiếu nguồn gốc", "Mức", "Đề xuất sửa"],
         [1.3, 3.0, 6.2, 8.0, 2.2, 5.0], hang, mau_muc=True)
    return bat_dau + len(ds)


def main():
    doc = Document()
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = Cm(29.7), Cm(21.0)
    sec.left_margin = sec.right_margin = sec.top_margin = sec.bottom_margin = Cm(2.0)

    doan(doc, "BÁO CÁO RÀ SOÁT TÍNH CHÍNH XÁC NỘI DUNG", 15, True, WD_ALIGN_PARAGRAPH.CENTER, sau=2)
    doan(doc, "Bản cuối Báo cáo toàn văn tổng kết đề tài (tải lên ngày 06/10/2026)", 13, False, WD_ALIGN_PARAGRAPH.CENTER, sau=12, nghieng=True)

    tat_ca = CU + MOI
    dem = {m: sum(1 for x in tat_ca if x[3] == m) for m in "ABC"}
    doan(doc, "1. Phạm vi", 13, True, truoc=6)
    doan(doc, "Bản cuối được đọc toàn văn (khoảng 27.000 chữ) và đối chiếu với văn bản pháp luật trong thư mục VBPL, tài liệu của Trường, "
              "bộ dữ liệu đã chuẩn hóa, bản PDF các tài liệu tham khảo và kết quả lần rà soát trước. Thông tư 20/2026 và Nghị định 109/2022 "
              "không có trong thư mục nên được tra cứu trên mạng. Chỉ kiểm tra tính chính xác của nội dung, không đánh giá bố cục, trình bày "
              "hay văn phong. Mức độ: A là lỗi phải sửa trước khi nộp; B là lỗi cần sửa; C là sai sót nhỏ.")

    doan(doc, "2. Kết luận chung", 13, True, truoc=6)
    doan(doc, f"Bản cuối đã sửa được phần lớn lỗi pháp luật và số liệu của lần rà soát trước. Còn {len(tat_ca)} lỗi về nội dung, gồm "
              f"{dem['A']} lỗi mức A, {dem['B']} lỗi mức B, {dem['C']} lỗi mức C; trong đó {len(CU)} lỗi cũ chưa sửa và {len(MOI)} lỗi mới "
              f"phát sinh trong các phần được viết lại, chủ yếu ở Bảng 1.2, Bảng 2.2, phần mô tả Viện REK, nhận định về Quyết định 217 và "
              f"cách dùng Điều 66 Luật số 93/2025/QH15, Điều 9a Nghị định 100/2026. Ngoài ra có {len(TLTK)} mục tài liệu tham khảo sai DOI, "
              f"tên bài hoặc tên tác giả (Bảng 3).")
    doan(doc, "Đã sửa đúng so với lần rà soát trước:", 13, True, truoc=4)
    gach(doc, DA_SUA)

    doan(doc, "3. Bảng 1. Lỗi cũ chưa sửa", 13, True, truoc=8)
    n = bang_phat_hien(doc, CU, 1)
    doan(doc, "4. Bảng 2. Lỗi mới phát sinh trong bản cuối", 13, True, truoc=10)
    bang_phat_hien(doc, MOI, n)
    doan(doc, "5. Bảng 3. Tài liệu tham khảo sai DOI, tên bài hoặc tên tác giả", 13, True, truoc=10)
    bang(doc, ["Tài liệu", "Báo cáo ghi", "Thông tin đúng"], [5.0, 10.0, 10.7], TLTK)
    doan(doc, "6. Văn bản cần bổ sung", 13, True, truoc=10)
    doan(doc, CAN_BO_SUNG)

    doc.save(RA)
    print("Đã lưu", RA, len(CU), len(MOI), dem)


if __name__ == "__main__":
    main()

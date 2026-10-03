# -*- coding: utf-8 -*-
"""Nội dung báo cáo toàn văn theo khung đã chốt ngày 03/10/2026 (cả ba chương theo chu trình 4 khâu).

Ký hiệu:
- ("CH", số, [dòng tên chương]) tiêu đề chương; ("H1"|"H2"|"TK", chữ) tiêu đề mục; ("P", chữ) đoạn, nhận **đậm**.
- ("GP", đầu đoạn) giữ nguyên đoạn của bản trước; ("GPS", đầu đoạn, {cũ: mới}) giữ đoạn và thay chữ (chỉ dùng cho đoạn
  không có chữ đậm).
- ("GK", đầu chú thích, {cũ: mới}) giữ khối bảng, hình từ chú thích đến dòng nguồn, đánh số lại.
- ("GKR", đầu đoạn đầu, đầu đoạn cuối, {cũ: mới}) giữ một dải đoạn liên tiếp.
- ("BANG", chú thích, tiêu đề cột, các hàng, nguồn, độ rộng cột) bảng mới.
"""

VIET_TAT = [
    ("Chỉ thị 02", "Chỉ thị số 02/CT-TTg ngày 30 tháng 01 năm 2026 của Thủ tướng Chính phủ về tăng cường thực thi quyền sở "
                   "hữu trí tuệ"),
    ("HEMIS", "Hệ thống cơ sở dữ liệu giáo dục đại học do Bộ Giáo dục và Đào tạo quản lý"),
    ("Kế hoạch 07", "Kế hoạch số 07/KH-ĐHTĐ ngày 01 tháng 7 năm 2024 của Trường Đại học Thành Đô về hoạt động khoa học công "
                    "nghệ giai đoạn 2024 - 2028"),
    ("Kết luận 51", "Kết luận số 51-KL/TW ngày 17 tháng 6 năm 2026 của Bộ Chính trị về đẩy mạnh công tác sở hữu trí tuệ phục "
                    "vụ phát triển kinh tế - xã hội trong tình hình mới"),
    ("Luật Giáo dục đại học", "Luật Giáo dục đại học số 125/2025/QH15 ngày 10 tháng 12 năm 2025"),
    ("Luật KH,CN&ĐMST", "Luật Khoa học, công nghệ và đổi mới sáng tạo số 93/2025/QH15 ngày 27 tháng 6 năm 2025"),
    ("Luật Sở hữu trí tuệ", "Luật Sở hữu trí tuệ số 50/2005/QH11, được sửa đổi, bổ sung bởi các Luật số 36/2009/QH12, "
                            "42/2019/QH14, 07/2022/QH15, 93/2025/QH15 và 131/2025/QH15; hợp nhất tại Văn bản hợp nhất số "
                            "67/VBHN-VPQH"),
    ("Nghị định 134", "Nghị định số 134/2026/NĐ-CP ngày 06 tháng 4 năm 2026 của Chính phủ sửa đổi, bổ sung một số điều của "
                      "Nghị định quy định chi tiết Luật Sở hữu trí tuệ về quyền tác giả, quyền liên quan"),
    ("Nghị định 267", "Nghị định số 267/2025/NĐ-CP ngày 14 tháng 10 năm 2025 của Chính phủ quy định chi tiết và hướng dẫn "
                      "một số điều của Luật Khoa học, công nghệ và đổi mới sáng tạo"),
    ("Quyết định 1068", "Quyết định số 1068/QĐ-TTg ngày 22 tháng 8 năm 2019 của Thủ tướng Chính phủ phê duyệt Chiến lược sở "
                        "hữu trí tuệ đến năm 2030"),
    ("Quyết định 1624", "Quyết định số 1624/QĐ-TTg ngày 21 tháng 8 năm 2026 của Thủ tướng Chính phủ sửa đổi, bổ sung Quyết "
                        "định số 1068/QĐ-TTg"),
    ("Quyết định 213", "Quy chế hoạt động khoa học công nghệ Trường Đại học Thành Đô ban hành kèm Quyết định số "
                       "213/QĐ-ĐHTĐ ngày 28 tháng 12 năm 2021"),
    ("Quyết định 217", "Quy chế quản trị tài sản trí tuệ tại Trường Đại học Thành Đô ban hành kèm Quyết định số "
                       "217/QĐ-ĐHTĐ ngày 21 tháng 11 năm 2024"),
    ("Thông tư 83", "Thông tư số 83/2026/TT-BGDĐT ngày 30 tháng 9 năm 2026 của Bộ Giáo dục và Đào tạo quy định Chuẩn cơ sở "
                    "giáo dục đại học"),
]

MO_DAU = [
    ("GKR", "MỞ ĐẦU", "Các mô hình quốc tế phần lớn được xây dựng", {}),
    ("GP", "3. Mục tiêu nghiên cứu"),
    ("GP", "Mục tiêu tổng quát:"),
    ("P", "**Mục tiêu cụ thể:** thứ nhất, làm rõ nội dung quản lý quyền sở hữu trí tuệ trong trường đại học theo bốn khâu "
          "sáng tạo, xác lập, khai thác, bảo vệ và các yêu cầu pháp lý đối với nhà trường; thứ hai, đánh giá thực trạng giai "
          "đoạn 2021 - 2025 theo bốn khâu, chỉ ra kết quả, hạn chế và nguyên nhân; thứ ba, đề xuất giải pháp kèm kế hoạch "
          "thực hiện, thí điểm và chỉ số theo dõi."),
    ("GP", "Câu hỏi nghiên cứu:"),
    ("GP", "4. Đối tượng nghiên cứu"),
    ("GP", "Đối tượng nghiên cứu là"),
    ("GKR", "5. Phạm vi nghiên cứu", "- Về thời gian:", {}),
    ("GP", "6. Phương pháp nghiên cứu"),
    ("P", "Đề tài kết hợp ba cách tiếp cận đã xác định trong thuyết minh. Tiếp cận hệ thống xem quản lý quyền sở hữu trí tuệ "
          "trong mối quan hệ với quy chế, bộ máy, nguồn lực và dữ liệu của nhà trường. Tiếp cận đa bên xem xét quyền lợi của "
          "nhà trường, giảng viên, người học và doanh nghiệp. Tiếp cận nghiên cứu trường hợp phân tích cụ thể tại Trường Đại "
          "học Thành Đô. Thực trạng và giải pháp được trình bày theo bốn khâu sáng tạo, xác lập, khai thác và bảo vệ nêu tại "
          "Chương 1. Các phương pháp cụ thể gồm:"),
    ("GP", "- Phân tích, tổng hợp tài liệu:"),
    ("GP", "- So sánh:"),
    ("GP", "- Nghiên cứu trường hợp dựa trên dữ liệu hành chính:"),
    ("P", "Dữ liệu được làm sạch, chuẩn hóa tên đơn vị và xác định năm theo mã số, năm phê duyệt, năm nghiệm thu, năm nộp "
          "đơn hoặc năm cấp văn bằng. Do các danh mục dùng đơn vị thống kê khác nhau và không có trường liên kết, 582 bản ghi "
          "được giữ là tổng số bản ghi, không quy đổi thành số sản phẩm độc lập. Hồ sơ tài sản trí tuệ được xếp theo bốn "
          "trạng thái: đã nộp đơn, đã chấp nhận đơn hợp lệ, đã cấp văn bằng hoặc giấy chứng nhận, chưa xác minh. Mức hoàn "
          "thành kế hoạch được tính theo tỷ lệ thực hiện so với chỉ tiêu."),
    ("P", "- **Phương pháp chuyên gia:** thuyết minh dự kiến tham vấn ý kiến chuyên gia, khảo sát và phỏng vấn sâu. Trong "
          "thời gian thực hiện, đề tài chưa tổ chức tham vấn chuyên gia, khảo sát hay phỏng vấn chính thức; các nhận định "
          "được kiểm tra bằng đối chiếu với văn bản gốc và dữ liệu hành chính. Việc lấy ý kiến của chủ nhiệm đề tài, Hội đồng "
          "nghiệm thu và khảo sát nhận thức của giảng viên được đưa vào kế hoạch thí điểm tại Mục 3.2."),
    ("GP", "Giới hạn phương pháp."),
    ("P", "**Quy ước tên gọi văn bản.** Mỗi văn bản được gọi bằng một tên ngắn thống nhất, liệt kê tại Danh mục chữ viết tắt; "
          "số hiệu và ngày ban hành đầy đủ ghi tại Danh mục chữ viết tắt và Tài liệu tham khảo. Điều khoản cụ thể chỉ nêu khi "
          "cần đối chiếu."),
    ("GP", "7. Bố cục của đề tài"),
    ("P", "Ngoài Mở đầu, Kết luận và kiến nghị, Tài liệu tham khảo, báo cáo gồm ba chương: Chương 1. Cơ sở lý luận về quản lý "
          "quyền sở hữu trí tuệ trong trường đại học; Chương 2. Thực trạng quản lý quyền sở hữu trí tuệ tại Trường Đại học "
          "Thành Đô; Chương 3. Giải pháp nâng cao hiệu quả quản lý quyền sở hữu trí tuệ tại Trường Đại học Thành Đô. Cả ba "
          "chương đều trình bày theo bốn khâu sáng tạo, xác lập, khai thác và bảo vệ."),
]

CHUONG_1 = [
    ("CH", 1, ["CƠ SỞ LÝ LUẬN VỀ QUẢN LÝ QUYỀN SỞ HỮU TRÍ TUỆ", "TRONG TRƯỜNG ĐẠI HỌC"]),
    ("H1", "1.1. Khái niệm và vai trò của quyền sở hữu trí tuệ trong trường đại học"),
    ("GP", "Tài sản trí tuệ. Tổ chức Sở hữu trí tuệ thế giới"),
    ("GP", "Quyền sở hữu trí tuệ. Theo Luật Sở hữu trí tuệ"),
    ("P", "Trong trường đại học, tài sản trí tuệ thường được chia thành ba nhóm để quản lý: nhóm quyền tác giả, gồm giáo "
          "trình, bài giảng, bài báo, phần mềm, cơ sở dữ liệu học liệu; nhóm tài sản công nghệ, gồm sáng chế, giải pháp hữu "
          "ích, kiểu dáng công nghiệp, bí mật kinh doanh phát sinh từ đề tài; nhóm tài sản nhận diện, gồm nhãn hiệu, tên "
          "thương mại, biểu trưng của trường. Đây là cách phân nhóm phục vụ quản lý, không thay thế phân loại pháp lý."),
    ("P", "**Quản lý quyền sở hữu trí tuệ trong trường đại học** là các hoạt động của nhà trường nhằm ban hành quy chế, tổ "
          "chức bộ máy và vận hành quy trình để tài sản trí tuệ được sáng tạo, xác lập quyền, khai thác và bảo vệ (Bradley et "
          "al., 2013). Khác với doanh nghiệp, nhà trường phải cân bằng giữa phổ biến tri thức vì lợi ích chung và bảo hộ để "
          "thu hồi chi phí, tái đầu tư (Perkmann et al., 2013). Tài sản trí tuệ trong trường có bốn đặc điểm cần lưu ý: khó "
          "nhận diện nếu không có khai báo; hình thành từ nhiều nguồn kinh phí, quyết định ai là chủ sở hữu; nhiều người cùng "
          "sáng tạo, cần xác định mức đóng góp từ đầu; luôn có áp lực giữa công bố sớm và giữ bí mật để bảo hộ (Shane, 2004; "
          "Võ, 2025)."),
    ("P", "**Vai trò.** Quản lý tốt quyền sở hữu trí tuệ mang lại cho nhà trường ba lợi ích thiết thực:"),
    ("GP", "- Nâng chất lượng đào tạo và nghiên cứu."),
    ("GP", "- Tạo nguồn thu ngoài học phí."),
    ("P", "- **Đáp ứng chuẩn và kiểm định chất lượng.** Chuẩn cơ sở giáo dục đại học tính văn bằng vào sản phẩm khoa học của "
          "trường và tách nguồn thu từ thương mại hóa trong nhóm chỉ số khoa học công nghệ. Thông tư 83, có hiệu lực từ ngày "
          "15 tháng 11 năm 2026, tính một bằng giải pháp hữu ích bằng 3 sản phẩm quy đổi và một bằng sáng chế bằng 5 sản phẩm "
          "quy đổi."),
    ("P", "Giai đoạn 2025 - 2026, khung pháp lý về sở hữu trí tuệ thay đổi nhanh. Luật KH,CN&ĐMST và Nghị định 267 chuyển "
          "nhiều quyền quyết định về tổ chức chủ trì, đồng thời đặt thêm nghĩa vụ đối với tác giả và việc theo dõi kết quả; "
          "Luật Sở hữu trí tuệ được sửa đổi; Luật Giáo dục đại học, Thông tư 83 và Quyết định 1624 đặt yêu cầu mới về quy chế, "
          "dữ liệu và công khai. Các yêu cầu cụ thể được nêu theo từng khâu tại Mục 1.2."),
    ("H1", "1.2. Nội dung quản lý quyền sở hữu trí tuệ"),
    ("P", "Chiến lược sở hữu trí tuệ đến năm 2030 định hướng phát triển đồng bộ các khâu sáng tạo, xác lập, khai thác và bảo "
          "vệ quyền sở hữu trí tuệ. Trong trường đại học, bốn khâu này nối tiếp nhau thành một chu trình (Hình 1.1). Mỗi tiểu "
          "mục dưới đây nêu việc nhà trường cần làm và yêu cầu pháp lý liên quan."),
    ("GK", "Hình 1.1.", {}),
    ("H2", "1.2.1. Sáng tạo tài sản trí tuệ"),
    ("P", "Sáng tạo là khâu hình thành kết quả nghiên cứu, giáo trình, phần mềm và sản phẩm của người học. Nhà trường không "
          "trực tiếp sáng tạo nhưng định hướng và tạo điều kiện qua ba công cụ: giao nhiệm vụ, cấp kinh phí đề tài; hợp tác "
          "với doanh nghiệp; không gian thử nghiệm sản phẩm. Việc quan trọng nhất ở khâu này là xác định ngay khi duyệt đề "
          "tài: đề tài có thể tạo ra đối tượng quyền nào, ai là chủ sở hữu, ai là đồng tác giả."),
    ("P", "**Yêu cầu pháp lý.** Theo Luật KH,CN&ĐMST và Nghị định 267, tổ chức chủ trì ngoài công lập được tự động giao quyền "
          "sở hữu phần kết quả tương ứng với kinh phí ngân sách nhà nước và phải theo dõi riêng các kết quả này. Theo Luật Sở "
          "hữu trí tuệ, nhà trường là chủ sở hữu quyền tài sản đối với tác phẩm giao cho giảng viên biên soạn, trừ khi có "
          "thỏa thuận khác. Quyết định 1624 kế thừa yêu cầu của Quyết định 1068 về xác định trước đối tượng quyền cần đạt với "
          "kết quả dùng ngân sách nhà nước."),
    ("P", "**Yêu cầu quản lý.** Nhà trường cần quy định rõ quyền sở hữu theo từng nguồn hình thành: nhiệm vụ trường giao, đề "
          "tài dùng cơ sở vật chất của trường, sản phẩm của người học, hợp đồng với doanh nghiệp; đồng thời có cơ chế khuyến "
          "khích để giảng viên hướng tới sản phẩm có thể bảo hộ."),
    ("H2", "1.2.2. Xác lập quyền sở hữu trí tuệ"),
    ("P", "Xác lập quyền là làm phát sinh hoặc xác nhận quyền đối với tài sản trí tuệ. Mỗi loại đối tượng có cách xác lập "
          "riêng. Sáng chế, giải pháp hữu ích, kiểu dáng công nghiệp, nhãn hiệu phải nộp đơn và được cấp văn bằng. Quyền tác "
          "giả phát sinh tự động khi tác phẩm được định hình; giấy chứng nhận đăng ký là chứng cứ khi khai thác hoặc tranh "
          "chấp. Bí mật kinh doanh được bảo vệ bằng biện pháp bảo mật, không nộp đơn. Bằng sáng chế có hiệu lực 20 năm, bằng "
          "giải pháp hữu ích 10 năm, kiểu dáng 5 năm và nhãn hiệu 10 năm; kiểu dáng và nhãn hiệu được gia hạn."),
    ("P", "**Yêu cầu pháp lý.** Luật Sở hữu trí tuệ trao cho tổ chức chủ trì nhiệm vụ dùng ngân sách nhà nước quyền đăng ký "
          "sáng chế, kiểu dáng, thiết kế bố trí là kết quả của nhiệm vụ. Sáng chế mất tính mới nếu bị bộc lộ trước khi nộp "
          "đơn; chỉ một số trường hợp bộc lộ được giữ tính mới khi đơn được nộp trong 12 tháng. Quyết định 1624 tiếp tục yêu "
          "cầu trường khối kỹ thuật, công nghệ đăng ký bảo hộ đồng thời với việc công bố kết quả có tính ứng dụng cao."),
    ("P", "**Yêu cầu quản lý.** Nhà trường cần một quy trình khai báo, sàng lọc và quyết định nộp đơn trước khi công bố; có "
          "kinh phí, người làm thủ tục và công cụ theo dõi hồ sơ đến khi được cấp văn bằng. Thủ tục thẩm định đơn kéo dài "
          "nhiều tháng, nên việc theo dõi các mốc và hạn nộp phí là bắt buộc."),
    ("H2", "1.2.3. Khai thác quyền sở hữu trí tuệ"),
    ("P", "Khai thác gồm khai thác phi thương mại, tức đưa giáo trình, kết quả nghiên cứu vào đào tạo và quản trị của trường, "
          "và khai thác thương mại theo bốn hình thức: chuyển nhượng quyền; chuyển quyền sử dụng, thường gọi là cấp phép; góp "
          "vốn bằng quyền sở hữu trí tuệ; trực tiếp sản xuất, cung ứng dịch vụ. Sau khi trả thù lao, thưởng cho tác giả, khoản "
          "thu được đưa vào quỹ phát triển khoa học và công nghệ để tái đầu tư."),
    ("P", "**Yêu cầu pháp lý.** Luật KH,CN&ĐMST cho tổ chức chủ trì tự quyết hình thức, giá và cách chia lợi nhuận khi thương "
          "mại hóa. Với phần kết quả dùng ngân sách nhà nước, tác giả được thưởng tối thiểu 30% lợi nhuận sau thuế; phần không "
          "dùng ngân sách do chủ sở hữu tự quyết; nhiệm vụ phê duyệt trước ngày 01 tháng 10 năm 2025 áp dụng quy định cũ, trừ "
          "lợi nhuận từ văn bằng đã được cấp. Nghị định 267 yêu cầu chia lợi nhuận công khai; đồng tác giả chia thưởng theo "
          "thỏa thuận; tổ chức trung gian hưởng tối thiểu 10% nếu các bên không thỏa thuận khác. Luật Sở hữu trí tuệ quy định "
          "thù lao cho tác giả sáng chế, kiểu dáng, thiết kế bố trí theo thỏa thuận; nếu không có thỏa thuận là 10% lợi nhuận "
          "trước thuế khi tự sử dụng hoặc 15% số tiền nhận được khi chuyển giao. Luật Giáo dục đại học cho phép trường định "
          "giá, góp vốn, thành lập doanh nghiệp quản lý tài sản trí tuệ."),
    ("P", "**Yêu cầu quản lý.** Cần phân biệt bốn khoản lợi ích thường bị nhầm lẫn. **Thù lao** là khoản chủ sở hữu trả cho "
          "tác giả sáng chế, kiểu dáng, thiết kế bố trí theo Luật Sở hữu trí tuệ. **Thưởng** là khoản trích từ lợi nhuận "
          "thương mại hóa theo Luật KH,CN&ĐMST. **Nhuận bút** là khoản trả cho tác giả sách, giáo trình. **Phân chia nguồn "
          "thu** là việc chủ sở hữu phân bổ khoản thu cho trường, đơn vị có tác giả và quỹ. Quy chế phải xác định lợi ích theo "
          "nguồn kinh phí, thời điểm giao nhiệm vụ và loại đối tượng, không thấp hơn mức tối thiểu của luật. Pháp luật chưa "
          "hướng dẫn cách tính lợi nhuận của một kết quả hình thành từ nhiều nguồn kinh phí, nên nhà trường cần tự quy định."),
    ("H2", "1.2.4. Bảo vệ quyền sở hữu trí tuệ"),
    ("P", "Bảo vệ gồm giữ bí mật thông tin nghiên cứu trước khi nộp đơn, kiểm tra trùng lặp, theo dõi sao chép trái phép, duy "
          "trì hiệu lực văn bằng, tự bảo vệ quyền và phối hợp xử lý xâm phạm. Bảo vệ diễn ra suốt vòng đời tài sản, từ khi có "
          "ý tưởng đến khi khai thác, không chỉ sau khi có văn bằng."),
    ("P", "**Yêu cầu pháp lý.** Luật Sở hữu trí tuệ quy định các hành vi xâm phạm và quyền tự bảo vệ của chủ sở hữu. Nghị "
          "định 134 xác định tác phẩm có dùng trí tuệ nhân tạo chỉ được bảo hộ quyền tác giả khi con người đóng góp sáng tạo "
          "trực tiếp. Thông tư 83 yêu cầu cơ sở giáo dục đại học có quy chế quản trị về sở hữu trí tuệ, liêm chính khoa học, "
          "liêm chính học thuật. Chỉ thị 02 yêu cầu tăng cường thực thi quyền sở hữu trí tuệ."),
    ("P", "**Yêu cầu quản lý.** Nhà trường cần quy định nghĩa vụ bảo mật và cam kết bảo mật khi hợp tác, thử nghiệm sản phẩm; "
          "quy định kiểm tra trùng lặp và xử lý vi phạm liêm chính; lịch theo dõi hạn nộp phí duy trì văn bằng; đầu mối tiếp "
          "nhận, xử lý khi phát hiện xâm phạm."),
    ("TK", "TIỂU KẾT CHƯƠNG 1"),
    ("P", "Quản lý quyền sở hữu trí tuệ trong trường đại học là quản lý một chu trình bốn khâu: sáng tạo, xác lập, khai thác "
          "và bảo vệ. Mỗi khâu có việc riêng mà nhà trường phải làm, từ xác định đối tượng quyền khi duyệt đề tài, sàng lọc và "
          "nộp đơn trước khi công bố, phân chia lợi ích khi khai thác, đến bảo mật, liêm chính và duy trì văn bằng."),
    ("P", "Khung pháp lý 2025 - 2026 chuyển nhiều quyền quyết định về nhà trường, đồng thời đặt thêm nghĩa vụ: thưởng tối "
          "thiểu cho tác giả, ghi nhận mức đóng góp, theo dõi riêng kết quả từ ngân sách, quy chế về liêm chính học thuật. Bốn "
          "khâu và các yêu cầu này là căn cứ để đánh giá thực trạng tại Chương 2 và đề xuất giải pháp tại Chương 3."),
]

CHUONG_2 = [
    ("CH", 2, ["THỰC TRẠNG QUẢN LÝ QUYỀN SỞ HỮU TRÍ TUỆ", "TẠI TRƯỜNG ĐẠI HỌC THÀNH ĐÔ"]),
    ("P", "Chương này đánh giá thực trạng quản lý quyền sở hữu trí tuệ tại Trường Đại học Thành Đô giai đoạn 2021 - 2025 theo "
          "bốn khâu nêu tại Mục 1.2. Dữ liệu lấy từ các danh mục thống kê của Phòng Khoa học Công nghệ, danh sách nhân sự năm "
          "2026, tệp theo dõi đơn và văn bằng, các quy chế nội bộ và Kế hoạch 07. Văn bản ban hành sau năm 2025 chỉ dùng để mô "
          "tả yêu cầu hiện hành, không dùng để giải thích kết quả của kỳ đánh giá."),
    ("H1", "2.1. Khái quát về Trường Đại học Thành Đô"),
    ("GP", "Trường Đại học Thành Đô là trường đại học tư thục"),
    ("GP", "Năm 2026, Nhà trường có 252 nhân sự"),
    ("GK", "Bảng 2.1.", {}),
    ("P", "Nhân lực trình độ cao tập trung ở ba viện đào tạo. Ba viện chiếm 57,9% nhân sự nhưng có 81 trên 94 tiến sĩ, tức "
          "86,2%, và 22 trên 25 người có học hàm. Đây là lực lượng có thể nhận diện kết quả nghiên cứu cần bảo hộ. Ngược lại, "
          "thủ tục xác lập quyền, quản trị thương hiệu và pháp chế do khối Quản trị và Dịch vụ đảm nhận."),
    ("P", "**Quy chế nội bộ.** Nhà trường có bốn văn bản chứa quy định về sở hữu trí tuệ:"),
    ("GP", "- Quyết định 213 năm 2021."),
    ("GP", "- Quyết định 217 ngày 21 tháng 11 năm 2024."),
    ("GP", "- Quy chế chi tiêu nội bộ ban hành ngày 01 tháng 8 năm 2026"),
    ("GP", "- Điều lệ Quỹ Học bổng sau tiến sĩ Ngô Xuân Độ năm 2025."),
    ("P", "Nhà trường ban hành quy chế về sở hữu trí tuệ từ năm 2021 và quy chế chuyên biệt từ năm 2024, trước khi các luật "
          "mới năm 2025 ra đời. Đây là tầm nhìn sớm. Quyết định 217 không thay thế Chương VI Quyết định 213, nên hai văn bản "
          "cùng hiệu lực. Nội dung của các văn bản này được đánh giá theo từng khâu tại Mục 2.2."),
    ("P", "**Bộ máy quản lý.** Theo Quyết định 213 và Quyết định 217, Phòng Khoa học Công nghệ là đầu mối tiếp nhận, nhận diện, "
          "theo dõi và xúc tiến thương mại hóa; Bộ phận Pháp chế làm thủ tục xác lập quyền. Thực tế, công việc chia cho bốn "
          "đơn vị (Hình 2.1). Phòng Khoa học Công nghệ tiếp nhận hồ sơ đề tài, có 2 nhân sự trình độ thạc sĩ kiêm nhiều mảng "
          "việc và chưa có vị trí chuyên trách sở hữu trí tuệ. Bộ phận quản trị thương hiệu thuộc Trung tâm Tuyển sinh và Quản "
          "trị thương hiệu lập danh mục nhãn hiệu, biểu trưng. Bộ phận Pháp chế thuộc Trung tâm Dịch vụ và Quản trị hành chính "
          "tổng hợp; theo thông tin được cung cấp, đây là đầu mối của Mạng lưới Trung tâm Hỗ trợ công nghệ và đổi mới sáng tạo "
          "từ năm 2023, nhưng chưa có văn bản xác nhận trong hồ sơ. Viện Nghiên cứu giáo dục và Chuyển giao tri thức thống kê "
          "hợp đồng khai thác sách."),
    ("GK", "Hình 2.3.", {"Hình 2.3.": "Hình 2.1."}),
    ("GP", "Ba trong bốn đầu mối thuộc khối Quản trị và Dịch vụ"),
    ("H1", "2.2. Thực trạng quản lý quyền sở hữu trí tuệ"),
    ("H2", "2.2.1. Sáng tạo tài sản trí tuệ"),
    ("P", "**Kết quả nghiên cứu.** Các danh mục thống kê ghi nhận 582 bản ghi sản phẩm khoa học trong năm năm (Bảng 2.2)."),
    ("GK", "Bảng 2.2.", {}),
    ("GP", "Con số 582 là tổng số bản ghi"),
    ("GK", "Hình 2.1.", {"Hình 2.1.": "Hình 2.2."}),
    ("GPS", "Hình 2.1 cho thấy số bản ghi tăng", {"Hình 2.1 cho thấy": "Hình 2.2 cho thấy"}),
    ("GPS", "Với quản lý quyền sở hữu trí tuệ, cơ cấu này có hai hàm ý.",
     {", không chỉ năng lực nghiên cứu (Mục 2.3.3).": ", không chỉ năng lực nghiên cứu."}),
    ("GP", "Trong hai năm cuối kỳ, Nhà trường chủ trì ba đề tài cấp quốc gia"),
    ("P", "**Kinh phí cho đề tài.** Nhà trường có ba kênh tài trợ nghiên cứu với cơ chế khác nhau (Bảng 2.3)."),
    ("GK", "Bảng 2.4.", {"Bảng 2.4.": "Bảng 2.3."}),
    ("P", "Đề tài cấp cơ sở là kênh duy nhất trong kỳ tạo ra sản phẩm có tiềm năng, được cấp 424,75 triệu đồng trong năm năm. "
          "Quỹ Học bổng sau tiến sĩ Ngô Xuân Độ mới hoạt động từ năm 2025; ba đề tài cấp quốc gia đang thực hiện. Hình 2.3 "
          "cho thấy số đề tài cấp cơ sở theo hình thức kinh phí qua các năm."),
    ("GK", "Hình 2.4.", {"Hình 2.4.": "Hình 2.3."}),
    ("GP", "Cơ chế cấp kinh phí thay đổi sau năm 2022."),
    ("P", "**Sản phẩm đề tài có tiềm năng bảo hộ.** Rà soát sản phẩm nghiệm thu của 38 đề tài cấp cơ sở, 11 đề tài có sản "
          "phẩm cụ thể ngoài báo cáo và bài báo, có tiềm năng tạo lập tài sản trí tuệ (Bảng 2.4). Đây là đánh giá sơ bộ theo mô "
          "tả sản phẩm. Khả năng bảo hộ thực tế còn phụ thuộc vào tính mới, trình độ sáng tạo và tình trạng bộc lộ, chưa được "
          "thẩm định."),
    ("GK", "Bảng 2.7.", {"Bảng 2.7.": "Bảng 2.4."}),
    ("GK", "Hình 2.8.", {"Hình 2.8.": "Hình 2.4.", "Bảng 2.7": "Bảng 2.4"}),
    ("GPS", "Hình 2.8 cho thấy 9 trên 11", {"Hình 2.8 cho thấy": "Hình 2.4 cho thấy"}),
    ("P", "Biểu mẫu đề xuất, thuyết minh tại Quyết định 213 đã có mục về sản phẩm dự kiến và đăng ký sở hữu trí tuệ, nhưng "
          "chưa yêu cầu xác định chủ thể quyền, đồng tác giả, kể cả người học và doanh nghiệp, và lịch công bố dự kiến. Quyền "
          "lợi của các bên vì vậy chưa được xác lập từ lúc duyệt đề tài."),
    ("P", "**Nhận thức về sở hữu trí tuệ.** Đề tài không khảo sát trực tiếp. Một số dữ kiện cho thấy kỹ năng nhận diện và bảo "
          "hộ tài sản trí tuệ của đội ngũ còn hạn chế: 6 đề tài mã số 2021 - 2024 có sản phẩm tiềm năng sở hữu công nghiệp "
          "nhưng chưa có đơn; hai đề tài năm 2025 ghi dự kiến đăng ký giải pháp hữu ích nhưng chưa có đơn; 87 giáo trình chưa "
          "được đăng ký quyền tác giả; Kế hoạch 07 thống kê 12 lượt tập huấn chung giai đoạn 2019 - 2023 nhưng không tách riêng "
          "tập huấn về sở hữu trí tuệ. Ở chiều ngược lại, đơn sáng chế từ đề tài chiết xuất lá Quế hoa và 5 kiểu dáng công "
          "nghiệp hợp tác với doanh nghiệp cho thấy một nhóm giảng viên đã có ý thức bảo hộ."),
    ("H2", "2.2.2. Xác lập quyền sở hữu trí tuệ"),
    ("P", "**Quy trình và kinh phí.** Quyết định 213 quy định đăng ký một cửa: tác giả nộp hồ sơ tại Phòng Khoa học Công nghệ, "
          "Phòng trình Hiệu trưởng ký và làm thủ tục nộp đơn, lệ phí; quy chế không đặt điều kiện đã nghiệm thu mới được nộp "
          "đơn. Quyết định 217 yêu cầu tác giả xin ý kiến Phòng trước khi công bố tài sản có thể bảo hộ. Các biểu mẫu đề tài "
          "đã có mục về đăng ký sở hữu trí tuệ, nhưng chưa yêu cầu đánh giá khả năng bảo hộ và tình trạng bộc lộ. Về kinh phí, "
          "Quyết định 213 cho phép chi lệ phí và thuê ngoài, Quyết định 217 coi lệ phí xác lập quyền là khoản được trừ khi chia "
          "lợi ích. Điều còn thiếu là dự toán riêng cho phí nộp đơn, phí đại diện, phí duy trì; cách tạm ứng khi đơn nộp sau "
          "nghiệm thu; và người chịu trách nhiệm đề xuất chi. Số chi thực tế cho xác lập quyền cũng chưa được thống kê."),
    ("GP", "Về khuyến khích, Quy chế chi tiêu nội bộ năm 2026"),
    ("GK", "Hình 2.5.", {}),
    ("GP", "Xét riêng giờ quy đổi, văn bằng được tính cao"),
    ("P", "**Kết quả xác lập.** Bảng 2.5 đối sánh ba lớp: đối tượng quyền theo Luật Sở hữu trí tuệ, tài sản được quy chế nội "
          "bộ liệt kê, và thực tế phát sinh giai đoạn 2021 - 2025."),
    ("GK", "Bảng 2.5.", {}),
    ("GP", "Nhà trường phát sinh tài sản ở nhiều nhóm"),
    ("GPS", "Tệp theo dõi ghi nhận 11 hồ sơ", {"(Bảng 2.6, Hình 2.7)": "(Bảng 2.6, Hình 2.6)"}),
    ("GK", "Bảng 2.6.", {}),
    ("GK", "Hình 2.7.", {"Hình 2.7.": "Hình 2.6."}),
    ("GP", "Mười một hồ sơ trong kỳ hình thành từ ba luồng:"),
    ("GP", "- Luồng thương hiệu:"),
    ("GP", "- Luồng hợp tác doanh nghiệp:"),
    ("GP", "- Luồng nghiên cứu:"),
    ("P", "Số văn bằng, giấy chứng nhận xác định được trong kỳ là 4; nếu 5 kiểu dáng được xác nhận, con số là 9. Tổng số hồ sơ "
          "không được hiểu là tổng số quyền đã được cấp. Giai đoạn tới, Nhà trường có thể tận dụng đà hợp tác để tạo thêm tài "
          "sản do chính mình sở hữu từ kết quả nghiên cứu, nhất là công thức và quy trình, những đối tượng mà kiểu dáng công "
          "nghiệp không bảo hộ."),
    ("P", "**Chuyển hóa từ đề tài sang đơn đăng ký.** Hình 2.7 theo dõi chuỗi từ đề tài cấp cơ sở đến đơn đăng ký sở hữu công "
          "nghiệp."),
    ("GK", "Hình 2.9.", {"Hình 2.9.": "Hình 2.7.", "Bảng 2.7": "Bảng 2.4"}),
    ("GP", "Trong chuỗi sở hữu công nghiệp, 9 trên 38 đề tài"),
    ("GP", "Đề tài chiết xuất lá Quế hoa mã số 09-2025"),
    ("P", "**Công cụ theo dõi hồ sơ.** Các danh mục thống kê của Phòng Khoa học Công nghệ chưa có danh mục tài sản trí tuệ đã "
          "đăng ký hoặc được cấp văn bằng, dù Quyết định 217 đã giao nhiệm vụ này. Tệp theo dõi đơn và văn bằng hiện có gồm "
          "hai bảng chưa thống nhất trạng thái của 5 kiểu dáng công nghiệp và không liên kết với danh mục đề tài. Danh mục bài "
          "báo không ghi mã đề tài, nên chưa biết bao nhiêu bài phát sinh từ đề tài. Vì vậy, Nhà trường đếm được số đơn và "
          "giấy chứng nhận, nhưng chưa theo dõi được thời gian xử lý, chi phí và giá trị mà các quyền mang lại. Từ năm 2026, "
          "đây lại là thông tin phải công khai theo Luật Giáo dục đại học."),
    ("H2", "2.2.3. Khai thác quyền sở hữu trí tuệ"),
    ("P", "**Quy định về lợi ích và phân chia nguồn thu.** Bảng 2.7 tập hợp các quy định nội bộ và quy định của luật về lợi "
          "ích của tác giả và phân chia nguồn thu."),
    ("GK", "Bảng 2.3.", {"Bảng 2.3.": "Bảng 2.7."}),
    ("GPS", "Các tỷ lệ trong Bảng 2.3 không so sánh trực tiếp",
     {"Các tỷ lệ trong Bảng 2.3": "Các tỷ lệ trong Bảng 2.7",
      " Chỉ điểm a và điểm b của Quyết định 213 có cùng cơ sở tính, nên được mô phỏng chung tại Hình 2.2.":
      " Chỉ điểm a và điểm b của Quyết định 213 có cùng cơ sở tính."}),
    ("GPS", "Trên cùng nguồn thu, phần của tác giả", {"Trên cùng nguồn thu,": "Khi tính trên cùng nguồn thu,"}),
    ("GP", "Chưa thấy trường hợp nào một tài sản"),
    ("GP", "- Quan hệ giữa quy định chia lợi ích"),
    ("GP", "- Quan hệ giữa mức trích 50%"),
    ("GPS", "- Thuật ngữ: điểm a gọi là khen thưởng", {"(Mục 1.3.2)": "(Mục 1.2.3)"}),
    ("GP", "Đối chiếu với luật hiện hành, có ba điểm cần rà soát."),
    ("P", "**Kết quả khai thác.** Khai thác tài sản trí tuệ được xem xét ở ba mức."),
    ("GP", "- Sử dụng nội bộ."),
    ("GP", "- Chuyển giao trong hệ thống."),
    ("GP", "- Khai thác có thu phí."),
    ("GP", "Như vậy, khai thác chủ yếu ở dạng phi thương mại"),
    ("H2", "2.2.4. Bảo vệ quyền sở hữu trí tuệ"),
    ("P", "Bảo vệ quyền là khâu có ít dữ liệu nhất. Hồ sơ hiện có cho thấy năm điểm:"),
    ("P", "- **Quy định đã có.** Quyết định 217 quy định các hành vi xâm phạm quyền tác giả như mạo danh, công bố trái phép, "
          "trích dẫn không đầy đủ; yêu cầu tác giả giữ bí mật và xin ý kiến trước khi công bố tài sản có thể bảo hộ."),
    ("P", "- **Bảo mật trước khi công bố.** Hồ sơ chưa có phiếu hay sổ ghi nhận việc xin ý kiến trước khi công bố, nên mức độ "
          "thực hiện quy định này chưa đánh giá được."),
    ("P", "- **Tranh chấp, xâm phạm.** Hồ sơ không ghi nhận tranh chấp, khiếu nại hay xử lý xâm phạm liên quan đến Nhà trường. "
          "Điều này chưa đủ để kết luận không có xâm phạm, vì Nhà trường chưa có cơ chế theo dõi."),
    ("P", "- **Duy trì hiệu lực văn bằng.** Chưa có sổ theo dõi hạn nộp phí duy trì và gia hạn. Nếu 5 kiểu dáng công nghiệp "
          "được cấp cùng năm, thời hạn gia hạn sẽ trùng nhau."),
    ("P", "- **Liêm chính học thuật và trí tuệ nhân tạo.** Trong Quyết định 217 chưa thấy quy định về kiểm tra trùng lặp và "
          "công khai việc dùng trí tuệ nhân tạo, trong khi Thông tư 83 đã yêu cầu quy chế quản trị về liêm chính khoa học, "
          "liêm chính học thuật."),
    ("H1", "2.3. Đánh giá chung"),
    ("GP", "Kết quả hai năm 2024 - 2025 được đối chiếu với Kế hoạch 07."),
    ("GK", "Hình 2.10.", {"Hình 2.10.": "Hình 2.8."}),
    ("GPS", "Hình 2.10 cho thấy 6 trên 9", {"Hình 2.10 cho thấy": "Hình 2.8 cho thấy"}),
    ("GP", "Chỉ tiêu 1.11 về công nhận sáng chế"),
    ("H2", "2.3.1. Kết quả đạt được"),
    ("GP", "Thứ nhất, Nhà trường có tầm nhìn sớm về thể chế."),
    ("GP", "Thứ hai, trong kỳ có 11 hồ sơ tài sản trí tuệ."),
    ("GP", "Thứ ba, năng lực nghiên cứu tăng nhanh."),
    ("GP", "Thứ tư, đội ngũ có tiềm năng tạo lập tài sản trí tuệ"),
    ("GP", "Thứ năm, biểu mẫu đề tài tại Quyết định 213"),
    ("H2", "2.3.2. Hạn chế và nguyên nhân"),
    ("P", "Hạn chế được xếp theo bốn khâu, mỗi hạn chế kèm nguyên nhân. Các nguyên nhân chủ quan là nhận định rút ra từ hồ sơ "
          "hiện có, cần được kiểm chứng thêm trong quá trình thí điểm."),
    ("P", "**Ở khâu sáng tạo.** Kết quả nghiên cứu tăng nhanh nhưng chủ yếu là bài báo; quyền sở hữu và quyền lợi của đồng tác "
          "giả, người học, doanh nghiệp chưa được xác định từ lúc duyệt đề tài; kỹ năng nhận diện tài sản có thể bảo hộ còn "
          "hạn chế. Nguyên nhân: biểu mẫu đề tài chưa yêu cầu xác định chủ thể quyền và lịch công bố; chưa có tập huấn riêng về "
          "sở hữu trí tuệ; cơ chế khuyến khích năm 2026 thưởng bài báo ngay khi đăng nhưng chỉ ghi nhận văn bằng khi được cấp."),
    ("P", "**Ở khâu xác lập.** Đây là hạn chế lớn nhất: 9 đề tài có sản phẩm tiềm năng sở hữu công nghiệp nhưng mới 1 đề tài "
          "có đơn; 6 đề tài mã số 2021 - 2024 chưa có đơn; tài sản do Nhà trường đơn sở hữu từ kết quả nghiên cứu còn ít; 2 sản "
          "phẩm thuộc quyền tác giả chưa đăng ký; dữ liệu theo dõi hồ sơ rời rạc. Nguyên nhân có thể là: chưa có bước sàng lọc "
          "khả năng bảo hộ trước khi công bố, nên việc khởi động thủ tục phụ thuộc vào chủ nhiệm đề tài; chưa có dự toán và "
          "người đề xuất chi cho bước đăng ký; đầu mối kiêm nhiệm, chưa có chuyên trách; chưa có danh mục liên kết đề tài với "
          "tài sản trí tuệ."),
    ("P", "**Ở khâu khai thác.** Khai thác có thu phí còn nhỏ, mới có hai hợp đồng chuyển giao quyền sử dụng tác phẩm; văn bằng "
          "hiện có chưa có giao dịch chuyển giao. Quy định về lợi ích của tác giả chưa rõ thứ tự áp dụng, chưa phân biệt thưởng "
          "với thù lao; mức trần tại Quyết định 213 cần rà soát theo Luật KH,CN&ĐMST. Nguyên nhân: các quy chế ban hành ở những "
          "thời điểm khác nhau, trước khi luật mới ra đời; chưa có bộ phận định giá, xúc tiến chuyển giao; số tài sản sẵn sàng "
          "chuyển giao còn ít."),
    ("P", "**Ở khâu bảo vệ.** Chưa có cơ chế theo dõi xâm phạm, sổ theo dõi hiệu lực văn bằng, quy định về kiểm tra trùng lặp "
          "và dùng trí tuệ nhân tạo. Nguyên nhân có thể là số văn bằng còn ít nên việc theo dõi chưa được đặt thành yêu cầu, và "
          "trách nhiệm theo dõi chưa được giao cụ thể."),
    ("P", "**Nguyên nhân khách quan.** Pháp luật thay đổi nhanh trong giai đoạn 2025 - 2026, nên các quy chế ban hành năm 2021 "
          "và 2024 cần rà soát; thủ tục xác lập quyền sở hữu công nghiệp kéo dài và tốn chi phí, Nhà trường tự cân đối từ nguồn "
          "thu; cơ cấu ngành chủ yếu tạo ra tác phẩm thuộc quyền tác giả, tiềm năng sở hữu công nghiệp tập trung ở lĩnh vực "
          "dược. Các nguyên nhân này giải thích vì sao quy mô tài sản còn nhỏ, nhưng chưa giải thích vì sao 6 đề tài có sản "
          "phẩm tiềm năng lại chưa có đơn."),
    ("TK", "TIỂU KẾT CHƯƠNG 2"),
    ("P", "Giai đoạn 2021 - 2025, Nhà trường có tầm nhìn thể chế sớm, năng lực công bố tăng nhanh, 11 hồ sơ tài sản trí tuệ và "
          "trường hợp đầu tiên chuyển đề tài thành đơn sáng chế. Tiềm năng rõ nhất ở khối Y - Dược: 11 trên 38 đề tài có sản "
          "phẩm tiềm năng, trong đó 9 thuộc sở hữu công nghiệp, nhưng mới 1 đề tài có đơn."),
    ("P", "Xét theo bốn khâu, điểm yếu lớn nhất nằm ở khâu xác lập: chưa có bước sàng lọc trước khi công bố, chưa có dự toán "
          "cho bước đăng ký và chưa có công cụ theo dõi hồ sơ. Khâu khai thác vướng ở quy định về lợi ích chưa thống nhất; khâu "
          "bảo vệ chưa được theo dõi. Đây là căn cứ để Chương 3 đề xuất giải pháp theo từng khâu."),
]

CHUONG_3 = [
    ("CH", 3, ["GIẢI PHÁP NÂNG CAO HIỆU QUẢ QUẢN LÝ QUYỀN SỞ HỮU TRÍ TUỆ", "TẠI TRƯỜNG ĐẠI HỌC THÀNH ĐÔ"]),
    ("P", "Các giải pháp được xây dựng trên ba căn cứ: các hạn chế tại Mục 2.3.2; yêu cầu pháp lý nêu tại Mục 1.2; và điều kiện "
          "thực tế của một trường tư thục định hướng ứng dụng. Nguyên tắc chung là dùng nguồn lực sẵn có, hoàn thiện quy chế "
          "và biểu mẫu hiện hành thay vì thay thế, việc cần nhiều nguồn lực chỉ làm khi đủ điều kiện."),
    ("P", "**Mục tiêu đến năm 2030:** ban hành Quy chế quản trị tài sản trí tuệ sửa đổi trước ngày 31 tháng 5 năm 2027; từ năm "
          "2027, mọi đề tài nghiệm thu có phiếu rà soát; số đơn sở hữu công nghiệp từ kết quả nghiên cứu đạt từ 2 đơn năm 2027 "
          "và 3 đến 5 đơn mỗi năm từ năm 2028; có ít nhất 1 hợp đồng chuyển giao hoặc cấp phép mỗi năm; dữ liệu tài sản trí tuệ "
          "đủ để báo cáo trên HEMIS. Các mức này là mục tiêu tham khảo, được điều chỉnh sau thí điểm."),
    ("H1", "3.1. Các giải pháp đề xuất"),
    ("P", "Bốn giải pháp tương ứng với bốn khâu tại Chương 2. Mỗi giải pháp nêu vấn đề cần giải quyết, việc cần làm, đơn vị "
          "thực hiện và sản phẩm đầu ra."),

    ("H2", "3.1.1. Nâng cao hiệu quả hoạt động sáng tạo và nhận diện tài sản trí tuệ"),
    ("P", "**Vấn đề.** Kết quả nghiên cứu tăng nhanh nhưng chủ yếu là bài báo; quyền sở hữu, đồng tác giả chưa được xác định từ "
          "đầu; kỹ năng nhận diện tài sản có thể bảo hộ còn hạn chế (Mục 2.2.1)."),
    ("P", "**Việc cần làm:**"),
    ("P", "- **Xác định đối tượng quyền ngay khi duyệt đề tài.** Hội đồng xét duyệt xem xét các thông tin về đối tượng quyền dự "
          "kiến, chủ thể quyền và đồng tác giả trong thuyết minh, được bổ sung theo Khâu 1 của quy trình tại Mục 3.1.2. Với đề "
          "tài có người học hoặc doanh nghiệp tham gia, quyền sở hữu và quyền lợi được ghi rõ trong hợp đồng ngay từ đầu."),
    ("P", "- **Ưu tiên lĩnh vực có tiềm năng.** Khi xét duyệt đề tài cấp cơ sở, ưu tiên đề tài có sản phẩm có thể bảo hộ, nhất "
          "là lĩnh vực dược của Viện Y - Dược; theo dõi riêng kết quả của ba đề tài cấp quốc gia để chuẩn bị đăng ký."),
    ("P", "- **Khuyến khích sáng tạo có định hướng bảo hộ.** Sửa Quy chế chi tiêu nội bộ để ghi nhận 50% giờ quy đổi khi đơn "
          "được chấp nhận hợp lệ và 50% khi được cấp văn bằng; xem xét thưởng tiền cho bằng sáng chế, giải pháp hữu ích. Khoản "
          "thưởng nội bộ này tách biệt với thưởng khi thương mại hóa và thù lao theo luật."),
    ("P", "- **Đào tạo và nâng cao nhận thức.** Tập huấn bắt buộc cho giảng viên mới trong 06 tháng đầu về nhận diện đối tượng "
          "bảo hộ và bộc lộ an toàn trước khi công bố; đưa học phần sở hữu trí tuệ vào chương trình đào tạo từ năm học 2027 - "
          "2028; phát hành cẩm nang sở hữu trí tuệ; tổ chức Tháng Sở hữu trí tuệ và Đổi mới sáng tạo vào tháng 4 hằng năm."),
    ("P", "- **Không gian sáng tạo mở thử nghiệm.** Dùng không gian này để giảng viên, người học và doanh nghiệp cùng hoàn thiện "
          "sản phẩm, với điều kiện mọi người tham gia ký cam kết bảo mật hoặc Nhà trường nộp đơn trước khi thử nghiệm rộng."),
    ("P", "**Thực hiện:** Phòng Khoa học Công nghệ chủ trì sửa biểu mẫu, phối hợp Phòng Đào tạo xây dựng học phần; Phòng Tài "
          "chính - Kế toán đề xuất sửa Quy chế chi tiêu nội bộ; Viện Nghiên cứu giáo dục và Chuyển giao tri thức vận hành "
          "Không gian sáng tạo mở thử nghiệm. Thời gian: từ quý IV năm 2026."),
    ("P", "**Sản phẩm:** biểu mẫu đề xuất, thuyết minh, hợp đồng đã sửa; chương trình tập huấn, cẩm nang; đề xuất sửa Quy chế "
          "chi tiêu nội bộ."),

    ("H2", "3.1.2. Sàng lọc và hỗ trợ xác lập quyền sở hữu trí tuệ"),
    ("P", "**Vấn đề.** 9 đề tài có sản phẩm tiềm năng sở hữu công nghiệp nhưng mới 1 đề tài có đơn; chưa có bước sàng lọc trước "
          "khi công bố; chưa có dự toán và người đề xuất chi; đầu mối kiêm nhiệm; dữ liệu theo dõi hồ sơ rời rạc (Mục 2.2.2)."),
    ("P", "**Việc cần làm:**"),
    ("P", "- **Ban hành quy trình 8 khâu từ khai báo đến khai thác** (Hình 3.1), có hai luồng. **Luồng sớm** dành cho kết quả "
          "cần nộp đơn trước khi công bố: hồ sơ được quyết định đăng ký và cấp kinh phí ngay trong quá trình nghiên cứu, không "
          "chờ nghiệm thu. **Luồng thường** dành cho các kết quả còn lại, được kiểm tra tại nghiệm thu. Cách làm này phù hợp "
          "Quyết định 213, vốn không yêu cầu đã nghiệm thu mới được nộp đơn."),
    ("GK", "Hình 3.2.", {"Hình 3.2.": "Hình 3.1."}),
    ("P", "Nội dung từng khâu như sau:"),
    ("GP", "- Khâu 1. Khai báo."),
    ("GP", "- Khâu 2. Sàng lọc và phân nhánh"),
    ("GP", "- Khâu 3. Tra cứu và đánh giá khả năng bảo hộ."),
    ("GP", "- Khâu 4. Xem xét bảo mật trước khi công bố."),
    ("GP", "- Khâu 5. Kiểm tra tại nghiệm thu."),
    ("GP", "- Khâu 6. Quyết định đăng ký và cấp kinh phí."),
    ("GP", "- Khâu 7. Nộp đơn, theo dõi và duy trì."),
    ("GPS", "- Khâu 8. Khai thác và chia lợi ích.", {"theo các lớp tại Giải pháp 1": "theo các lớp tại Mục 3.1.3"}),
    ("P", "- **Bố trí kinh phí cho bước đăng ký.** Lập dòng dự toán hằng năm cho phí xác lập quyền trong Quỹ nghiên cứu khoa học "
          "từ năm 2027; chi cả cho đơn nộp trước nghiệm thu hoặc sau khi đề tài đã quyết toán, không trừ vào kinh phí đề tài. "
          "Phòng Khoa học Công nghệ đề xuất chi trong 15 ngày làm việc; Bộ phận Pháp chế được tạm ứng. Quy mô dự toán tính theo "
          "số hồ sơ dự kiến nhân mức phí hiện hành. Đề nghị bổ sung Điều lệ Quỹ Học bổng sau tiến sĩ Ngô Xuân Độ để hỗ trợ phí "
          "đăng ký cho sản phẩm đề tài cấp cơ sở."),
    ("P", "- **Giao đầu mối và cơ chế phối hợp.** Giai đoạn đầu, giao một nhân sự của Phòng Khoa học Công nghệ và một nhân sự "
          "của Bộ phận Pháp chế làm đầu mối kiêm nhiệm, cử tập huấn tại Cục Sở hữu trí tuệ; thuê tổ chức đại diện sở hữu công "
          "nghiệp cho hồ sơ sáng chế, giải pháp hữu ích. Ban hành quy chế phối hợp theo mô hình tại Hình 3.2, giao ban hằng "
          "quý. Xem xét vị trí chuyên trách khi khối lượng đạt ngưỡng, ví dụ từ 15 kết quả được sàng lọc hoặc 5 đơn mỗi năm."),
    ("GK", "Hình 3.1.", {"Hình 3.1.": "Hình 3.2.", "Mục 2.2.2": "Mục 2.1"}),
    ("P", "- **Lập danh mục số tài sản trí tuệ.** Mỗi tài sản có một mã, liên kết với mã đề tài, ghi chủ thể quyền, đồng tác "
          "giả, nguồn kinh phí, ngày bộc lộ, số đơn, trạng thái theo bốn nhóm thống nhất và hạn nộp phí duy trì; kết quả từ "
          "ngân sách nhà nước được đánh dấu để theo dõi riêng theo Nghị định 267. Giai đoạn đầu dùng bảng tính dùng chung, "
          "không cần mua phần mềm; hoàn thành trước ngày 31 tháng 5 năm 2027 để phục vụ báo cáo trên HEMIS."),
    ("P", "- **Xử lý ngay các hồ sơ tồn đọng.** Rà soát tình trạng bộc lộ của 6 đề tài mã số 2021 - 2024 và hai đề tài 06-2025, "
          "07-2025 để xác định sản phẩm còn đăng ký được; xác minh trạng thái 5 kiểu dáng công nghiệp; rà soát 87 giáo trình, "
          "chỉ đăng ký quyền tác giả cho tài liệu cần chứng cứ khi khai thác."),
    ("P", "**Thực hiện:** Phòng Khoa học Công nghệ chủ trì; Bộ phận Pháp chế làm thủ tục, nộp đơn; Hội đồng nghiệm thu kiểm tra "
          "tại nghiệm thu; Hiệu trưởng và Phòng Tài chính - Kế toán quyết định, cấp kinh phí. Thời gian: thí điểm từ quý IV năm "
          "2026, áp dụng toàn trường từ quý III năm 2027."),
    ("P", "**Sản phẩm:** quy trình 8 khâu, phiếu khai báo, phiếu rà soát; dòng dự toán phí xác lập quyền; quyết định giao đầu "
          "mối và quy chế phối hợp; danh mục số tài sản trí tuệ."),

    ("H2", "3.1.3. Thúc đẩy khai thác quyền sở hữu trí tuệ"),
    ("P", "**Vấn đề.** Khai thác có thu phí mới có hai hợp đồng; quy định về lợi ích của tác giả chưa rõ thứ tự áp dụng, chưa "
          "phân biệt thưởng với thù lao; mức trần tại Quyết định 213 cần rà soát theo luật mới (Mục 2.2.3)."),
    ("P", "**Việc cần làm:**"),
    ("GP", "- Lợi ích của tác giả theo ba lớp."),
    ("P", "- **Ban hành mẫu thỏa thuận chia thưởng giữa đồng tác giả** theo Nghị định 267, dùng chung với văn bản xác định mức "
          "đóng góp của các thành viên lập khi nghiệm thu."),
    ("P", "- **Tổ chức khai thác theo giai đoạn.** Giai đoạn 2026 - 2028, định giá và đàm phán qua Phòng Khoa học Công nghệ, "
          "Bộ phận Pháp chế và tổ chức tư vấn thuê ngoài, có Viện Nghiên cứu giáo dục và Chuyển giao tri thức phối hợp. Hợp "
          "đồng với tổ chức trung gian phải thỏa thuận rõ mức hưởng, vì nếu không, Nghị định 267 dành cho tổ chức này tối "
          "thiểu 10%. Trung tâm tư vấn, định giá hoặc doanh nghiệp quản lý tài sản trí tuệ chỉ lập khi đã có văn bằng có đối "
          "tác quan tâm, có ít nhất một hợp đồng chuyển giao và phương án tài chính tự trang trải được."),
    ("P", "- **Tận dụng hợp tác doanh nghiệp và chương trình thí điểm định giá.** Phát triển kênh hợp tác đã tạo ra 5 kiểu dáng "
          "công nghiệp thành kênh thương mại hóa sản phẩm dược liệu; chuẩn bị hồ sơ tham gia chương trình thí điểm xác định "
          "giá trị ít nhất 100 quyền sở hữu trí tuệ theo Quyết định 1624."),
    ("P", "- **Giới thiệu tài sản sẵn sàng chuyển giao** trên chuyên trang của cổng thông tin Nhà trường, chỉ đăng thông tin "
          "được phép công bố."),
    ("P", "**Thực hiện:** Bộ phận Pháp chế soạn thảo quy định về lợi ích trong Quy chế sửa đổi; Phòng Tài chính - Kế toán mô "
          "phỏng tỷ lệ trên một số tình huống chuyển giao; Phòng Khoa học Công nghệ và Viện Nghiên cứu giáo dục và Chuyển giao "
          "tri thức tổ chức khai thác. Thời gian: quy định ban hành trước ngày 31 tháng 5 năm 2027; khai thác từ năm 2027."),
    ("P", "**Sản phẩm:** quy định lợi ích ba lớp; mẫu thỏa thuận chia thưởng; danh mục tài sản sẵn sàng chuyển giao; phấn đấu "
          "ít nhất 01 hợp đồng chuyển giao hoặc cấp phép mỗi năm."),

    ("H2", "3.1.4. Tăng cường bảo vệ quyền sở hữu trí tuệ"),
    ("P", "**Vấn đề.** Chưa có cơ chế theo dõi xâm phạm, sổ theo dõi hiệu lực văn bằng; chưa có quy định về kiểm tra trùng lặp "
          "và dùng trí tuệ nhân tạo; việc bảo mật trước khi công bố chưa được ghi nhận (Mục 2.2.4)."),
    ("P", "**Việc cần làm:**"),
    ("P", "- **Bảo mật trước khi công bố.** Áp dụng Khâu 4 của quy trình: trước khi gửi bài, báo cáo hội thảo hay trình diễn, "
          "chủ nhiệm đề tài xin ý kiến Phòng Khoa học Công nghệ; người tham gia hợp tác, thử nghiệm ký cam kết bảo mật; hồ sơ "
          "chưa nộp đơn và bí mật kinh doanh chỉ người được giao xử lý mới truy cập được."),
    ("P", "- **Liêm chính học thuật.** Mở rộng quy định về xâm phạm quyền tác giả của Quyết định 217; bổ sung kiểm tra trùng "
          "lặp, công khai việc dùng trí tuệ nhân tạo theo Nghị định 134 và quy trình xử lý vi phạm, đáp ứng Thông tư 83."),
    ("P", "- **Duy trì hiệu lực văn bằng.** Lập lịch nộp phí duy trì, gia hạn trong danh mục số, cảnh báo trước 03 tháng; quyết "
          "định duy trì văn bằng dựa trên nhu cầu khai thác."),
    ("P", "- **Theo dõi và xử lý xâm phạm.** Giao Bộ phận Pháp chế làm đầu mối tiếp nhận thông tin xâm phạm; quy định quy trình "
          "xử lý tranh chấp nội bộ và xử lý quyền khi tác giả chuyển công tác."),
    ("P", "- **Quản lý nhãn hiệu hệ sinh thái.** Quy định việc quản lý và cấp phép sử dụng nhãn hiệu của Nhà trường cho các "
          "pháp nhân thành viên."),
    ("P", "**Thực hiện:** Bộ phận Pháp chế chủ trì; Phòng Khoa học Công nghệ quản lý danh mục số; bộ phận quản trị thương hiệu "
          "quản lý nhãn hiệu. Thời gian: đưa vào Quy chế sửa đổi trước ngày 31 tháng 5 năm 2027."),
    ("P", "**Sản phẩm:** quy định về bảo mật, liêm chính và xử lý xâm phạm trong Quy chế sửa đổi; mẫu cam kết bảo mật; lịch "
          "duy trì văn bằng."),

    ("H1", "3.2. Kế hoạch tổ chức thực hiện"),
    ("P", "**Thể chế hóa các giải pháp.** Các nội dung về quyền sở hữu, lợi ích, quy trình, bảo mật và liêm chính nêu tại Mục "
          "3.1 được đưa vào một Quy chế quản trị tài sản trí tuệ sửa đổi, xây dựng trên nền Quyết định 217 và hợp nhất các quy "
          "định về sở hữu trí tuệ của Quyết định 213. Bộ phận Pháp chế soạn thảo, trình trong quý IV năm 2026 và ban hành trước "
          "ngày 31 tháng 5 năm 2027, trước kỳ tự đánh giá đầu tiên theo Thông tư 83."),
    ("P", "**Phân công và thời hạn.** Bảng 3.1 tổng hợp việc chính, đơn vị thực hiện, thời gian và sản phẩm của từng giải pháp."),
    ("BANG", "Bảng 3.1. Phân công thực hiện các giải pháp",
     ["Giải pháp", "Việc chính", "Chủ trì; phối hợp", "Thời gian", "Sản phẩm"],
     [
         ["Thể chế hóa", "Sửa Quy chế quản trị tài sản trí tuệ trên nền Quyết định 217, hợp nhất quy định của Quyết định 213",
          "Bộ phận Pháp chế; Phòng Khoa học Công nghệ, Phòng Tài chính - Kế toán", "Trình quý IV/2026; ban hành trước "
          "31/5/2027", "Quy chế sửa đổi"],
         ["3.1.1. Sáng tạo và nhận diện", "Sửa biểu mẫu đề tài; ưu tiên đề tài có sản phẩm bảo hộ được; tập huấn giảng viên "
          "mới; học phần sở hữu trí tuệ; đề xuất sửa Quy chế chi tiêu nội bộ", "Phòng Khoa học Công nghệ; Phòng Đào tạo, "
          "Phòng Tài chính - Kế toán", "Từ quý IV/2026; học phần từ năm học 2027 - 2028", "Biểu mẫu sửa đổi; chương trình tập "
          "huấn; cẩm nang"],
         ["3.1.2. Sàng lọc và xác lập", "Quy trình 8 khâu; dòng dự toán phí; giao đầu mối, quy chế phối hợp; danh mục số; rà "
          "soát hồ sơ tồn đọng", "Phòng Khoa học Công nghệ; Bộ phận Pháp chế, Hội đồng nghiệm thu, Phòng Tài chính - Kế toán",
          "Thí điểm quý IV/2026 - quý II/2027; toàn trường từ quý III/2027", "Phiếu khai báo, phiếu rà soát; dự toán; danh "
          "mục số"],
         ["3.1.3. Khai thác", "Quy định lợi ích ba lớp; mẫu thỏa thuận đồng tác giả; khai thác qua tư vấn thuê ngoài; chuẩn bị "
          "thí điểm định giá", "Bộ phận Pháp chế; Phòng Khoa học Công nghệ, Viện Nghiên cứu giáo dục và Chuyển giao tri thức",
          "Quy định trước 31/5/2027; khai thác từ năm 2027", "Mẫu thỏa thuận; danh mục tài sản sẵn sàng chuyển giao"],
         ["3.1.4. Bảo vệ", "Quy định bảo mật, liêm chính, xử lý xâm phạm; lịch duy trì văn bằng; quản lý nhãn hiệu hệ sinh "
          "thái", "Bộ phận Pháp chế; Phòng Khoa học Công nghệ, bộ phận quản trị thương hiệu", "Trước 31/5/2027",
          "Mẫu cam kết bảo mật; lịch duy trì văn bằng"],
     ],
     "Nguồn: Nhóm nghiên cứu đề xuất.", [1500, 2600, 2000, 1500, 1531]),
    ("GP", "Giai đoạn đầu chủ yếu dùng nguồn lực sẵn có"),
    ("P", "**Lộ trình.** Các giải pháp được triển khai theo ba giai đoạn, gắn với thời điểm có hiệu lực của các văn bản mới "
          "(Hình 3.3)."),
    ("GK", "Hình 3.3.", {}),
    ("GKR", "Giai đoạn 1, từ quý IV năm 2026", "- Hợp tác mở rộng với doanh nghiệp",
     {"khi đạt điều kiện tại Giải pháp 5": "khi đạt điều kiện nêu tại Mục 3.1.3",
      "theo yêu cầu của Thông tư 83/2026/TT-BGDĐT khi văn bản có hiệu lực": "theo yêu cầu của Thông tư 83"}),
    ("P", "**Thí điểm.**"),
    ("GKR", "Đây là kế hoạch, chưa được triển khai.", "Sản phẩm đầu ra là báo cáo đánh giá thí điểm",
     {"mục tiêu tại Bảng 3.4": "mục tiêu tại Bảng 3.2"}),
    ("P", "**Theo dõi và đánh giá.**"),
    ("GPS", "Bộ chỉ số tại Bảng 3.4 phân tầng", {"Bảng 3.4": "Bảng 3.2", "(Hình 2.10)": "(Hình 2.8)"}),
    ("GK", "Bảng 3.4.", {"Bảng 3.4.": "Bảng 3.2."}),
    ("GP", "Phòng Khoa học Công nghệ tổng hợp và báo cáo các chỉ số hằng năm"),
    ("TK", "TIỂU KẾT CHƯƠNG 3"),
    ("P", "Chương 3 đề xuất bốn giải pháp tương ứng với bốn khâu: định hướng sáng tạo và nhận diện tài sản trí tuệ ngay khi "
          "duyệt đề tài; sàng lọc trước khi công bố, nộp đơn sớm không chờ nghiệm thu, bố trí kinh phí, đầu mối và danh mục số "
          "cho bước xác lập; sửa quy định về lợi ích ba lớp và tổ chức khai thác theo giai đoạn; bảo mật, liêm chính và duy trì "
          "văn bằng."),
    ("P", "Các giải pháp được thể chế hóa trong Quy chế quản trị tài sản trí tuệ sửa đổi, triển khai theo lộ trình ba giai đoạn "
          "đến năm 2030, thí điểm tại Viện Y - Dược và theo dõi bằng bộ chỉ số phân tầng. Phần lớn việc dùng nguồn lực sẵn có; "
          "các việc cần nhiều nguồn lực chỉ làm khi đủ điều kiện."),
]

KET_LUAN = [
    ("GP", "KẾT LUẬN VÀ KIẾN NGHỊ"),
    ("GP", "1. Kết luận"),
    ("GP", "Kết quả nghiên cứu đáp ứng ba mục tiêu cụ thể"),
    ("P", "**Về lý luận và pháp lý**, đề tài làm rõ nội dung quản lý quyền sở hữu trí tuệ trong trường đại học theo bốn khâu "
          "sáng tạo, xác lập, khai thác, bảo vệ, cùng các yêu cầu của khung pháp lý 2025 - 2026 đối với từng khâu: giao quyền "
          "tự động và quyền đăng ký của tổ chức chủ trì, mức thưởng tối thiểu 30% cho tác giả với kết quả dùng ngân sách nhà "
          "nước, yêu cầu ghi nhận đóng góp, thỏa thuận chia thưởng và quy chế về liêm chính học thuật."),
    ("GP", "Về thực trạng, Nhà trường đã hình thành nền tảng thể chế từ sớm"),
    ("P", "**Về giải pháp**, đề tài đề xuất bốn giải pháp theo bốn khâu, trọng tâm là quy trình 8 khâu với sàng lọc trước khi "
          "công bố, luồng nộp đơn sớm không chờ nghiệm thu và bước kiểm tra tại nghiệm thu. Các giải pháp có phân công, lộ "
          "trình đến năm 2030, kế hoạch thí điểm và bộ chỉ số theo dõi."),
    ("P", "**Về đóng góp**, đề tài cung cấp cho Nhà trường căn cứ cụ thể để sửa quy chế theo luật mới, chuẩn hóa quy trình, bố "
          "trí kinh phí, số hóa dữ liệu và theo dõi bằng bộ chỉ số. Kết quả có thể tham khảo cho các trường đại học tư thục có "
          "điều kiện tương tự."),
    ("GKR", "2. Kiến nghị với cơ quan quản lý nhà nước", "Thứ tư, tiếp tục hợp tác với doanh nghiệp", {}),
]

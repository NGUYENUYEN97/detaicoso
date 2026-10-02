# -*- coding: utf-8 -*-
"""Nội dung rút gọn: Mở đầu và Chương 1 (bản v2 của nhóm tác giả, rút gọn còn khoảng 60 - 70 trang).

Ký hiệu: ("H0", chữ) tiêu đề phần có ngắt trang; ("CH", chỉ số) giữ nguyên đoạn tên chương của bản v2;
("H1", chữ) mục cấp 1; ("H2", chữ) mục cấp 2; ("H3", chữ) tiểu mục a), b); ("P", chữ) đoạn thân, nhận **đậm**;
("K", đầu, cuối, {cũ: mới}) giữ nguyên các khối bảng, hình của bản v2, có thể thay chữ; ("TK", chữ) tiểu kết.
"""

MO_DAU = [
    ("H0", "MỞ ĐẦU"),
    ("H1", "1. Tính cấp thiết của đề tài"),
    ("P", "Tài sản trí tuệ là nguồn lực quan trọng đối với năng lực cạnh tranh, uy tín học thuật và khả năng tự chủ tài "
          "chính của trường đại học. Giai đoạn 2025 - 2026, pháp luật về lĩnh vực này thay đổi nhanh. Luật Khoa học, công "
          "nghệ và đổi mới sáng tạo năm 2025 và nghị định hướng dẫn giao cho nhà trường quyền tự quyết về thương mại hóa và "
          "đặt mức thưởng tối thiểu cho tác giả. Luật Sở hữu trí tuệ được sửa đổi, Luật Giáo dục đại học năm 2025 được ban "
          "hành, Chiến lược sở hữu trí tuệ đến năm 2030 được sửa đổi và Chuẩn cơ sở giáo dục đại học mới ra đời. Tất cả đều "
          "đặt yêu cầu mới đối với quản lý tài sản trí tuệ trong nhà trường."),
    ("P", "Trường Đại học Thành Đô đã sớm ban hành Quy chế hoạt động khoa học công nghệ năm 2021, gọi tắt là Quyết định "
          "213, và Quy chế quản trị tài sản trí tuệ năm 2024, gọi tắt là Quyết định 217. Năng lực công bố tăng "
          "nhanh trong giai đoạn 2021 - 2025, và 11 trên 38 đề tài cấp cơ sở đã tạo ra sản phẩm có tiềm năng tạo lập "
          "tài sản trí tuệ, nhất là ở khối ngành Y - Dược. Tuy vậy, việc chuyển kết quả nghiên cứu thành quyền sở hữu "
          "trí tuệ còn ở quy mô nhỏ, và các quy chế ban hành trước năm 2025 cần được cập nhật. Nghiên cứu "
          "giải pháp nâng cao hiệu quả quản lý quyền sở hữu trí tuệ vì vậy vừa là yêu cầu tuân thủ pháp luật, vừa là cơ "
          "hội để Nhà trường chủ động thích ứng với khung pháp lý mới."),
    ("H1", "2. Tổng quan tình hình nghiên cứu"),
    ("H2", "2.1. Tình hình nghiên cứu trong nước"),
    ("P", "Các công trình trong nước tập trung vào ba hướng. Hướng thứ nhất là lý luận và chu trình quản lý. Phạm (2019) "
          "xác định bốn nội dung quản lý sở hữu trí tuệ trong trường đại học: xây dựng quy chế, tổ chức bộ máy, vận hành "
          "chu trình xác lập, khai thác và nâng cao nhận thức. Lê và Nguyễn (2019) đề xuất mô hình quản trị năm bước từ "
          "lập kế hoạch đến đánh giá hiệu quả, đồng thời chỉ ra các rào cản về kinh phí, định giá và bộ phận chuyên "
          "trách."),
    ("P", "Hướng thứ hai là quy chế nội bộ và quyền lợi của tác giả. Nguyễn (2025) nhấn mạnh yêu cầu ứng dụng công nghệ "
          "thông tin trong quản trị tài sản trí tuệ. Võ (2025) phân tích quyền của người trực tiếp sáng tạo nhưng không "
          "giữ quyền tài sản, cho thấy nhiều quy chế chưa quy định rõ quyền lợi của tác giả. Lê và Hoàng (2021) đề xuất "
          "thành lập bộ phận chuyên trách và xây dựng cách chia doanh thu minh bạch giữa nhà trường và tác giả."),
    ("P", "Hướng thứ ba là kinh nghiệm quốc tế và thực trạng chung. Phạm và Nguyễn (2018) phân tích mô hình quản lý ba "
          "cấp của Đại học Thanh Hoa, Trung Quốc. Báo cáo của Cục Sở hữu trí tuệ và Bộ Khoa học và Công nghệ (2024) cho "
          "thấy số đơn sáng chế và hợp đồng chuyển giao của khối đại học còn thấp so với tiềm năng, chủ yếu do thiếu quy "
          "chế đồng bộ và cơ chế thương mại hóa chuyên nghiệp."),
    ("H2", "2.2. Tình hình nghiên cứu ngoài nước"),
    ("P", "Thứ nhất, về thể chế và mô hình quản lý, các nghiên cứu cho thấy trường đại học ngày càng vận hành như một chủ "
          "thể khởi nghiệp (Etzkowitz, 2003), và hiệu quả thương mại hóa phụ thuộc vào cách phân bổ quyền, lợi ích giữa "
          "nhà trường và nhà khoa học (Goldfarb & Henrekson, 2003; Siegel et al., 2007; Thursby & Kemp, 2002). "
          "Maresova và cộng sự (2019) khẳng định không có một mô hình chung cho mọi trường. Holgersson (2021) phân tích "
          "sự chuyển dịch ở châu Âu từ quyền của nhà sáng chế sang quyền sở hữu của trường đại học."),
    ("P", "Thứ hai, về đơn vị chuyển giao công nghệ, Perkmann và cộng sự (2013) nhấn mạnh cần cân bằng giữa bảo hộ và "
          "hợp tác mở. Bstieler và cộng sự (2015) cùng O’Dwyer và cộng sự (2023) cho thấy quy chế sở hữu trí tuệ minh "
          "bạch giúp xây dựng lòng tin trong hợp tác với doanh nghiệp. Öztürk (2026) xác định thiếu nhân sự chuyên môn "
          "là điểm nghẽn lớn nhất của đơn vị chuyển giao. Pandey và cộng sự (2025) cùng Bulsara và Vaghela (2025) đề "
          "xuất kết hợp đơn vị chuyển giao với vườn ươm, giảm thủ tục và linh hoạt trong định giá, chia doanh thu."),
    ("P", "Thứ ba, về quy trình và đo lường, các nghiên cứu nhấn mạnh việc chuẩn hóa tiếp nhận, sàng lọc bản khai báo "
          "sáng chế (Holgersson & Aaboen, 2019; Rocha et al., 2023), ứng dụng phần mềm quản lý vòng đời tài sản "
          "(Holgersson, 2021) và bộ chỉ số gồm đơn, hợp đồng chuyển giao, doanh thu và mức tuân thủ quy chế (Maresova "
          "et al., 2019; Pandey et al., 2025)."),
    ("H2", "2.3. Đánh giá tổng hợp và khoảng trống nghiên cứu"),
    ("P", "Các mô hình quốc tế phần lớn được xây dựng cho trường đại học nghiên cứu có quyền tự chủ cao và nguồn lực lớn, "
          "nên không thể áp dụng nguyên trạng cho trường đại học định hướng ứng dụng ở Việt Nam. Các nghiên cứu trong "
          "nước chủ yếu phân tích pháp luật chung hoặc mô tả thực trạng ở một số trường công lập; rất ít công trình dùng "
          "dữ liệu hành chính của một trường để theo dõi chuỗi từ đề tài đến đơn và văn bằng. Chưa có công trình đánh "
          "giá thực trạng quản lý quyền sở hữu trí tuệ tại một trường đại học tư thục như Trường Đại học Thành Đô trong "
          "đối chiếu với khung pháp lý 2025 - 2026. Đề tài được thực hiện để lấp khoảng trống này và cung cấp căn cứ "
          "cho việc hoàn thiện quy chế, tổ chức, quy trình và cơ chế khai thác của Nhà trường."),
    ("H1", "3. Mục tiêu và câu hỏi nghiên cứu"),
    ("P", "**Mục tiêu tổng quát:** đề xuất hệ thống giải pháp khả thi nhằm nâng cao hiệu quả quản lý quyền sở hữu trí tuệ "
          "tại Trường Đại học Thành Đô, đáp ứng khung pháp lý mới và yêu cầu của Chuẩn cơ sở giáo dục đại học."),
    ("P", "**Mục tiêu cụ thể:** thứ nhất, hệ thống hóa cơ sở lý luận và pháp lý, xây dựng khung phân tích và bộ tiêu "
          "chí đánh giá; thứ hai, đánh giá thực trạng quản lý quyền sở hữu trí tuệ tại Nhà trường giai đoạn 2021 - "
          "2025, chỉ ra kết quả, hạn chế và nguyên nhân; thứ ba, đề xuất giải pháp kèm lộ trình, kế hoạch thí điểm và "
          "bộ chỉ số theo dõi."),
    ("P", "**Câu hỏi nghiên cứu:** Khung pháp lý mới đặt ra yêu cầu gì đối với quản lý quyền sở hữu trí tuệ của trường "
          "đại học tư thục? Nguồn tài sản trí tuệ tiềm năng của Nhà trường lớn đến đâu và đã được chuyển thành quyền đến "
          "mức nào? Yếu tố nào có thể đang cản trở việc chuyển tiềm năng thành kết quả? Giải pháp nào khả thi trong "
          "giai đoạn 2026 - 2030?"),
    ("H1", "4. Đối tượng và phạm vi nghiên cứu"),
    ("P", "**Đối tượng:** hoạt động quản lý quyền sở hữu trí tuệ tại Trường Đại học Thành Đô, gồm thể chế, tổ chức, "
          "nguồn lực và chu trình tạo lập, xác lập, bảo vệ, khai thác tài sản trí tuệ."),
    ("P", "**Phạm vi:** tài sản trí tuệ phát sinh từ hoạt động khoa học công nghệ và đào tạo của giảng viên, người lao "
          "động và người học, không gồm tài sản của các pháp nhân khác trong hệ sinh thái; số liệu thực trạng giai đoạn "
          "2021 - 2025 và danh sách nhân sự năm 2026; văn bản pháp luật hiện hành đến tháng 9 năm 2026; giải pháp hướng "
          "tới giai đoạn 2026 - 2030."),
    ("H1", "5. Cách tiếp cận và phương pháp nghiên cứu"),
    ("P", "Đề tài tiếp cận theo chu trình quản lý tạo lập, xác lập, khai thác và bảo vệ, kết hợp bộ tiêu chí đánh giá "
          "theo chuỗi đầu vào, quá trình, đầu ra và kết quả xây dựng tại Chương 1."),
    ("P", "Phương pháp phân tích văn bản được dùng để đối chiếu quy chế, kế hoạch của Nhà trường với pháp luật hiện "
          "hành. Dữ liệu định lượng gồm 582 bản ghi sản phẩm khoa học giai đoạn 2021 - 2025 của Phòng Khoa "
          "học Công nghệ, hồ sơ 38 đề tài cấp cơ sở, danh mục theo dõi đơn và văn bằng, danh sách 252 nhân sự năm "
          "2026; dữ liệu cá nhân chỉ được dùng ở dạng tổng hợp."),
    ("P", "Dữ liệu được làm sạch, chuẩn hóa tên đơn vị và ghi rõ năm theo mã số, năm phê duyệt, năm nghiệm thu, năm nộp "
          "đơn hoặc năm cấp văn bằng. Hồ sơ tài sản trí tuệ được xếp theo bốn trạng thái: đã nộp đơn, đã chấp nhận đơn "
          "hợp lệ, đã cấp văn bằng hoặc giấy chứng nhận, chưa xác minh. Tên tác giả bài báo được đối chiếu với danh sách "
          "nhân sự năm 2026; kết quả chỉ mang tính thăm dò. Mức độ tập trung công bố được đo bằng hệ số Gini; mức hoàn "
          "thành kế hoạch được tính theo tỷ lệ so với chỉ tiêu. Kết quả được tổng hợp bằng ma trận điểm mạnh, điểm yếu, "
          "thời cơ và thách thức làm căn cứ đề xuất giải pháp."),
    ("P", "**Quy ước tên gọi văn bản.** Mỗi văn bản được gọi bằng một tên ngắn thống nhất trong toàn báo cáo, liệt kê tại "
          "Bảng 1.1; số hiệu và ngày ban hành đầy đủ ghi tại Tài liệu tham khảo. Điều khoản cụ thể chỉ nêu tại Bảng 1.1 và "
          "khi đối chiếu quy chế tại Mục 2.2.1; các phần khác chỉ nêu nội dung quy định."),
    ("H1", "6. Ý nghĩa khoa học và thực tiễn"),
    ("P", "**Về khoa học:** đề tài đưa ra khung phân tích và bộ tiêu chí đánh giá có thể tính từ dữ liệu hành chính của "
          "trường đại học, và nêu giả thuyết về vai trò của bước sàng lọc trước công bố trong chuỗi từ đề tài đến văn "
          "bằng, kèm cách đo để kiểm chứng."),
    ("P", "**Về thực tiễn:** đề tài cung cấp căn cứ để Nhà trường hoàn thiện quy chế theo luật mới, chuẩn hóa quy trình, "
          "bố trí nguồn lực và theo dõi bằng bộ chỉ số; kết quả có thể tham khảo cho các trường đại học tư thục có điều "
          "kiện tương tự."),
    ("H1", "7. Kết cấu của báo cáo"),
    ("P", "Ngoài Mở đầu, Kết luận và kiến nghị, Tài liệu tham khảo, báo cáo gồm ba chương: Chương 1. Cơ sở lý luận về "
          "quản lý quyền sở hữu trí tuệ trong cơ sở giáo dục đại học; Chương 2. Thực trạng quản lý quyền sở hữu trí tuệ "
          "tại Trường Đại học Thành Đô; Chương 3. Giải pháp nâng cao hiệu quả quản lý quyền sở hữu trí tuệ tại Trường "
          "Đại học Thành Đô đáp ứng khung pháp lý mới."),
]

CHUONG_1 = [
    ("CH", 159),
    ("H1", "1.1. Khái niệm, đặc điểm và vai trò của quản lý quyền sở hữu trí tuệ trong trường đại học"),
    ("H2", "1.1.1. Một số khái niệm"),
    ("P", "**Tài sản trí tuệ.** Tổ chức Sở hữu trí tuệ thế giới (2020) coi tài sản trí tuệ là các sản phẩm sáng tạo của "
          "trí tuệ con người, như tác phẩm, sáng chế, kiểu dáng, nhãn hiệu và chương trình máy tính. Fisher (2001) và "
          "Guan (2014) nhấn mạnh hai mặt của tài sản trí tuệ: đó là kết quả của lao động trí óc, và có thể mang lại lợi "
          "ích kinh tế hoặc pháp lý cho chủ sở hữu. Trong trường đại học, tài sản trí tuệ gồm các kết quả học thuật như "
          "công trình nghiên cứu, bài báo, giáo trình, sáng chế và phần mềm (Cục Sở hữu trí tuệ, n.d.)."),
    ("P", "**Quyền sở hữu trí tuệ.** Theo Luật Sở hữu trí tuệ, đây là quyền của tổ chức, cá nhân đối với tài sản trí "
          "tuệ, gồm ba nhóm: quyền tác giả và quyền liên quan; quyền sở hữu công nghiệp đối với sáng chế, kiểu dáng công "
          "nghiệp, thiết kế bố trí, nhãn hiệu, tên thương mại, chỉ dẫn địa lý, bí mật kinh doanh; quyền đối với giống cây "
          "trồng. Các điều ước quốc tế cũng xác định quyền sở hữu trí tuệ bằng cách liệt kê các nhóm quyền này."),
    ("P", "Để thuận tiện cho quản lý, tài sản trí tuệ trong trường có thể chia thành ba nhóm theo nguồn hình thành và đơn "
          "vị phụ trách. Đây là cách phân nhóm quản lý, không thay thế phân loại pháp lý:"),
    ("P", "- Nhóm quyền tác giả: giáo trình, bài giảng, bài báo, báo cáo nghiên cứu, luận văn, phần mềm và cơ sở dữ liệu "
          "học liệu."),
    ("P", "- Nhóm tài sản công nghệ, thuộc quyền sở hữu công nghiệp: sáng chế, giải pháp hữu ích, kiểu dáng công nghiệp, "
          "thiết kế bố trí và bí mật kinh doanh phát sinh từ đề tài, dự án."),
    ("P", "- Nhóm tài sản nhận diện, cũng thuộc quyền sở hữu công nghiệp: nhãn hiệu, tên thương mại, biểu trưng của trường. "
          "Nhóm này được tách riêng vì do bộ phận thương hiệu quản lý và có cách khai thác khác."),
    ("P", "**Quản lý quyền sở hữu trí tuệ trong cơ sở giáo dục đại học** là các hoạt động có chủ đích của nhà trường nhằm "
          "ban hành quy chế, tổ chức bộ máy và vận hành quy trình để tài sản trí tuệ được tạo lập, xác lập quyền, khai "
          "thác và bảo vệ (Bradley et al., 2013; Goldfarb & Henrekson, 2003). Khác với doanh nghiệp, vốn ưu tiên lợi "
          "nhuận (Teece, 2018), trường đại học phải cân bằng hai mục tiêu: phổ biến tri thức vì lợi ích chung và bảo hộ "
          "để thu hồi chi phí, tái đầu tư, tạo thu nhập cho nhà khoa học (Etzkowitz, 2003; Nguyễn, 2025; Perkmann et "
          "al., 2013). Vì vậy, nhà trường không thể áp dụng nguyên mô hình quản lý của doanh nghiệp."),
    ("H2", "1.1.2. Đặc điểm của tài sản trí tuệ trong trường đại học"),
    ("P", "Tài sản trí tuệ trong trường đại học có sáu đặc điểm chính, mỗi đặc điểm đặt ra một yêu cầu quản lý:"),
    ("P", "- **Vô hình và khó nhận diện.** Nhiều tri thức có giá trị, như phương pháp, thuật toán, ý tưởng ban đầu, nằm "
          "trong suy nghĩ của giảng viên. Nhà trường chỉ quản lý được khi kết quả được khai báo."),
    ("P", "- **Có tính chất hàng hóa công cộng.** Một người dùng tri thức không làm giảm khả năng dùng của người khác "
          "(Fisher, 2001; Guan, 2014). Nhà trường phải chọn kết quả nào cần bảo hộ độc quyền, kết quả nào nên công bố "
          "rộng rãi."),
    ("P", "- **Hình thành từ nhiều nguồn.** Tài sản có thể đến từ nhiệm vụ dùng ngân sách nhà nước, đề tài cấp cơ sở, sản "
          "phẩm của người học hay hợp tác với doanh nghiệp (Bradley et al., 2013). Nguồn hình thành quyết định ai là chủ "
          "sở hữu và lợi ích được chia thế nào."),
    ("P", "- **Nhiều chủ thể cùng sáng tạo.** Giảng viên, người học, chuyên gia doanh nghiệp cùng tham gia, nên cần xác "
          "định mức đóng góp và quyền của từng người ngay từ đầu (Perkmann et al., 2013)."),
    ("P", "- **Áp lực giữa công bố và bảo hộ.** Giảng viên cần công bố sớm để được ghi nhận, nhưng sáng chế cần giữ bí mật "
          "đến khi nộp đơn. Nếu quy chế không rõ về thưởng, thù lao và cách chia nguồn thu, nhà khoa học dễ tự khai thác "
          "kết quả bên ngoài (Shane, 2004; Võ, 2025)."),
    ("P", "- **Căn cứ xác lập và thời hạn bảo hộ khác nhau.** Quyền tác giả phát sinh khi tác phẩm được định hình, không "
          "cần đăng ký. Sáng chế, kiểu dáng, nhãn hiệu chỉ được bảo hộ khi được cấp văn bằng. Bí mật kinh doanh được bảo hộ "
          "khi được giữ bí mật. Bằng sáng chế có hiệu lực 20 năm, bằng giải pháp hữu ích 10 năm; kiểu dáng 5 năm, nhãn "
          "hiệu 10 năm và được gia hạn. Văn bằng do Việt Nam cấp chỉ có hiệu lực tại Việt Nam. Nhà trường vì vậy phải chọn "
          "đúng hình thức bảo hộ cho từng loại tài sản."),
    ("H2", "1.1.3. Vai trò của quản lý quyền sở hữu trí tuệ đối với nhà trường"),
    ("P", "Quản lý quyền sở hữu trí tuệ tốt mang lại cho trường đại học ba lợi ích thiết thực:"),
    ("P", "- **Nâng chất lượng đào tạo và nghiên cứu.** Kết quả nghiên cứu được quản lý bài bản sẽ nhanh chóng được đưa "
          "vào bài giảng. Quy chế rõ ràng về quyền đối với đồ án, khóa luận khuyến khích người học giải quyết bài toán "
          "thực tế. Cơ chế minh bạch giúp giữ chân nhà khoa học và hình thành nhóm nghiên cứu mạnh."),
    ("P", "- **Tạo nguồn thu ngoài học phí.** Sáng chế, phần mềm, giáo trình khi được chuyển nhượng hoặc chuyển quyền sử "
          "dụng tạo ra khoản thu hợp pháp. Năng lực quản lý tốt cũng giúp thu hút hợp đồng nghiên cứu với doanh nghiệp "
          "(O’Dwyer et al., 2023). Sau khi trả thù lao, thưởng cho tác giả, phần còn lại có thể đưa vào quỹ phát triển "
          "khoa học và công nghệ để tái đầu tư."),
    ("P", "- **Đáp ứng chuẩn và kiểm định chất lượng.** Chuẩn cơ sở giáo dục đại học tính văn bằng vào sản phẩm khoa học "
          "của trường và tính nguồn thu từ chuyển giao vào thu khoa học công nghệ. Chuẩn mới theo Thông tư 83 tính văn bằng "
          "cao hơn trước, như tổng hợp tại Bảng 1.1."),
    ("H1", "1.2. Nội dung quản lý quyền sở hữu trí tuệ trong trường đại học"),
    ("H2", "1.2.1. Chu trình quản lý bốn khâu"),
    ("P", "Chiến lược sở hữu trí tuệ đến năm 2030 định hướng phát triển đồng bộ các khâu sáng tạo, xác lập, khai thác và "
          "bảo vệ quyền sở hữu trí tuệ. Trong trường đại học, bốn khâu này tạo thành chu trình "
          "tại Hình 1.1."),
    ("K", 267, 269, {}),
    ("P", "- **Tạo lập.** Giảng viên, người học tạo ra kết quả nghiên cứu, giáo trình, phần mềm. Nhà trường định hướng qua "
          "nhiệm vụ khoa học công nghệ, hợp tác doanh nghiệp và Không gian sáng tạo mở thử nghiệm. Việc quan trọng ở khâu "
          "này là xác định đối tượng quyền cần đạt ngay khi duyệt thuyết minh."),
    ("P", "- **Xác lập quyền.** Đây là việc làm phát sinh hoặc xác nhận quyền. Sáng chế, "
          "giải pháp hữu ích, kiểu dáng, nhãn hiệu phải nộp đơn và được cấp văn bằng. Bí mật kinh doanh được bảo vệ bằng "
          "biện pháp bảo mật, không nộp đơn. Quyền tác giả phát sinh tự động; giấy chứng nhận đăng ký là chứng cứ khi "
          "khai thác hoặc tranh chấp. Với sáng chế, đơn nên được nộp trước khi công bố để giữ tính mới."),
    ("P", "- **Khai thác.** Khai thác phi thương mại là đưa giáo trình, kết quả nghiên cứu vào đào tạo và quản trị của "
          "trường. Khai thác thương mại gồm bốn hình thức: chuyển nhượng quyền; chuyển quyền sử dụng, thường gọi là cấp "
          "phép; góp vốn bằng quyền sở hữu trí tuệ để thành lập doanh nghiệp; trực tiếp sản xuất, cung ứng dịch vụ. Giá và cách thanh toán do các bên thỏa "
          "thuận."),
    ("P", "- **Bảo vệ.** Gồm bảo mật thông tin nghiên cứu, kiểm tra trùng lặp, theo dõi sao chép trái phép, tự bảo vệ "
          "quyền và phối hợp xử lý xâm phạm. Bảo vệ không phải khâu cuối mà diễn ra "
          "suốt vòng đời tài sản, từ khi có ý tưởng đến khi khai thác."),
    ("P", "Sau khi trả thù lao, thưởng cho tác giả, khoản thu từ khai thác được đưa trở lại quỹ phát triển khoa học và "
          "công nghệ, tạo vòng tái đầu tư cho chu kỳ nghiên cứu tiếp theo."),
    ("H2", "1.2.2. Chức năng quản lý của nhà trường"),
    ("P", "Để vận hành chu trình, nhà trường thực hiện bốn chức năng: hoạch định, gồm xác định lĩnh vực có khả năng tạo tài "
          "sản trí tuệ, ban hành và cập nhật quy chế; tổ chức, gồm giao đầu mối và phân định trách nhiệm giữa phòng chức "
          "năng, viện, khoa và tác giả; thực hiện, gồm tiếp nhận khai báo, tra cứu, hỗ trợ nộp đơn, đàm phán hợp đồng, "
          "chi trả thưởng, thù lao, nhuận bút; kiểm tra, gồm theo dõi hồ sơ, giám sát liêm chính và đánh giá hiệu quả để "
          "điều chỉnh chính sách."),
    ("H1", "1.3. Khung pháp lý về quản lý quyền sở hữu trí tuệ trong cơ sở giáo dục đại học"),
    ("H2", "1.3.1. Các văn bản pháp luật chính"),
    ("P", "Bảng 1.1 tổng hợp các văn bản trực tiếp chi phối hoạt động sở hữu trí tuệ của nhà trường, tên gọi tắt dùng "
          "trong báo cáo, quy định chính và yêu cầu đặt ra cho nhà trường."),
    ("BANG", "Bảng 1.1. Các văn bản pháp luật chính về quản lý quyền sở hữu trí tuệ trong cơ sở giáo dục đại học",
     ["Văn bản và tên gọi tắt", "Quy định chính liên quan đến nhà trường", "Yêu cầu đối với nhà trường"],
     [
         [["**Luật Sở hữu trí tuệ**",
           "Luật số 50/2005/QH11, sửa đổi bởi các Luật số 36/2009/QH12, 42/2019/QH14, 07/2022/QH15, 93/2025/QH15 và "
           "131/2025/QH15; hợp nhất tại Văn bản hợp nhất số 67/VBHN-VPQH"],
          ["Điều 39: nhà trường là chủ sở hữu quyền tài sản đối với tác phẩm giao cho giảng viên biên soạn, trừ khi có "
           "thỏa thuận khác.",
           "Điều 60: sáng chế mất tính mới nếu bị bộc lộ trước khi nộp đơn; chỉ một số trường hợp bộc lộ được giữ tính mới "
           "khi đơn được nộp trong 12 tháng.",
           "Điều 86: tổ chức chủ trì nhiệm vụ dùng ngân sách nhà nước có quyền đăng ký sáng chế, kiểu dáng, thiết kế bố "
           "trí là kết quả của nhiệm vụ.",
           "Điều 135: thù lao cho tác giả sáng chế, kiểu dáng, thiết kế bố trí theo thỏa thuận; nếu không có thỏa thuận "
           "là 10% lợi nhuận trước thuế khi tự sử dụng hoặc 15% số tiền nhận được khi chuyển giao."],
          ["Nộp đơn trước khi công bố.", "Quy định thù lao trong quy chế hoặc hợp đồng.",
           "Đăng ký kết quả của đề tài cấp quốc gia."]],
         [["**Luật KH,CN&ĐMST**",
           "Luật Khoa học, công nghệ và đổi mới sáng tạo số 93/2025/QH15, hiệu lực từ 01/10/2025"],
          ["Điều 25: tổ chức chủ trì tự động được giao quyền sở hữu phần kết quả tương ứng với kinh phí ngân sách nhà "
           "nước, không phải bồi hoàn.",
           "Điều 27: tổ chức tự quyết hình thức, giá và cách chia lợi nhuận khi thương mại hóa.",
           "Điều 28: với phần dùng ngân sách, thưởng tác giả tối thiểu 30% lợi nhuận sau thuế; phần không dùng ngân sách "
           "do chủ sở hữu tự quyết.",
           "Điều 66: quỹ phát triển khoa học và công nghệ của tổ chức được chi cho đăng ký, bảo hộ, khai thác.",
           "Điều 73: nhiệm vụ phê duyệt trước 01/10/2025 theo quy định cũ, trừ lợi nhuận từ văn bằng đã cấp của nhiệm "
           "vụ giao từ 01/01/2023."],
          ["Xác định lợi ích theo nguồn kinh phí và thời điểm giao nhiệm vụ.",
           "Không để mức trần làm thưởng thấp hơn 30%.", "Lập dự toán chi cho bảo hộ."]],
         [["**Nghị định 267**",
           "Nghị định số 267/2025/NĐ-CP hướng dẫn Luật KH,CN&ĐMST, hiệu lực từ 14/10/2025"],
          ["Điều 17: hồ sơ đánh giá cuối kỳ có văn bản xác định mức đóng góp của các thành viên.",
           "Điều 32: tổ chức chủ trì ngoài công lập được giao quyền sở hữu tự động và phải theo dõi riêng kết quả.",
           "Điều 34: chia lợi nhuận công khai; tổ chức trung gian hưởng tối thiểu 10% nếu không có thỏa thuận khác; "
           "đồng tác giả chia thưởng theo thỏa thuận."],
          ["Ghi nhận mức đóng góp khi nghiệm thu.", "Lập danh mục riêng kết quả từ ngân sách.",
           "Ban hành mẫu thỏa thuận chia thưởng."]],
         [["**Luật Giáo dục đại học**", "Luật số 125/2025/QH15, hiệu lực từ 01/01/2026"],
          ["Điều 27: đăng ký, bảo hộ, khai thác tài sản trí tuệ là nội dung của hoạt động khoa học công nghệ.",
           "Điều 28: được định giá, góp vốn, thành lập doanh nghiệp quản lý tài sản trí tuệ; phải lập quỹ phát triển "
           "khoa học và công nghệ và công khai kết quả hằng năm trên Nền tảng số quốc gia."],
          ["Chuẩn bị phương án khai thác.", "Có dữ liệu tài sản trí tuệ đủ để công khai."]],
         [["**Thông tư 83**",
           "Thông tư số 83/2026/TT-BGDĐT quy định Chuẩn cơ sở giáo dục đại học, hiệu lực từ 15/11/2026, thay Thông tư "
           "số 01/2024/TT-BGDĐT"],
          ["Yêu cầu quy chế quản trị về sở hữu trí tuệ, liêm chính khoa học, liêm chính học thuật.",
           "Tính văn bằng vào sản phẩm khoa học quy đổi: bằng giải pháp hữu ích tính 3, bằng sáng chế tính 5; tách riêng "
           "nguồn thu từ thương mại hóa.",
           "Dữ liệu thống nhất trên HEMIS; công bố kết quả tự đánh giá trước ngày 31/5 hằng năm."],
          ["Ban hành quy chế sửa đổi trước 31/5/2027.", "Thống kê văn bằng đúng định nghĩa của Chuẩn."]],
         [["**Kết luận 51** và **Quyết định 1624**",
           "Kết luận số 51-KL/TW năm 2026 của Bộ Chính trị; Quyết định số 1624/QĐ-TTg năm 2026 sửa đổi Chiến lược sở hữu "
           "trí tuệ đến năm 2030 ban hành kèm Quyết định số 1068/QĐ-TTg"],
          ["Kết luận 51: chuyển từ quản lý hành chính sang kiến tạo hệ sinh thái sở hữu trí tuệ.",
           "Kế thừa từ Quyết định 1068: dùng chỉ số sở hữu trí tuệ để đánh giá trường; trường khối kỹ thuật, công nghệ "
           "đăng ký bảo hộ đồng thời với công bố.",
           "Điểm mới: thí điểm định giá ít nhất 100 quyền; trung tâm tư vấn định giá, khai thác trong trường; hướng tới "
           "đưa sở hữu trí tuệ thành nội dung học bắt buộc."],
          ["Sàng lọc trước khi công bố.", "Chuẩn bị tham gia thí điểm định giá.", "Đưa sở hữu trí tuệ vào đào tạo."]],
         [["**Văn bản khác**",
           "Luật Chuyển giao công nghệ số 07/2017/QH14; Nghị định số 134/2026/NĐ-CP, gọi tắt là Nghị định 134; Chỉ thị số "
           "02/CT-TTg năm 2026, gọi tắt là Chỉ thị 02"],
          ["Luật Chuyển giao công nghệ: chuyển giao, định giá, góp vốn bằng kết quả nghiên cứu.",
           "Nghị định 134: tác phẩm có dùng trí tuệ nhân tạo chỉ được bảo hộ quyền tác giả khi con người đóng góp sáng tạo "
           "trực tiếp.",
           "Chỉ thị 02: tăng cường thực thi quyền, xây dựng cơ sở dữ liệu quốc gia về thực thi."],
          ["Quy định việc dùng trí tuệ nhân tạo trong quy chế liêm chính."]],
     ],
     "Nguồn: Nhóm nghiên cứu tổng hợp từ các văn bản nêu trong bảng.", [2300, 4500, 2268]),
    ("P", "Bảng 1.1 cho thấy hai thay đổi lớn. Thứ nhất, quyền quyết định chuyển về nhà trường: tổ chức chủ trì tự động có "
          "quyền sở hữu phần kết quả dùng ngân sách, có quyền đăng ký và tự quyết cách thương mại hóa. Thứ hai, đi kèm "
          "quyền là nghĩa vụ: thưởng tối thiểu cho tác giả, ghi nhận mức đóng góp, theo dõi riêng kết quả từ ngân sách, "
          "công khai kết quả và đáp ứng Chuẩn cơ sở giáo dục đại học."),
    ("H2", "1.3.2. Hệ quả đối với quy chế nội bộ"),
    ("P", "Để không nhầm lẫn các khoản lợi ích, báo cáo dùng bốn thuật ngữ. **Thù lao** là khoản chủ sở hữu trả cho tác "
          "giả sáng chế, kiểu dáng, thiết kế bố trí theo Luật Sở hữu trí tuệ. **Thưởng** là khoản trích từ lợi nhuận "
          "thương mại hóa theo Luật KH,CN&ĐMST. **Nhuận bút** là khoản trả cho tác giả sách, giáo trình. **Phân chia "
          "nguồn thu** là việc chủ sở hữu phân bổ khoản thu cho trường, đơn vị có tác giả và quỹ. Bốn khoản này khác nhau "
          "về căn cứ, người hưởng và cơ sở tính, nên tỷ lệ của khoản này không so sánh trực tiếp được với khoản khác."),
    ("P", "Từ đó, có ba điểm áp dụng. Một là, muốn biết tác giả được hưởng gì, phải xác định cùng lúc nguồn kinh phí, "
          "thời điểm giao nhiệm vụ và loại đối tượng. Hai là, thù lao và thưởng là hai khoản riêng, không thay thế nhau. "
          "Ba là, quy chế nội bộ có thể quy định mức có lợi hơn cho tác giả, nhưng không được thấp hơn mức tối thiểu của "
          "luật. Pháp luật chưa hướng dẫn cách tính lợi nhuận của một kết quả hình thành từ nhiều nguồn kinh phí, nên nhà "
          "trường cần tự quy định trong quy chế."),
    ("P", "Quy chế sở hữu trí tuệ của nhà trường vì vậy cần có bốn nhóm nội dung: xác định quyền sở hữu theo từng nguồn "
          "kinh phí, gồm nhiệm vụ trường giao, đề tài dùng cơ sở vật chất của trường, sản phẩm của người học và hợp đồng "
          "với doanh nghiệp; quy trình khai báo, đánh giá và quyết định nộp đơn trước khi công bố; cơ chế chia lợi ích "
          "theo nguồn hình thành và loại đối tượng; đầu mối, trách nhiệm phối hợp và kinh phí cho đăng ký, khai thác. Quy "
          "chế hiện hành của Trường Đại học Thành Đô được đối chiếu với các yêu cầu này tại Mục 2.2.1."),
    ("H1", "1.4. Tiêu chí đánh giá, yếu tố ảnh hưởng và khung phân tích"),
    ("H2", "1.4.1. Tiêu chí đánh giá hiệu quả"),
    ("P", "Hiệu quả quản lý được đánh giá theo chuỗi bốn nhóm: đầu vào, quá trình, đầu ra và kết quả. Mười sáu tiêu chí được "
          "trình bày tại Bảng 1.2; đây cũng là căn cứ để đánh giá khả năng đo lường của dữ liệu tại Chương 2."),
    ("K", 272, 274, {"Bảng 1.1.": "Bảng 1.2."}),
    ("P", "Cách đo theo chuỗi giúp phân biệt nhà trường đã đầu tư, tổ chức bao nhiêu với việc đã tạo ra bao nhiêu quyền và "
          "mang lại giá trị gì. Nếu chỉ đo đầu vào và đầu ra, nhà trường chưa biết hoạt động sở hữu trí tuệ có hiệu quả "
          "thực sự hay không."),
    ("H2", "1.4.2. Các yếu tố ảnh hưởng"),
    ("P", "Hiệu quả quản lý chịu tác động của sáu nhóm yếu tố: thể chế, gồm pháp luật và quy chế nội bộ; tổ chức bộ máy và "
          "đầu mối; nguồn lực tài chính cho nộp đơn, duy trì văn bằng; dữ liệu theo dõi vòng đời tài sản; năng lực của cán "
          "bộ quản lý và giảng viên; động lực và văn hóa, gồm cơ chế chia lợi ích, khen thưởng và ý thức tôn trọng quyền "
          "sở hữu trí tuệ. Khi một yếu tố yếu, ví dụ chưa có dự toán cho nộp đơn, kết quả nghiên cứu dễ dừng ở bài báo thay "
          "vì được đăng ký bảo hộ."),
    ("H2", "1.4.3. Khung phân tích của đề tài"),
    ("P", "Kết hợp chu trình quản lý, khung pháp lý mới, bộ tiêu chí và các yếu tố ảnh hưởng, đề tài xây dựng khung phân "
          "tích tại Hình 1.2. Khung này nối lý luận ở Chương 1 với đánh giá thực trạng ở Chương 2 và giải pháp ở Chương 3."),
    ("K", 317, 319, {"Hình 1.3.": "Hình 1.2."}),
    ("H1", "1.5. Kinh nghiệm và bài học cho Trường Đại học Thành Đô"),
    ("P", "**Kinh nghiệm quốc tế.** Đạo luật Bayh-Dole năm 1980 của Hoa Kỳ giao quyền sở hữu kết quả nghiên cứu dùng ngân "
          "sách cho trường đại học, thúc đẩy việc thành lập văn phòng chuyển giao công nghệ và doanh nghiệp khởi nguồn "
          "(Shane, 2004). Văn phòng chuyển giao không chỉ làm thủ tục mà còn tiếp thị công nghệ, định giá và kết nối nhà "
          "đầu tư (Siegel et al., 2007; Thursby & Kemp, 2002). Các trường thành công thường dành cho nhà sáng chế một tỷ lệ "
          "đáng kể trong doanh thu cấp phép (Siegel et al., 2007). Luật KH,CN&ĐMST của Việt Nam cũng đi theo hướng "
          "giao quyền tự động cho tổ chức chủ trì."),
    ("P", "**Kinh nghiệm trong nước.** Một số trường đại học tự chủ như Đại học Quốc gia Thành phố Hồ Chí Minh, Đại học "
          "Bách khoa Hà Nội, Trường Đại học Tôn Đức Thắng đã ban hành quy chế sở hữu trí tuệ riêng, có quy định khai báo "
          "trước khi công bố; lập quỹ chi trả phí nộp đơn sáng chế cho giảng viên; và đưa chỉ số sở hữu trí tuệ vào đánh "
          "giá giảng viên, thi đua của khoa."),
    ("P", "**Bài học cho Nhà trường.** Là trường tư thục định hướng ứng dụng, Trường Đại học Thành Đô không nên sao chép mô "
          "hình của đại học nghiên cứu lớn. Bốn bài học phù hợp là:"),
    ("P", "- Tập trung nguồn lực vào nhóm sản phẩm có khả năng bảo hộ cao nhất, nhất là lĩnh vực dược của Viện Y - Dược, "
          "cùng nhóm giáo trình, học liệu."),
    ("P", "- Phân luồng xử lý: thủ tục đơn giản cho quyền tác giả; tập trung chuyên môn cho sáng chế, giải pháp hữu ích."),
    ("P", "- Thuê tổ chức đại diện sở hữu công nghiệp khi chưa có nhân sự chuyên trách."),
    ("P", "- Dùng Không gian sáng tạo mở thử nghiệm để người học, giảng viên và doanh nghiệp cùng thử nghiệm, hoàn thiện "
          "sản phẩm trước khi quyết định nộp đơn hoặc chuyển giao. Mọi người tham gia phải ký cam kết bảo mật, hoặc nhà "
          "trường nộp đơn trước khi đưa sản phẩm ra thử nghiệm rộng. Thời hạn 12 tháng giữ tính mới khi kết quả đã lỡ "
          "bộc lộ chỉ là phương án dự phòng, không thay cho nguyên tắc nộp đơn trước."),
    ("TK", "TIỂU KẾT CHƯƠNG 1"),
    ("P", "Chương 1 đã làm rõ bốn vấn đề làm căn cứ cho đánh giá thực trạng và đề xuất giải pháp."),
    ("P", "Thứ nhất, quản lý quyền sở hữu trí tuệ trong trường đại học là quản lý chu trình tạo lập, xác lập, khai thác và "
          "bảo vệ, phải cân bằng giữa phổ biến tri thức và bảo hộ để tạo giá trị. Tài sản trí tuệ trong trường có nhiều "
          "nguồn, nhiều chủ thể và chịu áp lực giữa công bố và bảo hộ."),
    ("P", "Thứ hai, khung pháp lý 2025 - 2026 chuyển nhiều quyền quyết định về nhà trường: tổ chức chủ trì được giao quyền "
          "tự động và có quyền đăng ký; thưởng cho tác giả tối thiểu 30% lợi nhuận sau thuế với phần kết quả từ ngân sách, "
          "tách biệt với thù lao; mức đóng góp và thỏa thuận chia thưởng giữa đồng tác giả phải được ghi nhận. Vì vậy, quy chế nội bộ phải xác định lợi ích theo nguồn kinh phí, thời điểm giao "
          "nhiệm vụ và loại đối tượng."),
    ("P", "Thứ ba, bộ 16 tiêu chí theo chuỗi đầu vào, quá trình, đầu ra, kết quả và sáu nhóm yếu tố ảnh hưởng là công cụ "
          "đánh giá thực trạng tại Chương 2."),
    ("P", "Thứ tư, kinh nghiệm trong và ngoài nước cho thấy Nhà trường nên tập trung vào nhóm sản phẩm dược, phân luồng xử "
          "lý, thuê ngoài khi cần và dùng Không gian sáng tạo mở thử nghiệm với nguyên tắc bảo mật."),
]

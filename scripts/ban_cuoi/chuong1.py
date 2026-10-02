# -*- coding: utf-8 -*-
"""Chương 1 bản cuối.

Đầu vào: Chuong_1_Co_so_ly_luan_chuan.md (giữ nguyên văn phong của nhóm nghiên cứu).
Các chỉnh sửa trong SUA được đối chiếu với thư mục VBPL:
  - 67/VBHN-VPQH ngày 23/3/2026 (Luật Sở hữu trí tuệ hợp nhất): Điều 19, 20, 22, 25, 28, 39, 60, 86, 86a, 93,
    133a, 135 và các chú thích về Luật 93/2025/QH15, Luật 131/2025/QH15;
  - Luật Giáo dục đại học 125/2025/QH15 (Điều 27, 28, 45);
  - Quyết định 1624/QĐ-TTg (khoản 4, 15, 18, 20 Điều 1); Kết luận 51-KL/TW (mục 2.1, 2.2);
  - Thông tư 01/2024/TT-BGDĐT (mục 6.2.1 Phụ lục) và Thông tư 83/2026/TT-BGDĐT (Điều hiệu lực);
  - Nghị định 134/2026/NĐ-CP (Điều 5a); Quyết định 213/QĐ-ĐHTĐ (ngày ban hành).
Hai sơ đồ ảnh cũ không khớp nội dung chữ nên được thay bằng Bảng 1.1 và Bảng 1.2.
"""
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from khung import GOC, VanBanChung, luu  # noqa: E402

VAO = os.path.join(GOC, "Chuong_1_Co_so_ly_luan_chuan.md")

SUA = [
    # 1.2.3: mô tả đúng Thông tư 01/2024 (văn bằng có trong công thức 6.2.1) và nêu Thông tư 83/2026 sắp thay thế
    ("- Trong Chuẩn cơ sở giáo dục đại học quốc gia: Thông tư số 01/2024/TT-BGDĐT của Bộ Giáo dục và Đào tạo quy định "
     "Tiêu chuẩn 6 về nghiên cứu và đổi mới sáng tạo gồm hai tiêu chí định lượng cốt lõi: tỷ trọng nguồn thu từ hoạt động "
     "khoa học và công nghệ trên tổng thu của nhà trường (Tiêu chí 6.1) và số lượng công bố khoa học bình quân trên một "
     "giảng viên cơ hữu (Tiêu chí 6.2). Chuẩn không quy định trực tiếp chỉ số riêng rẽ về số lượng văn bằng sở hữu trí tuệ.",
     "- Trong Chuẩn cơ sở giáo dục đại học: Thông tư số 01/2024/TT-BGDĐT ngày 05 tháng 02 năm 2024 của Bộ trưởng Bộ Giáo "
     "dục và Đào tạo quy định Tiêu chuẩn 6 về nghiên cứu và đổi mới sáng tạo gồm hai tiêu chí định lượng: tỷ trọng thu từ "
     "hoạt động khoa học và công nghệ trên tổng thu đối với cơ sở có đào tạo tiến sĩ, Tiêu chí 6.1, và số công bố khoa học "
     "bình quân trên một giảng viên toàn thời gian, Tiêu chí 6.2. Văn bằng sở hữu trí tuệ có mặt trực tiếp trong công "
     "thức tính Tiêu chí 6.2: mỗi bằng độc quyền giải pháp hữu ích được tính như một công bố, mỗi bằng độc quyền sáng chế "
     "được tính gấp năm lần."),
    ("qua đó nâng cao tỷ trọng thu từ khoa học công nghệ để đáp ứng Tiêu chí 6.1. Đây chính là mối liên kết thực tiễn mà đề "
     "tài sẽ khảo sát và định lượng cụ thể tại Chương 2.",
     "qua đó nâng cao tỷ trọng thu từ khoa học công nghệ để đáp ứng Tiêu chí 6.1. Thông tư số 83/2026/TT-BGDĐT ngày 30 "
     "tháng 9 năm 2026 sẽ thay thế Thông tư số 01/2024/TT-BGDĐT từ ngày 15 tháng 11 năm 2026; những thay đổi liên quan "
     "đến sở hữu trí tuệ được phân tích như thời cơ và thách thức của Nhà trường tại Mục 3.1.2."),
    # 1.3.1: căn cứ sửa đổi và văn bản hợp nhất hiện hành
    ("Luật Sở hữu trí tuệ năm 2005 (được sửa đổi, bổ sung các năm 2009, 2019, 2022 và các luật liên quan).",
     "Luật Sở hữu trí tuệ số 50/2005/QH11, được sửa đổi, bổ sung bởi các Luật số 36/2009/QH12, số 42/2019/QH14, số "
     "07/2022/QH15, Luật Khoa học, công nghệ và đổi mới sáng tạo số 93/2025/QH15 và gần nhất là Luật số 131/2025/QH15 có "
     "hiệu lực từ ngày 01 tháng 4 năm 2026; nội dung hiện hành được hợp nhất tại Văn bản hợp nhất số 67/VBHN-VPQH ngày 23 "
     "tháng 3 năm 2026."),
    ("Luật Sở hữu trí tuệ đã bãi bỏ Điều 86a riêng rẽ trước đây và tích hợp quy định dẫn chiếu thống nhất tại điểm c khoản "
     "1 Điều 86.",
     "Điều 86a trước đây đã được bãi bỏ theo điểm h khoản 7 Điều 71 Luật số 93/2025/QH15, và nội dung này được quy định "
     "thống nhất tại điểm c khoản 1 Điều 86 do khoản 22 Điều 1 Luật số 131/2025/QH15 bổ sung."),
    ("Khoản 2 Điều 135 trước đây (quy định khung trần khống chế thù lao 10 - 15% và 15 - 20% đối với nhiệm vụ sử dụng ngân "
     "sách nhà nước) đã chính thức bị bãi bỏ bởi điểm h khoản 7 Điều 71 Luật Khoa học, công nghệ và đổi mới sáng tạo số "
     "93/2025/QH15. Theo quy định hiện hành tại khoản 1 Điều 135 Luật Sở hữu trí tuệ,",
     "Khoản 2 Điều 135 trước đây, quy định riêng về thù lao đối với sáng chế, kiểu dáng công nghiệp, thiết kế bố trí là kết "
     "quả của nhiệm vụ sử dụng ngân sách nhà nước, đã bị bãi bỏ theo điểm h khoản 7 Điều 71 Luật Khoa học, công nghệ và đổi "
     "mới sáng tạo số 93/2025/QH15. Theo khoản 1 Điều 135 Luật Sở hữu trí tuệ được sửa đổi theo điểm b khoản 7 Điều 71 "
     "Luật này,"),
    ("tương ứng với giá trị mà sáng chế, kiểu dáng công nghiệp đóng góp vào hoạt động sản xuất, kinh doanh;",
     "tương ứng với giá trị mà sáng chế, kiểu dáng công nghiệp, thiết kế bố trí đóng góp vào sản phẩm, dịch vụ hoặc hoạt "
     "động sản xuất, kinh doanh;"),
    ("Việc bãi bỏ khoản 2 Điều 135 có ý nghĩa bước ngoặt: pháp luật hiện hành đã xóa bỏ hoàn toàn mức trần khống chế thù "
     "lao đối với mọi nguồn kinh phí,",
     "Việc bãi bỏ khoản 2 Điều 135 có ý nghĩa bước ngoặt: pháp luật hiện hành không còn quy định riêng về mức thù lao theo "
     "nguồn kinh phí và không đặt trần thù lao,"),
    ("Ba là, về bảo hộ quyền tác giả và liêm chính học thuật: Các Điều 14, 19, 20, 22 và 25 phân định minh bạch giữa quyền "
     "nhân thân không thể chuyển giao của tác giả (quyền đứng tên, bảo vệ sự toàn vẹn của tác phẩm) và quyền tài sản (quyền "
     "sao chép, phân phối, làm tác phẩm phái sinh) thuộc về nhà trường khi đầu tư kinh phí giao nhiệm vụ biên soạn giáo "
     "trình, bài giảng, tài liệu điện tử và cơ sở dữ liệu. Luật cũng quy định chặt chẽ các giới hạn quyền tác giả để phục vụ "
     "mục đích giảng dạy, nghiên cứu phi thương mại, đồng thời thiết lập chế tài nghiêm khắc đối với hành vi đạo văn và sao "
     "chép học thuật trái phép.",
     "Ba là, về bảo hộ quyền tác giả và liêm chính học thuật: Điều 19 và Điều 20 phân định quyền nhân thân của tác giả, như "
     "quyền đứng tên và bảo vệ sự toàn vẹn của tác phẩm, với quyền tài sản như sao chép, phân phối, làm tác phẩm phái sinh. "
     "Theo khoản 1 Điều 39, tổ chức giao nhiệm vụ sáng tạo tác phẩm cho người thuộc tổ chức mình là chủ sở hữu các quyền "
     "tài sản, trừ trường hợp có thỏa thuận khác; vì vậy nhà trường là chủ sở hữu đối với giáo trình, bài giảng, tài liệu "
     "điện tử và cơ sở dữ liệu được giao biên soạn. Điều 14 và Điều 22 xác định các loại hình tác phẩm, trong đó có chương "
     "trình máy tính và sưu tập dữ liệu; Điều 25 quy định các trường hợp ngoại lệ phục vụ giảng dạy, nghiên cứu phi thương "
     "mại; Điều 28 quy định các hành vi xâm phạm quyền tác giả, là căn cứ xử lý hành vi sao chép học thuật trái phép."),
    ("Điển hình là Nghị định số 134/2026/NĐ-CP của Chính phủ", "Điển hình là Nghị định số 134/2026/NĐ-CP ngày 06 tháng 4 "
     "năm 2026 của Chính phủ"),
    ("trường hợp sản phẩm hoàn toàn do thuật toán tạo ra tự động thì không được bảo hộ quyền tác giả.",
     "khi không có sự đóng góp như vậy, quyền tác giả không phát sinh."),
    ("và Luật Giáo dục năm 2019 (được sửa đổi, bổ sung bởi Luật số 123/2025/QH15) cùng Luật Giáo dục đại học năm 2012 (được "
     "sửa đổi, bổ sung năm 2018) về quyền tự chủ nghiên cứu khoa học, chuyển giao công nghệ và hợp tác doanh nghiệp của các "
     "cơ sở giáo dục.",
     "Luật Giáo dục số 43/2019/QH14 được sửa đổi, bổ sung bởi Luật số 123/2025/QH15; và Luật Giáo dục đại học số "
     "125/2025/QH15 ngày 10 tháng 12 năm 2025, có hiệu lực từ ngày 01 tháng 01 năm 2026 và thay thế Luật Giáo dục đại học "
     "năm 2012. Luật Giáo dục đại học mới xác định tại điểm e khoản 3 Điều 27 rằng đăng ký bản quyền hoặc bảo hộ, khai "
     "thác và phát triển tài sản trí tuệ là một nội dung của hoạt động khoa học, công nghệ và đổi mới sáng tạo trong cơ sở "
     "giáo dục đại học; Điều 28 quy định quyền thành lập doanh nghiệp quản lý tài sản trí tuệ, quyền định giá, xác lập "
     "quyền sở hữu, khai thác, góp vốn, phân chia lợi ích từ tài sản trí tuệ và nghĩa vụ công khai kết quả hoạt động khoa "
     "học, công nghệ và đổi mới sáng tạo trên Nền tảng số quốc gia."),
    # 1.3.2: đối chiếu toàn văn Luật 93/2025/QH15 (Điều 25, 27, 28, 37, 66, 71, 72)
    ("Thứ nhất, phân cấp toàn diện quyền định đoạt tài sản trí tuệ: Các Điều 27 và 28 Luật số 93/2025/QH15 quy định kết quả "
     "nghiên cứu hình thành từ nhiệm vụ khoa học và công nghệ có sử dụng ngân sách nhà nước được giao quyền sở hữu trực tiếp "
     "cho tổ chức chủ trì là cơ sở giáo dục đại học. Nhà trường được toàn quyền định giá, chuyển giao, góp vốn hoặc thành "
     "lập doanh nghiệp khởi nguồn tri thức mà không phải thực hiện các thủ tục phê duyệt xử lý tài sản công phức tạp như "
     "giai đoạn trước.",
     "Thứ nhất, về quyền đối với kết quả nghiên cứu: khoản 2 Điều 25 Luật số 93/2025/QH15 quy định tổ chức chủ trì nhiệm vụ "
     "khoa học, công nghệ và đổi mới sáng tạo được Nhà nước tự động giao quyền quản lý, sử dụng, quyền sở hữu phần kết quả "
     "tương ứng với kinh phí từ ngân sách nhà nước, không phải thực hiện thủ tục giao quyền và không phải bồi hoàn chi phí; "
     "khoản 1 Điều 25 xác định tổ chức, cá nhân đóng góp tài sản, tài chính là chủ sở hữu kết quả tương ứng với tỷ lệ đóng "
     "góp. Điều 27 cho phép chủ sở hữu tự quyết định việc thương mại hóa, lựa chọn hình thức, giá, phương án góp vốn và "
     "phân chia lợi nhuận mà không phải thực hiện thủ tục xử lý tài sản công như giai đoạn trước. Điểm b khoản 2 Điều 37 "
     "khẳng định tổ chức ngoài công lập cũng được giao quyền sở hữu hoặc quyền sử dụng kết quả nghiên cứu từ nhiệm vụ sử "
     "dụng ngân sách nhà nước do mình thực hiện, và điểm d khoản 2 cho hưởng ưu đãi như đối với tổ chức công lập, qua đó "
     "mở ra cơ hội trực tiếp cho trường đại học tư thục. Luật Sở hữu trí tuệ dẫn chiếu cơ chế giao quyền này tại điểm c "
     "khoản 1 Điều 86 để trao quyền đăng ký sáng chế, kiểu dáng công nghiệp, thiết kế bố trí cho tổ chức được giao quyền."),
    ("Thứ hai, đổi mới cơ chế phân chia lợi ích thương mại hóa: Luật quy định tỷ lệ chia sẻ lợi nhuận thu được từ thương mại "
     "hóa kết quả nghiên cứu cho nhóm tác giả và các nhà khoa học trực tiếp thực hiện đề tài đạt mức tối thiểu 30%, tạo hành "
     "lang pháp lý mở nhằm thúc đẩy tinh thần dấn thân nghiên cứu ứng dụng của đội ngũ cán bộ khoa học.",
     "Thứ hai, về phân chia lợi nhuận: Điều 28 phân biệt hai trường hợp. Với phần lợi nhuận tương ứng kết quả không sử dụng "
     "ngân sách nhà nước, chủ sở hữu tự quyết định việc xử lý lợi nhuận, bao gồm thưởng cho tác giả, theo khoản 2. Với phần "
     "tương ứng kết quả sử dụng ngân sách nhà nước, điểm a khoản 3 quy định thưởng cho tác giả tối thiểu 30% lợi nhuận thu "
     "được từ cho thuê, bán, chuyển nhượng, chuyển giao quyền sử dụng, tự khai thác kết quả, hoặc tối thiểu 30% giá trị kết "
     "quả khi góp vốn, liên doanh, liên kết, thành lập doanh nghiệp; khoản 4 bổ sung rằng tác giả sáng chế, kiểu dáng công "
     "nghiệp, thiết kế bố trí còn được hưởng quyền lợi theo Luật Sở hữu trí tuệ. Cùng với đó, điểm b khoản 2 Điều 66 cho "
     "phép quỹ phát triển khoa học và công nghệ của tổ chức chi cho đăng ký, bảo hộ, quản lý, khai thác quyền sở hữu trí tuệ."),
    ("Luật Khoa học, công nghệ và đổi mới sáng tạo số 93/2025/QH15 đã tạo ra bước đột phá thể chế mang tính đồng bộ cao khi "
     "không chỉ quy định tỷ lệ phân chia tối thiểu 30% lợi nhuận thương mại hóa cho nhóm tác giả nghiên cứu, mà còn trực "
     "tiếp bãi bỏ Điều 86a, Điều 133a và khoản 2 Điều 135 của Luật Sở hữu trí tuệ. Bằng việc bãi bỏ khoản 2 Điều 135, cơ chế "
     "quản lý nhà nước đã chính thức xóa bỏ mức trần khống chế thù lao (15% và 20%) đối với kết quả nghiên cứu sử dụng ngân "
     "sách nhà nước.",
     "Luật số 93/2025/QH15 đã tạo ra bước đột phá thể chế mang tính đồng bộ khi vừa quy định mức thưởng tối thiểu 30% cho "
     "tác giả kết quả sử dụng ngân sách nhà nước, vừa sửa đổi Luật Sở hữu trí tuệ tại khoản 7 Điều 71: điểm b sửa khoản 1 "
     "Điều 135 theo nguyên tắc thỏa thuận với mức mặc định 10% và 15%; điểm h bãi bỏ Điều 86a, Điều 133a và khoản 2 Điều "
     "135. Pháp luật vì vậy không còn quy định riêng về khung thù lao theo nguồn kinh phí như trước."),
    ("Pháp luật hiện hành đã trao toàn quyền tự chủ thỏa thuận cho cơ sở giáo dục đại học và không hề khống chế mức trần thù "
     "lao đối với bất kỳ nguồn kinh phí nào.",
     "Pháp luật hiện hành trao quyền tự chủ thỏa thuận cho cơ sở giáo dục đại học, không đặt mức trần thù lao, đồng thời bảo "
     "đảm mức sàn 30% cho tác giả đối với kết quả sử dụng ngân sách nhà nước."),
    ("Vì vậy, trên bình diện pháp lý quốc gia hoàn toàn không có bất kỳ rào cản nào ngăn cản Trường Đại học Thành Đô quy "
     "định tỷ lệ phân chia lợi nhuận thương mại hóa ở mức 30%, 50% hoặc cao hơn trong quy chế nội bộ.",
     "Vì vậy, pháp luật không ngăn cản Trường Đại học Thành Đô quy định tỷ lệ phân chia lợi nhuận thương mại hóa cho tác "
     "giả ở mức cao hơn mức tối thiểu và mức mặc định trong quy chế nội bộ."),
    ("Rào cản duy nhất hiện nay kìm hãm động lực thương mại hóa của nhà khoa học là các rào cản do chính Nhà trường tự đặt "
     "ra trong các văn bản quản trị nội bộ: mức khống chế trần tiền thưởng tối đa 100 triệu đồng trong Quy chế hoạt động "
     "khoa học công nghệ (ban hành kèm theo Quyết định số 213/QĐ-ĐHTĐ ngày 27/5/2021) và mức trần 50 triệu đồng trong Quy "
     "chế chi tiêu nội bộ. Đây chính là điểm nghẽn thể chế cốt lõi tại chỗ, cung cấp căn cứ pháp lý và thực tiễn vững chắc "
     "để đề tài đề xuất bãi bỏ hoàn toàn các mức trần hành chính này, hoàn thiện Quy chế quản trị tài sản trí tuệ mới của "
     "Trường Đại học Thành Đô tại Chương 3.",
     "Các quy chế nội bộ của Nhà trường, gồm Quy chế hoạt động khoa học công nghệ ban hành kèm Quyết định số 213/QĐ-ĐHTĐ "
     "ngày 28 tháng 12 năm 2021 và Quy chế quản trị tài sản trí tuệ ban hành kèm Quyết định số 217/QĐ-ĐHTĐ ngày 21 tháng "
     "11 năm 2024, đều được xây dựng trước khi các luật mới có hiệu lực. Một số quy định như khoản nộp ngân sách và mức trần "
     "100 triệu đồng tại điểm a khoản 4 Điều 36 Quyết định 213 vì vậy cần được cập nhật theo Điều 28 Luật số 93/2025/QH15. "
     "Đây là căn cứ để Chương 2 đánh giá tính tương thích của quy chế nội bộ và Chương 3 đề xuất hợp nhất quy chế theo hướng "
     "đón đầu luật mới."),
    # 1.3.3: đối chiếu Kết luận 51-KL/TW và Quyết định 1624/QĐ-TTg
    ("coi quyền sở hữu trí tuệ là nguồn lực kinh tế đặc biệt;", "quán triệt quan điểm quyền sở hữu trí tuệ là nguồn lực quan "
     "trọng của quốc gia;"),
    ("đã xác định rõ cơ sở giáo dục đại học cùng các viện nghiên cứu và doanh nghiệp là lực lượng nòng cốt trong việc tạo ra "
     "và khai thác quyền sở hữu trí tuệ. Chiến lược đặt ra bốn yêu cầu mới đối với các trường đại học:",
     "đặt ra bốn nhóm yêu cầu mới đối với cơ sở giáo dục đại học, được quy định tại Mục II và Mục III Điều 1 Quyết định số "
     "1068/QĐ-TTg đã được sửa đổi:"),
    ("  1. Sử dụng các chỉ số đo lường sở hữu trí tuệ làm căn cứ đánh giá hiệu quả hoạt động hàng năm của cơ sở giáo dục đại "
     "học.",
     "  1. Sử dụng các chỉ số đo lường về sở hữu trí tuệ làm căn cứ đánh giá hiệu quả hoạt động của cơ sở giáo dục đại học "
     "và xác định các đối tượng quyền sở hữu trí tuệ cần đạt được đối với kết quả nghiên cứu sử dụng ngân sách nhà nước, "
     "theo điểm b khoản 4 Mục III."),
    ("  2. Quy định các trường khối kỹ thuật, công nghệ tiến hành thủ tục đăng ký bảo hộ sở hữu công nghiệp đồng thời với "
     "việc công bố bài báo khoa học.",
     "  2. Các cơ sở giáo dục khối kỹ thuật, công nghệ tiến hành thủ tục đăng ký bảo hộ sở hữu công nghiệp đồng thời với việc "
     "công bố bài báo khoa học về các kết quả nghiên cứu có tính ứng dụng cao, theo điểm b khoản 4 Mục III."),
    ("  3. Thí điểm hỗ trợ xác định giá trị quyền sở hữu trí tuệ cho các cơ sở giáo dục đại học và phát triển các trung tâm "
     "chuyển giao công nghệ chuyên trách.",
     "  3. Thí điểm hỗ trợ xác định giá trị của ít nhất 100 quyền sở hữu trí tuệ của viện nghiên cứu, cơ sở giáo dục đại học, "
     "doanh nghiệp khởi nghiệp sáng tạo, theo điểm đ khoản 5 Mục II; phát triển các trung tâm tư vấn, hỗ trợ định giá, khai "
     "thác thương mại quyền sở hữu trí tuệ trong cơ sở giáo dục đại học, theo điểm a khoản 6 Mục III."),
    ("  4. Đưa nội dung sở hữu trí tuệ cùng kỹ năng khai thác thương mại thành nội dung học bắt buộc trong chương trình đào "
     "tạo đại học.",
     "  4. Xây dựng chương trình giáo dục, đào tạo về sở hữu trí tuệ và nghiên cứu đưa sở hữu trí tuệ, kỹ năng khai thác "
     "thương mại quyền sở hữu trí tuệ thành nội dung học bắt buộc tại cơ sở giáo dục đại học, theo điểm b khoản 8 Mục III."),
    # 1.4.1: chu trình trình bày bằng bảng thay cho sơ đồ ảnh không khớp nội dung
    ("Gắn với dòng đời của tài sản trí tuệ trong trường đại học, chu trình quản lý vận hành qua bốn giai đoạn kế tiếp nhau:",
     "Gắn với dòng đời của tài sản trí tuệ trong trường đại học, chu trình quản lý vận hành qua bốn giai đoạn kế tiếp nhau, "
     "được tóm tắt tại Hình 1.1. Hoạt động bảo vệ quyền đồng thời diễn ra xuyên suốt cả bốn giai đoạn như đã phân tích tại "
     "Mục 1.1.3.\n[[HINH_1_1]]"),
    # 1.4.2: bộ 16 tiêu chí trình bày bằng bảng
    ("Hiệu quả quản lý quyền sở hữu trí tuệ được lượng hóa và đánh giá khách quan thông qua chuỗi logic bốn mắt xích:",
     "Hiệu quả quản lý quyền sở hữu trí tuệ được lượng hóa và đánh giá khách quan thông qua chuỗi logic bốn mắt xích: đầu "
     "vào, quá trình, đầu ra và kết quả. Mười sáu tiêu chí thuộc bốn nhóm được trình bày tại Bảng 1.1; đây cũng là bộ tiêu "
     "chí được dùng để đánh giá khả năng đo lường của hệ thống dữ liệu tại Chương 2.\n[[BANG_1_2]]\nChuỗi bốn nhóm tiêu chí "
     "cho phép phân biệt giữa việc nhà trường đã đầu tư và tổ chức bao nhiêu với việc các nỗ lực đó đã tạo ra bao nhiêu "
     "quyền được xác lập và mang lại giá trị gì. Một hệ thống quản lý chỉ đo được đầu vào và đầu ra mà không đo được kết quả "
     "sẽ không đánh giá được hiệu quả thực chất của hoạt động sở hữu trí tuệ."),
    # 1.5.2: không dùng ngoặc đơn để giải thích thuật ngữ tiếng Anh
    ("mô hình Không gian sáng tạo mở thử nghiệm (được định danh là Living Lab trong Thuyết minh và Đề cương nghiên cứu đã "
     "được phê duyệt) được tiếp cận",
     "mô hình Không gian sáng tạo mở thử nghiệm, được thuyết minh đề tài gọi là mô hình Living Lab, được tiếp cận"),
    ("Theo đó, sáng chế không bị coi là mất tính mới nếu được người có quyền đăng ký hoặc người có được thông tin trực tiếp "
     "hoặc gián tiếp từ người đó công bố công khai, với điều kiện đơn đăng ký sáng chế phải được nộp tại Cục Sở hữu trí tuệ "
     "trong thời hạn 12 tháng kể từ ngày công bố; hoặc giải pháp được trưng bày tại các cuộc triển lãm chính thức được công "
     "nhận.",
     "Theo đó, sáng chế không bị coi là mất tính mới nếu được người có quyền đăng ký hoặc người có được thông tin trực tiếp "
     "hoặc gián tiếp từ người đó bộc lộ công khai, với điều kiện đơn đăng ký sáng chế được nộp tại Việt Nam trong thời hạn "
     "12 tháng kể từ ngày bộc lộ; khoản 4 mở rộng ngoại lệ này cho trường hợp sáng chế bị bộc lộ trong đơn hoặc văn bằng do "
     "cơ quan quản lý nhà nước công bố không phù hợp với quy định hoặc đơn do người không có quyền đăng ký nộp."),
    # 1.5.3: khung phân tích khớp với Chương 2, Chương 3 bản cuối
    ("đề tài xây dựng Khung phân tích logic xuyên suốt cho toàn bộ công trình nghiên cứu:",
     "đề tài xây dựng Khung phân tích logic xuyên suốt cho toàn bộ công trình nghiên cứu, được trình bày tại Hình 1.3.\n"
     "[[HINH_1_3]]"),
    (" (đặc biệt là Khoa Dược, khối kỹ thuật công nghệ và nhóm giáo trình, học liệu đào tạo)",
     ", đặc biệt là lĩnh vực dược của Viện Y - Dược, khối kỹ thuật công nghệ và nhóm giáo trình, học liệu đào tạo"),
    ("văn hóa học thuật tôn trọng quyền sở hữu trí tuệ, liêm chính nghiên cứu trong toàn trường.",
     "văn hóa học thuật tôn trọng quyền sở hữu trí tuệ, liêm chính nghiên cứu trong toàn trường.\n[[HINH_1_2]]"),
    ("Các trường đại học hàng đầu thế giới thường duy trì tỷ lệ chia sẻ lợi nhuận cho nhà sáng chế từ 40% đến 60%, thậm chí "
     "lên đến 70% đối với các công nghệ giai đoạn đầu, tạo động lực vật chất tối đa cho các nhà khoa học (Etzkowitz, năm 2003).",
     "Tỷ lệ chia sẻ doanh thu dành cho nhà sáng chế là một yếu tố động lực quan trọng; các trường đại học thành công trong "
     "chuyển giao thường dành cho nhà sáng chế một tỷ lệ đáng kể trong doanh thu cấp phép (Siegel et al., 2007)."),
    # tiểu kết
    ("tác động gián tiếp đến các tiêu chí đánh giá chuẩn cơ sở giáo dục đại học.",
     "tác động đến các tiêu chí của Chuẩn cơ sở giáo dục đại học."),
    ("làm rõ quy định về quyền đăng ký theo Điều 86 Luật Sở hữu trí tuệ dẫn chiếu sang Luật Khoa học, công nghệ và đổi mới "
     "sáng tạo số 93/2025/QH15; phân tích việc bãi bỏ Điều 86a, Điều 133a và khoản 2 Điều 135 Luật Sở hữu trí tuệ đã xóa bỏ "
     "hoàn toàn mức trần khống chế thù lao,",
     "làm rõ cơ chế tự động giao quyền sở hữu kết quả nghiên cứu tại Điều 25 và mức thưởng tối thiểu 30% cho tác giả kết "
     "quả sử dụng ngân sách nhà nước tại Điều 28 Luật số 93/2025/QH15, quyền đăng ký theo điểm c khoản 1 Điều 86 Luật Sở "
     "hữu trí tuệ do Luật số 131/2025/QH15 bổ sung; phân tích việc "
     "Luật số 93/2025/QH15 bãi bỏ Điều 86a, Điều 133a và khoản 2 Điều 135 Luật Sở hữu trí tuệ, theo đó pháp luật không còn "
     "đặt trần thù lao,"),
    ("định vị mô hình Không gian sáng tạo mở thử nghiệm (Living Lab) như", "định vị mô hình Không gian sáng tạo mở thử "
     "nghiệm như"),
    ("thiết lập bộ tiêu chí đánh giá hiệu quả 4 cấp độ và nhận diện 6 nhóm yếu tố ảnh hưởng",
     "thiết lập bộ 16 tiêu chí đánh giá hiệu quả theo chuỗi đầu vào, quá trình, đầu ra, kết quả và nhận diện 6 nhóm yếu tố "
     "ảnh hưởng"),
]

# Các khối được thay hoàn toàn (xóa khỏi bản markdown trước khi dựng)
BO_KHOI = [
    (r"- Giai đoạn tạo lập và nhận diện quyền:.*?\*Hình 1\.1\. Chu trình quản lý quyền sở hữu trí tuệ trong trường đại học\*\n",
     ""),
    (r"1\. \*\*Nhóm tiêu chí đầu vào:\*\*.*?\*Hình 1\.2\. Khung tiêu chí đánh giá hiệu quả quản lý quyền sở hữu trí tuệ trong "
     r"cơ sở giáo dục đại học\*\n", ""),
    (r"\| Cấp độ phân tích \|.*?\*Hình 1\.3\. Khung phân tích logic xuyên suốt của đề tài\*\n", ""),
]

BANG_1_1 = dict(
    tieu_de="Chu trình bốn giai đoạn quản lý quyền sở hữu trí tuệ trong trường đại học",
    cot=["Giai đoạn", "Hoạt động nghiệp vụ chủ yếu", "Sản phẩm quản lý"],
    dong=[["1. Tạo lập và nhận diện quyền",
           "Tiếp nhận ý tưởng, sàng lọc đề tài tiềm năng, khai báo phát minh nội bộ, thẩm định tính mới trước khi công bố",
           "Phiếu khai báo, biên bản rà soát khả năng bảo hộ"],
          ["2. Xác lập quyền",
           "Chuẩn bị hồ sơ, bố trí kinh phí, nộp đơn cấp văn bằng sở hữu công nghiệp hoặc đăng ký quyền tác giả",
           "Đơn đăng ký, văn bằng, giấy chứng nhận"],
          ["3. Khai thác và thương mại hóa",
           "Sử dụng trong giảng dạy, chuyển giao quyền sử dụng, góp vốn, thành lập doanh nghiệp, sản xuất thử nghiệm",
           "Hợp đồng chuyển giao, cấp phép, doanh thu"],
          ["4. Bảo vệ và phân chia lợi ích",
           "Giám sát xâm phạm, bảo vệ quyền lợi của nhà trường và tác giả, chi trả thù lao, trích quỹ tái đầu tư",
           "Sổ theo dõi hiệu lực, chứng từ chi trả, quỹ phát triển khoa học công nghệ"]],
    nguon="Nguồn: Nhóm nghiên cứu tổng hợp từ Bradley và cộng sự, năm 2013; Tổ chức Sở hữu trí tuệ thế giới, năm 2020.",
    rong=[3.2, 7.4, 4.4], can=["left", "left", "left"],
)

BANG_1_2 = dict(
    tieu_de="Bộ tiêu chí đánh giá hiệu quả quản lý quyền sở hữu trí tuệ theo chuỗi",
    cot=["Nhóm", "Tiêu chí"],
    dong=[["Đầu vào", "1. Kinh phí nộp đơn, duy trì văn bằng, khen thưởng sáng tạo\n2. Số lượng và trình độ nhân lực quản "
                      "lý sở hữu trí tuệ, chuyển giao công nghệ\n3. Cơ sở dữ liệu tra cứu sáng chế, hạ tầng kỹ thuật, phần mềm "
                      "kiểm tra trùng lặp"],
          ["Quá trình", "4. Mức độ tương thích của quy chế nội bộ với pháp luật hiện hành\n5. Thời gian xử lý trung bình một "
                        "hồ sơ đề xuất bảo hộ\n6. Mức độ chuẩn hóa biểu mẫu và quy trình phối hợp\n7. Số lớp tập huấn sở hữu "
                        "trí tuệ cho giảng viên, sinh viên"],
          ["Đầu ra", "8. Số đơn đăng ký sở hữu công nghiệp\n9. Số văn bằng bảo hộ được cấp\n10. Số giấy chứng nhận đăng ký "
                     "quyền tác giả cho giáo trình, bài giảng, phần mềm\n11. Số công trình có sản phẩm chuyển hóa được thành "
                     "tài sản trí tuệ có khả năng bảo hộ"],
          ["Kết quả", "12. Số hợp đồng chuyển giao, cấp phép sử dụng tài sản trí tuệ\n13. Nguồn thu từ khai thác tài sản trí "
                      "tuệ\n14. Tỷ trọng nguồn thu khoa học công nghệ trên tổng thu, hướng đến mức 5% theo Chuẩn cơ sở giáo "
                      "dục đại học\n15. Tác động đến kiểm định chất lượng và thứ hạng xếp hạng\n16. Mức độ hài lòng và động lực "
                      "đổi mới sáng tạo của giảng viên"]],
    nguon="Nguồn: Nhóm nghiên cứu xây dựng trên cơ sở tiếp cận chuỗi đầu vào, quá trình, đầu ra, kết quả.",
    rong=[2.6, 12.4], can=["center", "left"],
)

BANG_1_3 = dict(
    tieu_de="Khung phân tích của đề tài",
    cot=["Cấp độ phân tích", "Nội dung cốt lõi và mắt xích liên kết"],
    dong=[["1. Bối cảnh và căn cứ pháp lý",
           "- Luật Sở hữu trí tuệ hợp nhất năm 2026, Luật Khoa học, công nghệ và đổi mới sáng tạo năm 2025, Luật Giáo dục "
           "đại học năm 2025, Chiến lược sở hữu trí tuệ đến năm 2030 được sửa đổi năm 2026.\n- Chuẩn cơ sở giáo dục đại học "
           "theo Thông tư số 01/2024/TT-BGDĐT, được thay thế bởi Thông tư số 83/2026/TT-BGDĐT từ ngày 15 tháng 11 năm 2026, "
           "và tiêu chuẩn kiểm định chất lượng."],
          ["2. Khung lý thuyết quản lý quyền sở hữu trí tuệ",
           "- Bốn chức năng quản lý: hoạch định, tổ chức, thực hiện, kiểm tra và giám sát.\n- Chu trình bốn giai đoạn: tạo "
           "lập và nhận diện, xác lập, khai thác và thương mại hóa, bảo vệ và phân chia lợi ích.\n- Mười sáu tiêu chí theo "
           "chuỗi đầu vào, quá trình, đầu ra, kết quả.\n- Sáu nhóm yếu tố ảnh hưởng: thể chế, tổ chức bộ máy, tài chính, "
           "dữ liệu, con người, động lực và văn hóa."],
          ["3. Thực trạng tại Trường Đại học Thành Đô, Chương 2",
           "- Nguồn hình thành tài sản: 582 bản ghi sản phẩm khoa học, 38 đề tài cấp cơ sở, 12 tài sản trí tuệ giai đoạn "
           "2021 - 2025.\n- Thể chế, tổ chức và nguồn lực: Quyết định số 213/QĐ-ĐHTĐ, Quyết định số 217/QĐ-ĐHTĐ, Quy chế chi "
           "tiêu nội bộ, Kế hoạch số 07/KH-ĐHTĐ; bộ máy, kinh phí, cơ chế khuyến khích, dữ liệu.\n- Chuỗi chuyển hóa từ đề "
           "tài sang đơn đăng ký và mức độ khai thác.\n- Đánh giá tổng hợp: kết quả, hạn chế, nguyên nhân khách quan và chủ "
           "quan."],
          ["4. Hệ thống giải pháp, Chương 3",
           "- Phân tích điểm mạnh, điểm yếu, thời cơ, thách thức và năm nguyên tắc xây dựng giải pháp.\n- Năm nhóm giải "
           "pháp: thể chế, bộ máy, quy trình 8 khâu, tài chính, đào tạo và văn hóa.\n- Mô-đun Không gian sáng tạo mở thử "
           "nghiệm trong quy trình 8 khâu, gắn với cam kết bảo mật và giữ ngày ưu tiên.\n- Thí điểm rà soát tại Viện Y - "
           "Dược, lộ trình triển khai và bộ chỉ số theo dõi."]],
    nguon="Nguồn: Nhóm nghiên cứu xây dựng.",
    rong=[4.0, 11.0], can=["left", "left"],
)


# Trích dẫn trong bài theo APA 7: (Tác giả, năm); hai tác giả dùng "&", từ ba tác giả dùng "et al.".
# Tài liệu không kiểm chứng được trong thư mục Zotero và "Co so ly luan" bị lược trích dẫn.
APA = [
    ("(Bradley và cộng sự, năm 2013; Goldfarb và Henrekson, năm 2003)", "(Bradley et al., 2013; Goldfarb & Henrekson, 2003)"),
    ("(Bradley và cộng sự, năm 2013)", "(Bradley et al., 2013)"),
    ("(Etzkowitz, năm 2003; Shane, năm 2004; Perkmann và cộng sự, năm 2013)", "(Etzkowitz, 2003; Perkmann et al., 2013; "
     "Shane, 2004)"),
    ("(Fisher, năm 2001; Guan, năm 2014)", "(Fisher, 2001; Guan, 2014)"),
    ("(Fisher, năm 2001)", "(Fisher, 2001)"),
    ("(Nguyễn Minh Huyền Trang, năm 2025)", "(Nguyễn, 2025)"),
    ("(Rialti và cộng sự, năm 2022)", "(O’Dwyer et al., 2023)"),  # Zotero ghi sai tác giả; đã đối chiếu DOI 10.1007/s10961-022-09932-2
    ("(Shane, năm 2004; Mowery và cộng sự, năm 2004)", "(Shane, 2004)"),
    ("(Shane, năm 2004; Võ Nguyên Hoàng Phúc, năm 2025)", "(Shane, 2004; Võ, 2025)"),
    ("(Siegel và cộng sự, năm 2007; Thursby và Thursby, năm 2002)", "(Siegel et al., 2007; Thursby & Kemp, 2002)"),
    ("(Teece, năm 2018)", "(Teece, 2018)"),
    (" (Tewari và Bhardwaj, năm 2020)", ""),
    ("(Thursby và Thursby, năm 2002)", "(Perkmann et al., 2013)"),
    ("(Tổ chức Sở hữu trí tuệ thế giới, năm 2020)", "(Tổ chức Sở hữu trí tuệ thế giới, 2020)"),
    ("(Văn phòng Quốc hội, năm 2026)", "(Văn phòng Quốc hội, 2026)"),
    ("Tổ chức Sở hữu trí tuệ thế giới (năm 2020)", "Tổ chức Sở hữu trí tuệ thế giới (2020)"),
    ("Bổ sung cho góc nhìn này, Milliken và Allen (năm 2013) nhấn mạnh đây là",
     "Điểm chung của các sản phẩm này là"),  # nguồn Milliken và Allen không truy xuất được
    ("Fisher (năm 2001)", "Fisher (2001)"),
    ("Guan (năm 2014)", "Guan (2014)"),
    ("Cục Sở hữu trí tuệ (năm 2020) đã phân nhánh", "tài liệu tập huấn của Cục Sở hữu trí tuệ (n.d.) dành cho cán bộ "
     "trường đại học, viện nghiên cứu đã phân nhánh"),
]


def chuan_bi():
    md = open(VAO, encoding="utf-8").read()
    for mau, moi in BO_KHOI:
        md, n = re.subn(mau, moi, md, flags=re.S)
        assert n == 1, mau[:60]
    for cu, moi in SUA:
        assert md.count(cu) == 1, f"{md.count(cu)} lần: {cu[:80]}"
        md = md.replace(cu, moi)
    for cu, moi in APA:
        assert md.count(cu) == 1, f"APA {md.count(cu)} lần: {cu[:80]}"
        md = md.replace(cu, moi)
    assert not re.search(r", năm \d{4}\)|\(năm \d{4}\)", md), re.findall(r".{40}(?:, năm \d{4}\)|\(năm \d{4}\))", md)
    return md


# Lượt chỉnh sửa theo bản góp ý ba chương (tháng 10 năm 2026). Mỗi mục: (đầu đoạn hiện có, nội dung mới).
# Nội dung mới là chuỗi hoặc danh sách chuỗi (các chuỗi sau được chèn thành đoạn mới ngay sau đoạn được thay).
SUA_LAN6 = [
    ("- Giải quyết xung đột lợi ích giữa nhà trường và nhà nghiên cứu:",
     "- Giải quyết xung đột lợi ích giữa nhà trường và nhà nghiên cứu: Nhà trường đóng vai trò là chủ đầu tư cung cấp cơ "
     "sở vật chất, phòng thí nghiệm và kinh phí; trong khi giảng viên là người trực tiếp lao động sáng tạo (Shane, 2004; "
     "Võ, 2025). Nếu quy chế nội bộ không quy định minh bạch các khoản thưởng, thù lao dành cho tác giả và cách phân "
     "chia khoản thu từ chuyển giao, nhà khoa học có thể giữ lại kết quả nghiên cứu để tự khai thác bên ngoài, gây thất "
     "thoát tài sản của nhà trường. Đồng thời, nhà quản lý phải giải quyết hài hòa xung đột giữa áp lực công bố bài báo "
     "sớm để tính điểm chức danh và yêu cầu giữ bí mật để nộp đơn đăng ký bảo hộ sáng chế."),
    ("- Tạo lập chu trình tái đầu tư khép kín:",
     "- Tạo lập chu trình tái đầu tư khép kín: Sau khi thực hiện các nghĩa vụ đối với tác giả theo pháp luật và quy chế "
     "nội bộ, gồm thù lao theo Điều 135 Luật Sở hữu trí tuệ đối với sáng chế, kiểu dáng công nghiệp, thiết kế bố trí và "
     "tiền thưởng theo Điều 28 Luật Khoa học, công nghệ và đổi mới sáng tạo số 93/2025/QH15 khi kết quả thuộc phạm vi "
     "điều chỉnh của điều này, phần còn lại của nguồn thu từ thương mại hóa có thể được trích bổ sung vào quỹ phát triển "
     "khoa học và công nghệ của trường để nâng cấp phòng thí nghiệm và tài trợ cho các đề tài nghiên cứu tiếp theo."),
    ("Việc bãi bỏ khoản 2 Điều 135 có ý nghĩa bước ngoặt:",
     ["Việc bãi bỏ khoản 2 Điều 135 có nghĩa là Luật Sở hữu trí tuệ không còn quy định một khung thù lao riêng cho sáng "
      "chế, kiểu dáng công nghiệp, thiết kế bố trí là kết quả của nhiệm vụ sử dụng ngân sách nhà nước, và khoản 1 Điều "
      "135 không đặt mức trần thù lao. Điều này không có nghĩa là nguồn kinh phí không còn ảnh hưởng đến lợi ích của tác "
      "giả: đối với kết quả sử dụng ngân sách nhà nước, Điều 28 Luật số 93/2025/QH15 đặt ra yêu cầu riêng về thưởng cho "
      "tác giả, được phân tích tại Mục 1.3.2.",
      "Để tránh đồng nhất các khoản lợi ích khác nhau, Đề tài sử dụng bốn thuật ngữ theo nghĩa sau. Thù lao của tác giả "
      "là khoản chủ sở hữu sáng chế, kiểu dáng công nghiệp, thiết kế bố trí phải trả cho tác giả theo Điều 135 Luật Sở "
      "hữu trí tuệ trong suốt thời hạn bảo hộ, theo thỏa thuận hoặc theo mức mặc định tính trên lợi nhuận trước thuế hay "
      "trên tổng số tiền nhận được mỗi lần trước thuế. Thưởng cho tác giả là khoản trích từ lợi nhuận thương mại hóa kết "
      "quả nghiên cứu theo Điều 28 Luật số 93/2025/QH15, bắt buộc tối thiểu 30% lợi nhuận sau thuế đối với phần kết quả "
      "sử dụng ngân sách nhà nước, và do chủ sở hữu tự quyết định đối với phần không sử dụng ngân sách nhà nước. Nhuận "
      "bút là khoản chi trả cho tác giả tác phẩm như sách, giáo trình theo pháp luật về quyền tác giả và quy chế chi tiêu "
      "của tổ chức. Phân chia lợi nhuận, hay phân chia nguồn thu, là việc chủ sở hữu phân bổ khoản thu cho các bên như tổ "
      "chức chủ trì, đơn vị có tác giả, quỹ phát triển khoa học và công nghệ. Bốn khoản này khác nhau về căn cứ pháp lý, "
      "đối tượng hưởng và cơ sở tính, nên tỷ lệ của khoản này không thể so sánh trực tiếp với tỷ lệ của khoản khác."]),
    ("Thứ hai, về phân chia lợi nhuận: Điều 28 phân biệt hai trường hợp.",
     ["Thứ hai, về phân chia lợi nhuận từ thương mại hóa: Điều 28 phân biệt theo nguồn hình thành kết quả. Theo khoản "
      "2, với phần lợi nhuận tương ứng kết quả không sử dụng ngân sách nhà nước, chủ sở hữu tự quyết định việc xử lý lợi "
      "nhuận, bao gồm thưởng cho tác giả. Theo khoản 3, với phần lợi nhuận tương ứng phần kết quả sử dụng ngân sách nhà "
      "nước, tổ chức chủ trì sử dụng lợi nhuận sau thuế để: thưởng cho tác giả tối thiểu 30% lợi nhuận thu được từ cho "
      "thuê, bán, chuyển nhượng, chuyển giao quyền sử dụng, tự khai thác, sử dụng kết quả, hoặc tối thiểu 30% giá trị kết "
      "quả khi góp vốn, hợp tác, liên doanh, liên kết, thành lập doanh nghiệp; thưởng cho cá nhân có đóng góp trực tiếp "
      "vào hoạt động tổ chức thương mại hóa; tái đầu tư cho hoạt động khoa học, công nghệ và đổi mới sáng tạo; và mục đích "
      "khác. Khoản 4 quy định khi kết quả sử dụng ngân sách nhà nước là sáng chế, thiết kế bố trí, kiểu dáng công nghiệp, "
      "giống cây trồng được bảo hộ, tác giả vừa hưởng khoản thưởng nêu trên vừa hưởng các quyền lợi khác theo Luật Sở hữu "
      "trí tuệ, trong đó có thù lao theo Điều 135. Khoản 5 xác định phần thưởng là mức dành chung cho các đồng tác giả; "
      "khoản 6 giao Chính phủ quy định chi tiết. Cùng với đó, điểm b khoản 2 Điều 66 cho phép quỹ phát triển khoa học và "
      "công nghệ của tổ chức chi cho đăng ký, bảo hộ, quản lý, khai thác quyền sở hữu trí tuệ.",
      "Về áp dụng theo thời gian, khoản 3 Điều 73 quy định nhiệm vụ khoa học và công nghệ đã được phê duyệt giao chủ trì "
      "trước ngày 01 tháng 10 năm 2025 tiếp tục thực hiện theo Luật Khoa học và công nghệ số 29/2013/QH13, Nghị quyết số "
      "193/2025/QH15 và văn bản hướng dẫn có hiệu lực tại thời điểm phê duyệt. Khoản 7 Điều 73 quy định ngoại lệ: Điều 28 "
      "được áp dụng đối với lợi nhuận chưa phân chia từ thương mại hóa sáng chế, kiểu dáng công nghiệp, thiết kế bố trí, "
      "giống cây trồng là kết quả của nhiệm vụ được giao từ ngày 01 tháng 01 năm 2023 đến trước ngày 01 tháng 10 năm 2025 "
      "và đã được cấp văn bằng bảo hộ. Như vậy, để xác định nghĩa vụ đối với tác giả trong một trường hợp cụ thể cần trả "
      "lời đồng thời ba câu hỏi: kết quả có sử dụng ngân sách nhà nước hay không và ở phần nào; nhiệm vụ được giao vào "
      "thời điểm nào; đối tượng có phải là sáng chế, kiểu dáng công nghiệp, thiết kế bố trí đã được bảo hộ hay không."]),
    ("Luật số 93/2025/QH15 đã tạo ra bước đột phá thể chế mang tính đồng bộ",
     "Luật số 93/2025/QH15 điều chỉnh đồng thời hai cơ chế: cơ chế thưởng cho tác giả khi thương mại hóa kết quả nghiên "
     "cứu tại Điều 28, và cơ chế thù lao cho tác giả sáng chế, kiểu dáng công nghiệp, thiết kế bố trí trong Luật Sở hữu "
     "trí tuệ thông qua khoản 7 Điều 71: điểm b sửa khoản 1 Điều 135 theo nguyên tắc thỏa thuận với mức mặc định 10% và "
     "15%; điểm h bãi bỏ Điều 86a, Điều 133a và khoản 2 Điều 135. Sau sửa đổi, khung thù lao riêng theo nguồn kinh phí "
     "không còn nằm trong Luật Sở hữu trí tuệ, còn yêu cầu riêng đối với kết quả sử dụng ngân sách nhà nước được quy định "
     "tại Điều 28 dưới hình thức mức thưởng tối thiểu."),
    ("Sự đổi mới căn bản này mang lại kết luận pháp lý",
     "Từ các quy định trên, Đề tài rút ra ba nhận định làm căn cứ cho Chương 2 và Chương 3. Một là, nguồn kinh phí vẫn "
     "quyết định cơ chế áp dụng: với phần kết quả không sử dụng ngân sách nhà nước, chủ sở hữu tự quyết định việc phân "
     "chia lợi nhuận và mức thưởng theo khoản 2 Điều 28; với phần kết quả sử dụng ngân sách nhà nước thuộc phạm vi áp "
     "dụng của Điều 28, phần thưởng cho tác giả không thấp hơn 30% lợi nhuận sau thuế, hoặc 30% giá trị kết quả khi góp "
     "vốn. Hai là, thù lao theo Điều 135 là nghĩa vụ riêng của chủ sở hữu đối với tác giả sáng chế, kiểu dáng công "
     "nghiệp, thiết kế bố trí, có cơ sở tính khác và không bị thay thế bởi khoản thưởng theo khoản 4 Điều 28; Luật không "
     "đặt trần đối với thù lao, và mức 10%, 15% chỉ áp dụng khi các bên không có thỏa thuận. Ba là, quy chế nội bộ có "
     "thể quy định mức có lợi hơn cho tác giả, nhưng không thể quy định mức thấp hơn mức tối thiểu của Luật đối với các "
     "trường hợp thuộc phạm vi bắt buộc; một trường hợp cụ thể thuộc cơ chế nào phụ thuộc vào nguồn kinh phí, thời điểm "
     "giao nhiệm vụ và loại đối tượng như đã nêu."),
    ("Các quy chế nội bộ của Nhà trường, gồm Quy chế hoạt động khoa học công nghệ",
     "Các quy chế nội bộ của Nhà trường, gồm Quy chế hoạt động khoa học công nghệ ban hành kèm Quyết định số "
     "213/QĐ-ĐHTĐ ngày 28 tháng 12 năm 2021 và Quy chế quản trị tài sản trí tuệ ban hành kèm Quyết định số 217/QĐ-ĐHTĐ "
     "ngày 21 tháng 11 năm 2024, đều được xây dựng trước khi Luật số 93/2025/QH15 có hiệu lực. Điểm a khoản 4 Điều 36 "
     "Quyết định 213 áp dụng cho sản phẩm đề tài sử dụng ngân sách nhà nước do Trường chủ trì, đã nghiệm thu và được "
     "thương mại hóa: nguồn thu sau khi trừ các khoản chi phí cần thiết, hợp lệ được chia 40% nộp ngân sách nhà nước, 30% "
     "cho Trường và 30% khen thưởng tập thể tác giả, tối đa 100 triệu đồng một đề tài. Đối chiếu với Điều 28 và Điều 73 "
     "Luật số 93/2025/QH15, quy định này không mâu thuẫn với Luật trong mọi trường hợp, nhưng cần được rà soát ở ba điểm. "
     "Thứ nhất, mức trần có thể làm phần thưởng thấp hơn mức tối thiểu 30% khi kết quả thuộc phạm vi áp dụng Điều 28 và "
     "30% lợi nhuận sau thuế vượt 100 triệu đồng. Thứ hai, cơ sở tính của Quy chế là nguồn thu sau chi phí, còn cơ sở "
     "tính của Luật là lợi nhuận sau thuế. Thứ ba, khoản nộp ngân sách nhà nước không thuộc các mục đích sử dụng lợi "
     "nhuận liệt kê tại khoản 3 Điều 28. Với nhiệm vụ được phê duyệt trước ngày 01 tháng 10 năm 2025 và không thuộc "
     "khoản 7 Điều 73, việc phân chia tiếp tục theo pháp luật và văn bản có hiệu lực tại thời điểm phê duyệt. Đây là căn "
     "cứ để Chương 2 xác định phạm vi áp dụng của từng quy định nội bộ và Chương 3 đề xuất hoàn thiện quy chế."),
    ("3. Xác định công thức phân chia cụ thể tỷ lệ doanh thu thương mại hóa",
     "3. Xác định cơ chế phân chia lợi ích theo từng nguồn hình thành tài sản và từng loại đối tượng, phân biệt thưởng, "
     "thù lao, nhuận bút với phần chia cho nhóm nghiên cứu, đơn vị quản lý trực tiếp và quỹ phát triển khoa học công "
     "nghệ của nhà trường, phù hợp với luật mới."),
    ("3. Thực hiện: Tổ chức vận hành các quy trình nghiệp vụ",
     "3. Thực hiện: Tổ chức vận hành các quy trình nghiệp vụ tiếp nhận đề xuất, thẩm định tính mới, hỗ trợ nộp đơn đăng "
     "ký, tổ chức đàm phán hợp đồng chuyển giao công nghệ và chi trả thưởng, thù lao, nhuận bút cho tác giả."),
    ("3. Yếu tố nguồn lực tài chính:",
     "3. Yếu tố nguồn lực tài chính: Nguồn kinh phí bảo đảm cho việc nộp đơn, duy trì hiệu lực văn bằng và kinh phí thử "
     "nghiệm hoàn thiện công nghệ. Khi các khoản này chưa được dự toán và bố trí sẵn, kết quả nghiên cứu dễ dừng lại ở "
     "dạng bài báo thay vì được đăng ký bảo hộ."),
    ("Thứ ba, cập nhật hệ thống pháp luật quốc gia hiện hành,",
     "Thứ ba, cập nhật hệ thống pháp luật quốc gia hiện hành, làm rõ cơ chế tự động giao quyền sở hữu kết quả nghiên cứu "
     "tại Điều 25 Luật số 93/2025/QH15, quyền đăng ký theo điểm c khoản 1 Điều 86 Luật Sở hữu trí tuệ do Luật số "
     "131/2025/QH15 bổ sung; phân biệt thưởng cho tác giả theo Điều 28 Luật số 93/2025/QH15, với mức tối thiểu 30% lợi "
     "nhuận sau thuế đối với phần kết quả sử dụng ngân sách nhà nước, và thù lao theo Điều 135 Luật Sở hữu trí tuệ, vốn "
     "theo thỏa thuận, không có trần và chỉ áp dụng mức mặc định khi không có thỏa thuận; làm rõ quy định chuyển tiếp tại "
     "khoản 3 và khoản 7 Điều 73. Trên cơ sở đó, Chương 1 xác định rằng việc rà soát quy chế nội bộ phải căn cứ vào nguồn "
     "kinh phí, thời điểm giao nhiệm vụ và loại đối tượng của từng trường hợp, không thể áp một tỷ lệ chung cho mọi tài "
     "sản."),
]


SO_DO = {
    "[[HINH_1_1]]": ("Chu trình bốn giai đoạn quản lý quyền sở hữu trí tuệ trong trường đại học", "chu_trinh.png",
                     "Nguồn: Nhóm nghiên cứu xây dựng trên cơ sở Bradley et al., 2013 và Tổ chức Sở hữu trí tuệ thế giới, "
                     "2020."),
    "[[HINH_1_2]]": ("Sáu nhóm yếu tố ảnh hưởng đến hiệu quả quản lý quyền sở hữu trí tuệ", "yeu_to.png",
                     "Nguồn: Nhóm nghiên cứu tổng hợp."),
    "[[HINH_1_3]]": ("Khung phân tích của đề tài", "khung_phan_tich.png", "Nguồn: Nhóm nghiên cứu xây dựng."),
}


def noi_dung(v):
    import so_do
    tu = len(v.doc.paragraphs)
    md = chuan_bi()
    thu_muc = os.path.join(GOC, "Ban_cuoi", "so_do")
    if not os.path.exists(os.path.join(thu_muc, "chu_trinh.png")):
        so_do.ve_tat_ca()
    for dong in md.split("\n"):
        s = dong.strip()
        if not s or s == "---":
            continue
        if s.startswith("# "):
            v.doan("chuong1", "CHƯƠNG 1")
            v.doan("chuong2", "CƠ SỞ LÝ LUẬN VÀ PHÁP LÝ VỀ QUẢN LÝ QUYỀN SỞ HỮU TRÍ TUỆ")
            v.doan("chuong3", "TRONG CƠ SỞ GIÁO DỤC ĐẠI HỌC")
        elif s.startswith("## "):
            v.doan("h1", s[3:])
        elif s.startswith("### "):
            v.doan("h2", s[4:])
        elif s == "[[BANG_1_2]]":
            b = BANG_1_2
            v.bang(b["tieu_de"], b["cot"], b["dong"], b["nguon"], b["rong"], can=b["can"], tien_to="1")
        elif s in SO_DO:
            ten, anh, nguon = SO_DO[s]
            v.so_do(ten, os.path.join(thu_muc, anh), nguon, tien_to="1")
        else:
            assert not s.startswith("|") and not s.startswith("![") and not s.startswith("[["), s[:60]
            v.doan_md("than", s)
    sua_lan6(v.doc, tu)


def sua_lan6(doc, tu):
    """Lượt chỉnh sửa theo bản góp ý ba chương: tách thưởng, thù lao, nhuận bút, phân chia lợi nhuận; đối chiếu đầy
    đủ Điều 28, Điều 73 Luật số 93/2025/QH15 và Điều 135 Luật Sở hữu trí tuệ."""
    from khung import thay_doan
    for bat_dau, moi in SUA_LAN6:
        thay_doan(doc, bat_dau, moi, tu=tu)


def dung():
    v = VanBanChung()
    noi_dung(v)
    return luu(v, "Chuong_1_Co_so_ly_luan_va_phap_ly.docx",
               "Chương 1. Cơ sở lý luận và pháp lý về quản lý quyền sở hữu trí tuệ trong cơ sở giáo dục đại học", 1)


if __name__ == "__main__":
    dung()

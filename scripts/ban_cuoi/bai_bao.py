# -*- coding: utf-8 -*-
"""Bài báo khoa học đầu ra của đề tài, gửi Tạp chí Nghiên cứu Khoa học và Phát triển.

    python3 scripts/ban_cuoi/bai_bao.py

Đầu ra: Ban_cuoi/Bai_bao_phuong_an_du_lieu_noi_bo.docx (chỉ dùng khi Nhà trường đồng ý công bố số liệu)

Bài viết theo cấu trúc chuẩn JSRD: tiêu đề, tóm tắt tiếng Việt và tiếng Anh, từ khóa,
Đặt vấn đề, Tổng quan nghiên cứu, Phương pháp nghiên cứu, Kết quả nghiên cứu, Bàn luận,
Kết luận, Tài liệu tham khảo (APA 7) và Tờ khai minh bạch sử dụng AI. Số liệu lấy từ
scripts/chuong2 (đã kiểm tra với dữ liệu gốc); biểu đồ là biểu đồ Excel gốc nhúng trong Word.
"""
import os
import re
import sys

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import khung  # noqa: E402
import tai_lieu as TL  # noqa: E402

BG = khung.BG
B = BG.B
RA_DOCX = os.path.join(khung.THU_MUC_RA, "Bai_bao_phuong_an_du_lieu_noi_bo.docx")
SO_DO = os.path.join(khung.THU_MUC_RA, "so_do")

# số liệu dùng trong bài, đối chiếu với dữ liệu Chương 2
assert (len(B.dt_du_dk), len(B.DE_TAI), len(B.nop_don), len(B.du_dk_den_2024)) == (11, 38, 1, 8)
assert B.shcn == 9 and len(B.dt_duoc) == 10 and B.bb[0] == 23 and B.bb[4] == 161
assert B.pt(B.cagr(B.bb[0], B.bb[4], 4)) == "62,7%"

TIEU_DE = "Từ đề tài đến văn bằng: Chuỗi chuyển hóa tài sản trí tuệ tại Trường Đại học Thành Đô"
TAC_GIA = "Nguyễn Thị Tố Uyên, Trần Đăng Bộ"
DON_VI = "Trường Đại học Thành Đô"

TOM_TAT = (
    "Bài báo đánh giá quản lý quyền sở hữu trí tuệ tại Trường Đại học Thành Đô trước khung pháp lý mới giai đoạn "
    "2025 - 2026, nhằm xác định điểm nghẽn trong chuyển hóa kết quả nghiên cứu thành tài sản trí tuệ được bảo hộ. "
    "Tiếp cận theo chu trình tạo lập, xác lập, bảo vệ và khai thác, nghiên cứu trường hợp sử dụng toàn bộ dữ liệu hành "
    "chính giai đoạn 2021 - 2025 của Nhà trường, gồm danh mục sản phẩm khoa học, hồ sơ đề tài cấp cơ sở, danh mục tài "
    "sản trí tuệ và quy chế nội bộ, kết hợp đối chiếu văn bản với luật mới. Kết quả cho thấy tiềm năng tài sản trí tuệ "
    "từ đề tài cấp cơ sở là đáng kể, tập trung ở khối ngành Y - Dược, nhưng tỷ lệ nộp đơn còn thấp. Điểm nghẽn quyết "
    "định nằm ở khâu nối giữa nghiệm thu và đăng ký, nơi chưa có biểu mẫu rà soát khả năng bảo hộ; cấu trúc khuyến khích "
    "ưu tiên công bố, thiếu kinh phí nộp đơn và độ trễ thể chế khuếch đại điểm nghẽn này. Hợp tác doanh nghiệp là kênh "
    "hình thành tài sản hiệu quả nhất. Nghiên cứu bổ sung bằng chứng về vai trò của khâu sàng lọc tại nghiệm thu và đề "
    "xuất phiếu rà soát bắt buộc cùng quy chế hợp nhất theo luật mới.")

TU_KHOA = "Quản lý quyền sở hữu trí tuệ; Tài sản trí tuệ; Chuyển giao công nghệ; Nghiệm thu đề tài; Trường đại học tư thục"

ABSTRACT = (
    "This article evaluates the management of intellectual property rights at Thanh Do University amid the substantial "
    "revision of Vietnam's science, technology and intellectual property legislation in 2025 - 2026, with the aim of "
    "identifying where the conversion of research results into protected intellectual assets breaks down. Adopting a "
    "life-cycle perspective covering creation, registration, protection and exploitation, the study uses a single-case "
    "design based on the university's complete administrative records for 2021 - 2025, including research output "
    "inventories, institutional research project files, the intellectual property portfolio, the staff register and "
    "internal regulations, complemented by a document analysis comparing internal rules with the new laws. The results "
    "show a considerable pool of protectable outputs from institutional research projects, concentrated in medicine and "
    "pharmacy, alongside a low rate of conversion into filed applications. The decisive bottleneck lies neither in the "
    "scope of internal regulations nor in research capacity, but in the link between project acceptance and filing, "
    "where no protectability screening form exists; a publication-oriented incentive structure, the absence of a "
    "budget line for filing fees and an institutional lag behind the new laws amplify this bottleneck. Collaboration "
    "with industry is the university's most productive channel for building intellectual assets. The study adds "
    "empirical evidence on the role of screening at project acceptance in an application-oriented private university "
    "and proposes a mandatory screening form and a consolidated regulation aligned with the Law on Science, Technology "
    "and Innovation as two priority interventions.")

KEYWORDS = ("Intellectual property management; Intellectual assets; Technology transfer; Research project acceptance; "
            "Private university")

DAT_VAN_DE = [
    "Tài sản trí tuệ ngày càng được coi là thước đo năng lực đổi mới và nguồn lực tài chính ngoài học phí của cơ sở giáo "
    "dục đại học. Tại Việt Nam, giai đoạn 2025 - 2026 chứng kiến sự thay đổi dồn dập của khung pháp lý: Luật Khoa học, "
    "công nghệ và đổi mới sáng tạo số 93/2025/QH15 trao cho tổ chức chủ trì quyền sở hữu kết quả nghiên cứu và quyền tự "
    "quyết về thương mại hóa; Luật số 131/2025/QH15 sửa đổi Luật Sở hữu trí tuệ; Luật Giáo dục đại học số "
    "125/2025/QH15 bổ sung nghĩa vụ công khai thông tin về sở hữu trí tuệ; Kết luận số 51-KL/TW của Bộ Chính trị xác "
    "định sở hữu trí tuệ là yếu tố cốt lõi của tự chủ chiến lược.",
    "Nghiên cứu quốc tế đã chỉ ra nhiều yếu tố quyết định hiệu quả thương mại hóa tài sản trí tuệ đại học, từ thể chế "
    "phân bổ quyền đến năng lực đơn vị chuyển giao công nghệ (Siegel et al., 2007; Thursby & Kemp, 2002). Các nghiên "
    "cứu trong nước giai đoạn gần đây chủ yếu phân tích quy định và cơ hội, thách thức ở cấp hệ thống (Nguyễn, 2025; "
    "Võ, 2025). Khoảng trống còn lại là thiếu bằng chứng định lượng ở cấp trường, theo chuỗi từ đề tài đến văn bằng, "
    "cho thấy chính xác mắt xích nào làm tiềm năng không thành tài sản, đặc biệt ở trường đại học tư thục định hướng "
    "ứng dụng.",
    "Bài báo sử dụng trường hợp Trường Đại học Thành Đô để trả lời ba câu hỏi: tiềm năng tài sản trí tuệ từ hoạt động "
    "nghiên cứu của Nhà trường lớn đến đâu và đã được chuyển hóa đến mức nào; điểm nghẽn nào quyết định khoảng cách "
    "giữa tiềm năng và kết quả; và quy chế nội bộ cần điều chỉnh như thế nào trước khung pháp lý mới.",
]

TONG_QUAN = [
    "**Trường đại học như một chủ thể tạo lập tài sản trí tuệ.** Etzkowitz (2003) cho rằng nhóm nghiên cứu trong trường "
    "đại học vận hành như những đơn vị gần với doanh nghiệp, tạo nền tảng cho mô hình đại học khởi nghiệp. Shane (2004) "
    "cho thấy việc trao quyền sở hữu kết quả nghiên cứu do nhà nước tài trợ cho trường đại học theo Đạo luật Bayh-Dole "
    "làm tăng số sáng chế của trường đại học, chủ yếu ở những lĩnh vực mà cấp phép là kênh chuyển giao hiệu quả. Hai "
    "công trình này đặt cơ sở cho luận điểm rằng quyền sở hữu và cơ chế chia lợi ích định hình hành vi đăng ký của "
    "trường và nhà khoa học. Dù vậy, cả hai dựa trên bối cảnh Hoa Kỳ và các trường đại học nghiên cứu, nơi đơn vị chuyển "
    "giao công nghệ đã chuyên nghiệp hóa.",
    "**Thể chế, khuyến khích và năng lực chuyển giao.** Goldfarb và Henrekson (2003) so sánh chính sách từ trên xuống "
    "của Thụy Điển với cơ chế từ dưới lên của Hoa Kỳ và kết luận rằng khuyến khích đối với nhà khoa học và trường đại "
    "học quan trọng hơn việc nhà nước can thiệp trực tiếp. Thursby và Kemp (2002) cho thấy tăng trưởng cấp phép của "
    "trường đại học Hoa Kỳ gắn với mức độ sẵn sàng khai báo sáng chế của giảng viên và hiệu quả khác nhau đáng kể giữa "
    "các trường. Siegel và cộng sự (2007) tổng hợp rằng hiệu quả của đơn vị chuyển giao công nghệ phụ thuộc vào chính "
    "sách khuyến khích, năng lực nhân sự và cơ chế chia lợi ích. Bradley và cộng sự (2013) phê phán mô hình tuyến tính "
    "của chuyển giao công nghệ và nhấn mạnh tính đa kênh của quá trình này. Điểm chung của nhóm nghiên cứu này là coi "
    "khâu khai báo sáng chế như đã có sẵn và tập trung vào khâu sau đó, trong khi ở trường chưa có đơn vị chuyển giao "
    "chuyên nghiệp, chính khâu nhận diện mới là nơi tiềm năng bị bỏ sót.",
    "**Quy trình quản lý tài sản trí tuệ.** Các tổng quan gần đây dịch chuyển trọng tâm từ chiếm hữu sang khai thác tài "
    "sản trí tuệ và làm rõ các mô hình, quy trình quản lý chuyển giao trong trường đại học (Holgersson & Aaboen, 2019; "
    "Maresova et al., 2019). Rocha và cộng sự (2023) mô tả cách các đơn vị chuyển giao tại Bồ Đào Nha đánh giá bản khai "
    "báo sáng chế theo tính mới, khả năng áp dụng công nghiệp và trình độ sáng tạo trước khi quyết định bảo hộ, qua đó "
    "cho thấy vai trò của một bước sàng lọc có cấu trúc. Teece (2018) nhắc lại rằng khả năng thu lợi từ đổi mới phụ "
    "thuộc vào chế độ bảo hộ và tài sản bổ trợ, nên việc bảo hộ đúng lúc là điều kiện của khai thác. Tuy nhiên, các "
    "nghiên cứu này chủ yếu mô tả quy trình ở nơi đã có đơn vị chuyên trách, chưa lượng hóa hệ quả của việc thiếu bước "
    "sàng lọc tại trường chưa có đơn vị đó.",
    "**Hợp tác đại học với doanh nghiệp.** Perkmann và cộng sự (2013) cho thấy gắn kết học thuật với doanh nghiệp phổ "
    "biến hơn nhiều so với thương mại hóa theo nghĩa hẹp và phụ thuộc vào đặc điểm cá nhân nhà khoa học. O’Dwyer và "
    "cộng sự (2023) chỉ ra rằng niềm tin, sự tương thích mục tiêu và cơ chế phân định quyền là các yếu tố thúc đẩy hợp "
    "tác thành công. Các phát hiện này gợi ý rằng ở trường có năng lực đăng ký còn hạn chế, hợp tác doanh nghiệp có thể "
    "là kênh bù đắp để hình thành tài sản trí tuệ.",
    "**Nghiên cứu trong nước và khoảng trống.** Nguyễn (2025) phân tích cơ hội và thách thức của quản trị tài sản trí "
    "tuệ tại cơ sở giáo dục đại học trong bối cảnh cách mạng công nghiệp lần thứ tư; Võ (2025) bàn về quyền của chủ thể "
    "không giữ quyền tài sản đối với tác phẩm hình thành trong nhà trường và kinh nghiệm quốc tế để hoàn thiện chính "
    "sách. Các công trình này có giá trị định hướng chính sách nhưng dựa trên phân tích văn bản ở cấp hệ thống, chưa sử "
    "dụng dữ liệu hành chính của một trường để truy vết từng đề tài đến văn bằng, và ra đời trước khi Luật số "
    "93/2025/QH15 có hiệu lực. Bài báo lấp khoảng trống này bằng cách đo chuỗi chuyển hóa tại một trường đại học tư "
    "thục và đối chiếu quy chế nội bộ với luật mới.",
]

PHUONG_PHAP = [
    "Nghiên cứu sử dụng thiết kế nghiên cứu trường hợp đơn, lựa chọn Trường Đại học Thành Đô vì đây là trường đại học tư "
    "thục định hướng ứng dụng, có quy chế sở hữu trí tuệ từ năm 2021 và quy chế chuyên biệt từ năm 2024, đồng thời có dữ "
    "liệu hành chính đủ để truy vết từ đề tài đến văn bằng. Khung phân tích dựa trên chu trình quản lý bốn giai đoạn tạo "
    "lập, xác lập, bảo vệ và khai thác, đặt trong ba điều kiện thể chế, tổ chức và nguồn lực.",
    "Dữ liệu là toàn bộ hồ sơ hành chính giai đoạn 2021 - 2025, không chọn mẫu, gồm: 582 bản ghi sản phẩm khoa học từ "
    "bảy danh mục thống kê của Phòng Khoa học Công nghệ; hồ sơ 38 đề tài cấp cơ sở với mã số, đơn vị chủ trì, kinh phí "
    "và sản phẩm nghiệm thu; danh mục 12 tài sản trí tuệ do bộ phận quản trị thương hiệu theo dõi; danh sách 252 nhân "
    "sự năm 2026; Kế hoạch số 07/KH-ĐHTĐ; bốn văn bản nội bộ có quy định về sở hữu trí tuệ. Danh sách nhân sự chứa dữ "
    "liệu cá nhân nên chỉ được sử dụng ở dạng tổng hợp.",
    "Dữ liệu được xử lý qua bốn bước. Thứ nhất, các danh mục được làm sạch, chuẩn hóa tên đơn vị, và tên chủ nhiệm đề tài "
    "được so khớp với danh sách tác giả bài báo trong ba năm kể từ năm nghiệm thu để loại trừ bản ghi trùng; kết quả cho "
    "thấy khoảng 475 sản phẩm độc lập. Thứ hai, sản phẩm nghiệm thu của từng đề tài được phân loại theo đối tượng quyền "
    "có thể xác lập theo Luật Sở hữu trí tuệ hợp nhất (Văn phòng Quốc hội, 2026); đây là đánh giá sơ bộ của nhóm tác giả dựa trên mô tả trong hồ "
    "sơ. Thứ ba, chuỗi chuyển hóa được tính bằng tỷ lệ đề tài có sản phẩm đủ điều kiện và tỷ lệ đề tài đủ điều kiện đã "
    "nộp đơn; mức độ tập trung công bố được đo bằng hệ số Gini. Thứ tư, từng điều khoản về phân chia lợi ích trong quy "
    "chế nội bộ được đối chiếu với Điều 25, Điều 27 và Điều 28 Luật số 93/2025/QH15 và Điều 135 Luật Sở hữu trí tuệ, "
    "kèm mô phỏng phần lợi ích của tác giả theo khoản thu từ một hợp đồng chuyển giao.",
]

KQ_1 = [
    "Rà soát sản phẩm nghiệm thu của 38 đề tài cấp cơ sở cho thấy 11 đề tài, tức 28,9%, tạo ra sản phẩm cụ thể đủ điều "
    "kiện xác lập quyền ngoài báo cáo và bài báo. Tỷ lệ này đáng ghi nhận với một trường có phần lớn ngành đào tạo thuộc "
    "kinh tế, xã hội và ngôn ngữ. Tiềm năng tập trung rõ rệt: 10 trên 11 sản phẩm thuộc lĩnh vực dược, từ công thức cồn "
    "thuốc, quy trình bào chế viên ngậm đến rutin tinh khiết và sản phẩm cầm máu dạng màng; 9 sản phẩm thuộc nhóm sở hữu "
    "công nghiệp, trong đó 8 sản phẩm phù hợp với giải pháp hữu ích, loại hình không đòi hỏi trình độ sáng tạo như sáng "
    "chế và phù hợp với quy mô đề tài cấp cơ sở.",
    "Tuy vậy, Hình 1 cho thấy chuỗi chuyển hóa đứt ở bậc thứ hai. Chỉ 1 trên 11 đề tài đủ điều kiện, tức 9,1%, được "
    "nộp đơn, là đề tài chiết xuất lá Quế hoa năm 2025 với đơn sáng chế nộp ngay trong năm nghiệm thu. Nếu chỉ xét đề "
    "tài giao đến năm 2024, đã đủ thời gian để nộp đơn, có 8 đề tài đủ điều kiện nhưng chưa đề tài nào có đơn; chưa sản "
    "phẩm nào được nộp theo hình thức giải pháp hữu ích. Với sản phẩm đã bộc lộ công khai quá mười hai tháng, khả năng "
    "đăng ký sáng chế, giải pháp hữu ích theo Điều 60 Luật Sở hữu trí tuệ không còn, nên phần tiềm năng này khó khôi "
    "phục.",
    "Phân bố theo thời gian cho thấy tiềm năng không phải hiện tượng nhất thời: sản phẩm đủ điều kiện xuất hiện đều "
    "đặn từ năm 2022, với 3 đề tài năm 2022, 1 đề tài năm 2023, 4 đề tài năm 2024 và 3 đề tài năm 2025. Hai đề tài năm "
    "2025 đã ghi trong hồ sơ dự kiến đăng ký giải pháp hữu ích nhưng chưa có đơn; nếu còn trong thời hạn mười hai tháng "
    "kể từ ngày bộc lộ, đây là phần tiềm năng có thể thu hồi ngay. Ngoài nhóm sở hữu công nghiệp, 2 sản phẩm là bộ mẫu "
    "cây thuốc và bộ tiêu bản hiển vi, thuộc nhóm sưu tập dữ liệu và quyền tác giả, có giá trị sử dụng trực tiếp trong "
    "đào tạo. Như vậy, cơ cấu tiềm năng phù hợp với định hướng ứng dụng của Nhà trường và cho phép quản lý có trọng tâm "
    "vào một khối ngành, một loại hình bảo hộ.",
]

KQ_2 = [
    "Khoảng cách trên không bắt nguồn từ phạm vi quy chế. Điều 34 Quy chế ban hành kèm Quyết định số 213/QĐ-ĐHTĐ và Điều 3 "
    "Quy chế ban hành kèm Quyết định số 217/QĐ-ĐHTĐ đã liệt kê đầy đủ giải pháp hữu ích, sưu tập dữ liệu, công thức và quy "
    "trình, tức đúng những nhóm mà sản phẩm đề tài rơi vào. Quyết định 217 cũng đã phân công Phòng Khoa học Công nghệ "
    "nhận diện, theo dõi tài sản trí tuệ và giao bộ phận pháp chế thực hiện thủ tục xác lập quyền. Khoảng cách cũng "
    "không bắt nguồn từ năng lực nghiên cứu: số bài báo tăng bình quân 62,7% một năm, từ 23 bài năm 2021 lên 161 bài năm "
    "2025, và bài báo quốc tế đạt 283% chỉ tiêu của Kế hoạch số 07/KH-ĐHTĐ.",
    "Điểm nghẽn nằm ở một khoảng trống kỹ thuật: thiếu một biểu mẫu rà soát tại thời điểm nghiệm thu. Quy trình nghiệm "
    "thu đề tài trước đây chủ yếu tập trung đánh giá mức độ hoàn thành nhiệm vụ chuyên môn và bài báo công bố, chưa "
    "tích hợp tiêu chí sàng lọc và định hướng đăng ký bảo hộ quyền sở hữu trí tuệ. Việc khởi động thủ tục vì vậy phụ "
    "thuộc vào sự chủ động của từng tác giả theo Điều 35 Quyết định 213 và Điều 10 Quyết định 217. Dữ liệu củng cố nhận "
    "định này theo ba hướng. Một là, chức năng quản lý được phân cho bốn đầu mối theo thế mạnh nghiệp vụ, ba đầu mối "
    "thuộc khối Quản trị và Dịch vụ trong khi kết quả nghiên cứu phát sinh ở khối Đào tạo và Nghiên cứu, và giữa các "
    "đầu mối chưa có hồ sơ theo dõi dùng chung. Hai là, mức khai báo trong hồ sơ nghiệm thu thấp: 26 trên 38 đề tài khai "
    "tổng cộng 32 công bố, trong khi phép so khớp gắn được 107 bài báo với các đề tài, tức độ phủ khoảng 29,9%; nếu bài "
    "báo, loại sản phẩm dễ nhận diện nhất, còn khai thiếu thì công thức, quy trình càng ít được ghi nhận. Ba là, trong 16 "
    "tiêu chí đánh giá hiệu quả quản lý quyền sở hữu trí tuệ mà nhóm tác giả xây dựng, chỉ 5 tiêu chí tính được đầy đủ "
    "từ dữ liệu hiện có và nhóm tiêu chí kết quả không có tiêu chí nào tính được đầy đủ.",
    "Trường hợp đề tài lá Quế hoa là đối chứng tự nhiên: kênh chuyển hóa vận hành được khi chủ nhiệm đề tài chủ động và "
    "có kinh nghiệm, trong trường hợp này là người đồng thời chủ trì đề tài cấp quốc gia. Khi thiếu sự chủ động đó, "
    "không có bước nào trong quy trình tự động đưa sản phẩm đến bàn xét bảo hộ. Sự phụ thuộc vào cá nhân còn thể hiện "
    "ở mức độ tập trung công bố: tính trên toàn bộ nhân sự, 88 trên 247 người có họ tên đầy đủ, tức 35,6%, đứng tên ít "
    "nhất một bài báo trong năm năm; hệ số Gini về số bài là 0,829, và 10% người dẫn đầu chiếm 39,3% số lượt đứng tên. "
    "Khi lực lượng nghiên cứu nòng cốt còn mỏng, một quy trình chỉ vận hành nhờ sự chủ động của cá nhân sẽ khó mở rộng.",
    "Hình 2 tổng hợp chuỗi nguyên nhân từ khoảng trống kỹ thuật này đến kết quả sản phẩm đủ điều kiện chưa được đăng "
    "ký. Khoảng trống tại nghiệm thu cộng hưởng với việc hệ thống thống kê chưa có danh mục tài sản trí tuệ, nên việc "
    "nhận diện dựa hoàn toàn vào tác giả; khi tác giả đã nhận diện, việc khởi động thủ tục lại vướng ở chỗ chưa có dòng "
    "kinh phí nộp đơn; trong lúc đó kết quả được công bố trước và sau mười hai tháng thì mất tính mới. Mỗi mắt xích đều "
    "có thể can thiệp bằng một công cụ quản lý cụ thể, và mắt xích tại nghiệm thu có chi phí can thiệp thấp nhất.",
]

KQ_3 = [
    "Hình 3 đặt định mức khuyến khích dành cho văn bằng cạnh định mức dành cho công bố theo Quy chế chi tiêu nội bộ "
    "năm 2026. Xét riêng giờ nghiên cứu quy đổi, văn bằng được định giá cao: sáng chế chuẩn Việt Nam được tính 360 giờ, "
    "cao hơn bài báo WoS hạng Q1 với 300 giờ. Nhưng chỉ bài báo quốc tế được thưởng tiền, từ 10 đến 20 triệu đồng một "
    "bài, và được ghi nhận ngay khi đăng; văn bằng không có tiền thưởng và chỉ được ghi nhận khi được cấp, thường từ hai "
    "đến ba năm sau khi nộp đơn. Với cùng một kết quả, lựa chọn hợp lý của giảng viên là công bố trước, và việc công bố "
    "trước khi nộp đơn làm mất tính mới sau mười hai tháng.",
    "Về tài chính, đề tài cấp cơ sở là kênh duy nhất đến nay tạo ra sản phẩm đủ điều kiện xác lập quyền. Cả 11 đề tài đủ "
    "điều kiện đều thuộc nhóm 19 đề tài được cấp tiền, sử dụng 396,75 triệu đồng, tức 93,4% tổng kinh phí 424,75 triệu "
    "đồng mà Nhà trường cấp trong năm năm, với trung vị 8,5 triệu đồng một đề tài. Kinh phí bằng tiền vì vậy là điều "
    "kiện cần để tạo ra sản phẩm có thể bảo hộ, nhưng dự toán đề tài không có dòng chi cho phí nộp đơn, phí duy trì "
    "hiệu lực. Nhà trường có Quỹ Học bổng sau tiến sĩ Ngô Xuân Độ với ngân sách 5 tỷ đồng cho giai đoạn 2025 - 2029, có "
    "khoản chi đủ trang trải chi phí đăng ký, nhưng Quỹ chỉ dành cho nhà khoa học theo diện học bổng sau tiến sĩ; giữa "
    "hai kênh chưa có cơ chế chuyển tiếp.",
]

KQ_4 = [
    "Nhà trường đã có tầm nhìn thể chế sớm khi ban hành quy chế có chương riêng về sở hữu trí tuệ từ năm 2021 và quy "
    "chế quản trị tài sản trí tuệ chuyên biệt từ năm 2024, trước khi luật mới ra đời. Các văn bản này được xây dựng ở "
    "những thời điểm khác nhau, cho những kênh tài trợ khác nhau, nên hiện cùng tồn tại năm quy định về phân chia lợi ích, "
    "được tóm tắt tại Bảng 1. Đây là biểu hiện của độ trễ thể chế trước sự thay đổi dồn dập của pháp luật quốc gia giai "
    "đoạn 2025 - 2026, không phải khiếm khuyết trong tư duy quản trị.",
    "Hình 4 mô phỏng phần lợi ích của tác giả từ một hợp đồng chuyển giao. Nhìn chung, các tỷ lệ nội bộ cao hơn mức mặc "
    "định 15% tại điểm b khoản 1 Điều 135 Luật Sở hữu trí tuệ. Tuy nhiên, Luật số 93/2025/QH15, có hiệu lực từ ngày 01 "
    "tháng 10 năm 2025, quy định tổ chức chủ trì được tự động giao quyền sở hữu phần kết quả sử dụng ngân sách nhà nước "
    "theo khoản 2 Điều 25, được tự quyết định phương án thương mại hóa theo Điều 27, và phải thưởng cho tác giả tối "
    "thiểu 30% lợi nhuận theo điểm a khoản 3 Điều 28. Công thức tại điểm a khoản 4 Điều 36 Quyết định 213, xây dựng "
    "theo cơ chế của Luật Khoa học và công nghệ năm 2013 với khoản nộp ngân sách nhà nước và mức trần 100 triệu đồng, "
    "cho phần của tác giả thấp hơn mức sàn mới khi khoản thu vượt khoảng 333 triệu đồng. Quy định này sẽ áp dụng trực "
    "tiếp cho kết quả của ba đề tài cấp quốc gia với tổng kinh phí 4,67 tỷ đồng mà Nhà trường đang chủ trì. Việc hợp "
    "nhất quy chế vì vậy là cơ hội để Nhà trường đi tiên phong đón đầu luật mới; trong các văn bản hiện hành, Điều lệ "
    "Quỹ là văn bản gần nhất với tinh thần của Luật số 93/2025/QH15 khi không đặt trần và cho phép các bên thỏa thuận.",
]

KQ_5 = [
    "Nhà trường hiện sở hữu 12 tài sản trí tuệ đã được xác lập quyền hoặc đang xử lý đơn, hình thành từ ba luồng. Luồng "
    "thương hiệu gồm 4 tài sản do Nhà trường đơn sở hữu và đều đã có văn bằng. Luồng hợp tác doanh nghiệp gồm 5 kiểu "
    "dáng công nghiệp và 1 nhãn hiệu đồng sở hữu với một doanh nghiệp đối tác, là thành công nổi bật trong chiến lược "
    "hợp tác đại học với doanh nghiệp mà Ban Giám hiệu đã dày công kết nối; riêng năm 2024 có 5 kiểu dáng được cấp, nâng "
    "số tài sản lũy kế từ 4 lên 9. Luồng nghiên cứu mới có 2 đơn sáng chế nộp năm 2025 và 2026.",
    "Như vậy, kênh hợp tác doanh nghiệp đã tạo ra một nửa số tài sản trí tuệ trong thời gian ngắn, trong khi kênh đề tài "
    "cấp cơ sở, nơi có nhiều sản phẩm đủ điều kiện nhất, mới đóng góp một đơn. Khai thác có thu phí mới có hai hợp đồng "
    "chuyển giao quyền sử dụng tác phẩm đối với sách chuyên khảo; chín văn bằng đã được cấp chưa phát sinh giao dịch "
    "chuyển giao quyền. Bước tiếp theo là tận dụng đà hợp tác để phát triển thêm tài sản do Nhà trường đơn sở hữu từ kết "
    "quả nghiên cứu, nhất là công thức và quy trình, những đối tượng mà kiểu dáng công nghiệp chưa bảo hộ.",
    "Về bảo vệ quyền, giai đoạn nghiên cứu không ghi nhận tranh chấp hay xử lý xâm phạm liên quan đến Nhà trường, phù hợp "
    "với quy mô tài sản và mức khai thác hiện tại. Do 5 kiểu dáng được cấp cùng năm 2024, việc gia hạn sẽ dồn vào cùng "
    "một thời điểm; thời hạn văn bằng hiện được theo dõi trong bảng riêng của bộ phận quản trị thương hiệu, chưa liên "
    "kết với danh mục đề tài của Phòng Khoa học Công nghệ. Sự tách rời giữa hai danh mục là thêm một biểu hiện của khoảng "
    "trống dữ liệu đã nêu ở Mục 4.2.",
]

BAN_LUAN = [
    "**Giải thích kết quả.** Kết quả cho thấy ở một trường chưa có đơn vị chuyển giao chuyên nghiệp, tiềm năng tài sản "
    "trí tuệ không mất đi ở khâu tạo lập mà ở khâu nhận diện. Quy chế đã liệt kê đúng đối tượng, nhưng không có thời "
    "điểm bắt buộc nào để đối chiếu sản phẩm với danh mục đó. Nghiệm thu là thời điểm duy nhất mọi sản phẩm đề tài đều "
    "đi qua một hội đồng, nên cũng là điểm can thiệp có chi phí thấp nhất. Cấu trúc khuyến khích ưu tiên công bố và việc "
    "thiếu dòng kinh phí nộp đơn làm cho khoảng trống này không được bù đắp bằng sự chủ động của tác giả, ngoại trừ các "
    "trường hợp tác giả có kinh nghiệm như đề tài lá Quế hoa.",
    "**Đối chiếu với nghiên cứu trước.** Phát hiện về vai trò của khuyến khích phù hợp với Goldfarb và Henrekson (2003) "
    "và Siegel và cộng sự (2007): khi khuyến khích đặt vào công bố, hành vi đăng ký giảm. Phát hiện về khâu khai báo "
    "tương đồng với nhận định của Thursby và Kemp (2002) rằng mức độ sẵn sàng khai báo của giảng viên quyết định quy mô "
    "đầu vào của chuyển giao. Điểm khác biệt là các nghiên cứu này mặc định có đơn vị chuyển giao tiếp nhận bản khai "
    "báo, như mô tả của Rocha và cộng sự (2023); tại Trường Đại học Thành Đô, chính bước tiếp nhận có cấu trúc đó chưa "
    "tồn tại, nên nút thắt dịch chuyển sớm hơn trong chuỗi. Phát hiện về kênh hợp tác doanh nghiệp phù hợp với Perkmann "
    "và cộng sự (2013) và O’Dwyer và cộng sự (2023): quan hệ tin cậy với đối tác có thể tạo ra tài sản nhanh hơn kênh "
    "nội sinh khi năng lực đăng ký còn hạn chế. So với các nghiên cứu trong nước (Nguyễn, 2025; Võ, 2025), bài báo chuyển "
    "trọng tâm từ phân tích quy định sang đo lường thực thi, và cho thấy quy định đầy đủ chưa đủ để tạo ra kết quả.",
    "**Đóng góp của bài báo.** Về lý luận, bài báo bổ sung bằng chứng rằng trong mô hình chu trình quản lý tài sản trí tuệ, "
    "khâu nối giữa tạo lập và xác lập cần được xem là một khâu quản lý độc lập, đặc biệt ở trường chưa có đơn vị chuyển "
    "giao chuyên nghiệp. Về phương pháp, bài báo cho thấy dữ liệu hành chính sẵn có của trường đại học đủ để truy vết "
    "chuỗi từ đề tài đến văn bằng nếu được so khớp có hệ thống. Về thực tiễn và chính sách, kết quả gợi ý bốn can thiệp "
    "neo vào từng phát hiện: phiếu rà soát khả năng bảo hộ bắt buộc tại nghiệm thu, cùng việc rà soát lại tám đề tài giai "
    "đoạn 2021 - 2024, xuất phát từ điểm đứt ở bậc thứ hai của chuỗi; dòng kinh phí nộp đơn và cơ chế chuyển tiếp với "
    "Quỹ Học bổng sau tiến sĩ Ngô Xuân Độ, xuất phát từ khoảng trống tài chính; quy chế hợp nhất theo Luật số "
    "93/2025/QH15, bỏ mức trần và bảo đảm mức tối thiểu cho tác giả, xuất phát từ độ trễ thể chế; và danh mục tài sản "
    "trí tuệ dùng chung giữa các đầu mối, xuất phát từ khoảng trống dữ liệu.",
    "**Hạn chế của nghiên cứu.** Thứ nhất, đây là nghiên cứu trường hợp đơn nên kết quả không đại diện cho toàn bộ hệ "
    "thống giáo dục đại học tư thục. Thứ hai, việc phân loại sản phẩm theo đối tượng quyền là đánh giá sơ bộ từ mô tả "
    "trong hồ sơ nghiệm thu, chưa phải kết quả tra cứu tính mới; số sản phẩm thực sự đủ điều kiện bảo hộ có thể thấp hơn. "
    "Thứ ba, nghiên cứu dựa trên dữ liệu hành chính, chưa kết hợp khảo sát hay phỏng vấn giảng viên, nên chưa kiểm chứng "
    "trực tiếp động cơ lựa chọn giữa công bố và đăng ký. Thứ tư, dữ liệu về khai thác và giá trị kinh tế của tài sản trí "
    "tuệ còn mỏng, nên nghiên cứu chưa đánh giá được hiệu quả kinh tế của quản lý quyền.",
    "**Hướng nghiên cứu tiếp theo.** Các nghiên cứu tiếp theo có thể đánh giá tác động của phiếu rà soát sau một đến hai "
    "năm áp dụng bằng thiết kế trước và sau, mở rộng so sánh sang một số trường đại học tư thục khác, và khảo sát giảng "
    "viên để lượng hóa ảnh hưởng của cấu trúc khuyến khích đến quyết định công bố hay đăng ký.",
]

KET_LUAN = [
    "Bài báo đánh giá hiệu quả quản lý quyền sở hữu trí tuệ tại Trường Đại học Thành Đô nhằm xác định điểm nghẽn trong "
    "chuyển hóa kết quả nghiên cứu thành tài sản trí tuệ được bảo hộ trước khung pháp lý mới.",
    "Phát hiện quan trọng nhất là tiềm năng tài sản trí tuệ từ đề tài cấp cơ sở của Nhà trường là đáng kể, tập trung ở "
    "khối ngành Y - Dược, nhưng chuỗi chuyển hóa đứt ở khâu nối giữa nghiệm thu và đăng ký do thiếu một biểu mẫu rà soát "
    "tại thời điểm nghiệm thu. Cấu trúc khuyến khích ưu tiên công bố, việc thiếu dòng kinh phí nộp đơn và độ trễ thể chế "
    "của các quy chế ban hành trước năm 2025 cùng khuếch đại điểm nghẽn này. Trong khi đó, chiến lược hợp tác với doanh "
    "nghiệp đã chứng tỏ là kênh hình thành tài sản hiệu quả nhất.",
    "Bài báo đóng góp bằng chứng thực nghiệm cho việc xem khâu nhận diện tài sản trí tuệ như một khâu quản lý độc lập "
    "trong chu trình, và cho thấy dữ liệu hành chính có thể dùng để đo chuỗi chuyển hóa ở cấp trường. Hàm ý trực tiếp "
    "cho Nhà trường là ban hành phiếu rà soát bắt buộc tại nghiệm thu và hợp nhất quy chế theo Luật số 93/2025/QH15; hàm "
    "ý cho các trường đại học tư thục có điều kiện tương tự là ưu tiên thiết lập bước sàng lọc có cấu trúc trước khi đầu "
    "tư vào các thiết chế thương mại hóa phức tạp hơn.",
]

BANG_1 = dict(
    tieu_de="Các quy định nội bộ về phân chia lợi ích từ tài sản trí tuệ và đối chiếu với Luật số 93/2025/QH15",
    cot=["Văn bản", "Phạm vi áp dụng", "Phần của tác giả", "Mức trần"],
    dong=[["Quyết định 213, điểm a khoản 4 Điều 36 (2021)", "Đề tài sử dụng ngân sách nhà nước",
           "30% khen thưởng tập thể tác giả, sau 40% nộp ngân sách", "100 triệu đồng"],
          ["Quyết định 213, điểm b khoản 4 Điều 36 (2021)", "Tài sản của Trường được chuyển giao", "30%",
           "Không đặt trần"],
          ["Quyết định 217, Điều 13 (2024)", "Khi các bên không có thỏa thuận", "Hiệu trưởng quyết định",
           "Không quy định"],
          ["Quy chế chi tiêu nội bộ (2026)", "Đề tài có đăng ký sở hữu trí tuệ", "Chưa quy định",
           "50 triệu đồng kinh phí đề tài"],
          ["Điều lệ Quỹ Ngô Xuân Độ, Điều 9 (2025)", "Sản phẩm từ kinh phí của Quỹ",
           "50% năm đầu; 20% từ năm thứ hai", "Không đặt trần"],
          ["**Luật số 93/2025/QH15, điểm a khoản 3 Điều 28**", "**Kết quả sử dụng ngân sách nhà nước**",
           "**Tối thiểu 30% lợi nhuận**", "**Không đặt trần**"]],
    nguon="Nguồn: Nhóm tác giả tổng hợp từ các văn bản nội bộ của Trường Đại học Thành Đô và Luật số 93/2025/QH15.",
    rong=[4.2, 3.8, 4.4, 2.6], can=["left", "left", "left", "left"],
)

TO_KHAI = [
    ("h", "TỜ KHAI MINH BẠCH SỬ DỤNG AI VÀ CAM KẾT DỮ LIỆU GỐC"),
    ("p", "Tên bài báo: " + TIEU_DE + "."),
    ("p", "Tác giả: " + TAC_GIA + ". Đơn vị: " + DON_VI + "."),
    ("b", "Phần I. Minh bạch sử dụng trí tuệ nhân tạo (chọn một phương án)"),
    ("p", "☐ Phương án A: Nhóm tác giả cam kết không sử dụng công cụ trí tuệ nhân tạo trong toàn bộ quá trình nghiên cứu, "
          "xử lý số liệu và viết bản thảo."),
    ("p", "☐ Phương án B: Nhóm tác giả xác nhận có sử dụng công cụ trí tuệ nhân tạo trong phạm vi cho phép, không dùng "
          "trí tuệ nhân tạo để tạo nội dung khoa học, tạo dữ liệu giả hoặc tài liệu tham khảo giả. Phạm vi sử dụng được kê "
          "khai tại bảng dưới đây."),
    ("tbl", None),
    ("b", "Phần II. Cam kết dữ liệu gốc"),
    ("p", "Nhóm tác giả cam kết dữ liệu trong bài là trung thực, chính xác, được thu thập từ quá trình nghiên cứu thực tế "
          "của nhóm tác giả tại Trường Đại học Thành Đô; cam kết lưu trữ và sẵn sàng cung cấp dữ liệu gốc khi Ban Biên tập "
          "yêu cầu; chịu trách nhiệm trước pháp luật và trước Hội đồng Biên tập về nội dung bản thảo."),
    ("p", "Hà Nội, ngày ... tháng ... năm 2026"),
    ("p", "Đại diện nhóm tác giả (ký, ghi rõ họ tên)"),
]
BANG_AI = dict(
    cot=["Tên công cụ", "Phạm vi hỗ trợ", "Nội dung hoặc mục có sự hỗ trợ"],
    dong=[["...", "...", "..."], ["...", "...", "..."]],
)


def tieu_de_muc(v, text):
    p = v.doan("h1", text, bold=True)
    return p


def dung():
    v = khung.VanBanChung()
    dem = {}

    def dem_tu(ten, ds):
        dem[ten] = dem.get(ten, 0) + sum(len(re.sub(r"\*", "", x).split()) for x in ds)

    p = v.doan("chuong1", TIEU_DE.upper(), bold=True)
    for r in p.runs:
        r.font.size = Pt(15)
    p.paragraph_format.space_after = Pt(8)
    p = v.doan("chuong2", TAC_GIA, bold=True)
    p = v.doan("chuong3", DON_VI, italic=True)
    p.paragraph_format.space_after = Pt(12)

    v.doan_md("than", "**Tóm tắt:** " + TOM_TAT)
    v.doan_md("than", "**Từ khóa:** " + TU_KHOA)
    v.doan_md("than", "**Abstract:** *" + ABSTRACT + "*")
    v.doan_md("than", "**Keywords:** *" + KEYWORDS + "*")

    tieu_de_muc(v, "1. ĐẶT VẤN ĐỀ")
    v.than_md(*DAT_VAN_DE)
    dem_tu("Đặt vấn đề", DAT_VAN_DE)

    tieu_de_muc(v, "2. TỔNG QUAN NGHIÊN CỨU")
    v.than_md(*TONG_QUAN)
    dem_tu("Tổng quan", TONG_QUAN)

    tieu_de_muc(v, "3. PHƯƠNG PHÁP NGHIÊN CỨU")
    v.than_md(*PHUONG_PHAP)
    dem_tu("Phương pháp", PHUONG_PHAP)

    tieu_de_muc(v, "4. KẾT QUẢ NGHIÊN CỨU")
    v.doan("h2", "4.1. Tiềm năng tài sản trí tuệ từ đề tài cấp cơ sở đáng kể nhưng tỷ lệ chuyển hóa thấp")
    v.than_md(*KQ_1)
    v.hinh_bd("H2.15", tien_to="", nguon="Nguồn: Nhóm tác giả tính toán từ danh mục đề tài cấp cơ sở giai đoạn 2021 - "
                                          "2025 của Trường Đại học Thành Đô.")
    v.doan("h2", "4.2. Điểm nghẽn nằm ở khâu nối giữa nghiệm thu và đăng ký")
    v.than_md(*KQ_2)
    v.so_do("Chuỗi nguyên nhân dẫn đến việc sản phẩm đủ điều kiện chưa được đăng ký bảo hộ",
            os.path.join(SO_DO, "chuoi_nguyen_nhan.png"), "Nguồn: Nhóm tác giả xây dựng.", tien_to="", rong_cm=15.0)
    v.doan("h2", "4.3. Cấu trúc khuyến khích và tài chính ưu tiên công bố hơn đăng ký")
    v.than_md(*KQ_3)
    v.hinh_bd("H2.11", tien_to="", nguon="Nguồn: Nhóm tác giả tổng hợp từ Bảng 6 và Bảng 7, Quy chế chi tiêu nội bộ "
                                          "Trường Đại học Thành Đô năm 2026. Mức thưởng là tổng tiền thưởng cho một "
                                          "bài báo.")
    v.doan("h2", "4.4. Độ trễ thể chế trước luật mới và cơ hội đón đầu")
    v.than_md(KQ_4[0])
    b = BANG_1
    v.bang(b["tieu_de"], b["cot"], b["dong"], b["nguon"], b["rong"], can=b["can"], hang_dam=(), tien_to="")
    v.than_md(*KQ_4[1:])
    v.hinh_bd("H2.8", tien_to="", nguon="Nguồn: Nhóm tác giả mô phỏng từ Điều 36 Quyết định 213, Điều 9 Điều lệ Quỹ Học "
                                         "bổng sau tiến sĩ Ngô Xuân Độ, điểm b khoản 1 Điều 135 Luật Sở hữu trí tuệ và "
                                         "điểm a khoản 3 Điều 28 Luật số 93/2025/QH15; đơn vị: triệu đồng.")
    v.doan("h2", "4.5. Hợp tác doanh nghiệp là kênh hình thành tài sản trí tuệ hiệu quả nhất")
    v.than_md(*KQ_5)
    dem_tu("Kết quả", KQ_1 + KQ_2 + KQ_3 + KQ_4 + KQ_5)

    tieu_de_muc(v, "5. BÀN LUẬN")
    v.than_md(*BAN_LUAN)
    dem_tu("Bàn luận", BAN_LUAN)

    tieu_de_muc(v, "6. KẾT LUẬN")
    v.than_md(*KET_LUAN)
    dem_tu("Kết luận", KET_LUAN)

    van_ban = "\n".join([p.text for p in v.doc.paragraphs] +
                        [c.text for t in v.doc.tables for r in t.rows for c in r.cells])
    tieu_de_muc(v, "TÀI LIỆU THAM KHẢO")
    dung_tl = TL.duoc_trich(van_ban)
    for apa in TL.danh_muc(dung_tl):
        p = v.doan_md("than", apa)
        pf = p.paragraph_format
        pf.left_indent, pf.first_line_indent, pf.space_after = Pt(28), Pt(-28), Pt(4)

    # tờ khai minh bạch sử dụng AI, trang riêng
    for loai, t in TO_KHAI:
        if loai == "h":
            p = v.doan("chuong1", t, bold=True)
            p.paragraph_format.page_break_before = True
        elif loai == "b":
            v.doan_md("than", f"**{t}**")
        elif loai == "tbl":
            v.bang("Kê khai phạm vi sử dụng công cụ trí tuệ nhân tạo (Phương án B)", BANG_AI["cot"], BANG_AI["dong"],
                   None, [3.5, 5.5, 6.0], can=["left", "left", "left"], tien_to="")
        else:
            v.doan_md("than", t)

    v.doc.core_properties.title = TIEU_DE
    v.doc.core_properties.author = TAC_GIA
    v.doc.save(RA_DOCX)

    # báo cáo dung lượng theo chuẩn JSRD
    tong = sum(dem.values())
    khung_tl = {"Đặt vấn đề": (5, 7), "Tổng quan": (15, 20), "Phương pháp": (6, 10), "Kết quả": (40, 55),
                "Bàn luận": (15, 25), "Kết luận": (5, 8)}
    print("Đã ghi:", RA_DOCX)
    print(f"Tiêu đề: {len(TIEU_DE.split())} từ; tóm tắt: {len(TOM_TAT.split())} từ; abstract: {len(ABSTRACT.split())} "
          f"từ; từ khóa: {len(TU_KHOA.split(';'))} cụm; tài liệu: {len(dung_tl)}")
    for k, (a, b_) in khung_tl.items():
        tl = 100 * dem[k] / tong
        print(f"  {k:12s} {dem[k]:5d} từ  {tl:5.1f}%  (khung {a}-{b_}%){'' if a <= tl <= b_ else '  NGOÀI KHUNG'}")
    print(f"  Tổng thân bài {tong} từ")
    mo_coi = TL.trich_dan_mo_coi(van_ban)
    if mo_coi:
        print("CẢNH BÁO trích dẫn không có trong danh mục:", mo_coi)
    for x in khung.kiem_tra(v.doc):
        print("CẢNH BÁO:", x)
    return dem


if __name__ == "__main__":
    dung()

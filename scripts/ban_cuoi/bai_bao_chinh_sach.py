# -*- coding: utf-8 -*-
"""Bài báo đầu ra, phương án phân tích chính sách: chỉ sử dụng văn bản công khai.

    python3 scripts/ban_cuoi/bai_bao_chinh_sach.py

Đầu ra: Ban_cuoi/Bai_bao_Tap_chi_NCKH_PT.docx

Phương án này không công bố số liệu nội bộ của Trường Đại học Thành Đô. Dữ liệu là
12 văn bản pháp luật, chỉ đạo và chuẩn chất lượng ban hành công khai; mọi điều, khoản
được viện dẫn đã đối chiếu với tệp gốc trong thư mục VBPL và tệp Thông tư số
83/2026/TT-BGDĐT ở thư mục gốc. Phương án sử dụng dữ liệu của Nhà trường vẫn được dựng
bởi bai_bao.py, chỉ dùng khi Nhà trường đồng ý công bố số liệu.
"""
import os
import re
import sys

from docx.shared import Pt

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import khung  # noqa: E402
import tai_lieu as TL  # noqa: E402

RA_DOCX = os.path.join(khung.THU_MUC_RA, "Bai_bao_Tap_chi_NCKH_PT.docx")
SO_DO = os.path.join(khung.THU_MUC_RA, "so_do")

TIEU_DE = ("Khoảng cách giữa sản phẩm khoa học và tài sản trí tuệ tại trường đại học tư thục: Phân tích chính sách "
           "trong bối cảnh pháp lý mới")
TAC_GIA = "Nguyễn Thị Tố Uyên, Trần Đăng Bộ"
DON_VI = "Trường Đại học Thành Đô"

TOM_TAT = (
    "Bài báo phân tích những thay đổi của khung pháp lý giai đoạn 2025 - 2026 tác động như thế nào đến khoảng cách giữa "
    "sản phẩm khoa học và tài sản trí tuệ tại trường đại học tư thục, và quy chế nội bộ cần điều chỉnh gì. Nghiên "
    "cứu sử dụng phân tích văn bản chính sách theo chu trình tạo lập, xác lập, bảo vệ, khai thác tài sản trí "
    "tuệ, với dữ liệu là 12 văn bản luật, nghị định, chỉ đạo và chuẩn chất lượng ban hành công khai, đối chiếu với các "
    "nghiên cứu quốc tế về chuyển giao công nghệ đại học. Kết quả cho thấy pháp luật mới đã tháo gỡ phần lớn rào cản ở "
    "các khâu sau xác lập quyền: quyền sở hữu và quyền quyết định thương mại hóa được trao cho tổ chức chủ trì, kể cả "
    "tổ chức ngoài công lập; cơ chế chia lợi ích chuyển sang mức sàn và thỏa thuận; quỹ phát triển khoa học và công "
    "nghệ được phép chi cho đăng ký bảo hộ; chuẩn chất lượng mới nâng trọng số của bằng độc quyền giải pháp hữu ích. "
    "Tuy nhiên, khâu nhận diện sản phẩm có khả năng bảo hộ trước khi công bố vẫn hoàn toàn thuộc về quy chế nội bộ. "
    "Bài báo đề xuất khung rà soát quy chế nội bộ gồm chín nội dung và mô hình cổng rà soát tại nghiệm thu đề tài.")

TU_KHOA = "Tài sản trí tuệ; Quy chế nội bộ; Luật Khoa học, công nghệ và đổi mới sáng tạo; Trường đại học tư thục; Phân tích chính sách"

ABSTRACT = (
    "This article examines how the 2025 - 2026 overhaul of Vietnam's legal framework affects the gap between research "
    "outputs and intellectual assets in private universities, and what internal regulations need to change in "
    "response. Using policy document analysis structured around the intellectual property life cycle of creation, "
    "registration, protection and exploitation, the study analyses twelve publicly issued laws, decrees, policy "
    "directives and quality standards, and relates them to international research on university technology transfer. "
    "The findings show that the new legislation removes most barriers in the stages following registration: ownership "
    "and the right to decide on commercialisation are vested in the host institution, including non-public "
    "institutions; benefit sharing moves to a minimum floor combined with negotiated terms; institutional science and "
    "technology development funds may finance filing and protection; and the new quality standard triples the weight "
    "of utility solution patents. However, the identification of protectable outputs before disclosure remains "
    "entirely a matter for internal regulation. The article proposes a nine-item framework for reviewing internal "
    "regulations and a screening gate model at the acceptance stage of research projects. These findings suggest that, "
    "once national barriers have been lowered, the decisive lever for narrowing the gap lies in institutional design, "
    "particularly in a structured screening step placed before research results are published, which private "
    "universities can introduce at low cost when consolidating their regulations under the new law.")

KEYWORDS = ("Intellectual assets; Internal regulations; Law on Science, Technology and Innovation; Private university; "
            "Policy analysis")

DAT_VAN_DE = [
    "Khoảng cách giữa sản phẩm khoa học và tài sản trí tuệ của trường đại học không chỉ là vấn đề năng lực nghiên cứu "
    "mà còn là vấn đề thể chế: ai sở hữu kết quả, ai được quyết định khai thác, tác giả được hưởng bao nhiêu và chi phí đăng ký lấy từ đâu. "
    "Giai đoạn 2025 - 2026, các câu hỏi này được trả lời lại gần như toàn bộ bởi Luật Khoa học, công nghệ và đổi mới "
    "sáng tạo số 93/2025/QH15, Luật số 131/2025/QH15 sửa đổi Luật Sở hữu trí tuệ, Luật Giáo dục đại học số "
    "125/2025/QH15 và Chuẩn cơ sở giáo dục đại học ban hành kèm Thông tư số 83/2026/TT-BGDĐT.",
    "Nghiên cứu quốc tế đã chỉ ra vai trò của thể chế phân bổ quyền và cơ chế khuyến khích đối với chuyển giao công nghệ "
    "đại học (Goldfarb & Henrekson, 2003; Shane, 2004). Nghiên cứu trong nước chủ yếu ra đời trước khi Luật số "
    "93/2025/QH15 có hiệu lực (Nguyễn, 2025; Võ, 2025), nên chưa đánh giá được khung pháp lý mới giải quyết khâu nào và "
    "để lại khâu nào cho cơ sở giáo dục đại học tự thiết kế, nhất là trường tư thục.",
    "Bài báo trả lời hai câu hỏi: những thay đổi của khung pháp lý giai đoạn 2025 - 2026 tác động như thế nào đến từng "
    "khâu trong chu trình hình thành tài sản trí tuệ của trường đại học tư thục; và quy chế nội bộ cần điều chỉnh những "
    "nội dung gì để thu hẹp khoảng cách giữa sản phẩm khoa học và tài sản trí tuệ.",
]

TONG_QUAN = [
    "**Thể chế phân bổ quyền.** Shane (2004) cho thấy việc trao quyền sở hữu kết quả nghiên cứu do nhà nước tài trợ cho "
    "trường đại học theo Đạo luật Bayh-Dole làm tăng số sáng chế của trường đại học Hoa Kỳ, chủ yếu ở những lĩnh vực mà "
    "cấp phép là kênh chuyển giao hiệu quả. Etzkowitz (2003) lý giải sự hình thành của đại học khởi nghiệp từ đặc tính "
    "gần với doanh nghiệp của các nhóm nghiên cứu. Hai công trình này là cơ sở cho nhận định rằng trao quyền cho tổ "
    "chức chủ trì là điều kiện cần của thương mại hóa. Tuy vậy, chúng dựa trên bối cảnh trường đại học nghiên cứu có "
    "đơn vị chuyển giao công nghệ chuyên nghiệp.",
    "**Khuyến khích và năng lực tổ chức.** Goldfarb và Henrekson (2003) so sánh chính sách từ trên xuống của Thụy Điển "
    "với cơ chế từ dưới lên của Hoa Kỳ và kết luận rằng khuyến khích đối với nhà khoa học và nhà trường quan trọng hơn "
    "can thiệp trực tiếp của nhà nước. Thursby và Kemp (2002) cho thấy tăng trưởng cấp phép của trường đại học gắn với "
    "mức độ sẵn sàng khai báo sáng chế của giảng viên. Siegel và cộng sự (2007) tổng hợp rằng hiệu quả của đơn vị "
    "chuyển giao phụ thuộc vào chính sách khuyến khích, năng lực nhân sự và cơ chế chia lợi ích. Bradley và cộng sự "
    "(2013) phê phán mô hình tuyến tính và nhấn mạnh tính đa kênh của chuyển giao công nghệ. Các nghiên cứu này thống "
    "nhất rằng quyền sở hữu chỉ phát huy tác dụng khi đi kèm khuyến khích và năng lực tổ chức tương xứng.",
    "**Quy trình quản lý tài sản trí tuệ.** Holgersson và Aaboen (2019) chỉ ra sự dịch chuyển trọng tâm từ chiếm hữu sang "
    "khai thác tài sản trí tuệ; Maresova và cộng sự (2019) hệ thống hóa các mô hình và quy trình quản lý chuyển giao "
    "trong trường đại học. Rocha và cộng sự (2023) mô tả cách đơn vị chuyển giao đánh giá bản khai báo sáng chế theo "
    "tính mới, khả năng áp dụng công nghiệp và trình độ sáng tạo trước khi quyết định bảo hộ. Teece (2018) nhắc lại "
    "rằng khả năng thu lợi từ đổi mới phụ thuộc vào chế độ bảo hộ và tài sản bổ trợ. Điểm chung là các nghiên cứu đều "
    "giả định đã có một bước tiếp nhận và sàng lọc có cấu trúc, điều chưa chắc có ở trường chưa có đơn vị chuyên trách.",
    "**Hợp tác với doanh nghiệp.** Perkmann và cộng sự (2013) cho thấy gắn kết học thuật với doanh nghiệp phổ biến hơn "
    "nhiều so với thương mại hóa theo nghĩa hẹp; O’Dwyer và cộng sự (2023) xác định niềm tin và cơ chế phân định quyền "
    "rõ ràng là các yếu tố thúc đẩy hợp tác thành công, qua đó cho thấy quy chế nội bộ là điều kiện của hợp tác.",
    "**Nghiên cứu trong nước và khoảng trống.** Nguyễn (2025) phân tích cơ hội và thách thức của quản trị tài sản trí tuệ "
    "tại cơ sở giáo dục đại học; Võ (2025) bàn về quyền của chủ thể không giữ quyền tài sản đối với tác phẩm hình thành "
    "trong nhà trường; tài liệu tập huấn của Cục Sở hữu trí tuệ (n.d.) cung cấp khung nghiệp vụ cho cán bộ quản lý. Các "
    "công trình này chưa phân tích hệ thống những thay đổi của Luật số 93/2025/QH15, Luật số 131/2025/QH15 và Thông tư "
    "số 83/2026/TT-BGDĐT theo từng khâu của chu trình, và chưa chỉ ra khâu nào pháp luật để lại cho quy chế nội bộ. Bài "
    "báo lấp khoảng trống này với trọng tâm là trường đại học tư thục.",
]

PHUONG_PHAP = [
    "Nghiên cứu sử dụng phương pháp phân tích văn bản chính sách định tính. Khung phân tích là chu trình quản lý tài sản "
    "trí tuệ bốn giai đoạn gồm tạo lập và nhận diện, xác lập quyền, khai thác và thương mại hóa, bảo vệ và phân chia lợi "
    "ích (Hình 1), được nhóm tác giả xây dựng trong đề tài khoa học công nghệ cấp cơ sở năm 2026 trên cơ sở Bradley và "
    "cộng sự (2013) và Tổ chức Sở hữu trí tuệ thế giới (2020).",
    "Dữ liệu là 12 văn bản ban hành công khai, được chọn theo hai tiêu chí: có hiệu lực hoặc đã được ban hành đến ngày "
    "30 tháng 9 năm 2026, và có quy định trực tiếp về quyền đối với kết quả nghiên cứu, phân chia lợi ích, nguồn lực hoặc "
    "chỉ số về sở hữu trí tuệ trong cơ sở giáo dục đại học. Nhóm luật gồm Luật Sở hữu trí tuệ hợp nhất (Văn phòng Quốc "
    "hội, 2026), Luật số 93/2025/QH15, Luật số 131/2025/QH15, Luật số 125/2025/QH15 và Luật số 07/2017/QH14; nhóm văn "
    "bản dưới luật và chỉ đạo gồm Nghị định số 134/2026/NĐ-CP, Quyết định số 1068/QĐ-TTg, Quyết định số 1624/QĐ-TTg, Chỉ "
    "thị số 02/CT-TTg và Kết luận số 51-KL/TW; nhóm chuẩn chất lượng gồm Thông tư số 01/2024/TT-BGDĐT và Thông tư số "
    "83/2026/TT-BGDĐT.",
    "Mỗi văn bản được mã hóa theo điều, khoản, điểm và gán vào khâu tương ứng của chu trình cùng ba điều kiện thể chế, "
    "tổ chức, nguồn lực. Với từng khâu, nhóm tác giả xác định quy định hiện hành, đối tượng được trao quyền hoặc nghĩa "
    "vụ, và nội dung pháp luật để lại cho cơ sở giáo dục đại học tự quy định. Hai thông tư về chuẩn cơ sở giáo dục đại "
    "học được so sánh trực tiếp theo công thức tính chỉ số sản phẩm khoa học. Kết quả phân tích được đối chiếu với các "
    "nghiên cứu quốc tế ở Mục 2. Bài báo không sử dụng số liệu nội bộ của cơ sở giáo dục đại học nào.",
]

KQ_1 = [
    "Theo khoản 2 Điều 25 Luật số 93/2025/QH15, tổ chức chủ trì nhiệm vụ khoa học, công nghệ và đổi mới sáng tạo được "
    "Nhà nước tự động giao quyền quản lý, sử dụng, quyền sở hữu phần kết quả tương ứng với kinh phí ngân sách nhà nước, "
    "không phải thực hiện thủ tục giao quyền và không phải bồi hoàn chi phí. Luật cũng nêu ba trường hợp loại trừ, đáng "
    "chú ý là nhiệm vụ do Nhà nước đặt hàng đã nêu rõ yêu cầu Nhà nước nắm giữ quyền để phục vụ phòng bệnh, chữa bệnh "
    "hoặc nhu cầu cấp thiết khác của xã hội. Khoản 1 Điều 25 xác định tổ chức, cá nhân đóng góp tài sản, tài chính là "
    "chủ sở hữu kết quả tương ứng với tỷ lệ đóng góp theo thỏa thuận. Luật số 131/2025/QH15 bổ sung điểm c khoản 1 Điều "
    "86 Luật Sở hữu trí tuệ, trao quyền đăng ký sáng chế, kiểu dáng công nghiệp, thiết kế bố trí cho tổ chức được giao "
    "quyền theo cơ chế này.",
    "Về khai thác, Điều 27 Luật số 93/2025/QH15 cho phép chủ sở hữu tự quyết định việc thương mại hóa; tổ chức được giao "
    "quyền tự quyết định hình thức, phương án, giá, phân chia lợi nhuận và phương án góp vốn. Điểm d khoản 2 Điều 28 "
    "Luật Giáo dục đại học số 125/2025/QH15 khẳng định cơ sở giáo dục đại học được định giá, xác lập quyền sở hữu, khai "
    "thác, góp vốn, phân chia lợi ích từ tài sản trí tuệ; khoản 1 Điều 28 cho phép thành lập doanh nghiệp quản lý tài "
    "sản trí tuệ. Đối với trường tư thục, điểm quan trọng nhất nằm ở Điều 37 Luật số 93/2025/QH15: tổ chức ngoài công lập "
    "được bảo đảm quyền tiếp cận bình đẳng, được giao quyền sở hữu hoặc quyền sử dụng kết quả nghiên cứu từ nhiệm vụ sử "
    "dụng ngân sách nhà nước do mình thực hiện theo điểm b khoản 2, và được hưởng ưu đãi như tổ chức công lập theo điểm "
    "d khoản 2.",
    "Như vậy, ở khâu xác lập và khai thác, khung pháp lý mới đã trả lời dứt khoát câu hỏi ai sở hữu và ai quyết định. "
    "Đổi lại, khoản 4 Điều 27 đặt cho tổ chức chủ trì trách nhiệm công khai, minh bạch thông tin và báo cáo kết quả "
    "thương mại hóa, tức đòi hỏi tổ chức có hệ thống theo dõi tài sản trí tuệ của chính mình. Luật Chuyển giao công "
    "nghệ số 07/2017/QH14 tiếp tục điều chỉnh các hình thức chuyển giao, định giá công nghệ và góp vốn bằng kết quả "
    "nghiên cứu, nên quy chế nội bộ cần dẫn chiếu luật này khi xây dựng mẫu hợp đồng chuyển giao.",
]

KQ_2 = [
    "Điều 28 Luật số 93/2025/QH15 phân biệt hai trường hợp. Với phần lợi nhuận tương ứng kết quả không sử dụng ngân sách "
    "nhà nước, chủ sở hữu tự quyết định việc xử lý lợi nhuận, kể cả thưởng cho tác giả, theo khoản 2. Với phần tương ứng "
    "kết quả sử dụng ngân sách nhà nước, điểm a khoản 3 quy định thưởng cho tác giả tối thiểu 30% lợi nhuận thu được từ "
    "thương mại hóa, hoặc tối thiểu 30% giá trị kết quả khi góp vốn. Đồng thời, điểm h khoản 7 Điều 71 bãi bỏ khoản 2 "
    "Điều 135 Luật Sở hữu trí tuệ; khoản 1 Điều 135, được sửa theo điểm b khoản 7 Điều 71, đặt nguyên tắc thỏa thuận và "
    "chỉ áp dụng mức mặc định 10% lợi nhuận trước thuế khi chủ sở hữu tự sử dụng hoặc 15% số tiền nhận được khi chuyển "
    "giao quyền sử dụng nếu các bên không có thỏa thuận. Bảng 1 tổng hợp các thay đổi này cùng hàm ý cho quy chế nội bộ.",
    "Điểm then chốt là pháp luật chuyển từ một khung định sẵn sang mức sàn kết hợp thỏa thuận: không đặt trần cho phần "
    "của tác giả và để nhà trường tự thiết kế công thức cho các nguồn kinh phí ngoài ngân sách. Các quy chế nội bộ ban "
    "hành trước ngày 01 tháng 10 năm 2025 có thể còn dẫn chiếu cơ chế cũ, chẳng hạn khoản nộp ngân sách hoặc mức trần "
    "tiền thưởng; đó là độ trễ thể chế tất yếu trước sự thay đổi dồn dập của pháp luật, nhưng cần được rà soát sớm vì "
    "mức sàn 30% là quy định bắt buộc.",
]

KQ_3 = [
    "Chi phí nộp đơn và duy trì hiệu lực văn bằng phát sinh trước khi có doanh thu, nên nếu chỉ được trừ khi chia lợi "
    "ích thì không có nguồn để khởi động thủ tục. Khung pháp lý mới đã mở hai căn cứ cho khoản chi này. Điểm b khoản 2 "
    "Điều 66 Luật số 93/2025/QH15 cho phép doanh nghiệp, tổ chức, đơn vị sự nghiệp sử dụng quỹ phát triển khoa học và "
    "công nghệ cho đăng ký, bảo hộ, quản lý, khai thác quyền sở hữu trí tuệ, với quyền tự chủ trong quản lý quỹ theo "
    "khoản 4. Điểm d khoản 3 Điều 28 Luật số 125/2025/QH15 giao cơ sở giáo dục đại học trách nhiệm thành lập và vận hành "
    "quỹ phát triển khoa học và công nghệ.",
    "Ở tầm chiến lược, Quyết định số 1624/QĐ-TTg sửa đổi Chiến lược sở hữu trí tuệ đến năm 2030 yêu cầu sử dụng các chỉ "
    "số đo lường về sở hữu trí tuệ làm căn cứ đánh giá hiệu quả hoạt động của cơ sở giáo dục đại học, xác định đối "
    "tượng quyền cần đạt được đối với kết quả nghiên cứu sử dụng ngân sách nhà nước, và với cơ sở khối kỹ thuật, công "
    "nghệ thì đăng ký bảo hộ đồng thời với công bố bài báo về kết quả có tính ứng dụng cao. Kết luận số 51-KL/TW xác định "
    "yêu cầu chuyển từ tư duy quản lý hành chính sang kiến tạo hệ sinh thái sở hữu trí tuệ. Chỉ thị số 02/CT-TTg về tăng "
    "cường thực thi quyền sở hữu trí tuệ giao Bộ Giáo dục và Đào tạo nghiên cứu đưa nội dung giáo dục về sở hữu trí tuệ "
    "vào các hệ, cấp học phù hợp, qua đó đặt thêm yêu cầu về năng lực sở hữu trí tuệ của giảng viên và người học. Các "
    "văn bản này tạo căn cứ "
    "để nhà trường biến khoản chi đăng ký thành một dòng chi thường xuyên của quỹ, thay vì một khoản phát sinh ngoài dự "
    "toán.",
]

KQ_4 = [
    "Bảng 2 so sánh công thức tính chỉ số sản phẩm khoa học quy đổi trên giảng viên của hai thông tư về Chuẩn cơ sở giáo "
    "dục đại học. Theo Thông tư số 01/2024/TT-BGDĐT, bằng độc quyền giải pháp hữu ích được tính chung nhóm với bài báo "
    "trong danh mục của Hội đồng Giáo sư nhà nước, hệ số 1. Theo Thông tư số 83/2026/TT-BGDĐT, có hiệu lực từ ngày 15 "
    "tháng 11 năm 2026, bằng độc quyền giải pháp hữu ích được xếp cùng nhóm sách chuyên khảo với hệ số 3, cao hơn bài báo "
    "thuộc Web of Science hoặc Scopus với hệ số 2; bằng độc quyền sáng chế giữ hệ số 5.",
    "Thay đổi này có ý nghĩa thực tiễn lớn với trường định hướng ứng dụng. Giải pháp hữu ích không đòi hỏi trình độ sáng "
    "tạo như sáng chế, phù hợp với quy mô đề tài cấp cơ sở, nay có giá trị quy đổi bằng ba bài báo trong nước. Cùng với "
    "đó, Thông tư số 83/2026/TT-BGDĐT xếp quy định về sở hữu trí tuệ, liêm chính khoa học, liêm chính học thuật vào danh "
    "mục nội dung quản trị nội bộ bắt buộc chung, được tính trong chỉ số tỷ lệ nội dung quản trị đã ban hành bằng văn bản "
    "đúng thẩm quyền và còn hiệu lực, và đưa khoản thu từ thương mại hóa kết quả nghiên cứu, sở hữu trí tuệ vào danh mục "
    "dữ liệu báo cáo. Điểm đ khoản 3 Điều 28 Luật số 125/2025/QH15 bổ sung nghĩa vụ công khai và cập nhật hằng năm kết "
    "quả hoạt động khoa học, công nghệ và đổi mới sáng tạo trên Nền tảng số quốc gia. Ở khía cạnh liêm chính, Điều 5a "
    "Nghị định số 134/2026/NĐ-CP xác định quyền tác giả đối với tác phẩm tạo ra có sử dụng trí tuệ nhân tạo chỉ phát "
    "sinh khi con người có đóng góp đáng kể và mang tính quyết định, nên quy chế về liêm chính cần quy định việc khai "
    "báo sử dụng trí tuệ nhân tạo trong học liệu và công bố. Sở hữu trí tuệ vì vậy trở thành "
    "chỉ số quản trị được đo lường và công khai, không còn là hoạt động tự nguyện.",
]

KQ_5 = [
    "Đối chiếu bốn nhóm thay đổi trên với chu trình quản lý cho thấy pháp luật mới tập trung vào các khâu từ xác lập trở "
    "đi. Riêng khâu tạo lập và nhận diện, tức xác định sản phẩm nào của một đề tài có khả năng bảo hộ trước khi được "
    "công bố, không văn bản nào quy định cơ chế cụ thể ở cấp cơ sở. Đây lại là khâu chịu ràng buộc thời gian chặt nhất: "
    "theo khoản 3 Điều 60 Luật Sở hữu trí tuệ, sáng chế không bị coi là mất tính mới khi được bộc lộ công khai nếu đơn "
    "được nộp trong thời hạn mười hai tháng kể từ ngày bộc lộ; quá thời hạn này, khả năng đăng ký sáng chế, giải pháp "
    "hữu ích không còn. Khi khuyến khích công bố được ghi nhận ngay còn văn bằng chỉ được ghi nhận sau khi cấp, lựa chọn "
    "hợp lý của giảng viên là công bố trước, và nếu không có bước rà soát bắt buộc thì sản phẩm đi thẳng từ nghiệm thu "
    "sang công bố.",
    "Hình 2 đề xuất mô hình cổng rà soát tại nghiệm thu. Nghiệm thu là thời điểm duy nhất mọi sản phẩm của đề tài đều "
    "đi qua một hội đồng có chuyên môn, nên là điểm can thiệp có chi phí thấp nhất. Phiếu rà soát bắt buộc yêu cầu hội "
    "đồng xác định sản phẩm thuộc đối tượng quyền nào, đã bộc lộ hay chưa, và chuyển các sản phẩm đủ điều kiện sang đầu "
    "mối xác lập quyền trong thời hạn luật định. Từ phân tích trên, Bảng 3 đề xuất khung rà soát quy chế nội bộ gồm chín "
    "nội dung, mỗi nội dung gắn với một căn cứ pháp lý và một câu hỏi kiểm tra cụ thể.",
]

BAN_LUAN = [
    "**Giải thích kết quả.** Khung pháp lý mới giải quyết các rào cản thuộc thẩm quyền của Nhà nước: quyền sở hữu, quyền "
    "quyết định, mức sàn lợi ích, nguồn chi và chỉ số đánh giá. Khâu nhận diện thì không thể quy định thống nhất từ "
    "trung ương, vì phụ thuộc vào quy trình nghiệm thu, cơ cấu ngành và năng lực chuyên môn của từng trường. Đó là lý do "
    "khoảng cách giữa sản phẩm khoa học và tài sản trí tuệ có thể vẫn tồn tại ngay cả khi luật đã thông thoáng, và lý do "
    "quy chế nội bộ trở thành biến số quyết định.",
    "**Đối chiếu với nghiên cứu trước.** Việc trao quyền cho tổ chức chủ trì theo Điều 25 có logic tương đồng với Đạo "
    "luật Bayh-Dole mà Shane (2004) phân tích. Mức sàn 30% và quyền tự quyết về chia lợi ích phù hợp với nhận định của "
    "Goldfarb và Henrekson (2003) và Siegel và cộng sự (2007) rằng khuyến khích cho nhà khoa học là động lực chính. Khác "
    "biệt nằm ở chỗ các nghiên cứu quốc tế mặc định có đơn vị chuyển giao tiếp nhận bản khai báo, như mô tả của Rocha và "
    "cộng sự (2023), và coi mức độ khai báo của giảng viên là biến số đầu vào (Thursby & Kemp, 2002). Ở trường đại học tư "
    "thục chưa có đơn vị chuyên trách, chính bước tiếp nhận đó cần được thiết kế, nên trọng tâm can thiệp dịch chuyển "
    "sớm hơn trong chuỗi. So với các nghiên cứu trong nước (Nguyễn, 2025; Võ, 2025), bài báo cập nhật khung pháp lý sau "
    "ngày 01 tháng 10 năm 2025 và chỉ ra cụ thể phần việc thuộc về quy chế nội bộ.",
    "**Đóng góp của bài báo.** Về lý luận, bài báo đề xuất xem khâu nhận diện trước công bố là một khâu quản lý độc lập "
    "trong chu trình, nơi chính sách cấp trường có tác động lớn nhất khi chính sách quốc gia đã thông thoáng. Về thực "
    "tiễn, khung rà soát chín nội dung và mô hình cổng rà soát tại nghiệm thu là công cụ mà trường đại học tư thục có thể "
    "áp dụng ngay khi hợp nhất quy chế theo Luật số 93/2025/QH15 và chuẩn bị đánh giá theo Thông tư số "
    "83/2026/TT-BGDĐT. Về chính sách, kết quả gợi ý cơ quan quản lý cần sớm hướng dẫn cách xác định lợi nhuận làm căn cứ "
    "thưởng tác giả và hướng dẫn tính các chỉ số sở hữu trí tuệ trong Chuẩn cơ sở giáo dục đại học.",
    "**Hạn chế của nghiên cứu.** Thứ nhất, bài báo phân tích văn bản, chưa đo lường mức độ thực thi tại các trường; các "
    "nhận định về hành vi giảng viên dựa trên lý thuyết và nghiên cứu quốc tế, chưa được kiểm chứng bằng dữ liệu Việt "
    "Nam. Thứ hai, văn bản hướng dẫn chi tiết Điều 27 Luật số 93/2025/QH15 và hướng dẫn tính chỉ số của Thông tư số "
    "83/2026/TT-BGDĐT có thể làm thay đổi một số hàm ý. Thứ ba, khung rà soát mới được xây dựng từ phân tích văn bản, "
    "chưa qua lấy ý kiến chuyên gia hay thử nghiệm.",
    "**Hướng nghiên cứu tiếp theo.** Các nghiên cứu tiếp theo có thể khảo sát quy chế nội bộ của một nhóm trường đại học "
    "tư thục theo khung chín nội dung, đánh giá tác động của Thông tư số 83/2026/TT-BGDĐT đến số đơn giải pháp hữu ích "
    "sau một đến hai năm áp dụng, và thử nghiệm mô hình cổng rà soát tại nghiệm thu bằng thiết kế trước và sau.",
]

KET_LUAN = [
    "Bài báo phân tích tác động của khung pháp lý giai đoạn 2025 - 2026 đến khoảng cách giữa sản phẩm khoa học và tài "
    "sản trí tuệ tại trường đại học tư thục và xác định nội dung quy chế nội bộ cần điều chỉnh.",
    "Phát hiện chính là pháp luật mới đã tháo gỡ phần lớn rào cản ở các khâu từ xác lập quyền trở đi: tổ chức chủ trì, "
    "kể cả tổ chức ngoài công lập, được giao quyền sở hữu và quyền quyết định thương mại hóa; lợi ích của tác giả chuyển "
    "sang mức sàn 30% đối với kết quả sử dụng ngân sách nhà nước và thỏa thuận đối với các nguồn khác; quỹ phát triển "
    "khoa học và công nghệ được chi cho đăng ký bảo hộ; chuẩn chất lượng mới nâng trọng số của giải pháp hữu ích lên "
    "gấp ba và biến quy chế sở hữu trí tuệ thành nội dung quản trị bắt buộc. Khâu nhận diện sản phẩm có khả năng bảo hộ "
    "trước khi công bố là khoảng trống mà pháp luật để lại cho quy chế nội bộ.",
    "Bài báo đóng góp một cách nhìn theo từng khâu đối với khung pháp lý mới, cùng khung rà soát quy chế nội bộ và mô hình "
    "cổng rà soát tại nghiệm thu. Hàm ý cho các trường đại học tư thục là hợp nhất quy chế theo Luật số 93/2025/QH15, bỏ "
    "các quy định theo cơ chế cũ, và đặt bước rà soát khả năng bảo hộ vào quy trình nghiệm thu trước khi đầu tư vào các "
    "thiết chế thương mại hóa phức tạp hơn.",
]

BANG_1 = dict(
    tieu_de="Những thay đổi của khung pháp lý về quyền đối với kết quả nghiên cứu và hàm ý cho quy chế nội bộ",
    cot=["Nội dung", "Quy định hiện hành", "Căn cứ", "Hàm ý cho quy chế nội bộ"],
    dong=[["Quyền sở hữu kết quả sử dụng ngân sách nhà nước",
           "Tổ chức chủ trì được tự động giao quyền, không làm thủ tục giao quyền, không bồi hoàn",
           "Khoản 2 Điều 25 Luật số 93/2025/QH15", "Xác định nhà trường là chủ sở hữu; phân định với phần đóng góp của "
                                                   "đối tác theo khoản 1 Điều 25"],
          ["Quyền đăng ký sở hữu công nghiệp", "Tổ chức được giao quyền có quyền đăng ký sáng chế, kiểu dáng công "
                                               "nghiệp, thiết kế bố trí",
           "Điểm c khoản 1 Điều 86 Luật Sở hữu trí tuệ", "Giao đầu mối nộp đơn nhân danh nhà trường"],
          ["Thương mại hóa", "Chủ sở hữu tự quyết hình thức, giá, góp vốn, phân chia lợi nhuận",
           "Điều 27 Luật số 93/2025/QH15; điểm d khoản 2 Điều 28 Luật số 125/2025/QH15",
           "Quy định thẩm quyền quyết định và quy trình định giá nội bộ"],
          ["Thưởng tác giả, kết quả sử dụng ngân sách nhà nước", "Tối thiểu 30% lợi nhuận, hoặc 30% giá trị khi góp vốn",
           "Điểm a khoản 3 Điều 28 Luật số 93/2025/QH15", "Ghi rõ mức sàn; bỏ khoản nộp và mức trần theo cơ chế cũ"],
          ["Thưởng tác giả, kết quả ngoài ngân sách", "Chủ sở hữu tự quyết", "Khoản 2 Điều 28 Luật số 93/2025/QH15",
           "Công bố trước công thức chia cho từng nguồn kinh phí"],
          ["Thù lao khi không có thỏa thuận", "10% lợi nhuận khi tự sử dụng; 15% tiền chuyển giao",
           "Khoản 1 Điều 135 Luật Sở hữu trí tuệ; khoản 2 Điều 135 đã bãi bỏ",
           "Mức mặc định; quy chế có thể quy định cao hơn"],
          ["Tổ chức ngoài công lập", "Được giao quyền và hưởng ưu đãi như tổ chức công lập",
           "Điểm b, điểm d khoản 2 Điều 37 Luật số 93/2025/QH15", "Chủ động tham gia nhiệm vụ sử dụng ngân sách nhà nước"]],
    nguon="Nguồn: Nhóm tác giả tổng hợp từ Luật số 93/2025/QH15, Luật Sở hữu trí tuệ hợp nhất và Luật số 125/2025/QH15.",
    rong=[3.2, 4.2, 3.8, 4.2], can=["left", "left", "left", "left"],
)

BANG_2 = dict(
    tieu_de="Hệ số của các loại sản phẩm trong chỉ số sản phẩm khoa học quy đổi trên giảng viên",
    cot=["Loại sản phẩm", "Thông tư số 01/2024/TT-BGDĐT", "Thông tư số 83/2026/TT-BGDĐT"],
    dong=[["Bài báo trong danh mục Hội đồng Giáo sư nhà nước hoặc đạt tiêu chuẩn khoa học Việt Nam, không thuộc Web of "
           "Science, Scopus", "1", "1"],
          ["Bài báo thuộc Web of Science hoặc Scopus", "1 nếu thuộc danh mục Hội đồng Giáo sư nhà nước", "2"],
          ["Sách chuyên khảo", "3", "3"],
          ["**Bằng độc quyền giải pháp hữu ích**", "**1**", "**3**"],
          ["Bằng độc quyền sáng chế", "5", "5"]],
    nguon="Nguồn: Nhóm tác giả tổng hợp từ công thức chỉ số 6.2.1 tại Thông tư số 01/2024/TT-BGDĐT và Thông tư số "
          "83/2026/TT-BGDĐT.",
    rong=[7.0, 4.2, 3.8], can=["left", "center", "center"],
)

BANG_3 = dict(
    tieu_de="Khung rà soát quy chế nội bộ về sở hữu trí tuệ của trường đại học theo khung pháp lý mới",
    cot=["Nội dung", "Câu hỏi rà soát", "Căn cứ"],
    dong=[["1. Phạm vi đối tượng", "Quy chế đã bao quát giải pháp hữu ích, sưu tập dữ liệu, chương trình máy tính, bí "
                                 "mật kinh doanh, giáo trình chưa?", "Điều 14, 22, 39 Luật Sở hữu trí tuệ"],
          ["2. Quyền theo nguồn kinh phí", "Có phân định kết quả từ ngân sách nhà nước, kinh phí nhà trường, đối tác và "
                                          "người học không?", "Điều 25 Luật số 93/2025/QH15"],
          ["3. Nhận diện trước công bố", "Có phiếu rà soát bắt buộc tại nghiệm thu và thủ tục khai báo trước khi công bố "
                                        "không?", "Khoản 3 Điều 60 Luật Sở hữu trí tuệ; Quyết định số 1624/QĐ-TTg"],
          ["4. Thẩm quyền thương mại hóa", "Ai quyết định hình thức, giá, góp vốn; quy trình định giá ra sao?",
           "Điều 27 Luật số 93/2025/QH15; Điều 28 Luật số 125/2025/QH15"],
          ["5. Phân chia lợi ích", "Đã ghi mức sàn 30%, công thức cho các nguồn khác và bỏ các quy định theo cơ chế cũ "
                                  "chưa?", "Điều 28 Luật số 93/2025/QH15; Điều 135 Luật Sở hữu trí tuệ"],
          ["6. Kinh phí xác lập, duy trì", "Có dòng chi phí nộp đơn, phí duy trì từ quỹ phát triển khoa học và công nghệ "
                                          "không?", "Điều 66 Luật số 93/2025/QH15; Điều 28 Luật số 125/2025/QH15"],
          ["7. Đầu mối và dữ liệu", "Có đầu mối chuyên trách, danh mục tài sản trí tuệ và cơ chế công khai hằng năm "
                                   "không?", "Điều 27 Luật số 93/2025/QH15; Điều 28 Luật số 125/2025/QH15"],
          ["8. Liêm chính và trí tuệ nhân tạo", "Có quy định về liêm chính khoa học, học thuật và việc sử dụng trí tuệ "
                                               "nhân tạo không?", "Thông tư số 83/2026/TT-BGDĐT; Nghị định số "
                                                                  "134/2026/NĐ-CP"],
          ["9. Hiệu lực và thống nhất", "Các văn bản nội bộ có còn hiệu lực, thống nhất, không trùng lặp quy định "
                                       "không?", "Tiêu chí 1.1 Thông tư số 83/2026/TT-BGDĐT"]],
    nguon="Nguồn: Nhóm tác giả đề xuất.",
    rong=[3.6, 7.4, 4.4], can=["left", "left", "left"],
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
    ("p", "Nhóm tác giả cam kết dữ liệu trong bài là trung thực, chính xác; toàn bộ dữ liệu là văn bản pháp luật, chỉ đạo "
          "và chuẩn chất lượng ban hành công khai, được nhóm tác giả lưu trữ và sẵn sàng cung cấp khi Ban Biên tập yêu cầu; "
          "chịu trách nhiệm trước pháp luật và trước Hội đồng Biên tập về nội dung bản thảo."),
    ("p", "Hà Nội, ngày ... tháng ... năm 2026"),
    ("p", "Đại diện nhóm tác giả (ký, ghi rõ họ tên)"),
]


def dung():
    v = khung.VanBanChung()
    dem = {}

    def phan(ten, ds):
        dem[ten] = dem.get(ten, 0) + sum(len(re.sub(r"\*", "", x).split()) for x in ds)
        v.than_md(*ds)

    p = v.doan("chuong1", TIEU_DE.upper(), bold=True)
    for r in p.runs:
        r.font.size = Pt(15)
    p.paragraph_format.space_after = Pt(8)
    v.doan("chuong2", TAC_GIA, bold=True)
    v.doan("chuong3", DON_VI, italic=True).paragraph_format.space_after = Pt(12)
    v.doan_md("than", "**Tóm tắt:** " + TOM_TAT)
    v.doan_md("than", "**Từ khóa:** " + TU_KHOA)
    v.doan_md("than", "**Abstract:** *" + ABSTRACT + "*")
    v.doan_md("than", "**Keywords:** *" + KEYWORDS + "*")

    v.doan("h1", "1. ĐẶT VẤN ĐỀ", bold=True)
    phan("Đặt vấn đề", DAT_VAN_DE)
    v.doan("h1", "2. TỔNG QUAN NGHIÊN CỨU", bold=True)
    phan("Tổng quan", TONG_QUAN)
    v.doan("h1", "3. PHƯƠNG PHÁP NGHIÊN CỨU", bold=True)
    phan("Phương pháp", PHUONG_PHAP[:1])
    v.so_do("Chu trình quản lý tài sản trí tuệ trong trường đại học dùng làm khung phân tích",
            os.path.join(SO_DO, "chu_trinh.png"), "Nguồn: Nhóm tác giả xây dựng.", tien_to="", rong_cm=13.5)
    phan("Phương pháp", PHUONG_PHAP[1:])

    v.doan("h1", "4. KẾT QUẢ NGHIÊN CỨU", bold=True)
    v.doan("h2", "4.1. Quyền sở hữu và quyền quyết định thương mại hóa được trao cho tổ chức chủ trì, không phân biệt "
                 "công lập hay ngoài công lập")
    phan("Kết quả", KQ_1)
    v.doan("h2", "4.2. Cơ chế chia lợi ích chuyển từ khung định sẵn sang mức sàn và thỏa thuận")
    phan("Kết quả", KQ_2)
    for b in (BANG_1,):
        v.bang(b["tieu_de"], b["cot"], b["dong"], b["nguon"], b["rong"], can=b["can"], tien_to="")
    v.doan("h2", "4.3. Nguồn lực cho xác lập quyền đã có căn cứ pháp lý")
    phan("Kết quả", KQ_3)
    v.doan("h2", "4.4. Chuẩn chất lượng mới nâng giá trị của giải pháp hữu ích và biến quy chế sở hữu trí tuệ thành nội "
                 "dung quản trị bắt buộc")
    phan("Kết quả", KQ_4[:1])
    v.bang(BANG_2["tieu_de"], BANG_2["cot"], BANG_2["dong"], BANG_2["nguon"], BANG_2["rong"], can=BANG_2["can"],
           tien_to="")
    phan("Kết quả", KQ_4[1:])
    v.doan("h2", "4.5. Khâu nhận diện trước công bố là khoảng trống pháp luật để lại cho quy chế nội bộ")
    phan("Kết quả", KQ_5[:1])
    v.so_do("Mô hình cổng rà soát khả năng bảo hộ tại nghiệm thu đề tài", os.path.join(SO_DO, "cong_ra_soat.png"),
            "Nguồn: Nhóm tác giả đề xuất.", tien_to="", rong_cm=15.5)
    phan("Kết quả", KQ_5[1:])
    v.bang(BANG_3["tieu_de"], BANG_3["cot"], BANG_3["dong"], BANG_3["nguon"], BANG_3["rong"], can=BANG_3["can"],
           tien_to="")

    v.doan("h1", "5. BÀN LUẬN", bold=True)
    phan("Bàn luận", BAN_LUAN)
    v.doan("h1", "6. KẾT LUẬN", bold=True)
    phan("Kết luận", KET_LUAN)

    van_ban = "\n".join([p.text for p in v.doc.paragraphs] +
                        [c.text for t in v.doc.tables for r in t.rows for c in r.cells])
    v.doan("h1", "TÀI LIỆU THAM KHẢO", bold=True)
    dung_tl = TL.duoc_trich(van_ban)
    for apa in TL.danh_muc(dung_tl):
        p = v.doan_md("than", apa)
        pf = p.paragraph_format
        pf.left_indent, pf.first_line_indent, pf.space_after = Pt(28), Pt(-28), Pt(4)

    for loai, t in TO_KHAI:
        if loai == "h":
            v.doan("chuong1", t, bold=True).paragraph_format.page_break_before = True
        elif loai == "b":
            v.doan_md("than", f"**{t}**")
        elif loai == "tbl":
            v.bang("Kê khai phạm vi sử dụng công cụ trí tuệ nhân tạo (Phương án B)",
                   ["Tên công cụ", "Phạm vi hỗ trợ", "Nội dung hoặc mục có sự hỗ trợ"],
                   [["...", "...", "..."], ["...", "...", "..."]], None, [3.5, 5.5, 6.0],
                   can=["left", "left", "left"], tien_to="")
        else:
            v.doan_md("than", t)

    v.doc.core_properties.title = TIEU_DE
    v.doc.core_properties.author = TAC_GIA
    v.doc.save(RA_DOCX)

    tong = sum(dem.values())
    khung_tl = {"Đặt vấn đề": (5, 7), "Tổng quan": (15, 20), "Phương pháp": (6, 10), "Kết quả": (40, 55),
                "Bàn luận": (15, 25), "Kết luận": (5, 8)}
    print("Đã ghi:", RA_DOCX)
    print(f"Tiêu đề: {len(TIEU_DE.split())} âm tiết; tóm tắt: {len(TOM_TAT.split())}; abstract: "
          f"{len(ABSTRACT.split())} từ; từ khóa: {len(TU_KHOA.split(';'))} cụm; tài liệu: {len(dung_tl)}")
    for k, (a, b_) in khung_tl.items():
        tl = 100 * dem[k] / tong
        print(f"  {k:12s} {dem[k]:5d}  {tl:5.1f}%  (khung {a}-{b_}%){'' if a <= tl <= b_ else '  NGOÀI KHUNG'}")
    print(f"  Tổng thân bài {tong}")
    mo_coi = TL.trich_dan_mo_coi(van_ban)
    if mo_coi:
        print("CẢNH BÁO trích dẫn không có trong danh mục:", mo_coi)
    for x in khung.kiem_tra(v.doc):
        print("CẢNH BÁO:", x)


if __name__ == "__main__":
    dung()

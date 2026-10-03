# -*- coding: utf-8 -*-
"""Bài báo đầu ra (phương án chính, theo tiêu đề của chủ nhiệm đề tài).

    python3 scripts/ban_cuoi/bai_bao_tong_hop.py

Đầu ra: Ban_cuoi/Bai_bao_Tap_chi_NCKH_PT.docx

Ghép hai phương án trước: phần phân tích yêu cầu thể chế mới (bai_bao_chinh_sach.py) và phần
bộ chỉ số theo chuỗi kết quả (bai_bao_khung.py), thêm phần giải pháp rút từ Chương 3 báo cáo
tổng kết. Chỉ dùng văn bản ban hành công khai, không công bố số liệu nội bộ.
"""
import os
import re
import sys

from docx.shared import Pt

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import khung  # noqa: E402
import tai_lieu as TL  # noqa: E402
import bai_bao_chinh_sach as CS  # noqa: E402

RA_DOCX = os.path.join(khung.THU_MUC_RA, "Bai_bao_Tap_chi_NCKH_PT.docx")
SO_DO = os.path.join(khung.THU_MUC_RA, "so_do")

TIEU_DE = ("Nâng cao hiệu quả quản trị quyền sở hữu trí tuệ trong cơ sở giáo dục đại học trước những yêu cầu thể chế "
           "mới")
TAC_GIA = CS.TAC_GIA
DON_VI = CS.DON_VI

TOM_TAT = (
    "Bài báo phân tích những yêu cầu thể chế mới giai đoạn 2025 - 2026 đối với quản trị quyền sở hữu trí tuệ trong cơ "
    "sở giáo dục đại học và đề xuất giải pháp nâng cao hiệu quả. Nghiên cứu kết hợp phân tích 12 văn bản pháp luật, chỉ "
    "đạo và chuẩn chất lượng ban hành công khai với tổng hợp nghiên cứu quốc tế về chuyển giao công nghệ và đo lường "
    "chuyển giao tri thức, theo chu trình tạo lập, xác lập, khai thác và bảo vệ tài sản trí tuệ. Kết quả cho thấy khung "
    "pháp lý mới đặt ra năm nhóm yêu cầu về quyền sở hữu, phân chia lợi ích, nguồn lực, đo lường và tính thống nhất của "
    "quy chế nội bộ. Đối chiếu các yêu cầu này với công cụ quản trị cho thấy bốn khoảng trống: quy chế ban hành trước "
    "luật mới, khâu nhận diện sản phẩm trước khi công bố còn bỏ ngỏ, thiếu dòng kinh phí cho xác lập quyền và thiếu dữ "
    "liệu theo dõi. Từ đó, bài báo đề xuất năm nhóm giải pháp gắn với căn cứ pháp lý, trong đó cổng rà soát khả năng "
    "bảo hộ tại nghiệm thu là giải pháp then chốt, cùng bộ chỉ số theo dõi theo chuỗi đầu vào, quá trình, đầu ra, kết "
    "quả để nhà trường xác định điểm nghẽn và điều chỉnh kịp thời.")
TU_KHOA = ("Quản trị tài sản trí tuệ; Quyền sở hữu trí tuệ; Cơ sở giáo dục đại học; Yêu cầu thể chế; Luật Khoa học, công "
           "nghệ và đổi mới sáng tạo")

ABSTRACT = (
    "This article analyses the new institutional requirements of 2025 - 2026 for intellectual property rights "
    "governance in Vietnamese higher education institutions and proposes measures to improve its effectiveness. The "
    "study combines an analysis of twelve publicly issued laws, policy directives and quality standards with a "
    "synthesis of international research on university technology transfer and knowledge transfer metrics, structured "
    "around the intellectual property life cycle of creation, registration, exploitation and protection. The results "
    "show that the new legal framework sets five groups of requirements concerning ownership, benefit sharing, "
    "resources, measurement and the coherence of internal regulations. Comparing these requirements with institutional "
    "governance tools reveals four gaps: regulations issued before the new laws, an unaddressed stage of identifying "
    "protectable outputs before publication, the lack of a budget line for registration, and insufficient monitoring "
    "data. The article therefore proposes five groups of measures anchored in specific legal provisions, with a "
    "protectability screening gate at project acceptance as the key measure, together with a monitoring indicator set "
    "organised along the input, process, output and outcome chain that enables institutions to locate bottlenecks and "
    "adjust in time. The findings indicate that, once national barriers have been lowered, effectiveness depends "
    "mainly on institutional governance design rather than on further changes in national policy.")
KEYWORDS = ("Intellectual asset governance; Intellectual property rights; Higher education institutions; Institutional "
            "requirements; Law on Science, Technology and Innovation")

DAT_VAN_DE = [
    "Giai đoạn 2025 - 2026, khung thể chế về khoa học, công nghệ và sở hữu trí tuệ của Việt Nam thay đổi sâu rộng. Luật "
    "Khoa học, công nghệ và đổi mới sáng tạo số 93/2025/QH15, Luật số 131/2025/QH15 sửa đổi Luật Sở hữu trí tuệ, Luật "
    "Giáo dục đại học số 125/2025/QH15 và Chuẩn cơ sở giáo dục đại học ban hành kèm Thông tư số 83/2026/TT-BGDĐT cùng "
    "trả lời lại các câu hỏi cốt lõi của quản trị tài sản trí tuệ: ai sở hữu kết quả nghiên cứu, ai quyết định khai "
    "thác, tác giả được hưởng bao nhiêu, chi phí đăng ký lấy từ đâu và hiệu quả được đo bằng gì.",
    "Nghiên cứu quốc tế đã chỉ ra vai trò của thể chế phân bổ quyền, cơ chế khuyến khích và năng lực tổ chức đối với "
    "hiệu quả chuyển giao công nghệ đại học (Goldfarb & Henrekson, 2003; Siegel et al., 2007). Các nghiên cứu trong nước "
    "phần lớn ra đời trước khi Luật số 93/2025/QH15 có hiệu lực (Nguyễn, 2025; Võ, 2025), nên chưa chỉ ra cơ sở giáo dục "
    "đại học cần làm gì cụ thể để đáp ứng khung thể chế mới và theo dõi kết quả bằng cách nào.",
    "Bài báo trả lời ba câu hỏi: khung thể chế mới đặt ra những yêu cầu gì đối với quản trị quyền sở hữu trí tuệ trong "
    "cơ sở giáo dục đại học; giữa yêu cầu và công cụ quản trị nội bộ còn những khoảng trống nào; và giải pháp nào giúp "
    "nâng cao hiệu quả, được theo dõi bằng những chỉ số nào.",
]

TONG_QUAN = [
    "**Thể chế phân bổ quyền.** Shane (2004) cho thấy việc trao quyền sở hữu kết quả nghiên cứu do nhà nước tài trợ cho "
    "trường đại học theo Đạo luật Bayh-Dole làm tăng số sáng chế của trường đại học Hoa Kỳ, chủ yếu ở những lĩnh vực mà "
    "cấp phép là kênh chuyển giao hiệu quả. Etzkowitz (2003) lý giải sự hình thành đại học khởi nghiệp từ đặc tính gần "
    "với doanh nghiệp của các nhóm nghiên cứu. Các công trình này khẳng định trao quyền cho tổ chức chủ trì là điều kiện "
    "cần của thương mại hóa.",
    "**Khuyến khích và năng lực tổ chức.** Goldfarb và Henrekson (2003) kết luận rằng khuyến khích đối với nhà khoa học "
    "và nhà trường quan trọng hơn can thiệp trực tiếp của nhà nước. Thursby và Kemp (2002) cho thấy tăng trưởng cấp phép "
    "gắn với mức độ sẵn sàng khai báo sáng chế của giảng viên và hiệu quả khác nhau đáng kể giữa các trường. Siegel và "
    "cộng sự (2007) tổng hợp rằng hiệu quả của đơn vị chuyển giao phụ thuộc vào chính sách khuyến khích, năng lực nhân sự "
    "và cơ chế chia lợi ích. Bradley và cộng sự (2013) phê phán mô hình tuyến tính và nhấn mạnh tính đa kênh của chuyển "
    "giao công nghệ.",
    "**Quy trình quản lý và khai thác.** Holgersson và Aaboen (2019) chỉ ra sự dịch chuyển trọng tâm từ chiếm hữu sang "
    "khai thác tài sản trí tuệ; Maresova và cộng sự (2019) hệ thống hóa các mô hình và quy trình quản lý chuyển giao "
    "trong trường đại học. Rocha và cộng sự (2023) mô tả bước đánh giá bản khai báo sáng chế theo tính mới, khả năng áp "
    "dụng công nghiệp và trình độ sáng tạo trước khi quyết định bảo hộ. Perkmann và cộng sự (2013) và O’Dwyer và cộng sự "
    "(2023) cho thấy hợp tác với doanh nghiệp đòi hỏi cơ chế phân định quyền rõ ràng. Teece (2018) nhắc lại rằng khả năng "
    "thu lợi từ đổi mới phụ thuộc vào chế độ bảo hộ và tài sản bổ trợ, nên việc bảo hộ đúng lúc là điều kiện của khai "
    "thác. Điểm chung của nhóm nghiên cứu này là coi quy trình tiếp nhận, sàng lọc và theo dõi tài sản trí tuệ là một "
    "năng lực quản trị cần được thiết kế, không tự hình thành khi đã có quyền sở hữu.",
    "**Đo lường hiệu quả.** Nhóm chuyên gia của Ủy ban châu Âu đề xuất tập chỉ số cốt lõi về chuyển giao tri thức gồm "
    "khai báo sáng chế, đơn đăng ký, văn bằng, hợp đồng cấp phép, nguồn thu và doanh nghiệp khởi nguồn (Finne et al., "
    "2009); Campbell và cộng sự (2020) hướng tới bộ chỉ số hài hòa và mở rộng sang các kênh chuyển giao khác. Mô hình "
    "logic của W. K. Kellogg Foundation (2004) cung cấp cách đặt chỉ số theo chuỗi nguồn lực, hoạt động, đầu ra và kết "
    "quả. Các bộ chỉ số này chủ yếu đếm đầu ra và giả định đã có đơn vị chuyển giao chuyên nghiệp.",
    "**Nghiên cứu trong nước và khoảng trống.** Nguyễn (2025) phân tích cơ hội và thách thức của quản trị tài sản trí tuệ "
    "trong cơ sở giáo dục đại học; Võ (2025) bàn về quyền của chủ thể không giữ quyền tài sản đối với tác phẩm hình thành "
    "trong nhà trường; tài liệu tập huấn của Cục Sở hữu trí tuệ (n.d.) cung cấp khung nghiệp vụ cho cán bộ quản lý. Còn "
    "thiếu một nghiên cứu nối liền ba bước: xác định yêu cầu của khung thể chế mới theo từng khâu, chỉ ra khoảng trống "
    "giữa yêu cầu và công cụ quản trị nội bộ, rồi đề xuất giải pháp kèm cách theo dõi hiệu quả. Bài báo hướng tới lấp "
    "khoảng trống này.",
]

PHUONG_PHAP = [
    "Nghiên cứu sử dụng phương pháp phân tích văn bản chính sách kết hợp tổng hợp lý thuyết. Khung phân tích là chu "
    "trình quản lý tài sản trí tuệ bốn giai đoạn gồm tạo lập và nhận diện, xác lập quyền, khai thác và thương mại hóa, "
    "bảo vệ và phân chia lợi ích (Hình 1).",
    "Dữ liệu là 12 văn bản ban hành công khai, có hiệu lực hoặc đã ban hành đến ngày 30 tháng 9 năm 2026 và có quy định "
    "trực tiếp về quyền đối với kết quả nghiên cứu, phân chia lợi ích, nguồn lực hoặc chỉ số sở hữu trí tuệ trong cơ sở "
    "giáo dục đại học: Luật Sở hữu trí tuệ hợp nhất (Văn phòng Quốc hội, 2026), Luật số 93/2025/QH15, Luật số "
    "131/2025/QH15, Luật số 125/2025/QH15, Luật số 07/2017/QH14, Nghị định số 134/2026/NĐ-CP, Quyết định số 1068/QĐ-TTg, "
    "Quyết định số 1624/QĐ-TTg, Chỉ thị số 02/CT-TTg, Kết luận số 51-KL/TW, Thông tư số 01/2024/TT-BGDĐT và Thông tư số "
    "83/2026/TT-BGDĐT.",
    "Phân tích gồm ba bước. Bước thứ nhất mã hóa từng văn bản theo điều, khoản và gán vào khâu tương ứng của chu trình để "
    "xác định các nhóm yêu cầu. Bước thứ hai đối chiếu yêu cầu với những nội dung pháp luật để lại cho cơ sở tự quy định, "
    "kết hợp với các phát hiện của nghiên cứu quốc tế ở Mục 2, để chỉ ra khoảng trống. Bước thứ ba đề xuất giải pháp cho "
    "từng khoảng trống và xây dựng chỉ số theo dõi theo mô hình chuỗi kết quả (W. K. Kellogg Foundation, 2004). Các giải "
    "pháp được phát triển từ đề tài khoa học công nghệ cấp cơ sở năm 2026 của nhóm tác giả; bài báo không công bố số liệu "
    "nội bộ của cơ sở nào.",
]

KQ_YEU_CAU = [
    "Phân tích 12 văn bản cho thấy năm nhóm yêu cầu, được tóm tắt tại Bảng 1.",
    "**Thứ nhất, về quyền sở hữu và quyền quyết định khai thác.** Theo khoản 2 Điều 25 Luật số 93/2025/QH15, tổ chức chủ "
    "trì được Nhà nước tự động giao quyền quản lý, sử dụng, quyền sở hữu phần kết quả tương ứng với kinh phí ngân sách "
    "nhà nước, không phải làm thủ tục giao quyền và không phải bồi hoàn chi phí; Điều 27 cho phép tổ chức tự quyết định "
    "hình thức, giá, phân chia lợi nhuận và phương án góp vốn khi thương mại hóa. Luật số 131/2025/QH15 bổ sung điểm c "
    "khoản 1 Điều 86 Luật Sở hữu trí tuệ, trao quyền đăng ký sáng chế, kiểu dáng công nghiệp, thiết kế bố trí cho tổ "
    "chức được giao quyền. Điểm b và điểm d khoản 2 Điều 37 Luật số 93/2025/QH15 bảo đảm tổ chức ngoài công lập được giao "
    "quyền và hưởng ưu đãi như tổ chức công lập.",
    "**Thứ hai, về phân chia lợi ích.** Điểm a khoản 3 Điều 28 Luật số 93/2025/QH15 quy định thưởng cho tác giả tối thiểu "
    "30% lợi nhuận từ thương mại hóa kết quả sử dụng ngân sách nhà nước; với kết quả ngoài ngân sách, chủ sở hữu tự quyết "
    "theo khoản 2. Khoản 2 Điều 135 Luật Sở hữu trí tuệ bị bãi bỏ; khoản 1 chỉ áp dụng mức mặc định 10% hoặc 15% khi các "
    "bên không có thỏa thuận. Pháp luật vì vậy chuyển từ khung định sẵn sang mức sàn kết hợp thỏa thuận.",
    "**Thứ ba, về nguồn lực.** Điểm b khoản 2 Điều 66 Luật số 93/2025/QH15 cho phép quỹ phát triển khoa học và công nghệ "
    "của tổ chức chi cho đăng ký, bảo hộ, quản lý, khai thác quyền sở hữu trí tuệ; điểm d khoản 3 Điều 28 Luật số "
    "125/2025/QH15 giao cơ sở giáo dục đại học trách nhiệm thành lập và vận hành quỹ này.",
    "**Thứ tư, về đo lường và công khai.** Quyết định số 1624/QĐ-TTg yêu cầu sử dụng chỉ số sở hữu trí tuệ làm căn cứ "
    "đánh giá hiệu quả hoạt động của cơ sở giáo dục đại học. Điểm đ khoản 3 Điều 28 Luật số 125/2025/QH15 yêu cầu công "
    "khai và cập nhật hằng năm kết quả hoạt động khoa học, công nghệ và đổi mới sáng tạo; khoản 4 Điều 27 Luật số "
    "93/2025/QH15 yêu cầu báo cáo kết quả thương mại hóa. Bảng 2 cho thấy Thông tư số 83/2026/TT-BGDĐT nâng hệ số của "
    "bằng độc quyền giải pháp hữu ích trong chỉ số sản phẩm khoa học quy đổi từ 1 lên 3, cao hơn bài báo thuộc Web of "
    "Science hoặc Scopus với hệ số 2.",
    "**Thứ năm, về tính thống nhất của quy chế và năng lực.** Thông tư số 83/2026/TT-BGDĐT xếp quy định về sở hữu trí tuệ, "
    "liêm chính khoa học, liêm chính học thuật vào danh mục nội dung quản trị nội bộ bắt buộc chung, được đánh giá theo "
    "tiêu chí văn bản ban hành đúng thẩm quyền và còn hiệu lực. Điều 5a Nghị định số 134/2026/NĐ-CP xác định quyền tác "
    "giả đối với tác phẩm có sử dụng trí tuệ nhân tạo chỉ phát sinh khi con người có đóng góp đáng kể. Quyết định số "
    "1624/QĐ-TTg đặt yêu cầu nghiên cứu đưa sở hữu trí tuệ thành nội dung học bắt buộc tại cơ sở giáo dục đại học.",
]

KQ_KHOANG_TRONG = [
    "Đối chiếu năm nhóm yêu cầu với những nội dung pháp luật để lại cho cơ sở tự quy định cho thấy bốn khoảng trống mà "
    "quản trị nội bộ cần lấp.",
    "**Khoảng trống thứ nhất: độ trễ của quy chế nội bộ.** Các quy chế ban hành trước ngày 01 tháng 10 năm 2025 được xây "
    "dựng theo khung pháp luật khi đó, nên có thể còn khoản nộp ngân sách, mức trần tiền thưởng hoặc dẫn chiếu khung thù "
    "lao đã bị bãi bỏ. Một cơ sở cũng có thể có nhiều văn bản cùng quy định về phân chia lợi ích cho các kênh tài trợ "
    "khác nhau. Đây là độ trễ thể chế tất yếu trước sự thay đổi dồn dập của pháp luật, nhưng cần xử lý sớm vì mức sàn "
    "30% là quy định bắt buộc và quy chế còn hiệu lực, thống nhất là tiêu chí của chuẩn chất lượng.",
    "**Khoảng trống thứ hai: khâu nhận diện trước công bố.** Pháp luật mới tập trung vào các khâu từ xác lập quyền trở "
    "đi; không văn bản nào quy định cơ chế cụ thể để xác định sản phẩm nào của một đề tài có khả năng bảo hộ trước khi "
    "được công bố. Theo khoản 3 Điều 60 Luật Sở hữu trí tuệ, sáng chế đã bộc lộ công khai chỉ không bị coi là mất tính "
    "mới nếu đơn được nộp trong mười hai tháng kể từ ngày bộc lộ. Khi công bố được ghi nhận ngay còn văn bằng chỉ được ghi "
    "nhận sau khi cấp, lựa chọn hợp lý của giảng viên là công bố trước; nếu không có bước rà soát, sản phẩm đi thẳng từ "
    "nghiệm thu sang công bố. Ở cơ sở chưa có đơn vị chuyển giao chuyên nghiệp, bước tiếp nhận và đánh giá bản khai báo "
    "mà Rocha và cộng sự (2023) mô tả thường chưa tồn tại.",
    "**Khoảng trống thứ ba: dòng kinh phí cho xác lập quyền.** Chi phí nộp đơn và duy trì hiệu lực phát sinh trước khi có "
    "doanh thu. Nếu quy chế chỉ coi khoản chi này là chi phí được trừ khi chia lợi ích, thủ tục không có nguồn để khởi "
    "động, dù Điều 66 Luật số 93/2025/QH15 đã cho phép quỹ phát triển khoa học và công nghệ chi cho nội dung này.",
    "**Khoảng trống thứ tư: dữ liệu và đầu mối theo dõi.** Nghĩa vụ công khai và báo cáo ngầm định nhà trường có một danh "
    "mục tài sản trí tuệ thống nhất. Khi dữ liệu về đề tài, đơn, văn bằng, hợp đồng và kinh phí nằm rải rác ở nhiều đơn "
    "vị, nhà trường khó đáp ứng nghĩa vụ báo cáo và càng khó biết tiềm năng bị mất ở khâu nào.",
]

GIAI_PHAP = [
    "Từ bốn khoảng trống và năm nhóm yêu cầu, bài báo đề xuất năm nhóm giải pháp, tóm tắt tại Bảng 3. Mỗi giải pháp gắn "
    "với một khoảng trống, một căn cứ pháp lý và chỉ số theo dõi.",
    "**Giải pháp 1: Hợp nhất quy chế quản trị tài sản trí tuệ theo luật mới.** Mục tiêu là bảo đảm một văn bản duy nhất, "
    "còn hiệu lực và tương thích với Luật số 93/2025/QH15. Nội dung gồm: xác định quyền sở hữu theo nguồn kinh phí theo "
    "Điều 25; quy định thẩm quyền quyết định thương mại hóa và quy trình định giá theo Điều 27; ghi rõ mức sàn 30% cho "
    "tác giả đối với kết quả sử dụng ngân sách nhà nước, công bố trước công thức chia cho các nguồn khác và bỏ các quy "
    "định theo cơ chế cũ; bổ sung quy định về liêm chính khoa học, học thuật và khai báo sử dụng trí tuệ nhân tạo.",
    "**Giải pháp 2: Thiết lập cổng rà soát khả năng bảo hộ tại nghiệm thu.** Đây là giải pháp then chốt cho khoảng trống "
    "thứ hai (Hình 2). Nghiệm thu là thời điểm duy nhất mọi sản phẩm của đề tài đều đi qua một hội đồng chuyên môn, nên "
    "là điểm can thiệp có chi phí thấp nhất. Phiếu rà soát bắt buộc gồm năm câu hỏi: sản phẩm cụ thể là gì; thuộc đối "
    "tượng quyền nào; đã bộc lộ công khai chưa và vào ngày nào; chủ sở hữu xác định theo nguồn kinh phí ra sao; và hội "
    "đồng đề xuất nộp đơn, đăng ký quyền tác giả, giữ bí mật hay không bảo hộ. Sản phẩm đủ điều kiện được chuyển cho đầu "
    "mối xác lập quyền trong thời hạn luật định. Nên kết hợp với việc xác định trước đối tượng quyền dự kiến ngay từ "
    "thuyết minh đề tài.",
    "**Giải pháp 3: Xây dựng cơ chế tài chính cho xác lập quyền.** Mục tiêu là bảo đảm không sản phẩm đủ điều kiện nào "
    "dừng lại vì thiếu kinh phí. Nhà trường bố trí dòng chi thường xuyên cho phí nộp đơn, phí duy trì hiệu lực từ quỹ phát "
    "triển khoa học và công nghệ theo điểm b khoản 2 Điều 66 Luật số 93/2025/QH15, và ghi nhận văn bằng trong chế độ "
    "khuyến khích tương xứng với công bố. Với hệ số 3 của bằng độc quyền giải pháp hữu ích theo Thông tư số "
    "83/2026/TT-BGDĐT, khoản chi này đồng thời góp phần nâng chỉ số sản phẩm khoa học quy đổi của nhà trường.",
    "**Giải pháp 4: Kiện toàn đầu mối và hệ thống dữ liệu.** Nhà trường giao một đầu mối chuyên trách quản lý toàn bộ "
    "chu trình, thiết lập cơ chế phối hợp giữa đơn vị quản lý khoa học, bộ phận pháp chế, tài chính và các khoa, viện, và "
    "xây dựng danh mục tài sản trí tuệ số hóa dùng chung, liên kết đề tài với đơn, văn bằng và hợp đồng. Danh mục này "
    "phục vụ trực tiếp nghĩa vụ công khai theo Luật số 125/2025/QH15 và báo cáo thương mại hóa theo Luật số "
    "93/2025/QH15. Với nhóm sản phẩm phù hợp, hợp tác với doanh nghiệp và doanh nghiệp quản lý tài sản trí tuệ theo "
    "khoản 1 Điều 28 Luật số 125/2025/QH15 là kênh khai thác cần được tính đến.",
    "**Giải pháp 5: Phát triển năng lực và văn hóa sở hữu trí tuệ.** Nhà trường tổ chức tập huấn định kỳ cho giảng viên, "
    "thành viên hội đồng nghiệm thu và cán bộ quản lý về nhận diện đối tượng quyền và thời hạn mười hai tháng; xây dựng "
    "học phần sở hữu trí tuệ cho người học theo định hướng của Quyết định số 1624/QĐ-TTg; lồng ghép nội dung liêm chính "
    "học thuật và sử dụng trí tuệ nhân tạo.",
]

KQ_CHI_SO = [
    "Để biết giải pháp có phát huy tác dụng hay không, nhà trường cần theo dõi theo chuỗi đầu vào, quá trình, đầu ra và "
    "kết quả, đồng thời đặt chỉ số tại các mối nối giữa các mắt xích (Hình 3). Bốn chỉ số chuyển hóa gồm: tỷ lệ nhận "
    "diện, là phần công trình nghiệm thu có sản phẩm được xác nhận có khả năng bảo hộ; tỷ lệ xác lập kịp thời, là phần "
    "sản phẩm có khả năng bảo hộ được nộp đơn trong mười hai tháng; tỷ lệ khai thác, là phần tài sản đã xác lập có giao "
    "dịch hoặc được sử dụng; và tỷ suất khai thác trên chi phí, là nguồn thu từ khai thác so với chi phí xác lập và duy "
    "trì quyền.",
    "Khác với chỉ số đếm như số văn bằng hay nguồn thu, mỗi chỉ số chuyển hóa chỉ vào một mối nối cụ thể. Tỷ lệ nhận diện "
    "thấp cho thấy cần tác động vào định hướng đề tài và năng lực hội đồng; tỷ lệ nhận diện cao nhưng tỷ lệ xác lập kịp "
    "thời thấp cho thấy điểm nghẽn ở khâu nối giữa nghiệm thu và đăng ký, nơi Giải pháp 2 và Giải pháp 3 phát huy tác "
    "dụng; tỷ lệ xác lập cao nhưng tỷ lệ khai thác thấp cho thấy cần đẩy mạnh Giải pháp 4 ở khâu khai thác.",
    "Đối chiếu với các nghĩa vụ báo cáo hiện hành cho thấy dữ liệu bắt buộc tập trung ở nhóm kết quả, như văn bằng, nguồn "
    "thu và tỷ trọng thu khoa học, công nghệ theo Thông tư số 83/2026/TT-BGDĐT, trong khi cả bốn chỉ số chuyển hóa đều "
    "không thuộc dữ liệu bắt buộc. Hai chỉ số then chốt về nhận diện và xác lập kịp thời chỉ tính được khi có phiếu rà "
    "soát tại nghiệm thu. Vì vậy, Giải pháp 2 vừa là biện pháp quản trị, vừa là điều kiện để đo lường hiệu quả. Nhà "
    "trường có thể triển khai theo ba giai đoạn: dùng chỉ số đã có dữ liệu để lập đường cơ sở; đưa phiếu rà soát vào quy "
    "trình nghiệm thu để tính các chỉ số chuyển hóa; bổ sung khảo sát giảng viên và danh mục khai thác để đánh giá kết "
    "quả đầy đủ.",
]

BAN_LUAN = [
    "**Giải thích kết quả.** Khung thể chế mới đã giải quyết các rào cản thuộc thẩm quyền của Nhà nước: quyền sở hữu, "
    "quyền quyết định, mức sàn lợi ích, nguồn chi và chỉ số đánh giá. Khâu nhận diện thì không thể quy định thống nhất từ "
    "trung ương, vì phụ thuộc vào quy trình nghiệm thu, cơ cấu ngành và năng lực chuyên môn của từng trường. Do đó, hiệu "
    "quả quản trị quyền sở hữu trí tuệ trong giai đoạn tới phụ thuộc chủ yếu vào thiết kế quản trị ở cấp cơ sở, chứ "
    "không còn chủ yếu vào việc chờ chính sách quốc gia thông thoáng hơn.",
    "**Đối chiếu với nghiên cứu trước.** Việc trao quyền cho tổ chức chủ trì có logic tương đồng với Đạo luật Bayh-Dole mà "
    "Shane (2004) phân tích; mức sàn 30% và quyền tự quyết về chia lợi ích phù hợp với nhận định của Goldfarb và "
    "Henrekson (2003) và Siegel và cộng sự (2007) về vai trò của khuyến khích. Khác biệt là các nghiên cứu quốc tế mặc "
    "định có đơn vị chuyển giao tiếp nhận bản khai báo (Rocha et al., 2023) và coi mức độ khai báo là biến số đầu vào "
    "(Thursby & Kemp, 2002); ở cơ sở chưa có đơn vị này, chính bước tiếp nhận cần được thiết kế, nên trọng tâm can thiệp "
    "dịch chuyển sớm hơn trong chuỗi. So với tập chỉ số cốt lõi của Finne và cộng sự (2009), bộ chỉ số đề xuất giữ các "
    "chỉ số đếm nhưng bổ sung chỉ số chuyển hóa, phù hợp với cách nhìn chuyển từ chiếm hữu sang khai thác của Holgersson "
    "và Aaboen (2019). So với các nghiên cứu trong nước (Nguyễn, 2025; Võ, 2025), bài báo cập nhật khung pháp lý sau ngày "
    "01 tháng 10 năm 2025 và đi đến giải pháp cụ thể kèm cách theo dõi.",
    "**Đóng góp của bài báo.** Về lý luận, bài báo chỉ ra rằng khâu nhận diện trước công bố là một khâu quản trị độc lập, "
    "nơi chính sách cấp trường có tác động lớn nhất khi chính sách quốc gia đã thông thoáng, và đề xuất đặt chỉ số tại "
    "các mối nối của chuỗi kết quả. Về thực tiễn, năm nhóm giải pháp và bộ chỉ số theo dõi là công cụ mà cơ sở giáo dục "
    "đại học có thể áp dụng ngay khi hợp nhất quy chế theo Luật số 93/2025/QH15 và chuẩn bị đánh giá theo Thông tư số "
    "83/2026/TT-BGDĐT. Về chính sách, cơ quan quản lý cần sớm hướng dẫn cách xác định lợi nhuận làm căn cứ thưởng tác giả "
    "và cân nhắc bổ sung chỉ số về tỷ lệ sản phẩm có khả năng bảo hộ được nộp đơn vào hệ thống dữ liệu báo cáo.",
    "**Hạn chế của nghiên cứu.** Thứ nhất, bài báo phân tích văn bản và tổng hợp lý thuyết, chưa đo lường mức độ thực thi "
    "tại nhiều trường; các khoảng trống được suy ra từ đối chiếu yêu cầu với công cụ quản trị, cần kiểm chứng bằng dữ "
    "liệu. Thứ hai, văn bản hướng dẫn chi tiết Luật số 93/2025/QH15 và hướng dẫn tính chỉ số của Thông tư số "
    "83/2026/TT-BGDĐT có thể làm thay đổi một số hàm ý. Thứ ba, bộ chỉ số chưa được xác định trọng số và chưa qua lấy ý "
    "kiến chuyên gia.",
    "**Hướng nghiên cứu tiếp theo.** Các nghiên cứu tiếp theo có thể khảo sát quy chế và thực tiễn quản trị của một nhóm "
    "cơ sở giáo dục đại học để kiểm chứng bốn khoảng trống, lấy ý kiến chuyên gia để hoàn thiện bộ chỉ số, và đánh giá "
    "tác động của cổng rà soát tại nghiệm thu bằng thiết kế trước và sau.",
]

KET_LUAN = [
    "Bài báo phân tích yêu cầu của khung thể chế mới đối với quản trị quyền sở hữu trí tuệ trong cơ sở giáo dục đại học và "
    "đề xuất giải pháp nâng cao hiệu quả.",
    "Kết quả cho thấy khung thể chế giai đoạn 2025 - 2026 đặt ra năm nhóm yêu cầu về quyền sở hữu, phân chia lợi ích, "
    "nguồn lực, đo lường và tính thống nhất của quy chế. Giữa yêu cầu và công cụ quản trị nội bộ còn bốn khoảng trống: "
    "độ trễ của quy chế, khâu nhận diện trước công bố, dòng kinh phí cho xác lập quyền và dữ liệu theo dõi. Năm nhóm giải "
    "pháp được đề xuất tương ứng, trong đó cổng rà soát khả năng bảo hộ tại nghiệm thu là giải pháp then chốt vì vừa lấp "
    "khoảng trống nhận diện, vừa tạo dữ liệu cho các chỉ số chuyển hóa.",
    "Bài báo đóng góp một cách tiếp cận đi từ yêu cầu, qua khoảng trống, đến giải pháp và chỉ số theo dõi. Hàm ý cho các "
    "cơ sở giáo dục đại học là ưu tiên hợp nhất quy chế theo Luật số 93/2025/QH15 và đặt bước rà soát khả năng bảo hộ vào "
    "quy trình nghiệm thu, trước khi đầu tư vào các thiết chế thương mại hóa phức tạp hơn.",
]

BANG_GP = dict(
    tieu_de="Các nhóm giải pháp, căn cứ pháp lý và chỉ số theo dõi",
    cot=["Giải pháp", "Khoảng trống được lấp", "Căn cứ chủ yếu", "Chỉ số theo dõi"],
    dong=[["1. Hợp nhất quy chế theo luật mới", "Độ trễ của quy chế nội bộ",
           "Điều 25, 27, 28 Luật số 93/2025/QH15; Thông tư số 83/2026/TT-BGDĐT",
           "Mức độ tương thích của quy chế; số văn bản cùng điều chỉnh"],
          ["2. Cổng rà soát tại nghiệm thu", "Khâu nhận diện trước công bố", "Khoản 3 Điều 60 Luật Sở hữu trí tuệ; "
                                                                             "Quyết định số 1624/QĐ-TTg",
           "Tỷ lệ nhận diện; tỷ lệ xác lập kịp thời"],
          ["3. Cơ chế tài chính cho xác lập quyền", "Dòng kinh phí cho xác lập quyền",
           "Điều 66 Luật số 93/2025/QH15; Điều 28 Luật số 125/2025/QH15",
           "Kinh phí xác lập, duy trì; số đơn; tỷ suất khai thác trên chi phí"],
          ["4. Kiện toàn đầu mối và dữ liệu", "Dữ liệu và đầu mối theo dõi",
           "Điều 27 Luật số 93/2025/QH15; Điều 28 Luật số 125/2025/QH15",
           "Thời gian xử lý hồ sơ; tỷ lệ khai thác; hợp đồng, nguồn thu"],
          ["5. Năng lực và văn hóa sở hữu trí tuệ", "Hỗ trợ cả bốn khoảng trống",
           "Quyết định số 1624/QĐ-TTg; Nghị định số 134/2026/NĐ-CP; Chỉ thị số 02/CT-TTg",
           "Số lớp tập huấn; động lực đổi mới sáng tạo của giảng viên"]],
    nguon="Nguồn: Nhóm tác giả đề xuất.",
    rong=[3.6, 3.4, 4.6, 3.8], can=["left", "left", "left", "left"],
)

TO_KHAI = [("h", "TỜ KHAI MINH BẠCH SỬ DỤNG AI VÀ CAM KẾT DỮ LIỆU GỐC"),
           ("p", "Tên bài báo: " + TIEU_DE + "."), ("p", "Tác giả: " + TAC_GIA + ". Đơn vị: " + DON_VI + ".")] + \
          [x for x in CS.TO_KHAI if x[0] != "h" and not (x[1] or "").startswith(("Tên bài báo", "Tác giả"))]


def dung():
    v = khung.VanBanChung()
    dem = {}

    def phan(ten, ds):
        dem[ten] = dem.get(ten, 0) + sum(len(re.sub(r"\*", "", x).split()) for x in ds)
        v.than_md(*ds)

    def bang(b):
        v.bang(b["tieu_de"], b["cot"], b["dong"], b["nguon"], b["rong"], can=b["can"], tien_to="")

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
    v.so_do("Chu trình quản lý tài sản trí tuệ trong cơ sở giáo dục đại học", os.path.join(SO_DO, "chu_trinh.png"),
            "Nguồn: Nhóm tác giả xây dựng trên cơ sở Bradley và cộng sự (2013) và Tổ chức Sở hữu trí tuệ thế giới "
            "(2020).", tien_to="", rong_cm=12.5)
    phan("Phương pháp", PHUONG_PHAP[1:])

    v.doan("h1", "4. KẾT QUẢ NGHIÊN CỨU", bold=True)
    v.doan("h2", "4.1. Năm nhóm yêu cầu của khung thể chế mới")
    phan("Kết quả", KQ_YEU_CAU[:1])
    b1 = dict(CS.BANG_1, tieu_de="Các yêu cầu của khung thể chế mới về quyền đối với kết quả nghiên cứu và hàm ý cho "
                                 "quản trị nội bộ")
    bang(b1)
    phan("Kết quả", KQ_YEU_CAU[1:5])
    bang(CS.BANG_2)
    phan("Kết quả", KQ_YEU_CAU[5:])
    v.doan("h2", "4.2. Bốn khoảng trống giữa yêu cầu thể chế và công cụ quản trị nội bộ")
    phan("Kết quả", KQ_KHOANG_TRONG)
    v.doan("h2", "4.3. Năm nhóm giải pháp nâng cao hiệu quả quản trị")
    phan("Kết quả", GIAI_PHAP[:1])
    bang(BANG_GP)
    phan("Kết quả", GIAI_PHAP[1:3])
    v.so_do("Mô hình cổng rà soát khả năng bảo hộ tại nghiệm thu đề tài", os.path.join(SO_DO, "cong_ra_soat.png"),
            "Nguồn: Nhóm tác giả đề xuất.", tien_to="", rong_cm=15.5)
    phan("Kết quả", GIAI_PHAP[3:])
    v.doan("h2", "4.4. Theo dõi hiệu quả theo chuỗi đầu vào, quá trình, đầu ra và kết quả")
    phan("Kết quả", KQ_CHI_SO[:1])
    v.so_do("Chuỗi đầu vào, quá trình, đầu ra, kết quả và bốn chỉ số chuyển hóa", os.path.join(SO_DO, "chuoi_ket_qua.png"),
            "Nguồn: Nhóm tác giả đề xuất trên cơ sở W. K. Kellogg Foundation (2004).", tien_to="", rong_cm=15.5)
    phan("Kết quả", KQ_CHI_SO[1:])

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

# -*- coding: utf-8 -*-
"""Nội dung bài báo "Quản lý sở hữu trí tuệ tại trường đại học trước yêu cầu chính sách mới".

Chỉ dùng văn bản, nghiên cứu và hướng dẫn công khai; không dùng số liệu nội bộ của đề tài.
Văn bản được đối chiếu với tệp trong VBPL/ và Tai_lieu_tham_khao_PDF/, "Tong quan tai lieu/", "Co so ly luan/";
ngày chốt cập nhật văn bản: 30 tháng 9 năm 2026.
"""

TIEU_DE = "Quản lý sở hữu trí tuệ tại trường đại học trước yêu cầu chính sách mới"
TIEU_DE_EN = "Intellectual property management in universities in response to new policy requirements"
TAC_GIA = "Nguyễn Thị Tố Uyên, Trần Đăng Bộ"
DON_VI = "Trường Đại học Thành Đô"

TOM_TAT = (
    "Bài báo phân tích các yêu cầu từ chính sách mới đối với quản lý sở hữu trí tuệ trong trường đại học và chuyển các "
    "yêu cầu đó thành hệ thống giải pháp cùng điều kiện tổ chức thực hiện. Nghiên cứu sử dụng thiết kế định tính, phân "
    "tích tài liệu và đối chiếu chính sách trên mười hai tài liệu công khai, gồm sáu văn bản của Việt Nam cập nhật đến tháng 9 năm 2026, hai hướng dẫn quốc tế và bốn công trình nghiên cứu. Các yêu cầu được trích xuất theo căn cứ, chủ thể, phạm vi áp dụng và đối chiếu với năm cấu phần quản lý: quy chế, tổ chức, "
    "quy trình, tài chính và lợi ích, đào tạo và văn hóa. Kết quả xác định năm nhóm yêu cầu, gồm định hướng chính sách, "
    "nghĩa vụ pháp lý và quyền được trao; mỗi nhóm đòi hỏi thay đổi đồng thời ở nhiều cấu phần. Bài báo đề xuất năm nhóm "
    "giải pháp, một quy trình phối hợp có khai báo, sàng lọc và xem xét bảo mật trước công bố, cùng bộ chỉ số phân biệt "
    "hoạt động, đầu ra và kết quả khai thác. Đóng góp của bài báo là cách liên kết yêu cầu chính sách với nhiệm vụ, trách "
    "nhiệm và chỉ số ở cấp trường; hệ thống giải pháp cần được tham vấn chuyên gia hoặc thí điểm trước khi áp dụng rộng.")

TU_KHOA = "Chính sách sở hữu trí tuệ; Quản lý sở hữu trí tuệ; Tài sản trí tuệ; Trường đại học"

ABSTRACT = (
    "This article analyses the requirements that new policies place on intellectual property management in "
    "universities and translates them into a set of management measures and implementation conditions. The study "
    "adopts a qualitative design based on document analysis and policy comparison of twelve publicly available "
    "documents: six Vietnamese policy directives, strategies and laws issued or amended up to September 2026, two "
    "international guidance documents and four research works. Requirements were extracted by legal basis, "
    "responsible actor, scope of application and content to be specified, and then mapped onto five management "
    "components: regulations, organisation, procedures, finance and benefit sharing, and training and intellectual "
    "property culture. The analysis identifies five groups of requirements, combining policy orientations, legal "
    "obligations and rights granted to universities; each group calls for simultaneous changes in several components. "
    "The article proposes five groups of measures, a coordination procedure with disclosure, screening and "
    "confidentiality review before publication, and an indicator set that distinguishes activities, outputs and "
    "exploitation outcomes. Its contribution lies in linking each policy requirement to tasks, responsibilities and "
    "monitoring indicators at institutional level. The proposed system has not yet been validated in practice and "
    "should be tested through expert consultation or pilot implementation before wider adoption.")
KEYWORDS = "Intellectual property policy; Intellectual property management; Intellectual assets; Universities"

# ---------------------------------------------------------------------------------------------- 1
DAT_VAN_DE = [
    "Trường đại học vừa tạo ra tri thức mới qua nghiên cứu và đào tạo, vừa là nơi hình thành nhiều loại tài sản trí tuệ "
    "như sáng chế, kiểu dáng công nghiệp, phần mềm, giáo trình, cơ sở dữ liệu và bí quyết. Giá trị của chúng chỉ được hiện thực hóa khi được nhận diện kịp thời, bảo hộ bằng công cụ phù hợp và khai thác "
    "(Siegel et al., 2007).",
    "Giai đoạn 2025 - 2026, hoạt động này chịu tác động trực tiếp của nhiều văn bản mới. Kết luận số 51-KL/TW yêu cầu "
    "chuyển từ tư duy quản lý hành chính sang kiến tạo hệ sinh thái sở hữu trí tuệ (Bộ Chính trị, 2026). Chiến lược sở "
    "hữu trí tuệ đến năm 2030 được sửa đổi, yêu cầu dùng chỉ số sở hữu trí tuệ để đánh giá cơ sở giáo dục đại học (Thủ "
    "tướng Chính phủ, 2026). Luật Khoa học, công nghệ và đổi mới sáng tạo, Luật Giáo dục đại học và Luật Sở hữu trí tuệ "
    "sửa đổi thay đổi quyền đối với kết quả nghiên cứu, quyền đăng ký và cơ chế lợi ích của tác giả (Quốc hội, 2025a, "
    "2025b; Văn phòng Quốc hội, 2026).",
    "Các văn bản này có văn bản nêu định hướng, có văn bản đặt nghĩa vụ, có văn bản trao quyền. Vấn đề đặt ra là cụ thể "
    "hóa chúng thành nội dung quản lý và cơ chế tổ chức thực hiện ở cấp trường. Bài báo trả lời ba câu hỏi: những yêu "
    "cầu chính sách nào cần được cụ thể hóa; các yêu cầu đó cần được chuyển thành nội dung quản lý, giải pháp và cơ chế "
    "phối hợp nào; và cần những điều kiện, chỉ số theo dõi nào để triển khai. Phạm vi phân tích là tài liệu công khai; "
    "bài báo không đánh giá thực trạng của một trường cụ thể.",
]

# ---------------------------------------------------------------------------------------------- 2
TONG_QUAN_21 = [
    
    "**Chính sách và quy chế sở hữu trí tuệ.** Siegel và cộng sự (2007), tổng quan nghiên cứu về văn phòng chuyển giao công nghệ ở Hoa Kỳ và châu Âu, kết luận rằng trường đại học cần xây dựng chiến lược thương mại hóa nhất quán và "
    "khả thi, trong đó chiến lược sở hữu trí tuệ phải giải quyết trước vấn đề quyền sở hữu và phạm vi tài sản. Võ (2025), "
    "bằng phân tích quy định pháp luật và so sánh quy chế của một số trường trong nước với chính sách sở hữu trí tuệ của "
    "trường đại học nước ngoài, chỉ ra rằng quy chế trong nước chủ yếu quan tâm đến chủ sở hữu và phân chia lợi nhuận, còn "
    "quyền và nghĩa vụ của các chủ thể không giữ quyền tài sản đối với tác phẩm hình thành trong trường ít được quy định. "
    "Ở cấp chính sách, Khuyến nghị (EU) 2022/2415 đề nghị bảo đảm chính sách và thực hành quản lý tài sản trí tuệ được "
    "xác định, thực hiện, chia sẻ và công khai trong mọi tổ chức tham gia khai thác tri thức (Council of the European Union, 2022). "
    "Các công trình này làm rõ vai trò của quy chế, song tiếp cận từ một nhóm quyền hoặc từ chiến lược chung.",
    "**Tổ chức bộ máy, phân công và phối hợp.** Siegel và cộng sự (2007) tổng hợp các nghiên cứu cho thấy văn phòng chuyển "
    "giao công nghệ cần đạt quy mô tới hạn mới bù đắp được chi phí cố định, từ đó đề xuất hợp tác vùng khi từng trường chưa đạt quy mô tới hạn "
    "về nghiên cứu và chuyên môn chuyển giao. Nguyễn (2025), bằng phân tích tác động của Cách mạng công nghiệp lần thứ tư, "
    "nhấn mạnh yêu cầu phối hợp đồng bộ giữa các bộ phận trong trường. Các nghiên cứu này gợi ý mô hình tổ chức cần phù hợp quy mô, nhưng chưa phân tích việc phân công ở trường chưa có văn phòng chuyên trách.",
    "**Nhận diện, bảo hộ, quản lý và khai thác.** Siegel và cộng sự (2007) xác định bản khai báo sáng chế của nhà khoa học "
    "là đầu vào then chốt của chuyển giao, trong khi nhiều nhà khoa học không khai báo. Hướng dẫn của Tổ chức Sở hữu trí "
    "tuệ thế giới (World Intellectual Property Organization, 2020) trình bày quy trình thu thập thông tin, tra cứu sáng chế "
    "và phân tích tự do hoạt động, giúp sử dụng sáng chế thuộc phạm vi công cộng và tránh xâm phạm quyền của người khác. O’Dwyer và cộng sự (2023), qua nghiên cứu trường hợp một mạng lưới hợp tác dược phẩm gồm mười doanh "
    "nghiệp và tám cơ sở học thuật, cho thấy nỗi lo rò rỉ tri thức ở giai đoạn đầu được giải quyết dần, trong đó thỏa "
    "thuận sở hữu trí tuệ hình thành ở giai đoạn gắn kết. Các nghiên cứu này mô tả từng khâu, nhưng ít đặt chúng vào một "
    "quy trình có trách nhiệm và thời hạn ở cấp trường.",
    "**Khuyến khích, phân chia lợi ích và phát triển năng lực.** Siegel và cộng sự (2007) cho rằng cần có tỷ lệ chia hợp "
    "lý cho nhà sáng chế và cần điều chỉnh hệ thống đánh giá, đãi ngộ để hoạt động thương mại hóa được ghi nhận; đồng thời "
    "nhân sự chuyển giao cần kỹ năng thương mại chứ không chỉ kỹ năng pháp lý. Khuyến nghị (EU) 2022/2415 đề nghị hệ thống "
    "khuyến khích công bằng, phát triển kỹ năng cho mọi chủ thể và thống nhất định nghĩa, chỉ số đo lường. Nguyễn (2025) "
    "nêu yêu cầu đào tạo và nâng cao nhận thức cho giảng viên, người học. Các nghiên cứu quốc tế dựa trên pháp luật khác Việt Nam, nên mô hình khuyến khích không thể chuyển nguyên trạng.",
]
TONG_QUAN_22 = [
    "Từ các tài liệu đã đọc, có thể thấy nghiên cứu quốc tế cung cấp nhiều bằng chứng về từng cấu phần quản lý, còn nghiên "
    "cứu trong nước đã bàn về quyền của các chủ thể và ứng dụng công nghệ. Phần bài báo có thể bổ sung là sự liên kết giữa "
    "các yêu cầu chính sách ban hành giai đoạn 2025 - 2026 với các cấu phần quản lý ở cấp trường, theo đó mỗi yêu cầu được "
    "gắn với chủ thể, nội dung cần cụ thể hóa, điều kiện và chỉ số theo dõi.",
    "Bài báo sử dụng khung phân tích năm cấu phần: quy chế; tổ chức; quy trình; tài chính và lợi ích; đào tạo và văn hóa "
    "sở hữu trí tuệ. Năm cấu phần này tương ứng với các nhóm vấn đề trong tổng quan và các nhóm nguyên tắc của Khuyến nghị (EU) 2022/2415.",
]

# ---------------------------------------------------------------------------------------------- 3
PHUONG_PHAP = [
    "Nghiên cứu sử dụng thiết kế định tính, kết hợp phân tích tài liệu và đối chiếu chính sách. Tập tài liệu gồm ba loại: "
    "văn bản chính sách, pháp luật của Việt Nam; văn bản hướng dẫn quốc tế về quản lý tài sản trí tuệ; và nghiên cứu học "
    "thuật về quản lý sở hữu trí tuệ trong trường đại học.",
    "Tài liệu được chọn khi liên quan trực tiếp đến quản lý sở hữu trí tuệ trong trường đại học và xác định được nguồn, thời điểm, phạm vi áp dụng. Tập phân tích gồm mười hai tài liệu: sáu văn "
    "bản của Việt Nam, gồm Kết luận số 51-KL/TW, Quyết định số 1068/QĐ-TTg và Quyết định số 1624/QĐ-TTg sửa đổi, bổ sung "
    "quyết định này, Luật số 93/2025/QH15, Luật số 125/2025/QH15 và Văn bản hợp nhất số 67/VBHN-VPQH của Luật Sở hữu trí "
    "tuệ; hai văn bản quốc tế là Khuyến nghị (EU) 2022/2415 và hướng dẫn tra cứu sáng chế thuộc phạm vi công cộng của Tổ "
    "chức Sở hữu trí tuệ thế giới; bốn công trình nghiên cứu được trình bày tại Mục 2. Văn bản của Việt Nam tập trung vào giai đoạn 2025 - 2026; Quyết định số 1068/QĐ-TTg được dùng để xác định nội dung còn hiệu lực của Chiến lược. Ngày chốt cập nhật văn bản là 30 tháng 9 năm 2026.",
    "Mỗi văn bản được đọc toàn văn và trích xuất vào bảng gồm năm trường: căn cứ; yêu cầu; chủ thể và phạm vi áp dụng; nội dung quản lý cần cụ thể hóa; điều kiện thực hiện và chỉ số có thể theo dõi. Mỗi "
    "yêu cầu được phân loại thành định hướng chính sách, nghĩa vụ pháp lý hoặc quyền được trao cho trường đại học. Phần "
    "đề xuất của nhóm tác giả được trình bày tách biệt với nội dung của văn bản.",
    "Các yêu cầu sau đó được đối chiếu với năm cấu phần quản lý để xác định mỗi yêu cầu đòi hỏi thay đổi ở cấu phần nào. "
    "Hệ thống giải pháp được phát triển từ một đề tài cấp cơ sở của nhóm tác giả; bài báo chỉ kế thừa phần lập luận dựa trên tài liệu công khai, không sử dụng số liệu nội bộ. Nghiên cứu không thực hiện khảo sát hay phỏng vấn, nên kết quả là đề xuất có căn cứ văn bản, chưa phải mô "
    "hình đã được kiểm chứng.",
]

# ---------------------------------------------------------------------------------------------- 4.1
KQ_41 = [
    "Kết quả trích xuất cho thấy các yêu cầu chính sách hội tụ thành năm nhóm, được tổng hợp tại Bảng 1.",
]
KQ_41_SAU = [
    "**Thứ nhất, gắn sở hữu trí tuệ với chiến lược và đánh giá hoạt động.** Đây là định hướng chính sách (Bộ Chính trị, 2026; Thủ tướng Chính phủ, 2026), đòi hỏi mục tiêu sở hữu trí tuệ trong kế hoạch của trường và bộ chỉ số có định nghĩa rõ.",
    "**Thứ hai, gắn nhận diện, bảo hộ với nghiên cứu và công bố.** Chiến lược sửa đổi yêu cầu xác định trước đối tượng "
    "quyền cần đạt với kết quả sử dụng ngân sách nhà nước và đăng ký bảo hộ đồng thời với công bố kết quả có tính ứng dụng "
    "cao ở khối kỹ thuật, công nghệ (Thủ tướng Chính phủ, 2026). Luật Sở hữu trí tuệ chỉ cho phép nộp đơn sáng chế trong "
    "mười hai tháng và đơn kiểu dáng trong sáu tháng kể từ khi người có quyền đăng ký bộc lộ công khai (Văn phòng Quốc "
    "hội, 2026). Vì vậy, việc nhận diện phải diễn ra trước công bố.",
    "**Thứ ba, nâng cao năng lực hỗ trợ, khai thác.** Chiến lược định hướng phát triển trung tâm chuyển giao, tư vấn, định giá trong cơ sở giáo dục đại học (Thủ tướng Chính phủ, 2019, 2026); Luật Giáo dục đại học trao quyền "
    "thành lập doanh nghiệp quản lý tài sản trí tuệ (Quốc hội, 2025b). Đây là quyền, không phải nghĩa vụ phải lập thêm tổ "
    "chức.",
    "**Thứ tư, hoàn thiện cơ chế lợi ích, dữ liệu và trách nhiệm giải trình.** Điều 28 Luật số 93/2025/QH15 phân biệt theo "
    "nguồn hình thành: với phần không sử dụng ngân sách nhà nước, chủ sở hữu tự quyết định việc xử lý lợi nhuận, kể cả "
    "thưởng cho tác giả; với phần sử dụng ngân sách nhà nước, thưởng cho tác giả tối thiểu 30% lợi nhuận sau thuế, hoặc 30% "
    "giá trị kết quả khi góp vốn. Điều 73 giữ chế độ cũ cho nhiệm vụ phê duyệt trước ngày 01 tháng 10 năm 2025, trừ lợi "
    "nhuận chưa phân chia từ sáng chế, kiểu dáng công nghiệp, thiết kế bố trí, giống cây trồng đã được cấp văn bằng của "
    "nhiệm vụ giao từ ngày 01 tháng 01 năm 2023 (Quốc hội, 2025a). Bên cạnh khoản thưởng, chủ sở hữu sáng chế, kiểu dáng, thiết kế bố trí vẫn phải trả thù lao theo Điều 135 Luật Sở hữu trí tuệ, theo thỏa thuận hoặc mức mặc định tính trên số tiền trước thuế (Văn phòng Quốc hội, 2026). Các luật cũng yêu cầu công khai, báo cáo kết quả thương mại hóa và hoạt động khoa học hằng năm (Quốc hội, 2025a, 2025b).",
    "**Thứ năm, phát triển đào tạo và văn hóa sở hữu trí tuệ,** gồm chương trình đào tạo, nghiên cứu đưa sở hữu trí tuệ "
    "thành nội dung học bắt buộc, xử lý đạo văn và tuân thủ liêm chính trong nghiên cứu (Bộ Chính trị, 2026; Quốc hội, "
    "2025b; Thủ tướng Chính phủ, 2026).",
    "Mỗi nhóm yêu cầu không gắn với một cấu phần duy nhất: yêu cầu đăng ký đồng thời với công bố chỉ thực hiện được khi có nghĩa vụ khai báo, bước sàng lọc trước công bố, người chịu trách nhiệm, kinh phí nộp đơn và người nghiên cứu hiểu thời hạn bộc lộ. Đây là cơ sở để thiết kế giải pháp đồng bộ.",
]

BANG_1 = dict(
    tieu_de="Yêu cầu chính sách và nội dung quản lý cần cụ thể hóa trong trường đại học",
    cot=["Căn cứ", "Yêu cầu", "Chủ thể, phạm vi áp dụng", "Nội dung cần cụ thể hóa"],
    dong=[
        ["**Nhóm 1. Chiến lược và đánh giá**", "", "", ""],
        ["Kết luận số 51-KL/TW, mục 2.1", "Sở hữu trí tuệ là cấu phần của chiến lược khoa học, công nghệ, nhân lực", "Định hướng chính sách, toàn hệ thống",
         "Mục tiêu sở hữu trí tuệ trong kế hoạch của trường"],
        ["Quyết định số 1624/QĐ-TTg, điểm b khoản 4 Mục III",
         "Dùng chỉ số sở hữu trí tuệ để đánh giá hiệu quả hoạt động",
         "Định hướng; cơ sở giáo dục đại học, viện, doanh nghiệp",
         "Bộ chỉ số, định nghĩa, nguồn dữ liệu"],
        ["**Nhóm 2. Nhận diện, bảo hộ gắn với nghiên cứu, công bố**", "", "", ""],
        ["Quyết định số 1624/QĐ-TTg, điểm b khoản 4 Mục III",
         "Xác định đối tượng quyền cần đạt; đăng ký bảo hộ đồng thời với công bố",
         "Kết quả dùng ngân sách nhà nước; khối kỹ thuật, công nghệ",
         "Khai báo từ khâu đề xuất; sàng lọc trước công bố"],
        ["Luật Sở hữu trí tuệ, khoản 3 Điều 60, khoản 4 Điều 65, điểm c khoản 1 Điều 86",
         "Ân hạn 12 tháng với sáng chế, 6 tháng với kiểu dáng; quyền đăng ký của tổ chức được giao kết quả", "Nghĩa vụ pháp lý, điều kiện bảo hộ", "Lịch công bố; ghi nhận ngày bộc lộ; thẩm quyền nộp đơn"],
        ["**Nhóm 3. Năng lực hỗ trợ, khai thác**", "", "", ""],
        ["Quyết định số 1068/QĐ-TTg, khoản 5 Mục III; Quyết định số 1624/QĐ-TTg, khoản 5 Mục II, khoản 6 Mục III", "Trung tâm chuyển giao, tư vấn, định giá; doanh nghiệp khai thác; thí điểm định giá ít nhất 100 quyền", "Định hướng; cơ sở giáo dục đại học, viện",
         "Đầu mối hỗ trợ; phương án tổ chức phù hợp quy mô"],
        ["Luật số 125/2025/QH15, khoản 1, điểm d khoản 2 Điều 28",
         "Thành lập doanh nghiệp quản lý tài sản trí tuệ; định giá, khai thác, góp vốn, phân chia lợi ích",
         "Quyền được trao", "Thẩm quyền, điều kiện lựa chọn hình thức tổ chức"],
        ["**Nhóm 4. Lợi ích, dữ liệu, trách nhiệm giải trình**", "", "", ""],
        ["Luật số 93/2025/QH15, khoản 2 Điều 25, Điều 27, Điều 28, khoản 3, khoản 7 Điều 73",
         "Tự động giao quyền; tự quyết thương mại hóa; thưởng tác giả tối thiểu 30% lợi nhuận sau "
         "thuế với phần ngân sách; công khai, báo cáo", "Nghĩa vụ pháp lý theo nguồn kinh phí, thời điểm giao nhiệm vụ", "Cơ chế lợi ích theo nguồn, thời điểm, đối tượng; báo cáo"],
        ["Luật Sở hữu trí tuệ, khoản 1 Điều 135", "Thù lao cho tác giả sáng chế, kiểu dáng, thiết kế bố trí theo thỏa "
         "thuận; mặc định 10% hoặc 15%", "Chủ sở hữu văn bằng; nghĩa vụ pháp lý", "Tách thù lao khỏi thưởng; xác định "
         "thỏa thuận"],
        ["Luật số 93/2025/QH15, Điều 66; Luật số 125/2025/QH15, khoản 3 Điều 28",
         "Quỹ phát triển khoa học và công nghệ được chi cho đăng ký, bảo hộ; công khai hằng năm",
         "Tổ chức có quỹ; cơ sở giáo dục đại học", "Dòng kinh phí cho xác lập quyền; danh mục số"],
        ["**Nhóm 5. Đào tạo và văn hóa**", "", "", ""],
        ["Quyết định số 1624/QĐ-TTg, điểm b khoản 8 Mục III", "Chương trình đào tạo sở hữu trí tuệ; nghiên cứu đưa thành "
         "nội dung học bắt buộc", "Định hướng; cơ sở giáo dục đại học", "Học phần, bồi dưỡng theo vai trò"],
        ["Kết luận số 51-KL/TW, mục 2.2; Luật số 125/2025/QH15, điểm a khoản 3 Điều 28", "Văn hóa tôn trọng sáng tạo; xử lý đạo văn; liêm chính", "Toàn xã hội; cơ sở giáo dục đại học",
         "Quy định liêm chính, quyền của bên thứ ba"],
    ],
    nguon="Nguồn: Nhóm tác giả tổng hợp từ Bộ Chính trị (2026), Thủ tướng Chính phủ (2019, 2026), Quốc hội (2025a, "
          "2025b) và Văn phòng Quốc hội (2026).",
    rong=[4.2, 5.0, 3.8, 4.0],
)

# ---------------------------------------------------------------------------------------------- 4.2
KQ_42_MO = [
    "Năm nhóm giải pháp dưới đây tương ứng với năm cấu phần quản lý; mỗi nhóm được nối với yêu cầu đã xác định tại Mục "
    "4.1 và trình bày theo mục tiêu, nội dung, chủ thể và điều kiện.",
]
GP1 = [
    "Giải pháp này đáp ứng nhóm yêu cầu thứ tư và tạo căn cứ cho các nhóm còn lại. Quy chế cần bao quát cả đối tượng được bảo hộ theo đăng ký, như "
    "sáng chế, kiểu dáng, nhãn hiệu, lẫn đối tượng phát sinh quyền tự động hoặc được bảo vệ bằng biện pháp bảo mật, như "
    "giáo trình, phần mềm, cơ sở dữ liệu, bí mật kinh doanh; quy định nghĩa vụ khai báo, bảo mật trước công bố và quyền "
    "công bố; quy định quyền của tác giả và của các chủ thể không giữ quyền tài sản, theo hướng Võ (2025) đã nêu; quy định "
    "giải quyết tranh chấp nội bộ; và nêu thứ tự áp dụng với quy chế nghiên cứu, quy chế chi tiêu.",
    "Phân chia lợi ích cần được thiết kế theo nguồn hình thành tài sản, thời điểm giao nhiệm vụ và loại đối tượng. Với "
    "phần kết quả sử dụng ngân sách nhà nước thuộc phạm vi Điều 28 Luật số 93/2025/QH15, quy chế không được đặt mức thấp "
    "hơn 30% lợi nhuận sau thuế dành cho tác giả; với nhiệm vụ phê duyệt trước ngày 01 tháng 10 năm 2025, cần dẫn chiếu "
    "văn bản tại thời điểm phê duyệt, trừ trường hợp tại khoản 7 Điều 73; với tài sản không sử dụng ngân sách nhà nước, "
    "trường tự quyết định tỷ lệ (Quốc hội, 2025a). Thù lao theo Điều 135 Luật Sở hữu trí tuệ là khoản riêng, có cơ sở tính "
    "khác, nên quy chế cần nêu rõ quy chế hoặc hợp đồng có được coi là thỏa thuận về thù lao hay không (Văn phòng Quốc "
    "hội, 2026); nhuận bút cho sách, giáo trình là khoản thứ ba. Tỷ lệ cụ thể nên được mô phỏng trên một số tình huống "
    "chuyển giao trước khi ban hành; bài báo không đề xuất một tỷ lệ chung. Chủ trì là bộ phận pháp chế, phối hợp với đơn "
    "vị quản lý khoa học và tài chính.",
]

GP2 = [
    "Giải pháp này đáp ứng nhóm yêu cầu thứ ba. Nhóm tác giả đề xuất phân công theo bốn chức năng thay vì mặc định một "
    "mô hình: quản lý khoa học, gồm tiếp nhận khai báo và theo dõi danh mục; pháp lý, gồm thẩm định quyền, soạn hồ sơ, làm "
    "việc với cơ quan đăng ký; tài chính, gồm dự toán và chi trả; hỗ trợ khai thác và kết nối doanh nghiệp, gồm tìm đối "
    "tác, đàm phán, định giá.",
    "Văn phòng chuyên trách chỉ hiệu quả khi lượng sáng chế đủ "
    "lớn, và trường chưa đạt quy mô tới hạn có thể hợp tác vùng (Siegel et al., 2007). Vì vậy, trường có ít kết quả có "
    "thể bảo hộ nên giao nhiệm vụ kiêm nhiệm cho các phòng hiện có và thuê tổ chức đại diện sở hữu công nghiệp theo hồ sơ; "
    "khi số hồ sơ tăng mới xem xét vị trí chuyên trách. Trung tâm tư vấn, định giá hay doanh nghiệp quản lý tài sản trí "
    "tuệ chỉ nên đặt ra khi danh mục đã có tài sản sẵn sàng khai thác, có nhu cầu định giá thực tế và có phương án tài "
    "chính vận hành. Điều kiện là quy chế phối hợp và báo cáo khối lượng hồ sơ định kỳ.",
]

GP3 = [
    "Giải pháp này đáp ứng nhóm yêu cầu thứ hai. Mục tiêu là mỗi kết quả có khả năng bảo hộ được nhận diện, sàng lọc và xem xét bảo mật trước khi công bố, với cách xử lý phù hợp loại tài sản, theo tám khâu tại Hình 1.",
]
GP3_SAU = [
    "Khâu 1, người nghiên cứu khai báo đối tượng dự kiến, chủ thể quyền, nguồn kinh phí, lịch công bố, tình trạng bảo mật "
    "và nhu cầu hỗ trợ ngay từ thuyết minh. Khâu 2, đầu mối quản lý khoa học sàng lọc và phân nhánh: sáng chế, giải pháp "
    "hữu ích cần giữ bí mật đến khi nộp đơn; kiểu dáng cần nộp đơn trước khi trưng bày; nhãn hiệu cần tra cứu trước khi sử "
    "dụng; quyền tác giả phát sinh khi tác phẩm được định hình nên chỉ cần ghi nhận và đăng ký khi cần chứng cứ; bí mật "
    "kinh doanh được bảo vệ bằng biện pháp bảo mật (Văn phòng Quốc hội, 2026). Khâu 3, tra cứu và đánh giá khả năng bảo "
    "hộ; tra cứu sáng chế và phân tích tự do hoạt động còn giúp tránh xâm phạm quyền của bên thứ ba (World Intellectual "
    "Property Organization, 2020).",
    "Khâu 4, trước khi gửi bài báo, báo cáo hội thảo hoặc đưa sản phẩm vào thử nghiệm mở, đầu mối quyết định công bố, trì "
    "hoãn, nộp đơn trước hoặc ký cam kết bảo mật, và ghi nhận ngày bộc lộ nếu đã xảy ra. Khâu 5, nghiệm thu là một điểm "
    "kiểm tra trong quá trình, không phải thời điểm đầu tiên kết quả được xem xét. Khâu 6 đến Khâu 8 gồm quyết định xác lập "
    "quyền kèm dự toán, nộp đơn hoặc áp dụng biện pháp bảo mật, theo dõi, duy trì, và khai thác theo Giải pháp 1. Với kết "
    "quả hợp tác doanh nghiệp, thỏa thuận sở hữu trí tuệ cần được xác lập trong giai đoạn gắn kết, như O’Dwyer và cộng sự "
    "(2023) mô tả.",
]

GP4 = [
    "Giải pháp này đáp ứng nhóm yêu cầu thứ tư về nguồn lực và lợi ích. Mục tiêu là chi phí tra cứu, tư vấn, đăng ký, duy "
    "trì và khai thác có nguồn, có người đề xuất chi và có thời hạn. Quỹ phát triển khoa học và công nghệ của tổ chức được "
    "chi cho đăng ký, bảo hộ, quản lý, khai thác quyền sở hữu trí tuệ, và cơ sở giáo dục đại học có trách nhiệm thành lập, "
    "vận hành quỹ này (Quốc hội, 2025a, 2025b). Trên cơ sở đó, trường có thể lập dòng dự toán riêng cho chi phí xác lập "
    "quyền, tính từ số hồ sơ dự kiến và mức phí theo quy định, kèm quy định tạm ứng khi đơn được nộp sau khi đề tài đã "
    "quyết toán.",
    "Cơ chế khuyến khích nên ghi nhận đóng góp ở nhiều mốc, như khi đơn được chấp nhận hợp lệ và khi văn bằng được cấp, "
    "và cần được cân đối với chế độ dành cho công bố, vì hoạt động thương mại hóa cần được ghi nhận trong hệ thống đánh "
    "giá, đãi ngộ (Siegel et al., 2007). Khoản khuyến khích nội bộ này tách biệt với thưởng khi thương mại hóa và thù lao "
    "theo luật; lợi ích khi khai thác phân biệt theo nguồn kinh phí và loại giao dịch như chuyển nhượng, chuyển giao quyền "
    "sử dụng, tự khai thác, góp vốn. Mức cụ thể phụ thuộc khả năng tài chính của từng trường.",
]

GP5 = [
    "Giải pháp này đáp ứng nhóm yêu cầu thứ năm. Đào tạo nên phân theo vai trò: giảng viên cần kỹ năng nhận diện đối "
    "tượng bảo hộ, khai báo và bộc lộ an toàn; người học cần hiểu quyền tác giả, liêm chính và quyền của bên thứ ba; nhân "
    "sự quản lý cần kỹ năng tra cứu, sử dụng thông tin sở hữu công nghiệp và hỗ trợ khai thác, vì nhân sự chuyển giao cần "
    "kỹ năng thương mại bên cạnh kỹ năng pháp lý (Siegel et al., 2007). Nội dung tra cứu có thể dựa trên hướng dẫn công "
    "khai của Tổ chức Sở hữu trí tuệ thế giới (World Intellectual Property Organization, 2020). Cẩm nang ngắn, đầu mối giải "
    "đáp và quy định liêm chính được phổ biến giúp văn hóa tôn trọng sáng tạo trở thành chuẩn mực hằng ngày (Bộ Chính trị, "
    "2026); học phần sở hữu trí tuệ có thể được đưa vào chương trình theo lộ trình (Thủ tướng Chính phủ, 2026).",
]

# ---------------------------------------------------------------------------------------------- 4.3
KQ_43 = [
    "Hệ thống giải pháp nên được triển khai theo ba bước. Bước rà soát và chuẩn bị gồm đối chiếu quy chế hiện hành với "
    "Bảng 1, giao đầu mối và hoàn thiện biểu mẫu khai báo. Bước triển khai có lựa chọn áp dụng quy trình cho một đơn vị "
    "hoặc một nhóm đề tài có nhiều kết quả có khả năng bảo hộ, đo thời gian và khó khăn phát sinh. Bước đánh giá và hoàn "
    "thiện điều chỉnh quy chế, biểu mẫu, dự toán theo kết quả đo trước khi áp dụng toàn trường.",
    "Bảng 2 tổng hợp điều kiện thực hiện và chỉ số theo dõi. Chỉ số được chia thành ba loại: chỉ số hoạt động phản ánh "
    "việc quy trình có được vận hành; chỉ số đầu ra phản ánh đơn và văn bằng; chỉ số kết quả khai thác phản ánh hợp đồng, "
    "nguồn thu và lợi ích được chi trả. Cách phân loại này phù hợp với khuyến nghị thống nhất định nghĩa, chỉ số cho các "
    "kênh khai thác tri thức (Council of the European Union, 2022). Bài báo không đặt mức đích chung, vì mức đích phụ thuộc "
    "quy mô nghiên cứu và lĩnh vực của từng trường.",
]
KQ_43_SAU = [
    "Do văn bằng chỉ được cấp sau các giai đoạn thẩm định, chỉ số đầu ra và kết quả khai thác phản ánh tác động chậm hơn "
    "chỉ số hoạt động; vì vậy, năm đầu triển khai nên được đánh giá chủ yếu bằng chỉ số hoạt động.",
]

BANG_2 = dict(
    tieu_de="Điều kiện triển khai và chỉ số theo dõi hệ thống giải pháp",
    cot=["Nhóm giải pháp", "Điều kiện thực hiện", "Chỉ số", "Nguồn minh chứng"],
    dong=[
        ["1. Quy chế và trách nhiệm", "Rà soát văn bản nội bộ; đối chiếu pháp luật; tham vấn giảng viên", "**Hoạt động:** quy chế sửa đổi được ban hành; tỷ lệ tình huống chuyển giao xác định được "
         "quy định áp dụng", "Quyết định ban hành; bảng thứ tự áp dụng"],
        ["2. Đầu mối và phối hợp", "Phân công bốn chức năng; quy chế phối hợp; năng lực nhân sự hoặc dịch vụ thuê ngoài",
         "**Hoạt động:** thời gian xử lý trung bình từ khai báo đến quyết định; số hồ sơ do mỗi đầu mối xử lý",
         "Sổ theo dõi hồ sơ; biên bản giao ban"],
        ["3. Quy trình", "Phiếu khai báo, phiếu sàng lọc; công cụ tra cứu; danh mục số theo trạng thái pháp lý",
         "**Hoạt động:** tỷ lệ kết quả được khai báo và sàng lọc trước công bố; tỷ lệ đề tài nghiệm thu có phiếu xác nhận. "
         "**Đầu ra:** số đơn theo loại đối tượng", "Phiếu khai báo; danh mục số; thông báo của cơ quan đăng ký"],
        ["4. Kinh phí và khuyến khích", "Dòng dự toán cho xác lập quyền; quy định tạm ứng; cơ chế lợi ích theo nguồn kinh "
         "phí", "**Đầu ra:** số văn bằng, giấy chứng nhận được cấp, tách theo nguồn hình thành. **Kết quả khai thác:** số "
         "hợp đồng; nguồn thu; khoản chi trả cho tác giả tách theo thưởng, thù lao, nhuận bút",
         "Văn bằng; hợp đồng; sổ kế toán"],
        ["5. Năng lực và văn hóa", "Chương trình đào tạo theo vai trò; cẩm nang; đầu mối giải đáp",
         "**Hoạt động:** số lượt bồi dưỡng theo nhóm đối tượng; tỷ lệ nhân sự mới được bồi dưỡng",
         "Danh sách tập huấn; đề cương học phần"],
    ],
    nguon="Nguồn: Nhóm tác giả đề xuất.",
    rong=[3.2, 4.4, 6.0, 3.4],
)

# ---------------------------------------------------------------------------------------------- 5
BAN_LUAN = [
    "Kết quả phân tích cho thấy năm nhóm giải pháp phụ thuộc lẫn nhau. Quy trình sàng lọc trước công bố chỉ có hiệu lực "
    "khi quy chế đặt nghĩa vụ khai báo và bảo mật; nghĩa vụ đó chỉ được thực hiện khi có đầu mối tiếp nhận trong thời hạn "
    "hợp lý; quyết định nộp đơn chỉ thành hiện thực khi có dự toán và người đề xuất chi; còn người nghiên cứu chỉ khai báo "
    "khi hiểu thời hạn bộc lộ và thấy đóng góp của mình được ghi nhận. Vì vậy, chỉ bổ sung biểu mẫu khó đáp ứng yêu cầu đăng ký đồng thời với công bố. Nhận định này phù hợp với kết luận của Siegel và cộng sự (2007) rằng thương mại hóa đòi hỏi một chiến lược nhất quán thay vì các "
    "biện pháp đơn lẻ.",
    "Đối chiếu với các nghiên cứu đã tổng quan, hệ thống đề xuất kế thừa ba điểm. Thứ nhất, việc đặt khai báo làm đầu vào "
    "của quy trình dựa trên nhận định rằng bản khai báo là đầu vào then chốt của chuyển giao (Siegel et al., 2007). Thứ "
    "hai, việc đưa tra cứu sáng chế và phân tích tự do hoạt động vào quy trình và chương trình đào tạo dựa trên hướng dẫn "
    "của Tổ chức Sở hữu trí tuệ thế giới (World Intellectual Property Organization, 2020). Thứ ba, việc xác lập thỏa thuận "
    "sở hữu trí tuệ trong giai đoạn gắn kết của hợp tác phù hợp với kết quả của O’Dwyer và cộng sự (2023). Điểm khác là các "
    "nghiên cứu quốc tế chủ yếu khảo sát trường có văn phòng chuyển giao chuyên trách, trong khi bài báo đề xuất phân công "
    "theo chức năng, cho phép trường nhỏ dùng nhân sự kiêm nhiệm và dịch vụ thuê ngoài. Điểm khác thứ hai là cơ chế lợi ích: ở Việt Nam, trường phải đồng thời áp dụng mức thưởng tối thiểu theo nguồn kinh phí và thù lao theo Luật Sở hữu trí tuệ, nên cần tách các khoản này ngay trong quy chế.",
    "Đóng góp của bài báo nằm ở cách liên kết: Bảng 1 và Bảng 2 chuyển yêu cầu chính sách thành nội dung quản lý, chủ thể, điều kiện và chỉ số. Cách làm này giúp trường "
    "phân biệt ba loại yêu cầu: nghĩa vụ phải tuân thủ, như mức thưởng tối thiểu hay thời hạn nộp đơn; định hướng cần cụ "
    "thể hóa, như dùng chỉ số sở hữu trí tuệ trong đánh giá; và quyền có thể lựa chọn, như thành lập doanh nghiệp quản lý "
    "tài sản trí tuệ. Sự phân biệt này hạn chế rủi ro bỏ sót nghĩa vụ pháp lý hoặc đầu tư vào bộ máy khi chưa có nhu cầu, và phù hợp với tinh thần của Khuyến nghị (EU) 2022/2415 về việc công khai chính sách quản lý tài sản "
    "trí tuệ và thống nhất chỉ số đo lường (Council of the European Union, 2022).",
    "Điều kiện áp dụng khác nhau theo loại hình trường, nguồn lực và đặc điểm tài sản. Trường khối kỹ thuật, công nghệ, y dược cần tập trung vào tra cứu, bảo mật trước công bố và kinh phí nộp đơn. Trường khối kinh tế, xã hội, ngôn ngữ chủ yếu tạo ra tác phẩm, nên trọng tâm là quyền của các chủ thể, ghi nhận tác phẩm và liêm chính, như Võ (2025) đã phân tích. Trường có nhiều nhiệm vụ sử "
    "dụng ngân sách nhà nước cần chú ý quy định chuyển tiếp tại Điều 73 Luật số 93/2025/QH15, vì cùng một trường có thể "
    "đồng thời có nhiệm vụ chịu hai chế độ khác nhau. Trường có nguồn lực hạn chế có thể ưu tiên Giải pháp 1, 3 và 5, vốn chủ yếu là điều chỉnh văn bản, quy trình và đào tạo.",
    "Bài báo có một số giới hạn. Nghiên cứu dựa trên tài liệu công khai, nên chưa phản ánh cách các trường đang thực hiện "
    "các văn bản mới; hệ thống giải pháp chưa được kiểm chứng qua triển khai, nên chưa thể kết luận về hiệu quả; và kết quả "
    "không đại diện cho thực trạng của mọi trường. Văn bản quy định chi tiết Điều 28 Luật số 93/2025/QH15 có thể làm thay đổi cách tính lợi ích trình bày ở đây. Nghiên cứu tiếp theo có thể tham vấn chuyên gia quản lý khoa học, pháp chế, tài chính, hoặc thí điểm quy trình tại một số đơn vị và dùng Bảng 2 để đo kết quả trước và sau khi áp dụng.",
]

# ---------------------------------------------------------------------------------------------- 6
KET_LUAN = [
    "Bài báo phân tích các yêu cầu từ chính sách mới đối với quản lý sở hữu trí tuệ trong trường đại học và chuyển chúng "
    "thành hệ thống giải pháp cùng điều kiện tổ chức thực hiện.",
    "Kết quả cho thấy các văn bản ban hành hoặc sửa đổi giai đoạn 2025 - 2026 đặt ra năm nhóm yêu cầu: gắn sở hữu trí tuệ "
    "với chiến lược và đánh giá; gắn nhận diện, bảo hộ với nghiên cứu và công bố; nâng cao năng lực hỗ trợ, khai thác; "
    "hoàn thiện cơ chế lợi ích, dữ liệu và trách nhiệm giải trình; phát triển đào tạo và văn hóa sở hữu trí tuệ. Các yêu "
    "cầu này gồm cả định hướng chính sách, nghĩa vụ pháp lý và quyền được trao, và mỗi nhóm đòi hỏi thay đổi đồng thời ở "
    "nhiều cấu phần quản lý. Phân tích này ủng hộ luận điểm định hướng rằng yêu cầu chính sách cần được cụ thể hóa đồng bộ "
    "trong quy chế, tổ chức, quy trình, nguồn lực và năng lực của trường; mức độ hiệu quả của cách làm đó còn cần được "
    "kiểm chứng.",
    "Đóng góp của bài báo là một cách liên kết từng yêu cầu chính sách với nhiệm vụ, trách nhiệm và chỉ số theo dõi ở cấp "
    "trường. Về hàm ý tổ chức thực hiện, mỗi trường nên bắt đầu bằng việc phân loại yêu cầu theo nghĩa vụ, định hướng và "
    "quyền; ưu tiên hoàn thiện quy chế, quy trình sàng lọc trước công bố và đào tạo; lựa chọn hình thức tổ chức theo quy mô "
    "hồ sơ thực tế; và đánh giá năm đầu triển khai bằng chỉ số hoạt động trước khi kỳ vọng vào chỉ số khai thác.",
]

TAI_LIEU = [
    "Bộ Chính trị. (2026). *Kết luận số 51-KL/TW ngày 17 tháng 6 năm 2026 về đẩy mạnh công tác sở hữu trí tuệ phục vụ "
    "phát triển kinh tế - xã hội trong tình hình mới*.",
    "Council of the European Union. (2022). Council Recommendation (EU) 2022/2415 of 2 December 2022 on the guiding "
    "principles for knowledge valorisation. *Official Journal of the European Union, L 317*, 141-148.",
    "Nguyễn, M. H. T. (2025). Quản trị tài sản trí tuệ tại các cơ sở giáo dục đại học: Cơ hội và thách thức trong bối "
    "cảnh cuộc Cách mạng Công nghiệp 4.0. *Tạp chí Khoa học Trường Đại học Sư phạm Thành phố Hồ Chí Minh, 22*(1), 123-131. "
    "https://doi.org/10.54607/hcmue.js.22.1.4287(2025)",
    "O’Dwyer, M., Filieri, R., & O’Malley, L. (2023). Establishing successful university-industry collaborations: Barriers "
    "and enablers deconstructed. *The Journal of Technology Transfer, 48*(3), 900-931. "
    "https://doi.org/10.1007/s10961-022-09932-2",
    "Quốc hội. (2025a). *Luật Khoa học, công nghệ và đổi mới sáng tạo số 93/2025/QH15 ngày 27 tháng 6 năm 2025*.",
    "Quốc hội. (2025b). *Luật Giáo dục đại học số 125/2025/QH15 ngày 10 tháng 12 năm 2025*.",
    "Siegel, D. S., Veugelers, R., & Wright, M. (2007). Technology transfer offices and commercialization of university "
    "intellectual property: Performance and policy implications. *Oxford Review of Economic Policy, 23*(4), 640-660. "
    "https://doi.org/10.1093/oxrep/grm036",
    "Thủ tướng Chính phủ. (2019). *Quyết định số 1068/QĐ-TTg ngày 22 tháng 8 năm 2019 phê duyệt Chiến lược sở hữu trí tuệ "
    "đến năm 2030*.",
    "Thủ tướng Chính phủ. (2026). *Quyết định số 1624/QĐ-TTg ngày 21 tháng 8 năm 2026 sửa đổi, bổ sung một số điều của "
    "Quyết định số 1068/QĐ-TTg ngày 22 tháng 8 năm 2019 phê duyệt Chiến lược sở hữu trí tuệ đến năm 2030*.",
    "Văn phòng Quốc hội. (2026). *Văn bản hợp nhất số 67/VBHN-VPQH ngày 23 tháng 3 năm 2026 hợp nhất Luật Sở hữu trí "
    "tuệ*.",
    "Võ, N. H. P. (2025). Quyền của chủ thể không giữ quyền tài sản đối với các tác phẩm hình thành trong nhà trường, kinh "
    "nghiệm quốc tế để hoàn thiện chính sách sở hữu trí tuệ của các trường đại học tại Việt Nam. *Tạp chí Khoa học Trường Đại "
    "học Mở Hà Nội*, (129), 75-86. https://doi.org/10.59266/houjs.2025.606",
    "World Intellectual Property Organization. (2020). *Identifying inventions in the public domain: A guide for "
    "inventors and entrepreneurs*. World Intellectual Property Organization.",
]

# Khóa nhận diện trích dẫn trong bài, dùng để kiểm tra khớp hai chiều với TAI_LIEU (cùng thứ tự)
KHOA_TRICH = ["Bộ Chính trị, 2026", "Council of the European Union, 2022", "Nguyễn (2025)", "O’Dwyer và cộng sự (2023)",
              "Quốc hội, 2025a", "2025b", "Siegel", "Thủ tướng Chính phủ, 2019", "Thủ tướng Chính phủ, 2026",
              "Văn phòng Quốc hội, 2026", "Võ (2025)", "World Intellectual Property Organization, 2020"]

TO_KHAI_GHI_CHU = (
    "Ghi chú cho nhóm tác giả khi điền tờ khai: bản thảo được soạn với sự hỗ trợ của công cụ trí tuệ nhân tạo Claude "
    "của Anthropic trong các việc: tổ chức cấu trúc theo khung do chủ nhiệm đề tài xây dựng, trích xuất điều khoản từ văn "
    "bản pháp luật có trong hồ sơ, biên soạn câu chữ, định dạng tài liệu tham khảo và dựng tệp Word. Nhóm tác giả cần "
    "đọc lại toàn bộ nội dung, đối chiếu điều khoản và tài liệu, rồi kê khai đúng phạm vi tại Phương án B.")

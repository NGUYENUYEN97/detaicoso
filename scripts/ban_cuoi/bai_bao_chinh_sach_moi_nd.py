# -*- coding: utf-8 -*-
"""Nội dung bài báo "Quản lý quyền sở hữu trí tuệ trong trường đại học trước yêu cầu mới của chính sách".

Chỉ dùng văn bản, nghiên cứu và hướng dẫn công khai; không dùng số liệu nội bộ của đề tài.
Văn bản được đối chiếu với tệp trong VBPL/ và Tai_lieu_tham_khao_PDF/, "Tong quan tai lieu/", "Co so ly luan/";
ngày chốt cập nhật văn bản: 30 tháng 9 năm 2026.
Bản sửa theo phản biện: thêm cột điểm thay đổi và hệ quả quản lý; tách tính chất yêu cầu, bổ sung điều kiện pháp lý;
Hình 1 có luồng chính không chờ nghiệm thu; bổ sung Nghị định số 267/2025/NĐ-CP; mô tả lại phương pháp; định nghĩa
chỉ số; Bàn luận về các lựa chọn có đánh đổi.
"""

TIEU_DE = "Quản lý quyền sở hữu trí tuệ trong trường đại học trước yêu cầu mới của chính sách"
TIEU_DE_EN = "Intellectual property rights management in universities in response to new policy requirements"
TAC_GIA = "Nguyễn Thị Tố Uyên, Trần Đăng Bộ"
DON_VI = "Trường Đại học Thành Đô"

TOM_TAT = (
    "Bài báo xác định những thay đổi trong yêu cầu chính sách đối với quản lý quyền sở hữu trí tuệ trong trường đại "
    "học và chuyển các thay đổi đó thành cơ chế thực hiện ở cấp trường. Nghiên cứu sử dụng thiết kế định tính, phân tích "
    "tài liệu và đối chiếu chính sách trên mười ba tài liệu công khai được chọn có chủ đích, gồm bảy văn bản của Việt Nam "
    "cập nhật đến tháng 9 năm 2026, hai hướng dẫn quốc tế và bốn công trình nghiên cứu. Mỗi yêu cầu được phân loại theo "
    "tính chất, gồm định hướng, nghĩa vụ, quyền, điều kiện pháp lý, và theo mức độ thay đổi. Kết quả cho thấy thay đổi "
    "thực chất tập trung ở quyền đối với kết quả nghiên cứu, cơ chế lợi ích của tác giả và quyền tổ chức khai thác; đăng "
    "ký bảo hộ đồng thời với công bố là yêu cầu kế thừa, nay có thêm công cụ pháp lý để thực hiện. Bài báo đề xuất năm "
    "nhóm giải pháp, một quy trình cho phép quyết định và cấp kinh phí nộp đơn trước công bố mà không chờ nghiệm thu, cùng "
    "bộ chỉ số có định nghĩa. Đóng góp của bài báo là cách phân biệt yêu cầu mới với yêu cầu kế thừa và gắn mỗi thay đổi "
    "với quyết định quản lý cụ thể; hệ thống giải pháp cần được kiểm chứng trước khi áp dụng.")

TU_KHOA = "Chính sách sở hữu trí tuệ; Quản lý quyền sở hữu trí tuệ; Tài sản trí tuệ; Trường đại học"

ABSTRACT = (
    "This article identifies how new policies have changed the requirements for intellectual property rights management "
    "in universities and translates those changes into implementation mechanisms at institutional level. The study "
    "adopts a qualitative design based on document analysis and policy comparison of thirteen purposively selected "
    "public documents: seven Vietnamese policy and legal documents current to September 2026, two international guidance "
    "documents and four research works. Each requirement was classified by its nature, as a policy orientation, legal "
    "obligation, right or legal condition, and by its degree of change from earlier rules. Substantive changes are "
    "concentrated in rights over research results, inventor benefit mechanisms and the right to organise "
    "commercialisation; filing for protection at the time of publication and using intellectual property indicators in "
    "evaluation are inherited requirements that now have stronger legal tools. The article proposes five groups of "
    "measures, a procedure that allows filing decisions and fees to be approved before publication without waiting for "
    "project acceptance, and indicators defined by numerator, denominator and data source. Its contribution lies in "
    "distinguishing new from inherited requirements and linking each change to a specific management decision. The "
    "proposed system has not yet been validated and should be tested through expert consultation or pilot "
    "implementation before wider adoption.")
KEYWORDS = "Intellectual property policy; Intellectual property rights management; Intellectual assets; Universities"

# ---------------------------------------------------------------------------------------------- 1
DAT_VAN_DE = [
    "Trường đại học vừa tạo ra tri thức mới qua nghiên cứu và đào tạo, vừa là nơi hình thành nhiều loại tài sản trí tuệ "
    "như sáng chế, kiểu dáng công nghiệp, phần mềm, giáo trình, cơ sở dữ liệu và bí quyết. Một phần giá trị của các tài "
    "sản này phụ thuộc vào việc nhận diện kịp thời, lựa chọn công cụ bảo hộ phù hợp và khai thác (Siegel et al., 2007); "
    "phần khác hình thành qua công bố, chia sẻ và ứng dụng tri thức.",
    "Giai đoạn 2025 - 2026, hoạt động này chịu tác động của nhiều văn bản. Kết luận số 51-KL/TW yêu cầu chuyển từ tư duy "
    "quản lý hành chính sang kiến tạo hệ sinh thái sở hữu trí tuệ (Bộ Chính trị, 2026). Chiến lược sở hữu trí tuệ đến năm "
    "2030 được sửa đổi (Thủ tướng Chính phủ, 2026). Luật Khoa học, công nghệ và đổi mới sáng tạo cùng nghị định quy định "
    "chi tiết, Luật Giáo dục đại học và Luật Sở hữu trí tuệ sửa đổi điều chỉnh quyền đối với kết quả nghiên cứu, quyền "
    "đăng ký và lợi ích của tác giả (Chính phủ, 2025; Quốc hội, 2025a, 2025b; Văn phòng Quốc hội, 2026).",
    "Không phải mọi nội dung trong các văn bản này đều mới: có nội dung thay đổi thực chất, có nội dung được kế thừa và "
    "nhắc lại. Nếu không phân biệt, trường đại học dễ dàn trải nguồn lực hoặc bỏ sót thay đổi buộc phải điều chỉnh. Bài "
    "báo trả lời ba câu hỏi: yêu cầu nào thực sự thay đổi và buộc trường điều chỉnh quyết định, quy trình hay cách phân "
    "chia lợi ích nào; các thay đổi đó cần được chuyển thành giải pháp và cơ chế phối hợp nào; và cần điều kiện, chỉ số "
    "nào để theo dõi. Phạm vi phân tích là tài liệu công khai; bài báo không đánh giá thực trạng của một trường cụ thể.",
]

# ---------------------------------------------------------------------------------------------- 2
TONG_QUAN_21 = [
    "**Chính sách và quy chế sở hữu trí tuệ.** Siegel và cộng sự (2007), tổng quan nghiên cứu về văn phòng chuyển giao "
    "công nghệ ở Hoa Kỳ và châu Âu, kết luận rằng trường đại học cần xây dựng chiến lược thương mại hóa nhất quán và khả "
    "thi, trong đó chiến lược sở hữu trí tuệ phải giải quyết trước vấn đề quyền sở hữu và phạm vi tài sản. Võ (2025), "
    "bằng phân tích quy định pháp luật và so sánh quy chế của một số trường trong nước với chính sách sở hữu trí tuệ của "
    "trường đại học nước ngoài, chỉ ra rằng quy chế trong nước chủ yếu quan tâm đến chủ sở hữu và phân chia lợi nhuận, còn "
    "quyền và nghĩa vụ của các chủ thể không giữ quyền tài sản đối với tác phẩm hình thành trong trường ít được quy định. "
    "Ở cấp chính sách, Khuyến nghị (EU) 2022/2415 đề nghị bảo đảm chính sách và thực hành quản lý tài sản trí tuệ được "
    "xác định, thực hiện, chia sẻ và công khai trong mọi tổ chức tham gia khai thác tri thức (Council of the European "
    "Union, 2022). Các công trình này làm rõ vai trò của quy chế, song tiếp cận từ một nhóm quyền hoặc từ chiến lược chung.",
    "**Tổ chức bộ máy, phân công và phối hợp.** Siegel và cộng sự (2007) tổng hợp các nghiên cứu cho thấy văn phòng chuyển "
    "giao công nghệ cần đạt quy mô tới hạn mới bù đắp được chi phí cố định, từ đó đề xuất hợp tác vùng khi từng trường "
    "chưa đạt quy mô tới hạn về nghiên cứu và chuyên môn chuyển giao. Nguyễn (2025), bằng phân tích tác động của Cách mạng "
    "công nghiệp lần thứ tư, nhấn mạnh yêu cầu phối hợp đồng bộ giữa các bộ phận trong trường. Các nghiên cứu này gợi ý mô "
    "hình tổ chức cần phù hợp quy mô, nhưng chưa phân tích việc phân công ở trường chưa có văn phòng chuyên trách.",
    "**Nhận diện, bảo hộ, quản lý và khai thác.** Siegel và cộng sự (2007) xác định bản khai báo sáng chế của nhà khoa học "
    "là đầu vào then chốt của chuyển giao, trong khi nhiều nhà khoa học không khai báo. Hướng dẫn của Tổ chức Sở hữu trí "
    "tuệ thế giới (World Intellectual Property Organization, 2020) trình bày quy trình thu thập thông tin, tra cứu sáng chế "
    "và phân tích tự do hoạt động, giúp sử dụng sáng chế thuộc phạm vi công cộng và tránh xâm phạm quyền của người khác. "
    "O’Dwyer và cộng sự (2023), qua nghiên cứu trường hợp một mạng lưới hợp tác dược phẩm gồm mười doanh nghiệp và tám cơ "
    "sở học thuật, cho thấy nỗi lo rò rỉ tri thức ở giai đoạn đầu được giải quyết dần, trong đó thỏa thuận sở hữu trí tuệ "
    "hình thành ở giai đoạn gắn kết. Các nghiên cứu này mô tả từng khâu, nhưng ít đặt chúng vào một quy trình có trách "
    "nhiệm và thời hạn ở cấp trường.",
    "**Khuyến khích, phân chia lợi ích và phát triển năng lực.** Siegel và cộng sự (2007) cho rằng cần có tỷ lệ chia hợp "
    "lý cho nhà sáng chế và cần điều chỉnh hệ thống đánh giá, đãi ngộ để hoạt động thương mại hóa được ghi nhận; đồng thời "
    "nhân sự chuyển giao cần kỹ năng thương mại chứ không chỉ kỹ năng pháp lý. Khuyến nghị (EU) 2022/2415 đề nghị hệ thống "
    "khuyến khích công bằng, phát triển kỹ năng cho mọi chủ thể và thống nhất định nghĩa, chỉ số đo lường. Nguyễn (2025) "
    "nêu yêu cầu đào tạo và nâng cao nhận thức cho giảng viên, người học. Hai nghiên cứu quốc tế được lựa chọn dựa trên "
    "pháp luật khác Việt Nam, nên mô hình khuyến khích không thể chuyển nguyên trạng.",
]
TONG_QUAN_22 = [
    "Trong các công trình được lựa chọn, nghiên cứu quốc tế cung cấp bằng chứng về từng cấu phần quản lý, còn nghiên cứu "
    "trong nước bàn về quyền của các chủ thể và yêu cầu phối hợp. Các công trình trong nước công bố năm 2025 chưa phân "
    "tích nghị định quy định chi tiết Luật Khoa học, công nghệ và đổi mới sáng tạo và các văn bản sửa đổi năm 2026. Phần "
    "bài báo có thể bổ sung là việc phân biệt yêu cầu mới với yêu cầu kế thừa và gắn mỗi thay đổi với quyết định quản lý, "
    "chủ thể và chỉ số theo dõi ở cấp trường.",
    "Bài báo sử dụng khung phân tích năm cấu phần: quy chế; tổ chức; quy trình; tài chính và lợi ích; đào tạo và văn hóa "
    "sở hữu trí tuệ. Khung này được xác định trước khi trích xuất, kế thừa từ đề tài cấp cơ sở của nhóm tác giả và đối "
    "chiếu với các nhóm nguyên tắc của Khuyến nghị (EU) 2022/2415.",
]

# ---------------------------------------------------------------------------------------------- 3
PHUONG_PHAP = [
    "Nghiên cứu sử dụng thiết kế định tính, kết hợp phân tích tài liệu và đối chiếu chính sách. Tài liệu được chọn có "
    "chủ đích, không qua tìm kiếm hệ thống trên cơ sở dữ liệu thư mục. Văn bản của Việt Nam được chọn từ tập văn bản nhóm "
    "tác giả thu thập khi thực hiện đề tài cấp cơ sở, theo hai tiêu chí: được ban hành hoặc sửa đổi từ năm 2025 đến ngày "
    "chốt 30 tháng 9 năm 2026; có quy định trực tiếp về quyền đối với kết quả nghiên cứu, bảo hộ, khai thác, phân chia lợi "
    "ích hoặc đào tạo sở hữu trí tuệ trong cơ sở giáo dục đại học. Bảy văn bản đáp ứng tiêu chí là Kết luận số 51-KL/TW, "
    "Quyết định số 1624/QĐ-TTg, Luật số 93/2025/QH15, Nghị định số 267/2025/NĐ-CP, Luật số 125/2025/QH15, Văn bản hợp "
    "nhất số 67/VBHN-VPQH của Luật Sở hữu trí tuệ và Quyết định số 1068/QĐ-TTg; quyết định sau được đưa vào để so sánh nội "
    "dung Chiến lược trước và sau sửa đổi.",
    "Hai hướng dẫn quốc tế và bốn công trình nghiên cứu được chọn từ thư viện tài liệu của đề tài theo hai tiêu chí: có "
    "toàn văn và bàn trực tiếp về ít nhất một cấu phần quản lý. Bốn công trình gồm hai nghiên cứu quốc tế về văn phòng "
    "chuyển giao và hợp tác đại học với doanh nghiệp, hai nghiên cứu trong nước công bố năm 2025. Do số lượng ít và cách "
    "chọn không hệ thống, nhận xét về nghiên cứu chỉ giới hạn trong các công trình được lựa chọn.",
    "Mỗi văn bản được đọc toàn văn và trích xuất vào bảng gồm các trường: căn cứ; yêu cầu; chủ thể và phạm vi; tính chất; "
    "mức độ thay đổi; hệ quả quản lý. Tính chất gồm định hướng chính sách, nghĩa vụ pháp lý, quyền được trao và điều kiện "
    "pháp lý, tức điều kiện để được hưởng bảo hộ; mỗi dòng chỉ chứa yêu cầu cùng tính chất. Mức độ thay đổi gồm mới, sửa "
    "đổi và kế thừa, được xác định bằng cách so sánh văn bản sửa đổi với văn bản gốc và với chú thích về văn bản sửa đổi "
    "trong văn bản hợp nhất; với luật, nghị định thay thế văn bản cũ không có trong tập phân tích, bài báo chỉ ghi là "
    "quy định của văn bản mới kèm thời điểm hiệu lực, không suy ra mức độ thay đổi. Các yêu cầu sau đó được đối chiếu với năm cấu phần để xác định quyết định quản lý cần điều "
    "chỉnh. Việc trích xuất có sử dụng công cụ hỗ trợ như kê khai tại tờ khai kèm bản thảo; mỗi điều khoản được nhóm tác "
    "giả đối chiếu lại với toàn văn. Bảng trích xuất được lưu làm tài liệu kiểm chứng. Nghiên cứu chưa thực hiện mã hóa "
    "độc lập bởi hai người và không khảo sát hay phỏng vấn, nên kết quả là đề xuất có căn cứ văn bản, chưa phải mô hình "
    "đã được kiểm chứng.",
]

# ---------------------------------------------------------------------------------------------- 4.1
KQ_41 = [
    "Bảng 1 tổng hợp các yêu cầu liên quan trực tiếp đến năm nhóm giải pháp, kèm tính chất, điểm thay đổi và hệ quả "
    "quản lý.",
]

BANG_1 = dict(
    tieu_de="Yêu cầu chính sách, điểm thay đổi và hệ quả quản lý đối với trường đại học",
    cot=["Căn cứ", "Yêu cầu", "Tính chất", "Điểm thay đổi", "Hệ quả quản lý"],
    dong=[
        ["**Nhóm 1. Quyền đối với kết quả và lợi ích của tác giả**", "", "", "", ""],
        ["Luật số 93/2025/QH15, khoản 2 Điều 25, Điều 27; Nghị định số 267/2025/NĐ-CP, khoản 2 Điều 32",
         "Tổ chức chủ trì được giao tự động quyền sở hữu phần kết quả tương ứng kinh phí ngân sách, tự quyết phương án "
         "thương mại hóa", "Quyền",
         "Mới: không còn thủ tục giao quyền, bàn giao tài sản; Nghị định số 70/2018/NĐ-CP hết hiệu lực",
         "Quy chế phải tự quy định ai quyết định khai thác, giá và phân chia"],
        ["Luật số 93/2025/QH15, khoản 3 Điều 28, khoản 3, khoản 7 Điều 73; Nghị định số 267/2025/NĐ-CP, khoản 1 Điều 34",
         "Thưởng tác giả tối thiểu 30% lợi nhuận sau thuế, hoặc 30% giá trị kết quả khi góp vốn, với phần kết quả từ "
         "ngân sách", "Nghĩa vụ",
         "Chế độ mới của Luật số 93/2025/QH15; nhiệm vụ phê duyệt trước 01/10/2025 giữ chế độ cũ, trừ trường hợp tại khoản 7 Điều 73",
         "Phân loại tài sản theo nguồn và thời điểm giao; mức trần nội bộ không được làm thưởng thấp hơn mức tối thiểu"],
        ["Nghị định số 267/2025/NĐ-CP, điểm g khoản 2 Điều 17, khoản 3 Điều 34",
         "Hồ sơ đánh giá cuối kỳ có văn bản xác định mức đóng góp của thành viên; thưởng chia theo thỏa thuận giữa các "
         "đồng tác giả", "Nghĩa vụ", "Quy định của nghị định mới, hiệu lực từ 14/10/2025",
         "Ghi nhận đóng góp ngay trong quy trình, không chờ đến khi có nguồn thu"],
        ["Luật Sở hữu trí tuệ, khoản 1 Điều 135",
         "Chủ sở hữu trả thù lao cho tác giả sáng chế, kiểu dáng, thiết kế bố trí theo thỏa thuận; 10% hoặc 15% khi "
         "không có thỏa thuận", "Nghĩa vụ",
         "Sửa đổi: quy định riêng cho kết quả từ ngân sách bị bãi bỏ",
         "Tách thù lao khỏi thưởng; nêu rõ quy chế, hợp đồng có phải là thỏa thuận về thù lao"],
        ["**Nhóm 2. Bảo hộ gắn với nghiên cứu và công bố**", "", "", "", ""],
        ["Quyết định số 1068/QĐ-TTg, điểm b khoản 4 Mục III, được Quyết định số 1624/QĐ-TTg giữ lại",
         "Dùng chỉ số sở hữu trí tuệ để đánh giá; xác định đối tượng quyền cần đạt; trường khối kỹ thuật đăng ký đồng thời "
         "với công bố", "Định hướng", "Kế thừa từ năm 2019",
         "Khai báo từ thuyết minh; quyết định nộp đơn trước công bố"],
        ["Luật Sở hữu trí tuệ, khoản 3 Điều 60",
         "Sáng chế không mất tính mới nếu người có quyền đăng ký, hoặc người có thông tin từ người đó, bộc lộ công khai "
         "và đơn được nộp tại Việt Nam trong 12 tháng", "Điều kiện pháp lý",
         "Kế thừa, sửa đổi năm 2019",
         "Ghi ngày bộc lộ; chỉ dùng như phương án dự phòng"],
        ["Luật Sở hữu trí tuệ, khoản 4 Điều 65",
         "Kiểu dáng không mất tính mới khi bộc lộ như trên nếu đơn được nộp trong 6 tháng", "Điều kiện pháp lý",
         "Sửa đổi bởi Luật số 131/2025/QH15, hiệu lực từ 01/4/2026", "Đưa lịch trưng bày, giới thiệu sản phẩm vào sàng lọc"],
        ["Luật Sở hữu trí tuệ, điểm c khoản 1 Điều 86; Nghị định số 267/2025/NĐ-CP, khoản 6 Điều 32",
         "Tổ chức được giao quyền có quyền đăng ký sáng chế, kiểu dáng, thiết kế bố trí", "Quyền",
         "Sửa đổi: Điều 86a bãi bỏ, nội dung chuyển vào Điều 86, gắn với giao quyền tự động",
         "Xác định người ký đơn và trình tự quyết định nộp đơn"],
        ["**Nhóm 3. Tổ chức hỗ trợ và khai thác**", "", "", "", ""],
        ["Quyết định số 1624/QĐ-TTg, điểm a khoản 6 Mục III, điểm đ khoản 5 Mục II",
         "Phát triển trung tâm tư vấn, hỗ trợ định giá, khai thác thương mại trong cơ sở giáo dục đại học; thí điểm định "
         "giá ít nhất 100 quyền", "Định hướng",
         "Sửa đổi: trước chỉ nêu trung tâm tư vấn; nay thêm định giá, khai thác",
         "Phân công chức năng định giá, khai thác; có thể thuê ngoài"],
        ["Luật số 125/2025/QH15, khoản 1, điểm d khoản 2 Điều 28",
         "Được thành lập doanh nghiệp quản lý tài sản trí tuệ; định giá, góp vốn, phân chia lợi ích", "Quyền",
         "Quy định của luật mới, hiệu lực từ 01/01/2026", "Lựa chọn khi đủ điều kiện, không bắt buộc"],
        ["Nghị định số 267/2025/NĐ-CP, khoản 2 Điều 34",
         "Tổ chức trung gian, môi giới hưởng tối thiểu 10% lợi nhuận khi các bên không có thỏa thuận khác",
         "Nghĩa vụ áp dụng khi không có thỏa thuận", "Quy định của nghị định mới", "Hợp đồng thuê tư vấn, môi giới phải thỏa thuận mức hưởng"],
        ["**Nhóm 4. Kinh phí và công khai**", "", "", "", ""],
        ["Luật số 125/2025/QH15, điểm d khoản 3 Điều 28", "Thành lập và vận hành quỹ phát triển khoa học và công nghệ",
         "Nghĩa vụ", "Quy định của luật mới", "Dòng dự toán cho xác lập quyền đặt trong quỹ"],
        ["Luật số 93/2025/QH15, điểm b khoản 2 Điều 66",
         "Quỹ được chi cho đăng ký, bảo hộ, quản lý, khai thác quyền sở hữu trí tuệ", "Quyền", "Quy định của luật mới, hiệu lực từ 01/10/2025",
         "Căn cứ chi cho đơn nộp trong khi đề tài đang thực hiện"],
        ["Luật số 125/2025/QH15, điểm đ khoản 3 Điều 28",
         "Công khai năng lực, kết quả; cập nhật hằng năm trên nền tảng số quốc gia", "Nghĩa vụ", "Quy định của luật mới",
         "Danh mục số có phân quyền; chỉ công khai dữ liệu được phép"],
        ["**Nhóm 5. Đào tạo và liêm chính**", "", "", "", ""],
        ["Quyết định số 1624/QĐ-TTg, điểm b khoản 8 Mục III",
         "Nghiên cứu đưa sở hữu trí tuệ và kỹ năng khai thác thương mại thành nội dung học bắt buộc", "Định hướng",
         "Sửa đổi: trước chỉ yêu cầu chương trình đào tạo, bồi dưỡng", "Học phần; bồi dưỡng theo vai trò"],
        ["Luật số 125/2025/QH15, điểm a khoản 3 Điều 28", "Tuân thủ đạo đức, liêm chính trong nghiên cứu", "Nghĩa vụ",
         "Quy định của luật mới", "Quy định liêm chính, quyền của bên thứ ba"],
    ],
    nguon="Nguồn: Nhóm tác giả tổng hợp từ Chính phủ (2025), Thủ tướng Chính phủ (2019, 2026), Quốc hội (2025a, "
          "2025b) và Văn phòng Quốc hội (2026). Mức độ thay đổi xác định bằng so sánh với văn bản gốc và chú thích của "
          "văn bản hợp nhất.",
    rong=[3.4, 4.6, 1.9, 3.3, 3.6],
)

KQ_41_SAU = [
    "**Thay đổi thực chất tập trung ở quyền và lợi ích.** Các dòng mới hoặc sửa đổi ở Nhóm 1 cùng hướng: Nhà nước rút "
    "khỏi thủ tục giao quyền, chuyển quyết định khai thác về tổ chức chủ trì, đồng thời đặt mức sàn cho tác giả và yêu "
    "cầu ghi nhận đóng góp (Chính phủ, 2025; Quốc hội, 2025a). Hệ quả là trường không còn dựa được vào thủ tục bên ngoài; "
    "các quyết định trước đây do cơ quan quản lý nhiệm vụ thực hiện nay phải có chủ thể, trình tự và căn cứ trong quy chế "
    "của trường. Cơ chế lợi ích cũng không thể dùng một tỷ lệ chung, vì cùng một trường có tài sản chịu chế độ khác nhau "
    "theo nguồn kinh phí, thời điểm giao nhiệm vụ và loại đối tượng.",
    "**Yêu cầu kế thừa nay có thêm công cụ thực hiện.** Dùng chỉ số sở hữu trí tuệ trong đánh giá và đăng ký đồng thời "
    "với công bố đã có trong Chiến lược từ năm 2019 (Thủ tướng Chính phủ, 2019). Điểm khác là quyền đăng ký của tổ chức "
    "được giao quyền và khoản chi của quỹ cho đăng ký, bảo hộ nay được quy định rõ (Quốc hội, 2025a; Văn phòng Quốc hội, "
    "2026), nên trở ngại còn lại nằm chủ yếu ở cách trường tổ chức quyết định và chi tiền. Thời hạn sau bộc lộ không phải "
    "thời hạn nộp đơn chung: khoản 3 Điều 60 và khoản 4 Điều 65 Luật Sở hữu trí tuệ chỉ giữ tính mới khi người có quyền "
    "đăng ký, hoặc người có thông tin từ người đó, bộc lộ và đơn được nộp trong 12 tháng đối với sáng chế, 6 tháng đối "
    "với kiểu dáng. Vì vậy, nguyên tắc quản lý vẫn là nộp đơn trước khi công bố.",
    "**Quyền mới về tổ chức là lựa chọn, không phải nghĩa vụ.** Quyền thành lập doanh nghiệp quản lý tài sản trí tuệ và "
    "định hướng phát triển trung tâm định giá mở thêm phương án, nhưng không buộc mọi trường lập thêm tổ chức (Quốc hội, "
    "2025b; Thủ tướng Chính phủ, 2026). Ngược lại, nghĩa vụ công khai và mức hưởng mặc định của tổ chức trung gian đòi "
    "hỏi trường điều chỉnh ngay cách quản lý dữ liệu và hợp đồng thuê ngoài.",
]

# ---------------------------------------------------------------------------------------------- 4.2
KQ_42_MO = [
    "Năm nhóm giải pháp dưới đây tương ứng với năm cấu phần quản lý. Mỗi nhóm nêu các dòng của Bảng 1 mà nó xử lý và "
    "quyết định quản lý cần thay đổi.",
]
GP1 = [
    "Giải pháp này xử lý Nhóm 1 của Bảng 1. Vì quyết định khai thác đã chuyển về tổ chức chủ trì, quy chế cần nêu rõ chủ "
    "thể quyết định khai thác, giá, phương án góp vốn và phân chia lợi nhuận; nghĩa vụ khai báo, bảo mật trước công bố và "
    "quyền công bố; quyền của tác giả và của các chủ thể không giữ quyền tài sản, theo hướng Võ (2025) đã nêu; và thứ tự "
    "áp dụng với quy chế nghiên cứu, quy chế chi tiêu.",
    "Phân chia lợi ích cần được thiết kế theo nguồn hình thành tài sản, thời điểm giao nhiệm vụ và loại đối tượng. Với "
    "phần kết quả từ ngân sách thuộc phạm vi Điều 28 Luật số 93/2025/QH15, quy chế không được đặt mức thấp hơn 30% lợi "
    "nhuận sau thuế dành cho tác giả; với nhiệm vụ phê duyệt trước ngày 01 tháng 10 năm 2025, áp dụng văn bản tại thời "
    "điểm phê duyệt, trừ trường hợp tại khoản 7 Điều 73; với tài sản không sử dụng ngân sách, trường tự quyết định tỷ lệ "
    "(Quốc hội, 2025a). Quy chế cần kèm mẫu văn bản xác định mức đóng góp và mẫu thỏa thuận chia thưởng giữa các đồng tác "
    "giả, vì Nghị định số 267/2025/NĐ-CP yêu cầu cả hai (Chính phủ, 2025). Thù lao theo Điều 135 Luật Sở hữu trí tuệ là "
    "khoản riêng, có cơ sở tính khác (Văn phòng Quốc hội, 2026); nhuận bút là khoản thứ ba. Nghị định chưa quy định cách "
    "xác định lợi nhuận sau thuế của một kết quả hình thành từ nhiều nguồn kinh phí, nên quy chế cần tự nêu phương pháp "
    "phân bổ chi phí. Tỷ lệ cụ thể nên được mô phỏng trên một số tình huống chuyển giao trước khi ban hành.",
]

GP2 = [
    "Giải pháp này xử lý Nhóm 3 của Bảng 1. Nhóm tác giả đề xuất phân công theo bốn chức năng thay vì mặc định một mô "
    "hình: quản lý khoa học, gồm tiếp nhận khai báo và theo dõi danh mục; pháp lý, gồm thẩm định quyền, soạn hồ sơ, làm "
    "việc với cơ quan đăng ký; tài chính, gồm dự toán và chi trả; hỗ trợ khai thác, gồm tìm đối tác, đàm phán, định giá.",
    "Theo tổng hợp của Siegel và cộng sự (2007), văn phòng chuyên trách cần đạt quy mô tới hạn mới bù đắp được chi phí cố "
    "định; quy mô cần thiết còn tùy chức năng được giao và loại tài sản. Trường có ít kết quả có thể bảo hộ có thể giao "
    "nhiệm vụ kiêm nhiệm cho các phòng hiện có và thuê tổ chức đại diện sở hữu công nghiệp theo hồ sơ, xem xét vị trí "
    "chuyên trách khi số hồ sơ tăng. Hợp đồng thuê tư vấn, môi giới cần thỏa thuận rõ mức hưởng, vì khi không có thỏa "
    "thuận, tổ chức trung gian hưởng tối thiểu 10% lợi nhuận (Chính phủ, 2025). Trung tâm định giá hay doanh nghiệp quản "
    "lý tài sản trí tuệ nên được đặt ra khi danh mục đã có tài sản sẵn sàng khai thác và có phương án tài chính vận hành.",
]

GP3 = [
    "Giải pháp này xử lý Nhóm 2 của Bảng 1. Mục tiêu là kết quả có khả năng bảo hộ được quyết định và nộp đơn trước khi "
    "công bố, với cách xử lý phù hợp loại tài sản, theo quy trình tại Hình 1.",
]
GP3_SAU = [
    "Luồng chính gồm bảy khâu. Khâu 1, người nghiên cứu khai báo đối tượng dự kiến, chủ thể quyền, nguồn kinh phí và lịch "
    "công bố ngay từ thuyết minh; từ thời điểm này, thông tin cần bảo vệ được giữ bí mật. Khâu 2, đầu mối quản lý khoa "
    "học sàng lọc và phân nhánh theo loại tài sản (Văn phòng Quốc hội, 2026). Khâu 3, tra cứu và đánh giá khả năng bảo "
    "hộ; tra cứu sáng chế và phân tích tự do hoạt động còn giúp tránh xâm phạm quyền của bên thứ ba (World Intellectual "
    "Property Organization, 2020). Khâu 4, lãnh đạo trường quyết định nộp đơn, giữ bí mật hoặc công bố, đồng thời cấp "
    "kinh phí. Khâu 5, nộp đơn hoặc duy trì biện pháp bảo mật. Khâu 6, công bố hoặc trình diễn chỉ sau khi đã có ngày nộp "
    "đơn hoặc quyết định không bảo hộ. Khâu 7, khai thác theo Giải pháp 1.",
    "Luồng này không chờ nghiệm thu: kết quả cần bảo vệ trong khi đề tài đang thực hiện được chuyển thẳng từ Khâu 3 sang "
    "Khâu 4. Nghiệm thu là điểm kiểm tra đặt bên cạnh luồng chính, dùng để đối chiếu khai báo, cập nhật tình trạng quyền, "
    "lập văn bản xác định mức đóng góp và đưa kết quả chưa khai báo quay lại Khâu 2. Nếu kết quả đã lỡ bộc lộ, đầu mối ghi "
    "ngày bộc lộ và xem xét thời hạn tại khoản 3 Điều 60 hoặc khoản 4 Điều 65 Luật Sở hữu trí tuệ. Với kết quả hợp tác "
    "doanh nghiệp, thỏa thuận sở hữu trí tuệ cần được xác lập trong giai đoạn gắn kết, như O’Dwyer và cộng sự (2023) mô tả.",
]

GP4 = [
    "Giải pháp này xử lý Nhóm 4 của Bảng 1. Mục tiêu là chi phí tra cứu, đăng ký, duy trì và khai thác có nguồn, có người "
    "đề xuất chi và có thời hạn. Cơ sở giáo dục đại học có nghĩa vụ thành lập quỹ phát triển khoa học và công nghệ, và quỹ "
    "được chi cho đăng ký, bảo hộ, quản lý, khai thác quyền sở hữu trí tuệ (Quốc hội, 2025a, 2025b). Trường có thể lập "
    "dòng dự toán riêng trong quỹ cho chi phí xác lập quyền, tính từ số hồ sơ dự kiến và mức phí theo quy định. Đơn nộp "
    "trong khi đề tài đang thực hiện được chi từ dòng này, không trừ vào kinh phí đề tài, để người nghiên cứu không phải "
    "chọn giữa kinh phí nghiên cứu và nộp đơn; quy định tạm ứng áp dụng cho cả đơn nộp trước nghiệm thu và sau khi đề tài "
    "đã quyết toán.",
    "Cơ chế khuyến khích nội bộ có thể ghi nhận đóng góp ở nhiều mốc, như khi đơn được chấp nhận hợp lệ và khi văn bằng "
    "được cấp, và cân đối với chế độ dành cho công bố, vì hoạt động thương mại hóa cần được ghi nhận trong hệ thống đánh "
    "giá, đãi ngộ (Siegel et al., 2007). Khoản này tách biệt với thưởng khi thương mại hóa và thù lao theo luật; mức cụ "
    "thể phụ thuộc khả năng tài chính của từng trường.",
]

GP5 = [
    "Giải pháp này xử lý Nhóm 5 của Bảng 1. Đào tạo nên phân theo vai trò: giảng viên cần kỹ năng nhận diện đối tượng bảo "
    "hộ, khai báo và bộc lộ an toàn; người học cần hiểu quyền tác giả, liêm chính và quyền của bên thứ ba; nhân sự quản "
    "lý cần kỹ năng tra cứu và hỗ trợ khai thác, vì nhân sự chuyển giao cần kỹ năng thương mại bên cạnh kỹ năng pháp lý "
    "(Siegel et al., 2007). Nội dung tra cứu có thể dựa trên hướng dẫn công khai của Tổ chức Sở hữu trí tuệ thế giới "
    "(World Intellectual Property Organization, 2020). Học phần sở hữu trí tuệ và kỹ năng khai thác thương mại có thể "
    "được đưa vào chương trình theo lộ trình (Thủ tướng Chính phủ, 2026); quy định liêm chính được phổ biến cùng cẩm nang "
    "ngắn và đầu mối giải đáp (Bộ Chính trị, 2026).",
]

# ---------------------------------------------------------------------------------------------- 4.3
KQ_43 = [
    "Hệ thống giải pháp nên được triển khai theo ba bước: rà soát quy chế hiện hành theo Bảng 1 và giao đầu mối; áp dụng "
    "quy trình cho một đơn vị hoặc một nhóm đề tài có nhiều kết quả có khả năng bảo hộ; đánh giá, điều chỉnh rồi áp dụng "
    "toàn trường. Bảng 2 chỉ giữ những chỉ số định nghĩa được tử số, mẫu số và nguồn dữ liệu, phân thành chỉ số hoạt động, "
    "đầu ra và kết quả khai thác, phù hợp khuyến nghị thống nhất định nghĩa, chỉ số cho các kênh khai thác tri thức "
    "(Council of the European Union, 2022). Bài báo không đặt mức đích chung, vì mức đích phụ thuộc quy mô nghiên cứu và "
    "lĩnh vực của từng trường.",
]
KQ_43_SAU = [
    "Số văn bằng được cấp không được gán cho riêng một giải pháp: đây là đầu ra chịu tác động của nhiều cấu phần và có độ "
    "trễ do các giai đoạn thẩm định, nên chỉ được theo dõi như đầu ra chung của hệ thống. Năm đầu triển khai nên được "
    "đánh giá chủ yếu bằng chỉ số hoạt động.",
]

BANG_2 = dict(
    tieu_de="Chỉ số theo dõi, định nghĩa và nguồn dữ liệu",
    cot=["Chỉ số", "Định nghĩa, cách tính", "Loại; giải pháp", "Nguồn dữ liệu"],
    dong=[
        ["1. Tỷ lệ kết quả khai báo trước công bố",
         "Tử số: kết quả thuộc nhánh sáng chế, giải pháp hữu ích, kiểu dáng có phiếu khai báo ghi ngày trước lần công bố "
         "đầu tiên. Mẫu số: kết quả thuộc các nhánh này nêu trong báo cáo tổng kết của đề tài nghiệm thu trong năm",
         "Hoạt động; 3", "Phiếu khai báo; báo cáo tổng kết; danh mục công bố"],
        ["2. Thời gian xử lý",
         "Số ngày làm việc từ ngày nhận phiếu khai báo đầy đủ đến ngày có quyết định tại Khâu 4, trừ thời gian chờ người "
         "nghiên cứu bổ sung; báo cáo trung vị, tách đơn nộp trước và sau nghiệm thu",
         "Hoạt động; 2, 3", "Sổ theo dõi hồ sơ"],
        ["3. Tỷ lệ hồ sơ được tạm ứng đúng hạn",
         "Tử số: hồ sơ được tạm ứng phí trước hạn nộp đơn ghi trong quyết định. Mẫu số: hồ sơ có quyết định nộp đơn trong "
         "năm", "Hoạt động; 4", "Quyết định; chứng từ tạm ứng"],
        ["4. Tỷ lệ đơn nộp trước công bố",
         "Tử số: đơn sáng chế, giải pháp hữu ích, kiểu dáng có ngày nộp đơn trước ngày bộc lộ đầu tiên. Mẫu số: đơn các "
         "loại này nộp trong năm", "Đầu ra; 3, 4", "Tờ khai; danh mục công bố"],
        ["5. Tỷ lệ người được bồi dưỡng theo vai trò",
         "Tử số: chủ nhiệm đề tài mới và nhân sự đầu mối hoàn thành bồi dưỡng. Mẫu số: số người thuộc hai nhóm này trong "
         "năm", "Hoạt động; 5", "Danh sách bồi dưỡng"],
        ["6. Hợp đồng và nguồn thu",
         "Hợp đồng chuyển nhượng hoặc chuyển quyền sử dụng đã ký trong năm; nguồn thu là tiền thực thu trong năm, ghi "
         "riêng giá trị ký kết", "Kết quả khai thác; 1, 2", "Hợp đồng; sổ kế toán"],
        ["7. Khoản chi trả cho tác giả",
         "Số tiền thực chi trong năm, tách thưởng, thù lao, nhuận bút", "Kết quả khai thác; 1", "Sổ kế toán"],
    ],
    nguon="Nguồn: Nhóm tác giả đề xuất.",
    rong=[3.6, 7.6, 2.4, 3.4],
)

# ---------------------------------------------------------------------------------------------- 5
BAN_LUAN = [
    "Kết quả gợi ý năm nhóm giải pháp phụ thuộc lẫn nhau. Quy trình sàng lọc khó vận hành nếu quy chế chưa đặt nghĩa vụ "
    "khai báo và bảo mật, nếu đầu mối không tiếp nhận trong thời hạn hợp lý, hoặc nếu quyết định nộp đơn không đi kèm "
    "nguồn chi. Bài báo giả định rằng người nghiên cứu sẽ khai báo nhiều hơn khi hiểu điều kiện giữ tính mới và thấy đóng "
    "góp được ghi nhận; giả định hành vi này chưa được kiểm chứng trong nghiên cứu.",
    "Cơ chế đề xuất buộc trường lựa chọn giữa một số phương án có đánh đổi. Thứ nhất, công bố sớm và bảo hộ: chờ nộp đơn "
    "có thể làm chậm công bố, ảnh hưởng đến tiến độ của người nghiên cứu, trong khi giá trị học thuật và xã hội của nhiều "
    "kết quả hình thành chính qua công bố, chia sẻ. Vì vậy, Khâu 4 phải cho phép quyết định công bố mà không bảo hộ, và "
    "luồng không chờ nghiệm thu nhằm rút ngắn thời gian chờ chứ không giữ mọi kết quả lại. Thời hạn sau bộc lộ chỉ là "
    "phương án dự phòng, vì không phải mọi trường hợp bộc lộ đều đáp ứng điều kiện luật định. Thứ hai, thuê ngoài và phát "
    "triển năng lực nội bộ: thuê tổ chức đại diện linh hoạt về chi phí và chuyên môn soạn đơn, nhưng không tích lũy năng "
    "lực nhận diện sớm trong trường; phương án hợp lý hơn với trường quy mô nhỏ là giữ Khâu 2 trong trường và thuê ngoài "
    "Khâu 3, Khâu 5 theo hồ sơ. Thứ ba, tăng số đơn và kiểm soát chất lượng tài sản: chỉ tiêu số đơn có thể khuyến khích "
    "nộp đơn ít giá trị khai thác, phát sinh phí duy trì. Do đó, Bảng 2 dùng tỷ lệ đơn nộp trước công bố thay cho số đơn "
    "tuyệt đối, và quy trình cần có quyết định định kỳ về việc tiếp tục hay ngừng duy trì văn bằng.",
    "Đối chiếu với các công trình được lựa chọn, hệ thống đề xuất kế thừa ba điểm: khai báo là đầu vào của quy trình "
    "(Siegel et al., 2007); tra cứu sáng chế và phân tích tự do hoạt động được đưa vào quy trình và đào tạo (World "
    "Intellectual Property Organization, 2020); thỏa thuận sở hữu trí tuệ được xác lập trong giai đoạn gắn kết của hợp "
    "tác (O’Dwyer et al., 2023). Điểm khác là hai nghiên cứu quốc tế được lựa chọn khảo sát trường có văn phòng chuyển "
    "giao chuyên trách, trong khi bài báo đề xuất phân công theo chức năng cho trường chưa có văn phòng. Điểm khác thứ hai "
    "là ở Việt Nam, trường phải đồng thời áp dụng mức thưởng tối thiểu theo nguồn kinh phí, thù lao theo Luật Sở hữu trí "
    "tuệ và quy định của Nghị định số 267/2025/NĐ-CP về đồng tác giả, nên các khoản này cần được tách ngay trong quy chế.",
    "Đóng góp của bài báo là cách phân loại yêu cầu theo hai chiều: tính chất, gồm định hướng, nghĩa vụ, quyền và điều "
    "kiện pháp lý; và mức độ thay đổi, gồm mới, sửa đổi và kế thừa. Cách phân loại này giúp trường ưu tiên điều chỉnh ở "
    "nơi pháp luật thực sự thay đổi, phân biệt điều kiện để được bảo hộ với nghĩa vụ phải thực hiện, và tránh đầu tư bộ "
    "máy chỉ vì một quyền mới được trao, phù hợp với tinh thần công khai chính sách và thống nhất chỉ số của Khuyến nghị "
    "(EU) 2022/2415 (Council of the European Union, 2022). Điều kiện áp dụng khác nhau theo loại trường: trường khối kỹ "
    "thuật, y dược cần tập trung vào luồng nộp đơn trước công bố và kinh phí; trường khối kinh tế, xã hội chủ yếu tạo ra "
    "tác phẩm, nên trọng tâm là quyền của các chủ thể và liêm chính, như Võ (2025) đã phân tích; trường có nhiều nhiệm vụ "
    "từ ngân sách cần chú ý quy định chuyển tiếp tại Điều 73 Luật số 93/2025/QH15.",
    "Bài báo có một số giới hạn. Tài liệu được chọn có chủ đích và số nghiên cứu ít, nên nhận xét về nghiên cứu không khái "
    "quát cho toàn bộ tài liệu quốc tế. Việc phân loại do nhóm tác giả thực hiện, chưa có mã hóa độc lập. Nghiên cứu dựa "
    "trên văn bản, chưa phản ánh cách các trường đang thực hiện; hệ thống giải pháp chưa được kiểm chứng qua triển khai. "
    "Nghị định số 267/2025/NĐ-CP đã được đối chiếu, nhưng chưa có quy định về phương pháp xác định lợi nhuận của một kết "
    "quả hình thành từ nhiều nguồn kinh phí. Nghiên cứu tiếp theo có thể nhờ người thứ hai mã hóa độc lập bảng trích "
    "xuất, tham vấn chuyên gia quản lý khoa học, pháp chế, tài chính, hoặc thí điểm quy trình tại một số đơn vị và dùng "
    "Bảng 2 để đo kết quả trước và sau khi áp dụng.",
]

# ---------------------------------------------------------------------------------------------- 6
KET_LUAN = [
    "Bài báo xác định những thay đổi trong yêu cầu chính sách đối với quản lý quyền sở hữu trí tuệ trong trường đại học "
    "và chuyển các thay đổi đó thành cơ chế thực hiện ở cấp trường.",
    "Kết quả cho thấy thay đổi thực chất tập trung ở quyền đối với kết quả nghiên cứu, cơ chế lợi ích của tác giả và "
    "quyền tổ chức khai thác: quyền được giao tự động, mức thưởng tối thiểu theo nguồn kinh phí, yêu cầu ghi nhận đóng góp "
    "của đồng tác giả và quyền thành lập doanh nghiệp quản lý tài sản trí tuệ. Yêu cầu dùng chỉ số sở hữu trí tuệ trong "
    "đánh giá và đăng ký bảo hộ đồng thời với công bố là nội dung kế thừa, nay có thêm quyền đăng ký và căn cứ chi rõ "
    "hơn. Thời hạn sau bộc lộ là điều kiện để giữ tính mới, không thay thế nguyên tắc nộp đơn trước công bố. Từ các thay "
    "đổi này, quyết định khai thác, phân chia lợi ích, nộp đơn và chi tiền phải có chủ thể, trình tự và căn cứ trong quy "
    "chế của trường; quy trình cần cho phép quyết định và cấp kinh phí nộp đơn trước khi công bố mà không chờ nghiệm thu.",
    "Đóng góp của bài báo là cách phân biệt yêu cầu theo tính chất và mức độ thay đổi, gắn mỗi thay đổi với quyết định "
    "quản lý và chỉ số có định nghĩa. Về hàm ý, mỗi trường nên bắt đầu bằng việc rà soát quy chế theo các dòng mới và "
    "sửa đổi của Bảng 1; ưu tiên luồng nộp đơn trước công bố và nguồn chi tương ứng; lựa chọn hình thức tổ chức theo quy "
    "mô hồ sơ thực tế; và đánh giá năm đầu bằng chỉ số hoạt động, có cân nhắc các đánh đổi giữa công bố và bảo hộ, thuê "
    "ngoài và năng lực nội bộ, số lượng và chất lượng đơn.",
]

TAI_LIEU = [
    "Bộ Chính trị. (2026). *Kết luận số 51-KL/TW ngày 17 tháng 6 năm 2026 về đẩy mạnh công tác sở hữu trí tuệ phục vụ "
    "phát triển kinh tế - xã hội trong tình hình mới*.",
    "Chính phủ. (2025). *Nghị định số 267/2025/NĐ-CP ngày 14 tháng 10 năm 2025 quy định chi tiết và hướng dẫn một số điều "
    "của Luật Khoa học, công nghệ và đổi mới sáng tạo về chương trình, nhiệm vụ khoa học, công nghệ và đổi mới sáng tạo và "
    "một số quy định về thúc đẩy hoạt động nghiên cứu khoa học, phát triển công nghệ và đổi mới sáng tạo*.",
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
KHOA_TRICH = ["Bộ Chính trị, 2026", "Chính phủ, 2025", "Council of the European Union, 2022", "Nguyễn (2025)",
              "O’Dwyer", "Quốc hội, 2025a", "2025b", "Siegel", "Thủ tướng Chính phủ, 2019",
              "Thủ tướng Chính phủ, 2026", "Văn phòng Quốc hội, 2026", "Võ (2025)",
              "World Intellectual Property Organization, 2020"]

TO_KHAI_GHI_CHU = (
    "Ghi chú cho nhóm tác giả khi điền tờ khai: bản thảo được soạn với sự hỗ trợ của công cụ trí tuệ nhân tạo Claude "
    "của Anthropic trong các việc: tổ chức cấu trúc theo khung do chủ nhiệm đề tài xây dựng, trích xuất điều khoản từ văn "
    "bản pháp luật có trong hồ sơ, so sánh văn bản trước và sau sửa đổi, biên soạn câu chữ, sửa bản thảo theo ý kiến phản "
    "biện, định dạng tài liệu tham khảo và dựng tệp Word. Nhóm tác giả cần đọc lại toàn bộ nội dung, đối chiếu điều khoản "
    "và tài liệu, rồi kê khai đúng phạm vi tại Phương án B.")

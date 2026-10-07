# -*- coding: utf-8 -*-
"""Bài báo đầu ra (phương án dự phòng): khung đánh giá hiệu quả quản lý quyền sở hữu trí tuệ.

    python3 scripts/ban_cuoi/bai_bao_khung.py

Đầu ra: Ban_cuoi/Bai_bao_phuong_an_khung_danh_gia.docx (phương án dự phòng)

Bài lý luận phát triển từ Mục 1.4 của báo cáo tổng kết: bộ chỉ số theo chuỗi đầu vào,
quá trình, đầu ra, kết quả, bổ sung bốn chỉ số chuyển hóa và phân loại nguồn dữ liệu theo
Thông tư số 83/2026/TT-BGDĐT, Luật số 93/2025/QH15, Luật số 125/2025/QH15. Bài không sử dụng
số liệu nội bộ của cơ sở giáo dục đại học nào.
"""
import collections
import os
import re
import sys

from docx.shared import Pt

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import khung  # noqa: E402
import tai_lieu as TL  # noqa: E402

RA_DOCX = os.path.join(khung.THU_MUC_RA, "Bai_bao_phuong_an_khung_danh_gia.docx")
SO_DO = os.path.join(khung.THU_MUC_RA, "so_do")

TIEU_DE = ("Khung đánh giá hiệu quả quản lý quyền sở hữu trí tuệ tại trường đại học định hướng ứng dụng: Tiếp cận "
           "chuỗi kết quả")
TAC_GIA = "Nguyễn Thị Tố Uyên, Trần Đăng Bộ"
DON_VI = "Trường Đại học Thành Đô"

# ---------------------------------------------------------------------------
# Bộ chỉ số: (mã, nhóm, tên, cách tính, nguồn dữ liệu, loại dữ liệu)
# A: dữ liệu bắt buộc báo cáo hoặc công khai theo quy định; B: có trong hồ sơ hành chính
# thông thường của nhà trường; C: nhà trường phải thiết lập công cụ thu thập mới.
# ---------------------------------------------------------------------------
CHI_SO = [
    ("V1", "Đầu vào", "Kinh phí xác lập, duy trì quyền và khen thưởng sáng tạo", "Tổng chi trong năm",
     "Sổ kế toán, quỹ phát triển khoa học và công nghệ", "B"),
    ("V2", "Đầu vào", "Nhân lực quản lý sở hữu trí tuệ, chuyển giao công nghệ",
     "Số người quy đổi toàn thời gian; tỷ lệ được đào tạo nghiệp vụ", "Hồ sơ nhân sự", "B"),
    ("V3", "Đầu vào", "Hạ tầng tra cứu sáng chế, kiểm tra trùng lặp", "Có hoặc không; số lượt tra cứu",
     "Đơn vị quản lý khoa học", "B"),
    ("Q1", "Quá trình", "Mức độ tương thích của quy chế nội bộ với pháp luật hiện hành",
     "Quy chế được ban hành đúng thẩm quyền, còn hiệu lực và cập nhật theo luật mới", "Chỉ số 1.1, Thông tư số "
                                                                                       "83/2026/TT-BGDĐT", "A"),
    ("Q2", "Quá trình", "Thời gian xử lý một hồ sơ đề xuất bảo hộ", "Số ngày bình quân từ khai báo đến quyết định "
                                                                     "nộp đơn", "Sổ theo dõi hồ sơ", "C"),
    ("Q3", "Quá trình", "Mức độ chuẩn hóa biểu mẫu và quy trình phối hợp", "Số khâu có biểu mẫu chuẩn / tổng số khâu",
     "Quy trình nội bộ", "C"),
    ("Q4", "Quá trình", "Tập huấn sở hữu trí tuệ cho giảng viên, người học", "Số lớp; số lượt người tham gia",
     "Đơn vị tổ chức đào tạo", "B"),
    ("R1", "Đầu ra", "Đơn đăng ký sở hữu công nghiệp", "Số đơn nộp trong năm theo loại", "Sổ theo dõi đơn", "B"),
    ("R2", "Đầu ra", "Văn bằng bảo hộ được cấp", "Số văn bằng theo loại; quy đổi theo chỉ số 6.2.1",
     "Dữ liệu báo cáo theo Thông tư số 83/2026/TT-BGDĐT", "A"),
    ("R3", "Đầu ra", "Giấy chứng nhận đăng ký quyền tác giả", "Số giấy chứng nhận cho giáo trình, bài giảng, phần "
                                                              "mềm", "Sổ theo dõi", "B"),
    ("R4", "Đầu ra", "Công trình có sản phẩm có khả năng bảo hộ", "Số công trình được hội đồng nghiệm thu xác nhận",
     "Phiếu rà soát tại nghiệm thu", "C"),
    ("K1", "Kết quả", "Hợp đồng chuyển giao, cấp phép sử dụng", "Số hợp đồng còn hiệu lực trong năm",
     "Báo cáo theo khoản 4 Điều 27 Luật số 93/2025/QH15", "A"),
    ("K2", "Kết quả", "Nguồn thu từ khai thác tài sản trí tuệ", "Tổng thu trong năm",
     "Dữ liệu báo cáo theo Thông tư số 83/2026/TT-BGDĐT", "A"),
    ("K3", "Kết quả", "Tỷ trọng thu khoa học, công nghệ và đổi mới sáng tạo", "Bình quân ba năm, theo chỉ số 6.1",
     "Thông tư số 83/2026/TT-BGDĐT", "A"),
    ("K4", "Kết quả", "Đóng góp vào kiểm định chất lượng và xếp hạng", "Phần đóng góp của văn bằng vào chỉ số 6.2.1 "
                                                                       "và các bảng xếp hạng", "Báo cáo tự đánh giá",
     "C"),
    ("K5", "Kết quả", "Động lực đổi mới sáng tạo của giảng viên", "Điểm thang đo khảo sát", "Khảo sát định kỳ", "C"),
    ("CH1", "Chuyển hóa", "Tỷ lệ nhận diện", "R4 / số công trình nghiệm thu trong kỳ", "Phiếu rà soát tại nghiệm thu",
     "C"),
    ("CH2", "Chuyển hóa", "Tỷ lệ xác lập kịp thời", "Số sản phẩm của R4 được nộp đơn trong 12 tháng / R4",
     "Phiếu rà soát, sổ theo dõi đơn", "C"),
    ("CH3", "Chuyển hóa", "Tỷ lệ khai thác", "Số tài sản có giao dịch hoặc được sử dụng / số tài sản đã xác lập",
     "Danh mục tài sản trí tuệ", "C"),
    ("CH4", "Chuyển hóa", "Tỷ suất khai thác trên chi phí", "K2 / phần chi xác lập và duy trì quyền trong V1",
     "Sổ kế toán", "B"),
]
NHOM = ["Đầu vào", "Quá trình", "Đầu ra", "Kết quả", "Chuyển hóa"]
DEM = {n: collections.Counter(c[5] for c in CHI_SO if c[1] == n) for n in NHOM}
TONG = collections.Counter(c[5] for c in CHI_SO)
assert len(CHI_SO) == 20 and (TONG["A"], TONG["B"], TONG["C"]) == (5, 7, 8)
assert DEM["Kết quả"]["A"] == 3 and DEM["Chuyển hóa"]["A"] == 0 and DEM["Chuyển hóa"]["C"] == 3

TOM_TAT = (
    "Bài báo đề xuất khung đánh giá hiệu quả quản lý quyền sở hữu trí tuệ cho trường đại học định hướng ứng "
    "dụng, trong bối cảnh pháp luật mới yêu cầu sử dụng chỉ số sở hữu trí tuệ để đánh giá cơ sở giáo dục đại học. "
    "Nghiên cứu sử dụng thiết kế phát triển khung khái niệm: tổng hợp các công trình và bộ chỉ số quốc tế về đo lường "
    "chuyển giao tri thức, xây dựng chỉ số theo mô hình chuỗi kết quả gắn với chu trình quản lý tài sản trí tuệ, và "
    "đối chiếu với các văn bản pháp luật, chuẩn chất lượng hiện hành để phân loại khả năng có dữ liệu. Kết quả là bộ "
    "chỉ số gồm bốn nhóm đầu vào, quá trình, đầu ra, kết quả và bốn chỉ số chuyển hóa giữa các mắt xích. Phân tích cho "
    "thấy dữ liệu bắt buộc báo cáo tập trung ở nhóm kết quả, trong khi các chỉ số chuyển hóa và phần lớn chỉ số quá "
    "trình đòi hỏi nhà trường tự thiết lập công cụ thu thập. Nếu chỉ dựa vào dữ liệu bắt buộc, nhà trường biết đạt "
    "được bao nhiêu nhưng không biết điểm nghẽn nằm ở đâu. Bài báo đóng góp một công cụ tự đánh giá có thể áp dụng ở "
    "cấp trường, kèm các mẫu hình chẩn đoán điểm nghẽn và lộ trình triển khai theo mức độ sẵn có của dữ liệu.")
TU_KHOA = "Đánh giá hiệu quả; Quản lý quyền sở hữu trí tuệ; Chuỗi kết quả; Bộ chỉ số; Trường đại học định hướng ứng dụng"

ABSTRACT = (
    "This article proposes a framework for evaluating the effectiveness of intellectual property rights management in "
    "application-oriented universities, at a time when new Vietnamese legislation requires intellectual property "
    "indicators to be used in assessing higher education institutions. The study follows a conceptual framework "
    "development design: it synthesises international studies and indicator sets on knowledge transfer measurement, "
    "constructs indicators using a results-chain logic model aligned with the intellectual property management cycle, "
    "and cross-checks them against current laws and quality standards to classify data availability. The result is "
    "an indicator set organised into four groups, namely inputs, processes, outputs and outcomes, complemented by four "
    "conversion indicators that link successive stages of the chain. The analysis shows that mandatory reporting data "
    "are concentrated in the outcome group, whereas the conversion indicators and most process indicators require "
    "universities to build their own data collection tools. Relying only on mandatory data, a university can tell how "
    "much it has achieved but not where the bottleneck lies. The article contributes a self-assessment instrument "
    "applicable at institutional level, together with diagnostic patterns for locating bottlenecks and a phased "
    "implementation path based on data availability. The framework is particularly relevant to institutions without a "
    "professional technology transfer office, where potential intellectual assets are most often lost at the "
    "identification stage.")
KEYWORDS = ("Effectiveness evaluation; Intellectual property management; Results chain; Indicator set; "
            "Application-oriented university")

DAT_VAN_DE = [
    "Quản lý quyền sở hữu trí tuệ đang trở thành nội dung được đo lường trong quản trị đại học. Quyết định "
    "số 1624/QĐ-TTg sửa đổi Chiến lược sở hữu trí tuệ đến năm 2030 yêu cầu sử dụng các chỉ số đo lường về sở hữu trí "
    "tuệ làm căn cứ đánh giá hiệu quả hoạt động của cơ sở giáo dục đại học. Luật Giáo dục đại học số 125/2025/QH15 bổ "
    "sung nghĩa vụ công khai hằng năm kết quả hoạt động khoa học, công nghệ và đổi mới sáng tạo; Chuẩn cơ sở "
    "giáo dục đại học ban hành kèm Thông tư số 83/2026/TT-BGDĐT tính văn bằng bảo hộ vào chỉ số sản phẩm khoa học quy "
    "đổi.",
    "Tuy nhiên, các chỉ số hiện hành chủ yếu đếm kết quả cuối cùng như số văn bằng hay nguồn thu, nên cho biết nhà "
    "trường đạt được bao nhiêu nhưng không cho biết hiệu quả, cũng không chỉ ra tiềm năng bị mất ở khâu nào. Các bộ chỉ số quốc tế về chuyển giao tri thức được thiết kế cho trường đại học "
    "nghiên cứu có đơn vị chuyển giao công nghệ chuyên nghiệp (Campbell et al., 2020; Finne et al., 2009), trong khi "
    "nghiên cứu trong nước chưa đề xuất một khung đo lường hoàn chỉnh cho trường đại học định hướng ứng dụng.",
    "Bài báo trả lời hai câu hỏi: hiệu quả quản lý quyền sở hữu trí tuệ ở trường định hướng ứng dụng nên được "
    "đo theo những chiều nào; và bộ chỉ số nào khả thi với nguồn dữ liệu hiện có cùng các yêu cầu báo cáo mới.",
]

TONG_QUAN = [
    "**Đo lường hiệu quả chuyển giao bằng quan hệ đầu vào, đầu ra.** Thursby và Kemp (2002) đánh giá hiệu quả sản xuất "
    "của hoạt động cấp phép tại các trường đại học Hoa Kỳ bằng cách so sánh đầu ra với nguồn lực đầu vào, và cho thấy "
    "hiệu quả khác nhau đáng kể giữa các trường. Siegel và cộng sự (2007) tổng hợp rằng hiệu quả của đơn vị chuyển giao "
    "công nghệ chịu ảnh hưởng của chính sách khuyến khích, năng lực nhân sự và cơ chế chia lợi ích. Hướng tiếp cận này "
    "khẳng định hiệu quả phải được đo bằng quan hệ giữa các mắt xích, không chỉ bằng số lượng đầu ra; tuy vậy, các phép "
    "đo so sánh như vậy cần dữ liệu của nhiều trường và khó áp dụng cho tự đánh giá ở một trường.",
    "**Các bộ chỉ số quốc tế.** Nhóm chuyên gia của Ủy ban châu Âu đề xuất một tập chỉ số cốt lõi về chuyển giao tri "
    "thức từ tổ chức nghiên cứu công, gồm các chỉ số như khai báo sáng chế, đơn đăng ký, văn bằng, hợp đồng cấp phép, "
    "nguồn thu cấp phép và doanh nghiệp khởi nguồn (Finne et al., 2009). Campbell và cộng sự (2020) tiếp tục hướng tới "
    "một bộ chỉ số hài hòa ở cấp châu Âu, mở rộng sang các kênh chuyển giao tri thức khác ngoài sở hữu trí tuệ. Các bộ "
    "chỉ số này có ưu điểm chuẩn hóa định nghĩa để so sánh, nhưng phần lớn là chỉ số đếm ở đầu ra và kết quả, và giả "
    "định đã có đơn vị chuyển giao thu thập dữ liệu.",
    "**Từ chiếm hữu sang khai thác và vai trò của quá trình.** Holgersson và Aaboen (2019) chỉ ra sự dịch chuyển trọng "
    "tâm quản lý tài sản trí tuệ tại đơn vị chuyển giao từ chiếm hữu sang khai thác, hàm ý rằng số văn bằng không đủ để "
    "đánh giá thành công. Perkmann và cộng sự (2013) cho thấy gắn kết học thuật với doanh nghiệp rộng hơn nhiều so với "
    "thương mại hóa theo nghĩa hẹp. Bradley và cộng sự (2013) phê phán mô hình tuyến tính của chuyển giao công nghệ; "
    "Maresova và cộng sự (2019) hệ thống hóa các mô hình, quy trình và vai trò của trường đại học trong quản lý chuyển "
    "giao. Rocha và cộng sự (2023) mô tả bước đánh giá bản khai báo sáng chế theo tính mới, khả năng áp dụng công "
    "nghiệp và trình độ sáng tạo, cho thấy quá trình nhận diện là một khâu có thể và cần được đo lường.",
    "**Mô hình chuỗi kết quả.** Mô hình logic của W. K. Kellogg Foundation (2004) mô tả quan hệ nhân quả giữa nguồn lực, "
    "hoạt động, đầu ra và kết quả của một chương trình, và được dùng rộng rãi để thiết kế chỉ số đánh giá. Ưu điểm của "
    "mô hình là buộc người đánh giá làm rõ giả định về mối liên hệ giữa các mắt xích, qua đó cho phép đặt chỉ số tại "
    "chính các mối liên hệ đó.",
    "**Nghiên cứu trong nước và khoảng trống.** Nguyễn (2025) phân tích cơ hội và thách thức của quản trị tài sản trí tuệ "
    "trong cơ sở giáo dục đại học; Võ (2025) bàn về chính sách đối với tác phẩm hình thành trong nhà trường; tài liệu "
    "tập huấn của Cục Sở hữu trí tuệ (n.d.) cung cấp khung nghiệp vụ cho cán bộ quản lý. Các công trình này chưa xây "
    "dựng bộ chỉ số đánh giá hiệu quả. Tổng hợp lại, còn ba khoảng trống: các khung hiện có thiếu chỉ số đo mức chuyển "
    "hóa giữa các mắt xích; chưa có khung thiết kế cho trường chưa có đơn vị chuyển giao chuyên nghiệp; và chưa có khung "
    "gắn chỉ số với nguồn dữ liệu bắt buộc theo pháp luật Việt Nam hiện hành.",
]

PHUONG_PHAP = [
    "Nghiên cứu sử dụng thiết kế phát triển khung khái niệm, gồm ba bước. Bước thứ nhất tổng hợp có chọn lọc các công "
    "trình về đo lường chuyển giao tri thức, quản lý tài sản trí tuệ đại học và hai bộ chỉ số của Ủy ban châu Âu nêu tại "
    "Mục 2 để xác định các chiều cần đo. Bước thứ hai xây dựng chỉ số theo mô hình chuỗi kết quả của W. K. Kellogg "
    "Foundation (2004), đặt trên chu trình quản lý tài sản trí tuệ bốn giai đoạn tạo lập và nhận diện, xác lập, khai "
    "thác, bảo vệ và phân chia lợi ích (Hình 1); mỗi chỉ số được mô tả bằng tên, cách tính và nguồn dữ liệu.",
    "Bước thứ ba đối chiếu từng chỉ số với sáu văn bản quy định nghĩa vụ báo cáo hoặc công khai: Luật Sở hữu trí tuệ hợp "
    "nhất (Văn phòng Quốc hội, 2026), Luật số 93/2025/QH15, Luật số 125/2025/QH15, Quyết định số 1624/QĐ-TTg, Thông tư "
    "số 01/2024/TT-BGDĐT và Thông tư số 83/2026/TT-BGDĐT. Mỗi chỉ số được xếp vào một trong ba loại dữ liệu: loại A là "
    "dữ liệu bắt buộc báo cáo hoặc công khai theo quy định; loại B là dữ liệu có trong hồ sơ hành chính thông thường của "
    "nhà trường; loại C là dữ liệu nhà trường phải thiết lập công cụ thu thập mới.",
    "Khung được phát triển trong đề tài khoa học công nghệ cấp cơ sở năm 2026 của nhóm tác giả. Bài báo trình bày cấu "
    "trúc khung và lập luận xây dựng, không công bố kết quả áp dụng tại một cơ sở cụ thể.",
]

KQ_1 = [
    "Khung đề xuất đo hiệu quả quản lý quyền sở hữu trí tuệ ở bốn mắt xích và ba mối nối giữa chúng (Hình 2). Đầu vào "
    "phản ánh nguồn lực nhà trường bố trí; quá trình phản ánh mức độ vận hành của thể chế và quy trình; đầu ra phản ánh "
    "sản phẩm có khả năng bảo hộ và quyền đã được xác lập; kết quả phản ánh giá trị mà quyền mang lại cho nhà trường, "
    "tác giả và chỉ số chất lượng.",
    "Điểm mới của khung nằm ở bốn chỉ số chuyển hóa. Tỷ lệ nhận diện đo phần công trình nghiệm thu có sản phẩm được xác "
    "nhận là có khả năng bảo hộ, tức mức độ quá trình nghiên cứu và quá trình rà soát chuyển thành đầu ra tiềm năng. Tỷ "
    "lệ xác lập kịp thời đo phần sản phẩm có khả năng bảo hộ được nộp đơn trong mười hai tháng; mốc này xuất phát từ "
    "khoản 3 Điều 60 Luật Sở hữu trí tuệ, theo đó sáng chế không bị coi là mất tính mới khi đã bộc lộ công khai nếu đơn "
    "được nộp trong thời hạn mười hai tháng kể từ ngày bộc lộ. Tỷ lệ khai thác đo phần tài sản đã xác lập có giao dịch "
    "hoặc được sử dụng, phù hợp với cách nhìn chuyển từ chiếm hữu sang khai thác (Holgersson & Aaboen, 2019). Tỷ suất "
    "khai thác trên chi phí so sánh nguồn thu từ khai thác với chi phí xác lập và duy trì quyền, là thước đo hiệu quả "
    "theo nghĩa hẹp nhất.",
    "Khác với chỉ số đếm, mỗi chỉ số chuyển hóa chỉ vào một mối nối cụ thể. Khi một tỷ lệ thấp, nhà trường biết khâu nào "
    "cần can thiệp, kể cả khi số lượng đầu ra vẫn tăng.",
    "Chuỗi kết quả được đặt chồng lên chu trình quản lý tài sản trí tuệ ở Hình 1 để bảo đảm mỗi giai đoạn của chu trình "
    "đều có ít nhất một chỉ số theo dõi. Giai đoạn tạo lập và nhận diện được đo bằng số công trình có sản phẩm có khả "
    "năng bảo hộ và tỷ lệ nhận diện; giai đoạn xác lập được đo bằng số đơn, số văn bằng, số giấy chứng nhận quyền tác giả "
    "và tỷ lệ xác lập kịp thời; giai đoạn khai thác được đo bằng hợp đồng, nguồn thu và tỷ lệ khai thác; giai đoạn bảo vệ "
    "và phân chia lợi ích được phản ánh qua mức độ tương thích của quy chế về chia lợi ích và động lực của giảng viên. "
    "Cách đặt chồng này tránh tình trạng một giai đoạn quan trọng, thường là giai đoạn nhận diện, không có chỉ số nào và "
    "vì vậy không được quản lý.",
]

KQ_2 = [
    "Bảng 1 trình bày 20 chỉ số, gồm 16 chỉ số ở bốn nhóm và 4 chỉ số chuyển hóa, kèm cách tính, nguồn dữ liệu và loại "
    "dữ liệu. Nhóm đầu vào gồm kinh phí, nhân lực và hạ tầng tra cứu. Nhóm quá trình gồm mức độ tương thích của quy chế "
    "nội bộ với pháp luật, thời gian xử lý hồ sơ, mức độ chuẩn hóa biểu mẫu và tập huấn. Nhóm đầu ra gồm đơn, văn bằng, "
    "giấy chứng nhận quyền tác giả và số công trình có sản phẩm có khả năng bảo hộ. Nhóm kết quả gồm hợp đồng, nguồn "
    "thu, tỷ trọng thu khoa học công nghệ, đóng góp vào kiểm định, xếp hạng và động lực của giảng viên.",
    "Các chỉ số được lựa chọn theo ba nguyên tắc. Thứ nhất, mỗi chỉ số phải gắn với một quyết định quản lý cụ thể, chẳng "
    "hạn thời gian xử lý hồ sơ gắn với việc bố trí nhân lực, tỷ lệ khai thác gắn với việc duy trì hay ngừng duy trì một "
    "văn bằng. Thứ hai, chỉ số phải tính được bằng dữ liệu mà một trường không có đơn vị chuyển giao chuyên nghiệp vẫn có "
    "thể thu thập; vì vậy khung không đưa vào các chỉ số đòi hỏi định giá tài sản trí tuệ hay đo tác động kinh tế xã hội "
    "dài hạn. Thứ ba, chỉ số phải tương thích với định nghĩa trong văn bản pháp luật và chuẩn chất lượng hiện hành để dữ "
    "liệu thu thập cho quản lý nội bộ có thể dùng lại cho báo cáo bắt buộc.",
    "Một số chỉ số được neo trực tiếp vào chuẩn mới. Văn bằng được quy đổi theo chỉ số 6.2.1 của Thông tư số "
    "83/2026/TT-BGDĐT, trong đó bằng độc quyền giải pháp hữu ích có hệ số 3 và bằng độc quyền sáng chế có hệ số 5; tỷ "
    "trọng thu được tính theo chỉ số 6.1 với ngưỡng không thấp hơn 5% đối với cơ sở có đào tạo tiến sĩ. Mức độ tương "
    "thích của quy chế gắn với chỉ số 1.1 về tỷ lệ nội dung quản trị nội bộ bắt buộc đã ban hành, trong đó quy định về "
    "sở hữu trí tuệ, liêm chính khoa học, liêm chính học thuật là một nội dung bắt buộc chung. Việc neo chỉ số vào chuẩn "
    "giúp một lần thu thập dữ liệu phục vụ được cả hai mục đích: tự đánh giá nội bộ và báo cáo theo Chuẩn cơ sở giáo dục "
    "đại học. Điều này đặc biệt quan trọng với trường định hướng ứng dụng, khi việc nâng hệ số của bằng độc quyền giải "
    "pháp hữu ích lên 3 làm cho loại văn bằng phù hợp nhất với quy mô đề tài cấp cơ sở có đóng góp đáng kể vào chỉ số "
    "6.2.1.",
]

KQ_3 = [
    "Bảng 2 tổng hợp loại dữ liệu theo nhóm chỉ số. Trong 20 chỉ số, 5 chỉ số có dữ liệu loại A, 7 chỉ số loại B và 8 "
    "chỉ số loại C. Dữ liệu bắt buộc tập trung ở nhóm kết quả với 3 trên 5 chỉ số, gồm hợp đồng chuyển giao theo khoản 4 "
    "Điều 27 Luật số 93/2025/QH15, nguồn thu và tỷ trọng thu theo Thông tư số 83/2026/TT-BGDĐT. Ngược lại, không chỉ số "
    "chuyển hóa nào có dữ liệu bắt buộc, và 3 trên 4 chỉ số chuyển hóa thuộc loại C.",
    "Phát hiện này có hàm ý quan trọng: hệ thống báo cáo bắt buộc được thiết kế để kiểm tra kết quả, không để chẩn đoán "
    "nguyên nhân. Một trường chỉ đáp ứng nghĩa vụ báo cáo sẽ biết mình có bao nhiêu văn bằng và nguồn thu, nhưng không "
    "biết bao nhiêu sản phẩm có khả năng bảo hộ đã bị bỏ qua, hay bao nhiêu sản phẩm đã mất tính mới vì công bố trước "
    "khi nộp đơn. Hai chỉ số then chốt CH1 và CH2 đều phụ thuộc vào một công cụ duy nhất là phiếu rà soát khả năng bảo "
    "hộ tại nghiệm thu; thiếu công cụ này, khung đánh giá mất phần có giá trị chẩn đoán nhất.",
    "Nhóm dữ liệu loại B cũng cần được lưu ý. Các dữ liệu này đã tồn tại trong hồ sơ hành chính nhưng thường phân tán ở "
    "nhiều đơn vị: kinh phí ở bộ phận tài chính, nhân lực ở bộ phận tổ chức, đơn và văn bằng ở bộ phận pháp chế hoặc "
    "quản trị thương hiệu, đề tài ở đơn vị quản lý khoa học. Khi không có một danh mục tài sản trí tuệ dùng chung, việc "
    "tính các chỉ số đòi hỏi đối chiếu thủ công giữa nhiều nguồn. Khoản 4 Điều 27 Luật số 93/2025/QH15 yêu cầu tổ chức "
    "chủ trì công khai thông tin và báo cáo kết quả thương mại hóa, còn điểm đ khoản 3 Điều 28 Luật số 125/2025/QH15 yêu "
    "cầu công khai kết quả hoạt động khoa học, công nghệ và đổi mới sáng tạo hằng năm; cả hai nghĩa vụ này đều ngầm định "
    "nhà trường phải có một danh mục tài sản trí tuệ thống nhất. Danh mục đó vì vậy vừa là điều kiện để tính bộ chỉ số, "
    "vừa là điều kiện để thực hiện nghĩa vụ pháp lý.",
]

KQ_4 = [
    "Bảng 3 đề xuất năm mẫu hình chẩn đoán dựa trên cách kết hợp các chỉ số. Mỗi mẫu hình gắn với một điểm nghẽn có thể "
    "xảy ra và công cụ can thiệp có căn cứ pháp lý. Chẳng hạn, khi tỷ lệ nhận diện cao nhưng tỷ lệ xác lập kịp thời thấp, "
    "điểm nghẽn nằm ở khâu nối giữa nghiệm thu và đăng ký, nơi cần đầu mối tiếp nhận và dòng kinh phí nộp đơn; điểm b "
    "khoản 2 Điều 66 Luật số 93/2025/QH15 cho phép quỹ phát triển khoa học và công nghệ chi cho đăng ký, bảo hộ, quản lý, "
    "khai thác quyền sở hữu trí tuệ. Khi tỷ lệ xác lập cao nhưng tỷ lệ khai thác thấp, điểm nghẽn chuyển sang thị trường "
    "và định giá, nơi Điều 27 Luật này trao cho tổ chức quyền tự quyết về hình thức, giá và phân chia lợi nhuận.",
    "Vì CH1 và CH2 phụ thuộc vào phiếu rà soát tại nghiệm thu, bài báo đề xuất nội dung tối thiểu của phiếu gồm năm câu "
    "hỏi: sản phẩm cụ thể của đề tài là gì; sản phẩm thuộc đối tượng quyền nào trong các nhóm sáng chế, giải pháp hữu "
    "ích, kiểu dáng công nghiệp, quyền tác giả, sưu tập dữ liệu, bí mật kinh doanh; sản phẩm đã được bộc lộ công khai hay "
    "chưa và vào ngày nào; chủ sở hữu được xác định theo nguồn kinh phí như thế nào theo Điều 25 Luật số 93/2025/QH15; "
    "và hội đồng đề xuất nộp đơn, đăng ký quyền tác giả, giữ bí mật hay không bảo hộ. Ngày bộc lộ là thông tin then "
    "chốt, vì nó xác định thời hạn còn lại để nộp đơn và là căn cứ tính CH2.",
    "Với trường mới bắt đầu theo dõi, chưa có ngưỡng tham chiếu bên ngoài cho các chỉ số chuyển hóa. Cách sử dụng phù hợp "
    "là so sánh với chính nhà trường qua các năm và giữa các khối ngành trong trường. Riêng CH2 có thể đặt mục tiêu gần "
    "mức tuyệt đối đối với các sản phẩm đã được hội đồng xác nhận có khả năng bảo hộ, vì mỗi sản phẩm không được nộp "
    "đơn trong thời hạn là một tài sản tiềm năng bị mất vĩnh viễn.",
    "Từ mức độ sẵn có của dữ liệu, khung có thể triển khai theo ba giai đoạn. Giai đoạn đầu sử dụng các chỉ số loại A và "
    "B, vốn đã có hoặc dễ tổng hợp, để lập đường cơ sở. Giai đoạn hai đưa phiếu rà soát vào quy trình nghiệm thu và sổ "
    "theo dõi hồ sơ để tính được R4, CH1, CH2 và Q2. Giai đoạn ba bổ sung khảo sát giảng viên và danh mục khai thác để "
    "tính K5 và CH3. Cách triển khai này cho phép nhà trường bắt đầu ngay mà không phải chờ hoàn thiện toàn bộ hệ thống "
    "dữ liệu.",
]

BAN_LUAN = [
    "**Giải thích kết quả.** Việc dữ liệu bắt buộc tập trung ở nhóm kết quả phản ánh mục đích của chuẩn chất lượng là "
    "kiểm tra mức đáp ứng tối thiểu, không phải hỗ trợ quản lý nội bộ. Với trường định hướng ứng dụng, nơi phần lớn sản "
    "phẩm có khả năng bảo hộ là giải pháp hữu ích, công thức, quy trình phát sinh từ đề tài cấp cơ sở, giá trị nằm ở việc "
    "không để sản phẩm bị bỏ sót ở khâu nhận diện. Vì vậy, chỉ số chuyển hóa quan trọng hơn chỉ số đếm đối với nhóm "
    "trường này. Trường tư thục còn chịu ràng buộc chặt về nguồn lực, nên cần biết mỗi khoản chi cho xác lập và duy "
    "trì quyền mang lại gì; tỷ suất khai thác trên chi phí và tỷ lệ khai thác giúp quyết định nên tiếp tục đầu tư vào "
    "loại tài sản nào, thay vì mở rộng danh mục một cách dàn trải.",
    "**Đối chiếu với nghiên cứu trước.** Khung kế thừa nhận định của Thursby và Kemp (2002) và Siegel và cộng sự (2007) "
    "rằng hiệu quả phải được đo bằng quan hệ giữa đầu ra và nguồn lực. Điểm khác là thay cho phép đo so sánh giữa nhiều "
    "trường, khung dùng các tỷ lệ nội bộ mà một trường có thể tự tính và theo dõi qua thời gian. So với tập chỉ số cốt "
    "lõi của Finne và cộng sự (2009) và Campbell và cộng sự (2020), khung giữ các chỉ số đếm để bảo đảm khả năng so sánh "
    "nhưng bổ sung chỉ số quá trình và chỉ số chuyển hóa, phù hợp với lập luận của Holgersson và Aaboen (2019) về khai "
    "thác và của Rocha và cộng sự (2023) về đánh giá bản khai báo. So với các nghiên cứu trong nước (Nguyễn, 2025; Võ, "
    "2025), bài báo chuyển từ phân tích định tính sang đề xuất công cụ đo lường cụ thể.",
    "**Đóng góp của bài báo.** Về lý luận, bài báo vận dụng mô hình chuỗi kết quả vào quản lý quyền sở hữu trí tuệ đại "
    "học và đề xuất đặt chỉ số tại các mối nối, không chỉ tại các mắt xích. Về thực tiễn, bộ 20 chỉ số, cách phân loại "
    "dữ liệu và năm mẫu hình chẩn đoán là công cụ tự đánh giá mà trường đại học định hướng ứng dụng có thể sử dụng khi "
    "chuẩn bị đánh giá theo Thông tư số 83/2026/TT-BGDĐT và thực hiện nghĩa vụ công khai theo Luật số 125/2025/QH15. Về "
    "chính sách, kết quả gợi ý cơ quan quản lý có thể cân nhắc bổ sung chỉ số về tỷ lệ sản phẩm có khả năng bảo hộ được "
    "nộp đơn vào hệ thống dữ liệu báo cáo.",
    "**Hạn chế của nghiên cứu.** Thứ nhất, khung được xây dựng bằng lập luận lý thuyết và đối chiếu văn bản, chưa được "
    "kiểm định thực nghiệm về độ tin cậy và giá trị. Thứ hai, bài báo chưa xác định trọng số giữa các chỉ số, nên chưa "
    "thể tổng hợp thành một chỉ số chung. Thứ ba, việc phân loại dữ liệu phụ thuộc vào văn bản hướng dẫn tính chỉ số và "
    "danh mục dữ liệu báo cáo; khi các hướng dẫn này thay đổi, một số chỉ số có thể chuyển loại.",
    "**Hướng nghiên cứu tiếp theo.** Các nghiên cứu tiếp theo có thể lấy ý kiến chuyên gia bằng phương pháp Delphi để "
    "hoàn thiện định nghĩa và trọng số, áp dụng thử khung tại một số trường đại học định hướng ứng dụng để kiểm định "
    "khả năng tính toán, và theo dõi các chỉ số chuyển hóa qua nhiều năm để đánh giá tác động của các can thiệp như phiếu "
    "rà soát tại nghiệm thu.",
]

KET_LUAN = [
    "Bài báo đề xuất khung đánh giá hiệu quả quản lý quyền sở hữu trí tuệ cho trường đại học định hướng ứng dụng trong bối "
    "cảnh chỉ số sở hữu trí tuệ trở thành căn cứ đánh giá cơ sở giáo dục đại học.",
    "Kết quả chính là bộ 20 chỉ số theo chuỗi đầu vào, quá trình, đầu ra, kết quả, trong đó bốn chỉ số chuyển hóa đo mức "
    "độ các mắt xích nối tiếp nhau. Đối chiếu với pháp luật hiện hành cho thấy dữ liệu bắt buộc tập trung ở nhóm kết "
    "quả, còn các chỉ số chuyển hóa, vốn có giá trị chẩn đoán cao nhất, đều do nhà trường tự thiết lập. Hai chỉ số then "
    "chốt về nhận diện và xác lập kịp thời phụ thuộc vào phiếu rà soát khả năng bảo hộ tại nghiệm thu.",
    "Bài báo đóng góp một công cụ tự đánh giá cấp trường, kèm mẫu hình chẩn đoán điểm nghẽn và lộ trình triển khai theo "
    "mức độ sẵn có của dữ liệu. Hàm ý cho các trường đại học định hướng ứng dụng là không dừng ở việc đáp ứng chỉ số bắt "
    "buộc, mà cần thiết lập sớm các công cụ thu thập dữ liệu ở khâu nhận diện để biết tiềm năng tài sản trí tuệ đang bị "
    "mất ở đâu.",
]

BANG_3 = dict(
    tieu_de="Các mẫu hình chẩn đoán điểm nghẽn từ bộ chỉ số",
    cot=["Mẫu hình chỉ số", "Điểm nghẽn có thể xảy ra", "Công cụ can thiệp và căn cứ"],
    dong=[["CH1 thấp", "Ít sản phẩm có khả năng bảo hộ, hoặc hội đồng nghiệm thu chưa nhận diện",
           "Xác định đối tượng quyền từ thuyết minh đề tài; tập huấn; phiếu rà soát tại nghiệm thu"],
          ["CH1 cao, CH2 thấp", "Khâu nối giữa nghiệm thu và đăng ký: thiếu đầu mối, kinh phí, công bố trước khi nộp đơn",
           "Đầu mối tiếp nhận; dòng chi từ quỹ phát triển khoa học và công nghệ (Điều 66 Luật số 93/2025/QH15); khai "
           "báo trước công bố (Điều 60 Luật Sở hữu trí tuệ)"],
          ["CH2 cao, CH3 thấp", "Khai thác: thiếu kết nối thị trường, định giá", "Hợp tác doanh nghiệp; tự quyết "
                                                                                    "thương mại hóa (Điều 27 Luật số "
                                                                                    "93/2025/QH15); doanh nghiệp quản "
                                                                                    "lý tài sản trí tuệ (Điều 28 Luật "
                                                                                    "số 125/2025/QH15)"],
          ["CH3 cao, CH4 thấp", "Thu không bù chi do danh mục dàn trải, chi phí duy trì cao hoặc giá chuyển giao thấp",
           "Rà soát danh mục định kỳ; ngừng duy trì tài sản không khai thác; chuẩn hóa định giá"],
          ["Q1 thấp", "Quy chế chưa cập nhật theo luật mới, nhiều văn bản cùng điều chỉnh",
           "Hợp nhất quy chế theo Luật số 93/2025/QH15; chỉ số 1.1 Thông tư số 83/2026/TT-BGDĐT"]],
    nguon="Nguồn: Nhóm tác giả đề xuất.",
    rong=[3.0, 5.8, 6.6], can=["left", "left", "left"],
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
    ("p", "Nhóm tác giả cam kết nội dung trong bài là trung thực, chính xác; tài liệu và văn bản được sử dụng là tài liệu "
          "công bố và văn bản ban hành công khai, được nhóm tác giả lưu trữ và sẵn sàng cung cấp khi Ban Biên tập yêu cầu; "
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
    v.so_do("Chu trình quản lý tài sản trí tuệ trong trường đại học", os.path.join(SO_DO, "chu_trinh.png"),
            "Nguồn: Nhóm tác giả xây dựng trên cơ sở Bradley và cộng sự (2013) và Tổ chức Sở hữu trí tuệ thế giới "
            "(2020).", tien_to="", rong_cm=13.0)
    phan("Phương pháp", PHUONG_PHAP[1:])

    v.doan("h1", "4. KẾT QUẢ NGHIÊN CỨU", bold=True)
    v.doan("h2", "4.1. Hiệu quả cần được đo ở bốn mắt xích và tại các mối nối giữa chúng")
    phan("Kết quả", KQ_1[:1])
    v.so_do("Chuỗi kết quả và bốn chỉ số chuyển hóa trong đánh giá hiệu quả quản lý quyền sở hữu trí tuệ",
            os.path.join(SO_DO, "chuoi_ket_qua.png"), "Nguồn: Nhóm tác giả đề xuất trên cơ sở W. K. Kellogg Foundation "
                                                       "(2004).", tien_to="", rong_cm=15.5)
    phan("Kết quả", KQ_1[1:])
    v.doan("h2", "4.2. Bộ 20 chỉ số gắn với chuẩn chất lượng mới")
    phan("Kết quả", KQ_2[:1])
    v.bang("Bộ chỉ số đánh giá hiệu quả quản lý quyền sở hữu trí tuệ theo chuỗi kết quả",
           ["Mã", "Chỉ số", "Cách tính", "Nguồn dữ liệu", "Loại"],
           [[c[0], c[2], c[3], c[4], c[5]] for c in CHI_SO],
           "Nguồn: Nhóm tác giả đề xuất. Loại dữ liệu: A là dữ liệu bắt buộc báo cáo hoặc công khai; B là dữ liệu có trong "
           "hồ sơ hành chính; C là dữ liệu cần công cụ thu thập mới.",
           [1.2, 4.2, 4.6, 4.0, 1.2], can=["center", "left", "left", "left", "center"],
           hang_dam=(), tien_to="")
    phan("Kết quả", KQ_2[1:])
    v.doan("h2", "4.3. Dữ liệu bắt buộc tập trung ở nhóm kết quả, chỉ số chuyển hóa phải tự thiết lập")
    phan("Kết quả", KQ_3[:1])
    v.bang("Phân bố loại dữ liệu theo nhóm chỉ số", ["Nhóm chỉ số", "Loại A", "Loại B", "Loại C", "Tổng"],
           [[n, str(DEM[n]["A"]), str(DEM[n]["B"]), str(DEM[n]["C"]), str(sum(DEM[n].values()))] for n in NHOM] +
           [["Tổng", str(TONG["A"]), str(TONG["B"]), str(TONG["C"]), str(len(CHI_SO))]],
           "Nguồn: Nhóm tác giả tổng hợp từ Bảng 1.", [5.0, 2.5, 2.5, 2.5, 2.5], dong_tong=True, tien_to="")
    phan("Kết quả", KQ_3[1:])
    v.doan("h2", "4.4. Mẫu hình chẩn đoán điểm nghẽn và lộ trình triển khai")
    phan("Kết quả", KQ_4[:1])
    b = BANG_3
    v.bang(b["tieu_de"], b["cot"], b["dong"], b["nguon"], b["rong"], can=b["can"], tien_to="")
    phan("Kết quả", KQ_4[1:])

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

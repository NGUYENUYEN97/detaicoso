# -*- coding: utf-8 -*-
"""Chương 3 bản cuối.

Đầu vào: Chuong_3_He_thong_giai_phap_ra_soat.docx (bản đã rà soát, chấp nhận toàn bộ theo dõi thay đổi).
Bản cuối:
  - thay Mục 3.1.2 bằng phân tích điểm mạnh, điểm yếu, thời cơ, thách thức đầy đủ, kèm ma trận và bảng phương án kết hợp;
  - khớp các nhận định với Chương 2 bản cuối (bốn đầu mối, kinh phí đề tài, cơ chế chuyển tiếp với Quỹ, thưởng văn bằng);
  - lộ trình gắn với mốc hiệu lực của Thông tư 83/2026/TT-BGDĐT và kỳ tự đánh giá tháng 5 năm 2027;
  - đánh số lại bảng, chuẩn hóa tên đơn vị trong bảng, thêm tiểu kết.
"""
import os
import re
import subprocess
import sys
import tempfile

import docx
from docx.table import Table
from docx.text.paragraph import Paragraph

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from khung import GOC, VanBanChung, luu  # noqa: E402

VAO = os.path.join(GOC, "Chuong_3_He_thong_giai_phap_ra_soat.docx")
ACCEPT = ("/root/.claude/skills/synced/511a21b4-da62-4b21-9cb6-da109b6fbd84_b74c1dbc-fffb-4c9d-9a90-c57409abedcc/docx/"
          "scripts/accept_changes.py")

TIEN_TO_DAM = re.compile(r"^(- )?(Thứ (nhất|hai|ba|tư|năm|sáu|bảy), |Yêu cầu \d: |Nguyên tắc \d: |Khâu \d: |Nội dung \d: |"
                         r"Về [^:]{3,60}: )")


def doc_khoi():
    """Đọc bản rà soát đã chấp nhận thay đổi thành danh sách khối (loại, nội dung)."""
    tmp = tempfile.mkdtemp()
    ra = os.path.join(tmp, "c3.docx")
    subprocess.run(["python3", ACCEPT, VAO, ra], check=True, capture_output=True)
    d = docx.Document(ra)
    khoi = []
    for el in d.element.body.iterchildren():
        tag = el.tag.split("}")[1]
        if tag == "p":
            p = Paragraph(el, d)
            t = p.text.strip()
            if not t or re.match(r"^Bảng 3\.\d\. ", t):
                continue
            ds_dam = [r.bold for r in p.runs if r.text.strip()]
            khoi.append(["p", t, bool(ds_dam) and all(ds_dam)])
        elif tag == "tbl":
            khoi.append(["tbl", Table(el, d), None])
    return khoi


SUA = [
    # Luật 93/2025/QH15 (toàn văn), Quyết định 217 ngày 21/11/2024, văn phong xây dựng
    ("là quy định nội bộ tự đặt ra của Nhà trường, cần được bãi bỏ để bảo đảm tương thích với pháp luật hiện hành và tạo "
     "động lực thực sự cho hoạt động sáng tạo.",
     "được xây dựng năm 2021 theo khung pháp luật khi đó, nay cần được cập nhật để tương thích với pháp luật hiện hành và "
     "tạo động lực thực sự cho hoạt động sáng tạo. Bên cạnh đó, khoản 2 Điều 25 Luật này quy định tổ chức chủ trì nhiệm vụ "
     "sử dụng ngân sách nhà nước được Nhà nước tự động giao quyền quản lý, sử dụng, quyền sở hữu phần kết quả tương ứng, "
     "không phải bồi hoàn chi phí; Điều 27 cho phép tổ chức được giao quyền tự quyết định hình thức, giá và phân chia lợi "
     "nhuận khi thương mại hóa; điểm a khoản 3 Điều 28 quy định thưởng cho tác giả tối thiểu 30% lợi nhuận thu được từ "
     "thương mại hóa phần kết quả sử dụng ngân sách nhà nước; điểm b khoản 2 Điều 66 cho phép quỹ phát triển khoa học và "
     "công nghệ của tổ chức chi cho đăng ký, bảo hộ, quản lý, khai thác quyền sở hữu trí tuệ."),
    ("Quyết định số 217/QĐ-ĐHTĐ ban hành Quy chế quản trị tài sản trí tuệ năm 2024, Quy chế chi tiêu nội bộ",
     "Quyết định số 217/QĐ-ĐHTĐ ngày 21 tháng 11 năm 2024 ban hành Quy chế quản trị tài sản trí tuệ, Quy chế chi tiêu nội "
     "bộ"),
    ("với năm quy định khác nhau về phân chia lợi ích, dẫn đến sự chồng chéo trong áp dụng thực tế.",
     "với năm quy định khác nhau về phân chia lợi ích. Các văn bản này được ban hành ở những thời điểm và cho những kênh "
     "tài trợ khác nhau, trước khi Luật số 93/2025/QH15 và Luật số 131/2025/QH15 có hiệu lực, nay cần được hợp nhất để "
     "thống nhất cách áp dụng."),
    ("15% số tiền nhận được mỗi lần khi chuyển giao quyền sử dụng.",
     "15% số tiền nhận được mỗi lần khi chuyển giao quyền sử dụng; đối với kết quả sử dụng ngân sách nhà nước như ba đề "
     "tài cấp quốc gia, bảo đảm mức thưởng cho tác giả không thấp hơn 30% lợi nhuận theo điểm a khoản 3 Điều 28 Luật số "
     "93/2025/QH15."),
    ("điểm d khoản 3 Điều 28 Luật Giáo dục đại học số 125/2025/QH15. Nguồn kinh phí này",
     "điểm d khoản 3 Điều 28 Luật Giáo dục đại học số 125/2025/QH15, phù hợp với điểm b khoản 2 Điều 66 Luật số "
     "93/2025/QH15 cho phép quỹ phát triển khoa học và công nghệ của tổ chức chi cho đăng ký, bảo hộ, quản lý, khai thác "
     "quyền sở hữu trí tuệ. Nguồn kinh phí này"),
    ("Thực trạng cho thấy việc đăng ký bảo hộ sở hữu trí tuệ tại Nhà trường hiện nay mang tính tự phát, phụ thuộc vào sự "
     "chủ động của từng cá nhân giảng viên. Quy trình tiếp nhận - xử lý - nộp đơn chưa được chuẩn hóa,",
     "Thực trạng tại Chương 2 cho thấy tiềm năng tài sản trí tuệ của đội ngũ là rất lớn, đặc biệt ở khối ngành Y - Dược, "
     "nhưng quy trình hiện hành còn một khoảng trống kỹ thuật: chưa có biểu mẫu rà soát khả năng bảo hộ tại thời điểm "
     "nghiệm thu, nên việc đăng ký phụ thuộc vào sự chủ động của từng tác giả. Quy trình tiếp nhận, xử lý, nộp đơn vì vậy "
     "cần được chuẩn hóa để tránh"),
    ("dẫn đến nguy cơ bỏ sót các kết quả nghiên cứu có khả năng bảo hộ,", "nguy cơ bỏ sót các kết quả nghiên cứu có khả "
     "năng bảo hộ,"),
    ("cho thấy năng lực nhận diện tài sản trí tuệ trong đội ngũ còn hạn chế:",
     "cho thấy đội ngũ cần được hỗ trợ thêm về kỹ năng nhận diện và bảo hộ tài sản trí tuệ:"),
    ("Thiết lập Cơ chế phối hợp liên phòng ban theo mô hình tam giác vận hành. Trong mô hình này, ba trục chức năng phối hợp "
     "chặt chẽ theo quy chế liên thông:",
     "Thiết lập Cơ chế phối hợp liên phòng ban theo mô hình tam giác vận hành. Trong mô hình này, ba trục chức năng phối hợp "
     "chặt chẽ theo quy chế liên thông, với hai đơn vị phối hợp, như thể hiện tại hình dưới đây.\n[[HINH_PHOI_HOP]]"),
    ("áp dụng cho tất cả các đề tài cấp cơ sở và các nhiệm vụ khoa học công nghệ sử dụng ngân sách:",
     "áp dụng cho tất cả các đề tài cấp cơ sở và các nhiệm vụ khoa học công nghệ sử dụng ngân sách, được tóm tắt tại hình "
     "dưới đây và mô tả cụ thể sau đó.\n[[HINH_TAM_KHAU]]"),
    ("và các mốc chiến lược quốc gia về sở hữu trí tuệ:",
     "và các mốc chiến lược quốc gia về sở hữu trí tuệ, được tóm tắt tại hình dưới đây.\n[[HINH_LO_TRINH]]"),
    ("phân tích tại Hình 2.9 Chương 2.", "phân tích tại Hình 2.{SO_HINH_KH} Chương 2."),
    # 3.1.3
    ("Dữ liệu thực trạng tại Chương 2 phản ánh khoảng trống lớn về quy trình và động lực: hệ số Gini về phân bố công bố "
     "khoa học lên tới 0,829, cả giai đoạn năm năm mới phát sinh một đơn sáng chế từ đề tài cơ sở. Thực trạng này cho thấy "
     "không chỉ là vấn đề quy trình hay thủ tục, mà là vấn đề hệ thống - từ tư duy quản trị, động lực tài chính, năng lực "
     "chuyên trách đến văn hóa sáng tạo. Để khắc phục triệt để các nguyên nhân này, hệ thống giải pháp được xây dựng dựa "
     "trên năm nguyên tắc chỉ đạo:",
     "Dữ liệu thực trạng tại Chương 2 cho thấy điểm nghẽn nằm ở khâu nối giữa nghiệm thu và đăng ký: 11 trên 38 đề tài có "
     "sản phẩm đủ điều kiện xác lập quyền nhưng chỉ 1 đề tài được nộp đơn, trong khi năng lực công bố tập trung ở một nhóm "
     "nhỏ với hệ số Gini 0,829. Điểm nghẽn này mang tính hệ thống, liên quan đồng thời đến quy chế, quy trình, động lực tài "
     "chính, năng lực chuyên trách và dữ liệu. Kết hợp với phân tích tại Mục 3.1.2, hệ thống giải pháp được xây dựng dựa "
     "trên năm nguyên tắc chỉ đạo:"),
    ("tránh tình trạng đề xuất giải pháp chung chung không gắn với căn cứ pháp lý.",
     "đồng thời chuẩn bị cho các yêu cầu của Thông tư số 83/2026/TT-BGDĐT, tránh tình trạng đề xuất giải pháp chung chung "
     "không gắn với căn cứ pháp lý."),
    ("mà kế thừa và phát huy vai trò hiện có của Bộ phận Pháp chế (thuộc Trung tâm Dịch vụ và Quản trị hành chính tổng "
     "hợp) - đơn vị hiện đang trực tiếp đảm nhiệm công tác soạn thảo quy chế, tiếp nhận hồ sơ, thực hiện thủ tục xác lập "
     "quyền và lưu trữ bằng chứng nhận tài sản trí tuệ.",
     "mà kế thừa sự phân công đã có tại Điều 11 Quyết định 217: Phòng Khoa học Công nghệ nhận diện, lập hồ sơ theo dõi và "
     "xúc tiến thương mại hóa tài sản trí tuệ; Bộ phận Pháp chế thuộc Trung tâm Dịch vụ và Quản trị hành chính tổng hợp "
     "thực hiện thủ tục xác lập quyền và là đầu mối của Mạng lưới Trung tâm Hỗ trợ công nghệ và đổi mới sáng tạo."),
    # 3.2
    ("đến phát triển văn hóa và đào tạo (nền tảng nhận thức).",
     "đến phát triển văn hóa và đào tạo (nền tảng nhận thức). Mỗi nhóm giải pháp tương ứng với một hoặc nhiều phương án kết "
     "hợp tại Bảng 3.2."),
    # Giải pháp 1
    ("Bố trí nguồn kinh phí tư vấn pháp lý chuyên sâu từ ngân sách thường xuyên cho việc thuê chuyên gia sở hữu trí tuệ tham "
     "gia rà soát dự thảo Quy chế. Rà soát đồng bộ Quy chế chi tiêu nội bộ và Điều lệ Quỹ Học bổng sau tiến sĩ Ngô Xuân Độ để "
     "bảo đảm thống nhất về cơ chế tài chính hỗ trợ hoạt động sở hữu trí tuệ. Thời gian hoàn thành dự kiến: Quý IV năm 2026.",
     "Bố trí nguồn kinh phí tư vấn pháp lý chuyên sâu từ ngân sách thường xuyên cho việc thuê chuyên gia sở hữu trí tuệ tham "
     "gia rà soát dự thảo Quy chế. Rà soát đồng bộ Quy chế chi tiêu nội bộ và Điều lệ Quỹ Học bổng sau tiến sĩ Ngô Xuân Độ để "
     "bảo đảm thống nhất về cơ chế tài chính hỗ trợ hoạt động sở hữu trí tuệ. Thời gian hoàn thành dự kiến: trình dự thảo "
     "trong quý IV năm 2026 và ban hành trước ngày 31 tháng 5 năm 2027, thời hạn công bố kết quả tự đánh giá Chuẩn cơ sở giáo "
     "dục đại học năm đầu áp dụng Thông tư số 83/2026/TT-BGDĐT."),
    # Giải pháp 2
    ("chức năng quản lý sở hữu trí tuệ hiện đang được phân tán ở nhiều đầu mối: Bộ phận Pháp chế trực tiếp thực hiện các "
     "thủ tục pháp lý và lưu trữ, Phòng Khoa học Công nghệ quản lý chỉ tiêu nghiên cứu, Trung tâm Tuyển sinh và Quản trị "
     "thương hiệu quản lý nhãn hiệu.",
     "chức năng quản lý sở hữu trí tuệ hiện đang được phân tán ở bốn đầu mối: Phòng Khoa học Công nghệ tiếp nhận hồ sơ nhưng "
     "chỉ có 2 nhân sự, Bộ phận Pháp chế thực hiện thủ tục và giữ đầu mối Mạng lưới Trung tâm Hỗ trợ công nghệ và đổi mới "
     "sáng tạo, Trung tâm Tuyển sinh và Quản trị thương hiệu xác lập quyền đối với nhãn hiệu, Viện Nghiên cứu giáo dục và "
     "Chuyển giao tri thức nắm khâu khai thác."),
    # Giải pháp 4
    ("Chi phí nộp đơn đăng ký sáng chế tại Cục Sở hữu trí tuệ dao động từ 3 đến 5 triệu đồng (chưa bao gồm phí thuê đại diện "
     "sở hữu công nghiệp), trong khi Điều 38 Quyết định 213 không có mục chi riêng cho lệ phí đăng ký và giảng viên phải chờ "
     "đợi quy trình xét duyệt kéo dài từ 18 đến 36 tháng mới nhận được kết quả.",
     "Theo kết quả Chương 2, đề tài cấp cơ sở có mức kinh phí trung vị 8,5 triệu đồng và không có dòng chi cho lệ phí nộp "
     "đơn, phí thẩm định, phí đại diện sở hữu công nghiệp và phí duy trì hiệu lực; Điều 38 Quyết định 213 không có mục chi "
     "riêng cho các khoản này; trong khi đó văn bằng chỉ được quy đổi giờ, không có tiền thưởng và thủ tục thẩm định thường "
     "kéo dài từ hai đến ba năm."),
    ("Mức trích lập dự kiến từ 5% đến 10% tổng Quỹ Phát triển khoa học công nghệ hằng năm.",
     "Mức trích lập dự kiến từ 5% đến 10% tổng Quỹ Phát triển khoa học công nghệ hằng năm. Đồng thời sửa đổi Điều lệ Quỹ "
     "Học bổng sau tiến sĩ Ngô Xuân Độ, ngân sách 5 tỷ đồng giai đoạn 2025 - 2029, để sản phẩm hình thành từ đề tài cấp cơ "
     "sở đã qua rà soát tại Khâu 4 được đề nghị hỗ trợ chi phí đăng ký, tạo cơ chế chuyển tiếp giữa hai kênh tài trợ."),
    ("Cơ chế này tạo động lực tức thì cho giảng viên.",
     "Bổ sung mức thưởng bằng tiền cho bằng độc quyền sáng chế và giải pháp hữu ích, tối thiểu tương đương mức thưởng cho "
     "bài báo WoS có cùng mức giờ quy đổi. Cơ chế này tạo động lực tức thì cho giảng viên và xóa bỏ sự chênh lệch động lực "
     "giữa công bố và đăng ký."),
    # Giải pháp 5
    ("sẵn sàng kết nối tự động với Nền tảng số quốc gia theo Điều 28 Luật Giáo dục đại học số 125/2025/QH15.",
     "sẵn sàng kết nối tự động với Nền tảng số quốc gia theo điểm đ khoản 3 Điều 28 Luật Giáo dục đại học số 125/2025/QH15 "
     "và cung cấp dữ liệu cho báo cáo trên HEMIS."),
    ("Viện Nghiên cứu phối hợp tổ chức Hội thảo chuyên đề.",
     "Viện Nghiên cứu giáo dục và Chuyển giao tri thức phối hợp tổ chức Hội thảo chuyên đề."),
    # 3.4.1
    ("Giai đoạn 1 (Năm 2026): Chuẩn hóa thể chế và kích hoạt phong trào",
     "Giai đoạn 1, từ quý IV năm 2026 đến quý II năm 2027: Chuẩn hóa thể chế và kích hoạt phong trào"),
    ("- Ban hành Quy chế quản lý sở hữu trí tuệ hợp nhất sửa đổi, trong đó bãi bỏ mức trần thù lao 100 triệu đồng.",
     "- Ban hành Quy chế quản lý sở hữu trí tuệ hợp nhất trước ngày 31 tháng 5 năm 2027, trong đó bãi bỏ mức trần thù lao "
     "100 triệu đồng và bổ sung nội dung liêm chính khoa học, liêm chính học thuật."),
    ("- Chuẩn bị nội dung tham gia Đề án tăng cường hoạt động sở hữu trí tuệ của cơ sở giáo dục đại học do Bộ Giáo dục và "
     "Đào tạo dự thảo trong quý II năm 2027.", None),
    ("- Tổ chức đợt tập huấn đầu tiên về sở hữu trí tuệ cho toàn thể giảng viên.",
     "- Tổ chức đợt tập huấn đầu tiên về sở hữu trí tuệ cho toàn thể giảng viên.\n"
     "- Lập danh mục số tài sản trí tuệ phục vụ báo cáo trên HEMIS theo Thông tư số 83/2026/TT-BGDĐT.\n"
     "- Chuẩn bị nội dung tham gia Đề án tăng cường hoạt động sở hữu trí tuệ của cơ sở giáo dục đại học do Bộ Giáo dục và "
     "Đào tạo dự thảo trong quý II năm 2027."),
    ("Giai đoạn 2 (Năm 2027): Thí điểm thương mại hóa và ươm tạo doanh nghiệp",
     "Giai đoạn 2, từ quý III năm 2027 đến hết năm 2028: Thí điểm thương mại hóa và ươm tạo doanh nghiệp"),
    ("Giai đoạn 3 (Giai đoạn 2028 - 2030): Vận hành toàn diện và tích hợp hệ sinh thái",
     "Giai đoạn 3, giai đoạn 2029 - 2030: Vận hành toàn diện và tích hợp hệ sinh thái"),
    # 3.4.2
    ("được tổng hợp tại Bảng 3.2 dưới đây.", "được tổng hợp tại Bảng 3.3."),
    ("Ban Giám hiệu giao Phòng Khoa học Công nghệ là đơn vị đầu mối tổng hợp, theo dõi và báo cáo kết quả thực hiện các chỉ "
     "số hằng năm trong Báo cáo tổng kết công tác khoa học công nghệ của Nhà trường.",
     "Ban Giám hiệu giao Phòng Khoa học Công nghệ là đơn vị đầu mối tổng hợp, theo dõi và báo cáo kết quả thực hiện các chỉ "
     "số hằng năm trong Báo cáo tổng kết công tác khoa học công nghệ của Nhà trường. Các chỉ số về văn bằng và nguồn thu từ "
     "khai thác tài sản trí tuệ dùng chung định nghĩa với Chuẩn cơ sở giáo dục đại học để có thể sử dụng trực tiếp trong báo "
     "cáo trên HEMIS."),
]

# Mục 3.1.2 mới ------------------------------------------------------------------
SWOT_MO = (
    "Trước khi xác định nguyên tắc và giải pháp, Đề tài tổng hợp các yếu tố bên trong rút ra từ kết quả đánh giá tại "
    "Chương 2 và các yếu tố bên ngoài từ khung pháp lý, chủ trương và yêu cầu quản lý sắp có hiệu lực thành ma trận điểm "
    "mạnh, điểm yếu, thời cơ và thách thức, thường gọi là ma trận SWOT. Điểm mạnh và điểm yếu phản ánh những gì Nhà trường "
    "đang có trong giai đoạn 2021 - 2025; thời cơ và thách thức phản ánh những thay đổi bên ngoài tác động đến Nhà trường "
    "trong giai đoạn triển khai giải pháp, trong đó có Thông tư số 83/2026/TT-BGDĐT ngày 30 tháng 9 năm 2026 quy định Chuẩn "
    "cơ sở giáo dục đại học, có hiệu lực từ ngày 15 tháng 11 năm 2026 và thay thế Thông tư số 01/2024/TT-BGDĐT. Thông tư "
    "này chưa áp dụng cho giai đoạn đánh giá tại Chương 2 nên chỉ được xem xét ở góc độ thời cơ và thách thức. Kết quả tổng "
    "hợp được trình bày tại Bảng 3.1.")

MANH = [
    "**1. Có quy định nội bộ từ sớm:** Quyết định 213 năm 2021 có chương riêng về sở hữu trí tuệ; Quyết định 217 năm 2024 "
    "là quy chế chuyên biệt, phạm vi tài sản rộng, đã quy định công bố, bảo mật, phân công đầu mối và hành vi xâm phạm "
    "quyền tác giả.",
    "**2. Năng lực nghiên cứu tăng nhanh:** số bài báo tăng từ 23 lên 161 bài giai đoạn 2021 - 2025; 71 trên 115 bài quốc tế "
    "có phân hạng Q; 6 trên 9 chỉ tiêu xác định được của Kế hoạch 07/KH-ĐHTĐ đạt hoặc vượt.",
    "**3. Đã có sản phẩm tiềm năng bảo hộ:** 11 trên 38 đề tài cấp cơ sở, trong đó 9 thuộc sở hữu công nghiệp và 8 có thể "
    "phù hợp với giải pháp hữu ích, tập trung ở lĩnh vực dược với 42 người trình độ tiến sĩ và tương đương tại Viện Y - Dược.",
    "**4. Kênh chuyển hóa và hợp tác doanh nghiệp đã có kết quả:** 1 đơn sáng chế từ đề tài năm 2025 và 1 đơn năm 2026; 4 "
    "văn bằng, giấy chứng nhận đã cấp; 5 kiểu dáng công nghiệp và 1 đơn nhãn hiệu từ chiến lược hợp tác doanh nghiệp mà Ban "
    "Giám hiệu đã dày công kết nối.",
    "**5. Quy chế và biểu mẫu đã có nền tảng:** Điều 35, Điều 38 Quyết định 213 có căn cứ chi lệ phí, thuê ngoài; biểu mẫu "
    "đề xuất, thuyết minh, hợp đồng, nghiệm thu đã có mục đăng ký sở hữu trí tuệ; 3 đề tài cấp quốc gia tổng 4,67 tỷ đồng; "
    "Quỹ Học bổng sau tiến sĩ Ngô Xuân Độ theo thông tin được cung cấp.",
]

YEU = [
    "**1. Chuyển hóa chưa tương xứng tiềm năng:** 1 trên 9 đề tài có sản phẩm tiềm năng sở hữu công nghiệp có đơn; 6 đề "
    "tài mã số 2021 - 2024 có sản phẩm tiềm năng sở hữu công nghiệp chưa có đơn.",
    "**2. Quy định chưa làm rõ phạm vi áp dụng sau thay đổi pháp luật:** chưa thống nhất thứ tự áp dụng các quy định về lợi "
    "ích của tác giả, chưa phân biệt thưởng, thù lao, phần chia nguồn thu; điểm a Điều 36 Quyết định 213 cần rà soát theo "
    "Điều 28, Điều 73 Luật số 93/2025/QH15; biểu mẫu chưa có nội dung sàng lọc khả năng bảo hộ.",
    "**3. Phối hợp chưa liên thông:** chức năng phân công cho 4 đơn vị nhưng chưa có luồng hồ sơ chung; chưa có vị trí "
    "chuyên trách sở hữu trí tuệ.",
    "**4. Kinh phí cho bước đăng ký chưa được dự toán riêng:** quy chế có căn cứ chi nhưng chưa thấy dự toán, cách tạm ứng "
    "và người đề xuất chi; theo Quy chế chi tiêu nội bộ năm 2026, văn bằng chưa có tiền thưởng và chỉ được ghi nhận khi được "
    "cấp.",
    "**5. Dữ liệu và khai thác yếu:** chưa có danh mục tài sản trí tuệ trong hệ thống thống kê, trạng thái 5 kiểu dáng chưa "
    "thống nhất, 7 trên 16 tiêu chí chưa tính được; khai thác có thu phí mới có 2 hợp đồng.",
]

THOI_CO = [
    "**1. Khung pháp luật mới trao quyền:** Luật số 93/2025/QH15 tự động giao quyền sở hữu kết quả cho tổ chức chủ trì, "
    "kể cả tổ chức ngoài công lập, và cho phép quỹ phát triển khoa học và công nghệ chi cho bảo hộ sở hữu trí tuệ; điểm c khoản 1 Điều 86 do Luật số 131/2025/QH15 bổ sung trao quyền "
    "đăng ký cho tổ chức chủ trì; Điều 28 Luật Giáo dục đại học số 125/2025/QH15 cho phép thành lập doanh nghiệp, định giá, "
    "góp vốn bằng tài sản trí tuệ.",
    "**2. Chủ trương chiến lược thuận lợi:** Kết luận số 51-KL/TW; Quyết định số 1624/QĐ-TTg dùng chỉ số sở hữu trí tuệ để "
    "đánh giá cơ sở giáo dục đại học, thí điểm hỗ trợ xác định giá trị ít nhất 100 quyền sở hữu trí tuệ và chuẩn bị Đề án "
    "của Bộ Giáo dục và Đào tạo giai đoạn 2027 - 2030.",
    "**3. Chuẩn cơ sở giáo dục đại học mới định giá văn bằng cao hơn:** bằng giải pháp hữu ích được tính 3 sản phẩm quy đổi "
    "thay cho 1, bằng sáng chế 5 sản phẩm; khoản thu từ thương mại hóa kết quả nghiên cứu, sở hữu trí tuệ được tách riêng.",
    "**4. Hạ tầng hỗ trợ mở rộng:** Mạng lưới Trung tâm Hỗ trợ công nghệ và đổi mới sáng tạo; cơ sở dữ liệu quốc gia về "
    "thực thi quyền sở hữu trí tuệ theo Chỉ thị số 02/CT-TTg; Nghị định số 134/2026/NĐ-CP làm rõ quyền tác giả khi có sử "
    "dụng trí tuệ nhân tạo.",
    "**5. Kênh hợp tác doanh nghiệp sẵn có:** đối tác đã cùng Nhà trường có 6 hồ sơ đồng sở hữu, có thể trở thành "
    "kênh thương mại hóa cho sản phẩm dược liệu.",
]
THACH_THUC = [
    "**1. Yêu cầu tuân thủ tăng từ ngày 15 tháng 11 năm 2026:** nội dung quản trị bắt buộc về sở hữu trí tuệ, liêm chính "
    "khoa học, liêm chính học thuật; dữ liệu kết quả hoạt động phải nhất quán trên HEMIS; kết quả tự đánh giá phải công bố "
    "trước ngày 31 tháng 5 hằng năm.",
    "**2. Sức ép công bố quốc tế:** ngưỡng công bố WoS, Scopus là 0,3 trên một giảng viên quy đổi; nếu tính trên danh sách "
    "giảng viên năm 2026, chỉ số năm 2025 của Nhà trường ước khoảng 0,31, sát ngưỡng, và văn bằng không được tính; nguy cơ công bố trước khi nộp đơn làm mất tính mới sau mười hai tháng.",
    "**3. Thủ tục dài, chi phí tự cân đối:** thẩm định hình thức, công bố đơn, thẩm định nội dung theo Điều 119 Luật Sở hữu "
    "trí tuệ kéo dài nhiều tháng, phát sinh chi phí "
    "tra cứu, soạn đơn, lệ phí và duy trì; là trường tư thục, Nhà trường phải tự cân đối từ nguồn thu của mình.",
    "**4. Pháp luật thay đổi nhanh, cần đối chiếu liên văn bản:** nhiều luật, nghị định, thông tư mới được ban hành "
    "trong giai đoạn 2025 - 2026, trong đó Nghị định số 267/2025/NĐ-CP đã quy định chi tiết việc giao quyền, thương mại "
    "hóa và phân chia lợi nhuận; quy chế nội bộ phải được đối chiếu đồng thời với các văn bản này. Bảng tổng hợp cuối "
    "Phụ lục II Thông tư số 83/2026/TT-BGDĐT chưa nêu giải pháp hữu ích dù công thức đã tính.",
    "**5. Yêu cầu tự chủ năng lực và rủi ro tranh chấp:** cần năng lực tự tra cứu, soạn đơn để phát triển tài sản đơn sở "
    "hữu; nguy cơ tranh chấp đối với sản phẩm đồng sáng tạo giữa giảng viên, người học, doanh nghiệp và sản phẩm có sử "
    "dụng trí tuệ nhân tạo.",
]
SWOT_NGUON = (
    "Nguồn: Nhóm nghiên cứu tổng hợp từ kết quả Chương 2 và các văn bản nêu tại Mục 3.1.1. Ước tính chỉ số công bố tạm lấy "
    "145 giảng viên theo danh sách năm 2026, quy đổi được 128,0 giảng viên theo Bảng 1 Phụ lục I Thông tư số "
    "83/2026/TT-BGDĐT, coi 40 bài báo có phân hạng Q năm 2025 là bài thuộc WoS hoặc Scopus, chưa áp dụng hệ số lĩnh vực.")
SWOT_PHAN_TICH = [
    "Bảng 3.1 cho thấy ba đặc điểm. Thứ nhất, điểm mạnh của Nhà trường nằm ở nền tảng và đầu vào: quy chế đã có, năng lực "
    "nghiên cứu tăng, sản phẩm có tiềm năng bảo hộ đã hình thành và quy chế đã có căn cứ chi; điểm yếu nằm ở "
    "khâu nối và khâu vận hành: quy trình, động lực, bộ máy và dữ liệu. Thứ hai, phần lớn thời cơ bên ngoài đòi hỏi đúng "
    "những năng lực mà Nhà trường đang yếu: quyền đăng ký, định giá, góp vốn và chỉ số văn bằng chỉ tạo ra giá trị khi có "
    "quy trình nhận diện và nộp đơn vận hành thường xuyên. Thứ ba, sức ép thời gian là thách thức rõ nhất: yêu cầu tuân "
    "thủ của Chuẩn mới có hiệu lực ngay trong năm 2026, còn sức ép công bố quốc tế có thể làm mất tính mới của chính các "
    "sản phẩm có tiềm năng bảo hộ.",
    "Kết hợp các yếu tố bên trong và bên ngoài, Đề tài xác định bốn nhóm phương án tại Bảng 3.2, làm căn cứ lựa chọn và sắp "
    "xếp thứ tự ưu tiên của năm nhóm giải pháp tại Mục 3.2.",
]
KET_HOP = [
    ["Phát huy điểm mạnh để tận dụng thời cơ",
     "Đưa các sản phẩm dược có tiềm năng và kết quả ba đề tài cấp quốc gia vào sàng lọc để đăng ký giải pháp hữu ích, sáng "
     "chế theo quyền tại điểm c khoản 1 Điều 86 Luật Sở hữu trí tuệ; cùng Viện Nghiên cứu giáo dục và Chuyển giao tri thức "
     "chuẩn bị phương án tư vấn, định giá để tham gia chương trình thí điểm xác định giá trị quyền sở hữu trí tuệ theo "
     "Quyết định số 1624/QĐ-TTg.", "Giải pháp 2, 3, 4"],
    ["Khắc phục điểm yếu nhờ thời cơ",
     "Dựa vào Điều 28, Điều 73 Luật số 93/2025/QH15 và Điều 135 Luật Sở hữu trí tuệ để làm rõ phạm vi áp dụng các quy định "
     "về lợi ích của tác giả, rà soát mức trần 100 triệu đồng cho các trường hợp thuộc phạm vi Luật mới; dùng chỉ số văn "
     "bằng của Chuẩn mới và Quyết định số 1624/QĐ-TTg để đưa văn bằng vào đánh giá, khen thưởng.", "Giải pháp 1, 4"],
    ["Phát huy điểm mạnh để vượt thách thức",
     "Biến cách làm của đơn sáng chế từ đề tài năm 2025 thành bước sàng lọc trước khi công bố và tại nghiệm thu, để vừa "
     "giữ ngưỡng công bố quốc tế vừa không mất tính mới; dùng Quyết định 217 làm nền cho nội dung quản trị bắt buộc về sở "
     "hữu trí tuệ và liêm chính.", "Giải pháp 1, 3"],
    ["Giảm điểm yếu và phòng tránh thách thức",
     "Lập danh mục số tài sản trí tuệ để quản lý nội bộ, có phân quyền truy cập, và trích dữ liệu được phép công bố cho "
     "báo cáo trên HEMIS, Nền tảng số quốc gia; lập dự toán phí xác lập quyền và giao đầu mối; đào tạo giảng viên về bộc lộ an toàn trước khi công bố.", "Giải pháp 2, 4, 5"],
]

KET_HOP_SAU = (
    "Trong bốn nhóm phương án, hai nhóm khắc phục điểm yếu nhờ thời cơ và phát huy điểm mạnh để vượt thách thức được ưu tiên "
    "triển khai trước, vì chúng xử lý trực tiếp điểm nghẽn ở khâu nối giữa nghiệm thu và đăng ký, đồng thời đáp ứng yêu cầu "
    "tuân thủ của Chuẩn mới trước kỳ tự đánh giá tháng 5 năm 2027. Hai nhóm còn lại được triển khai theo lộ trình tại Mục "
    "3.4.1.")

BO_SUNG_GP1 = (
    "- Về liêm chính khoa học, liêm chính học thuật: Hoàn thiện và mở rộng Điều 14 Quyết định 217 về hành vi xâm phạm "
    "quyền tác giả, bổ sung trách nhiệm kiểm tra trùng lặp, công khai việc sử dụng trí tuệ nhân tạo phù hợp với Điều 5a "
    "Nghị định số 134/2026/NĐ-CP và quy trình xử lý vi phạm, để Quy "
    "chế hợp nhất đáp ứng nội dung quản trị bắt buộc về sở hữu trí tuệ, liêm chính khoa học, liêm chính học thuật theo Thông "
    "tư số 83/2026/TT-BGDĐT.")

BANG_PHOI_HOP = dict(
    tieu_de="Cơ chế phối hợp liên phòng ban trong quản lý sở hữu trí tuệ",
    cot=["Đơn vị", "Vai trò", "Nhiệm vụ chính", "Sản phẩm đầu ra"],
    dong=[["Bộ phận Pháp chế, Trung tâm Dịch vụ và Quản trị hành chính tổng hợp", "Đầu mối thể chế và thủ tục",
           "Soạn thảo, cập nhật quy chế\nTiếp nhận, thẩm định hồ sơ\nLàm việc với Cục Sở hữu trí tuệ\nQuản lý kho lưu trữ tài "
           "sản trí tuệ\nĐăng ký quyền tác giả giáo trình số", "Quy chế hợp nhất\nVăn bằng bảo hộ\nSổ theo dõi đơn"],
          ["Phòng Khoa học Công nghệ", "Đầu mối chuyên môn và ngân sách",
           "Đánh giá tính mới\nPhân bổ kinh phí nghiên cứu\nQuản lý chỉ tiêu khoa học công nghệ\nTổ chức rà soát khả năng bảo "
           "hộ tại nghiệm thu\nCập nhật danh mục số tài sản trí tuệ",
           "Biên bản rà soát\nQuyết toán kinh phí\nBáo cáo chỉ tiêu và dữ liệu"],
          ["Viện Nghiên cứu giáo dục và Chuyển giao tri thức, các viện đào tạo", "Đầu mối tạo sinh tài sản và thương mại hóa",
           "Phát sinh sáng chế, giải pháp hữu ích\nVận hành Không gian sáng tạo mở thử nghiệm\nĐịnh giá, đàm phán chuyển giao",
           "Hồ sơ sáng chế\nHợp đồng chuyển giao\nBáo cáo định giá"],
          ["Phòng Tài chính - Kế toán", "Đơn vị phối hợp",
           "Bố trí kinh phí nộp đơn\nChi trả thù lao\nGiám sát sử dụng quỹ", "Chứng từ giải ngân\nPhụ lục hợp đồng"],
          ["Trung tâm Tuyển sinh và Quản trị thương hiệu", "Đơn vị phối hợp",
           "Quản lý nhãn hiệu, bộ nhận diện\nCấp phép sử dụng nhãn hiệu", "Văn bằng nhãn hiệu\nQuy chế thương hiệu"]],
    nguon="Nguồn: Nhóm nghiên cứu đề xuất trên cơ sở Điều 11 Quyết định 217 và thực trạng tại Mục 2.2.2.",
    rong=[3.6, 2.8, 5.2, 3.4], can=["left", "left", "left", "left"],
)

BANG_CHI_SO = dict(
    tieu_de="Bộ chỉ số theo dõi đánh giá hiệu quả triển khai hệ thống giải pháp, phân tầng theo chuỗi kết quả",
    cot=["Tầng", "Chỉ số theo dõi", "Nguồn dữ liệu", "Mục tiêu đến năm 2030"],
    dong=[["Thể chế và tổ chức",
           "1. Quy chế quản trị tài sản trí tuệ sửa đổi được ban hành\n2. Đầu mối sở hữu trí tuệ được giao và số cuộc giao "
           "ban liên phòng ban mỗi năm\n3. Tỷ lệ giảng viên mới hoàn thành tập huấn",
           "Quyết định ban hành; biên bản giao ban; danh sách tập huấn",
           "Trước ngày 31/5/2027\n04 cuộc mỗi năm\n100% giảng viên mới"],
          ["Sàng lọc",
           "4. Tỷ lệ đề tài nghiệm thu có phiếu rà soát\n5. Số kết quả được khai báo và sàng lọc mỗi năm\n6. Tỷ lệ kết quả "
           "thuộc nhánh sáng chế, giải pháp hữu ích, kiểu dáng được xem xét bảo mật trước khi công bố\n7. Thời gian từ "
           "khi có kết quả đánh giá tại Khâu 3 đến quyết định xác lập quyền, tách theo luồng sớm và luồng thường",
           "Phiếu khai báo, phiếu rà soát; danh mục số tài sản trí tuệ",
           "100% từ năm 2027\nTheo dõi, làm căn cứ quyết định nhân sự\n100%\nKhông quá 15 ngày làm việc"],
          ["Đơn nộp",
           "8. Số đơn sở hữu công nghiệp nộp mới từ kết quả nghiên cứu\n9. Số đăng ký quyền tác giả cho giáo trình, phần "
           "mềm có nhu cầu khai thác",
           "Danh mục số, nhóm trạng thái đã nộp đơn",
           "Từ 02 đơn năm 2027, 03 đến 05 đơn mỗi năm từ năm 2028, điều chỉnh sau thí điểm\nTheo nhu cầu đã rà soát"],
          ["Văn bằng",
           "10. Số văn bằng, giấy chứng nhận được cấp, tách theo nguồn: nghiên cứu, thương hiệu, hợp tác doanh nghiệp\n11. "
           "Tỷ lệ hồ sơ trong danh mục có trạng thái đã xác minh",
           "Danh mục số, nhóm trạng thái đã cấp; văn bằng gốc",
           "Từ 03 văn bằng mỗi năm theo mục 1.11 Kế hoạch 07/KH-ĐHTĐ, không gồm nhãn hiệu; văn bằng từ kết quả nghiên cứu "
           "theo dõi riêng\n100%"],
          ["Khai thác",
           "12. Số hợp đồng chuyển giao, cấp phép\n13. Nguồn thu từ khai thác và khoản chi trả cho tác giả, tách theo "
           "thưởng, thù lao, nhuận bút",
           "Hợp đồng; sổ kế toán",
           "Từ 01 hợp đồng mỗi năm theo mục 1.12 Kế hoạch 07/KH-ĐHTĐ\nTheo dõi hằng năm; công khai số liệu tổng hợp"]],
    nguon="Nguồn: Nhóm nghiên cứu đề xuất. Mục tiêu về đơn và văn bằng là mức tham khảo, cần được điều chỉnh sau thí điểm; "
          "văn bằng chỉ được cấp sau các giai đoạn thẩm định theo Điều 119 Luật Sở hữu trí tuệ nên được đánh giá chậm hơn "
          "các tầng sàng lọc và đơn nộp.",
    rong=[2.4, 6.4, 3.4, 4.0], can=["left", "left", "left", "left"],
)

BANG_TRIEN_KHAI = dict(
    tieu_de="Tổ chức thực hiện các giải pháp theo hạn chế được xử lý",
    cot=["Giải pháp và hạn chế được xử lý", "Chủ trì; phối hợp", "Nguồn lực", "Thời gian", "Sản phẩm đầu ra",
         "Cách đánh giá"],
    dong=[["**Giải pháp 1.** Hạn chế thứ hai: chưa rõ phạm vi, thứ tự áp dụng quy định về lợi ích của tác giả",
           "Bộ phận Pháp chế; Phòng Khoa học Công nghệ, Phòng Tài chính - Kế toán",
           "Ngân sách thường xuyên; chuyên gia rà soát dự thảo", "Quý IV/2026 đến 31/5/2027",
           "Quy chế sửa đổi; bảng phạm vi và thứ tự áp dụng; bảng mô phỏng tỷ lệ",
           "Ban hành đúng hạn; mỗi tình huống chuyển giao xác định được quy định áp dụng"],
          ["**Giải pháp 2.** Hạn chế thứ ba: phối hợp chưa liên thông, chưa có chuyên trách",
           "Ban Giám hiệu; Phòng Khoa học Công nghệ, Bộ phận Pháp chế",
           "Nhân sự hiện có kiêm nhiệm; kinh phí tập huấn", "Từ quý IV/2026; xem xét chuyên trách từ năm 2028",
           "Quyết định giao đầu mối; quy chế phối hợp; báo cáo khối lượng hồ sơ",
           "Số hồ sơ xử lý đúng hạn; quyết định nhân sự dựa trên ngưỡng"],
          ["**Giải pháp 3.** Hạn chế thứ nhất, thứ tư, thứ năm, thứ bảy: chuyển hóa thấp, thiếu sàng lọc, dữ liệu rời rạc",
           "Phòng Khoa học Công nghệ; Bộ phận Pháp chế, Hội đồng nghiệm thu, các viện",
           "Biểu mẫu hiện có được hoàn thiện; công cụ tra cứu; danh mục số",
           "Thí điểm quý IV/2026 - quý II/2027; toàn trường từ quý III/2027",
           "Phiếu khai báo, phiếu rà soát; quy trình 8 khâu; danh mục số",
           "Chỉ số 4 đến 8 và 11 tại Bảng 3.4"],
          ["**Giải pháp 4.** Nguyên nhân về nguồn lực, động lực; hạn chế thứ sáu: khai thác nhỏ",
           "Phòng Tài chính - Kế toán; Phòng Khoa học Công nghệ, Bộ phận Pháp chế, Viện Nghiên cứu giáo dục và Chuyển giao "
           "tri thức", "Quỹ nghiên cứu khoa học; nguồn thu dịch vụ khoa học công nghệ",
           "Dự toán từ năm 2027; phương án khai thác ở giai đoạn 2",
           "Dòng dự toán, quy định tạm ứng; đề xuất sửa Quy chế chi tiêu; phương án tổ chức khai thác",
           "Thời gian từ kết quả Khâu 3 đến nộp đơn, kể cả đơn nộp trước nghiệm thu; chỉ số 12, 13 tại Bảng 3.4; điều kiện "
           "chuyển bước"],
          ["**Giải pháp 5.** Hạn chế thứ nhất về kỹ năng nhận diện; giáo trình chưa đăng ký",
           "Phòng Khoa học Công nghệ; Phòng Đào tạo, Viện Nghiên cứu giáo dục và Chuyển giao tri thức",
           "Kinh phí đào tạo thường xuyên; chuyên gia bên ngoài", "Từ năm 2027; học phần từ năm học 2027 - 2028",
           "Chương trình tập huấn; cẩm nang; danh mục giáo trình cần đăng ký",
           "Chỉ số 3 và 9 tại Bảng 3.4; số giáo trình được rà soát"]],
    nguon="Nguồn: Nhóm nghiên cứu đề xuất trên cơ sở Mục 2.5.2 và Mục 2.5.3.",
    rong=[3.4, 2.8, 2.5, 2.0, 2.8, 2.6], can=["left", "left", "left", "left", "left", "left"],
)

MUC_TRIEN_KHAI = [
    "Năm nhóm giải pháp được thiết kế để mỗi nhóm xử lý một hoặc một số hạn chế đã xác định tại Mục 2.5.2 và các nguyên "
    "nhân tại Mục 2.5.3. Bảng 3.3 tổng hợp đơn vị chủ trì, đơn vị phối hợp, nguồn lực, thời gian, sản phẩm đầu ra và cách "
    "đánh giá của từng giải pháp; các chỉ số đánh giá được định nghĩa tại Bảng 3.4.",
]
MUC_TRIEN_KHAI_SAU = (
    "Bảng 3.3 cho thấy giai đoạn đầu chủ yếu dùng nguồn lực sẵn có: quy chế và biểu mẫu hiện hành được hoàn thiện thay vì "
    "thay thế, nhân sự hiện có được giao nhiệm vụ kiêm nhiệm, căn cứ chi tại Điều 35 và Điều 38 Quyết định 213 được chuyển "
    "thành dòng dự toán. Các đề xuất đòi hỏi nguồn lực lớn hơn, gồm vị trí chuyên trách, trung tâm tư vấn, định giá và "
    "doanh nghiệp quản lý tài sản trí tuệ, chỉ được đặt ra khi đạt điều kiện chuyển bước được đo bằng số liệu của chính "
    "các giải pháp giai đoạn đầu.")

TIEU_KET = [
    "Trên cơ sở kết quả đánh giá thực trạng tại Chương 2 và khung pháp lý mới, Chương 3 đã xây dựng hệ thống giải pháp nâng "
    "cao hiệu quả quản lý quyền sở hữu trí tuệ tại Trường Đại học Thành Đô. Phân tích điểm mạnh, điểm yếu, thời cơ và thách "
    "thức cho thấy Nhà trường có nền tảng về quy chế, biểu mẫu, năng lực nghiên cứu và sản phẩm có tiềm năng bảo hộ, nhưng "
    "yếu ở khâu nối giữa nghiệm thu và đăng ký. Thời cơ từ Luật số 93/2025/QH15, Luật số 131/2025/QH15, Luật Giáo dục đại "
    "học số 125/2025/QH15, Quyết định số 1624/QĐ-TTg và Chuẩn cơ sở giáo dục đại học mới chỉ được hiện thực hóa khi khâu này "
    "được khắc phục, trong khi yêu cầu tuân thủ của Thông tư số 83/2026/TT-BGDĐT và sức ép công bố quốc tế đặt ra giới hạn "
    "về thời gian.",
    "Từ đó, Đề tài đề xuất năm nguyên tắc và năm nhóm giải pháp: hoàn thiện quy chế với cơ chế lợi ích của tác giả xác định "
    "theo nguồn hình thành tài sản, thời điểm giao nhiệm vụ và loại đối tượng; giao đầu mối và cơ chế phối hợp, bố trí nhân "
    "sự theo giai đoạn; chuẩn hóa quy trình 8 khâu với bước khai báo, sàng lọc, xem xét bảo mật trước khi công bố và phân "
    "nhánh theo loại đối tượng; lập dòng dự toán cho bước xác lập quyền và chuẩn bị phương án tổ chức khai thác theo điều "
    "kiện chuyển bước; phát triển đào tạo và văn hóa sở hữu trí tuệ. Mỗi giải pháp gắn với hạn chế được xử lý, có chủ trì, "
    "phối hợp, nguồn lực, thời gian, sản phẩm đầu ra và cách đánh giá. Kế hoạch thí điểm tại Viện Y - Dược, lộ trình ba giai "
    "đoạn đến năm 2030 và bộ chỉ số phân tầng từ sàng lọc, đơn nộp, văn bằng đến khai thác bảo đảm hệ thống giải pháp có thể "
    "kiểm chứng, điều chỉnh và nhân rộng.",
]


def ap_sua(khoi, so_hinh_kh):
    for cu, moi in SUA:
        if moi is not None:
            moi = moi.replace("{SO_HINH_KH}", str(so_hinh_kh))
        vt = [i for i, k in enumerate(khoi) if k[0] == "p" and cu in k[1]]
        assert len(vt) == 1, f"{len(vt)} khối chứa: {cu[:80]}"
        i = vt[0]
        if moi is None:
            assert khoi[i][1] == cu
            del khoi[i]
            continue
        moi_t = khoi[i][1].replace(cu, moi)
        phan = moi_t.split("\n")
        khoi[i][1] = phan[0]
        for j, t in enumerate(phan[1:], 1):
            khoi.insert(i + j, ["p", t, False])
    # thay Mục 3.1.2 cũ bằng đánh dấu
    a = next(i for i, k in enumerate(khoi) if k[0] == "p" and k[1].startswith("3.1.2. "))
    b = next(i for i, k in enumerate(khoi) if k[0] == "p" and k[1].startswith("3.1.3. "))
    khoi[a:b] = [["swot", None, None]]
    # bổ sung nội dung liêm chính vào Giải pháp 1
    i = next(i for i, k in enumerate(khoi) if k[0] == "p" and k[1].startswith("- Về quan hệ pháp lý:"))
    khoi.insert(i + 1, ["p", BO_SUNG_GP1, False])
    return khoi


# Lượt chỉnh sửa theo bản góp ý ba chương (tháng 10 năm 2026). Mỗi mục: (đầu đoạn hiện có, nội dung mới hoặc danh sách).
SUA_LAN6 = [
    # 3.1.1
    ("Thứ ba, Luật Khoa học, Công nghệ và Đổi mới sáng tạo số 93/2025/QH15",
     "**Thứ ba,** Luật Khoa học, công nghệ và đổi mới sáng tạo số 93/2025/QH15 có hiệu lực thi hành từ ngày 01 tháng 10 năm "
     "2025 điều chỉnh hai cơ chế lợi ích của tác giả. Một là, tại điểm b và điểm h khoản 7 Điều 71, Luật sửa khoản 1 Điều "
     "135 Luật Sở hữu trí tuệ: chủ sở hữu trả thù lao cho tác giả sáng chế, kiểu dáng công nghiệp, thiết kế bố trí theo "
     "thỏa thuận; nếu không có thỏa thuận thì mức thù lao là 10% lợi nhuận trước thuế khi chủ sở hữu tự sử dụng, hoặc 15% "
     "tổng số tiền nhận được mỗi lần trước thuế khi chuyển giao quyền sử dụng; khoản 2 Điều 135 bị bãi bỏ. Hai là, Điều 28 "
     "quy định phân chia lợi nhuận từ thương mại hóa theo nguồn hình thành: chủ sở hữu tự quyết định đối với phần không sử "
     "dụng ngân sách nhà nước; đối với phần sử dụng ngân sách nhà nước, tổ chức chủ trì dùng lợi nhuận sau thuế để thưởng "
     "cho tác giả tối thiểu 30%, thưởng cho người tổ chức thương mại hóa, tái đầu tư và mục đích khác, và tác giả sáng chế, "
     "kiểu dáng công nghiệp, thiết kế bố trí còn hưởng quyền lợi theo Luật Sở hữu trí tuệ; khoản 3 và khoản 7 Điều 73 quy "
     "định chuyển tiếp theo thời điểm giao nhiệm vụ. Bên cạnh đó, khoản 2 Điều 25 quy định tổ chức chủ trì nhiệm vụ sử dụng "
     "ngân sách nhà nước được Nhà nước tự động giao quyền quản lý, sử dụng, quyền sở hữu phần kết quả tương ứng, không phải "
     "bồi hoàn chi phí; Điều 27 cho phép tổ chức được giao quyền tự quyết định hình thức, giá và phân chia lợi nhuận khi "
     "thương mại hóa; điểm b khoản 2 Điều 66 cho phép quỹ phát triển khoa học và công nghệ của tổ chức chi cho đăng ký, bảo "
     "hộ, quản lý, khai thác quyền sở hữu trí tuệ. Các quy định này là căn cứ để rà soát điểm a khoản 4 Điều 36 Quyết định "
     "số 213/QĐ-ĐHTĐ, gồm mức trần 100 triệu đồng, khoản nộp ngân sách nhà nước và cơ sở tính, đối với các trường hợp thuộc "
     "phạm vi áp dụng của Luật mới, như phân tích tại Mục 2.2.1."),
    ("Thứ tư, Luật sửa đổi, bổ sung một số điều của Luật Sở hữu trí tuệ số 131/2025/QH15",
     "**Thứ tư,** Luật sửa đổi, bổ sung một số điều của Luật Sở hữu trí tuệ số 131/2025/QH15 có hiệu lực thi hành từ ngày "
     "01 tháng 4 năm 2026, tại khoản 22 Điều 1 đã bổ sung điểm c khoản 1 Điều 86, theo đó tổ chức được giao quyền quản lý, "
     "sử dụng, quyền sở hữu kết quả của nhiệm vụ khoa học, công nghệ và đổi mới sáng tạo sử dụng ngân sách nhà nước có quyền "
     "đăng ký sáng chế, kiểu dáng công nghiệp, thiết kế bố trí là kết quả của nhiệm vụ đó. Quy định này tạo cơ sở pháp lý để "
     "Nhà trường, với tư cách tổ chức chủ trì ba đề tài cấp quốc gia do Quỹ Phát triển khoa học và công nghệ quốc gia tài "
     "trợ, chủ động đăng ký các đối tượng nêu trên là kết quả của các đề tài này; việc đăng ký cần được chuẩn bị trước khi "
     "kết quả được công bố."),
    ("Bên cạnh đó, tư cách thành viên Mạng lưới Trung tâm Hỗ trợ công nghệ",
     "Bên cạnh đó, theo thông tin được cung cấp, Nhà trường tham gia Mạng lưới Trung tâm Hỗ trợ công nghệ và đổi mới sáng "
     "tạo từ năm 2023; nếu được xác nhận bằng văn bản, tư cách này giúp Nhà trường tiếp cận cơ sở dữ liệu sáng chế phục vụ "
     "tra cứu và đánh giá tính mới của kết quả nghiên cứu."),
    # 3.1.3
    ("Dữ liệu thực trạng tại Chương 2 cho thấy điểm nghẽn nằm ở khâu nối",
     "Dữ liệu thực trạng tại Chương 2 cho thấy điểm nghẽn nằm ở khâu nối giữa nghiệm thu và đăng ký: 9 trên 38 đề tài có "
     "sản phẩm tiềm năng sở hữu công nghiệp nhưng chỉ 1 đề tài có đơn, trong khi năng lực công bố tập trung ở một nhóm nhỏ "
     "với hệ số Gini 0,832. Điểm nghẽn này liên quan đồng thời đến quy chế, quy trình, cơ chế tài chính, năng lực đầu mối và "
     "dữ liệu. Kết hợp với phân tích tại Mục 3.1.2, hệ thống giải pháp được xây dựng dựa trên năm nguyên tắc chỉ đạo:"),
    ("Nguyên tắc 3: Bảo đảm tính khả thi và kế thừa",
     "**Nguyên tắc 3:** Bảo đảm tính khả thi và kế thừa, gắn liền với điều kiện thực tế của Trường Đại học Thành Đô. Giải "
     "pháp không đề xuất thay đổi đột ngột bộ máy tổ chức hiện hành mà kế thừa sự phân công đã có tại Điều 11 Quyết định "
     "217: Phòng Khoa học Công nghệ nhận diện, lập hồ sơ theo dõi và xúc tiến thương mại hóa tài sản trí tuệ; Bộ phận Pháp "
     "chế thuộc Trung tâm Dịch vụ và Quản trị hành chính tổng hợp thực hiện thủ tục xác lập quyền. Quy chế, biểu mẫu hiện "
     "có được hoàn thiện thay vì thay thế; các đề xuất cần nhiều nguồn lực được triển khai theo giai đoạn, kèm điều kiện "
     "chuyển bước."),
    ("Nguyên tắc 4: Chủ động rà soát bộc lộ công khai",
     "**Nguyên tắc 4:** Chủ động sàng lọc và xem xét bảo mật từ khâu đề xuất, trước mọi hoạt động công bố, trình diễn và "
     "tại thời điểm nghiệm thu. Nguyên tắc này nhằm hạn chế việc bỏ sót kết quả có tiềm năng bảo hộ và ngăn ngừa nguy cơ "
     "bộc lộ công khai làm mất tính mới của sáng chế trước khi kịp nộp đơn đăng ký."),
    ("Nguyên tắc 5: Hài hòa lợi ích kinh tế và quyền nhân thân",
     "**Nguyên tắc 5:** Hài hòa lợi ích kinh tế và quyền nhân thân giữa Nhà trường, tập thể tác giả và các pháp nhân thành "
     "viên trong hệ sinh thái giáo dục. Cơ chế phân chia lợi ích được xác định theo nguồn hình thành tài sản, thời điểm giao "
     "nhiệm vụ và loại đối tượng, phân biệt thưởng, thù lao, nhuận bút và phần chia nguồn thu; phần do Nhà trường tự quyết "
     "phải đủ hấp dẫn để tạo động lực cho giảng viên, đồng thời bảo đảm nguồn tái đầu tư cho quỹ phát triển khoa học công "
     "nghệ của Nhà trường."),
    ("Trên cơ sở phân tích thực trạng tại Chương 2 và các căn cứ pháp lý nêu trên, Đề tài xây dựng hệ thống gồm 05 nhóm",
     "Trên cơ sở phân tích thực trạng tại Chương 2 và các căn cứ pháp lý nêu trên, Đề tài xây dựng hệ thống gồm 05 nhóm giải "
     "pháp, mỗi nhóm giải pháp được trình bày theo kết cấu 4 trụ cột: Mục tiêu - Nội dung thực hiện - Chủ thể thực hiện - "
     "Điều kiện bảo đảm. Các giải pháp được sắp xếp theo trình tự từ hoàn thiện thể chế nội bộ, kiện toàn tổ chức, chuẩn hóa "
     "quy trình, hoàn thiện cơ chế tài chính đến phát triển văn hóa và đào tạo. Mỗi nhóm giải pháp gắn với một hoặc nhiều "
     "hạn chế tại Mục 2.5.2 và một hoặc nhiều phương án kết hợp tại Bảng 3.2; chủ trì, phối hợp, nguồn lực, thời gian, sản "
     "phẩm đầu ra và cách đánh giá của từng giải pháp được tổng hợp tại Mục 3.2.6."),
    # Giải pháp 1
    ("Kết quả phân tích thực trạng tại Chương 2 cho thấy hiện nay Nhà trường đang vận hành đồng thời bốn văn bản",
     "Kết quả phân tích tại Chương 2 cho thấy các quy định nội bộ liên quan đến lợi ích của tác giả nằm ở bốn văn bản, gồm "
     "Quyết định số 213/QĐ-ĐHTĐ, Quyết định số 217/QĐ-ĐHTĐ ngày 21 tháng 11 năm 2024 ban hành Quy chế quản trị tài sản trí "
     "tuệ, Quy chế chi tiêu nội bộ và Điều lệ Quỹ Học bổng sau tiến sĩ Ngô Xuân Độ, được ban hành ở những thời điểm và cho "
     "những kênh tài trợ khác nhau. Các quy định này chưa xung đột trong cùng một tình huống, nhưng chưa làm rõ phạm vi và "
     "thứ tự áp dụng, chưa phân biệt thưởng, thù lao, nhuận bút và phần chia nguồn thu, và điểm a khoản 4 Điều 36 Quyết định "
     "213 cần được rà soát theo Điều 28, Điều 73 Luật số 93/2025/QH15 đối với các trường hợp thuộc phạm vi của Luật mới; đây "
     "là hạn chế thứ hai tại Mục 2.5.2. Quyết định 217 đã mở rộng phạm vi tài sản, quy định công bố, bảo mật, phân công "
     "trách nhiệm và hành vi xâm phạm quyền tác giả tại Điều 14, nhưng không dẫn chiếu Chương VI Quyết định 213; Thông tư "
     "số 83/2026/TT-BGDĐT xếp liêm chính khoa học, liêm chính học thuật cùng nhóm nội dung quản trị bắt buộc về sở hữu trí "
     "tuệ."),
    ("Giải pháp này nhằm khắc phục triệt để tình trạng nêu trên",
     "Giải pháp này nhằm hoàn thiện Quy chế quản trị tài sản trí tuệ trên nền Quyết định 217, hợp nhất các quy định về sở "
     "hữu trí tuệ tại Chương VI Quyết định 213, xác định rõ phạm vi và thứ tự áp dụng, bảo đảm phù hợp với pháp luật hiện "
     "hành và tạo nền tảng thể chế cho các giải pháp tiếp theo."),
    ("Bộ phận Pháp chế chủ trì tham mưu soạn thảo Quy chế",
     "Theo khoản 3 Điều 17 Quyết định 217, Bộ phận Pháp chế chủ trì tham mưu sửa đổi Quy chế khi pháp luật thay đổi. Quy "
     "chế sửa đổi, hợp nhất các quy định về sở hữu trí tuệ của Quyết định số 213/QĐ-ĐHTĐ và Quyết định số 217/QĐ-ĐHTĐ, cần "
     "bảo đảm các nội dung trọng tâm sau:"),
    ("- Về cơ chế thù lao và phân chia lợi ích:",
     "- **Về lợi ích của tác giả và phân chia nguồn thu:** Thiết kế theo ba lớp, xác định theo nguồn hình thành tài sản, "
     "thời điểm giao nhiệm vụ và loại đối tượng. Lớp thứ nhất là nghĩa vụ theo luật: đối với phần kết quả sử dụng ngân sách "
     "nhà nước thuộc phạm vi Điều 28 Luật số 93/2025/QH15, gồm nhiệm vụ giao từ ngày 01 tháng 10 năm 2025 và trường hợp tại "
     "khoản 7 Điều 73, thưởng cho tác giả không thấp hơn 30% lợi nhuận sau thuế hoặc 30% giá trị kết quả khi góp vốn, không "
     "áp mức trần làm phần thưởng thấp hơn mức này và không giữ khoản nộp ngân sách nhà nước như một tỷ lệ cố định; đối với "
     "nhiệm vụ phê duyệt trước ngày 01 tháng 10 năm 2025 không thuộc khoản 7 Điều 73, áp dụng văn bản có hiệu lực tại thời "
     "điểm phê duyệt; đối với sáng chế, kiểu dáng công nghiệp, thiết kế bố trí, quy định thù lao theo Điều 135 Luật Sở hữu "
     "trí tuệ như một khoản riêng và xác định rõ quy chế hoặc hợp đồng có phải là thỏa thuận về thù lao hay không. Lớp thứ "
     "hai là phần do Nhà trường tự quyết đối với tài sản không sử dụng ngân sách nhà nước theo khoản 2 Điều 28: tỷ lệ thưởng "
     "cho tác giả, phần cho đơn vị có tác giả và quỹ phát triển khoa học công nghệ, quan hệ với quy định trích 50% kinh phí "
     "chuyển giao công nghệ tại Quy chế chi tiêu nội bộ. Lớp thứ ba là nhuận bút cho sách, giáo trình theo Quy chế chi tiêu "
     "nội bộ, giữ nguyên khoản 2 Điều 13 Quyết định 217. Các phương án tỷ lệ cụ thể cho lớp thứ hai, như dành tỷ lệ cao hơn "
     "cho tác giả trong năm đầu phát sinh doanh thu, chỉ là phương án tham khảo và cần được mô phỏng trên một số tình huống "
     "chuyển giao trước khi đưa vào quy chế; Quy chế không áp một công thức chung cho mọi tài sản và mọi nguồn kinh phí."),
    # Giải pháp 2
    ("Thực trạng phân tích tại Chương 2 cho thấy chức năng quản lý sở hữu trí tuệ hiện đang được phân tán",
     "Thực trạng tại Chương 2 cho thấy công việc quản lý sở hữu trí tuệ được phân cho bốn đơn vị, Phòng Khoa học Công nghệ "
     "có 2 nhân sự kiêm nhiều mảng và chưa có vị trí chuyên trách, giữa các đơn vị chưa có luồng hồ sơ liên thông; đây là "
     "hạn chế thứ ba tại Mục 2.5.2. Điều 11 Quyết định 217 đã giao Phòng Khoa học Công nghệ quản lý chung và Bộ phận Pháp "
     "chế thực hiện thủ tục, nhưng mức độ thực hiện các nhiệm vụ này chưa được đánh giá. Giải pháp này nhằm bảo đảm các "
     "nhiệm vụ đã giao có người thực hiện, có quy chế phối hợp và có số liệu về khối lượng công việc để quyết định bước kiện "
     "toàn tiếp theo."),
    ("Giao Bộ phận Pháp chế (thuộc Trung tâm Dịch vụ và Quản trị hành chính tổng hợp) tiếp tục",
     "Bố trí nhân sự theo giai đoạn. Ở giai đoạn 1, giao một nhân sự hiện có của Phòng Khoa học Công nghệ và một nhân sự của "
     "Bộ phận Pháp chế làm đầu mối kiêm nhiệm, cử tham gia tập huấn của Cục Sở hữu trí tuệ về tra cứu sáng chế, soạn bản mô "
     "tả và định giá tài sản trí tuệ; thuê tổ chức đại diện sở hữu công nghiệp cho các hồ sơ sáng chế, giải pháp hữu ích. Ở "
     "giai đoạn 2, căn cứ số liệu của thí điểm và năm đầu áp dụng, xem xét bố trí vị trí chuyên trách khi khối lượng công "
     "việc đạt ngưỡng đề xuất, chẳng hạn từ 15 kết quả được sàng lọc hoặc từ 5 đơn mỗi năm; ngưỡng cụ thể do Ban Giám hiệu "
     "quyết định sau thí điểm."),
    ("Phối hợp với Viện Nghiên cứu giáo dục và Chuyển giao tri thức xây dựng Đề án thành lập Trung tâm",
     "Việc thành lập Trung tâm tư vấn, định giá và thương mại hóa tài sản trí tuệ theo định hướng tại điểm a khoản 6 Mục "
     "III Điều 1 Quyết định số 1068/QĐ-TTg đã được sửa đổi tại Quyết định số 1624/QĐ-TTg được đặt ở giai đoạn 2, với điều "
     "kiện chuyển bước: danh mục đã có tài sản sẵn sàng chuyển giao, có nhu cầu định giá thực tế và có nguồn kinh phí vận "
     "hành. Trước đó, các chức năng định giá, đàm phán được thực hiện qua hợp đồng với tổ chức tư vấn bên ngoài, với sự "
     "phối hợp của Viện Nghiên cứu giáo dục và Chuyển giao tri thức."),
    ("Bố trí 02 nhân sự chuyên trách có trình độ cử nhân luật",
     "Giai đoạn 1 không phát sinh biên chế mới, chỉ điều chỉnh nhiệm vụ của nhân sự hiện có; kinh phí tập huấn và thuê đại "
     "diện sở hữu công nghiệp lấy từ dòng dự toán tại Giải pháp 4. Cử cán bộ tham gia các lớp tập huấn chuyên sâu của Cục "
     "Sở hữu trí tuệ, tối thiểu 02 khóa mỗi năm. Duy trì giao ban liên phòng ban hằng quý về công tác sở hữu trí tuệ do Phó "
     "Hiệu trưởng phụ trách khoa học chủ trì, trong đó báo cáo số kết quả được khai báo, sàng lọc và nộp đơn làm căn cứ "
     "quyết định vị trí chuyên trách."),
    # Giải pháp 3
    ("Thực trạng tại Chương 2 cho thấy tiềm năng tài sản trí tuệ của đội ngũ là rất lớn",
     "Thực trạng tại Chương 2 cho thấy 9 đề tài có sản phẩm tiềm năng sở hữu công nghiệp nhưng mới 1 đề tài có đơn, ứng với "
     "hạn chế thứ nhất và thứ bảy tại Mục 2.5.2, và Đề tài nhận định khoảng trống kỹ thuật: thiếu một biểu mẫu rà soát tại "
     "thời điểm nghiệm thu. Các biểu mẫu tại Quyết định 213 đã có mục về đăng ký sở hữu trí tuệ nhưng chưa có nội dung sàng "
     "lọc khả năng bảo hộ và tình trạng bộc lộ; danh mục tài sản trí tuệ chưa liên kết với danh mục đề tài. Giải pháp này "
     "nhằm hoàn thiện biểu mẫu hiện có và chuẩn hóa quy trình để mỗi kết quả có tiềm năng được khai báo, sàng lọc và xem "
     "xét bảo mật trước khi công bố hoặc trình diễn, với cách xử lý phù hợp từng loại đối tượng."),
    ("Ban hành Quy trình chuẩn 8 khâu quản lý tài sản trí tuệ từ ý tưởng đến thương mại hóa",
     "Ban hành Quy trình chuẩn 8 khâu quản lý tài sản trí tuệ từ khai báo đến khai thác, áp dụng cho tất cả các đề tài cấp "
     "cơ sở và các nhiệm vụ khoa học công nghệ sử dụng ngân sách, được tóm tắt tại hình dưới đây và mô tả cụ thể sau đó."),
    ("Khâu 1:",
     "**Khâu 1:** Khai báo kết quả và nhu cầu bảo hộ. Hoàn thiện các biểu mẫu hiện có thay vì ban hành biểu mẫu mới: Mẫu "
     "01 đề xuất và Mẫu 06 thuyết minh bổ sung các trường loại đối tượng dự kiến theo năm nhánh, chủ thể quyền và đồng tác "
     "giả, nguồn kinh phí, lịch công bố dự kiến, tình trạng bảo mật và nhu cầu hỗ trợ; Mẫu 11 hợp đồng bổ sung nghĩa vụ khai "
     "báo kết quả trước khi công bố, phù hợp Điều 10 Quyết định 217; báo cáo tiến độ và Mẫu 15 báo cáo tổng kết cập nhật "
     "các trường này. Chủ nhiệm đề tài khai báo, Phòng Khoa học Công nghệ tiếp nhận, phù hợp định hướng xác định đối tượng "
     "quyền sở hữu trí tuệ cần đạt được tại điểm b khoản 4 Mục III Điều 1 Quyết định số 1068/QĐ-TTg đã được sửa đổi."),
    ("Khâu 2:",
     "**Khâu 2:** Sàng lọc sơ bộ và phân nhánh. Trong 10 ngày làm việc kể từ khi nhận phiếu khai báo, Phòng Khoa học Công "
     "nghệ phân loại kết quả theo năm nhánh: sáng chế, giải pháp hữu ích phải giữ bí mật đến khi nộp đơn; kiểu dáng công "
     "nghiệp cần nộp đơn trước khi trưng bày, giới thiệu sản phẩm; nhãn hiệu cần tra cứu trước khi sử dụng; quyền tác giả "
     "đối với giáo trình, phần mềm, sưu tập dữ liệu được ghi nhận và đăng ký khi cần chứng cứ để khai thác; bí mật kinh "
     "doanh như công thức, quy trình không công bố được bảo vệ bằng biện pháp bảo mật thay vì nộp đơn."),
    ("Khâu 3:",
     "**Khâu 3:** Tra cứu và đánh giá khả năng bảo hộ. Với các nhánh sáng chế, giải pháp hữu ích, kiểu dáng công nghiệp, "
     "nhãn hiệu, Bộ phận Pháp chế hoặc tổ chức đại diện sở hữu công nghiệp được thuê tra cứu trên các cơ sở dữ liệu sáng "
     "chế, nhãn hiệu, đánh giá sơ bộ tính mới, trình độ sáng tạo, khả năng áp dụng và đề xuất loại hình đăng ký."),
    ("Khâu 4:",
     "**Khâu 4:** Xem xét bảo mật trước khi công bố hoặc trình diễn. Trước khi gửi bài báo, báo cáo hội thảo, luận văn hoặc "
     "đưa sản phẩm vào thử nghiệm mở, trình diễn tại Không gian sáng tạo mở thử nghiệm, chủ nhiệm đề tài xin ý kiến Phòng "
     "Khoa học Công nghệ theo Điều 10 Quyết định 217; Phòng thông báo công bố, trì hoãn công bố hoặc yêu cầu nộp đơn trước "
     "theo Điều 11, hoặc yêu cầu người tham gia ký cam kết bảo mật. Khi kết quả đã bị bộc lộ, Phòng xác định ngày bộc lộ để "
     "tính thời hạn mười hai tháng theo khoản 3 Điều 60 Luật Sở hữu trí tuệ. Cách làm này phù hợp với định hướng đăng ký "
     "bảo hộ đồng thời với công bố tại điểm b khoản 4 Mục III Điều 1 Quyết định số 1068/QĐ-TTg đã được sửa đổi."),
    ("Khâu 5:",
     "**Khâu 5:** Xác nhận tại nghiệm thu. Mẫu 16 phiếu nghiệm thu được bổ sung phiếu rà soát gồm các trường: loại đối "
     "tượng, chủ thể quyền, nguồn kinh phí và thời điểm giao nhiệm vụ, tình trạng bộc lộ và ngày bộc lộ nếu có, kết quả "
     "sàng lọc tại Khâu 2 và Khâu 3, đề xuất hướng xử lý. Hội đồng nghiệm thu xác nhận phiếu rà soát; kết quả được đề xuất "
     "đăng ký được chuyển sang Khâu 6 trong thời hạn quy định."),
    ("Khâu 6:",
     "**Khâu 6:** Quyết định xác lập quyền và dự toán. Trên cơ sở phiếu rà soát, Phòng Khoa học Công nghệ trình Hiệu trưởng "
     "quyết định đăng ký theo Điều 35 Quyết định 213, kèm dự toán phí nộp đơn, phí đại diện, phí duy trì, nguồn chi, người "
     "đề xuất chi và thời hạn nộp đơn; Phòng Tài chính - Kế toán bố trí tạm ứng theo Giải pháp 4. Thời hạn từ khi có phiếu "
     "rà soát đến khi có quyết định không quá 15 ngày làm việc."),
    ("Khâu 7:",
     "**Khâu 7:** Nộp đơn, theo dõi và duy trì. Bộ phận Pháp chế nộp đơn tại Cục Sở hữu trí tuệ hoặc đăng ký quyền tác giả "
     "tại Cục Bản quyền tác giả, theo dõi các mốc thẩm định theo Điều 119 Luật Sở hữu trí tuệ, cập nhật trạng thái vào danh "
     "mục số theo bốn nhóm thống nhất: đã nộp đơn, đã chấp nhận đơn hợp lệ, đã cấp văn bằng hoặc giấy chứng nhận, chưa xác "
     "minh; lập lịch phí duy trì và cảnh báo trước 03 tháng. Với nhánh bí mật kinh doanh, Phòng Khoa học Công nghệ lập danh "
     "mục và áp dụng biện pháp bảo mật."),
    ("Khâu 8:",
     "**Khâu 8:** Khai thác và phân chia lợi ích. Phòng Khoa học Công nghệ chủ trì xúc tiến thương mại hóa theo Điều 11 "
     "Quyết định 217, phối hợp với Viện Nghiên cứu giáo dục và Chuyển giao tri thức và tổ chức tư vấn bên ngoài để định giá, "
     "tìm đối tác, đàm phán hợp đồng; lợi ích được phân chia theo lớp áp dụng tại Giải pháp 1, căn cứ nguồn hình thành tài "
     "sản, thời điểm giao nhiệm vụ và loại đối tượng."),
    ("Phòng Khoa học Công nghệ chủ trì ban hành Quy trình 8 khâu",
     "Phòng Khoa học Công nghệ chủ trì hoàn thiện biểu mẫu, ban hành quy trình và thực hiện các khâu 2, 4, 8. Chủ nhiệm đề "
     "tài thực hiện khâu 1. Bộ phận Pháp chế thực hiện các khâu 3 và 7. Hội đồng nghiệm thu thực hiện khâu 5. Hiệu trưởng "
     "và Phòng Tài chính - Kế toán thực hiện khâu 6. Viện Nghiên cứu giáo dục và Chuyển giao tri thức vận hành Không gian "
     "sáng tạo mở thử nghiệm theo nguyên tắc bảo mật tại khâu 4."),
    ("Hạ tầng Không gian sáng tạo mở thử nghiệm.",
     "Phiếu khai báo và phiếu rà soát được ban hành kèm quy trình; tài khoản truy cập công cụ tra cứu sáng chế; danh mục số "
     "tài sản trí tuệ dùng chung giữa Phòng Khoa học Công nghệ, Bộ phận Pháp chế và bộ phận quản trị thương hiệu; nhân sự "
     "đầu mối theo Giải pháp 2."),
    # Giải pháp 4
    ("3.2.4. Giải pháp 4:",
     "3.2.4. Giải pháp 4: Hoàn thiện cơ chế tài chính cho bước xác lập quyền và phương án tổ chức khai thác tài sản trí tuệ "
     "theo giai đoạn"),
    ("Một trong những nguyên nhân chủ quan quan trọng nhất khiến giảng viên chưa chủ động",
     "Chương 2 cho thấy quy chế đã có căn cứ chi cho bước xác lập quyền tại Điều 35, Điều 38 Quyết định 213 và Điều 13 "
     "Quyết định 217, nhưng chưa thấy dự toán riêng, cách tạm ứng, thanh toán khi đơn được nộp sau nghiệm thu, thời hạn và "
     "người chịu trách nhiệm đề xuất chi; số liệu chi thực tế chưa được thống kê. Quy chế chi tiêu nội bộ năm 2026 ghi nhận "
     "văn bằng khi được cấp và chưa có tiền thưởng cho văn bằng, trong khi bài báo quốc tế được thưởng ngay khi đăng; hoạt "
     "động khai thác có thu phí còn nhỏ, ứng với hạn chế thứ sáu tại Mục 2.5.2. Giải pháp này nhằm chuyển căn cứ chi hiện "
     "có thành một dòng dự toán vận hành được, điều chỉnh thời điểm ghi nhận văn bằng và chuẩn bị phương án tổ chức khai "
     "thác theo giai đoạn."),
    ("Nội dung 1: Trích lập Dòng kinh phí hỗ trợ đăng ký sở hữu trí tuệ",
     "**Nội dung 1:** Lập dòng dự toán hằng năm cho phí xác lập quyền trong Quỹ nghiên cứu khoa học của Trường theo Chương "
     "VII Quyết định 213, trên cơ sở Điều 35 và Điều 38 Quyết định 213, phù hợp với điểm b khoản 2 Điều 66 Luật số "
     "93/2025/QH15 và trách nhiệm thành lập quỹ phát triển khoa học và công nghệ tại điểm d khoản 3 Điều 28 Luật Giáo dục "
     "đại học số 125/2025/QH15. Dòng dự toán chi trả lệ phí nộp đơn, phí thẩm định, phí công bố, phí đại diện sở hữu công "
     "nghiệp cho các kết quả đã qua Khâu 5, kể cả khi đề tài đã quyết toán; quy định rõ người đề xuất chi là Phòng Khoa học "
     "Công nghệ, thời hạn đề xuất trong 15 ngày làm việc kể từ khi có phiếu rà soát, hình thức tạm ứng cho Bộ phận Pháp chế "
     "và hồ sơ thanh toán. Quy mô dự toán được tính từ số hồ sơ dự kiến nhân với mức phí theo quy định hiện hành về phí, lệ "
     "phí sở hữu công nghiệp, thay vì ấn định trước một tỷ lệ trích quỹ; tỷ lệ trích được để mở và điều chỉnh sau năm đầu. "
     "Sau khi đối chiếu văn bản Điều lệ, đề nghị bổ sung vào Điều lệ Quỹ Học bổng sau tiến sĩ Ngô Xuân Độ quy định cho phép "
     "hỗ trợ chi phí đăng ký đối với sản phẩm đề tài cấp cơ sở đã qua Khâu 5."),
    ("Nội dung 2: Sửa đổi Quy chế chi tiêu nội bộ",
     "**Nội dung 2:** Đề xuất sửa đổi Quy chế chi tiêu nội bộ năm 2026 theo hướng ghi nhận 50% định mức giờ nghiên cứu quy "
     "đổi khi Cục Sở hữu trí tuệ ra thông báo chấp nhận đơn hợp lệ và 50% còn lại khi văn bằng được cấp; xem xét mức thưởng "
     "bằng tiền cho bằng độc quyền sáng chế, giải pháp hữu ích, với mức cụ thể do Phòng Tài chính - Kế toán tính toán trên "
     "số hồ sơ dự kiến trước khi trình. Đây là thưởng theo quy chế nội bộ, tách biệt với thưởng khi thương mại hóa theo "
     "Điều 28 Luật số 93/2025/QH15 và thù lao theo Điều 135 Luật Sở hữu trí tuệ."),
    ("Nội dung 3: Xây dựng Đề án thành lập Doanh nghiệp quản lý tài sản trí tuệ",
     "**Nội dung 3:** Chuẩn bị phương án tổ chức khai thác theo giai đoạn. Ở giai đoạn 1 và 2, việc định giá, đàm phán, ký "
     "hợp đồng chuyển giao được thực hiện qua Phòng Khoa học Công nghệ, Bộ phận Pháp chế và hợp đồng thuê tổ chức tư vấn bên "
     "ngoài. Việc xây dựng Đề án thành lập doanh nghiệp quản lý tài sản trí tuệ theo khoản 1 Điều 28 Luật Giáo dục đại học "
     "số 125/2025/QH15 chỉ được đặt ra khi đạt các điều kiện chuyển bước: danh mục có tài sản đã được cấp văn bằng và có đối "
     "tác quan tâm, đã có ít nhất một hợp đồng chuyển giao hoặc cấp phép được thực hiện, và phương án tài chính cho thấy "
     "doanh nghiệp có thể tự trang trải chi phí vận hành. Khi đó, doanh nghiệp có chức năng định giá tài sản trí tuệ theo "
     "phương pháp chi phí, phương pháp thu nhập và phương pháp thị trường; góp vốn bằng tài sản trí tuệ; đàm phán và ký kết "
     "hợp đồng chuyển nhượng, cấp phép sử dụng các đối tượng sở hữu trí tuệ của Nhà trường."),
    ("Phòng Tài chính - Kế toán chủ trì xây dựng cơ chế trích lập kinh phí",
     "Phòng Tài chính - Kế toán chủ trì lập dòng dự toán, quy định tạm ứng và đề xuất sửa đổi Quy chế chi tiêu nội bộ. "
     "Phòng Khoa học Công nghệ đề xuất danh sách hồ sơ cần kinh phí. Viện Nghiên cứu giáo dục và Chuyển giao tri thức phối "
     "hợp xây dựng phương án tổ chức khai thác. Bộ phận Pháp chế rà soát tính pháp lý. Ban Giám hiệu trình Hội đồng trường "
     "phê duyệt các nội dung thuộc thẩm quyền."),
    ("Trích 5% đến 10% từ Quỹ Phát triển khoa học công nghệ hằng năm.",
     "Dòng dự toán phí xác lập quyền được lập từ năm 2027 căn cứ số hồ sơ dự kiến; nguồn bổ sung từ hợp đồng dịch vụ khoa "
     "học công nghệ, chuyển giao công nghệ và các nguồn tài trợ hợp pháp khác."),
    # Giải pháp 5
    ("Chương 2 không sử dụng khảo sát nhận thức, nhưng các dữ kiện hành vi",
     "Chương 2 không sử dụng khảo sát nhận thức, nhưng các dữ kiện hành vi cho thấy đội ngũ cần được hỗ trợ thêm về kỹ năng "
     "nhận diện và bảo hộ tài sản trí tuệ: 6 đề tài mã số 2021 - 2024 có sản phẩm tiềm năng sở hữu công nghiệp chưa có đơn; "
     "hai đề tài năm 2025 ghi dự kiến đăng ký giải pháp hữu ích nhưng chưa có đơn; 87 giáo trình chưa được đăng ký quyền tác "
     "giả, dù quyền tác giả đã phát sinh và Nhà trường là chủ sở hữu quyền tài sản theo khoản 1 Điều 39 Luật Sở hữu trí tuệ "
     "khi giao nhiệm vụ biên soạn và không có thỏa thuận khác, nên việc đăng ký có ý nghĩa tạo chứng cứ khi khai thác hoặc "
     "tranh chấp; Kế hoạch 07/KH-ĐHTĐ chỉ thống kê 12 lượt tập huấn chung trong giai đoạn 2019 - 2023 và không tách riêng "
     "tập huấn về sở hữu trí tuệ. Giải pháp này nhằm nâng kỹ năng nhận diện, bộc lộ an toàn và chủ động đón đầu định hướng "
     "nghiên cứu đưa sở hữu trí tuệ thành nội dung học bắt buộc tại điểm b khoản 8 Mục III Điều 1 Quyết định số "
     "1068/QĐ-TTg đã được sửa đổi tại Quyết định số 1624/QĐ-TTg."),
    ("Nội dung 4: Phát hành Cẩm nang sở hữu trí tuệ",
     ["**Nội dung 4:** Phát hành Cẩm nang sở hữu trí tuệ dành cho giảng viên và sinh viên. Xây dựng Chuyên trang Cơ sở dữ "
      "liệu số tài sản trí tuệ trên cổng thông tin điện tử của Nhà trường để tôn vinh tác giả, minh bạch hóa thông tin và "
      "sẵn sàng kết nối với Nền tảng số quốc gia theo điểm đ khoản 3 Điều 28 Luật Giáo dục đại học số 125/2025/QH15 và cung "
      "cấp dữ liệu cho báo cáo trên HEMIS.",
      "**Nội dung 5:** Rà soát 87 giáo trình, tài liệu giảng dạy giai đoạn 2021 - 2025 để xác định chủ thể quyền, đồng tác "
      "giả và tài liệu có nhu cầu khai thác ngoài Trường; chỉ đăng ký quyền tác giả cho tài liệu cần chứng cứ, thay vì đăng "
      "ký đại trà."]),
    # 3.3
    ("3.3. Phương án đề xuất thí điểm", "3.3. Kế hoạch thí điểm mô hình sàng lọc tại nghiệm thu"),
    ("3.3.1. Mục đích và phạm vi", "3.3.1. Mục đích và phạm vi thí điểm dự kiến"),
    ("Để bảo đảm tính khả thi và giảm thiểu rủi ro khi áp dụng đại trà",
     "Thí điểm dưới đây là kế hoạch, chưa được triển khai. Để bảo đảm tính khả thi và giảm rủi ro khi áp dụng toàn trường, "
     "Đề tài đề xuất thí điểm Khâu 1, Khâu 4 và Khâu 5 của quy trình, gồm phiếu khai báo, xem xét bảo mật trước khi công bố "
     "và phiếu rà soát tại nghiệm thu, trước khi trình ban hành chính thức. Phạm vi dự kiến là các đề tài cấp cơ sở năm 2026 "
     "của Viện Y - Dược, nơi phát sinh 9 trên 11 sản phẩm tiềm năng giai đoạn 2021 - 2025 và đơn sáng chế duy nhất trong kỳ; "
     "Viện Nghiên cứu giáo dục và Chuyển giao tri thức tham gia với vai trò hỗ trợ khai thác. Thời gian dự kiến từ quý IV năm "
     "2026 đến quý II năm 2027."),
    ("Mục đích của việc thí điểm bao gồm:",
     "Mục đích của việc thí điểm bao gồm: kiểm chứng tính khả thi và mức độ phù hợp của phiếu khai báo, phiếu rà soát; đo "
     "thời gian bổ sung cho mỗi phiên nghiệm thu; thu thập phản hồi của Hội đồng nghiệm thu và chủ nhiệm đề tài; phát hiện "
     "và điều chỉnh các điểm vướng mắc trước khi áp dụng toàn trường. Sản phẩm đầu ra là báo cáo đánh giá thí điểm, nêu số "
     "phiếu, số kết quả được sàng lọc, số đề xuất đăng ký, số đơn được nộp và thời gian xử lý."),
    ("3.3.2. Kịch bản đề xuất thí điểm", "3.3.2. Kịch bản thí điểm và các giả định cần kiểm chứng"),
    ("Trong kịch bản thí điểm, biểu mẫu đánh giá nghiệm thu được bổ sung",
     "Trong kịch bản thí điểm, Mẫu 16 phiếu nghiệm thu được bổ sung phiếu rà soát với các trường thông tin: loại đối tượng "
     "dự kiến theo năm nhánh; chủ thể quyền và đồng tác giả, kể cả người học và doanh nghiệp; nguồn kinh phí và thời điểm "
     "giao nhiệm vụ; lịch công bố, tình trạng bộc lộ và ngày bộc lộ nếu có; tình trạng bảo mật; kết quả tra cứu sơ bộ; nhu "
     "cầu hỗ trợ và đề xuất hướng xử lý, gồm nộp đơn, đăng ký quyền tác giả, giữ bí mật hoặc không xử lý."),
    ("Hai trường hợp thực tế minh họa cho kịch bản này",
     "Trường hợp đề tài chiết xuất lá Quế hoa mã số 09-2025, đề tài cấp cơ sở duy nhất trong kỳ có đơn sáng chế, nộp năm "
     "2025, cho thấy kết quả chiết tách có thể được đưa vào đơn khi chủ nhiệm đề tài chủ động. Đơn sáng chế số 1-2026-07185 "
     "về phương pháp chiết tách hợp chất từ cây Piper aduncum L. được nộp năm 2026, ngoài kỳ đánh giá. Hồ sơ hiện có chưa "
     "cho biết hai đơn được nộp trước hay sau khi kết quả được công bố, nên thí điểm cần ghi nhận cả ngày nộp đơn và ngày "
     "công bố để kiểm chứng. Bước sàng lọc nhằm biến cách làm của trường hợp đơn lẻ thành quy trình áp dụng cho mọi đề tài; "
     "đối với 6 đề tài mã số 2021 - 2024 có sản phẩm tiềm năng sở hữu công nghiệp chưa có đơn, Phòng Khoa học Công nghệ cần "
     "rà soát tình trạng bộc lộ để xác định sản phẩm nào còn khả năng đăng ký."),
    # 3.4.1
    ("Giai đoạn 1, từ quý IV năm 2026 đến quý II năm 2027:",
     "Giai đoạn 1, từ quý IV năm 2026 đến quý II năm 2027: Hoàn thiện thể chế và khởi động quy trình"),
    ("- Ban hành Quy chế quản lý sở hữu trí tuệ hợp nhất",
     "- Ban hành Quy chế quản trị tài sản trí tuệ sửa đổi trước ngày 31 tháng 5 năm 2027, trong đó xác định phạm vi và thứ "
     "tự áp dụng các quy định về lợi ích của tác giả theo Điều 28, Điều 73 Luật số 93/2025/QH15 và Điều 135 Luật Sở hữu trí "
     "tuệ, hoàn thiện nội dung liêm chính khoa học, liêm chính học thuật."),
    ("- Bổ sung 02 nhân sự chuyên trách sở hữu trí tuệ",
     "- Giao đầu mối kiêm nhiệm tại Phòng Khoa học Công nghệ và Bộ phận Pháp chế; cử cán bộ tham gia tập huấn tại Cục Sở hữu "
     "trí tuệ."),
    ("- Ban hành Quy trình chuẩn 8 khâu",
     "- Hoàn thiện biểu mẫu đề xuất, thuyết minh, hợp đồng, nghiệm thu với phiếu khai báo và phiếu rà soát; ban hành quy "
     "trình 8 khâu."),
    ("- Kích hoạt Dòng kinh phí hỗ trợ đăng ký", "- Lập dòng dự toán phí xác lập quyền cho năm 2027."),
    ("- Thực hiện thí điểm mô hình rà soát bộc lộ bắt buộc tại Viện Y - Dược.",
     "- Triển khai kế hoạch thí điểm tại Viện Y - Dược và đánh giá kết quả."),
    ("Giai đoạn 2, từ quý III năm 2027 đến hết năm 2028:",
     "Giai đoạn 2, từ quý III năm 2027 đến hết năm 2028: Áp dụng toàn trường và chuẩn bị khai thác"),
    ("Trên nền tảng thể chế đã được chuẩn hóa, giai đoạn này tập trung",
     "Trên nền tảng thể chế đã được hoàn thiện, giai đoạn này tập trung áp dụng quy trình toàn trường và chuẩn bị điều kiện "
     "khai thác:"),
    ("- Xây dựng và trình Đề án thành lập Doanh nghiệp quản lý tài sản trí tuệ",
     ["- Áp dụng quy trình 8 khâu cho toàn bộ đề tài cấp cơ sở và nhiệm vụ cấp quốc gia; xem xét bố trí vị trí chuyên trách "
      "khi khối lượng hồ sơ đạt ngưỡng.",
      "- Đánh giá điều kiện chuyển bước; xây dựng phương án trung tâm tư vấn, định giá hoặc doanh nghiệp quản lý tài sản trí "
      "tuệ khi đạt điều kiện tại Giải pháp 4."]),
    ("- Thực hiện thí điểm ít nhất 01 hợp đồng chuyển giao",
     "- Phấn đấu ký ít nhất 01 hợp đồng chuyển giao công nghệ hoặc cấp phép sử dụng mỗi năm, đáp ứng mục 1.12 Kế hoạch "
     "07/KH-ĐHTĐ."),
    ("- Vận hành chính thức Doanh nghiệp quản lý tài sản trí tuệ",
     "- Vận hành doanh nghiệp quản lý tài sản trí tuệ nếu đã đủ điều kiện thành lập; định giá, góp vốn khi có tài sản phù "
     "hợp."),
    ("- Khai thác thường xuyên tư cách thành viên Mạng lưới",
     "- Khai thác các mạng lưới hỗ trợ công nghệ và đổi mới sáng tạo mà Nhà trường tham gia."),
    ("- Thương mại hóa thường xuyên các kết quả nghiên cứu",
     "- Theo dõi nguồn thu từ khai thác tài sản trí tuệ, khoản chi trả cho tác giả và phần trích vào quỹ phát triển khoa "
     "học công nghệ."),
    # 3.4.2
    ("Để đo lường hiệu quả triển khai hệ thống giải pháp, Đề tài đề xuất Bộ chỉ số",
     "Để đo lường hiệu quả triển khai hệ thống giải pháp, Đề tài đề xuất Bộ chỉ số theo dõi đánh giá hiệu quả, được tổng hợp "
     "tại Bảng 3.4. Bộ chỉ số phân tầng theo chuỗi kết quả: sàng lọc, đơn nộp, văn bằng được cấp và khai thác, vì mỗi tầng "
     "có độ trễ khác nhau và văn bằng chỉ được cấp sau các giai đoạn thẩm định theo Điều 119 Luật Sở hữu trí tuệ. Các chỉ số "
     "được dùng làm căn cứ đánh giá hiệu quả hoạt động sở hữu trí tuệ của Nhà trường theo yêu cầu tại điểm b khoản 4 Mục III "
     "Điều 1 Quyết định số 1068/QĐ-TTg đã được sửa đổi tại Quyết định số 1624/QĐ-TTg. Chỉ tiêu văn bằng tại mục 1.11 Kế "
     "hoạch 07/KH-ĐHTĐ, 03 văn bằng mỗi năm giai đoạn 2026 - 2028, cần được thống kê đúng phạm vi của chỉ tiêu, không gồm "
     "nhãn hiệu, và tách riêng văn bằng hình thành từ kết quả nghiên cứu, tránh tình trạng chỉ tiêu gộp được hoàn thành mà "
     "không có kết quả nghiên cứu nào được bảo hộ như phân tích tại Hình 2.{SO_HINH_KH} Chương 2."),
    ("Bộ chỉ số nêu trên được thiết kế theo nguyên tắc đo lường được",
     "Bộ chỉ số nêu trên được thiết kế theo nguyên tắc đo lường được, có thời hạn cụ thể và gắn với trách nhiệm của từng "
     "đơn vị trong cơ chế phối hợp liên phòng ban. Ban Giám hiệu giao Phòng Khoa học Công nghệ là đơn vị đầu mối tổng hợp, "
     "theo dõi và báo cáo kết quả thực hiện các chỉ số hằng năm trong Báo cáo tổng kết công tác khoa học công nghệ của Nhà "
     "trường. Các chỉ số về văn bằng và nguồn thu từ khai thác tài sản trí tuệ dùng chung định nghĩa với Chuẩn cơ sở giáo "
     "dục đại học để có thể sử dụng trực tiếp trong báo cáo trên HEMIS. Các mục tiêu về đơn và văn bằng là mức tham khảo, "
     "được điều chỉnh sau khi có kết quả thí điểm."),
]


SO_DO = {
    "[[HINH_PHOI_HOP]]": ("Mô hình phối hợp liên phòng ban trong quản lý quyền sở hữu trí tuệ", "phoi_hop.png",
                          "Nguồn: Nhóm nghiên cứu đề xuất trên cơ sở Điều 11 Quyết định 217 và thực trạng tại Mục 2.2.2."),
    "[[HINH_TAM_KHAU]]": ("Quy trình 8 khâu quản lý tài sản trí tuệ từ khai báo đến khai thác, kèm luồng nộp đơn sớm và "
                          "phân nhánh theo loại đối tượng", "tam_khau.png",
                          "Nguồn: Nhóm nghiên cứu đề xuất."),
    "[[HINH_LO_TRINH]]": ("Lộ trình triển khai hệ thống giải pháp giai đoạn 2026 - 2030", "lo_trinh.png",
                          "Nguồn: Nhóm nghiên cứu đề xuất."),
}


def noi_dung(v, so_hinh_kh=10):
    import so_do
    tu = len(v.doc.paragraphs)
    thu_muc = os.path.join(GOC, "Ban_cuoi", "so_do")
    if not os.path.exists(os.path.join(thu_muc, "tam_khau.png")):
        so_do.ve_tat_ca()
    khoi = ap_sua(doc_khoi(), so_hinh_kh)
    so_bang_goc = 0
    for loai, t, dam in khoi:
        if loai == "swot":
            v.doan("h2", "3.1.2. Phân tích điểm mạnh, điểm yếu, thời cơ và thách thức")
            v.than_md(SWOT_MO)
            v.bang("Ma trận điểm mạnh, điểm yếu, thời cơ và thách thức trong quản lý quyền sở hữu trí tuệ tại Trường Đại "
                   "học Thành Đô", ["Điểm mạnh", "Điểm yếu"],
                   [["\n".join(MANH), "\n".join(YEU)], ["Thời cơ", "Thách thức"],
                    ["\n".join(THOI_CO), "\n".join(THACH_THUC)]],
                   SWOT_NGUON, [7.5, 7.5], can=["left", "left"], hang_dam=(1,), tien_to="3")
            v.than_md(*SWOT_PHAN_TICH)
            v.bang("Các phương án kết hợp từ ma trận và giải pháp tương ứng", ["Nhóm phương án", "Nội dung phương án",
                                                                                 "Giải pháp tương ứng"],
                   KET_HOP, "Nguồn: Nhóm nghiên cứu xây dựng từ Bảng 3.1.", [3.4, 9.2, 2.4],
                   can=["left", "left", "center"], tien_to="3")
            v.than_md(KET_HOP_SAU)
            continue
        if loai == "tbl":
            so_bang_goc += 1
            if so_bang_goc == 1:
                continue  # bảng phối hợp được thay bằng sơ đồ HINH_PHOI_HOP
            b = BANG_CHI_SO
            v.bang(b["tieu_de"], b["cot"], b["dong"], b["nguon"], b["rong"], can=b["can"], tien_to="3")
            continue
        if t in SO_DO:
            ten, anh, nguon = SO_DO[t]
            v.so_do(ten, os.path.join(thu_muc, anh), nguon, tien_to="3")
            continue
        if t == "CHƯƠNG 3":
            v.doan("chuong1", "CHƯƠNG 3")
            v.doan("chuong2", "HỆ THỐNG GIẢI PHÁP NÂNG CAO HIỆU QUẢ")
            v.doan("chuong3", "QUẢN LÝ QUYỀN SỞ HỮU TRÍ TUỆ TẠI TRƯỜNG ĐẠI HỌC THÀNH ĐÔ")
            v.doan("chuong3", "ĐÁP ỨNG KHUNG PHÁP LÝ MỚI")
            continue
        if t.startswith("HỆ THỐNG GIẢI PHÁP NÂNG CAO"):
            continue
        if t.startswith("3.3. "):
            v.doan("h2", "3.2.6. Tổ chức thực hiện các giải pháp")
            v.than_md(*MUC_TRIEN_KHAI)
            b = BANG_TRIEN_KHAI
            v.bang(b["tieu_de"], b["cot"], b["dong"], b["nguon"], b["rong"], can=b["can"], tien_to="3")
            v.than_md(MUC_TRIEN_KHAI_SAU)
        if re.match(r"^3\.\d\. ", t):
            v.doan("h1", t)
        elif re.match(r"^3\.\d\.\d\. ", t):
            v.doan("h2", t)
        elif re.match(r"^[a-d]\) ", t) and dam:
            v.doan("h3", t, bold=True)
        elif dam:
            v.doan_md("than", f"**{t}**")
        else:
            m = TIEN_TO_DAM.match(t)
            if m:
                t = (m.group(1) or "") + f"**{m.group(2).strip()}** " + t[m.end():]
            v.doan_md("than", t)
    v.doan("h1", "TIỂU KẾT CHƯƠNG 3")
    v.than_md(*TIEU_KET)
    sua_lan6(v.doc, tu, so_hinh_kh)
    sua_lan7(v.doc, tu)
    assert v.so_bang == 4, v.so_bang


def sua_lan6(doc, tu, so_hinh_kh):
    """Lượt chỉnh sửa theo bản góp ý ba chương: thống nhất pháp lý với Chương 1, 2; giải pháp theo giai đoạn, có điều
    kiện chuyển bước; quy trình có khai báo, sàng lọc, bảo mật trước công bố; thí điểm ghi là kế hoạch."""
    from khung import thay_doan
    for bat_dau, moi in SUA_LAN6:
        if isinstance(moi, str):
            moi = moi.replace("{SO_HINH_KH}", str(so_hinh_kh))
        thay_doan(doc, bat_dau, moi, tu=tu)


# Lượt chỉnh sửa lần 7 theo góp ý về quy trình nộp đơn sớm, mức độ kết luận, Nghị định số 267/2025/NĐ-CP và dữ liệu
# công khai. Mỗi mục: (đầu đoạn, [(cụm cũ, cụm mới)], [đoạn mới chèn sau]).
SUA_LAN7 = [
    # 3.1.1: bổ sung Nghị định số 267/2025/NĐ-CP
    ("Thứ ba, Luật Khoa học, công nghệ và đổi mới sáng tạo số 93/2025/QH15 có hiệu lực",
     [("như phân tích tại Mục 2.2.1.",
       "như phân tích tại Mục 2.2.1. Nghị định số 267/2025/NĐ-CP, có hiệu lực từ ngày 14 tháng 10 năm 2025, quy định chi "
       "tiết các nội dung này: khoản 2 Điều 32 giao tự động quyền sở hữu phần kết quả tương ứng với kinh phí ngân sách nhà "
       "nước cho tổ chức chủ trì không phải là đơn vị công lập, như Nhà trường, và yêu cầu theo dõi riêng thông tin về kết "
       "quả; Điều 34 yêu cầu phân chia lợi nhuận công khai, minh bạch, quy định tổ chức trung gian, môi giới hưởng tối "
       "thiểu 10% lợi nhuận khi các bên không có thỏa thuận khác, và phần thưởng giữa các đồng tác giả được chia theo thỏa "
       "thuận của họ.")], []),
    # 3.1.2: mức độ kết luận
    ("Kết quả phân tích tại Chương 2 cho thấy các quy định nội bộ liên quan đến lợi ích",
     [("Các quy định này chưa xung đột trong cùng một tình huống, nhưng chưa làm rõ",
       "Qua đối chiếu theo từng tình huống, Đề tài chưa xác định được trường hợp các quy định này áp dụng không tương "
       "thích cho cùng một tài sản; tuy vậy, các quy định chưa làm rõ"),
      ("cần được rà soát theo Điều 28, Điều 73 Luật số 93/2025/QH15 đối với",
       "cần được rà soát theo Điều 28, Điều 73 Luật số 93/2025/QH15 và Điều 34 Nghị định số 267/2025/NĐ-CP đối với")],
     []),
    # 3.2: nguyên tắc
    ("Dữ liệu thực trạng tại Chương 2 cho thấy điểm nghẽn nằm ở khâu nối",
     [("Dữ liệu thực trạng tại Chương 2 cho thấy điểm nghẽn nằm ở khâu nối giữa nghiệm thu và đăng ký",
       "Dữ liệu thực trạng tại Chương 2 gợi ý điểm nghẽn nằm ở khâu nối giữa kết quả nghiên cứu và đăng ký"),
      ("trong khi năng lực công bố tập trung ở một nhóm nhỏ với hệ số Gini 0,832. Điểm nghẽn này liên quan đồng thời",
       "trong khi phân tích thăm dò trên danh sách nhân sự năm 2026 cho thấy hoạt động công bố tập trung ở một nhóm "
       "tương đối nhỏ. Điểm nghẽn này có thể liên quan đồng thời"),
      ("năng lực đầu mối và dữ liệu.",
       "năng lực đầu mối và dữ liệu; mức độ tác động của từng yếu tố cần được kiểm chứng trong quá trình triển khai.")],
     []),
    # 3.2.3: mục tiêu
    ("Thực trạng tại Chương 2 cho thấy 9 đề tài có sản phẩm tiềm năng",
     [("và Đề tài nhận định khoảng trống kỹ thuật: thiếu một biểu mẫu rà soát tại thời điểm nghiệm thu.",
       "và Đề tài nêu giả thuyết rằng việc thiếu bước sàng lọc trước khi công bố và tại nghiệm thu có thể là một yếu tố "
       "cản trở."),
      ("với cách xử lý phù hợp từng loại đối tượng.",
       "với cách xử lý phù hợp từng loại đối tượng. Kết quả cần được bảo vệ trước khi công bố phải được quyết định đăng "
       "ký và cấp kinh phí ngay trong quá trình nghiên cứu, không chờ đến nghiệm thu.")], []),
    ("Ban hành Quy trình chuẩn 8 khâu quản lý tài sản trí tuệ",
     [("được tóm tắt tại hình dưới đây và mô tả cụ thể sau đó.",
       "được tóm tắt tại hình dưới đây và mô tả cụ thể sau đó. Quy trình có hai luồng. Luồng sớm áp dụng cho kết quả đã "
       "được sàng lọc và đánh giá tại Khâu 2, Khâu 3 và cần nộp đơn trước khi công bố hoặc trình diễn: hồ sơ được chuyển "
       "thẳng sang Khâu 6 để quyết định đăng ký, cấp kinh phí và nộp đơn, kể cả khi đề tài đang thực hiện. Luồng thường "
       "áp dụng cho các kết quả còn lại và đi qua Khâu 5. Nghiệm thu là bước kiểm tra, cập nhật tình trạng quyền, không "
       "phải điều kiện để quyết định đăng ký. Cách làm này phù hợp Điều 35 Quyết định 213, vốn không đặt điều kiện đã "
       "nghiệm thu đối với việc tác giả nộp đơn tại Phòng Khoa học Công nghệ.")], []),
    ("Khâu 4:",
     [("Phòng thông báo công bố, trì hoãn công bố hoặc yêu cầu nộp đơn trước theo Điều 11, hoặc yêu cầu người tham gia ký "
       "cam kết bảo mật.",
       "Phòng thông báo công bố, trì hoãn công bố hoặc yêu cầu nộp đơn trước theo Điều 11, hoặc yêu cầu người tham gia ký "
       "cam kết bảo mật. Khi yêu cầu nộp đơn trước, Phòng chuyển hồ sơ đã có kết quả đánh giá tại Khâu 3 sang Khâu 6 theo "
       "luồng sớm, và việc công bố chỉ được thực hiện sau khi đã có ngày nộp đơn.")], []),
    ("Khâu 5:",
     [("Xác nhận tại nghiệm thu.", "Kiểm tra tại nghiệm thu."),
      ("kết quả sàng lọc tại Khâu 2 và Khâu 3, đề xuất hướng xử lý. Hội đồng nghiệm thu xác nhận phiếu rà soát; kết quả "
       "được đề xuất đăng ký được chuyển sang Khâu 6 trong thời hạn quy định.",
       "kết quả sàng lọc tại Khâu 2 và Khâu 3, tình trạng quyền của các kết quả đã đi theo luồng sớm như đã nộp đơn, đã có "
       "số đơn, đã chấp nhận đơn hợp lệ, và đề xuất hướng xử lý. Nghiệm thu không phải điều kiện để quyết định đăng ký hay "
       "cấp kinh phí: Hội đồng nghiệm thu kiểm tra việc khai báo đã đầy đủ chưa, phát hiện kết quả chưa được khai báo, "
       "xác nhận tình trạng quyền và mức độ đóng góp của các thành viên. Đối với nhiệm vụ sử dụng ngân sách nhà nước, hồ sơ "
       "đánh giá cuối kỳ có văn bản xác định mức độ đóng góp của thành viên, có xác nhận của các thành viên, để làm căn cứ "
       "phân chia lợi nhuận theo điểm g khoản 2 Điều 17 Nghị định số 267/2025/NĐ-CP; đề tài cấp cơ sở áp dụng cùng cách "
       "làm. Kết quả chưa được xử lý trong quá trình nghiên cứu và được đề xuất đăng ký được chuyển sang Khâu 6 theo luồng "
       "thường.")], []),
    ("Khâu 6:",
     [("Quyết định xác lập quyền và dự toán. Trên cơ sở phiếu rà soát, Phòng Khoa học Công nghệ",
       "Quyết định xác lập quyền và cấp kinh phí. Khâu này nhận hồ sơ từ hai luồng: luồng sớm, gồm kết quả đã được đánh giá "
       "khả năng bảo hộ tại Khâu 3 và cần nộp đơn trước khi công bố, trình diễn, kể cả khi đề tài đang thực hiện; luồng "
       "thường, gồm kết quả được đề xuất tại Khâu 5. Phòng Khoa học Công nghệ"),
      ("Thời hạn từ khi có phiếu rà soát đến khi có quyết định không quá 15 ngày làm việc.",
       "Thời hạn từ khi có kết quả đánh giá tại Khâu 3 đối với luồng sớm, hoặc từ khi có phiếu rà soát đối với luồng "
       "thường, đến khi có quyết định không quá 15 ngày làm việc; với luồng sớm, quyết định phải được ban hành và đơn phải "
       "được nộp trước ngày công bố dự kiến.")], []),
    ("Phiếu khai báo và phiếu rà soát được ban hành kèm quy trình",
     [("danh mục số tài sản trí tuệ dùng chung giữa Phòng Khoa học Công nghệ, Bộ phận Pháp chế và bộ phận quản trị thương "
       "hiệu;",
       "danh mục số tài sản trí tuệ dùng chung giữa Phòng Khoa học Công nghệ, Bộ phận Pháp chế và bộ phận quản trị thương "
       "hiệu, có phân quyền truy cập: hồ sơ chưa nộp đơn, bí mật kinh doanh và thông tin bị ràng buộc bởi hợp đồng chỉ "
       "người được giao xử lý mới truy cập được; danh mục theo dõi riêng kết quả hình thành từ ngân sách nhà nước theo "
       "khoản 2 Điều 32 Nghị định số 267/2025/NĐ-CP;")], []),
    # 3.2.1
    ("- Về lợi ích của tác giả và phân chia nguồn thu:",
     [("Quy chế không áp một công thức chung cho mọi tài sản và mọi nguồn kinh phí.",
       "Quy chế không áp một công thức chung cho mọi tài sản và mọi nguồn kinh phí. Quy chế cần kèm mẫu thỏa thuận phân "
       "chia phần thưởng giữa các đồng tác giả, vì theo khoản 5 Điều 28 Luật số 93/2025/QH15 và khoản 3 Điều 34 Nghị định "
       "số 267/2025/NĐ-CP, phần thưởng là mức dành chung cho các đồng tác giả và được chia theo thỏa thuận giữa họ; mẫu "
       "này dùng chung với văn bản xác định mức độ đóng góp lập tại Khâu 5 của Giải pháp 3.")], []),
    # 3.2.4
    ("Chương 2 cho thấy quy chế đã có căn cứ chi cho bước xác lập quyền",
     [("thanh toán khi đơn được nộp sau nghiệm thu,", "thanh toán khi đơn được nộp trước hoặc sau nghiệm thu,")], []),
    ("Nội dung 1: Lập dòng dự toán hằng năm cho phí xác lập quyền",
     [("cho các kết quả đã qua Khâu 5, kể cả khi đề tài đã quyết toán;",
       "cho các kết quả đã được đánh giá khả năng bảo hộ tại Khâu 3 và có quyết định tại Khâu 6, trong thời gian thực "
       "hiện đề tài hoặc sau nghiệm thu, kể cả khi đề tài đã quyết toán; khoản chi cho đơn nộp trước nghiệm thu được chi từ "
       "dòng dự toán này, không trừ vào kinh phí của đề tài;"),
      ("thời hạn đề xuất trong 15 ngày làm việc kể từ khi có phiếu rà soát,",
       "thời hạn đề xuất trong 15 ngày làm việc kể từ khi có kết quả đánh giá tại Khâu 3 đối với luồng sớm, hoặc từ khi "
       "có phiếu rà soát đối với luồng thường,"),
      ("đối với sản phẩm đề tài cấp cơ sở đã qua Khâu 5.",
       "đối với sản phẩm đề tài cấp cơ sở đã được đánh giá tại Khâu 3, không phân biệt đã nghiệm thu hay chưa.")], []),
    ("Nội dung 3: Chuẩn bị phương án tổ chức khai thác theo giai đoạn.",
     [("và hợp đồng thuê tổ chức tư vấn bên ngoài.",
       "và hợp đồng thuê tổ chức tư vấn bên ngoài; hợp đồng với tổ chức trung gian, môi giới cần thỏa thuận rõ mức hưởng, "
       "vì khi không có thỏa thuận, khoản 2 Điều 34 Nghị định số 267/2025/NĐ-CP dành cho tổ chức này tối thiểu 10% lợi "
       "nhuận thu được.")], []),
    # 3.2.5
    ("Nội dung 4: Phát hành Cẩm nang sở hữu trí tuệ",
     [("Xây dựng Chuyên trang Cơ sở dữ liệu số tài sản trí tuệ trên cổng thông tin điện tử của Nhà trường để tôn vinh tác "
       "giả, minh bạch hóa thông tin và sẵn sàng kết nối với Nền tảng số quốc gia theo điểm đ khoản 3 Điều 28 Luật Giáo dục "
       "đại học số 125/2025/QH15 và cung cấp dữ liệu cho báo cáo trên HEMIS.",
       "Xây dựng Chuyên trang giới thiệu tài sản trí tuệ trên cổng thông tin điện tử của Nhà trường để tôn vinh tác giả và "
       "giới thiệu tài sản sẵn sàng chuyển giao; chuyên trang chỉ đăng thông tin đã được phép công bố, tách khỏi danh mục "
       "số dùng cho quản lý nội bộ tại Giải pháp 3. Dữ liệu được phép công bố được trích từ danh mục số để kết nối với Nền "
       "tảng số quốc gia theo điểm đ khoản 3 Điều 28 Luật Giáo dục đại học số 125/2025/QH15 và cung cấp cho báo cáo trên "
       "HEMIS.")], []),
    # 3.3 thí điểm
    ("Thí điểm dưới đây là kế hoạch, chưa được triển khai.",
     [("Đề tài đề xuất thí điểm Khâu 1, Khâu 4 và Khâu 5 của quy trình, gồm phiếu khai báo, xem xét bảo mật trước khi công "
       "bố và phiếu rà soát tại nghiệm thu,",
       "Đề tài đề xuất thí điểm Khâu 1, Khâu 4, Khâu 5 và luồng sớm từ Khâu 3 sang Khâu 6 của quy trình, gồm phiếu khai "
       "báo, xem xét bảo mật trước khi công bố, quyết định đăng ký và cấp kinh phí trước nghiệm thu khi cần, và phiếu rà "
       "soát tại nghiệm thu,")], []),
    ("Mục đích của việc thí điểm bao gồm:",
     [("số đề xuất đăng ký, số đơn được nộp và thời gian xử lý.",
       "số đề xuất đăng ký, số đơn được nộp theo từng luồng và thời gian xử lý của mỗi luồng.")], []),
    # 3.4.1 lộ trình
    ("- Hoàn thiện Chuyên trang Cơ sở dữ liệu số tài sản trí tuệ",
     [("- Hoàn thiện Chuyên trang Cơ sở dữ liệu số tài sản trí tuệ, sẵn sàng kết nối với Nền tảng số quốc gia.",
       "- Hoàn thiện danh mục số tài sản trí tuệ có phân quyền truy cập; chuẩn bị trích xuất dữ liệu được phép công bố để "
       "kết nối với Nền tảng số quốc gia và báo cáo trên HEMIS.")], []),
    ("- Số hóa toàn bộ dữ liệu tài sản trí tuệ",
     [("- Số hóa toàn bộ dữ liệu tài sản trí tuệ và công khai minh bạch hằng năm trên Nền tảng số quốc gia.",
       "- Số hóa đầy đủ dữ liệu tài sản trí tuệ để quản lý nội bộ; hằng năm chỉ công khai dữ liệu được phép công bố, gồm "
       "văn bằng đã cấp, đơn đã được công bố chính thức và số liệu tổng hợp; hồ sơ chưa nộp đơn, bí mật kinh doanh và "
       "thông tin bị ràng buộc bởi hợp đồng được giới hạn quyền truy cập.")], []),
    ("- Ban hành Quy chế quản trị tài sản trí tuệ sửa đổi trước ngày 31 tháng 5 năm 2027",
     [("theo Điều 28, Điều 73 Luật số 93/2025/QH15 và Điều 135 Luật Sở hữu trí tuệ,",
       "theo Điều 28, Điều 73 Luật số 93/2025/QH15, Điều 32, Điều 34 Nghị định số 267/2025/NĐ-CP và Điều 135 Luật Sở hữu "
       "trí tuệ,")], []),
]


def sua_lan7(doc, tu):
    """Lượt chỉnh sửa lần 7: luồng nộp đơn sớm, mức độ kết luận, Nghị định số 267/2025/NĐ-CP, dữ liệu công khai."""
    from khung import sua_cum
    for bat_dau, cap, them in SUA_LAN7:
        sua_cum(doc, bat_dau, cap, tu=tu, them_sau=them)


def dung():
    v = VanBanChung()
    noi_dung(v)
    return luu(v, "Chuong_3_He_thong_giai_phap.docx",
               "Chương 3. Hệ thống giải pháp nâng cao hiệu quả quản lý quyền sở hữu trí tuệ tại Trường Đại học Thành Đô", 3)


if __name__ == "__main__":
    dung()

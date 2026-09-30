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
    "là quy chế chuyên biệt, phạm vi tài sản rộng và đã phân công đầu mối.",
    "**2. Năng lực nghiên cứu tăng nhanh:** số bài báo tăng từ 23 lên 161 bài giai đoạn 2021 - 2025; 71 trên 115 bài quốc tế "
    "có phân hạng Q; 7 trên 10 chỉ tiêu Kế hoạch 07/KH-ĐHTĐ đạt hoặc vượt.",
    "**3. Đã có sản phẩm đủ điều kiện bảo hộ:** 11 trên 38 đề tài cấp cơ sở, trong đó 8 sản phẩm phù hợp với giải pháp hữu "
    "ích, tập trung ở lĩnh vực dược với 42 người trình độ tiến sĩ và tương đương tại Viện Y - Dược.",
    "**4. Kênh chuyển hóa và hợp tác doanh nghiệp đã vận hành:** 2 đơn sáng chế năm 2025 và 2026; 9 văn bằng, trong đó "
    "5 kiểu dáng công nghiệp và nhãn hiệu từ chiến lược hợp tác doanh nghiệp mà Ban Giám hiệu đã dày công kết nối.",
    "**5. Có nguồn lực gắn cơ chế sở hữu trí tuệ:** Quỹ Học bổng sau tiến sĩ Ngô Xuân Độ ngân sách 5 tỷ đồng; 3 đề tài cấp "
    "quốc gia tổng 4,67 tỷ đồng; Viện Nghiên cứu giáo dục và Chuyển giao tri thức có tư cách tổ chức khoa học và công nghệ, "
    "đã ký hợp đồng chuyển giao quyền sử dụng tác phẩm.",
]
YEU = [
    "**1. Chuyển hóa chưa tương xứng tiềm năng:** 1 trên 11 đề tài đủ điều kiện được nộp đơn; 8 đề tài đủ điều kiện giai "
    "đoạn 2021 - 2024 chưa được nộp đơn.",
    "**2. Quy định chưa được hợp nhất sau thay đổi pháp luật:** 4 văn bản với 5 quy định chia lợi ích; trần 100 triệu "
    "đồng chưa tương thích với Điều 28 Luật số 93/2025/QH15; chưa có biểu mẫu rà soát khả năng bảo hộ khi nghiệm thu.",
    "**3. Phối hợp chưa liên thông:** chức năng phân công cho 4 đơn vị nhưng chưa có luồng hồ sơ chung; chưa có vị trí "
    "chuyên trách sở hữu trí tuệ.",
    "**4. Động lực và kinh phí lệch về công bố:** văn bằng không có tiền thưởng và chỉ được ghi nhận khi được cấp; đề tài cấp "
    "cơ sở không có dòng chi phí nộp đơn, kinh phí trung vị 8,5 triệu đồng.",
    "**5. Dữ liệu và khai thác yếu:** không có danh mục tài sản trí tuệ trong hệ thống thống kê, độ phủ khai báo khoảng "
    "29,9%, 7 trên 16 tiêu chí chưa tính được; khai thác có thu phí mới có 2 hợp đồng.",
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
    "**5. Kênh hợp tác doanh nghiệp sẵn có:** đối tác đã cùng Nhà trường xác lập 6 tài sản đồng sở hữu, có thể trở thành "
    "kênh thương mại hóa cho sản phẩm dược liệu.",
]
THACH_THUC = [
    "**1. Yêu cầu tuân thủ tăng từ ngày 15 tháng 11 năm 2026:** nội dung quản trị bắt buộc về sở hữu trí tuệ, liêm chính "
    "khoa học, liêm chính học thuật; dữ liệu kết quả hoạt động phải nhất quán trên HEMIS; kết quả tự đánh giá phải công bố "
    "trước ngày 31 tháng 5 hằng năm.",
    "**2. Sức ép công bố quốc tế:** ngưỡng công bố WoS, Scopus là 0,3 trên một giảng viên quy đổi, ước tính năm 2025 của Nhà "
    "trường khoảng 0,31 và văn bằng không được tính; nguy cơ công bố trước khi nộp đơn làm mất tính mới sau mười hai tháng.",
    "**3. Thủ tục dài, chi phí tự cân đối:** thẩm định sở hữu công nghiệp thường kéo dài hai đến ba năm, phát sinh chi phí "
    "tra cứu, soạn đơn, lệ phí và duy trì; là trường tư thục, Nhà trường phải tự cân đối từ nguồn thu của mình.",
    "**4. Pháp luật thay đổi nhanh, hướng dẫn chưa đồng bộ:** sáu văn bản lớn trong giai đoạn 2025 - 2026; bảng tổng hợp "
    "cuối Phụ lục II Thông tư số 83/2026/TT-BGDĐT chưa nêu giải pháp hữu ích dù công thức đã tính.",
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
    "nghiên cứu tăng, sản phẩm có khả năng bảo hộ đã hình thành và nguồn lực tài chính đã được thiết lập; điểm yếu nằm ở "
    "khâu nối và khâu vận hành: quy trình, động lực, bộ máy và dữ liệu. Thứ hai, phần lớn thời cơ bên ngoài đòi hỏi đúng "
    "những năng lực mà Nhà trường đang yếu: quyền đăng ký, định giá, góp vốn và chỉ số văn bằng chỉ tạo ra giá trị khi có "
    "quy trình nhận diện và nộp đơn vận hành thường xuyên. Thứ ba, thách thức lớn nhất không nằm ở thiếu nguồn lực mà ở sức "
    "ép thời gian: yêu cầu tuân thủ của Chuẩn mới có hiệu lực ngay trong năm 2026, còn sức ép công bố quốc tế có thể làm "
    "mất tính mới của chính các sản phẩm đủ điều kiện bảo hộ.",
    "Kết hợp các yếu tố bên trong và bên ngoài, Đề tài xác định bốn nhóm phương án tại Bảng 3.2, làm căn cứ lựa chọn và sắp "
    "xếp thứ tự ưu tiên của năm nhóm giải pháp tại Mục 3.2.",
]
KET_HOP = [
    ["Phát huy điểm mạnh để tận dụng thời cơ",
     "Đưa các sản phẩm dược đủ điều kiện và kết quả ba đề tài cấp quốc gia vào đăng ký giải pháp hữu ích, sáng chế theo "
     "quyền tại điểm c khoản 1 Điều 86 Luật Sở hữu trí tuệ; cùng Viện Nghiên cứu giáo dục và Chuyển giao tri thức lập trung "
     "tâm tư vấn, định giá để tham gia chương trình thí điểm xác định giá trị quyền sở hữu trí tuệ theo Quyết định số "
     "1624/QĐ-TTg.", "Giải pháp 2, 3, 4"],
    ["Khắc phục điểm yếu nhờ thời cơ",
     "Dựa vào Điều 28 Luật số 93/2025/QH15 và nguyên tắc thỏa thuận thù lao để hợp nhất quy chế, cập nhật mức trần 100 triệu đồng; "
     "dùng chỉ số văn bằng của Chuẩn mới và Quyết định số 1624/QĐ-TTg để đưa văn bằng vào đánh giá, khen thưởng ngang với "
     "công bố.", "Giải pháp 1, 4"],
    ["Phát huy điểm mạnh để vượt thách thức",
     "Biến cách làm của hai đơn sáng chế năm 2025 và 2026 thành khâu rà soát bắt buộc trước khi công bố, để vừa giữ ngưỡng "
     "công bố quốc tế vừa không mất tính mới; dùng Quyết định 217 làm nền cho nội dung quản trị bắt buộc về sở hữu trí tuệ "
     "và liêm chính.", "Giải pháp 1, 3"],
    ["Giảm điểm yếu và phòng tránh thách thức",
     "Lập danh mục số tài sản trí tuệ liên thông với báo cáo trên HEMIS và Nền tảng số quốc gia; bố trí dòng kinh phí nộp "
     "đơn và nhân sự chuyên trách; đào tạo giảng viên về bộc lộ an toàn trước khi công bố.", "Giải pháp 2, 4, 5"],
]
KET_HOP_SAU = (
    "Trong bốn nhóm phương án, hai nhóm khắc phục điểm yếu nhờ thời cơ và phát huy điểm mạnh để vượt thách thức được ưu tiên "
    "triển khai trước, vì chúng xử lý trực tiếp điểm nghẽn ở khâu nối giữa nghiệm thu và đăng ký, đồng thời đáp ứng yêu cầu "
    "tuân thủ của Chuẩn mới trước kỳ tự đánh giá tháng 5 năm 2027. Hai nhóm còn lại được triển khai theo lộ trình tại Mục "
    "3.4.1.")

BO_SUNG_GP1 = (
    "- Về liêm chính khoa học, liêm chính học thuật: Bổ sung quy định về trách nhiệm tôn trọng quyền tác giả, kiểm tra trùng "
    "lặp, công khai việc sử dụng trí tuệ nhân tạo phù hợp với Điều 5a Nghị định số 134/2026/NĐ-CP và xử lý vi phạm, để Quy "
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
    tieu_de="Bộ chỉ số theo dõi đánh giá hiệu quả triển khai hệ thống giải pháp",
    cot=["Nhóm giải pháp", "Chỉ số theo dõi", "Mục tiêu đến năm 2030"],
    dong=[["Giải pháp 1: Thể chế nội bộ",
           "1.1. Tỷ lệ văn bản nội bộ được hợp nhất và đồng bộ hóa\n1.2. Mức thù lao trả cho tác giả sáng chế so với mức "
           "mặc định của pháp luật",
           "100% văn bản thống nhất\nKhông thấp hơn mức mặc định tại khoản 1 Điều 135 Luật Sở hữu trí tuệ"],
          ["Giải pháp 2: Tổ chức bộ máy",
           "2.1. Số nhân sự chuyên trách sở hữu trí tuệ\n2.2. Số cuộc họp giao ban liên phòng ban mỗi năm\n2.3. Số lượt tra "
           "cứu sáng chế tiền kiểm mỗi năm\n2.4. Tỷ lệ tài sản trí tuệ được cập nhật trong danh mục số",
           "02 nhân sự\n04 cuộc mỗi năm\nTăng 20% mỗi năm\n100%"],
          ["Giải pháp 3: Quy trình 8 khâu",
           "3.1. Tỷ lệ đề tài được rà soát khả năng bảo hộ tại nghiệm thu\n3.2. Số đơn đăng ký bảo hộ nộp mới mỗi năm\n3.3. "
           "Số văn bằng hình thành từ kết quả nghiên cứu được cấp",
           "100% đề tài\n05 đến 10 đơn mỗi năm\n03 đến 05 văn bằng mỗi năm"],
          ["Giải pháp 4: Cơ chế tài chính",
           "4.1. Tỷ lệ đề tài đủ điều kiện được tài trợ kinh phí nộp đơn\n4.2. Thời gian từ khi đơn được chấp nhận hợp lệ đến "
           "khi ghi nhận 50% giờ quy đổi\n4.3. Số hợp đồng chuyển giao, cấp phép được ký kết",
           "100%\nDưới 30 ngày\n01 đến 02 hợp đồng mỗi năm"],
          ["Giải pháp 5: Đào tạo và văn hóa",
           "5.1. Tỷ lệ sinh viên học học phần sở hữu trí tuệ\n5.2. Tỷ lệ giảng viên mới hoàn thành tập huấn\n5.3. Số lượt truy "
           "cập chuyên trang dữ liệu số tài sản trí tuệ",
           "100% sinh viên từ năm 2028\n100% giảng viên mới\nTăng 30% mỗi năm"]],
    nguon="Nguồn: Nhóm nghiên cứu đề xuất.",
    rong=[3.4, 7.0, 4.6], can=["left", "left", "left"],
)

TIEU_KET = [
    "Trên cơ sở kết quả đánh giá thực trạng tại Chương 2 và khung pháp lý mới, Chương 3 đã xây dựng hệ thống giải pháp nâng "
    "cao hiệu quả quản lý quyền sở hữu trí tuệ tại Trường Đại học Thành Đô. Phân tích điểm mạnh, điểm yếu, thời cơ và thách "
    "thức cho thấy Nhà trường có nền tảng về quy chế, năng lực nghiên cứu, sản phẩm có khả năng bảo hộ và nguồn lực tài "
    "chính, nhưng yếu ở khâu nối giữa nghiệm thu và đăng ký. Thời cơ từ Luật số 93/2025/QH15, Luật số 131/2025/QH15, Luật "
    "Giáo dục đại học số 125/2025/QH15, Quyết định số 1624/QĐ-TTg và Chuẩn cơ sở giáo dục đại học mới chỉ được hiện thực hóa "
    "khi khâu này được khắc phục, trong khi yêu cầu tuân thủ của Thông tư số 83/2026/TT-BGDĐT và sức ép công bố quốc tế đặt "
    "ra giới hạn về thời gian.",
    "Từ đó, Đề tài đề xuất năm nguyên tắc và năm nhóm giải pháp: hoàn thiện quy chế nội bộ hợp nhất; kiện toàn bộ máy và cơ "
    "chế phối hợp liên phòng ban; chuẩn hóa quy trình 8 khâu với khâu rà soát bắt buộc tại nghiệm thu; xây dựng cơ chế tài "
    "chính linh hoạt, cơ chế chuyển tiếp với Quỹ Học bổng sau tiến sĩ Ngô Xuân Độ và doanh nghiệp quản lý tài sản trí tuệ; "
    "phát triển đào tạo và văn hóa sở hữu trí tuệ. Mỗi giải pháp được trình bày đủ bốn trụ cột: mục tiêu, nội dung thực "
    "hiện, chủ thể thực hiện và điều kiện bảo đảm. Phương án thí điểm tại Viện Y - Dược, lộ trình ba giai đoạn đến năm 2030 "
    "và bộ chỉ số theo dõi bảo đảm hệ thống giải pháp có thể kiểm chứng, điều chỉnh và nhân rộng.",
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


SO_DO = {
    "[[HINH_PHOI_HOP]]": ("Mô hình phối hợp liên phòng ban trong quản lý quyền sở hữu trí tuệ", "phoi_hop.png",
                          "Nguồn: Nhóm nghiên cứu đề xuất trên cơ sở Điều 11 Quyết định 217 và thực trạng tại Mục 2.2.2."),
    "[[HINH_TAM_KHAU]]": ("Quy trình 8 khâu quản lý tài sản trí tuệ từ ý tưởng đến thương mại hóa", "tam_khau.png",
                          "Nguồn: Nhóm nghiên cứu đề xuất."),
    "[[HINH_LO_TRINH]]": ("Lộ trình triển khai hệ thống giải pháp giai đoạn 2026 - 2030", "lo_trinh.png",
                          "Nguồn: Nhóm nghiên cứu đề xuất."),
}


def noi_dung(v, so_hinh_kh=10):
    import so_do
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
    assert v.so_bang == 3, v.so_bang


def dung():
    v = VanBanChung()
    noi_dung(v)
    return luu(v, "Chuong_3_He_thong_giai_phap.docx",
               "Chương 3. Hệ thống giải pháp nâng cao hiệu quả quản lý quyền sở hữu trí tuệ tại Trường Đại học Thành Đô", 3)


if __name__ == "__main__":
    dung()

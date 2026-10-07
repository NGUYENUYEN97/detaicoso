# -*- coding: utf-8 -*-
"""Sửa Chương 2 bản cuối Báo cáo toàn văn theo góp ý, xuất tệp có theo dõi thay đổi.

Đầu vào : bản đã sửa Chương 1 (Bao_cao_toan_van_ban_cuoi_sua_Chuong1.docx)
Đầu ra  : Bao_cao_toan_van_ban_cuoi_sua_Chuong1_2.docx

Nội dung:
1. Văn bản: rút gọn 2.1.1, 2.1.2, 2.2.3; bớt các câu cảnh báo lặp lại; viết 2.1.3 theo phát hiện;
   kết luận rõ hơn ở 2.2.1; thêm đoạn đánh giá ở 2.2.4; chia nguyên nhân khách quan, chủ quan ở 2.3.3.
2. Độ chính xác (đối chiếu nguồn gốc):
   - Bảng 2.2 trả về số liệu đã đối chiếu với các danh mục thống kê (scripts/chuong2/du_lieu.py);
   - Quyết định 213 ngày 28/12/2021, Quyết định 217 là quy chế riêng; Điều 11, 13 Quyết định 217 về đầu mối;
   - giờ quy đổi, mức thưởng theo Phụ lục 5 và Bảng 7 Quy chế chi tiêu nội bộ năm 2026;
   - trạng thái hồ sơ theo sổ theo dõi đơn nhãn hiệu, kiểu dáng công nghiệp, sáng chế của Trường;
   - Điều 3, Điều 10, Điều 14 - 16 Quyết định 217; Điều 35, 36, 38 Quyết định 213;
   - Điều 94, 110, 119 Luật Sở hữu trí tuệ; điểm a khoản 3 Điều 28 Luật số 93/2025/QH15; Kết luận 51-KL/TW.
3. Hình: thay hình vẽ bằng ký tự bằng biểu đồ Excel nhúng (sửa dữ liệu được) và sơ đồ hình vẽ Word;
   bỏ Hình 2.2 (mô phỏng) và Hình 2.6 (bộ 16 tiêu chí không có nguồn); đánh số lại hình, bảng theo thứ tự xuất hiện.
"""
import copy
import os
import re
import sys
import zipfile

from lxml import etree

DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, DIR)
sys.path.insert(0, os.path.join(DIR, "..", "chuong2"))
from track_changes import TrackEditor, q  # noqa: E402
import so_do_word as SD  # noqa: E402
import bieu_do as B  # noqa: E402
from build import chart_xml, xlsx_bytes, cao_cm, INLINE, CT_CHART, CT_XLSX, RT_PACKAGE, EMU_CM  # noqa: E402

GOC = os.path.abspath(os.path.join(DIR, "..", ".."))
THU_MUC = os.path.join(GOC, "Ban_cuoi", "Ban_hoan_thien_03-10-2026")
VAO = os.path.join(THU_MUC, "Bao_cao_toan_van_ban_cuoi_sua_Chuong1.docx")
RA = os.path.join(THU_MUC, "Bao_cao_toan_van_ban_cuoi_sua_Chuong1_2.docx")
RT_CHART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"


# ---------------------------------------------------------------------------
# Biểu đồ Excel nhúng (dùng lại cấu hình và số liệu đã đối chiếu ở scripts/chuong2)
# ---------------------------------------------------------------------------
def _h(i, so, **sua):
    h = copy.deepcopy(B.HINH[i])
    h["so"] = so
    h.update(sua)
    return h


BIEU_DO = {
    "Hình 2.1. Cơ cấu và diễn biến sản phẩm khoa học": _h(1, 2),  # dữ liệu sau khi loại trùng, xem dưới
    "Hình 2.4. Số đề tài cấp cơ sở theo hình thức kinh phí": _h(9, 3),
    "Hình 2.5. Giờ nghiên cứu quy đổi và mức thưởng": _h(10, 4),
    "Hình 2.8. Sản phẩm đề tài cấp cơ sở có tiềm năng": _h(13, 5),
    "Hình 2.7. Hồ sơ tài sản trí tuệ của Nhà trường": _h(
        11, 6,
        cot=["Loại hình", "Đã cấp văn bằng, giấy chứng nhận", "Chờ cấp bằng", "Đã nộp đơn"],
        nguon="Nguồn: Sổ theo dõi đơn nhãn hiệu, kiểu dáng công nghiệp, sáng chế và bảng thống kê văn bằng của "
              "Trường Đại học Thành Đô."),
}
assert [d[1:] for d in BIEU_DO["Hình 2.7. Hồ sơ tài sản trí tuệ của Nhà trường"]["dong"]] == \
    [[2, 0, 0], [2, 0, 1], [0, 5, 0], [0, 0, 1]]
# Nhãn hiệu Double2n: sổ theo dõi ghi Chờ cấp bằng; đơn sáng chế 1-2025-07378: đã nộp năm 2025
# (dòng "Chấp nhận đơn hợp lệ" trong sổ không ghi số đơn nên không gắn với đơn này)
BIEU_DO["Hình 2.7. Hồ sơ tài sản trí tuệ của Nhà trường"]["dong"] = [
    ["Quyền tác giả", 2, 0, 0], ["Nhãn hiệu", 2, 1, 0], ["Kiểu dáng công nghiệp", 5, 0, 0], ["Sáng chế", 0, 0, 1]]

for _i, _d in enumerate(BIEU_DO["Hình 2.1. Cơ cấu và diễn biến sản phẩm khoa học"]["dong"]):
    _d[1] = [17, 42, 49, 71, 107][_i]
    _d[-1] = [56, 100, 84, 127, 190][_i]
    assert _d[-1] == sum(_d[1:-1]), _d

SO_DO = {
    "Hình 2.3. Phân công đầu mối": ("Hình 2.1", SD.dau_moi),
    "Hình 2.9. Chuỗi chuyển hóa": ("Hình 2.7", SD.pheu),
    "Hình 2.11. Chẩn đoán nguy cơ": ("Hình 2.8", SD.nguyen_nhan),
}
BO_HINH = ["Hình 2.2. Mô phỏng phần dành cho tác giả", "Hình 2.6. Khả năng tính toán bộ tiêu chí",
           "Hình 2.10. Tỷ lệ thực hiện so với chỉ tiêu"]

PPR_HINH = ('<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:keepNext/>'
            '<w:spacing w:before="60" w:after="60" w:line="240" w:lineRule="auto"/><w:ind w:firstLine="0"/>'
            '<w:jc w:val="center"/></w:pPr>')

# Đánh số lại theo thứ tự xuất hiện
DOI_SO = {
    "Hình 2.11": "Hình 2.8",
    "Hình 2.3": "Hình 2.1", "Hình 2.1": "Hình 2.2", "Hình 2.4": "Hình 2.3", "Hình 2.5": "Hình 2.4",
    "Hình 2.8": "Hình 2.5", "Hình 2.7": "Hình 2.6", "Hình 2.9": "Hình 2.7",
    "Bảng 2.4": "Bảng 2.3", "Bảng 2.7": "Bảng 2.4", "Bảng 2.3": "Bảng 2.7",
}


# ---------------------------------------------------------------------------
# Văn bản mới
# ---------------------------------------------------------------------------
T = {}
T["mo"] = ("Chương này đánh giá thực trạng quản lý quyền sở hữu trí tuệ tại Trường Đại học Thành Đô giai đoạn 2021 - 2025 "
           "trên cơ sở rà soát hồ sơ, thống kê dữ liệu khoa học công nghệ của Nhà trường và đối chiếu các quy chế nội bộ "
           "với khung pháp lý hiện hành. Thực trạng được trình bày theo bốn khâu sáng tạo, xác lập, khai thác và bảo vệ đã "
           "xác định ở Chương 1, sau đó đánh giá kết quả, hạn chế và nguyên nhân.")
T["211"] = ("Trường Đại học Thành Đô được thành lập năm 2009 theo Quyết định số 679/QĐ-TTg của Thủ tướng Chính phủ trên cơ "
            "sở nâng cấp Trường Cao đẳng Tư thục Công nghệ Thành Đô (thành lập năm 2004). Nhà trường phát triển theo định "
            "hướng ứng dụng, đa ngành, đạt chuẩn kiểm định chất lượng cơ sở giáo dục và đạt 3 sao theo định hướng ứng dụng "
            "trên bảng xếp hạng UPM năm 2020. Hoạt động khoa học và công nghệ được xác định là một trụ cột phát triển, gắn "
            "với chuyển giao tri thức và giải quyết các bài toán thực tiễn của doanh nghiệp, địa phương.")
T["212a"] = ("Năm 2026, Nhà trường có 252 nhân sự cơ hữu, trong đó 145 giảng viên. Toàn trường có 94 người có trình độ tiến "
             "sĩ và tương đương (37,3%), trong đó 25 giáo sư, phó giáo sư. Đội ngũ tiến sĩ tập trung tại Viện Y - Dược "
             "(42 người) và Viện Quản trị và Công nghệ (29 người), như thể hiện tại Bảng 2.1.")
T["212b"] = ("Số liệu cho thấy Nhà trường có đội ngũ đủ điều kiện để hình thành tài sản trí tuệ, nhất là trong lĩnh vực dược "
             "liệu và công nghệ. Tuy nhiên, Bảng 2.1 không có đơn vị hoặc vị trí làm nghiệp vụ sở hữu trí tuệ, như tra cứu "
             "thông tin sáng chế hay soạn thảo bản mô tả sáng chế.")
T["212c"] = ("Các đơn vị đào tạo, nghiên cứu được tổ chức theo mô hình viện, gồm Viện Y - Dược, Viện Quản trị và Công nghệ, "
             "Viện Ngôn ngữ - Văn hóa - Quốc tế và Viện Nghiên cứu giáo dục và Chuyển giao tri thức. Mỗi viện có Phòng Học "
             "vụ và Hợp tác đối ngoại và Phòng Công nghệ, Đổi mới sáng tạo và Khởi nghiệp.")
T["212d"] = ("Theo Điều 11 Quy chế quản trị tài sản trí tuệ ban "
             "hành kèm theo Quyết định số 217/QĐ-ĐHTĐ, Phòng Khoa học công nghệ quản lý, giám sát hoạt động sở hữu trí tuệ, "
             "xây dựng quy trình, biểu mẫu khai báo, lập hồ sơ theo dõi tài sản trí tuệ và xúc tiến thương mại hóa; Bộ phận "
             "Pháp chế tư vấn pháp lý và thực hiện thủ tục xác lập quyền; người đứng đầu các đơn vị yêu cầu người lao động, "
             "người học ghi nhận tài sản trí tuệ mới phát sinh và phối hợp đăng ký, khai thác. Theo Điều 35 Quy chế hoạt động "
             "khoa học công nghệ ban hành kèm theo Quyết định số 213/QĐ-ĐHTĐ, Phòng Khoa học Công nghệ tiếp nhận đơn của tác "
             "giả, trình Hiệu trưởng ký và nộp đơn, lệ phí tại cơ quan nhà nước có thẩm quyền. Phòng Tài chính - Kế toán "
             "tham gia tham mưu tỷ lệ phân chia lợi ích (Điều 13 Quyết định số 217/QĐ-ĐHTĐ). Hình 2.1 thể hiện sự phân công "
             "này.")
T["212e"] = ("Như vậy, công tác quản lý quyền sở hữu trí tuệ hiện được thực hiện bởi nhiều đầu mối khác nhau, trong khi chưa "
             "hình thành bộ phận chuyên trách. Điều này có thể làm tăng nhu cầu phối hợp và chia sẻ thông tin trong quá "
             "trình quản lý.")
T["213a"] = ("Giai đoạn 2021 - 2025, hoạt động nghiên cứu khoa học của Nhà trường tăng rõ rệt, đặc biệt trong hai năm 2024 - "
             "2025. Số sản phẩm khoa học tăng từ 56 năm 2021 lên 192 năm 2025, gấp 3,4 lần, tương ứng mức tăng bình quân "
             "khoảng 36% mỗi năm. Sau mức giảm năm 2023 (85 sản phẩm), số sản phẩm tăng lên 128 năm 2024 và tiếp tục tăng "
             "50% trong năm 2025 (Hình 2.2, Bảng 2.2).")
T["213b"] = ("Công bố khoa học là hoạt động chủ đạo. Trong kỳ có 405 bài báo, gồm 290 bài trong nước và 115 bài quốc tế, "
             "trong đó 71 bài đăng trên tạp chí có phân hạng; 87 giáo trình, tài liệu giảng dạy; 12 sách có mã ISBN; 37 "
             "tham luận hội thảo (21 quốc gia, 16 quốc tế); 38 đề tài cấp cơ sở và 3 đề tài cấp quốc gia do Quỹ NAFOSTED "
             "tài trợ với tổng kinh phí 4,67 tỷ đồng. Bài báo quốc tế tăng từ 11 bài năm 2023 lên 33 bài năm 2024 và 52 "
             "bài năm 2025; số bài đăng trên tạp chí có phân hạng tăng từ 1 lên 40 bài trong cùng thời gian. Năng lực tạo "
             "ra kết quả nghiên cứu của Nhà trường vì vậy đã tăng đáng kể; các mục sau xem xét bao nhiêu kết quả trong số "
             "đó được nhận diện và xác lập quyền.")
T["213c"] = ("Trong 38 đề tài cấp cơ sở, 19 đề tài được Nhà trường cấp kinh phí với tổng số 424,75 triệu đồng, 16 đề tài chỉ "
             "được quy đổi giờ nghiên cứu khoa học và 3 đề tài tự tìm nguồn tài trợ (Hình 2.3). Năm 2021 chưa có đề tài "
             "được cấp kinh phí; từ năm 2022, kinh phí cấp hằng năm dao động từ 70 đến 145 triệu đồng và cao nhất vào năm "
             "2025. Kinh phí bình quân của một đề tài được cấp tăng từ 28,0 triệu đồng năm 2022 lên 36,25 triệu đồng năm "
             "2025.")
T["221a"] = ("Hoạt động sáng tạo được điều chỉnh bởi Quy chế hoạt động khoa học công nghệ ban hành kèm theo Quyết định số "
             "213/QĐ-ĐHTĐ ngày 28/12/2021, Quy chế quản trị tài sản trí tuệ ban hành kèm theo Quyết định số 217/QĐ-ĐHTĐ năm "
             "2024 và Quy chế chi tiêu nội bộ năm 2026. Các quy chế khuyến khích nghiên cứu bằng hai công cụ là quy đổi giờ "
             "nghiên cứu khoa học và thưởng bằng tiền (Hình 2.4).")
T["221b"] = ("Phụ lục 5 Quy chế chi tiêu nội bộ năm 2026 quy đổi 600 giờ cho bằng độc quyền sáng chế cấp tại Mỹ, châu Âu, "
             "Đông Bắc Á, 360 giờ cho bằng sáng chế cấp tại Việt Nam, Trung Quốc, ASEAN và 180 giờ cho giải pháp hữu ích, "
             "cao hơn mức quy đổi cho bài báo quốc tế (240 đến 300 giờ). Tuy nhiên, giờ vượt định mức không được chi trả "
             "bằng tiền mà chỉ quy đổi thành điểm cộng khi xét thưởng. Về thưởng bằng tiền, Bảng 7 của Quy chế quy định mức "
             "thưởng cho bài báo quốc tế: 20 triệu đồng với bài WoS Q1, 15 triệu đồng Q2, 12 triệu đồng Q3, 10 triệu đồng Q4 "
             "và 8 triệu đồng với bài Scopus; văn bằng bảo hộ không có mức thưởng tương ứng. Hỗ trợ tài chính dành cho sở "
             "hữu công nghiệp hiện có ở khâu xét duyệt đề tài: đề tài có đăng ký sở hữu trí tuệ được cấp tối đa 50 triệu "
             "đồng, với điều kiện có chứng nhận đăng ký thành công khi nghiệm thu. Như vậy, cơ chế khuyến khích hiện nay "
             "tập trung nhiều vào công bố khoa học; đối với hoạt động xác lập quyền sở hữu công nghiệp, Nhà trường chưa có "
             "chính sách thưởng bằng tiền tương ứng.")
T["221c"] = ("Bảng 2.3 cho thấy kinh phí cho xác lập quyền chưa được tách thành dòng dự toán riêng ở cả ba kênh tài trợ. Đối "
             "với đề tài cấp cơ sở, Điều 35 Quyết định số 213/QĐ-ĐHTĐ giao Phòng Khoa học Công nghệ nộp đơn và lệ phí, còn "
             "các khoản tra cứu thông tin sáng chế, soạn bản mô tả hay thuê tổ chức đại diện chỉ có thể bố trí qua mục chi "
             "dịch vụ thuê ngoài phục vụ nghiên cứu (Điều 38), chưa có định mức riêng.")
T["221d"] = ("Trong 38 đề tài cấp cơ sở, 11 đề tài (28,9%) có sản phẩm có thể bảo hộ (Bảng 2.4, Hình 2.5). Chín sản phẩm "
             "thuộc nhóm sở hữu công nghiệp, gồm các quy trình chiết xuất, công thức bào chế và mô hình kỹ thuật; hai sản "
             "phẩm thuộc nhóm quyền tác giả là bộ mẫu cây thuốc và bộ tiêu bản hiển vi. Viện Y - Dược chủ trì 9 trên 11 "
             "đề tài. Đến hết năm 2025, chỉ đề tài 09-2025 về cao chiết lá Quế hoa đã nộp đơn đăng ký sáng chế.")
T["222a"] = ("Bảng 2.5 cho thấy Quyết định số 217/QĐ-ĐHTĐ đã liệt kê đầy đủ các nhóm đối tượng, kể cả bí mật thương mại, "
             "kiểu dáng công nghiệp và thiết kế bố trí mạch tích hợp (Điều 3). Khoảng cách nằm ở thực tế xác lập: các đối "
             "tượng đã được cấp văn bằng chủ yếu là nhãn hiệu và quyền tác giả đối với bộ biểu trưng, trong khi 8 sản phẩm "
             "có thể đăng ký giải pháp hữu ích hoặc sáng chế từ đề tài chưa được nộp đơn.")
T["222b"] = ("Tính đến hết năm 2025, Nhà trường có 11 hồ sơ tài sản trí tuệ (Bảng 2.6, Hình 2.6). Bốn hồ sơ đã được cấp văn "
             "bằng, gồm 2 giấy chứng nhận quyền tác giả đối với bộ biểu trưng và 2 văn bằng nhãn hiệu Thanh do University, "
             "Thado Edupark. Sáu hồ sơ đồng sở hữu với doanh nghiệp đang chờ cấp bằng theo sổ theo dõi, gồm nhãn hiệu "
             "Double2n và 5 kiểu dáng công nghiệp bao bì sản phẩm thảo dược. Đơn sáng chế số 1-2025-07378 về hợp chất từ "
             "lá Quế hoa đã được chấp nhận hợp lệ. Trong kỳ, Nhà trường chưa có bằng độc quyền sáng chế hoặc giải pháp hữu "
             "ích nào.")
T["222c"] = ("Hình 2.7 cho thấy điểm nghẽn nằm ở bước chuyển từ kết quả nghiên cứu sang đơn đăng ký: trong 9 đề tài có sản "
             "phẩm thuộc nhóm sở hữu công nghiệp, mới có 1 đề tài nộp đơn (11,1%). Tám đề tài còn lại đã nghiệm thu nhưng "
             "chưa nộp đơn, trong đó có đề tài nghiệm thu từ năm 2022. Đây là phát hiện chính ở khâu xác lập quyền: năng "
             "lực tạo ra kết quả có thể bảo hộ đã hình thành, nhưng việc chuyển kết quả thành hồ sơ đăng ký chưa được thực "
             "hiện thường xuyên.")
T["223a"] = "Khai thác quyền sở hữu trí tuệ được đánh giá qua quy định phân chia lợi ích và kết quả khai thác thực tế."
T["223b"] = ("Bảng 2.7 cho thấy các quy định về phân chia lợi ích đã có nhưng nằm rải rác trong nhiều văn bản. Khoản 4 Điều "
             "36 Quyết định số 213/QĐ-ĐHTĐ có hai điểm áp dụng cho hai phạm vi khác nhau: điểm a áp dụng cho đề tài sử dụng "
             "ngân sách nhà nước, chia 40% nộp ngân sách, 30% cho Nhà trường, 30% thưởng tập thể tác giả và giới hạn tiền "
             "thưởng tối đa 100 triệu đồng mỗi đề tài; điểm b áp dụng cho tài sản trí tuệ thuộc sở hữu của Trường, tác giả "
             "hưởng 30% kinh phí chuyển giao sau chi phí, không có mức trần. Từ ngày 01/10/2025, điểm a khoản 3 Điều 28 Luật "
             "số 93/2025/QH15 yêu cầu thưởng cho tác giả tối thiểu 30% lợi nhuận thu được từ phần kết quả sử dụng ngân sách "
             "nhà nước, không đặt mức trần; vì vậy, mức trần 100 triệu đồng tại điểm a cần được rà soát. Quyết định số "
             "217/QĐ-ĐHTĐ chỉ quy định trường hợp không có thỏa thuận thì Hiệu trưởng quyết định tỷ lệ phân chia.")
T["223c"] = ("Về kết quả thực tế, giai đoạn 2021 - 2025 Nhà trường chưa ký hợp đồng chuyển nhượng, chuyển quyền sử dụng sáng "
             "chế hoặc hợp đồng chuyển giao công nghệ nào phát sinh doanh thu, và chưa có doanh nghiệp khởi nguồn từ kết quả "
             "nghiên cứu. Các tài sản đã được xác lập là nhãn hiệu và quyền tác giả đối với bộ biểu trưng được sử dụng cho "
             "nhận diện thương hiệu và truyền thông tuyển sinh. Do chưa có nguồn thu từ khai thác, các quy định phân chia "
             "lợi ích nêu trên chưa được áp dụng trên thực tế.")
T["224a"] = ("Quyết định số 217/QĐ-ĐHTĐ đã hình thành khung bảo vệ quyền ở cấp quy chế. Điều 10 quy định nghĩa vụ bảo mật "
             "của mọi đơn vị, người lao động, người học, cộng tác viên và yêu cầu tác giả xin ý kiến Phòng Khoa học công "
             "nghệ trước khi bộc lộ công khai tài sản có thể bảo hộ; Điều 14 liệt kê các hành vi xâm phạm quyền tác giả; "
             "Điều 15 quy định các biện pháp tự bảo vệ, yêu cầu cơ quan nhà nước xử lý và khởi kiện; Điều 16 quy định xử lý "
             "kỷ luật và khuyến khích hòa giải tranh chấp nội bộ. Hai nhãn hiệu đã được cấp văn bằng là căn cứ pháp lý để "
             "Nhà trường bảo vệ tên gọi và hình ảnh của mình.")
T["224b"] = ("Tuy nhiên, các quy định trên mới dừng ở mức nguyên tắc. Hồ sơ của Nhà trường chưa có quy trình, biểu mẫu để "
             "thực hiện việc xin ý kiến trước khi công bố theo Điều 10, chưa có cơ chế theo dõi hành vi xâm phạm nhãn hiệu "
             "trên môi trường số và chưa có hệ thống lưu trữ tập trung để theo dõi thời hạn, nghĩa vụ phí của văn bằng.")
T["224c"] = ("Dữ liệu quản lý cũng còn phân tán. Các danh mục hiện có cho phép thống kê số công bố, số đề tài có sản phẩm có "
             "thể bảo hộ, số đơn đăng ký và số văn bằng, nhưng chưa ghi nhận doanh thu khai thác, thù lao trả cho tác giả, "
             "chi phí xác lập và duy trì quyền. Đây là các thông tin mà Danh mục quyền sở hữu trí tuệ theo Điều 9a Nghị định "
             "số 65/2023/NĐ-CP, được bổ sung bởi Nghị định số 100/2026/NĐ-CP, yêu cầu chủ sở hữu ghi nhận.")
T["224d"] = ("Nhìn chung, Nhà trường đã bước đầu hình thành các biện pháp kiểm soát liêm chính học thuật và quản lý quyền sở "
             "hữu trí tuệ. Tuy nhiên, dữ liệu theo dõi và cơ chế quản trị tập trung đối với tài sản trí tuệ vẫn cần tiếp tục "
             "hoàn thiện.")
T["231b"] = ("Hai là, tiềm lực nghiên cứu tăng rõ rệt: số sản phẩm khoa học tăng từ 56 năm 2021 lên 192 năm 2025, trong đó "
             "có 115 bài báo quốc tế trong kỳ; đội ngũ 94 tiến sĩ và tương đương, tập trung tại Viện Y - Dược và Viện Quản "
             "trị và Công nghệ, là nguồn tạo ra giải pháp kỹ thuật.")
T["231c"] = ("Ba là, công tác xác lập quyền đã có kết quả cụ thể với 11 hồ sơ tài sản trí tuệ, trong đó 2 nhãn hiệu và 2 "
             "giấy chứng nhận quyền tác giả đã được cấp, và lần đầu tiên có đơn đăng ký sáng chế từ kết quả đề tài cấp cơ "
             "sở (đề tài 09-2025).")
T["231d"] = ("Bốn là, trong hai năm 2024 - 2025, Nhà trường đạt hoặc vượt 6 trên 9 chỉ tiêu của Kế hoạch số 07/KH-ĐHTĐ "
             "(Hình 2.8): bài báo quốc tế đạt 283% (85/30 bài), sách có mã ISBN 267% (8/3 cuốn), bài kỷ yếu hội thảo cấp "
             "Trường 223% (134/60 bài), bài báo trong nước 165% (181/110 bài), đề tài cấp Bộ, Nhà nước 150% (3/2 đề tài) "
             "và đề tài cấp cơ sở 107% (15/14 đề tài). Chỉ tiêu tham luận quốc tế và giáo trình đạt 90 - 91%; chỉ tiêu "
             "chuyển giao công nghệ chưa phát sinh.")
T["232a"] = ("Thứ nhất, kết quả nghiên cứu có thể bảo hộ chưa được chuyển thành đơn đăng ký. Trong 9 đề tài cấp cơ sở có sản "
             "phẩm thuộc nhóm sở hữu công nghiệp, mới có 1 đề tài nộp đơn sáng chế; 8 đề tài còn lại đã nghiệm thu nhưng "
             "chưa nộp đơn. Hai sản phẩm thuộc nhóm quyền tác giả là bộ mẫu cây thuốc và bộ tiêu bản hiển vi chưa được đăng "
             "ký.")
T["232b"] = ("Thứ hai, quy chế nội bộ chưa đồng bộ với khung pháp lý mới. Mức trần 100 triệu đồng tại điểm a khoản 4 Điều 36 "
             "Quyết định số 213/QĐ-ĐHTĐ chưa phù hợp với yêu cầu thưởng cho tác giả tối thiểu 30% lợi nhuận theo điểm a "
             "khoản 3 Điều 28 Luật số 93/2025/QH15 đối với kết quả sử dụng ngân sách nhà nước; văn bằng bảo hộ chưa có mức "
             "thưởng bằng tiền; chi phí tra cứu, soạn đơn và duy trì văn bằng chưa có dòng dự toán riêng.")
T["232c"] = ("Thứ ba, chưa có hoạt động khai thác thương mại phát sinh doanh thu: chưa có hợp đồng chuyển nhượng, chuyển quyền "
             "sử dụng hoặc chuyển giao công nghệ, chưa có doanh nghiệp khởi nguồn từ kết quả nghiên cứu.")
T["232d"] = ("Thứ tư, mô hình quản lý phân tán, chưa có bộ phận và nhân sự chuyên trách; dữ liệu về tài sản trí tuệ nằm rải "
             "rác trong nhiều danh mục, chưa có cơ sở dữ liệu theo dõi vòng đời tài sản.")
T["233a"] = "Các hạn chế trên xuất phát từ cả nguyên nhân khách quan và nguyên nhân chủ quan (Hình 2.8)."
T["233b"] = ("Về nguyên nhân khách quan. Thứ nhất, thủ tục xác lập quyền đối với sáng chế kéo dài: theo Điều 110 và Điều 119 "
             "Luật Sở hữu trí tuệ, đơn sáng chế hợp lệ được công bố vào tháng thứ mười chín kể từ ngày nộp đơn, trừ khi có "
             "yêu cầu công bố sớm, và được thẩm định nội dung trong mười hai tháng kể từ ngày công bố hoặc ngày có yêu cầu "
             "thẩm định, chưa kể thời gian sửa đổi, bổ sung đơn. Thứ hai, chi phí bảo hộ phát sinh ở nhiều khâu, gồm tra cứu, "
             "soạn bản mô tả, phí và lệ phí nộp đơn, thẩm định, cấp văn bằng, và phí, lệ phí duy trì hiệu lực hằng năm đối "
             "với bằng sáng chế, giải pháp hữu ích (Điều 94 Luật Sở hữu trí tuệ), trong khi kinh phí bình quân của một đề "
             "tài cấp cơ sở được cấp vốn chỉ khoảng 22 triệu đồng. Thứ ba, thị trường chuyển giao công nghệ và quyền sở hữu "
             "trí tuệ còn hạn chế; phát triển thị trường quyền sở hữu trí tuệ, thúc đẩy thương mại hóa kết quả nghiên cứu "
             "vẫn là nhiệm vụ được Kết luận số 51-KL/TW yêu cầu đẩy mạnh.")
T["233c"] = ("Về nguyên nhân chủ quan. Thứ nhất, chưa có quy trình sàng lọc trước công bố: Điều 10 Quyết định số 217/QĐ-ĐHTĐ "
             "yêu cầu tác giả xin ý kiến Phòng Khoa học công nghệ trước khi bộc lộ, nhưng chưa có trình tự, biểu mẫu và thời "
             "hạn thực hiện. Thứ hai, chưa có đầu mối chuyên trách: nhiệm vụ quản lý tài sản trí tuệ được giao kiêm nhiệm "
             "cho Phòng Khoa học công nghệ và Bộ phận Pháp chế. Thứ ba, cơ chế hỗ trợ xác lập quyền còn hạn chế: văn bằng bảo "
             "hộ không có mức thưởng bằng tiền như bài báo quốc tế, chi phí xác lập quyền chưa có dòng dự toán riêng. Thứ "
             "tư, dữ liệu quản lý phân tán ở nhiều danh mục, chưa ghi nhận chi phí, doanh thu và thời hạn của từng tài sản.")
T["tk1"] = ("Giai đoạn 2021 - 2025, Trường Đại học Thành Đô đã hình thành tiềm lực nghiên cứu đáng kể: 94 tiến sĩ và tương "
            "đương, số sản phẩm khoa học tăng từ 56 (năm 2021) lên 192 (năm 2025), 38 đề tài cấp cơ sở, trong đó 11 đề tài có sản phẩm có "
            "thể bảo hộ. Khung quy chế về sở hữu trí tuệ đã có, gồm Quyết định số 213/QĐ-ĐHTĐ, Quyết định số 217/QĐ-ĐHTĐ và "
            "Quy chế chi tiêu nội bộ năm 2026.")
T["tk2"] = ("Điểm nghẽn chính nằm ở khâu xác lập và khai thác: trong 9 đề tài có sản phẩm thuộc nhóm sở hữu công nghiệp mới "
            "có 1 đơn sáng chế, chưa có bằng độc quyền sáng chế hoặc giải pháp hữu ích, chưa có hợp đồng chuyển giao phát "
            "sinh doanh thu. Nguyên nhân chủ quan gồm chưa có quy trình sàng lọc trước công bố, chưa có đầu mối chuyên trách, "
            "cơ chế hỗ trợ xác lập quyền còn hạn chế và dữ liệu phân tán; nguyên nhân khách quan gồm thủ tục xác lập quyền "
            "kéo dài, chi phí bảo hộ và thị trường chuyển giao công nghệ còn hạn chế. Đây là căn cứ để đề xuất giải pháp ở "
            "Chương 3.")


# --- Vòng góp ý thứ hai: rút gọn 2.1.2, 2.2.3 (Bảng 2.7), 2.3.1; Mở đầu bớt số liệu
T["212a"] = ("Năm 2026, Nhà trường có 252 nhân sự cơ hữu, trong đó 145 giảng viên và 94 người có trình độ tiến sĩ và tương "
             "đương (37,3%). Đội ngũ tiến sĩ tập trung tại Viện Y - Dược (42 người) và Viện Quản trị và Công nghệ (29 "
             "người), như thể hiện tại Bảng 2.1. Đây là nguồn lực đủ điều kiện để hình thành tài sản trí tuệ, nhất là trong "
             "lĩnh vực dược liệu và công nghệ.")
T["212d"] = ("Theo Điều 11, Điều 13 Quyết định số 217/QĐ-ĐHTĐ và Điều 35 Quyết định số 213/QĐ-ĐHTĐ, công tác quản lý quyền "
             "sở hữu trí tuệ được phân công cho Phòng Khoa học công nghệ (đầu mối quản lý, tiếp nhận và nộp đơn, theo dõi "
             "tài sản), Bộ phận Pháp chế (tư vấn pháp lý, thực hiện thủ tục xác lập quyền), Phòng Tài chính - Kế toán "
             "(tham mưu phân chia lợi ích) và các viện (yêu cầu ghi nhận tài sản trí tuệ mới phát sinh), như thể hiện tại "
             "Hình 2.1.")
T["223b"] = ("Bảng 2.7 cho thấy quy định về phân chia lợi ích nằm ở ba văn bản nội bộ và chưa thống nhất với nhau. Điểm bất "
             "cập rõ nhất là mức trần 100 triệu đồng tại điểm a khoản 4 Điều 36 Quyết định số 213/QĐ-ĐHTĐ: từ ngày "
             "01/10/2025, điểm a khoản 3 Điều 28 Luật số 93/2025/QH15 yêu cầu thưởng cho tác giả tối thiểu 30% lợi nhuận "
             "từ phần kết quả sử dụng ngân sách nhà nước, không đặt mức trần.")
T["231d"] = ("Bốn là, các chỉ tiêu công bố khoa học của Kế hoạch số 07/KH-ĐHTĐ giai đoạn 2024 - 2025 đều đạt hoặc vượt kế "
             "hoạch, trong khi chỉ tiêu chuyển giao công nghệ chưa phát sinh.")
T["md1"] = ("Trường Đại học Thành Đô phát triển theo định hướng ứng dụng, với hoạt động nghiên cứu khoa học tăng nhanh trong "
            "giai đoạn 2021 - 2025. Rà soát 38 đề tài cấp cơ sở của giai đoạn này cho thấy 11 đề tài có sản phẩm có thể "
            "bảo hộ quyền sở hữu trí tuệ, trong đó 9 đề tài có sản phẩm thuộc nhóm sở hữu công nghiệp, chủ yếu trong lĩnh "
            "vực dược liệu.")
T["md2"] = ("Tuy nhiên, trong 9 đề tài này mới có 1 đề tài nộp đơn đăng ký sáng chế, và Nhà trường chưa có bằng độc quyền "
            "sáng chế hay giải pháp hữu ích nào. Khoảng cách giữa năng lực tạo ra kết quả nghiên cứu và kết quả xác lập "
            "quyền cho thấy vấn đề nằm ở khâu quản lý: dữ liệu về tài sản trí tuệ còn phân tán, việc phân công giữa các "
            "đầu mối chưa gắn với một quy trình chung và chưa có quy trình sàng lọc khả năng bảo hộ trước khi công bố kết "
            "quả nghiên cứu.")

# --- Vòng đối chiếu nguồn lần 3 (danh mục gốc trong Tai lieu thanh do)
T["213b"] = ("Công bố khoa học là hoạt động chủ đạo. Danh mục ghi nhận 405 bản ghi bài báo, gồm 290 bản ghi trong nước và "
             "115 bài quốc tế, trong đó 71 bài đăng trên tạp chí có phân hạng; 87 giáo trình, tài liệu giảng dạy; 12 sản "
             "phẩm thuộc nhóm sách, chương sách và tài liệu xuất bản có ISBN; 37 tham luận hội thảo (21 quốc gia, 16 quốc "
             "tế); 38 đề tài cấp cơ sở và 3 đề tài cấp quốc gia do Quỹ NAFOSTED tài trợ với tổng kinh phí được phê duyệt "
             "4,67 tỷ đồng. Bài báo quốc tế tăng từ 11 bài năm 2023 lên 33 bài năm 2024 và 52 bài năm 2025; số bài đăng "
             "trên tạp chí có phân hạng tăng từ 1 lên 40 bài trong cùng thời gian. Năng lực tạo ra kết quả nghiên cứu của "
             "Nhà trường vì vậy đã tăng đáng kể; các mục sau xem xét bao nhiêu kết quả trong số đó được nhận diện và xác "
             "lập quyền.")
T["213c"] = ("Theo danh mục đề tài cấp cơ sở, tổng kinh phí bằng tiền ghi cho 19 đề tài là 424,75 triệu đồng, chưa gồm khoản "
             "20 triệu đồng đề nghị hỗ trợ thêm của đề tài 09-2025; 16 đề tài chỉ được quy đổi giờ nghiên cứu khoa học và 3 "
             "đề tài tự tìm nguồn tài trợ (Hình 2.3). Năm 2021 chưa có đề tài được cấp kinh phí; từ năm 2022, kinh phí "
             "hằng năm dao động từ 70 đến 145 triệu đồng và cao nhất vào năm 2025. Kinh phí bình quân của một đề tài được "
             "cấp tăng từ 28,0 triệu đồng năm 2022 lên 36,25 triệu đồng năm 2025.")
T["221d"] = T["221d"].replace("11 đề tài (28,9%) có sản phẩm có thể bảo hộ",
                              "11 đề tài (28,9%) có sản phẩm mà nhóm nghiên cứu đánh giá là có thể bảo hộ")
T["222b"] = ("Tính đến hết năm 2025, Nhà trường có 11 hồ sơ tài sản trí tuệ (Bảng 2.6, Hình 2.6). Bốn hồ sơ đã được cấp văn "
             "bằng, gồm 2 giấy chứng nhận quyền tác giả đối với bộ biểu trưng và 2 văn bằng nhãn hiệu Thanh do University, "
             "Thado Edupark. Sáu hồ sơ đồng sở hữu với doanh nghiệp được sổ theo dõi ghi trạng thái chờ cấp bằng, gồm nhãn "
             "hiệu Double2n và 5 kiểu dáng công nghiệp bao bì sản phẩm thảo dược. Đơn sáng chế số 1-2025-07378 về hợp chất "
             "từ lá Quế hoa đã được nộp năm 2025. Trong kỳ, Nhà trường chưa có bằng độc quyền sáng chế hoặc giải pháp hữu "
             "ích nào.")
T["223c"] = ("Về kết quả thực tế, hồ sơ đã đối chiếu chưa có hợp đồng chuyển nhượng, chuyển quyền sử dụng hay chứng từ doanh "
             "thu từ khai thác tài sản trí tuệ trong kỳ, và chưa có doanh nghiệp khởi nguồn từ kết quả nghiên cứu. Danh mục "
             "đề tài ghi đề tài 08-2023 có sản phẩm chuyển giao công nghệ cho Nhà trường thương mại hóa, nhưng chưa có hồ "
             "sơ xác định hợp đồng và doanh thu. Các tài sản đã được xác lập là nhãn hiệu và quyền tác giả đối với bộ biểu "
             "trưng được sử dụng cho nhận diện thương hiệu và truyền thông tuyển sinh. Vì vậy, các quy định phân chia lợi "
             "ích nêu trên chưa có căn cứ để đánh giá qua thực tế áp dụng.")
T["224b"] = T["224b"].replace("Hồ sơ của Nhà trường chưa có quy trình", "Trong hồ sơ nhóm nghiên cứu tiếp cận được, chưa có quy trình")
T["231d"] = ("Bốn là, trong hai năm 2024 - 2025, phần lớn chỉ tiêu công bố của Kế hoạch số 07/KH-ĐHTĐ đạt hoặc vượt kế hoạch; "
             "chỉ tiêu tham luận hội thảo quốc tế đạt 10 trên 11, chỉ tiêu chuyển giao công nghệ chưa phát sinh.")
T["232c"] = ("Thứ ba, chưa ghi nhận hoạt động khai thác thương mại phát sinh doanh thu: hồ sơ chưa có hợp đồng chuyển nhượng, "
             "chuyển quyền sử dụng hoặc chuyển giao công nghệ có thu phí, chưa có doanh nghiệp khởi nguồn từ kết quả nghiên "
             "cứu.")
T["tk2"] = T["tk2"].replace("chưa có hợp đồng chuyển giao phát sinh doanh thu", "chưa ghi nhận hợp đồng chuyển giao phát sinh doanh thu")

# --- Vòng đối chiếu lần 4: 5 kiểu dáng đã được cấp bằng năm 2024 (xác nhận của chủ nhiệm đề tài, số bằng
#     3-00396xx-000 trong bảng thống kê văn bằng); phạm vi kinh phí 424,75 triệu; cách gọi nhân sự
T["212a"] = ("Danh sách nhân sự năm 2026 có 252 người, trong đó 145 người giữ chức vụ giảng viên và 94 người có trình độ "
             "tiến sĩ và tương đương (37,3%). Đội ngũ tiến sĩ tập trung tại Viện Y - Dược (42 người) và Viện Quản trị và "
             "Công nghệ (29 người), như thể hiện tại Bảng 2.1. Đây là nguồn lực đủ điều kiện để hình thành tài sản trí "
             "tuệ, nhất là trong lĩnh vực dược liệu và công nghệ.")
T["213c"] = ("Theo danh mục đề tài cấp cơ sở, tổng các khoản kinh phí ghi cho 19 đề tài là 424,75 triệu đồng, không bao gồm "
             "kinh phí tự tìm tài trợ và khoản đề nghị hỗ trợ thêm; 16 đề tài chỉ được quy đổi giờ nghiên cứu khoa học và 3 "
             "đề tài tự tìm nguồn tài trợ (Hình 2.3). Năm 2021 chưa có đề tài được cấp kinh phí; từ năm 2022, kinh phí "
             "hằng năm dao động từ 70 đến 145 triệu đồng và cao nhất vào năm 2025. Kinh phí bình quân của một đề tài được "
             "cấp tăng từ 28,0 triệu đồng năm 2022 lên 36,25 triệu đồng năm 2025.")
T["222a"] = ("Bảng 2.5 cho thấy Quyết định số 217/QĐ-ĐHTĐ đã liệt kê đầy đủ các nhóm đối tượng, kể cả bí mật thương mại, "
             "kiểu dáng công nghiệp và thiết kế bố trí mạch tích hợp (Điều 3). Khoảng cách nằm ở thực tế xác lập: các văn "
             "bằng đã được cấp là quyền tác giả đối với bộ biểu trưng, nhãn hiệu và kiểu dáng công nghiệp hình thành từ hợp "
             "tác với doanh nghiệp, trong khi 8 sản phẩm có thể đăng ký giải pháp hữu ích hoặc sáng chế từ đề tài chưa được "
             "nộp đơn.")
T["222b"] = ("Tính đến hết năm 2025, Nhà trường có 11 hồ sơ tài sản trí tuệ (Bảng 2.6, Hình 2.6), trong đó 9 hồ sơ đã được "
             "cấp văn bằng: 2 giấy chứng nhận quyền tác giả đối với bộ biểu trưng, 2 văn bằng nhãn hiệu Thanh do University, "
             "Thado Edupark và 5 bằng độc quyền kiểu dáng công nghiệp bao bì sản phẩm thảo dược, đồng sở hữu với doanh "
             "nghiệp, cấp năm 2024. Nhãn hiệu Double2n, đồng sở hữu với doanh nghiệp, đang chờ cấp bằng; đơn sáng chế số "
             "1-2025-07378 về hợp chất từ lá Quế hoa đã được nộp năm 2025. Như vậy, các văn bằng đã có đều thuộc nhóm "
             "thương hiệu và hợp tác doanh nghiệp; trong kỳ, Nhà trường chưa có bằng độc quyền sáng chế hoặc giải pháp hữu "
             "ích nào từ kết quả đề tài.")
T["231c"] = ("Ba là, công tác xác lập quyền đã có kết quả cụ thể: 9 trên 11 hồ sơ tài sản trí tuệ đã được cấp văn bằng, gồm "
             "2 giấy chứng nhận quyền tác giả, 2 nhãn hiệu và 5 kiểu dáng công nghiệp đồng sở hữu với doanh nghiệp; lần đầu "
             "tiên có đơn đăng ký sáng chế từ kết quả đề tài cấp cơ sở (đề tài 09-2025).")

# --- Loại bài trùng (quyết định của chủ nhiệm đề tài): danh mục bài báo trong nước có 3 bài ghi lặp, 4 bản ghi thừa.
#     Giữ mỗi bài ở năm xuất bản: graphene oxide (11/2024) bỏ 1 bản ghi 2025 và 1 bản ghi 2024; Paederia foetida
#     (28/12/2022) bỏ bản ghi 2023; Chlorpheniramine (9/2023) bỏ bản ghi 2025.
#     Trong nước: 17, 42, 49, 71, 107 = 286; tổng theo năm 56, 100, 84, 127, 190; cộng 21 tham luận quốc gia = 578.
TRONG_NUOC = [17, 42, 49, 71, 107]
TONG_NAM = [56, 100, 84, 127, 190]
T["213a"] = ("Giai đoạn 2021 - 2025, hoạt động nghiên cứu khoa học của Nhà trường tăng rõ rệt, đặc biệt trong hai năm 2024 - "
             "2025. Số sản phẩm khoa học tăng từ 56 năm 2021 lên 190 năm 2025, gấp 3,4 lần, tương ứng mức tăng bình quân "
             "khoảng 36% mỗi năm. Sau mức giảm năm 2023 (84 sản phẩm), số sản phẩm tăng lên 127 năm 2024 và tiếp tục tăng "
             "khoảng 50% trong năm 2025 (Hình 2.2, Bảng 2.2).")
T["213b"] = T["213b"].replace("Danh mục ghi nhận 405 bản ghi bài báo, gồm 290 bản ghi trong nước và 115 bài quốc tế",
                              "Trong kỳ có 401 bài báo, gồm 286 bài trong nước và 115 bài quốc tế")
T["231b"] = T["231b"].replace("lên 192 năm 2025", "lên 190 năm 2025")
T["tk1"] = T["tk1"].replace("lên 192 (năm 2025)", "lên 190 (năm 2025)")

# --- Điều lệ Quỹ Học bổng sau tiến sĩ Ngô Xuân Độ (kèm Nghị quyết số 40/NQ-HĐT-ĐHTĐ ngày 27/5/2025 và quyết định
#     ủy quyền cho Viện Nghiên cứu giáo dục và Chuyển giao tri thức điều hành): Điều 3, 9, 10 khớp số liệu báo cáo.
#     Văn bản này dẫn Quyết định 679/QĐ-TTg ngày 27/5/2009.
T["211"] = T["211"].replace("được thành lập năm 2009 theo Quyết định số 679/QĐ-TTg của Thủ tướng Chính phủ",
                            "được thành lập theo Quyết định số 679/QĐ-TTg ngày 27/5/2009 của Thủ tướng Chính phủ")
T["223b"] = ("Bảng 2.7 cho thấy quy định về phân chia lợi ích nằm ở bốn văn bản nội bộ và chưa thống nhất với nhau: tác giả "
             "được 30% kinh phí chuyển giao theo điểm b khoản 4 Điều 36 Quyết định số 213/QĐ-ĐHTĐ, nhưng chỉ được 20% từ năm "
             "thứ hai theo Điều lệ Quỹ Học bổng sau tiến sĩ Ngô Xuân Độ. Điểm bất cập rõ nhất là mức trần 100 triệu đồng tại "
             "điểm a khoản 4 Điều 36 Quyết định số 213/QĐ-ĐHTĐ: từ ngày 01/10/2025, điểm a khoản 3 Điều 28 Luật số "
             "93/2025/QH15 yêu cầu thưởng cho tác giả tối thiểu 30% lợi nhuận từ phần kết quả sử dụng ngân sách nhà nước, "
             "không đặt mức trần.")

# Bảng 2.7 rút gọn: phạm vi áp dụng, tỷ lệ phân chia, điểm bất cập
BANG27 = dict(
    dau="Văn bản, điều khoản",
    rong=[2000, 2100, 2550, 2420],
    cot=["Văn bản", "Phạm vi áp dụng", "Tỷ lệ phân chia", "Điểm bất cập"],
    dong=[
        ["Quyết định 213, điểm a khoản 4 Điều 36", "Đề tài sử dụng ngân sách nhà nước",
         "40% ngân sách nhà nước; 30% Nhà trường; 30% tác giả, tối đa 100 triệu đồng mỗi đề tài",
         "Mức trần 100 triệu đồng chưa phù hợp yêu cầu tối thiểu 30% của Luật số 93/2025/QH15"],
        ["Quyết định 213, điểm b khoản 4 Điều 36", "Tài sản trí tuệ thuộc sở hữu của Trường",
         "Tác giả 30%; đơn vị có tác giả 20%; Quỹ nghiên cứu khoa học 50%",
         "Chỉ áp dụng cho kinh phí chuyển giao; chưa quy định khi Trường tự sản xuất, kinh doanh hoặc góp vốn"],
        ["Quyết định 217, Điều 13", "Tài sản trí tuệ thuộc sở hữu của Trường, khi không có thỏa thuận",
         "Hiệu trưởng quyết định tỷ lệ", "Chưa có tỷ lệ cụ thể; chưa dẫn chiếu tỷ lệ tại Quyết định 213"],
        ["Quy chế chi tiêu nội bộ 2026, mục 6.3.5.1", "Đề tài cấp cơ sở có đăng ký sở hữu trí tuệ",
         "Trích 50% kinh phí chuyển giao công nghệ về Nhà trường", "Chưa quy định phần của tác giả"],
        ["Điều lệ Quỹ Học bổng sau tiến sĩ Ngô Xuân Độ, Điều 9", "Kết quả nghiên cứu ứng dụng do Quỹ tài trợ",
         "Nghiệm thu năm đầu: tác giả 50%, Trường 50%; từ năm thứ hai: tác giả 20%, Trường 80%",
         "Tỷ lệ cho tác giả từ năm thứ hai thấp hơn mức 30% tại điểm b Điều 36 Quyết định 213"],
        ["Luật số 93/2025/QH15, điểm a khoản 3 Điều 28 (đối chiếu)", "Phần kết quả sử dụng ngân sách nhà nước",
         "Tác giả tối thiểu 30% lợi nhuận, không đặt mức trần", "Mốc đối chiếu cho điểm a Điều 36 Quyết định 213"],
    ],
)

# (đầu đoạn gốc, văn bản mới); None = xóa cả đoạn
DOAN = [
    ("Đánh giá thực trạng quản lý quyền sở hữu trí tuệ tại Trường Đại học Thành Đô là nội dung", T["mo"]),
    ("Trường Đại học Thành Đô được thành lập theo Quyết định số 679", T["211"]),
    ("Với triết lý giáo dục \"Trí - Năng - Nhân - Hòa\"", None),
    ("Tính đến năm 2026, tổng số nhân sự cơ hữu", T["212a"]),
    ("Số liệu Bảng 2.1 khẳng định nguồn nhân lực", None),
    ("Cơ cấu các đơn vị nghiên cứu và đào tạo của Nhà trường", None),
    ("Về mô hình quản lý quyền sở hữu trí tuệ, Nhà trường chưa thành lập", T["212d"]),
    ("(Nguồn: Nhóm nghiên cứu tổng hợp từ Quyết định số 213/QĐ-ĐHTĐ và Quyết định số 217/QĐ-ĐHTĐ)",
     "(Nguồn: Nhóm nghiên cứu tổng hợp từ Điều 11, Điều 13 Quyết định số 217/QĐ-ĐHTĐ và Điều 35 Quyết định số "
     "213/QĐ-ĐHTĐ)"),
    ("Thực trạng phân công tại Hình 2.3", T["212e"]),
    ("Giai đoạn 2021 - 2025 ghi nhận bước phát triển nhảy vọt", T["213a"]),
    ("(Nguồn: Nhóm nghiên cứu tổng hợp từ cơ sở dữ liệu khoa học công nghệ",
     "(Nguồn: Nhóm nghiên cứu tổng hợp từ các danh mục thống kê sản phẩm khoa học của Phòng Khoa học Công nghệ. Không gồm "
     "21 tham luận hội thảo quốc gia do danh mục không ghi năm)"),
    ("(Nguồn: Nhóm nghiên cứu tổng hợp từ các danh mục thống kê sản phẩm khoa học của Phòng",
     "(Nguồn: Nhóm nghiên cứu tổng hợp từ các danh mục thống kê sản phẩm khoa học của Phòng Khoa học Công nghệ. Tham luận "
     "hội thảo quốc gia không ghi năm nên chỉ ghi tổng của kỳ; đề tài cấp cơ sở tính theo năm ghi trong mã số đề tài (đề "
     "tài 14-2024 nghiệm thu năm 2025); đề tài cấp quốc gia tính theo năm phê duyệt kinh phí. Dòng tổng cộng là tổng số "
     "bản ghi; bài báo trong nước đã loại 4 bản ghi trùng lặp của 3 bài, mỗi bài giữ ở năm xuất bản)"),
    ("Theo số liệu Bảng 2.2 và Hình 2.1, số lượng sản phẩm", T["213b"]),
    ("Hoạt động nghiên cứu khoa học cấp cơ sở có sự chuyển biến căn bản", T["213c"]),
    ("Khung thể chế điều chỉnh hoạt động sáng tạo khoa học công nghệ", T["221a"]),
    ("(Nguồn: Nhóm nghiên cứu trích xuất từ Quyết định số 213",
     "(Nguồn: Nhóm nghiên cứu tổng hợp từ Phụ lục 5 và Bảng 7 Quy chế chi tiêu nội bộ năm 2026 của Trường Đại học Thành "
     "Đô)"),
    ("Quy chế hiện hành quy định mức quy đổi giờ nghiên cứu khoa học", T["221b"]),
    ("Bảng 2.4 phân tích ba kênh tài trợ nghiên cứu", T["221c"]),
    ("Mặc dù thiếu vắng động lực tài chính trực tiếp", T["221d"]),
    ("Bảng 2.5 đối sánh giữa quy định của Luật Sở hữu trí tuệ", T["222a"]),
    ("(Nguồn: Nhóm nghiên cứu tổng hợp từ sổ theo dõi nhãn hiệu, kiểu dáng công nghiệp và thống kê văn bằng",
     "(Nguồn: Nhóm nghiên cứu tổng hợp từ sổ theo dõi đơn nhãn hiệu, kiểu dáng công nghiệp, sáng chế và bảng thống kê "
     "văn bằng của Trường Đại học Thành Đô. Hồ sơ số 12 nộp năm 2026, nằm ngoài kỳ đánh giá)"),
    ("Theo Bảng 2.6 và Hình 2.7, tính đến hết năm 2025", T["222b"]),
    ("(Nguồn: Nhóm nghiên cứu tổng hợp từ danh mục thuyết minh và quyết toán đề tài cấp cơ sở)",
     "(Nguồn: Nhóm nghiên cứu tổng hợp từ danh mục đề tài cấp cơ sở giai đoạn 2021 - 2025. Năm tính theo mã số đề tài; "
     "kinh phí là số ghi trong danh mục, chưa đối chiếu chứng từ giải ngân, quyết toán)"),
    ("(Nguồn: Nhóm nghiên cứu phân loại từ 11 đề tài cấp cơ sở có sản phẩm ứng dụng)",
     "(Nguồn: Nhóm nghiên cứu phân loại từ Bảng 2.4)"),
    ("(Nguồn: Nhóm nghiên cứu mô hình hóa từ kết quả nghiệm thu 38 đề tài",
     "(Nguồn: Nhóm nghiên cứu tổng hợp từ Bảng 2.4 và Bảng 2.6)"),
    ("Hình 2.9 mô hình hóa chuỗi chuyển hóa", T["222c"]),
    ("Hoạt động khai thác quyền sở hữu trí tuệ và thương mại hóa", T["223a"]),
    ("(Nguồn: Nhóm nghiên cứu tổng hợp từ các văn bản nội bộ của Trường Đại học Thành Đô và văn bản quy phạm",
     "(Nguồn: Nhóm nghiên cứu tổng hợp từ Quyết định số 213/QĐ-ĐHTĐ, Quyết định số 217/QĐ-ĐHTĐ, Quy chế chi tiêu nội "
     "bộ năm 2026, Điều lệ Quỹ Học bổng sau tiến sĩ Ngô Xuân Độ và Luật số 93/2025/QH15)"),
    ("(Nguồn: Nhóm nghiên cứu tổng hợp từ danh mục đề tài cấp cơ sở, danh mục đề tài cấp quốc gia và Quyết định số 213",
     "(Nguồn: Nhóm nghiên cứu tổng hợp từ danh mục đề tài cấp cơ sở, danh mục đề tài cấp quốc gia, Quyết định số "
     "213/QĐ-ĐHTĐ và Điều lệ Quỹ Học bổng sau tiến sĩ Ngô Xuân Độ ban hành kèm theo Nghị quyết số 40/NQ-HĐT-ĐHTĐ ngày "
     "27/5/2025)"),
    ("Bảng 2.3 tổng hợp các quy định về lợi ích của tác giả", T["223b"]),
    ("Điểm a Điều 36 quy định: tác giả được hưởng", None),
    ("Hình 2.2 mô phỏng tác động của hai cơ chế", None),
    ("Về tình hình khai thác thực tế: giai đoạn 2021 - 2025", T["223c"]),
    ("Công tác bảo vệ quyền sở hữu trí tuệ và bảo đảm liêm chính học thuật", T["224a"]),
    ("Tuy nhiên, trong công tác bảo vệ quyền sở hữu công nghiệp", T["224b"]),
    ("Đánh giá mức độ sẵn sàng dữ liệu theo Hình 2.6", T["224c"]),
    ("Hai là, tiềm lực nghiên cứu khoa học tăng trưởng vượt bậc", T["231b"]),
    ("Ba là, công tác xác lập quyền bước đầu đạt được", T["231c"]),
    ("Bốn là, thực hiện Kế hoạch số 07/KH-ĐHTĐ", T["231d"]),
    ("Thứ nhất, tồn tại sự đứt gãy nghiêm trọng", T["232a"]),
    ("Thứ hai, khung thể chế nội bộ còn tồn tại điểm nghẽn", T["232b"]),
    ("Thứ ba, hoạt động thương mại hóa và chuyển giao công nghệ chưa phát sinh", T["232c"]),
    ("Thứ tư, mô hình quản lý còn phân tán", T["232d"]),
    ("Các hạn chế nêu trên bắt nguồn từ cả nguyên nhân", T["233a"]),
    ("Về nguyên nhân khách quan: Thời gian thẩm định", T["233b"]),
    ("Về nguyên nhân chủ quan: Cơ chế động lực tài chính", T["233c"]),
    ("Hình 2.11. Chẩn đoán nguy cơ", "Hình 2.8. Nguyên nhân của các hạn chế trong quản lý quyền sở hữu trí tuệ"),
    ("(Nguồn: Nhóm nghiên cứu mô hình hóa cơ chế xung đột",
     "(Nguồn: Nhóm nghiên cứu tổng hợp từ kết quả phân tích tại Mục 2.2)"),
    ("Hình 2.11 làm rõ chuỗi tác động tiêu cực", None),
    ("Nghiên cứu thực trạng quản lý quyền sở hữu trí tuệ tại Trường Đại học Thành Đô giai đoạn", T["tk1"]),
    ("Sự đối kháng giữa chính sách ưu tiên thưởng tiền mặt", T["tk2"]),
]
# Đoạn ngoài Chương 2 (Mở đầu)
DOAN_NGOAI = [
    ("- Về sản phẩm khoa học công nghệ (582 bản ghi)",
     "- Về sản phẩm khoa học công nghệ (578 bản ghi): số liệu được tổng hợp từ các danh mục thống kê của Phòng Khoa học "
     "Công nghệ giai đoạn 2021 - 2025 (bài báo, tham luận, sách, giáo trình và đề tài), sau khi loại 4 bản ghi bài báo "
     "trùng lặp. Do các loại hình có đơn vị ghi nhận khác nhau, tổng này phản ánh quy mô sản lượng, không tương ứng hoàn "
     "toàn với số sản phẩm độc lập (một đề tài có thể đồng thời tạo ra báo cáo, bài báo và tham luận)."),
    ("Trường Đại học Thành Đô đang chuyển mình mạnh mẽ", T["md1"]),
    ("Tuy nhiên, thực tiễn quản trị sở hữu trí tuệ tại Nhà trường đang bộc lộ", T["md2"]),
]
# Đoạn chèn thêm sau đoạn (đầu đoạn gốc làm neo)
CHEN = [
    ("Đánh giá mức độ sẵn sàng dữ liệu theo Hình 2.6", [T["224d"]]),
]
# Sửa một phần trong đoạn: (đầu đoạn gốc, chuỗi cũ, chuỗi mới)
MOT_PHAN = [
    ("Quyết định số 679/QĐ-TTg ngày 19 tháng 5 năm 2009", "ngày 19 tháng 5 năm 2009", "ngày 27 tháng 5 năm 2009"),
    ("năng lực nghiên cứu khoa học tăng trưởng vượt bậc với 582 sản phẩm",
     "năng lực nghiên cứu khoa học tăng trưởng vượt bậc với 582 sản phẩm (bài báo tăng gấp 7 lần)",
     "năng lực nghiên cứu khoa học tăng nhanh với 578 sản phẩm (số bài báo năm 2025 gấp 7 lần năm 2021)"),
    # "hao hụt" -> tỷ lệ chuyển hóa còn thấp (Chương 3, Kết luận)
    ("W1. Tỷ lệ chuyển hóa đề tài sang đơn bảo hộ còn thấp (hao hụt 88,9%;",
     "W1. Tỷ lệ chuyển hóa đề tài sang đơn bảo hộ còn thấp (hao hụt 88,9%; 6 đề tài giai đoạn 2021 - 2024 chưa nộp đơn)",
     "W1. Tỷ lệ chuyển hóa từ kết quả nghiên cứu sang đơn đăng ký còn thấp (11,1%); 6 đề tài giai đoạn 2021 - 2024 "
     "chưa nộp đơn"),
    ("Xử lý Hạn chế 1 (hao hụt 88,9%)", "(hao hụt 88,9%)", "(tỷ lệ chuyển hóa sang đơn đăng ký còn thấp, 11,1%)"),
    ("các hạn chế về tỷ lệ hao hụt chuyển hóa", "các hạn chế về tỷ lệ hao hụt chuyển hóa từ đề tài sang đơn đăng ký (88,9%)",
     "các hạn chế về tỷ lệ chuyển hóa từ kết quả nghiên cứu sang đơn đăng ký còn thấp (11,1%)"),
    # SWOT thống nhất với Chương 2
    ("S2. Năng lực nghiên cứu tăng nhanh", "5/9 chỉ tiêu Kế hoạch 07 đạt hoặc vượt", "6/9 chỉ tiêu Kế hoạch 07 đạt hoặc vượt"),
    ("S2. Năng lực nghiên cứu tăng nhanh", "1 đơn sáng chế Quế hoa đang trong giai đoạn thẩm định hình thức",
     "1 đơn sáng chế Quế hoa đã nộp năm 2025"),
    # Chương 3: dẫn chiếu tới các hình đã bỏ
    ("Xử lý Hạn chế 2 (bất cập mức trần", "Bảng 2.3 và Hình 2.2 cho thấy", "Bảng 2.7 và Mục 2.2.3 cho thấy"),
    ("Xử lý Hạn chế 3 (thiếu liên thông)", "Hình 2.6 có 7/16 tiêu chí khuyết trắng dữ liệu.",
     "Mục 2.2.4 cho thấy dữ liệu về doanh thu, thù lao, chi phí xác lập và duy trì quyền chưa được ghi nhận."),
]

# Ô bảng: (ô tiêu đề nhận diện bảng, hàng, cột, giá trị cũ, giá trị mới)
O = []
# Bảng 2.2: trả về số liệu đã đối chiếu với các danh mục thống kê gốc (cộng đúng 56, 100, 85, 128, 192; 582)
for hang, cu, moi in [
    (1, ["17", "62", "37", "72", "109", "297"], ["17", "42", "49", "71", "107", "286"]),
    (10, ["56", "100", "85", "128", "192", "582"], ["56", "100", "84", "127", "190", "578"]),
    (2, ["6", "8", "16", "33", "52", "115"], ["6", "13", "11", "33", "52", "115"]),
    (3, ["0", "0", "3", "29", "39", "71"], ["0", "2", "1", "28", "40", "71"]),
    (4, ["26", "20", "17", "13", "11", "87"], ["26", "31", "12", "7", "11", "87"]),
    (5, ["1", "0", "3", "4", "4", "12"], ["0", "2", "2", "2", "6", "12"]),
    (9, ["0", "0", "5", "1", "7", "13"], ["1", "2", "3", "5", "5", "16"]),
]:
    for c, (a, b) in enumerate(zip(cu, moi), start=1):
        if a != b:
            O.append(("Loại hình sản phẩm khoa học", hang, c, a, b))
O += [
    ("Kênh tài trợ", 2, 2, "Khoản chi phí khác tối đa bằng 50% chi trực tiếp",
     "Chi phí khác (xuất bản, đi lại, thiết bị, vật liệu) tối đa 50% chi trả trực tiếp cho nhà khoa học; chưa nêu riêng "
     "chi phí đăng ký bảo hộ"),
    ("Kênh tài trợ", 2, 3, "Quỹ học bổng sau tiến sĩ mang tính chất thí điểm hợp tác nghiên cứu đặc thù từ năm 2025",
     "Thành lập năm 2025; Viện Nghiên cứu giáo dục và Chuyển giao tri thức được ủy quyền điều hành; Trường là đồng chủ "
     "sở hữu văn bằng bảo hộ của sản phẩm do Quỹ tài trợ"),
    ("Số TT", 5, 5, "Đã nộp đơn, đang thẩm định", "Chờ cấp bằng"),
    *[("Số TT", r, 5, "Chưa xác minh trạng thái", "Đã cấp bằng") for r in range(6, 11)],
    *[("Số TT", r, 3, f"Đơn 3-2023-0284{r - 5}; Số {so}", f"Đơn 3-2023-0284{r - 5}; Bằng số {so}")
      for r, so in zip(range(6, 11), ["3-0039693-000", "3-0039694-000", "3-0039902-000", "3-0039695-000",
                                       "3-0039696-000"])],
    ("Số TT", 11, 5, "Đã nộp đơn, đang thẩm định hình thức", "Đã nộp đơn năm 2025"),
    ("Loại hình sản phẩm khoa học", 5, 0, "Sách chuyên khảo, tham khảo (ISBN)", "Sách, chương sách, tài liệu có ISBN"),
    ("Nhóm quyền", 1, 3, "405 bài báo, 12 cuốn sách", "401 bài báo; 12 sách, chương sách"),
    ("Nhóm quyền", 7, 4, "5 hồ sơ có số hiệu văn bằng, trạng thái chờ xác minh",
     "Đã cấp 5 bằng độc quyền năm 2024, đồng sở hữu với doanh nghiệp"),
    ("Nhóm quyền", 8, 4, "Đã cấp 2 văn bằng bảo hộ; 01 đơn đang xử lý", "Đã cấp 2 văn bằng bảo hộ; 01 đơn chờ cấp bằng"),
]


# ---------------------------------------------------------------------------
# Thực hiện
# ---------------------------------------------------------------------------
def van_ban(p):
    return "".join(t.text or "" for r in p.findall(q("r")) for t in r.findall(q("t")))


# Điểm mở rộng cho các chương khác (scripts/ra_soat/sua_chuong3_ban_cuoi.py)
BANG_THAY = [BANG27]
THAY_TAT_CA = []
DOAN_CHUA = []  # (chuỗi nằm trong đúng một đoạn, văn bản mới hoặc None để xóa)
BANG_XOA = []  # ô đầu của các bảng cần xóa cả bảng


def chen_bang(ed, mau, cau_hinh):
    """Chèn bảng mới (theo dõi thay đổi) ngay sau bảng mẫu, dùng lại định dạng ô của bảng mẫu."""
    hang = mau.findall(q("tr"))
    tc_dau, tc_than = hang[0].find(q("tc")), hang[1].find(q("tc"))
    moi = etree.Element(q("tbl"))
    moi.append(copy.deepcopy(mau.find(q("tblPr"))))
    grid = etree.SubElement(moi, q("tblGrid"))
    for w in cau_hinh["rong"]:
        etree.SubElement(grid, q("gridCol")).set(q("w"), str(w))

    def o(tc_mau, text, w):
        tc = etree.Element(q("tc"))
        tcpr = copy.deepcopy(tc_mau.find(q("tcPr")))
        tcpr.find(q("tcW")).set(q("w"), str(w))
        tc.append(tcpr)
        p_mau = tc_mau.find(q("p"))
        p = etree.SubElement(tc, q("p"))
        p.append(copy.deepcopy(p_mau.find(q("pPr"))))
        ed._danh_dau_doan(p, "ins")
        r_mau = next(r for r in p_mau.findall(q("r")) if r.find(q("t")) is not None)
        ins = ed._dau("ins")
        p.append(ins)
        r = etree.SubElement(ins, q("r"))
        r.append(copy.deepcopy(r_mau.find(q("rPr"))))
        for k, dong in enumerate(text.split("\n")):
            if k:
                etree.SubElement(r, q("br"))
            t = etree.SubElement(r, q("t"))
            t.text = dong
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        return tc

    for k, dong in enumerate([cau_hinh["cot"]] + cau_hinh["dong"]):
        tr = etree.SubElement(moi, q("tr"))
        trpr = etree.SubElement(tr, q("trPr"))
        if k == 0:
            etree.SubElement(trpr, q("tblHeader"))
        trpr.append(ed._dau("ins"))
        for text, w in zip(dong, cau_hinh["rong"]):
            tr.append(o(tc_dau if k == 0 else tc_than, text, w))
    mau.addnext(moi)
    return moi


def main():
    z = zipfile.ZipFile(VAO)
    items = {i.filename: z.read(i.filename) for i in z.infolist()}
    infos = list(z.infolist())
    root = etree.fromstring(items["word/document.xml"])
    rels = etree.fromstring(items["word/_rels/document.xml.rels"])
    ctypes = etree.fromstring(items["[Content_Types].xml"])
    ed = TrackEditor(root, "Claude")
    body = root.find(q("body"))

    # Vùng Chương 2 (bỏ qua mục lục)
    con = list(body)
    tieu_de = [i for i, el in enumerate(con) if el.tag == q("p") and van_ban(el).startswith("CHƯƠNG 2. THỰC TRẠNG")]
    i2 = tieu_de[-1]
    i3 = next(i for i, el in enumerate(con) if i > i2 and el.tag == q("p")
              and van_ban(el).startswith("CHƯƠNG 3. HỆ THỐNG"))
    vung2 = con[i2:i3]
    doan2 = [el for el in vung2 if el.tag == q("p")]
    tat_ca_doan = list(root.iter(q("p")))

    def tim(dau, ds=None):
        hits = [p for p in (ds or doan2) if van_ban(p).startswith(dau)]
        if len(hits) != 1:
            raise ValueError(f"{len(hits)} đoạn bắt đầu bằng: {dau}")
        return hits[0]

    def bang(o_dau):
        hits = [t for t in body.iter(q("tbl")) if van_ban(t.find(".//" + q("p"))) == o_dau]
        if len(hits) != 1:
            raise ValueError(f"{len(hits)} bảng có ô đầu: {o_dau}")
        return hits[0]

    # 1. Giải quyết trước mọi phần tử đích (trước khi văn bản bị thay)
    dich_doan = [(tim(d), m) for d, m in DOAN]
    dich_chen = [(tim(d), ds) for d, ds in CHEN]
    def tim_chua(neo):
        hits = [p for p in tat_ca_doan if neo in van_ban(p)]
        if len(hits) != 1:
            raise ValueError(f"{len(hits)} đoạn chứa: {neo}")
        return hits[0]

    dich_mot_phan = [(tim_chua(d), c, m) for d, c, m in MOT_PHAN]
    dich_ngoai = [(tim(d, tat_ca_doan), m) for d, m in DOAN_NGOAI] + [(tim_chua(a), m) for a, m in DOAN_CHUA]
    bang_cu = [(bang(b["dau"]), b) for b in BANG_THAY]
    bang_xoa = [bang(d) for d in BANG_XOA]
    dich_o = []
    for o_dau, r, c, cu, moi in O:
        tc = bang(o_dau).findall(q("tr"))[r].findall(q("tc"))[c]
        p = next(p for p in tc.iter(q("p")) if van_ban(p))
        if van_ban(p) != cu:
            raise ValueError(f"Ô {o_dau}[{r},{c}] = {van_ban(p)!r}, cần {cu!r}")
        dich_o.append((p, cu, moi))

    def khoi(dau_chu_thich):
        """(đoạn chú thích, các phần tử hình vẽ bằng ký tự, đoạn nguồn)."""
        cap = tim(dau_chu_thich)
        k = vung2.index(cap)
        giua = []
        for el in vung2[k + 1:]:
            if el.tag == q("p") and van_ban(el).startswith("(Nguồn"):
                return cap, giua, el
            giua.append(el)
        raise ValueError(dau_chu_thich)

    khoi_bd = {dau: khoi(dau) for dau in BIEU_DO}
    khoi_sd = {dau: khoi(dau) for dau in SO_DO}
    khoi_bo = [khoi(dau) for dau in BO_HINH]

    # 2. Hình: xóa hình vẽ bằng ký tự, chèn biểu đồ nhúng và sơ đồ hình vẽ
    def xoa(el):
        (ed.xoa_bang if el.tag == q("tbl") else ed.xoa_doan)(el)

    def giu_cung_trang(cap):
        """Chú thích hình luôn nằm cùng trang với hình ngay dưới (định dạng, không thay đổi nội dung)."""
        ppr = cap.find(q("pPr"))
        if ppr is None:
            ppr = etree.Element(q("pPr"))
            cap.insert(0, ppr)
        if ppr.find(q("keepNext")) is None:
            k = etree.Element(q("keepNext"))
            ps = ppr.find(q("pStyle"))
            (ps.addnext(k) if ps is not None else ppr.insert(0, k))

    ct_ns = f"{{{CT}}}"
    if not any(d.get("Extension") == "xlsx" for d in ctypes.findall(ct_ns + "Default")):
        etree.SubElement(ctypes, ct_ns + "Default", Extension="xlsx", ContentType=CT_XLSX)
    for n, (dau, h) in enumerate(BIEU_DO.items(), start=1):
        cap, giua, _ = khoi_bd[dau]
        for el in giua:
            xoa(el)
        items[f"word/embeddings/Microsoft_Excel_Worksheet{n}.xlsx"] = xlsx_bytes(h, False)
        items[f"word/charts/chart{n}.xml"] = chart_xml(h, "rId1")
        items[f"word/charts/_rels/chart{n}.xml.rels"] = (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{PR}">'
            f'<Relationship Id="rId1" Type="{RT_PACKAGE}" Target="../embeddings/Microsoft_Excel_Worksheet{n}.xlsx"/>'
            f'</Relationships>').encode("utf-8")
        etree.SubElement(ctypes, ct_ns + "Override", PartName=f"/word/charts/chart{n}.xml", ContentType=CT_CHART)
        rid = f"rIdBieuDo{n}"
        etree.SubElement(rels, f"{{{PR}}}Relationship", Id=rid, Type=RT_CHART, Target=f"charts/chart{n}.xml")
        run = INLINE.format(cx=int(15.5 * EMU_CM), cy=int(cao_cm(h) * EMU_CM), id=900 + n, rid=rid,
                            descr=f"Hình 2.{h['so']}. {h['tieu_de']}")
        giu_cung_trang(cap)
        ed.chen_doan_sau(cap, [run], PPR_HINH)
    for k, (dau, (so, ham)) in enumerate(SO_DO.items(), start=1):
        cap, giua, _ = khoi_sd[dau]
        for el in giua:
            xoa(el)
        giu_cung_trang(cap)
        ed.chen_doan_sau(cap, [ham().run_xml(950 + k, f"{so}. Sơ đồ")], PPR_HINH)
    for cap, giua, nguon in khoi_bo:
        for el in [cap, *giua, nguon]:
            xoa(el)

    # 3. Văn bản
    for p, cu, moi in dich_o:
        ed._replace_in(p, cu, moi)
    for p, ds in dich_chen:
        el = p
        for text in ds:
            el = ed.them_doan_sau(el, text)
    for p, moi in dich_doan:
        if moi is None:
            ed.xoa_doan(p)
        else:
            ed._replace_in(p, van_ban(p), moi)
    for p, cu, moi in dich_mot_phan:
        ed._replace_in(p, cu, moi)
    for p, moi in dich_ngoai:
        if moi is None:
            ed.xoa_doan(p)
        else:
            ed._replace_in(p, van_ban(p), moi)
    # Bảng 2.7: xóa bảng cũ, chèn bảng rút gọn ngay sau
    for cu, b in bang_cu:
        chen_bang(ed, cu, b)
        ed.xoa_bang(cu)
    for t in bang_xoa:
        ed.xoa_bang(t)
    for cu, moi in THAY_TAT_CA:
        ed.replace(cu, moi, tat_ca=True)

    # 4. Đánh số lại hình, bảng từ Chương 2 đến hết văn bản
    pham_vi = [p for el in con[i2:] for p in ([el] if el.tag == q("p") else el.iter(q("p")))]
    so_lan = 0
    for cu in sorted(DOI_SO, key=len, reverse=True):
        mau = re.compile(re.escape(cu) + r"(?!\d)")
        for p in pham_vi:
            while mau.search(van_ban(p)):
                ed._replace_in(p, cu, DOI_SO[cu])
                so_lan += 1
    con_sot = [van_ban(p) for p in pham_vi if re.search(r"Hình 2\.(2|6|10)(?!\d)", van_ban(p))]
    if con_sot:
        print("CẢNH BÁO, còn dẫn chiếu hình đã bỏ:", con_sot)

    items["word/document.xml"] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    items["word/_rels/document.xml.rels"] = etree.tostring(rels, xml_declaration=True, encoding="UTF-8",
                                                           standalone=True)
    items["[Content_Types].xml"] = etree.tostring(ctypes, xml_declaration=True, encoding="UTF-8", standalone=True)
    ten_cu = [i.filename for i in infos]
    with zipfile.ZipFile(RA, "w", zipfile.ZIP_DEFLATED) as zout:
        for ten in ten_cu + [t for t in items if t not in ten_cu]:
            zout.writestr(ten, items[ten])
    print(f"{len(DOAN)} đoạn, {len(O)} ô bảng, {len(BIEU_DO)} biểu đồ, {len(SO_DO)} sơ đồ, "
          f"{so_lan} dẫn chiếu đánh số lại -> {RA}")


if __name__ == "__main__":
    main()

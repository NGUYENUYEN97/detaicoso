# -*- coding: utf-8 -*-
"""Xuất bản chỉnh sửa theo bản góp ý ba chương vào thư mục riêng, giữ nguyên các tệp gốc trong Ban_cuoi.

    python3 scripts/ban_cuoi/ban_chinh_sua.py

Đầu ra trong Ban_cuoi/Ban_chinh_sua_02-10-2026/:
  - Chuong_1_Co_so_ly_luan_va_phap_ly.docx, Chuong_2_Thuc_trang_quan_ly_quyen_SHTT.docx,
    Chuong_3_He_thong_giai_phap.docx: ba chương đã sửa đồng bộ;
  - Bao_cao_tong_ket_de_tai.docx: bản toàn văn báo cáo đã sửa;
  - Du_lieu_bieu_do_Chuong_2.xlsx: dữ liệu và biểu đồ Chương 2, có sheet nhật ký chỉnh sửa;
  - Bang_giai_trinh_chinh_sua.docx: vị trí, nội dung cũ, nội dung sửa, căn cứ. Phần lời văn được sinh tự động bằng
    cách so khớp từng đoạn của báo cáo cũ (Ban_cuoi/Bao_cao_tong_ket_de_tai.docx) với báo cáo mới; bảng được so theo
    tên bảng; hình và sơ đồ ghi theo danh sách GIAI_TRINH_HINH;
  - Danh_sach_diem_chua_du_minh_chung.docx.
"""
import difflib
import os
import re
import sys

import docx
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from lxml import etree

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import khung  # noqa: E402
import chuong1  # noqa: E402
import chuong3  # noqa: E402
import bao_cao  # noqa: E402

BG = khung.BG
B, BD = BG.B, BG.BD
THU_MUC = os.path.join(khung.GOC, "Ban_cuoi", "Ban_chinh_sua_02-10-2026")
BAN_CU = os.path.join(khung.GOC, "Ban_cuoi", "Bao_cao_tong_ket_de_tai.docx")


# ---------------------------------------------------------------------------
# 1. Dựng ba chương, workbook và báo cáo toàn văn
# ---------------------------------------------------------------------------
def dung_chuong():
    khung.THU_MUC_RA = THU_MUC
    os.makedirs(THU_MUC, exist_ok=True)
    chuong1.dung()
    v = khung.VanBanChung()
    BG.noi_dung(v)
    h_kh = v.ket_qua["h_kh"]
    for p in v.doc.paragraphs:
        if re.match(r"^[ab]\) Nguyên nhân", p.text):
            p.paragraph_format.keep_with_next = True
    khung.luu(v, "Chuong_2_Thuc_trang_quan_ly_quyen_SHTT.docx",
              "Chương 2. Thực trạng quản lý quyền sở hữu trí tuệ tại Trường Đại học Thành Đô", 2)
    B.HINH[:] = v.hinh
    BD.RA_XLSX = os.path.join(THU_MUC, "Du_lieu_bieu_do_Chuong_2.xlsx")
    if BG.NHAT_KY_LAN6[0] not in BD.NHAT_KY:
        BD.NHAT_KY[:0] = BG.NHAT_KY_LAN6 + BG.NHAT_KY_GON
    BD.dung_workbook()
    BD.va_workbook(BD.RA_XLSX)
    print("Đã ghi:", BD.RA_XLSX)
    v3 = khung.VanBanChung()
    chuong3.noi_dung(v3, so_hinh_kh=h_kh)
    khung.luu(v3, "Chuong_3_He_thong_giai_phap.docx",
              "Chương 3. Hệ thống giải pháp nâng cao hiệu quả quản lý quyền sở hữu trí tuệ tại Trường Đại học Thành Đô", 3)


def dung_bao_cao():
    bao_cao.RA_DOCX = os.path.join(THU_MUC, "Bao_cao_tong_ket_de_tai.docx")
    bao_cao.RA_XLSX = os.path.join(THU_MUC, "Du_lieu_bieu_do_Chuong_2.xlsx")
    bao_cao.ghi_workbook = lambda v: None  # dữ liệu biểu đồ đã ghi ở Du_lieu_bieu_do_Chuong_2.xlsx
    bao_cao.main()
    return bao_cao.RA_DOCX


# ---------------------------------------------------------------------------
# 2. Bảng giải trình
# ---------------------------------------------------------------------------
TIEU_DE = re.compile(r"^(CHƯƠNG \d|MỞ ĐẦU|KẾT LUẬN VÀ KIẾN NGHỊ|TIỂU KẾT CHƯƠNG \d|\d\.\d\.\d\. |\d\.\d\. |\d\. [A-ZĐ]|"
                     r"[ab]\) Nguyên nhân)")

CAN_CU = [
    (r"Thuyết minh|thuyết minh đề tài gọi", "Góp ý mục 8: không coi tệp có tên FINAL là thuyết minh đã được phê duyệt"),
    (r"Quy chế chi tiêu|sau kỳ đánh giá|giai đoạn tới|ban hành ngày 01 tháng 8", "Góp ý mục 7: không dùng quy định năm "
                                                                                  "2026 để giải thích kết quả 2021 - 2025"),
    (r"độ trễ thể chế|không giải thích", "Góp ý mục 7: giảm khẳng định nhân quả; độ trễ thể chế chỉ giải thích nhu cầu "
                                         "cập nhật quy chế"),
    (r"nguồn lực không thiếu|dữ liệu không đủ|chưa đủ để xác định|dễ dừng lại|nguyên nhân hàng đầu|có thể phụ thuộc",
     "Góp ý mục 7: giảm mức khẳng định nhân quả"),
    (r"thưởng|thù lao|nhuận bút|Điều 28|Điều 135|Điều 73|100 triệu|phân chia|lợi ích của tác giả",
     "Góp ý mục 2: Điều 28, Điều 73 Luật số 93/2025/QH15; khoản 1 Điều 135 Luật Sở hữu trí tuệ, Văn bản hợp nhất số "
     "67/VBHN-VPQH; khoản 4 Điều 36 Quyết định 213; Điều 13 Quyết định 217"),
    (r"Kế hoạch 07|1\.11|chỉ tiêu", "Góp ý mục 6: Kế hoạch số 07/KH-ĐHTĐ, mục 2.2.4, trang 7 - 8"),
    (r"kiểu dáng|văn bằng|giấy chứng nhận|1-2026|1-2025|hồ sơ tài sản|trạng thái",
     "Góp ý mục 5: Tai lieu thanh do/Theo dõi đơn nhãn hiệu và KDCN.xlsx, Sheet1 và Sheet2"),
    (r"582|475|107 |Gini|0,83|tên khác nhau|bản ghi|so khớp",
     "Góp ý mục 7: scripts/ra_soat/so_khop_tac_gia.py; các danh mục thống kê của Phòng Khoa học Công nghệ"),
    (r"xung đột|thứ tự áp dụng|cách hiểu", "Góp ý mục 4: chỉ kết luận xung đột khi cùng một tình huống chịu các yêu cầu "
                                          "không tương thích"),
    (r"liêm chính|Điều 14", "Góp ý mục 4: Điều 14 Quyết định 217 đã có quy định về hành vi xâm phạm quyền tác giả"),
    (r"Khâu \d|phân nhánh|bảo mật trước|sàng lọc", "Góp ý mục 8: quy trình có nhận diện, tra cứu, xem xét bảo mật trước "
                                                   "khi công bố, phân nhánh theo loại đối tượng"),
    (r"biểu mẫu|Mẫu \d|Điều 35|Điều 38|dự toán|tạm ứng|lệ phí|cơ chế chi",
     "Góp ý mục 3: Điều 35, Điều 38 và các Mẫu 01, 06, 11, 15, 16 Quyết định 213; Điều 11 Quyết định 217"),
    (r"Điều 10|Điều 11|giả thuyết|kiểm chứng|chủ động của", "Góp ý mục 3: Điều 10, Điều 11 Quyết định 217; không kết "
                                                            "luận khi chưa có bằng chứng"),
    (r"Quy chế chi tiêu|sau kỳ đánh giá|giai đoạn tới", "Góp ý mục 7: không dùng quy định năm 2026 để giải thích kết quả "
                                                         "2021 - 2025"),
    (r"tiềm năng|mã số 2021 - 2024|chuỗi|nghiệm thu ngày", "Góp ý mục 7: danh mục đề tài cấp cơ sở 2021 - 2025, cột sản "
                                                           "phẩm và ngày nghiệm thu"),
    (r"độ trễ thể chế", "Góp ý mục 7: giảm khẳng định nhân quả; độ trễ thể chế chỉ giải thích nhu cầu cập nhật quy chế"),
    (r"thí điểm", "Góp ý mục 8: thí điểm ghi là kế hoạch"),
    (r"nhân sự|chuyên trách|kiêm nhiệm|Trung tâm tư vấn|doanh nghiệp quản lý|chuyển bước",
     "Góp ý mục 8: căn cứ nhu cầu, phương án theo giai đoạn và điều kiện chuyển bước"),
    (r"chỉ số|Bảng 3\.4", "Góp ý mục 8: chỉ số phân biệt sàng lọc, đơn nộp, văn bằng, khai thác"),
    (r"Mạng lưới|Quỹ Học bổng|thông tin được cung cấp|thống kê do Viện",
     "Góp ý mục 1: phân biệt điểm đã có văn bản với điểm chưa xác minh"),
    (r"nguồn lực không thiếu|nguyên nhân|dữ liệu không đủ", "Góp ý mục 7: giảm mức khẳng định nhân quả"),
]

GIAI_TRINH_HINH = [
    ("Chương 2, Hình 2.2",
     "Mô phỏng 5 đường: điểm a, điểm b Điều 36 Quyết định 213; Quỹ Ngô Xuân Độ năm đầu và từ năm thứ hai; mức mặc định "
     "15% Điều 135 Luật Sở hữu trí tuệ, trên cùng trục khoản thu sau chi phí",
     "Chỉ giữ 2 đường có cùng cơ sở tính là nguồn thu sau chi phí cần thiết, hợp lệ: điểm a, khen thưởng tập thể tác "
     "giả, đề tài sử dụng ngân sách nhà nước, tối đa 100 triệu đồng; điểm b, tác giả, tài sản thuộc sở hữu của Trường. "
     "Nguồn ghi rõ đối tượng hưởng, phạm vi, cơ sở tính, khoản khấu trừ, điều kiện; các quy định khác cơ sở tính chuyển "
     "sang Bảng 2.3",
     "Góp ý mục 2: không đặt tỷ lệ tính trên khoản thu sau chi phí cạnh tỷ lệ tính trên tổng tiền trước thuế"),
    ("Chương 2, Hình 2.6",
     "16 tiêu chí: 5 tính được đầy đủ, 4 một phần, 7 chưa tính được",
     "4 đầy đủ, 5 một phần, 7 chưa: tiêu chí kinh phí chuyển sang một phần do Điều 35, 38 Quyết định 213 đã có căn cứ "
     "chi; tiêu chí văn bằng chuyển sang một phần do trạng thái 5 kiểu dáng chưa thống nhất; tiêu chí cơ sở dữ liệu tra "
     "cứu chuyển sang chưa tính được do chưa có số liệu",
     "Góp ý mục 3 và mục 5"),
    ("Chương 2, Hình 2.7",
     "Tài sản trí tuệ theo loại hình: đã cấp văn bằng 9, đang xử lý 3, gồm đơn sáng chế năm 2026",
     "Hồ sơ trong kỳ 2021 - 2025 theo ba nhóm trạng thái: đã cấp văn bằng, giấy chứng nhận 4; có số hiệu văn bằng, "
     "trạng thái chưa xác minh 5; đã nộp đơn 2. Không gồm đơn năm 2026",
     "Góp ý mục 5: Theo dõi đơn nhãn hiệu và KDCN.xlsx"),
    ("Chương 2, Hình 2.8",
     "Sản phẩm đề tài cấp cơ sở đủ điều kiện xác lập quyền theo nhóm quyền có thể xác lập",
     "Sản phẩm đề tài cấp cơ sở có tiềm năng tạo lập tài sản trí tuệ theo nhóm quyền dự kiến; số liệu không đổi",
     "Góp ý mục 7: 11 đề tài là có tiềm năng tạo lập tài sản trí tuệ, chưa được thẩm định"),
    ("Chương 2, Hình 2.9",
     "Phễu 38 đề tài, 11 có sản phẩm đủ điều kiện, 1 đơn; đề tài giao đến 2024: 31, 8, 0",
     "Phễu sở hữu công nghiệp: 38 đề tài, 9 có sản phẩm tiềm năng sở hữu công nghiệp, 1 đơn; đề tài mã số 2021 - 2024, "
     "đều nghiệm thu trước 31/5/2025: 31, 6, 0. Hai sản phẩm quyền tác giả được tách riêng",
     "Góp ý mục 7: phễu sở hữu công nghiệp không gộp quyền tác giả; Bảng 2.7 bổ sung ngày nghiệm thu"),
    ("Chương 2, Hình 2.10",
     "10 chỉ tiêu, 7 đạt; chỉ tiêu văn bằng 6 trên 4, tức 150%, tính cả nhãn hiệu Thado Edupark; tên chỉ tiêu viết tắt",
     "9 chỉ tiêu xác định được, 6 đạt; mục 1.11 đưa ra khỏi biểu đồ, ghi chưa xác định, nếu 5 kiểu dáng được xác nhận "
     "cấp năm 2024 thì 5 trên 4, tức 125%; tên chỉ tiêu ghi theo nguyên văn kèm số mục; sheet DL_Ke_hoach_07 ghi đủ 10 chỉ "
     "tiêu",
     "Góp ý mục 6: Kế hoạch 07/KH-ĐHTĐ trang 7, mục 1.11 không gồm nhãn hiệu"),
    ("Chương 2, Hình 2.11",
     "Chuỗi nguyên nhân: trách nhiệm dồn về tác giả; chưa có dòng kinh phí nộp đơn; điểm nghẽn cốt lõi",
     "Giả thuyết về chuỗi nguyên nhân: việc khởi động thủ tục phụ thuộc vào sự chủ động của chủ nhiệm đề tài; chưa có dự "
     "toán riêng và người chịu trách nhiệm đề xuất chi; kết quả có thể được công bố trước khi sàng lọc; điểm nghẽn giả định",
     "Góp ý mục 3: Điều 10, Điều 11 Quyết định 217; nhận định chỉ giữ ở dạng giả thuyết khi chưa có bằng chứng"),
    ("Chương 3, Hình 3.1",
     "Phòng Tài chính - Kế toán: kinh phí nộp đơn, chi trả thù lao, giám sát quỹ",
     "Phòng Tài chính - Kế toán: dự toán phí xác lập quyền, chi trả thưởng, thù lao",
     "Góp ý mục 2 và mục 3: phân biệt thưởng, thù lao; chuyển trọng tâm sang dự toán"),
    ("Chương 3, Hình 3.2",
     "Quy trình 8 khâu từ ý tưởng đến thương mại hóa, khâu 4 rà soát tại nghiệm thu là mắt xích quyết định",
     "Quy trình 8 khâu từ khai báo đến khai thác: khai báo; sàng lọc, phân nhánh; tra cứu; xem xét bảo mật trước khi công "
     "bố hoặc trình diễn; xác nhận tại nghiệm thu; quyết định xác lập quyền và dự toán; nộp đơn hoặc bảo mật, theo dõi; "
     "khai thác. Bổ sung 5 nhánh: sáng chế, giải pháp hữu ích; kiểu dáng; nhãn hiệu; quyền tác giả; bí mật kinh doanh",
     "Góp ý mục 8: bước nhận diện, tra cứu, bảo mật trước công bố; phân nhánh theo đối tượng"),
    ("Chương 3, Hình 3.3",
     "Lộ trình: quy chế hợp nhất, nhân sự chuyên trách, dòng kinh phí; doanh nghiệp quản lý tài sản trí tuệ vận hành từ "
     "2029",
     "Lộ trình: hoàn thiện quy chế; phiếu khai báo, phiếu rà soát; đầu mối kiêm nhiệm và dự toán phí; xem xét chuyên trách "
     "khi đạt ngưỡng; trung tâm, doanh nghiệp chỉ khi đủ điều kiện",
     "Góp ý mục 8: phương án theo giai đoạn, điều kiện chuyển bước"),
]


def _chuan(t):
    return re.sub(r"\s+", " ", t).strip()


def doan_noi_dung(path):
    d = docx.Document(path)
    ps = d.paragraphs[bao_cao.dau_noi_dung(d):]
    kq, ngu_canh, dem, trong_chuong = [], "", 0, False
    for p in ps:
        t = _chuan(p.text)
        if not t:
            continue
        if t.startswith("TÀI LIỆU THAM KHẢO"):
            break
        if t.startswith("CHƯƠNG "):
            trong_chuong = True
        elif t.startswith("KẾT LUẬN VÀ KIẾN NGHỊ"):
            trong_chuong = False
        if TIEU_DE.match(t) and not (trong_chuong and re.match(r"^\d\. [A-ZĐ]", t)):
            ngu_canh, dem = t.split(" CƠ SỞ")[0].split(" THỰC TRẠNG")[0].split(" HỆ THỐNG")[0][:60], 0
            kq.append((t, f"{ngu_canh}, tên mục"))
            continue
        dem += 1
        kq.append((t, f"{ngu_canh}, đoạn {dem}"))
    return d, kq


def chuong_cua(ngu_canh, so_chuong_hien):
    return so_chuong_hien


def vi_tri_day_du(ds_vt):
    """Thêm tên chương vào vị trí để dễ tra."""
    kq, chuong = [], "Mở đầu"
    for t, vt in ds_vt:
        m = re.match(r"^CHƯƠNG (\d)", t)
        if m:
            chuong = f"Chương {m.group(1)}"
        elif t.startswith("KẾT LUẬN VÀ KIẾN NGHỊ"):
            chuong = "Kết luận và kiến nghị"
        kq.append((t, f"{chuong}; {vt}"))
    return kq


def phan_thay_doi(cu, moi):
    """Các từ bị xóa hoặc thêm giữa hai đoạn, dùng để xác định căn cứ theo đúng nội dung được sửa."""
    a, b = cu.split(), moi.split()
    kq = []
    for op, i1, i2, j1, j2 in difflib.SequenceMatcher(None, a, b, autojunk=False).get_opcodes():
        if op != "equal":
            kq += a[max(0, i1 - 3):i2 + 3] + b[max(0, j1 - 3):j2 + 3]
    return " ".join(kq)


def _can_cu_tu_khoa(cu, moi):
    goc = phan_thay_doi(cu, moi) if cu and moi else moi + " " + cu
    for mau, cc in CAN_CU:
        if re.search(mau, goc):
            return cc
    return "Góp ý mục 1 và yêu cầu kiểm tra cuối: thống nhất thuật ngữ, số liệu, thời kỳ giữa ba chương"


G = {
    1: "Góp ý mục 1: phân biệt quy định, thực hiện, thống kê, nhận định; không suy chưa thấy thành không có",
    2: "Góp ý mục 2: Điều 28, Điều 73 Luật số 93/2025/QH15; khoản 1 Điều 135 Luật Sở hữu trí tuệ; phân biệt thưởng, thù "
       "lao, nhuận bút, phân chia lợi nhuận",
    3: "Góp ý mục 3: Điều 35, Điều 38 và các biểu mẫu Quyết định 213; Điều 10, Điều 11 Quyết định 217",
    4: "Góp ý mục 4: chỉ kết luận xung đột khi cùng một tình huống chịu yêu cầu không tương thích; Điều 14 Quyết định 217",
    5: "Góp ý mục 5: Theo dõi đơn nhãn hiệu và KDCN.xlsx, Sheet1 và Sheet2; quyền tác giả chưa đăng ký không có nghĩa là "
       "chưa có quyền",
    6: "Góp ý mục 6: Kế hoạch số 07/KH-ĐHTĐ, mục 2.2.4, chỉ tiêu 1.11 không gồm nhãn hiệu",
    7: "Góp ý mục 7: thống kê, phễu sở hữu công nghiệp, năm theo loại dữ liệu, giảm khẳng định nhân quả",
    8: "Góp ý mục 8: giải pháp nối với hạn chế; quy trình có sàng lọc, bảo mật trước công bố; phương án theo giai đoạn; "
       "thí điểm là kế hoạch; chỉ số phân tầng",
}


def _g(*ds):
    return "; ".join(G[d] for d in ds)


CAN_CU_RIENG = [
    ("Khâu ", _g(8)), ("Bảng 2.3.", _g(2, 4)), ("Bảng 2.4.", _g(3, 1)), ("Bảng 2.5.", _g(5)), ("Bảng 2.6.", _g(5)),
    ("Bảng 2.7.", _g(7)), ("Bảng 3.1.", _g(1, 5, 6, 7)), ("Bảng 3.2.", _g(2, 4, 8)), ("Bảng 3.3.", _g(8)),
    ("Bảng 3.4.", _g(8, 6)), ("3.2.6.", _g(8)), ("3.2.4. Giải pháp 4", _g(3, 8)), ("2.3.2.", _g(5)), ("Hình 2.7.", _g(5)),
    ("Để đo lường hiệu quả", _g(8, 6)), ("Bộ chỉ số nêu trên", _g(8)), ("Nguồn: Nhóm nghiên cứu đề xuất. Mục tiêu", _g(8)),
    ("Hình 2.11 trình bày", _g(3, 7)), ("Về dữ liệu", _g(7)), ("Thứ năm, hệ thống dữ liệu", _g(5, 7)),
    ("Chương 2 phân tích thực trạng", _g(7)), ("Thứ tư, Luật sửa đổi", _g(8)), ("Nguyên tắc 3", _g(8)),
    ("Bố trí nhân sự theo giai đoạn", _g(8)), ("Giai đoạn 1 không phát sinh", _g(8)), ("Nội dung 1: Lập dòng", _g(3)),
    ("Nội dung 2: Đề xuất sửa đổi", _g(2, 7)), ("Nội dung 3: Chuẩn bị phương án", _g(8)),
    ("Chương 2 không sử dụng khảo sát", _g(5, 7)), ("Nội dung 4:", _g(5)), ("Ban hành Quy trình chuẩn", _g(8)),
    ("- Triển khai kế hoạch thí điểm", _g(8)), ("- Theo dõi nguồn thu", _g(2, 8)), ("Giai đoạn 1, từ quý IV", _g(8)),
    ("Giai đoạn 2, từ quý III", _g(8)), ("Trên cơ sở phân tích thực trạng tại Chương 2 và các căn cứ", _g(8)),
    ("Theo khoản 3 Điều 17", _g(2)), ("- Về lợi ích của tác giả", _g(2)), ("Thứ hai, thủ tục xác lập quyền", _g(7)),
    ("Thứ nhất, Nhà trường có tầm nhìn sớm", _g(3, 4)), ("Thứ năm, các biểu mẫu", _g(3)),
    ("Thứ nhất, kết quả nghiên cứu chưa", _g(7)), ("Thứ tư, tài sản do Nhà trường", _g(5)),
    ("Các nguyên nhân chủ quan dưới đây", _g(3, 7)), ("Thứ ba, năng lực nghiên cứu", _g(6)), ("Bảng 3.1 cho thấy", _g(7)),
    ("Dữ liệu thực trạng tại Chương 2", _g(7)), ("Hình 2.8 cho thấy", _g(7)), ("Hình 2.9 cho thấy", _g(7, 3)),
    ("Nguồn: Nhóm nghiên cứu tính toán từ danh mục", _g(7)), ("Ở mức chuyển giao trong hệ thống", _g(1)),
    ("Nguồn: Nhóm nghiên cứu tổng hợp từ Luật Sở hữu trí tuệ, Quyết định 213", _g(5)),
    ("Phòng Khoa học Công nghệ chủ trì hoàn thiện", _g(8)), ("Phiếu khai báo và phiếu rà soát", _g(3, 8)),
    ("Chương 2 cho thấy quy chế đã có căn cứ chi", _g(3)), ("Thứ nhất, chỉ đạo hoàn thiện", _g(2, 4)),
    ("Thứ hai, bổ sung phiếu", _g(3, 5)), ("Thứ ba, lập dòng dự toán", _g(3, 8)), ("Về thực trạng", _g(2, 4, 5, 7)),
    ("Về giải pháp", _g(8)), ("Các danh mục được làm sạch", _g(5, 7)), ("Thứ ba, Luật Khoa học", _g(2)),
    ("Hình 2.6 lượng hóa", _g(3, 5)), ("Hình 2.5 cho thấy", _g(7)), ("Hình 2.4 cho thấy", _g(7)),
    ("Bảng 2.4 cho thấy", _g(3, 1, 7)), ("Theo Điều 35 Quyết định 213 và Điều 11", _g(1, 3)),
    ("Hình 2.3 cho thấy", _g(3)), ("Nhà trường có bốn văn bản", _g(3, 4, 7)), ("Tính trên toàn bộ nhân sự", _g(7)),
    ("Trong hai năm cuối kỳ", _g(2)), ("Con số 582", _g(7)), ("Hình 2.1 cho thấy", _g(7)),
    ("Bảng 2.5 cho thấy", _g(5)), ("Tệp theo dõi đơn", _g(5)), ("Bảng 2.6 và Hình 2.7", _g(5)),
    ("Giai đoạn tới, Nhà trường", _g(5, 1)), ("Rà soát cột sản phẩm", _g(7)), ("Đề tài chiết xuất lá Quế hoa", _g(5, 7)),
    ("Hình 2.10 cho thấy", _g(6)), ("Chỉ tiêu 1.11", _g(6, 5)), ("Thứ hai, trong kỳ 2021 - 2025", _g(5)),
    ("Thứ tư, đội ngũ", _g(7)), ("Thứ hai, các quy định nội bộ", _g(2, 4)), ("Thứ sáu, hoạt động khai thác", _g(5)),
    ("Thứ nhất, độ trễ thể chế", _g(7)), ("Các nguyên nhân khách quan", _g(7)), ("Chương 2 đã phân tích", _g(5, 7)),
    ("Về khách quan", _g(3, 4, 7)), ("Trên cơ sở kết quả đánh giá", _g(3, 8)), ("Từ đó, Đề tài đề xuất", _g(8)),
]


def can_cu(cu, moi):
    for dau, cc in CAN_CU_RIENG:
        if moi.startswith(dau):
            return cc
    return _can_cu_tu_khoa(cu, moi)


def so_sanh_doan(cu, moi):
    a = [t for t, _ in cu]
    b = [t for t, _ in moi]
    sm = difflib.SequenceMatcher(None, a, b, autojunk=False)
    hang = []
    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op == "equal":
            continue
        cu_k, moi_k = cu[i1:i2], moi[j1:j2]
        if op == "replace" and len(cu_k) == len(moi_k):
            cap = list(zip(cu_k, moi_k))
        else:
            cap = [(("\n".join(t for t, _ in cu_k)) if cu_k else "", cu_k[0][1] if cu_k else "",
                    ("\n".join(t for t, _ in moi_k)) if moi_k else "", moi_k[0][1] if moi_k else "")]
            cap = [((x[0], x[1]), (x[2], x[3])) for x in cap]
        for (tc, vc), (tm, vm) in cap:
            if _chuan(tc) == _chuan(tm):
                continue
            vt = vm or vc
            hang.append([vt, tc or "(bổ sung)", tm or "(lược bỏ)", can_cu(tc, tm)])
    return hang


def bang_theo_ten(d):
    kq = []
    for tbl in d.tables:
        prev = tbl._tbl.getprevious()
        while prev is not None and prev.tag != qn("w:p"):
            prev = prev.getprevious()
        ten = "".join(x.text or "" for x in prev.iter(qn("w:t"))) if prev is not None else ""
        if not ten.startswith("Bảng"):
            continue
        noi = []
        for r in tbl.rows:
            o = []
            for c in r.cells:
                tx = _chuan(c.text.replace("\n", "; "))
                if not o or o[-1] != tx:
                    o.append(tx)
            noi.append(" | ".join(o))
        kq.append((ten, "\n".join(noi)))
    return kq


def so_sanh_bang(d_cu, d_moi):
    cu, moi = bang_theo_ten(d_cu), bang_theo_ten(d_moi)
    hang, da_dung = [], set()
    for ten_m, nd_m in moi:
        tieu_m = re.sub(r"^Bảng \d\.\d+\. ", "", ten_m)
        tot, k = 0, None
        for i, (ten_c, _) in enumerate(cu):
            r = difflib.SequenceMatcher(None, re.sub(r"^Bảng \d\.\d+\. ", "", ten_c), tieu_m).ratio()
            if r > tot and i not in da_dung:
                tot, k = r, i
        if k is not None and tot >= 0.55:
            da_dung.add(k)
            ten_c, nd_c = cu[k]
            if nd_c == nd_m and ten_c == ten_m:
                continue
            hang.append([f"{ten_m.split('. ')[0]}, toàn bảng", f"{ten_c}\n{nd_c}", f"{ten_m}\n{nd_m}",
                         can_cu(nd_c, f"{ten_m}\n{nd_m}")])
        else:
            hang.append([f"{ten_m.split('. ')[0]}, bảng mới", "(bổ sung)", f"{ten_m}\n{nd_m}",
                         can_cu("", f"{ten_m}\n{nd_m}")])
    for i, (ten_c, nd_c) in enumerate(cu):
        if i not in da_dung:
            hang.append([f"{ten_c.split('. ')[0]}, bản cũ", f"{ten_c}\n{nd_c}", "(lược bỏ hoặc thay bằng bảng mới)", ""])
    return hang


def _o(cell, text, dam=False, co=9.5):
    cell.text = ""
    dong = str(text).split("\n")
    p = cell.paragraphs[0]
    for i, t in enumerate(dong):
        if i:
            p = cell.add_paragraph()
        p.paragraph_format.space_after = Pt(1)
        r = p.add_run(t)
        r.bold = dam
        r.font.size = Pt(co)


def tai_lieu_ngang(tieu_de, mo_ta):
    d = docx.Document()
    st = d.styles["Normal"]
    st.font.name = "Times New Roman"
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    st.font.size = Pt(11)
    for z in d.settings.element.findall(qn("w:zoom")):
        z.set(qn("w:percent"), "100")  # mẫu mặc định của python-docx thiếu thuộc tính bắt buộc này
    s = d.sections[0]
    s.orientation = WD_ORIENT.LANDSCAPE
    s.page_width, s.page_height = Cm(29.7), Cm(21.0)
    for k in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(s, k, Cm(1.8))
    p = d.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(tieu_de)
    r.bold = True
    r.font.size = Pt(14)
    for t in mo_ta:
        d.add_paragraph(t).paragraph_format.space_after = Pt(4)
    return d


def bang_word(d, cot, rong, hang):
    t = d.add_table(rows=1, cols=len(cot))
    t.style = "Table Grid"
    for j, c in enumerate(cot):
        _o(t.rows[0].cells[j], c, dam=True, co=10)
    tr = t.rows[0]._tr
    trpr = tr.get_or_add_trPr()
    etree.SubElement(trpr, qn("w:tblHeader"))
    for h in hang:
        cells = t.add_row().cells
        for j, x in enumerate(h):
            _o(cells[j], x)
    for row in t.rows:
        for j, w in enumerate(rong):
            row.cells[j].width = Cm(w)
    return t


def ghi_giai_trinh(ban_moi):
    d_cu, cu = doan_noi_dung(BAN_CU)
    d_moi, moi = doan_noi_dung(ban_moi)
    cu, moi = vi_tri_day_du(cu), vi_tri_day_du(moi)
    hang_doan = so_sanh_doan(cu, moi)
    hang_bang = so_sanh_bang(d_cu, d_moi)
    hang = hang_doan + hang_bang + [list(x) for x in GIAI_TRINH_HINH]
    d = tai_lieu_ngang(
        "BẢNG GIẢI TRÌNH CHỈNH SỬA BA CHƯƠNG BÁO CÁO ĐỀ TÀI",
        ["Đề tài: Nghiên cứu giải pháp nâng cao hiệu quả quản lý quyền sở hữu trí tuệ tại Trường Đại học Thành Đô.",
         "Căn cứ: Bản góp ý và yêu cầu chỉnh sửa ba chương đề tài. Bản đối chiếu: Ban_cuoi/Bao_cao_tong_ket_de_tai.docx, "
         "bản trước chỉnh sửa; bản sửa: Bao_cao_tong_ket_de_tai.docx trong thư mục này. Ba tệp chương dùng cùng nội dung "
         "với bản toàn văn.",
         "Cách lập: phần lời văn được so khớp tự động từng đoạn giữa hai bản; vị trí ghi theo bản sửa, gồm chương, mục và "
         "thứ tự đoạn trong mục; bảng được so theo tên bảng; hình và sơ đồ được ghi riêng ở cuối. Mục lục, danh mục bảng, "
         "hình và tài liệu tham khảo được cập nhật tự động nên không liệt kê.",
         f"Tổng số mục giải trình: {len(hang)}, gồm {len(hang_doan)} mục lời văn, {len(hang_bang)} mục bảng, "
         f"{len(GIAI_TRINH_HINH)} mục hình, sơ đồ."])
    bang_word(d, ["TT", "Vị trí", "Nội dung cũ", "Nội dung sửa", "Căn cứ"], [1.0, 3.6, 9.0, 9.0, 3.5],
              [[i + 1] + h for i, h in enumerate(hang)])
    ra = os.path.join(THU_MUC, "Bang_giai_trinh_chinh_sua.docx")
    d.save(ra)
    print("Đã ghi:", ra, f"({len(hang)} mục)")
    return ra


# ---------------------------------------------------------------------------
# 3. Danh sách điểm chưa đủ minh chứng
# ---------------------------------------------------------------------------
CHUA_DU = [
    ("Trạng thái pháp lý của 5 kiểu dáng công nghiệp", "Sheet2 ghi số hiệu 3-00396xx-000, năm 2024; Sheet1 ghi chờ cấp "
     "bằng", "Số văn bằng trong kỳ là 4 hay 9; tỷ lệ chỉ tiêu 1.11 Kế hoạch 07 chưa xác định, có thể là 125%",
     "Bản sao văn bằng hoặc kết quả tra cứu trên cơ sở dữ liệu của Cục Sở hữu trí tuệ"),
    ("Dòng sáng chế ghi chấp nhận đơn hợp lệ tại Sheet1", "Không có số đơn, tên đơn", "Không gán được cho đơn nào",
     "Thông báo chấp nhận đơn hợp lệ của Cục Sở hữu trí tuệ"),
    ("Ngày nộp đơn và ngày công bố kết quả của hai đơn sáng chế", "Chỉ có số đơn 1-2025-07378, 1-2026-07185",
     "Chưa biết đơn được nộp trước hay sau khi công bố", "Tờ khai nộp đơn; danh sách bài báo, báo cáo liên quan"),
    ("Tình trạng bộc lộ của 6 đề tài mã số 2021 - 2024 có sản phẩm tiềm năng sở hữu công nghiệp", "Chưa có",
     "Chưa xác định sản phẩm nào còn khả năng đăng ký", "Danh mục công bố của từng đề tài, ngày công bố"),
    ("Điều lệ Quỹ Học bổng sau tiến sĩ Ngô Xuân Độ", "Không có văn bản trong hồ sơ", "Bảng 2.3, Bảng 2.4 ghi theo thông "
     "tin được cung cấp", "Văn bản Điều lệ, Điều 9 về phân chia lợi ích"),
    ("Quy chế chi tiêu nội bộ áp dụng giai đoạn 2021 - 2025", "Chỉ có quy chế ban hành 01/8/2026",
     "Chưa đánh giá được vai trò của cơ chế khuyến khích đối với kết quả của kỳ", "Quy chế chi tiêu các năm 2021 - 2025"),
    ("Chế độ lợi ích của hai đề tài cấp quốc gia phê duyệt trước 01/10/2025", "Không có Luật Khoa học và công nghệ năm "
     "2013, văn bản hướng dẫn và hợp đồng tài trợ", "Chỉ xác định được khung chuyển tiếp tại khoản 3, khoản 7 Điều 73",
     "Hợp đồng tài trợ, quyết định phê duyệt, văn bản hướng dẫn có hiệu lực tại thời điểm phê duyệt"),
    ("Mức độ thực hiện Điều 10, Điều 11 Quyết định 217", "Chưa thấy biểu mẫu khai báo, hồ sơ theo dõi dùng chung",
     "Nguyên nhân chủ quan được trình bày dưới dạng giả thuyết", "Biểu mẫu, sổ theo dõi do Phòng Khoa học Công nghệ "
     "lập; báo cáo thực hiện"),
    ("Số liệu chi thực tế cho xác lập quyền 2021 - 2025", "Chưa có", "Tiêu chí kinh phí chỉ tính được một phần",
     "Sổ chi lệ phí, phí đại diện theo năm"),
    ("Tư cách thành viên Mạng lưới Trung tâm Hỗ trợ công nghệ và đổi mới sáng tạo từ năm 2023; đăng ký hoạt động khoa học "
     "công nghệ, con dấu và 5 hợp đồng của Viện Nghiên cứu giáo dục và Chuyển giao tri thức",
     "Chỉ có thông tin được cung cấp, chưa có văn bản", "Các nội dung này được ghi là thông tin chưa xác minh",
     "Văn bản kết nạp; giấy chứng nhận đăng ký hoạt động; bản hợp đồng"),
    ("Liên kết bài báo với đề tài", "Danh mục bài báo không ghi mã đề tài; con số 107 bài của bản trước không tái lập "
     "được", "Không xác định được số bản ghi trùng; 582 chỉ là tổng số bản ghi", "Trường mã đề tài trong danh mục bài báo"),
    ("Ký hiệu, số, ngày của quy chế nội bộ", "Bìa Quy chế ghi Quyết định 213/QĐ-ĐHTĐ, Kế hoạch 07 ghi Nghị quyết "
     "213/NQ-ĐHTĐ; tệp Quy chế sở hữu trí tuệ không ghi số, ngày", "Ảnh hưởng cách trích dẫn văn bản",
     "Bản ký, đóng dấu của Quyết định 213 và Quyết định 217"),
    ("Thuyết minh đề tài được phê duyệt", "Các tệp thuyết minh không có mã số, quyết định giao", "Không dùng làm căn cứ "
     "cho nội dung đã phê duyệt", "Thuyết minh kèm quyết định giao nhiệm vụ"),
]


def ghi_chua_du():
    d = tai_lieu_ngang("DANH SÁCH CÁC ĐIỂM CHƯA ĐỦ MINH CHỨNG ĐỂ KẾT LUẬN",
                       ["Các điểm dưới đây chưa được kết luận trong ba chương; nội dung liên quan được ghi là chưa xác "
                        "minh, theo thông tin được cung cấp hoặc dưới dạng giả thuyết."])
    bang_word(d, ["TT", "Nội dung", "Hiện trạng hồ sơ", "Ảnh hưởng đến kết luận", "Đề nghị bổ sung"],
              [1.0, 6.6, 6.0, 6.2, 6.2], [[i + 1] + list(h) for i, h in enumerate(CHUA_DU)])
    ra = os.path.join(THU_MUC, "Danh_sach_diem_chua_du_minh_chung.docx")
    d.save(ra)
    print("Đã ghi:", ra)
    return ra


def main():
    dung_chuong()
    ban_moi = dung_bao_cao()
    ghi_giai_trinh(ban_moi)
    ghi_chua_du()


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "giai_trinh":
        ghi_giai_trinh(os.path.join(THU_MUC, "Bao_cao_tong_ket_de_tai.docx"))
        ghi_chua_du()
    else:
        main()

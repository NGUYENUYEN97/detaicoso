# -*- coding: utf-8 -*-
"""Định nghĩa các hình của Chương 2: bảng dữ liệu, cấu hình biểu đồ, lời bình.

Mỗi hình là một dict:
  id, tieu_de, nguon, cot (tiêu đề cột), dong (dữ liệu), dinh_dang (định dạng số
  theo cột), bieu_do (cấu hình biểu đồ), sau_doan (đoạn văn trong Word mà hình
  được chèn ngay sau), binh_luan (các đoạn nhận xét đặt dưới hình).
"""
from du_lieu import (NAM, NHAN_LUC, SAN_PHAM, THAM_LUAN_QUOC_GIA,
                     PHAN_HANG_QT, DON_VI_SAN_LUONG, NHAN_SU_CO_TEN, NHAN_SU_CO_BAI, GINI_BAI,
                     DE_TAI, KHUYEN_KHICH, KENH_TAI_TRO, TSTT, TSTT_KY, TIEU_CHI, KE_HOACH,
                     DA_CAP, DA_NOP, CHUA_XM)

# Bảng màu phân loại cố định (đã kiểm tra phân biệt được với người mù màu).
XANH, CAM, NGOC, VANG, HONG, LUC, TIM = ("#2A78D6", "#EB6834", "#1BAF7A", "#EDA100",
                                          "#E87BA4", "#008300", "#4A3AA7")
XAM = "#A8A7A1"
# Thang một sắc độ xanh cho dữ liệu có thứ bậc (Q1 đậm nhất).
XANH_THANG = ["#0B3D82", "#2A78D6", "#6FA8E8", "#B5D2F4"]


def so(x, d=0):
    """Định dạng số kiểu Việt Nam: dấu chấm phân nghìn, dấu phẩy thập phân."""
    s = f"{x:,.{d}f}"
    return s.replace(",", "_").replace(".", ",").replace("_", ".")


def pt(x, d=1):
    return so(100 * x, d) + "%"


def cagr(a, b, n):
    return (b / a) ** (1 / n) - 1


# ---------------------------------------------------------------------------
# Các đại lượng dẫn xuất dùng trong lời bình
# ---------------------------------------------------------------------------
tong_nl = sum(r[2] for r in NHAN_LUC)
ts_toan_truong = sum(r[5] for r in NHAN_LUC)
ts_ba_vien = sum(r[5] for r in NHAN_LUC[:3])
nl_ba_vien = sum(r[2] for r in NHAN_LUC[:3])
hoc_ham = sum(r[3] + r[4] for r in NHAN_LUC)
hoc_ham_ba_vien = sum(r[3] + r[4] for r in NHAN_LUC[:3])
khoi_qt = [r for r in NHAN_LUC if r[1] in ("TT Tuyển sinh và QTTH", "TT Dịch vụ và QTHC")]
nl_khoi_qt = sum(r[2] for r in khoi_qt)
dh_khac_khoi_qt = sum(r[7] + r[8] for r in khoi_qt)

sp = {ten: v for ten, v in SAN_PHAM}
bb = [a + b for a, b in zip(sp["Bài báo đăng tạp chí trong nước"], sp["Bài báo đăng tạp chí quốc tế"])]
gt = sp["Giáo trình, tài liệu giảng dạy"]
tong_nam = [sum(v[i] for _, v in SAN_PHAM) for i in range(5)]
tong_ban_ghi = sum(tong_nam) + THAM_LUAN_QUOC_GIA
ty_so = [b / g for b, g in zip(bb, gt)]

q = PHAN_HANG_QT
qt = sp["Bài báo đăng tạp chí quốc tế"]
q12 = [q["Q1"][i] + q["Q2"][i] for i in range(5)]
co_hang = [q["Q1"][i] + q["Q2"][i] + q["Q3"][i] + q["Q4"][i] for i in range(5)]

dt_nam = {n: [d for d in DE_TAI if d[1] == int(n)] for n in NAM}
dt_tien = [sum(1 for d in dt_nam[n] if d[4] == "Tiền mặt") for n in NAM]
dt_gio = [sum(1 for d in dt_nam[n] if d[4] == "Quy đổi giờ") for n in NAM]
dt_tu = [sum(1 for d in dt_nam[n] if d[4] == "Tự tìm tài trợ") for n in NAM]
kp_nam = [sum(d[5] for d in dt_nam[n]) for n in NAM]
kp_tong = sum(kp_nam)
dt_co_tien = [d for d in DE_TAI if d[4] == "Tiền mặt"]
dt_du_dk = [d for d in DE_TAI if d[6]]
kp_du_dk = sum(d[5] for d in dt_du_dk)
kp_sorted = sorted(d[5] for d in dt_co_tien)
kp_trung_vi = kp_sorted[len(kp_sorted) // 2]
dt_duoc = [d for d in dt_du_dk if d[2] == "Viện Y - Dược" or d[0] == "08-2023"]
dt_den_2024 = [d for d in DE_TAI if d[1] <= 2024]
du_dk_den_2024 = [d for d in dt_du_dk if d[1] <= 2024]
nop_don = [d for d in dt_du_dk if d[8].startswith("Đã nộp")]
# Phễu sở hữu công nghiệp: không gộp hai sản phẩm thuộc quyền tác giả (bộ mẫu cây thuốc, bộ tiêu bản)
dt_qtg = [d for d in dt_du_dk if d[7].startswith("Sưu tập dữ liệu")]
dt_shcn = [d for d in dt_du_dk if d not in dt_qtg]
dt_shcn_den_2024 = [d for d in dt_shcn if d[1] <= 2024]

tstt_nam = list(range(2021, 2027))
nguon_ts = ["Thương hiệu", "Hợp tác doanh nghiệp", "Nghiên cứu"]


def dem_ts(nam, nguon):
    return sum(1 for t in TSTT if t[3] == nam and t[6] == nguon)


luy_ke, s = [], 0
for n in tstt_nam:
    s += sum(dem_ts(n, g) for g in nguon_ts)
    luy_ke.append(s)

# ---------------------------------------------------------------------------
# Danh sách hình
# ---------------------------------------------------------------------------
HINH = []

# H2.1 ----------------------------------------------------------------------
ty_le_ts = [r[5] / r[2] for r in NHAN_LUC]
HINH.append(dict(
    tieu_de="Cơ cấu nhân lực theo đơn vị, trình độ và tỷ lệ tiến sĩ, năm 2026",
    nguon="Nguồn: Nhóm nghiên cứu tổng hợp từ danh sách nhân sự năm 2026 của Trường Đại học Thành Đô. "
          "Nhóm tiến sĩ và tương đương đã bao gồm người có học hàm giáo sư, phó giáo sư và bác sĩ chuyên khoa II.",
    cot=["Đơn vị", "Tiến sĩ và tương đương", "Thạc sĩ và tương đương", "Đại học", "Khác", "Tỷ lệ tiến sĩ"],
    dong=[[r[1], r[5], r[6], r[7], r[8], r[5] / r[2]] for r in NHAN_LUC],
    dinh_dang=[None, "0", "0", "0", "0", "0.0%"],
    bieu_do=dict(loai="column", xep_chong=True, khoang_cach=55,
                 chuoi=[dict(cot=1, mau=XANH_THANG[0]), dict(cot=2, mau=XANH_THANG[1]),
                        dict(cot=3, mau=XANH_THANG[2]), dict(cot=4, mau=XAM),
                        dict(cot=5, kieu="line", mau=CAM, truc_phu=True, nhan=True, duong=False,
                             vi_tri_nhan="above")],
                 truc_y="Số người", max_y=120, buoc_y=20,
                 truc_y2="Tỷ lệ tiến sĩ và tương đương", dd_y2="0%", max_y2=1),
    sau_doan="Trong 252 nhân sự có 145 người giữ vị trí giảng viên",
    binh_luan=[
        f"Hình 2.1 cho thấy nguồn nhân lực trình độ cao tập trung gần như tuyệt đối tại ba viện đào tạo. "
        f"Ba viện chiếm {pt(nl_ba_vien / tong_nl)} tổng nhân sự nhưng nắm {ts_ba_vien} trên {ts_toan_truong} "
        f"người có trình độ tiến sĩ và tương đương, tương đương {pt(ts_ba_vien / ts_toan_truong)}, cùng "
        f"{hoc_ham_ba_vien} trên {hoc_ham} người có học hàm. Tỷ lệ tiến sĩ tại Viện Quản trị và Công nghệ đạt "
        f"{pt(ty_le_ts[1])} và tại Viện Y - Dược đạt {pt(ty_le_ts[0])}, trong khi Viện Ngôn ngữ - Văn hóa - Quốc tế "
        f"chỉ đạt {pt(ty_le_ts[2])} do cơ cấu đội ngũ nghiêng về trình độ thạc sĩ.",
        f"Ở chiều ngược lại, hai trung tâm thuộc khối Quản trị và Dịch vụ có {nl_khoi_qt} nhân sự nhưng không có "
        f"người nào có trình độ tiến sĩ; {dh_khac_khoi_qt} người, tương đương {pt(dh_khac_khoi_qt / nl_khoi_qt)}, "
        f"có trình độ đại học hoặc thấp hơn. Hai trung tâm này đang giữ hai trong bốn đầu mối quản lý quyền sở hữu trí "
        f"tuệ, như phân tích tại Mục 2.2.2. Như vậy, năng lực chuyên môn để nhận diện kết quả nghiên cứu có khả năng bảo hộ nằm "
        f"ở khối Đào tạo và Nghiên cứu, còn thẩm quyền xử lý hồ sơ lại nằm ở khối có năng lực chuyên môn khoa học thấp "
        f"hơn. Sự lệch pha này là tiền đề để lý giải các điểm nghẽn về tổ chức ở các mục sau.",
    ],
))

# H2.2 ----------------------------------------------------------------------
loai_ngan = ["Bài báo trong nước", "Bài báo quốc tế", "Giáo trình, tài liệu", "Đề tài cấp cơ sở",
             "Sách", "Tham luận quốc tế", "Đề tài cấp quốc gia"]
HINH.append(dict(
    tieu_de="Cơ cấu và diễn biến sản phẩm khoa học theo năm, giai đoạn 2021 - 2025",
    nguon="Nguồn: Nhóm nghiên cứu tổng hợp từ các danh mục thống kê của Phòng Khoa học Công nghệ, Trường Đại học "
          "Thành Đô. Không gồm 21 tham luận hội thảo quốc gia do danh mục không có trường thời gian.",
    cot=["Năm"] + loai_ngan + ["Tổng số"],
    dong=[[NAM[i]] + [SAN_PHAM[k][1][i] for k in range(7)] + [tong_nam[i]] for i in range(5)],
    dinh_dang=[None] + ["0"] * 8,
    bieu_do=dict(loai="column", xep_chong=True, khoang_cach=45,
                 chuoi=[dict(cot=1, mau=XANH), dict(cot=2, mau=CAM), dict(cot=3, mau=NGOC),
                        dict(cot=4, mau=VANG), dict(cot=5, mau=HONG), dict(cot=6, mau=LUC),
                        dict(cot=7, mau=TIM),
                        dict(cot=8, kieu="line", mau="#3B3A36", nhan=True, vi_tri_nhan="above")],
                 truc_y="Số sản phẩm"),
    sau_doan="Con số 582 chỉ phản ánh",
    binh_luan=[
        f"Hình 2.2 cho thấy tổng sản lượng tăng từ {tong_nam[0]} sản phẩm năm 2021 lên {tong_nam[4]} sản phẩm năm "
        f"2025, tốc độ tăng bình quân {pt(cagr(tong_nam[0], tong_nam[4], 4))} một năm. Năm 2023 là năm duy nhất sản "
        f"lượng giảm, từ {tong_nam[1]} xuống {tong_nam[2]} sản phẩm, do số giáo trình giảm từ {gt[1]} xuống {gt[2]} "
        f"trong khi bài báo chỉ tăng nhẹ. Từ năm 2024, quy mô tăng trở lại và tăng mạnh.",
        f"Cơ cấu sản phẩm thay đổi rõ rệt giữa đầu và cuối kỳ. Năm 2021, bài báo chiếm {pt(bb[0] / tong_nam[0])} "
        f"và giáo trình chiếm {pt(gt[0] / tong_nam[0])} sản lượng; đến năm 2025, bài báo chiếm {pt(bb[4] / tong_nam[4])} "
        f"còn giáo trình chỉ còn {pt(gt[4] / tong_nam[4])}. Sự tăng trưởng của Nhà trường trong giai đoạn này về thực "
        f"chất là sự tăng trưởng của công bố khoa học, nhóm sản phẩm chỉ phát sinh quyền tác giả và gần như không có "
        f"khả năng khai thác thương mại.",
    ],
))

# H2.3 ----------------------------------------------------------------------
HINH.append(dict(
    tieu_de="Tương quan giữa bài báo khoa học và giáo trình, tài liệu giảng dạy, giai đoạn 2021 - 2025",
    nguon="Nguồn: Nhóm nghiên cứu tính toán từ Bảng 2.2. Bài báo gồm bài đăng tạp chí trong nước và quốc tế.",
    cot=["Năm", "Bài báo khoa học", "Giáo trình, tài liệu giảng dạy", "Số bài báo trên một giáo trình"],
    dong=[[NAM[i], bb[i], gt[i], ty_so[i]] for i in range(5)],
    dinh_dang=[None, "0", "0", "0.0"],
    bieu_do=dict(loai="column", khoang_cach=80,
                 chuoi=[dict(cot=1, mau=XANH, nhan=True), dict(cot=2, mau=NGOC, nhan=True),
                        dict(cot=3, kieu="line", mau=CAM, truc_phu=True, nhan=True, vi_tri_nhan="above")],
                 truc_y="Số sản phẩm", truc_y2="Số bài báo trên một giáo trình (lần)", dd_y2="0"),
    sau_doan="Số liệu tại Bảng 2.2 cho thấy hai xu hướng ngược chiều",
    binh_luan=[
        f"Hình 2.3 thể hiện rõ hai đường xu hướng tách xa nhau. Số bài báo tăng liên tục qua cả năm năm, từ {bb[0]} "
        f"lên {bb[4]} bài, gấp {so(bb[4] / bb[0], 1)} lần. Số giáo trình đạt đỉnh {gt[1]} tài liệu năm 2022 rồi giảm "
        f"mạnh, chỉ còn {gt[3]} tài liệu năm 2024 và {gt[4]} tài liệu năm 2025. Đường tỷ số cho thấy năm 2023 là điểm "
        f"gãy: tỷ số tăng từ {so(ty_so[1], 1)} lên {so(ty_so[2], 1)} lần và đạt mức cao nhất {so(ty_so[3], 1)} lần vào "
        f"năm 2024.",
        "Đối với quản lý quyền sở hữu trí tuệ, xu hướng này có hai hàm ý. Thứ nhất, "
        "nguồn lực khoa học đang dịch chuyển khỏi nhóm sản phẩm có khả năng khai thác trong đào tạo và chuyển giao quyền "
        "sử dụng. Thứ hai, số lượng giáo trình lớn của các năm 2021 - 2022 là một kho tài sản thuộc quyền tác giả chưa "
        "được xác lập quyền, cần được rà soát trước khi tiếp tục mở rộng biên soạn.",
    ],
))

# H2.4 ----------------------------------------------------------------------
khac_qt = q["Chưa xếp hạng hoặc tạp chí quốc tế khác"]
HINH.append(dict(
    tieu_de="Bài báo quốc tế theo phân hạng tạp chí và tỷ trọng nhóm Q1 - Q2, giai đoạn 2021 - 2025",
    nguon="Nguồn: Nhóm nghiên cứu tổng hợp từ cột phân hạng ISI, Scopus của danh mục bài báo quốc tế, "
          "Phòng Khoa học Công nghệ, Trường Đại học Thành Đô.",
    cot=["Năm", "Q1", "Q2", "Q3", "Q4", "Chưa xếp hạng, quốc tế khác", "Tỷ trọng Q1 - Q2"],
    dong=[[NAM[i], q["Q1"][i], q["Q2"][i], q["Q3"][i], q["Q4"][i], khac_qt[i], q12[i] / qt[i]] for i in range(5)],
    dinh_dang=[None, "0", "0", "0", "0", "0", "0.0%"],
    bieu_do=dict(loai="column", xep_chong=True, khoang_cach=55,
                 chuoi=[dict(cot=1, mau=XANH_THANG[0]), dict(cot=2, mau=XANH_THANG[1]),
                        dict(cot=3, mau=XANH_THANG[2]), dict(cot=4, mau=XANH_THANG[3]),
                        dict(cot=5, mau=XAM),
                        dict(cot=6, kieu="line", mau=CAM, truc_phu=True, nhan=True, vi_tri_nhan="above")],
                 truc_y="Số bài báo", truc_y2="Tỷ trọng Q1 - Q2", dd_y2="0%", max_y2=1),
    sau_doan="Điểm có ý nghĩa đối với quản lý quyền sở hữu trí tuệ là sự dịch chuyển",
    binh_luan=[
        f"Không chỉ tăng về số lượng, công bố quốc tế của Nhà trường còn chuyển dịch mạnh về chất lượng. Trong "
        f"{sum(qt)} bài báo quốc tế, {sum(co_hang)} bài, tương đương {pt(sum(co_hang) / sum(qt))}, đăng trên tạp chí "
        f"có phân hạng Q, trong đó {sum(q12)} bài thuộc nhóm Q1 - Q2. Hình 2.4 cho thấy giai đoạn 2021 - 2023 gần như "
        f"không có bài thuộc nhóm có phân hạng, với tỷ trọng Q1 - Q2 chỉ từ 0% đến {pt(q12[1] / qt[1])}. Bước ngoặt "
        f"diễn ra năm 2024 khi tỷ trọng này đạt {pt(q12[3] / qt[3])}, và năm 2025 duy trì ở mức {pt(q12[4] / qt[4])} "
        f"trên quy mô lớn hơn. Hai năm cuối kỳ đóng góp {sum(co_hang[3:])} trên {sum(co_hang)} bài có phân hạng.",
        "Kết quả này cho thấy đội ngũ của Nhà trường có năng lực tạo ra công bố chất lượng cao khi có định hướng rõ "
        "ràng. Cùng một đội ngũ, trong cùng giai đoạn, đã tạo ra bước nhảy về công bố nhưng không tạo ra bước nhảy "
        "tương ứng về đơn đăng ký sở hữu công nghiệp. Điều đó gợi ý rằng khoảng trống về tài sản trí tuệ không xuất "
        "phát chủ yếu từ năng lực nghiên cứu mà từ cách thiết kế động lực và quy trình, luận điểm được phân tích tiếp "
        "tại Hình 2.11.",
    ],
))

# H2.5 ----------------------------------------------------------------------
HINH.append(dict(
    tieu_de="Quy mô nhân lực và số bài báo bình quân một nhân sự theo đơn vị, giai đoạn 2021 - 2025",
    nguon="Nguồn: Nhóm nghiên cứu tính toán từ Bảng 2.3.",
    cot=["Đơn vị", "Nhân lực", "Bài báo trên một nhân sự", "Đề tài trên 100 nhân sự"],
    dong=[[r[0], r[1], r[4] / r[1], 100 * r[2] / r[1]] for r in DON_VI_SAN_LUONG],
    dinh_dang=[None, "0", "0.00", "0.0"],
    bieu_do=dict(loai="column", khoang_cach=90,
                 chuoi=[dict(cot=1, mau=XANH, nhan=True),
                        dict(cot=2, kieu="line", mau=CAM, truc_phu=True, nhan=True, duong=False,
                             vi_tri_nhan="right")],
                 truc_y="Số nhân sự", truc_y2="Số bài báo trên một nhân sự", dd_y2="0.0"),
    sau_doan="Điểm đáng chú ý là Viện Y - Dược, đơn vị có năng suất bình quân thấp nhất",
    binh_luan=[
        "Hình 2.5 đặt quy mô và năng suất trên cùng một mặt phẳng. Cột và điểm đánh dấu gần như đối xứng ngược nhau: "
        "đơn vị có cột cao nhất là Viện Y - Dược lại có điểm năng suất thấp nhất, trong khi hai đơn vị có cột thấp nhất "
        "lại có điểm năng suất cao nhất. Quy mô nhân lực vì vậy không phải là biến số giải thích năng suất nghiên cứu "
        "của Nhà trường.",
        "Hàm ý quản lý là các chính sách phân bổ nguồn lực theo quy mô đơn vị sẽ không trúng đích. Số bài báo và khả "
        "năng hình thành tài sản trí tuệ cần được theo dõi bằng hai thước đo riêng; nếu chỉ dùng số bài báo, Nhà trường "
        "sẽ đánh giá thấp chính đơn vị có tiềm năng tài sản trí tuệ lớn nhất, như thể hiện tại Hình 2.14.",
    ],
))

# H2.6 ----------------------------------------------------------------------
HINH.append(dict(
    tieu_de="Cơ cấu nhân sự theo tình trạng có công bố bài báo, giai đoạn 2021 - 2025",
    nguon=f"Nguồn: Nhóm nghiên cứu tính toán trên {NHAN_SU_CO_TEN} tên khác nhau của 252 nhân sự trong danh sách năm "
          "2026, đối chiếu với danh sách tác giả của hai danh mục bài báo theo quy tắc giữ dấu, đúng thứ tự họ tên.",
    cot=["Tình trạng", "Số nhân sự"],
    dong=[["Có ít nhất một bài báo", NHAN_SU_CO_BAI], ["Chưa có bài báo", NHAN_SU_CO_TEN - NHAN_SU_CO_BAI]],
    dinh_dang=[None, "0"],
    bieu_do=dict(loai="doughnut", mau_diem=[XANH, "#D9D8D3"], lo=58),
    sau_doan="Mức độ tập trung sản phẩm theo đơn vị, đo bằng chỉ số Herfindahl",
    binh_luan=[
        f"Hình 2.6 minh họa trực quan mức độ tập trung nói trên: gần hai phần ba nhân sự, cụ thể "
        f"{NHAN_SU_CO_TEN - NHAN_SU_CO_BAI} trên {NHAN_SU_CO_TEN} người, chưa đứng tên bài báo nào trong năm năm. Kết "
        f"hợp với hệ số Gini {so(GINI_BAI, 3)}, dữ liệu cho thấy hoạt động công bố của Nhà trường dựa trên khoảng {NHAN_SU_CO_BAI} người, "
        f"và trong nhóm này lại dựa chủ yếu vào khoảng 9 người dẫn đầu.",
    ],
))

# H2.7 ----------------------------------------------------------------------
nhom_tc = ["Đầu vào", "Quá trình", "Đầu ra", "Kết quả"]
dem_tc = {n: [sum(1 for t in TIEU_CHI if t[0] == n and t[2] == m) for m in "TMC"] for n in nhom_tc}
tong_T = sum(dem_tc[n][0] for n in nhom_tc)
tong_M = sum(dem_tc[n][1] for n in nhom_tc)
tong_C = sum(dem_tc[n][2] for n in nhom_tc)
HINH.append(dict(
    tieu_de="Khả năng tính toán bộ tiêu chí đánh giá hiệu quả quản lý quyền sở hữu trí tuệ từ dữ liệu hiện có",
    nguon="Nguồn: Đánh giá của nhóm nghiên cứu đối với 16 tiêu chí tại Mục 1.4.2, căn cứ vào các danh mục và văn bản "
          "hiện có của Trường Đại học Thành Đô. Chi tiết từng tiêu chí tại tệp dữ liệu kèm theo.",
    cot=["Nhóm tiêu chí", "Tính được đầy đủ", "Tính được một phần", "Chưa tính được"],
    dong=[[n] + dem_tc[n] for n in nhom_tc],
    dinh_dang=[None, "0", "0", "0"],
    bieu_do=dict(loai="bar", xep_chong=True, khoang_cach=50, dao_truc=True,
                 chuoi=[dict(cot=1, mau=NGOC, nhan=True, an_nhan_0=True),
                        dict(cot=2, mau=VANG, nhan=True, an_nhan_0=True),
                        dict(cot=3, mau=CAM, nhan=True, an_nhan_0=True)],
                 truc_y="Số tiêu chí"),
    sau_doan="Thứ tư, hệ quả của ba hạn chế trên",
    binh_luan=[
        f"Hình 2.7 lượng hóa nhận định trên. Trong 16 tiêu chí đã xây dựng, chỉ {tong_T} tiêu chí tính được đầy đủ, "
        f"{tong_M} tiêu chí tính được một phần và {tong_C} tiêu chí chưa tính được. Mức độ thiếu hụt tăng dần theo chuỗi "
        f"đánh giá: nhóm đầu ra có {dem_tc['Đầu ra'][0]} trên 4 tiêu chí tính được đầy đủ, còn nhóm kết quả không có "
        f"tiêu chí nào tính được đầy đủ. Nói cách khác, Nhà trường đo được mình đã xác lập bao nhiêu quyền nhưng chưa "
        f"đo được quyền đó mang lại giá trị gì, trong khi đây chính là nhóm thông tin mà Luật Giáo dục đại học số "
        f"125/2025/QH15 yêu cầu công khai.",
    ],
))

# H2.8 ----------------------------------------------------------------------
# Chỉ đặt cạnh nhau hai quy định có cùng cơ sở tính: nguồn thu sau khi trừ các khoản chi phí cần thiết, hợp lệ.
# Điều 135 Luật Sở hữu trí tuệ (tổng tiền trước thuế) và Điều 28 Luật số 93/2025/QH15 (lợi nhuận sau thuế) khác cơ sở
# tính nên được trình bày tại Bảng 2.3, không vẽ chung.
muc_thu = [0, 100, 200, 300, 400, 500, 600, 700]
HINH.append(dict(
    tieu_de="Mô phỏng phần dành cho tác giả theo điểm a và điểm b khoản 4 Điều 36 Quyết định 213 trên cùng cơ sở tính",
    nguon="Nguồn: Nhóm nghiên cứu mô phỏng từ khoản 4 Điều 36 Quy chế ban hành kèm Quyết định số 213/QĐ-ĐHTĐ. Trục "
          "hoành là nguồn thu sau khi trừ các khoản chi phí cần thiết, hợp lệ, cơ sở tính chung của hai điểm; đơn vị: "
          "triệu đồng. Điểm a áp dụng cho sản phẩm đề tài sử dụng ngân sách nhà nước do Trường chủ trì: 30% khen thưởng "
          "tập thể tác giả, tối đa 100 triệu đồng một đề tài, phần vượt chuyển vào quỹ khen thưởng, phúc lợi. Điểm b "
          "áp dụng cho tài sản trí tuệ thuộc sở hữu của Trường: tác giả hưởng 30%. Hai điểm có phạm vi khác nhau nên "
          "không cùng áp dụng cho một tài sản.",
    cot=["Nguồn thu sau chi phí hợp lệ", "Điểm a: khen thưởng tập thể tác giả, đề tài sử dụng ngân sách nhà nước",
         "Điểm b: tác giả, tài sản trí tuệ thuộc sở hữu của Trường"],
    dong=[[str(x), min(0.3 * x, 100), 0.3 * x] for x in muc_thu],
    dinh_dang=[None, "0", "0"],
    bieu_do=dict(loai="line",
                 chuoi=[dict(cot=1, mau=XANH, dam=2.25), dict(cot=2, mau=NGOC, dam=2.25, gach=True)],
                 truc_y="Phần dành cho tác giả (triệu đồng)",
                 truc_x="Nguồn thu sau khi trừ chi phí cần thiết, hợp lệ (triệu đồng)"),
    sau_doan="",
    binh_luan=[],
))

# H2.9 ----------------------------------------------------------------------
HINH.append(dict(
    tieu_de="Quy mô kinh phí của ba kênh tài trợ nghiên cứu",
    nguon="Nguồn: Nhóm nghiên cứu tổng hợp từ danh mục đề tài cấp cơ sở, danh mục đề tài cấp quốc gia và Điều lệ "
          "Quỹ Học bổng sau tiến sĩ Ngô Xuân Độ; đơn vị: triệu đồng. Kinh phí đề tài cấp cơ sở chỉ tính khoản Nhà "
          "trường cấp bằng tiền.",
    cot=["Kênh tài trợ", "Kinh phí (triệu đồng)"],
    dong=[[k, v] for k, v in KENH_TAI_TRO],
    dinh_dang=[None, "#,##0"],
    bieu_do=dict(loai="bar", khoang_cach=60, dao_truc=True,
                 chuoi=[dict(cot=1, mau=XANH, nhan=True, mau_diem=[CAM, XANH, XANH])],
                 truc_y="Triệu đồng", an_chu_giai=True),
    sau_doan="Bảng 2.5 cho thấy một nghịch lý về phân bổ nguồn lực",
    binh_luan=[
        f"Hình 2.9 cho thấy độ chênh về quy mô giữa các kênh. Kinh phí Nhà trường cấp cho toàn bộ 38 đề tài cấp cơ sở "
        f"trong năm năm chỉ bằng {pt(KENH_TAI_TRO[0][1] / KENH_TAI_TRO[2][1])} quy mô Quỹ Học bổng sau tiến sĩ và "
        f"{pt(KENH_TAI_TRO[0][1] / KENH_TAI_TRO[1][1])} tổng kinh phí ba đề tài cấp quốc gia. Tuy nhiên, kênh nhỏ nhất "
        f"này lại là kênh duy nhất đến nay đã tạo ra sản phẩm có khả năng xác lập quyền, với 11 đề tài có sản phẩm cụ "
        f"thể. Hai kênh lớn đều mới bắt đầu từ năm 2024 - 2025 và chưa có sản phẩm nghiệm thu.",
    ],
))

# H2.10 ---------------------------------------------------------------------
HINH.append(dict(
    tieu_de="Số đề tài cấp cơ sở theo hình thức kinh phí và kinh phí Nhà trường cấp, giai đoạn 2021 - 2025",
    nguon="Nguồn: Nhóm nghiên cứu tổng hợp từ danh mục đề tài khoa học công nghệ cấp cơ sở giai đoạn 2021 - 2025. "
          "Đề tài xếp theo năm ghi trong mã số; kinh phí đơn vị triệu đồng.",
    cot=["Năm", "Được cấp kinh phí bằng tiền", "Chỉ quy đổi giờ nghiên cứu", "Tự tìm nguồn tài trợ",
         "Kinh phí Nhà trường cấp (triệu đồng)"],
    dong=[[NAM[i], dt_tien[i], dt_gio[i], dt_tu[i], kp_nam[i]] for i in range(5)],
    dinh_dang=[None, "0", "0", "0", "0.00"],
    bieu_do=dict(loai="column", xep_chong=True, khoang_cach=60,
                 chuoi=[dict(cot=1, mau=XANH), dict(cot=2, mau=XAM), dict(cot=3, mau=VANG),
                        dict(cot=4, kieu="line", mau=CAM, truc_phu=True, nhan=True, vi_tri_nhan="above",
                             dd_nhan="0", an_nhan_0=True)],
                 truc_y="Số đề tài", truc_y2="Kinh phí (triệu đồng)", dd_y2="0"),
    sau_doan="Ngược lại, kênh sinh ra nhiều sản phẩm có khả năng bảo hộ nhất",
    binh_luan=[
        f"Hình 2.10 cho thấy cơ chế cấp kinh phí đề tài cấp cơ sở thay đổi căn bản sau năm 2022. Toàn bộ 6 đề tài năm "
        f"2021 và {dt_gio[1]} trên {len(dt_nam['2022'])} đề tài năm 2022 chỉ được quy đổi giờ nghiên cứu; từ năm 2023, "
        f"phần lớn đề tài được cấp kinh phí bằng tiền và tổng kinh phí tăng từ {so(kp_nam[1], 0)} triệu đồng năm 2022 "
        f"lên {so(kp_nam[4], 0)} triệu đồng năm 2025. Tuy vậy, mức cấp phân tán rất mạnh: trung vị chỉ "
        f"{so(kp_trung_vi, 1)} triệu đồng một đề tài, trong khi đề tài cao nhất được cấp 80 triệu đồng.",
        f"Phát hiện quan trọng nhất đến từ việc đối chiếu hình thức kinh phí với sản phẩm. Cả 11 đề tài có sản phẩm đủ "
        f"điều kiện xác lập quyền đều thuộc nhóm {len(dt_co_tien)} đề tài được cấp kinh phí bằng tiền; không đề tài nào "
        f"trong nhóm chỉ quy đổi giờ hoặc tự tìm tài trợ có sản phẩm như vậy. Mười một đề tài này sử dụng "
        f"{so(kp_du_dk, 2)} triệu đồng, tương đương {pt(kp_du_dk / kp_tong)} tổng kinh phí. Kinh phí bằng tiền vì vậy là "
        f"điều kiện cần để hình thành sản phẩm có thể bảo hộ, song kinh phí đó chỉ đủ để tạo ra sản phẩm mà không có "
        f"khoản nào dành cho bước đăng ký.",
    ],
))

# H2.11 ---------------------------------------------------------------------
HINH.append(dict(
    tieu_de="Giờ nghiên cứu quy đổi và mức thưởng bằng tiền theo loại sản phẩm khoa học",
    nguon="Nguồn: Nhóm nghiên cứu tổng hợp từ Bảng 6 và Bảng 7, Quy chế chi tiêu nội bộ Trường Đại học Thành Đô "
          "ban hành ngày 01 tháng 8 năm 2026. Sáng chế chuẩn quốc tế là bằng độc quyền theo chuẩn Mỹ, châu Âu, Đông Bắc "
          "Á; sáng chế chuẩn Việt Nam là bằng độc quyền theo chuẩn Việt Nam, Trung Quốc, ASEAN. Mức thưởng là tổng "
          "tiền thưởng cho một bài báo.",
    cot=["Loại sản phẩm", "Giờ nghiên cứu quy đổi", "Thưởng bằng tiền (triệu đồng)"],
    dong=[[k, g, t] for k, g, t in KHUYEN_KHICH],
    cao=10.0,
    dinh_dang=[None, "0", "0"],
    bieu_do=dict(loai="column", khoang_cach=60,
                 chuoi=[dict(cot=1, mau=XANH, nhan=True, mau_diem=[TIM, TIM, TIM, XANH, XANH, XANH, XANH, XANH]),
                        dict(cot=2, kieu="line", mau=CAM, truc_phu=True, nhan=True, duong=False,
                             vi_tri_nhan="above", an_nhan_0=True)],
                 truc_y="Giờ nghiên cứu quy đổi", truc_y2="Thưởng bằng tiền (triệu đồng)", dd_y2="0", max_y2=25,
                 max_y=900),
    sau_doan="Tuy nhiên, quy chế đặt điều kiện phải có chứng nhận",
    binh_luan=[
        "Hình 2.11 đặt hai công cụ khuyến khích cạnh nhau. Xét riêng giờ quy đổi, văn bằng sáng chế được định giá cao: "
        "một bằng độc quyền sáng chế chuẩn Việt Nam được quy đổi 360 giờ, cao hơn một bài báo WoS hạng Q1 với 300 giờ. "
        "Tuy nhiên, các điểm đánh dấu cho thấy chỉ bài báo quốc tế mới được thưởng bằng tiền, từ 10 đến 20 triệu đồng "
        "một bài, còn sáng chế và giải pháp hữu ích không có khoản thưởng tương ứng.",
        "Kết hợp với yếu tố thời gian, cấu trúc khuyến khích này nghiêng hẳn về công bố. Một bài báo Q1 mang lại 300 giờ "
        "và 20 triệu đồng ngay khi bài được đăng. Một sáng chế chuẩn Việt Nam mang lại 360 giờ, "
        "không có tiền thưởng và chỉ được ghi nhận sau khi có văn bằng, tức từ hai đến ba năm sau khi nộp đơn. Với "
        "cùng một kết quả nghiên cứu, lựa chọn hợp lý của giảng viên là công bố trước, và việc công bố trước khi nộp đơn "
        "chính là hành vi làm mất tính mới của sáng chế sau thời hạn mười hai tháng. Cấu trúc này phù hợp với hiện "
        "tượng công bố tăng mạnh tại Hình 2.4 trong khi đơn đăng ký hầu như vắng mặt.",
    ],
))

# H2.12 ---------------------------------------------------------------------
loai_ts = ["Quyền tác giả", "Nhãn hiệu", "Kiểu dáng công nghiệp", "Sáng chế"]
nhom_tt = [DA_CAP, CHUA_XM, DA_NOP]
dem_tt = {n: [sum(1 for t in TSTT_KY if t[0] == l and t[4] == n) for l in loai_ts] for n in nhom_tt}
da_cap = dem_tt[DA_CAP]
HINH.append(dict(
    tieu_de="Hồ sơ tài sản trí tuệ của Nhà trường trong kỳ 2021 - 2025 theo loại hình và tình trạng pháp lý",
    nguon="Nguồn: Nhóm nghiên cứu tổng hợp từ Bảng 2.6. Không gồm đơn sáng chế nộp năm 2026. Nhóm chưa xác minh là "
          "5 kiểu dáng công nghiệp có số hiệu văn bằng tại bảng thống kê văn bằng nhưng được ghi là chờ cấp bằng tại "
          "bảng theo dõi đơn.",
    cot=["Loại hình", DA_CAP, "Có số hiệu văn bằng, trạng thái chưa xác minh", DA_NOP],
    dong=[[loai_ts[i]] + [dem_tt[n][i] for n in nhom_tt] for i in range(4)],
    dinh_dang=[None, "0", "0", "0"],
    bieu_do=dict(loai="bar", xep_chong=True, khoang_cach=55, dao_truc=True,
                 chuoi=[dict(cot=1, mau=XANH, nhan=True, an_nhan_0=True),
                        dict(cot=2, mau=XAM, nhan=True, an_nhan_0=True),
                        dict(cot=3, mau=VANG, nhan=True, an_nhan_0=True)],
                 truc_y="Số hồ sơ"),
    sau_doan="",
    binh_luan=[],
))

# H2.13 ---------------------------------------------------------------------
HINH.append(dict(
    tieu_de="Tài sản trí tuệ của Nhà trường theo năm xác lập, nguồn hình thành và số lũy kế",
    nguon="Nguồn: Nhóm nghiên cứu tổng hợp từ Bảng 2.7. Năm là năm cấp văn bằng hoặc năm nộp đơn đối với hồ sơ "
          "đang xử lý.",
    cot=["Năm", "Thương hiệu", "Hợp tác doanh nghiệp", "Nghiên cứu", "Lũy kế"],
    dong=[[str(n)] + [dem_ts(n, g) for g in nguon_ts] + [luy_ke[i]] for i, n in enumerate(tstt_nam)],
    dinh_dang=[None, "0", "0", "0", "0"],
    bieu_do=dict(loai="column", xep_chong=True, khoang_cach=60,
                 chuoi=[dict(cot=1, mau=XANH), dict(cot=2, mau=CAM), dict(cot=3, mau=NGOC),
                        dict(cot=4, kieu="line", mau="#3B3A36", nhan=True, vi_tri_nhan="above")],
                 truc_y="Số tài sản"),
    sau_doan="Xét theo thời gian, tiến độ xác lập quyền có xu hướng tăng",
    binh_luan=[
        "Hình 2.13 cho thấy ba luồng hình thành tài sản có nhịp độ rất khác nhau. Luồng thương hiệu diễn ra rải rác nhưng "
        "đều đặn trong cả giai đoạn. Luồng hợp tác doanh nghiệp tạo ra một đợt tăng đột biến duy nhất vào năm 2024 với "
        "5 kiểu dáng công nghiệp, khiến đường lũy kế dốc lên từ 4 lên 9 tài sản chỉ trong một năm. Luồng nghiên cứu chỉ "
        "xuất hiện từ năm 2025 với mỗi năm một hồ sơ.",
        "Hình dạng đường lũy kế cho thấy danh mục tài sản của Nhà trường phát triển theo từng sự kiện chứ chưa theo một "
        "quy trình. Nếu loại bỏ đợt hợp tác năm 2024, tốc độ xác lập quyền chỉ vào khoảng một tài sản mỗi năm. Để luồng "
        "nghiên cứu trở thành luồng chủ đạo, Nhà trường cần cơ chế biến các trường hợp đơn lẻ của năm 2025 và 2026 thành "
        "hoạt động thường xuyên.",
    ],
))

# H2.14 ---------------------------------------------------------------------
nhom_q = ["Giải pháp hữu ích", "Giải pháp hữu ích hoặc sáng chế", "Sáng chế", "Sưu tập dữ liệu",
          "Kiểu dáng công nghiệp hoặc giải pháp hữu ích"]
dem_q = [sum(1 for d in dt_du_dk if d[7] == n) for n in nhom_q]
shcn = sum(dem_q[i] for i in (0, 1, 2, 4))
HINH.append(dict(
    tieu_de="Sản phẩm đề tài cấp cơ sở có tiềm năng tạo lập tài sản trí tuệ theo nhóm quyền dự kiến",
    nguon="Nguồn: Nhóm nghiên cứu tổng hợp từ Bảng 2.8.",
    cot=["Nhóm quyền dự kiến", "Số sản phẩm"],
    dong=[[nhom_q[i], dem_q[i]] for i in range(len(nhom_q))],
    dinh_dang=[None, "0"],
    bieu_do=dict(loai="bar", khoang_cach=60, dao_truc=True,
                 chuoi=[dict(cot=1, mau=XANH, nhan=True, mau_diem=[XANH, XANH, XANH, NGOC, XANH])],
                 truc_y="Số sản phẩm", an_chu_giai=True),
    sau_doan="Nguồn: Tổng hợp của nhóm nghiên cứu từ danh mục đề tài khoa học công nghệ cấp cơ sở",
    binh_luan=[
        f"Hình 2.14 cho thấy {shcn} trên 11 sản phẩm thuộc nhóm sở hữu công nghiệp, trong đó giải pháp hữu ích là "
        f"hình thức phù hợp với {dem_q[0] + dem_q[1] + dem_q[4]} sản phẩm. Giải pháp hữu ích không đòi hỏi trình độ "
        f"sáng tạo như sáng chế, thời gian bảo hộ ngắn hơn và phù hợp với quy mô kinh phí của đề tài cấp cơ sở. Loại hình "
        f"này đã được liệt kê tại Điều 34 Quyết định 213 và Điều 3 Quy chế quản trị tài sản trí tuệ năm 2024, song chưa "
        f"sản phẩm nào được nhận diện và nộp đơn theo hình thức này. Về "
        f"nguồn gốc, {len(dt_duoc)} trên 11 sản phẩm liên quan đến lĩnh vực dược, cho thấy tiềm năng "
        f"sở hữu công nghiệp của Nhà trường tập trung ở một lĩnh vực và có thể được quản lý có trọng tâm.",
    ],
))

# H2.15 ---------------------------------------------------------------------
HINH.append(dict(
    tieu_de="Chuỗi chuyển hóa từ đề tài cấp cơ sở sang đơn đăng ký sở hữu công nghiệp",
    nguon="Nguồn: Nhóm nghiên cứu tính toán từ danh mục đề tài cấp cơ sở và Bảng 2.7. Chuỗi chỉ xét sở hữu công "
          "nghiệp, không gồm 2 sản phẩm thuộc quyền tác giả. Cột thứ hai gồm các đề tài mang mã số từ năm 2021 đến năm "
          "2024, đều nghiệm thu trước ngày 31 tháng 5 năm 2025.",
    cot=["Bậc chuyển hóa", "Toàn bộ 38 đề tài, mã số 2021 - 2025", "Đề tài mã số 2021 - 2024"],
    dong=[["Đề tài đã nghiệm thu", len(DE_TAI), len(dt_den_2024)],
          ["Có sản phẩm tiềm năng sở hữu công nghiệp", len(dt_shcn), len(dt_shcn_den_2024)],
          ["Đã nộp đơn đăng ký", len(nop_don), sum(1 for d in nop_don if d[1] <= 2024)]],
    dinh_dang=[None, "0", "0"],
    bieu_do=dict(loai="bar", khoang_cach=45, dao_truc=True,
                 chuoi=[dict(cot=1, mau=XANH, nhan=True), dict(cot=2, mau=NGOC, nhan=True)],
                 truc_y="Số đề tài"),
    sau_doan="",
    binh_luan=[],
))

# H2.16 ---------------------------------------------------------------------
# Chỉ tiêu 1.11 chưa xác định được kết quả (trạng thái 5 kiểu dáng chưa thống nhất) nên không đưa vào biểu đồ.
kh_rows, kh_chua_xd = [], []
for muc, ten, nhom, k1, k2, t1, t2 in KE_HOACH:
    if t1 is None:
        kh_chua_xd.append([f"{muc}. {ten}", nhom, k1 + k2])
        continue
    kh_rows.append([f"{muc}. {ten}", nhom, k1 + k2, t1 + t2, (t1 + t2) / (k1 + k2)])
kh = {r[0].split(". ", 1)[1]: list(r) for r in kh_rows}
mau_nhom = {"Công bố": XANH, "Đề tài và học liệu": NGOC, "Tài sản trí tuệ": CAM}
HINH.append(dict(
    tieu_de="Tỷ lệ thực hiện so với chỉ tiêu Kế hoạch 07/KH-ĐHTĐ, cộng dồn hai năm 2024 - 2025",
    nguon="Nguồn: Nhóm nghiên cứu tính toán từ mục 2.2.4 Kế hoạch số 07/KH-ĐHTĐ ngày 01 tháng 7 năm 2024 và các danh mục "
          "thống kê của Phòng Khoa học Công nghệ; tên chỉ tiêu ghi theo nguyên văn Kế hoạch, kèm số mục. Mục 1.3 và "
          "1.4 được cộng chung do danh mục bài báo trong nước gồm cả tạp chí của Trường. Đề tài xếp theo năm ghi trong "
          "mã số hoặc năm phê duyệt kinh phí. Không gồm mục 1.5 do danh mục tham luận hội thảo quốc gia không có năm, "
          "và mục 1.11 do trạng thái pháp lý của 5 kiểu dáng công nghiệp chưa thống nhất. Màu xanh dương: công bố; xanh "
          "lục: đề tài và học liệu; cam: tài sản trí tuệ.",
    cot=["Chỉ tiêu", "Nhóm", "Kế hoạch 2024 - 2025", "Thực hiện 2024 - 2025", "Tỷ lệ thực hiện"],
    dong=kh_rows,
    dinh_dang=[None, None, "0", "0", "0%"],
    bieu_do=dict(loai="bar", khoang_cach=45, dao_truc=True,
                 chuoi=[dict(cot=4, mau=XANH, nhan=True, mau_diem=[mau_nhom[r[1]] for r in kh_rows])],
                 truc_y="Tỷ lệ thực hiện so với kế hoạch", dd_y="0%", an_chu_giai=True),
    sau_doan="",
    binh_luan=[],
))

for i, h in enumerate(HINH, 1):
    h["id"] = f"H2.{i}"
    h["so"] = i

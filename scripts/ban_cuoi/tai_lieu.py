# -*- coding: utf-8 -*-
"""Danh mục tài liệu tham khảo theo APA 7, dùng chung cho báo cáo tổng kết và bài báo.

Mỗi mục gồm: khóa trích dẫn trong bài (chuỗi xuất hiện trong văn bản), nhóm và
chuỗi APA có markup *nghiêng*. Chỉ ghi tài liệu đã đối chiếu được với bản gốc
trong kho (thư mục VBPL, "Co so ly luan", "Tai lieu thanh do") hoặc với trang
của nhà xuất bản; tài liệu không truy xuất được đã bị lược khỏi văn bản.

Ghi chú kiểm chứng:
- O'Dwyer, Filieri và O'Malley (2023): tệp Zotero ghi sai tác giả là "Rialti và cộng
  sự, 2022"; đã sửa theo DOI 10.1007/s10961-022-09932-2.
- Siegel và cộng sự (2007): DOI đúng là 10.1093/oxrep/grm036.
- Milliken và Allen (2013): không truy xuất được nhà xuất bản, đã lược trích dẫn.
- Điều lệ Quỹ Học bổng sau tiến sĩ Ngô Xuân Độ: tệp gốc không có trong kho; nội dung Điều 9
  lấy từ bản thảo của nhóm, cần bổ sung số hiệu, ngày ban hành khi có tệp.
- Nguyễn (2025), Võ (2025): lấy từ tệp Zotero của nhóm, có DOI nhưng chưa mở được
  trang tạp chí để đối chiếu số tập, trang của Võ (2025); cần xác minh trước khi nộp.
"""
import re

# nhóm: "vb" văn bản pháp luật và văn bản của Nhà trường; "vn" tài liệu tiếng Việt; "nn" tài liệu tiếng nước ngoài
TAI_LIEU = [
    # --- Văn bản quy phạm pháp luật, văn bản chỉ đạo, văn bản nội bộ ---------------------------
    ("Kết luận số 51-KL/TW", "vb",
     "Bộ Chính trị. (2026). *Kết luận số 51-KL/TW ngày 17 tháng 6 năm 2026 về đẩy mạnh công tác sở hữu trí tuệ phục "
     "vụ phát triển kinh tế - xã hội trong tình hình mới*."),
    ("Thông tư số 01/2024/TT-BGDĐT", "vb",
     "Bộ Giáo dục và Đào tạo. (2024). *Thông tư số 01/2024/TT-BGDĐT ngày 05 tháng 02 năm 2024 ban hành Chuẩn cơ sở "
     "giáo dục đại học*."),
    ("Thông tư số 83/2026/TT-BGDĐT", "vb",
     "Bộ Giáo dục và Đào tạo. (2026). *Thông tư số 83/2026/TT-BGDĐT ngày 30 tháng 9 năm 2026 quy định Chuẩn cơ sở "
     "giáo dục đại học*."),
    ("Nghị định số 17/2023/NĐ-CP", "vb",
     "Chính phủ. (2023). *Nghị định số 17/2023/NĐ-CP ngày 26 tháng 4 năm 2023 quy định chi tiết một số điều và biện "
     "pháp thi hành Luật Sở hữu trí tuệ về quyền tác giả, quyền liên quan*."),
    ("Nghị định số 134/2026/NĐ-CP", "vb",
     "Chính phủ. (2026). *Nghị định số 134/2026/NĐ-CP ngày 06 tháng 4 năm 2026 sửa đổi, bổ sung một số điều của Nghị "
     "định số 17/2023/NĐ-CP quy định chi tiết một số điều và biện pháp thi hành Luật Sở hữu trí tuệ về quyền tác giả, "
     "quyền liên quan*."),
    ("Luật số 36/2009/QH12", "vb",
     "Quốc hội. (2009). *Luật số 36/2009/QH12 ngày 19 tháng 6 năm 2009 sửa đổi, bổ sung một số điều của Luật Sở hữu "
     "trí tuệ*."),
    ("Luật số 07/2017/QH14", "vb",
     "Quốc hội. (2017). *Luật Chuyển giao công nghệ số 07/2017/QH14 ngày 19 tháng 6 năm 2017*."),
    ("Luật số 93/2025/QH15", "vb",
     "Quốc hội. (2025a). *Luật Khoa học, công nghệ và đổi mới sáng tạo số 93/2025/QH15 ngày 27 tháng 6 năm 2025*."),
    ("Luật số 123/2025/QH15", "vb",
     "Quốc hội. (2025b). *Luật số 123/2025/QH15 ngày 10 tháng 12 năm 2025 sửa đổi, bổ sung một số điều của Luật Giáo "
     "dục*."),
    ("Luật số 125/2025/QH15", "vb",
     "Quốc hội. (2025c). *Luật Giáo dục đại học số 125/2025/QH15 ngày 10 tháng 12 năm 2025*."),
    ("Luật số 131/2025/QH15", "vb",
     "Quốc hội. (2025d). *Luật số 131/2025/QH15 ngày 10 tháng 12 năm 2025 sửa đổi, bổ sung một số điều của Luật Sở "
     "hữu trí tuệ*."),
    ("Quyết định số 1068/QĐ-TTg", "vb",
     "Thủ tướng Chính phủ. (2019). *Quyết định số 1068/QĐ-TTg ngày 22 tháng 8 năm 2019 phê duyệt Chiến lược sở hữu "
     "trí tuệ đến năm 2030*."),
    ("Chỉ thị số 02/CT-TTg", "vb",
     "Thủ tướng Chính phủ. (2026a). *Chỉ thị số 02/CT-TTg ngày 30 tháng 01 năm 2026 về tăng cường thực thi quyền sở "
     "hữu trí tuệ*."),
    ("Quyết định số 1624/QĐ-TTg", "vb",
     "Thủ tướng Chính phủ. (2026b). *Quyết định số 1624/QĐ-TTg ngày 21 tháng 8 năm 2026 sửa đổi, bổ sung một số điều "
     "của Quyết định số 1068/QĐ-TTg ngày 22 tháng 8 năm 2019 phê duyệt Chiến lược sở hữu trí tuệ đến năm 2030*."),
    ("Quyết định số 213/QĐ-ĐHTĐ", "vb",
     "Trường Đại học Thành Đô. (2021). *Quyết định số 213/QĐ-ĐHTĐ ngày 28 tháng 12 năm 2021 ban hành Quy chế hoạt "
     "động khoa học công nghệ Trường Đại học Thành Đô*."),
    ("Kế hoạch số 07/KH-ĐHTĐ", "vb",
     "Trường Đại học Thành Đô. (2024a). *Kế hoạch số 07/KH-ĐHTĐ ngày 01 tháng 7 năm 2024 về hoạt động khoa học "
     "công nghệ giai đoạn 2024 - 2028*."),
    ("Quyết định số 217/QĐ-ĐHTĐ", "vb",
     "Trường Đại học Thành Đô. (2024b). *Quyết định số 217/QĐ-ĐHTĐ ngày 21 tháng 11 năm 2024 ban hành Quy chế quản "
     "trị tài sản trí tuệ tại Trường Đại học Thành Đô*."),
    ("Điều lệ Quỹ", "vb",
     "Trường Đại học Thành Đô. (2025). *Điều lệ Quỹ Học bổng sau tiến sĩ Ngô Xuân Độ*."),
    ("Quy chế chi tiêu nội bộ", "vb",
     "Trường Đại học Thành Đô. (2026). *Quy chế chi tiêu nội bộ Trường Đại học Thành Đô* (Ban hành theo Nghị quyết số "
     "02/NQ-HĐT-ĐHTĐ ngày 01 tháng 6 năm 2026)."),
    ("Văn phòng Quốc hội, 2026", "vb",
     "Văn phòng Quốc hội. (2026). *Văn bản hợp nhất số 67/VBHN-VPQH ngày 23 tháng 3 năm 2026 hợp nhất Luật Sở hữu trí "
     "tuệ*."),
    # --- Tài liệu tiếng Việt ------------------------------------------------------------------
    ("Cục Sở hữu trí tuệ (n.d.)", "vn",
     "Cục Sở hữu trí tuệ. (n.d.). *Tài liệu tập huấn về sở hữu trí tuệ dành cho cán bộ các trường đại học, viện nghiên "
     "cứu*. Truy cập ngày 30 tháng 9 năm 2026, từ https://shtt.hcmulaw.edu.vn/sach-dien-tu/"
     "bo-tai-lieu-tap-huan-ve-so-huu-tri-tue-cua-cuc-so-huu-tri-tue-327.html"),
    ("Nguyễn, 2025", "vn",
     "Nguyễn, M. H. T. (2025). Quản trị tài sản trí tuệ tại các cơ sở giáo dục đại học: Cơ hội và thách thức trong bối "
     "cảnh cuộc cách mạng công nghiệp 4.0. *Tạp chí Khoa học Trường Đại học Sư phạm Thành phố Hồ Chí Minh, 22*(1), "
     "123-131. https://doi.org/10.54607/hcmue.js.22.1.4287(2025)"),
    ("Tổ chức Sở hữu trí tuệ thế giới", "vn",
     "Tổ chức Sở hữu trí tuệ thế giới. (2020). *What is intellectual property?* (WIPO Publication No. 450). World "
     "Intellectual Property Organization."),
    ("Võ, 2025", "vn",
     "Võ, N. H. P. (2025). Quyền của chủ thể không giữ quyền tài sản đối với các tác phẩm hình thành trong nhà trường, "
     "kinh nghiệm quốc tế để hoàn thiện chính sách sở hữu trí tuệ của các trường đại học tại Việt Nam. *Tạp chí Khoa "
     "học Đại học Mở Thành phố Hồ Chí Minh*. https://doi.org/10.59266/houjs.2025.606"),
    # --- Tài liệu tiếng nước ngoài ------------------------------------------------------------
    ("Bradley et al., 2013", "nn",
     "Bradley, S. R., Hayter, C. S., & Link, A. N. (2013). Models and methods of university technology transfer. "
     "*Foundations and Trends in Entrepreneurship, 9*(6), 571-650. https://doi.org/10.1561/0300000048"),
    ("Etzkowitz, 2003", "nn",
     "Etzkowitz, H. (2003). Research groups as 'quasi-firms': The invention of the entrepreneurial university. "
     "*Research Policy, 32*(1), 109-121. https://doi.org/10.1016/S0048-7333(02)00009-4"),
    ("Fisher", "nn",
     "Fisher, W. (2001). Theories of intellectual property. In S. R. Munzer (Ed.), *New essays in the legal and "
     "political theory of property* (pp. 168-199). Cambridge University Press."),
    ("Goldfarb & Henrekson, 2003", "nn",
     "Goldfarb, B., & Henrekson, M. (2003). Bottom-up versus top-down policies towards the commercialization of "
     "university intellectual property. *Research Policy, 32*(4), 639-658. "
     "https://doi.org/10.1016/S0048-7333(02)00034-3"),
    ("Guan", "nn",
     "Guan, W. (2014). Intellectual property: Concept, history, and contentions. In *Intellectual property theory and "
     "practice: A critical examination of China's TRIPS compliance and beyond* (pp. 1-10). Springer. "
     "https://doi.org/10.1007/978-3-642-55265-6_1"),
    ("Holgersson & Aaboen, 2019", "nn",
     "Holgersson, M., & Aaboen, L. (2019). A literature review of intellectual property management in technology "
     "transfer offices: From appropriation to utilization. *Technology in Society, 59*, Article 101132. "
     "https://doi.org/10.1016/j.techsoc.2019.04.008"),
    ("Maresova et al., 2019", "nn",
     "Maresova, P., Stemberkova, R., & Fadeyi, O. (2019). Models, processes, and roles of universities in technology "
     "transfer management: A systematic review. *Administrative Sciences, 9*(3), Article 67. "
     "https://doi.org/10.3390/admsci9030067"),
    ("O’Dwyer et al., 2023", "nn",
     "O’Dwyer, M., Filieri, R., & O’Malley, L. (2023). Establishing successful university-industry collaborations: "
     "Barriers and enablers deconstructed. *The Journal of Technology Transfer, 48*(3), 900-931. "
     "https://doi.org/10.1007/s10961-022-09932-2"),
    ("Perkmann et al., 2013", "nn",
     "Perkmann, M., Tartari, V., McKelvey, M., Autio, E., Broström, A., D’Este, P., Fini, R., Geuna, A., Grimaldi, R., "
     "Hughes, A., Krabel, S., Kitson, M., Llerena, P., Lissoni, F., Salter, A., & Sobrero, M. (2013). Academic "
     "engagement and commercialisation: A review of the literature on university-industry relations. *Research "
     "Policy, 42*(2), 423-442. https://doi.org/10.1016/j.respol.2012.09.007"),
    ("Rocha et al., 2023", "nn",
     "Rocha, A., Cruz-Cunha, M. M., & Romero, F. (2023). University technology transfer: Assessment of invention "
     "disclosures by technology transfer offices. *International Journal of Entrepreneurship and Innovation "
     "Management, 27*(1/2), 119-136."),
    ("Shane, 2004", "nn",
     "Shane, S. (2004). Encouraging university entrepreneurship? The effect of the Bayh-Dole Act on university "
     "patenting in the United States. *Journal of Business Venturing, 19*(1), 127-151. "
     "https://doi.org/10.1016/S0883-9026(02)00114-3"),
    ("Siegel et al., 2007", "nn",
     "Siegel, D. S., Veugelers, R., & Wright, M. (2007). Technology transfer offices and commercialization of "
     "university intellectual property: Performance and policy implications. *Oxford Review of Economic Policy, "
     "23*(4), 640-660. https://doi.org/10.1093/oxrep/grm036"),
    ("Teece, 2018", "nn",
     "Teece, D. J. (2018). Profiting from innovation in the digital economy: Enabling technologies, standards, and "
     "licensing models in the wireless world. *Research Policy, 47*(8), 1367-1387. "
     "https://doi.org/10.1016/j.respol.2017.01.015"),
    ("Thursby & Kemp, 2002", "nn",
     "Thursby, J. G., & Kemp, S. (2002). Growth and productive efficiency of university intellectual property "
     "licensing. *Research Policy, 31*(1), 109-124. https://doi.org/10.1016/S0048-7333(00)00160-8"),
]

NHOM = [("vb", "A. Văn bản pháp luật, văn bản chỉ đạo và văn bản của Trường Đại học Thành Đô"),
        ("vn", "B. Tài liệu tiếng Việt"),
        ("nn", "C. Tài liệu tiếng nước ngoài")]


def _mau(muc):
    """Mẫu nhận diện trích dẫn: văn bản pháp luật theo số hiệu; tài liệu khác theo họ tác giả và năm,
    chấp nhận cả dạng (Tác giả, năm), Tác giả (năm), Tác giả và cộng sự (năm)."""
    khoa, nhom, apa = muc
    nam = re.search(r"\((\d{4}[a-z]?|n\.d\.)\)", apa).group(1)
    if nhom == "vb":
        so = re.search(r"\d+[-/][\w/\-]+", khoa)
        tac_gia = apa.split(". (")[0]
        return (re.escape(so.group(0)) if so else re.escape(khoa)) + "|" + re.escape(tac_gia) + r",\s*" + \
            re.escape(nam[:4])
    ho = re.split(r",| et al\.| &| \(", khoa)[0].strip()
    return re.escape(ho) + r"[^()]{0,40}?[(,]\s*" + re.escape(nam)


def duoc_trich(van_ban):
    """Trả về các mục được trích trong văn bản, giữ thứ tự danh mục."""
    return [m for m in TAI_LIEU if re.search(_mau(m), van_ban)]


def trich_dan_mo_coi(van_ban):
    """Trích dẫn dạng (Tác giả, năm) trong bài không khớp mục nào của danh mục."""
    kq = []
    for m in re.finditer(r"\(([^()]*?(?:\d{4}[a-z]?|n\.d\.))\)", van_ban):
        for phan in m.group(1).split(";"):
            phan = phan.strip()
            if not re.search(r"(\d{4}[a-z]?|n\.d\.)$", phan) or re.search(r"(năm|ngày|tháng|số)\s", phan):
                continue
            if re.fullmatch(r"\d{4}[a-z]?|n\.d\.", phan):  # trích dẫn dạng kể: Tác giả (năm)
                truoc = van_ban[max(0, m.start() - 60):m.start()]
                if not any(re.sub(r"\W.*", "", k) in truoc for k, _, _ in TAI_LIEU):
                    kq.append(truoc[-30:] + "(" + phan + ")")
                continue
            if not any(k.split(" (")[0].split(",")[0] in phan for k, _, _ in TAI_LIEU):
                kq.append(phan)
    return kq

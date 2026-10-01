# BÁO CÁO RÀ SOÁT CHƯƠNG 2, CHƯƠNG 3 VÀ BÀI BÁO THEO DỮ LIỆU GỐC

Ngày rà soát: 30/9/2026. Căn cứ đối chiếu là các tệp trong thư mục `Tai lieu thanh do` và `VBPL` do chủ nhiệm đề tài cung cấp.

## 1. Tệp kết quả

| Tệp | Nội dung | Cách dùng |
|---|---|---|
| `Chuong_2_Thuc_trang_rut_gon.docx` | Chương 2 bản rút gọn, khoảng 21 trang, 7 bảng, 9 biểu đồ Excel nhúng | Bản sạch, thay cho bản 43 trang trước; nhật ký sửa nằm trong sheet `Nhat_ky_chuan_hoa` của tệp Excel |
| `Du_lieu_bieu_do_Chuong_2.xlsx` | Dữ liệu, biểu đồ, nhật ký chuẩn hóa, danh mục đề tài và tài sản trí tuệ | Sửa số liệu tại đây hoặc trong `scripts/chuong2/du_lieu.py` |
| `Chuong_3_He_thong_giai_phap_ra_soat.docx` | Chương 3 có 47 chỉnh sửa ở chế độ theo dõi thay đổi | Mở bằng Word, thẻ Review, chấp nhận hoặc từ chối từng chỗ |
| `Bai_bao_khoa_hoc_Tap_chi_DHTD_ra_soat.docx` | Bài báo có 24 chỉnh sửa ở chế độ theo dõi thay đổi | Như trên |

Các tệp gốc không bị sửa.

## Cập nhật lượt 5: bản cuối gộp và bài báo (thư mục `Ban_cuoi`)

**Tệp kết quả**
- `Bao_cao_tong_ket_de_tai.docx`: báo cáo tổng kết gộp một bản gồm trang bìa, mục lục, danh mục bảng, danh mục hình, Mở đầu, ba chương, Kết luận và kiến nghị, Tài liệu tham khảo (38 mục, APA 7). Mục lục và danh mục là trường tự động của Word, đã có sẵn số trang; Word hỏi cập nhật trường khi mở tệp, chọn Yes để số trang khớp máy in.
- `Du_lieu_bieu_do.xlsx`: số liệu và biểu đồ gốc của 11 biểu đồ Chương 2.
- `Bai_bao_Tap_chi_NCKH_PT.docx`: bài báo đầu ra theo chuẩn JSRD, kèm Tờ khai minh bạch sử dụng AI và cam kết dữ liệu gốc (nhóm tác giả tự điền Phần I).
- `so_do/`: 8 sơ đồ dùng trong báo cáo và bài báo.
- Các tệp chương rời đã gỡ; có thể dựng lại riêng bằng `scripts/ban_cuoi/chuong1.py`, `scripts/chuong2/build_gon.py`, `scripts/ban_cuoi/chuong3.py`. Dựng toàn bộ: `python3 scripts/ban_cuoi/dung_tat_ca.py`.

**Nội dung đã cập nhật**
- Quyết định số 217/QĐ-ĐHTĐ ghi thống nhất ngày 21 tháng 11 năm 2024.
- Luật số 93/2025/QH15 (tệp `VBPL/93_2025_QH15_581164.docx`) được đối chiếu và đưa vào Mục 1.3.2, Chương 2, Chương 3: khoản 2 Điều 25 (tự động giao quyền), Điều 27 (tự quyết thương mại hóa), khoản 2 và điểm a khoản 3 Điều 28 (thưởng tác giả tối thiểu 30% lợi nhuận đối với kết quả sử dụng ngân sách nhà nước), Điều 37, điểm b khoản 2 Điều 66, Điều 71, hiệu lực từ 01/10/2025.
- Bốn điểm diễn đạt theo góp ý: độ trễ thể chế trước sự thay đổi dồn dập của pháp luật giai đoạn 2025 - 2026 và tầm nhìn sớm của Nhà trường; tiềm năng rất lớn của khối Y - Dược với khoảng trống kỹ thuật là thiếu biểu mẫu rà soát tại thời điểm nghiệm thu; 6 tài sản đồng sở hữu là thành công nổi bật của chiến lược hợp tác doanh nghiệp mà Ban Giám hiệu đã dày công kết nối; câu về quy trình nghiệm thu tại Mục 2.5.2 thay bằng câu được đề nghị.
- Bổ sung Mở đầu (7 mục) và Kết luận, kiến nghị (3 mục) theo đề cương báo cáo tổng kết.
- Sơ đồ thay cho mô tả dài: chu trình bốn giai đoạn, sáu nhóm yếu tố, khung phân tích (Chương 1); phân công bốn đầu mối, chuỗi nguyên nhân (Chương 2); phối hợp liên phòng ban, quy trình 8 khâu, lộ trình 2026 - 2030 (Chương 3).

**Trích dẫn và tài liệu tham khảo**
- Toàn bộ trích dẫn chuyển sang APA 7: (Tác giả, năm), "&" với hai tác giả, "et al." từ ba tác giả; tài liệu tiếng Việt ghi họ tác giả. Danh mục chia ba nhóm: văn bản pháp luật và văn bản Nhà trường; tài liệu tiếng Việt; tài liệu tiếng nước ngoài. Danh mục được sinh tự động từ `scripts/ban_cuoi/tai_lieu.py` và chỉ gồm tài liệu có trích dẫn trong bài.
- Sửa sai lệch: "Rialti và cộng sự, 2022" trong tệp Zotero là sai tác giả; bài có DOI 10.1007/s10961-022-09932-2 là của O'Dwyer, Filieri và O'Malley (2023), The Journal of Technology Transfer, 48(3), 900-931. DOI của Siegel và cộng sự (2007) là 10.1093/oxrep/grm036. Guan (2014) đã xác định đủ thông tin chương sách của Springer.
- Lược bỏ trích dẫn không truy xuất được: Milliken và Allen (2013).
- Bổ sung tài liệu quốc tế gần đây đã kiểm chứng: Holgersson và Aaboen (2019), Maresova và cộng sự (2019), Rocha và cộng sự (2023).
- **Cần nhóm xác minh trước khi nộp:** Võ (2025) chưa có số tập, trang; Nguyễn (2025) và Võ (2025) lấy từ tệp Zotero, có DOI nhưng chưa mở được trang tạp chí để đối chiếu; Điều lệ Quỹ Học bổng sau tiến sĩ Ngô Xuân Độ chưa có tệp gốc trong kho nên danh mục chưa ghi số hiệu, ngày ban hành; Cục Sở hữu trí tuệ (n.d.) là tài liệu không ghi năm.

**Bài báo: phương án chính (cập nhật theo góp ý chủ nhiệm đề tài)**
- `Bai_bao_Tap_chi_NCKH_PT.docx`: "Khung đánh giá hiệu quả quản lý quyền sở hữu trí tuệ tại trường đại học định hướng ứng dụng: Tiếp cận chuỗi kết quả". Bài lý luận phát triển từ Mục 1.4: bộ 20 chỉ số (16 chỉ số đầu vào, quá trình, đầu ra, kết quả và 4 chỉ số chuyển hóa CH1 tỷ lệ nhận diện, CH2 tỷ lệ xác lập kịp thời, CH3 tỷ lệ khai thác, CH4 tỷ suất khai thác trên chi phí); phân loại nguồn dữ liệu A/B/C theo Thông tư 83/2026, Luật 93, Luật 125 (5/7/8 chỉ số); 5 mẫu hình chẩn đoán điểm nghẽn; nội dung tối thiểu của phiếu rà soát tại nghiệm thu; lộ trình 3 giai đoạn. Không dùng số liệu nội bộ. Tóm tắt 249 âm tiết, Abstract 211 từ; Đặt vấn đề 7,0%, Tổng quan 16,9%, Phương pháp 7,0%, Kết quả 47,2%, Bàn luận 16,3%, Kết luận 5,7%; 3 bảng, 2 hình, 20 tài liệu. Bổ sung 3 tài liệu đã kiểm chứng: W. K. Kellogg Foundation (2004), Finne et al. (2009, DOI 10.2777/49910), Campbell et al. (2020, DOI 10.2760/907762).
- Hai phương án trước được giữ làm dự phòng: `Bai_bao_phuong_an_chinh_sach.docx`, `Bai_bao_phuong_an_du_lieu_noi_bo.docx`.

**Bài báo: các phương án dự phòng**
- `Bai_bao_phuong_an_chinh_sach.docx` (phân tích chính sách): "Khoảng cách giữa sản phẩm khoa học và tài sản trí tuệ tại trường đại học tư thục: Phân tích chính sách trong bối cảnh pháp lý mới". Chỉ dùng 12 văn bản công khai (Luật Sở hữu trí tuệ hợp nhất, Luật 93, 131, 125, 07/2017; Nghị định 134; Quyết định 1068, 1624; Chỉ thị 02; Kết luận 51; Thông tư 01/2024, 83/2026), không công bố số liệu nội bộ của Nhà trường. Phát hiện mới có căn cứ văn bản: Thông tư 83/2026 nâng hệ số bằng độc quyền giải pháp hữu ích từ 1 lên 3 (cao hơn bài WoS, Scopus hệ số 2), xếp quy định về sở hữu trí tuệ, liêm chính vào nội dung quản trị bắt buộc chung. Đề xuất khung rà soát quy chế nội bộ 9 nội dung và mô hình cổng rà soát tại nghiệm thu (sơ đồ `so_do/cong_ra_soat.png`). Tóm tắt 248 âm tiết, Abstract 226 từ; Đặt vấn đề 6,9%, Tổng quan 16,2%, Phương pháp 8,9%, Kết quả 45,5%, Bàn luận 15,5%, Kết luận 7,0%; 3 bảng, 2 hình, 28 tài liệu.
- `Bai_bao_phuong_an_du_lieu_noi_bo.docx` (phương án dự phòng): bài dùng số liệu của Nhà trường, chỉ nộp khi Ban Giám hiệu đồng ý công bố số liệu bằng văn bản.
- Danh mục tài liệu của từng tệp sắp xếp theo chữ cái tiếng Việt, hậu tố năm (2025a, 2025b) đánh lại riêng cho từng danh mục.

**Bài báo phương án dự phòng**
- Tiêu đề: "Từ đề tài đến văn bằng: Chuỗi chuyển hóa tài sản trí tuệ tại Trường Đại học Thành Đô" (19 âm tiết). Tóm tắt 242 âm tiết, Abstract 238 từ, 5 từ khóa.
- Dung lượng thân bài 4.755 âm tiết: Đặt vấn đề 6,9%; Tổng quan 16,9%; Phương pháp 8,9%; Kết quả 45,0%; Bàn luận 16,6%; Kết luận 5,7%, đều trong khung JSRD.
- 4 hình, 1 bảng, 24 tài liệu tham khảo; tiểu mục Kết quả đặt tên theo phát hiện; mỗi khuyến nghị neo vào một phát hiện.

## Cập nhật lượt 4: bản cuối ba chương (thư mục `Ban_cuoi`)

| Tệp | Nội dung |
|---|---|
| `Ban_cuoi/Chuong_1_Co_so_ly_luan_va_phap_ly.docx` | Chương 1, khoảng 21 trang, 3 bảng |
| `Ban_cuoi/Chuong_2_Thuc_trang_quan_ly_quyen_SHTT.docx` | Chương 2, khoảng 22 trang, 7 bảng, 9 biểu đồ Excel nhúng, có tiểu kết |
| `Ban_cuoi/Chuong_3_He_thong_giai_phap.docx` | Chương 3, khoảng 21 trang, 4 bảng, có phân tích SWOT đầy đủ và tiểu kết |
| `Ban_cuoi/Du_lieu_bieu_do_Chuong_2.xlsx` | Dữ liệu và biểu đồ Chương 2 |

Ba tệp dùng chung khổ A4, lề, phông Times New Roman 13, giãn dòng và kiểu bảng; tiêu đề chương, mục có cấp đề mục để tạo mục lục tự động. Đây là bản sạch, không còn theo dõi thay đổi.

**Chương 1, sửa theo văn bản gốc:**
- Luật Giáo dục đại học năm 2012 được thay bằng Luật số 125/2025/QH15, có hiệu lực từ 01/01/2026 (Điều 45).
- Nguồn gốc các thay đổi của Luật Sở hữu trí tuệ ghi đúng theo chú thích Văn bản hợp nhất 67/VBHN-VPQH: Điều 86a, Điều 133a, khoản 2 Điều 135 bãi bỏ theo điểm h khoản 7 Điều 71 Luật 93/2025/QH15; khoản 1 Điều 135 sửa theo điểm b; điểm c khoản 1 Điều 86 do khoản 22 Điều 1 Luật 131/2025/QH15 bổ sung.
- Bỏ các khẳng định không đối chiếu được: Luật 93/2025/QH15 "quy định tối thiểu 30%" cho tác giả và "Điều 27, 28 Luật 93" về giao quyền sở hữu, vì thư mục VBPL không có toàn văn Luật 93. Nếu có điều khoản cụ thể, có thể bổ sung lại.
- Thông tư 01/2024/TT-BGDĐT được mô tả đúng: công thức Tiêu chí 6.2 có tính bằng giải pháp hữu ích và sáng chế; giảng viên toàn thời gian thay vì cơ hữu.
- Quyết định 1624/QĐ-TTg: bốn nhóm yêu cầu ghi đúng khoản, điểm; "đưa vào nội dung học bắt buộc" sửa thành "nghiên cứu đưa"; thí điểm định giá ít nhất 100 quyền sở hữu trí tuệ.
- Quyết định 213 ghi ngày 28/12/2021 (bản cũ ghi 27/5/2021); mức 50 triệu đồng trong Quy chế chi tiêu nội bộ là kinh phí đề tài, không phải trần thù lao; bổ sung Quyết định 217.
- Điều 60 khoản 4 mô tả đúng (bản cũ ghi thêm trường hợp triển lãm, không có trong điều luật hiện hành); Điều 39 được dẫn cho quyền tài sản của tổ chức giao nhiệm vụ.
- Hai sơ đồ ảnh cũ không khớp nội dung chữ (sơ đồ tiêu chí chia định lượng, định tính) được thay bằng Bảng 1.1 chu trình bốn giai đoạn và Bảng 1.2 bộ 16 tiêu chí, khớp với Hình 2.5 Chương 2. Khung phân tích thành Bảng 1.3, khớp cấu trúc Chương 2 và Chương 3 bản cuối.
- Bỏ ngoặc đơn giải thích "Living Lab".

**Chương 3:**
- Mục 3.1.2 mới "Phân tích điểm mạnh, điểm yếu, thời cơ và thách thức": Bảng 3.1 ma trận 5 điểm mạnh, 5 điểm yếu rút từ Chương 2; 5 thời cơ, 5 thách thức từ khung pháp lý và Thông tư 83/2026/TT-BGDĐT. Bảng 3.2 bốn nhóm phương án kết hợp, gắn với năm giải pháp và thứ tự ưu tiên.
- Nguyên tắc 3 và Giải pháp 2 ghi đúng phân công tại Điều 11 Quyết định 217 và bốn đầu mối như Chương 2.
- Giải pháp 1 bổ sung nội dung liêm chính khoa học, liêm chính học thuật; thời hạn ban hành quy chế trước 31/5/2027.
- Giải pháp 4 bỏ số liệu chưa có nguồn (phí 3 đến 5 triệu đồng, 18 đến 36 tháng), thay bằng số liệu Chương 2; bổ sung cơ chế chuyển tiếp với Quỹ Ngô Xuân Độ 5 tỷ đồng và thưởng tiền cho văn bằng.
- Lộ trình chia lại theo mốc: quý IV/2026 đến quý II/2027; quý III/2027 đến hết 2028; 2029 - 2030. Bộ chỉ số thêm chỉ tiêu danh mục số và tách văn bằng từ kết quả nghiên cứu. Bảng phối hợp bỏ chữ viết tắt.
- Bảng đánh số 3.1 đến 3.4, tiêu đề bảng đặt phía trên; thêm tiểu kết.

**Chương 2:** thêm tiểu kết; nội dung giữ như lượt 3.

**Dựng lại:** `python3 scripts/ban_cuoi/dung_tat_ca.py`.

## Cập nhật lượt 3: rút gọn Chương 2, Thông tư 83/2026, Quyết định 217

- **Chương 2 rút gọn** từ 43 trang xuống khoảng 21 trang (khoảng 7.000 từ ngoài bảng). Bỏ phần năng suất theo đơn vị, phân hạng tạp chí, tương quan bài báo và giáo trình, cơ cấu nhân sự có công bố, quy mô ba kênh tài trợ, tài sản theo năm; số liệu cốt lõi của các phần này được giữ trong lời văn. Gộp hoạt động bảo vệ quyền vào Mục 2.3.2, bỏ phân mục của Mục 2.4, chuyển hạn chế dữ liệu thành Mục 2.2.4. Bảng hợp đồng của Viện Nghiên cứu giáo dục và Chuyển giao tri thức chuyển thành lời văn.
- **Thông tư số 83/2026/TT-BGDĐT** (Chuẩn cơ sở giáo dục đại học, ngày 30/9/2026, hiệu lực 15/11/2026, thay Thông tư 01/2024/TT-BGDĐT) chưa áp dụng cho giai đoạn đánh giá nên **không đưa vào Chương 2**. Nội dung được đặt tại **Mục 3.1.2 mới "Thời cơ và thách thức" của Chương 3** (Mục 3.1.2 cũ thành 3.1.3):
  - thời cơ: bằng giải pháp hữu ích tăng từ hệ số 1 lên 3, sáng chế hệ số 5 trong chỉ số sản phẩm quy đổi trên giảng viên quy đổi; Bảng 6B tách riêng thu từ thương mại hóa, sở hữu trí tuệ; ba đề tài cấp quốc gia và Quỹ 5 tỷ đồng;
  - thách thức: tiêu chí 1.1 có nội dung bắt buộc về sở hữu trí tuệ, liêm chính khoa học, liêm chính học thuật mà Quyết định 213 và 217 chưa có phần liêm chính; tiêu chí 1.3 về dữ liệu trên HEMIS; ngưỡng công bố WoS, Scopus 0,3 trong khi ước tính năm 2025 khoảng 0,31 và văn bằng không được tính.
  - Lưu ý: công thức mục 6.2.1 Phụ lục II xếp giải pháp hữu ích vào nhóm hệ số 3, nhưng bảng tổng hợp cuối Phụ lục II không nêu loại này.
- **Quyết định 217/QĐ-ĐHTĐ** được xác nhận là văn bản ban hành Quy chế quản trị tài sản trí tuệ năm 2024. Chương 2, Chương 3 và bài báo đã dùng số hiệu này.
- **Ngân sách Quỹ Ngô Xuân Độ 5 tỷ đồng** được xác nhận; số liệu trong các tài liệu vẫn giữ 5 tỷ đồng giai đoạn 2025 - 2029.
- **Chương 3:** thêm Mục 3.1.2 "Thời cơ và thách thức"; Mục 3.2.1 nêu thiếu nội dung liêm chính; tham chiếu Hình 2.16 đổi thành Hình 2.9. **Bài báo:** thêm Thông tư 83/2026 vào giải pháp và tài liệu tham khảo.
- **Biểu đồ:** chú giải đặt phía trên, vùng vẽ có bố cục cố định nên không còn bị đè; màu chữ nhãn chọn tự động trắng hoặc sẫm theo độ tương phản với màu nền; nhãn danh mục hiện đầy đủ; tiêu đề hình đặt ở dòng chú thích Word thay vì trong biểu đồ.

## 2. Tài liệu gốc mới được khai thác

1. **Quy chế quản trị tài sản trí tuệ năm 2024** (`QC SHTT ĐHTĐ.docx`). Các bản trước của Chương 2 và Chương 3 chưa sử dụng văn bản này. Quy chế gồm 17 điều. Điều 3 đã liệt kê giải pháp hữu ích, cơ sở dữ liệu, bí mật thương mại, giáo trình điện tử. Điều 10 yêu cầu tác giả xin ý kiến Phòng Khoa học Công nghệ trước khi bộc lộ. Điều 11 giao Phòng Khoa học Công nghệ nhận diện, lập hồ sơ theo dõi tài sản trí tuệ, và giao Bộ phận Pháp chế thực hiện thủ tục xác lập quyền. Điều 13 giao Hiệu trưởng quyết định tỷ lệ phân chia khi không có thỏa thuận. **Bản được cung cấp chưa ghi số và ngày ban hành.**
2. **Kế hoạch số 07/KH-ĐHTĐ ngày 01/7/2024** về hoạt động khoa học công nghệ giai đoạn 2024 - 2028. Tệp là bản quét; nhóm đã nhận dạng chữ và đối chiếu lại với ảnh gốc. Số liệu Nhà trường tự tổng kết cho giai đoạn 2021 - 2023 khớp với số liệu đã chuẩn hóa: đề tài cấp cơ sở 6, 10, 7; bài báo trong nước 17, 42, 50; bài báo quốc tế 6, 13, 11; giáo trình 26, 31, 12. Chỉ tiêu 2024 - 2025 của Kế hoạch được dùng để dựng Hình 2.16.

## 3. Sai lệch chính đã sửa

### 3.1. Chương 2

| Nội dung | Bản trước | Theo tài liệu gốc |
|---|---|---|
| Điều 34 Quyết định 213 | Không liệt kê giải pháp hữu ích; dùng làm một nguyên nhân | Có liệt kê bằng độc quyền giải pháp hữu ích; sửa Bảng 2.6, Mục 2.3.1, 2.5.3, Hình 2.14 |
| Văn bản nội bộ về phân chia lợi ích | Ba văn bản, ba hoặc bốn công thức | Bốn văn bản, năm quy định; Bảng 2.4 bổ sung Điều 13 Quy chế 2024 |
| Độ phủ khai báo (Mục 2.1.5) | Khai 28 bài; 25 trên 28 tìm thấy; độ phủ 26,2% | 26 đề tài khai 32 công bố; độ phủ khoảng 29,9%. Tỷ lệ 89,3% bị bỏ vì không tái lập được từ tài liệu gốc |
| Mục 2.5.2, ý thứ sáu | Ba đề tài ghi rõ không có bài báo như cam kết | Danh mục gốc không có ghi chú này. Thay bằng: 7 đề tài chỉ có báo cáo tổng kết, 5 trong số đó xếp loại Tốt |
| Tổ chức bộ máy | Chức năng bị chia nhỏ, không có phân công | Quy chế 2024 đã phân công; khoảng cách nằm ở khâu thực thi |
| Lệ phí đăng ký | Không có dòng chi | Điều 35 giao Phòng Khoa học Công nghệ nộp lệ phí nhưng Điều 38 không có mục chi tương ứng |

Các sửa đổi của lượt trước vẫn giữ nguyên: 38 đề tài, 424,75 triệu đồng, 16 tham luận quốc tế, Bảng 2.1, Bảng 2.7, Bảng 2.8… Danh sách đầy đủ có trong sheet `Nhat_ky_chuan_hoa`.

### 3.2. Chương 3 và bài báo: trích dẫn pháp luật

| Nội dung | Bản trước | Theo văn bản gốc |
|---|---|---|
| Thù lao tác giả theo Điều 135 Luật Sở hữu trí tuệ | Tối thiểu 30% lợi nhuận thuần | Khi không có thỏa thuận: 10% lợi nhuận trước thuế nếu tự sử dụng, 15% mỗi lần nhận tiền chuyển giao (điểm b khoản 7 Điều 71 Luật 93/2025/QH15; Văn bản hợp nhất 67/VBHN-VPQH) |
| Nghĩa vụ công khai, Luật 125/2025/QH15 | Điều 28 khoản 2 điểm đ | Điều 28 khoản 3 điểm đ |
| Điều 27 Luật 125/2025/QH15 | Tài sản trí tuệ là chức năng cốt lõi | Là một nội dung hoạt động khoa học công nghệ (điểm e khoản 3) |
| Điều 86 Luật Sở hữu trí tuệ | Chủ động nộp đơn, không cần chờ phê duyệt | Có quyền đăng ký (điểm c khoản 1, bổ sung tại khoản 22 Điều 1 Luật 131/2025/QH15) |
| Quyết định 1624/QĐ-TTg | Ngày 18/8/2026; học phần bắt buộc cho khối kỹ thuật (điểm h khoản 1); trung tâm định giá tại điểm d khoản 4 | Ngày 21/8/2026; "nghiên cứu đưa" sở hữu trí tuệ thành nội dung học bắt buộc tại mọi cơ sở giáo dục đại học (điểm b khoản 8); trung tâm định giá ở điểm a khoản 6; bổ sung yêu cầu đăng ký bảo hộ đồng thời với công bố (điểm b khoản 4) |
| Văn bản hợp nhất 67/VBHN-VPQH | Ủy ban Thường vụ Quốc hội, ngày 08/01/2026 | Văn phòng Quốc hội, ngày 23/3/2026 |
| Chỉ thị 02/CT-TTg | Không ghi ngày | Ngày 30/01/2026 |

### 3.3. Chương 3: nội dung khác

- **Quyết định 217/QĐ-ĐHTĐ:** là văn bản ban hành Quy chế quản trị tài sản trí tuệ năm 2024 (xác nhận ở lượt 3). Nội dung mô tả trong tài liệu tham khảo bài báo được sửa từ "sửa đổi, bổ sung Quy chế hoạt động khoa học và công nghệ" thành "ban hành Quy chế quản trị tài sản trí tuệ".
- **Mục 3.2.5 viện dẫn "kết quả khảo sát thực trạng tại Chương 2":** Chương 2 không có khảo sát. Đoạn này đã được thay bằng các dữ kiện hành vi kiểm chứng được.
- **Mục 3.3.2:** trình bày "kết quả mô phỏng" và "bài học từ thí điểm" (15 - 20 phút mỗi phiên nghiệm thu) như kết quả đã có. Đã chuyển thành giả định cần kiểm chứng, và minh họa bằng hai đơn sáng chế thật (năm 2025 và 2026).
- **Nơi thí điểm:** chuyển từ Viện Nghiên cứu giáo dục và Chuyển giao tri thức sang Viện Y - Dược. Lý do: Viện Nghiên cứu giáo dục và Chuyển giao tri thức không có đề tài cấp cơ sở nào năm 2025, còn 10 trên 11 sản phẩm đủ điều kiện liên quan đến lĩnh vực dược.
- **Giai đoạn 3 của lộ trình ghi "kết nối" với Mạng lưới Trung tâm Hỗ trợ công nghệ và đổi mới sáng tạo:** Chương 2 xác định Nhà trường đã là thành viên từ năm 2023, nên đã sửa cho khớp.
- **Tên đơn vị sửa theo danh sách nhân sự 2026 và Quy chế 2024:**
  - "Phòng Quản lý Khoa học và Công nghệ" → Phòng Khoa học Công nghệ;
  - "Phòng Kế hoạch Tài chính" → Phòng Tài chính - Kế toán;
  - "Phòng Tổ chức Cán bộ" → Bộ phận Hành chính - Nhân sự.
- **Tên quỹ:** "Quỹ Ngô Xuân Đỗ" → Quỹ Học bổng sau tiến sĩ Ngô Xuân Độ.

## 4. Nội dung đã kiểm chứng độc lập

- Phép so khớp tác giả bài báo với danh sách nhân sự được tái lập bằng mã riêng, cho kết quả 256 trên 405 bài khớp, đúng như Chương 2. Số bài theo đơn vị: Viện Y - Dược 55, Viện Quản trị và Công nghệ 57, Viện Nghiên cứu giáo dục và Chuyển giao tri thức 17, khớp với Chương 2. Các chỉ số tập trung chỉ chênh 1 - 3 đơn vị (89 người có bài so với 88; Gini 0,828 so với 0,829), do có 5 họ tên trùng nhau trong danh sách. Chương 2 giữ nguyên số liệu này.
- Số liệu nhân sự, sản phẩm khoa học, đề tài, tài sản trí tuệ, giờ quy đổi và mức thưởng đều khớp với tài liệu gốc sau khi sửa.

## 5. Nội dung chưa có tài liệu gốc trong thư mục (cần bổ sung minh chứng)

1. Điều lệ Quỹ Học bổng sau tiến sĩ Ngô Xuân Độ, Điều 9 (tỷ lệ 50/50 năm đầu, 20/80 từ năm thứ hai; khoản chi phí khác tối đa 50%). Ngân sách 5 tỷ đồng đã được xác nhận. Chỉ cần tệp Điều lệ nếu hội đồng yêu cầu minh chứng cho các tỷ lệ được trích.
2. Năm hợp đồng của Viện Nghiên cứu giáo dục và Chuyển giao tri thức (mức 10% doanh số; 100 triệu đồng dịch vụ). Chỉ cần khi hội đồng yêu cầu minh chứng.
3. Tư cách thành viên Mạng lưới Trung tâm Hỗ trợ công nghệ và đổi mới sáng tạo từ năm 2023.
4. Việc Viện Nghiên cứu giáo dục và Chuyển giao tri thức đã đăng ký hoạt động khoa học công nghệ và có con dấu riêng.
5. Sơ đồ cơ cấu tổ chức ngày 16/6/2026.
6. Chi phí nộp đơn sáng chế từ 3 đến 5 triệu đồng và thời gian xét duyệt từ 18 đến 36 tháng (Chương 3, Mục 3.2.4). Cần dẫn biểu phí hoặc văn bản quy định.
7. Ngày ban hành của Quyết định 217/QĐ-ĐHTĐ (bản quy chế được cung cấp để trống ngày).

## 6. Điểm chưa thống nhất giữa các tài liệu gốc

- **Số hiệu văn bản 213:** trang bìa quy chế ghi "Nghị quyết số 213/QĐ-ĐHTĐ", Kế hoạch 07 ghi "Nghị quyết số 213/NQ-HĐT-ĐHTĐ", tên tệp ghi "QĐ 213". Chương 2 hiện ghi "văn bản số 213/QĐ-ĐHTĐ, sau đây gọi là Quyết định 213".
- **Tham luận hội thảo quốc tế năm 2022 và 2023:** danh mục (xếp theo ngày tổ chức) cho 2 và 3; Kế hoạch 07 ghi 1 và 4. Chương 2 dùng số liệu danh mục.
- **Tệp tổng hợp `Du_lieu_thong_ke_KHCN_Thanh_Do_2021_2025.xlsx`:** ghi sai dãy bài báo trong nước theo năm và không nên dùng để trích dẫn.

## 7. Tài liệu chưa rà trong lượt này

- **Thuyết minh v4:** đã được phê duyệt từ đầu kỳ. Một số điểm không khớp:
  - thời gian thực hiện đến 30/11 nhưng Hoạt động 4.1 kéo sang tháng 12;
  - phạm vi thời gian ghi "từ năm 2020";
  - dòng tổng của bảng kinh phí;
  - lỗi "Chủ nghiệm".
- **Báo cáo tiến độ ngày 24/9/2026:** ghi ngày nghiệm thu 30/10/2026, sớm hơn các hoạt động tháng 11.
- **`Bao_cao_toan_van_Thuc_trang_va_Giai_phap_SHTT_Thanh_Do.md`:** là bản cũ, vẫn ghi 37 đề tài. Nên chuyển vào thư mục lưu trữ.

## 8. Dựng lại

```bash
python3 scripts/ban_cuoi/dung_tat_ca.py     # ba chương bản cuối vào thư mục Ban_cuoi
python3 scripts/ra_soat/ra_soat_chuong3.py  # Chương 3 có theo dõi thay đổi
python3 scripts/ra_soat/ra_soat_bai_bao.py  # Bài báo có theo dõi thay đổi
```

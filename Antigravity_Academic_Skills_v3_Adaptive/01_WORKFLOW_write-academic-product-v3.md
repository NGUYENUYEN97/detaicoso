# /write-academic-product-v3

## Mục đích
Workflow tổng quát để viết sản phẩm học thuật bằng Antigravity với khả năng tự nhận diện chuẩn trình bày/trích dẫn/template và kích hoạt các skill phù hợp.

## Step 1 — Nhận diện nhiệm vụ và tài liệu đầu vào
Xác định:
- Loại sản phẩm: chuyên đề tổng quan, 3 chuyên đề nghiên cứu, bài báo, đề tài cấp cơ sở, luận án, chương luận án.
- Chủ đề/đề tài.
- Dữ liệu/số liệu/phỏng vấn/khảo sát.
- Tài liệu tham khảo.
- Tài liệu mẫu/template/quy định trình bày.

Nếu có tài liệu mẫu, không hỏi lại dài dòng; tự trích xuất quy tắc trình bày và tạo hồ sơ format.

## Step 2 — Kích hoạt `academic-format-detector-router`
Tạo `FORMAT_REQUIREMENTS_PROFILE` gồm:
- Chuẩn bố cục.
- Chuẩn trích dẫn.
- Chuẩn Danh mục tài liệu tham khảo.
- Quy định tiêu đề, chữ viết tắt, bảng/hình, phụ lục.
- Các skill cần kích hoạt.

## Step 3 — Lập hồ sơ học thuật của tác giả
Cập nhật hoặc yêu cầu `AUTHOR_PROFILE_CARD`:
- Chuyên môn và kinh nghiệm của tác giả.
- Quan điểm học thuật riêng.
- Dữ liệu/số liệu riêng.
- Luận điểm cần bảo vệ/phản biện.

## Step 4 — Lập ma trận bằng chứng
Kích hoạt `evidence-and-citation-auditor`:
- Claim.
- Evidence.
- Source.
- Warrant.
- Author contribution.
- Risk.
- Citation status.

## Step 5 — Thiết kế cấu trúc sản phẩm
Kích hoạt `academic-product-architect`:
- Nếu format yêu cầu đề cương 4 chương, dùng template 4 chương.
- Nếu format yêu cầu chuyên đề tổng quan, dùng template chuyên đề tổng quan.
- Nếu format khác, tự tạo cấu trúc theo tài liệu mẫu và giải trình.

## Step 6 — Kích hoạt skill chuyên biệt theo nhu cầu
- Nếu có chuẩn [n]/Danh mục tài liệu tham khảo: dùng `citation-style-adaptive-csdt`.
- Nếu có khảo sát/phỏng vấn: dùng `survey-designer`.
- Nếu có phân tích định lượng: dùng `quantitative-analysis-guide`.
- Nếu có pháp luật/chính sách/quản lý nhà nước: dùng `legal-writing-style`.
- Nếu là luận án/chuỗi chuyên đề: dùng `thesis-dissertation-coherence`.

## Step 7 — Viết bản nháp học thuật
Viết theo CEW-C, bám dữ liệu, nguồn và quan điểm tác giả. Không viết vượt quá bằng chứng.

## Step 8 — Biên tập giọng học giả
Kích hoạt `human-scholar-voice-editor` để rà soát:
- Liệt kê máy móc.
- Cấu trúc đối xứng giả tạo.
- Thiếu phản biện.
- Sáo ngữ.
- Diễn ngôn pháp lý chưa chuẩn.
- Chữ viết tắt chưa nhất quán.

## Step 9 — Kiểm tra cuối
Kích hoạt:
- `evidence-and-citation-auditor` để kiểm tra nguồn và trích dẫn.
- `citation-style-adaptive-csdt` để kiểm tra [n] hoặc chuẩn trích dẫn tương ứng.
- `thesis-dissertation-coherence` nếu có chuỗi chuyên đề/luận án.
- `ai-disclosure-and-integrity` để kiểm tra tuyên bố AI.

## Step 10 — Xuất đầu ra
Luôn xuất:
1. Bản nháp/khung đề cương.
2. FORMAT_REQUIREMENTS_PROFILE.
3. EVIDENCE_MATRIX.
4. DECISION_LOG.
5. CITATION_CHECK_MATRIX.
6. ABBREVIATION_REGISTER nếu có.
7. Danh sách điểm cần tác giả xác nhận.
8. Checklist cuối.

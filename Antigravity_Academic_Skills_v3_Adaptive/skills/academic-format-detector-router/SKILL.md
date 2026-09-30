---
name: academic-format-detector-router
description: Tự nhận diện chuẩn trình bày, template, chuẩn trích dẫn, bố cục chương/phần, quy tắc Danh mục tài liệu tham khảo và điều phối các skill học thuật phù hợp.
---

# Academic Format Detector & Router

## Mục tiêu
Skill này đọc tài liệu mẫu/quy định trình bày/yêu cầu nơi nộp để tạo hồ sơ định dạng và tự kích hoạt các skill phù hợp mà không hỏi lại quá nhiều.

## Khi dùng
- Khi có file template của cơ sở đào tạo, tạp chí, hội thảo hoặc đề tài.
- Khi có yêu cầu “theo APA 7th”, “theo chuẩn [n]”, “theo mẫu 4 chương”, “theo quy định cơ sở đào tạo”.
- Khi cần xác định bố cục, cách trích dẫn, tài liệu tham khảo, chữ viết tắt, bảng/hình.

## Quy trình
1. Đọc tài liệu mẫu/quy định.
2. Trích xuất quy tắc về:
   - Loại sản phẩm.
   - Bố cục.
   - Heading.
   - Font/lề/giãn dòng nếu có.
   - Trích dẫn trong văn bản.
   - Danh mục tài liệu tham khảo.
   - Bảng/hình/phụ lục.
   - Chữ viết tắt.
3. Tạo `FORMAT_REQUIREMENTS_PROFILE`.
4. Xác định skill cần kích hoạt:
   - `citation-style-adaptive-csdt` nếu có quy tắc trích dẫn/tài liệu tham khảo.
   - `academic-product-architect` nếu cần đề cương/bố cục.
   - `survey-designer` nếu có khảo sát/phỏng vấn.
   - `quantitative-analysis-guide` nếu có dữ liệu định lượng.
   - `legal-writing-style` nếu có pháp luật/chính sách/quản lý nhà nước.
   - `thesis-dissertation-coherence` nếu có chuyên đề/luận án.

## Quy tắc linh hoạt
- Không mặc định mọi trường hợp đều dùng chuẩn của cơ sở đào tạo.
- Nếu tài liệu mẫu yêu cầu APA 7th, áp dụng APA 7th.
- Nếu tài liệu mẫu yêu cầu [n], áp dụng [n].
- Nếu không có chuẩn rõ, hỏi tối đa 3 câu ngắn; sau đó tạo bản nháp với nhãn “cần xác nhận chuẩn”.

## Đầu ra
- FORMAT_REQUIREMENTS_PROFILE.
- Danh sách skill được kích hoạt.
- Các điểm chưa rõ cần tác giả xác nhận.

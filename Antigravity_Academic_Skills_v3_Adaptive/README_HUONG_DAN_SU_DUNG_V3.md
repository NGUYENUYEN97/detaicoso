# Bộ Google Antigravity Academic Skills V3 – Adaptive Academic Product System

## 1. Mục tiêu
Bộ V3 giúp Antigravity hỗ trợ tác giả xây dựng khung đề cương, chuyên đề, bài báo, đề tài cấp cơ sở và luận án tiến sĩ theo nguyên tắc:

**Tác giả làm chủ chuyên môn – AI hỗ trợ quy trình – Nguồn và dữ liệu kiểm chứng – Định dạng linh hoạt theo yêu cầu từng nơi nộp – Tác giả đọc, duyệt, sửa và chịu trách nhiệm bản cuối.**

Bộ này không dùng để né công cụ nhận diện AI. Bộ này dùng để tạo sản phẩm học thuật có căn cứ, có dấu ấn tác giả, đúng quy cách trình bày/trích dẫn/tài liệu tham khảo theo từng chuẩn cụ thể.

## 2. Điểm mới của V3
1. Thêm cơ chế **Adaptive Format Detection**: tự nhận diện chuẩn trình bày từ tài liệu mẫu, ví dụ chuẩn của cơ sở đào tạo, APA 7th, Vancouver, IEEE hoặc template khác.
2. Thêm cơ chế **Citation Router**: tự chọn cách trích dẫn [n], tác giả-năm, footnote hoặc chuẩn khác theo hồ sơ định dạng.
3. Thêm 4 Skill mới: `citation-style-adaptive-csdt`, `survey-designer`, `quantitative-analysis-guide`, `legal-writing-style`.
4. Thêm 1 Skill điều phối: `academic-format-detector-router` để đọc template/format mẫu và kích hoạt các skill phù hợp.
5. Nâng cấp các Skill cũ để tương thích nhiều chuẩn trình bày, không khóa cứng vào một trường hợp cụ thể.
6. Bổ sung template: đề cương 4 chương của cơ sở đào tạo, chuyên đề tổng quan, 3 chuyên đề nghiên cứu, ma trận nhất quán, kiểm tra chữ viết tắt, kiểm tra trích dẫn [n] – Danh mục tài liệu tham khảo.

## 3. Cách cài đặt

### 3.1. Rule nền
Dán nội dung file `00_GLOBAL_RULE_adaptive-academic-integrity-v3.md` vào Global Rule hoặc Workspace Rule của Antigravity.

### 3.2. Skills
Copy toàn bộ thư mục con trong `skills/` vào:

```text
<workspace-root>/.agents/skills/
```

### 3.3. Workflow
Tạo workflow trong Antigravity với tên:

```text
/write-academic-product-v3
```

Dán nội dung file:

```text
01_WORKFLOW_write-academic-product-v3.md
```

## 4. Quy trình dùng nhanh
1. Nạp tài liệu mẫu, quy định trình bày, tài liệu tham khảo, dữ liệu, số liệu.
2. Gọi `/write-academic-product-v3`.
3. Antigravity phải tự tạo `FORMAT_REQUIREMENTS_PROFILE.md` trước khi viết.
4. Antigravity tự chọn chuẩn trích dẫn, cấu trúc, kiểu tài liệu tham khảo, kiểm tra chữ viết tắt, kiểm tra bảng/hình.
5. Tác giả duyệt khung đề cương và ma trận bằng chứng.
6. Antigravity viết bản nháp theo CEW-C.
7. Tác giả đọc, sửa, bổ sung quan điểm và dữ liệu riêng.
8. Antigravity kiểm tra cuối: nguồn, trích dẫn, định dạng, chữ viết tắt, nhất quán chuyên đề–luận án, tuyên bố AI.

## 5. Prompt gọi chuẩn
Xem file `02_PROMPT_MAU_SU_DUNG_V3.md`.

---
name: citation-style-adaptive-csdt
description: Xử lý trích dẫn và Danh mục tài liệu tham khảo linh hoạt: chuẩn [n] của cơ sở đào tạo, APA 7th hoặc chuẩn khác; kiểm tra [n] khớp DMTLTK và nguồn web có ngày truy cập.
---

# Citation Style Adaptive CSĐT

## Mục tiêu
Tự nhận diện và áp dụng chuẩn trích dẫn/tài liệu tham khảo theo tài liệu mẫu hoặc yêu cầu nơi nộp.

## Chế độ chuẩn [n]
Áp dụng khi tài liệu mẫu yêu cầu trích dẫn bằng số thứ tự trong ngoặc vuông.

Quy tắc:
- Trích dẫn trong văn bản dạng [n].
- Nếu trích dẫn nhiều tài liệu, đặt độc lập từng ngoặc vuông theo thứ tự tăng dần, ví dụ [19], [25], [41].
- Mọi [n] phải có tài liệu tương ứng trong Danh mục tài liệu tham khảo.
- Nguồn web phải có URL và ngày truy cập nếu quy định yêu cầu.
- Danh mục tài liệu tham khảo có thể xếp theo ngôn ngữ và ABC nếu template quy định.

## Chế độ APA 7th
Áp dụng khi yêu cầu APA 7th:
- Trích dẫn trong văn bản theo tác giả-năm.
- Danh mục tài liệu tham khảo xếp ABC theo họ tác giả/cơ quan.
- DOI/URL trình bày theo APA 7th.

## Chế độ chuẩn khác
Nếu yêu cầu Vancouver/IEEE/Chicago/Harvard:
- Tạo hồ sơ chuẩn.
- Nêu các quy tắc chính.
- Áp dụng nhất quán.
- Đánh dấu điểm cần tác giả xác nhận nếu thiếu thông tin.

## Kiểm tra tự động
- Extract citation markers.
- Match with reference list.
- Detect missing references.
- Detect uncited references.
- Check web access date.
- Check language grouping if required.
- Check author sorting if required.

## Đầu ra
- REFERENCE_STYLE_PROFILE.
- CITATION_CHECK_MATRIX.
- Danh sách lỗi và cách sửa.
